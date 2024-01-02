import sys
import os
import asyncio
import tabula
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QTextEdit, QFileDialog, QLabel
from PyQt5.QtGui import QTextCursor
from PyQt5.QtCore import Qt

class ConsoleRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.insertPlainText(message)

class QCMApplication(QWidget):
    def __init__(self):
        super().__init__()

        # Ajouter un widget Text pour afficher la console
        self.console_text = QTextEdit(self)
        self.console_text.setReadOnly(True)
        self.console_text.setPlaceholderText("Console output will be displayed here.")

        # Boutons
        self.start_button = QPushButton("Start", self)
        self.start_button.clicked.connect(self.start_application)

        self.quit_button = QPushButton("Quitter", self)
        self.quit_button.clicked.connect(self.quit_application)

        self.select_pdfs_button = QPushButton("Sélectionner QCM_pdfs", self)
        self.select_pdfs_button.clicked.connect(self.select_pdfs_folder)

        
        self.select_correct_file_button = QPushButton("Sélectionner QCM_correct.xlsx", self)
        self.select_correct_file_button.clicked.connect(self.select_correct_file)
        
        self.correct_button = QPushButton("Corriger", self)
        self.correct_button.clicked.connect(self.correct_quiz)

        self.store_results_button = QPushButton("Stockage de résultats", self)
        self.store_results_button.clicked.connect(self.extract_data_and_store_results)

        self.show_results_button = QPushButton("Afficher résultats", self)
        self.show_results_button.clicked.connect(self.show_results)

        self.statistics_button = QPushButton("Statistics", self)
        self.statistics_button.clicked.connect(self.show_statistics)

        # Étiquette pour afficher le nombre total de fichiers PDF
        self.pdf_count_label = QLabel(self)
        self.pdf_count_label.setText("Nombre total de fichiers PDF : ")

        # Organiser les widgets dans une mise en page verticale
        layout = QVBoxLayout(self)
        layout.addWidget(self.console_text)
        layout.addWidget(self.start_button)
        layout.addWidget(self.quit_button)
        layout.addWidget(self.select_pdfs_button)
        layout.addWidget(self.select_correct_file_button)
        layout.addWidget(self.correct_button)
        layout.addWidget(self.store_results_button)
        layout.addWidget(self.show_results_button)
        layout.addWidget(self.statistics_button)
        layout.addWidget(self.pdf_count_label)

        # Rediriger la sortie de la console vers le widget Text
        sys.stdout = ConsoleRedirector(self.console_text)

    def start_application(self):
        print("L'application a démarré.")

    def quit_application(self):
        print("Good bye!")
        self.close()

    def select_pdfs_folder(self):
        global folder_path
        folder_path = QFileDialog.getExistingDirectory(self, "Sélectionnez le dossier des PDFs")
        folder_path = os.path.join(folder_path, "QCM_pdfs")
        print("Dossier des PDFs sélectionné :", folder_path)
        self.display_pdf_count()
    
    def select_correct_file(self):
        global correct_answers_file
        correct_answers_file, _ = QFileDialog.getOpenFileName(self, "Sélectionnez QCM_correct.xlsx")
        print("Fichier QCM_correct.xlsx sélectionné :", correct_answers_file)
        self.select_correct_file()
        
    def display_pdf_count(self):
        if folder_path:
            pdf_count = count_pdfs_in_folder(folder_path)
            print(f"Nombre total de fichiers PDF : {pdf_count}")
            self.pdf_count_label.setText(f"Nombre total de fichiers PDF : {pdf_count}")
        else:
            print("Sélectionnez d'abord le dossier des PDFs.")
            self.pdf_count_label.setText("Sélectionnez d'abord le dossier des PDFs.")

    def correct_quiz(self):
        if not folder_path:
            print("Sélectionnez d'abord le dossier des PDFs.")
            return
        tables_list = asyncio.run(extract_tables_from_pdfs())
        save_tables_as_excel(tables_list)
        print("Tables extraites et sauvegardées avec succès.")
        print("Quiz corrigé avec succès.")

    def extract_data_and_store_results(self):
        notes = []
        names = []
        for filename in os.listdir(excel_folder):
            if filename.endswith(".xlsx"):
                file_path = os.path.join(excel_folder, filename)
                workbook = openpyxl.load_workbook(file_path)
                worksheet = workbook.active
                cell = worksheet["C2"]
                final_score = correct_excel_quiz(file_path, correct_answers_file)
                new_name = cell.value
                names.append(new_name)
                new_note = final_score
                final_score_formatted = "{0:.2f}".format(final_score)
                notes.append(final_score_formatted)

        self.store_results_in_excel(names, notes)
        print("Stockage de résultats avec succès")

    def correct_excel_quiz(self, student_file, correct_answers_file):
        student_answers = pd.read_excel(student_file)
        answers = pd.read_excel(correct_answers_file)
        correct_count = 0
        for i, row in student_answers.iterrows():
            answer = row['Answer']
            if answer == answers.iloc[i]['Correct Answer']:
                correct_count += 1
        final_score = correct_count
        return final_score

    def store_results_in_excel(self, names, notes):
        if not os.path.exists("resultats.xlsx"):
            workbook = openpyxl.Workbook()
            worksheet = workbook.active
            worksheet["A1"] = "Noms"
            worksheet["B1"] = "Notes"
            workbook.save("resultats.xlsx")

        df = pd.read_excel("resultats.xlsx")
        new_data = {"Noms": names, "Notes": notes}
        new_df = pd.DataFrame(data=new_data)
        df = pd.concat([df, new_df], ignore_index=True)
        df.to_excel("resultats.xlsx", index=False)

    def show_results(self):
        os.system("xdg-open resultats.xlsx")

    def show_statistics(self):
        statistics_window = QWidget()
        statistics_window.setWindowTitle("Statistics")

        statistics_button = QPushButton("Generate Statistics", statistics_window)
        statistics_button.clicked.connect(self.generate_statistics)
        layout = QVBoxLayout(statistics_window)
        layout.addWidget(statistics_button)

        statistics_window.show()

    def generate_statistics(self):
        df = pd.read_excel("resultats.xlsx")
        bins = [0, 5, 10, 15, 18, 20]
        labels = ['0-5', '6-10', '11-15', '16-18', '19-20']
        df['NoteGroup'] = pd.cut(df['Notes'], bins=bins, labels=labels)
        group_counts = df['NoteGroup'].value_counts()
        colors = ['#FF9999', '#66B2FF', '#99FF99', '#FFCC99', '#c2c2f0']
        best_student_info = df.loc[df['Notes'].idxmax(), ['Noms', 'Notes']]
        worst_student_info = df.loc[df['Notes'].idxmin(), ['Noms', 'Notes']]

        plt.figure(figsize=(8, 8))
        wedges, texts, autotexts = plt.pie(group_counts, labels=group_counts.index, autopct='%1.1f%%', startangle=140, colors=colors, textprops=dict(color="w"))
        plt.legend(wedges, labels, title="Intervalles de notes", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
        plt.title(f"Autocorrect_QCM result_graph\nMeilleure note: {best_student_info['Notes']} ({best_student_info['Noms']})\nMoins bonne note: {worst_student_info['Notes']} ({worst_student_info['Noms']})")
        plt.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    qcm_app = QCMApplication()
    qcm_app.show()
    sys.exit(app.exec_())

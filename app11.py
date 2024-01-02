import sys
import os
import asyncio
import tabula
import pandas as pd
import openpyxl
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QTextEdit, QFileDialog, QLabel
from PyQt5.QtGui import QTextCursor
from PyQt5.QtCore import Qt

# Définir la classe ConsoleRedirector pour rediriger la sortie de la console vers un widget Text
class ConsoleRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.insertPlainText(message)

# Définir des variables globales pour le dossier des PDFs, le fichier correct_answers et le dossier des excels
feuilles_corrigees = 0
folder_path = ""
correct_answers_file = ""
excel_folder = "QCM_excels"

# Définir le répertoire de travail
os.chdir('/home/user/alx/Autocorrect_QCM')

class QCMApplication(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        # Boutons
        self.start_button = QPushButton('Start', self)
        self.select_pdfs_button = QPushButton('Sélectionner QCM_pdfs', self)
        self.select_correct_file_button = QPushButton('Sélectionner QCM_correct.xlsx', self)
        self.correct_button = QPushButton('Corriger', self)
        self.store_results_button = QPushButton('Stockage de résultats', self)
        self.show_results_button = QPushButton('Afficher résultats', self)
        self.statistics_button = QPushButton('Statistics', self)
        self.quit_button = QPushButton('Quitter', self)

        # Labels
        self.pdf_count_label = QLabel(self)

        # Layout
        vbox = QVBoxLayout()
        vbox.addWidget(self.start_button)
        vbox.addWidget(self.select_pdfs_button)
        vbox.addWidget(self.select_correct_file_button)
        vbox.addWidget(self.correct_button)
        vbox.addWidget(self.store_results_button)
        vbox.addWidget(self.show_results_button)
        vbox.addWidget(self.statistics_button)
        vbox.addWidget(self.quit_button)
        vbox.addWidget(self.pdf_count_label)

        # Set layout
        self.setLayout(vbox)

        # Boutons connect
        self.start_button.clicked.connect(self.start_application)
        self.select_pdfs_button.clicked.connect(self.select_pdfs_folder)
        self.select_correct_file_button.clicked.connect(self.select_correct_file)
        self.correct_button.clicked.connect(self.correct_quiz)
        self.store_results_button.clicked.connect(self.extract_data_and_store_results)
        self.show_results_button.clicked.connect(self.show_results)
        self.statistics_button.clicked.connect(self.show_statistics)
        self.quit_button.clicked.connect(self.quit_application)

        # Initialiser l'application
        self.start_application()

    def start_application(self):
        print("L'application a démarré.")

    def select_pdfs_folder(self):
        global folder_path
        selected_folder = QFileDialog.getExistingDirectory(self, "Sélectionnez le dossier des PDFs", "")
        folder_path = os.path.join(selected_folder, "QCM_pdfs")
        print("Dossier des PDFs sélectionné :", folder_path)
        self.display_pdf_count()

    def select_correct_file(self):
        global correct_answers_file
        correct_answers_file, _ = QFileDialog.getOpenFileName(self, "Sélectionnez QCM_correct.xlsx", "", "Excel Files (*.xlsx)")
        print("Fichier QCM_correct.xlsx sélectionné :", correct_answers_file)

    def display_pdf_count(self):
        if folder_path:
            pdf_count = self.count_pdfs_in_folder(folder_path)
            print(f"Nombre total de fichiers PDF : {pdf_count}")
            self.pdf_count_label.setText(f"Nombre total de fichiers PDF : {pdf_count}")
        else:
            print("Sélectionnez d'abord le dossier des PDFs.")
            self.pdf_count_label.setText("Sélectionnez d'abord le dossier des PDFs.")

    def count_pdfs_in_folder(self, folder):
        pdf_count = sum(1 for file in os.listdir(folder) if file.lower().endswith('.pdf'))
        return pdf_count

    async def extract_tables_from_pdfs(self):
        tables_list = []
        for filename in os.listdir(folder_path):
            if filename.endswith(".pdf"):
                file_path = os.path.join(folder_path, filename)
                tables = tabula.read_pdf(file_path, pages="all")
                for table in tables:
                    tables_list.append(table)
        return tables_list

    def save_tables_as_excel(self, tables_list):
        for i, table in enumerate(tables_list):
            file_path = os.path.join(excel_folder, f'QCM_table{i}.xlsx')
            table.to_excel(file_path, index=False)

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

    def correct_quiz(self):
        if not folder_path or not correct_answers_file:
            print("Sélectionnez d'abord le dossier des PDFs et le fichier correct_answers.")
            return
        tables_list = asyncio.run(self.extract_tables_from_pdfs())
        self.save_tables_as_excel(tables_list)
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
                cell.value
                final_score = self.correct_excel_quiz(file_path, correct_answers_file)
                new_name = cell.value
                names.append(new_name)
                new_note = final_score
                final_score_formatted = "{0:.2f}".format(final_score)
                notes.append(final_score_formatted)

        self.store_results_in_excel(names, notes)
        print("Stockage de résultats avec succès")

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
        statistics_window = StatisticsWindow()
        statistics_window.show()

    def quit_application(self):
        print("Good bye!")
        sys.exit()

class StatisticsWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        # Créer des widgets pour afficher les statistiques
        label = QLabel('Statistiques', self)
        label.setAlignment(Qt.AlignCenter)

        # Ajouter d'autres widgets selon vos besoins (tableau, graphiques, etc.)

        # Créer un layout vertical pour organiser les widgets
        vbox = QVBoxLayout()
        vbox.addWidget(label)
        # Ajouter d'autres widgets au layout

        # Définir le layout pour la fenêtre
        self.setLayout(vbox)

        # Définir la taille de la fenêtre
        self.setGeometry(300, 300, 400, 300)
        # Définir le titre de la fenêtre
        self.setWindowTitle('Statistics Window')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = QCMApplication()
    sys.exit(app.exec_())

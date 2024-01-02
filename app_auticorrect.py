import os
import pandas as pd
import openpyxl
import tabula
import tkinter as tk
from tkinter import filedialog
from datetime import datetime, timedelta

class QCMApp:
    def __init__(self, master):
        self.master = master
        self.master.title("QCM Correction App")

        self.folder_path = ""
        self.correct_answers_file = ""

        self.create_widgets()

    def create_widgets(self):
        # Bouton pour sélectionner le dossier QCM_pdfs
        self.btn_select_folder = tk.Button(self.master, text="Sélectionner le dossier QCM_pdfs", command=self.select_folder)
        self.btn_select_folder.pack(pady=10)

        # Bouton pour sélectionner le fichier QCM_correct.xlsx
        self.btn_select_correct_answers = tk.Button(self.master, text="Sélectionner le fichier QCM_correct.xlsx", command=self.select_correct_answers)
        self.btn_select_correct_answers.pack(pady=10)

        # Bouton pour lancer la correction
        self.btn_correct = tk.Button(self.master, text="Corriger", command=self.correct)
        self.btn_correct.pack(pady=10)

    def select_folder(self):
        self.folder_path = filedialog.askdirectory(title="Sélectionner le dossier QCM_pdfs")

    def select_correct_answers(self):
        self.correct_answers_file = filedialog.askopenfilename(title="Sélectionner le fichier QCM_correct.xlsx", filetypes=[("Fichiers Excel", "*.xlsx")])

    def correct(self):
        if not self.folder_path or not self.correct_answers_file:
            tk.messagebox.showerror("Erreur", "Veuillez sélectionner le dossier QCM_pdfs et le fichier QCM_correct.xlsx")
            return

        tables_list = self.extract_tables_from_pdfs()
        excel_folder = self.create_excel_folder("QCM_excels")
        self.save_tables_as_excel(tables_list, excel_folder)
        names, notes = self.extract_data_from_excel(excel_folder)
        self.store_results_in_excel(names, notes)
        self.display_results(names, notes)

    def extract_tables_from_pdfs(self):
        tables_list = []
        for filename in os.listdir(self.folder_path):
            if filename.endswith(".pdf"):
                file_path = os.path.join(self.folder_path, filename)
                tables = tabula.read_pdf(file_path, pages="all")
                for table in tables:
                    tables_list.append(table)
        return tables_list

    def create_excel_folder(self, excel_folder):
        if not os.path.exists(excel_folder):
            os.makedirs(excel_folder)
        return excel_folder

    def save_tables_as_excel(self, tables_list, excel_folder):
        for i, table in enumerate(tables_list):
            file_path = os.path.join(excel_folder, 'QCM_table{}.xlsx'.format(i))
            table.to_excel(file_path, index=False)

    def extract_data_from_excel(self, excel_folder):
        notes = []
        names = []

        for filename in os.listdir(excel_folder):
            if filename.endswith(".xlsx"):
                file_path = os.path.join(excel_folder, filename)
                workbook = openpyxl.load_workbook(file_path)
                worksheet = workbook.active
                cell = worksheet["C2"]
                final_score = self.correct_excel_quiz(file_path)
                new_name = cell.value
                names.append(new_name)
                final_score_formatted = "{0:.2f}".format(final_score)
                notes.append(final_score_formatted)

        return names, notes

    def correct_excel_quiz(self, student_file):
        student_answers = pd.read_excel(student_file)
        answers = pd.read_excel(self.correct_answers_file)
        correct_count = 0

        for i, row in student_answers.iterrows():
            answer = row['Answer']
            if answer == answers.iloc[i]['Correct Answer']:
                correct_count += 1

        final_score = (correct_count)

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

    def display_results(self, names, notes):
        now = datetime.now()
        one_hour_later = now + timedelta(hours=1)
        result_text = f"Date et heure actuelles : {one_hour_later}\n\nNombre d'éléments dans la liste : {len(names)}\n\nNoms : {names}\n\nNotes : {notes}"
        tk.messagebox.showinfo("Résultats", result_text)

if __name__ == "__main__":
    root = tk.Tk()
    app = QCMApp(root)
    root.mainloop()

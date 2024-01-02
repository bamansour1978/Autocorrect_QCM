import pandas as pd
import matplotlib.pyplot as plt

def create_pie_chart_with_legend_and_extremes(input_file):
    # Charger les données depuis le fichier Excel
    df = pd.read_excel(input_file)

    # Diviser les étudiants en cinq groupes en fonction des intervalles de notes
    bins = [0, 5, 10, 15, 18, 20]
    labels = ['0-5', '6-10', '11-15', '16-18', '19-20']
    df['NoteGroup'] = pd.cut(df['Notes'], bins=bins, labels=labels)

    # Compter le nombre d'étudiants dans chaque groupe
    group_counts = df['NoteGroup'].value_counts()

    # Définir une palette de couleurs personnalisée
    colors = ['#FF9999', '#66B2FF', '#99FF99', '#FFCC99', '#c2c2f0']

    # Trouver la meilleure et la moins bonne note avec les noms correspondants
    best_student_info = df.loc[df['Notes'].idxmax(), ['Noms', 'Notes']]
    worst_student_info = df.loc[df['Notes'].idxmin(), ['Noms', 'Notes']]

    # Créer un graphique en secteurs (pie chart) avec des couleurs personnalisées
    plt.figure(figsize=(8, 8))
    wedges, texts, autotexts = plt.pie(group_counts, labels=group_counts.index, autopct='%1.1f%%', startangle=140, colors=colors, textprops=dict(color="w"))

    # Ajouter une légende (Clé illustrée)
    plt.legend(wedges, labels, title="Intervalles de notes", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))

    # Ajouter la meilleure et la moins bonne note avec les noms correspondants à la titre
    plt.title(f"Autocorrect_QCM result_graph\nMeilleure note: {best_student_info['Notes']} ({best_student_info['Noms']})\nMoins bonne note: {worst_student_info['Notes']} ({worst_student_info['Noms']})")

    # Afficher le graphique
    plt.show()

# Exemple d'utilisation
create_pie_chart_with_legend_and_extremes('resultats.xlsx')

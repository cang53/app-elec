from flask import Flask, render_template, request
import openpyxl

app = Flask(__name__)

@app.route('/')
def index():
    # Affiche un formulaire qui permet à l'utilisateur de soumettre ses données de consommation électrique
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    # Récupère les données de consommation électrique soumises par l'utilisateur
    date = request.form['date']
    time = request.form['time']
    kwh = request.form['kwh']

    # Ouvre la feuille de calcul Excel et ajoute les données de consommation électrique à la fin de la feuille
    wb = openpyxl.load_workbook('data.xlsx')
    sheet = wb['Sheet1']
    sheet.append([date, time, kwh])
    wb.save('data.xlsx')

    # Affiche un message de confirmation à l'utilisateur
    return "Données enregistrées avec succès!"

if __name__ == '__main__':
    app.run()

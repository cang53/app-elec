import openpyxl

def main():
    # Affiche un formulaire qui permet à l'utilisateur de soumettre ses données de consommation électrique
    print("Enregistrement de la consommation électrique")
    date = input("Date (jj/mm/aaaa): ")
    consumption_day = input("Consommation jour (kwh): ")
    consumption_night = input("Consommation nuit (kwh): ")

    # Ouvre la feuille de calcul Excel et ajoute les données de consommation électrique à la fin de la feuille
    wb = openpyxl.load_workbook('/Users/cangemici/Documents/Can/app elec/data.xlsx')
    sheet = wb['Feuil1']
    sheet.append([date, consumption_day, consumption_night])
    wb.save('data.xlsx')

    # Affiche un message de confirmation à l'utilisateur
    print("Données enregistrées avec succès!")

if __name__ == '__main__':
    main()



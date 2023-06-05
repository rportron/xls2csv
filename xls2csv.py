import xlrd #gestion des fichiers xls
import csv
import sys

#@author : Raoul Portron - 29 juin 2016

def xls2csv(extraction_file, ligne_donnees = 1):
    '''
    Ce programme transforme un fichier xls en csv :
     * il lit les valeurs de chaque ligne et de chaque colonne
     * il écrit les valeurs dans un fichier csv
     et enregistre le fichier csv en ajoutant l'extension .csv au nom du fichier dans le répertoire courant
    '''
    ligne_donnees = ligne_donnees - 1 # la numérotation débute à 0
    extraction_excel = xlrd.open_workbook(extraction_file)
    feuille_xls_active = extraction_excel.sheet_by_index(0)
    fichier_csv = str(extraction_file) + '.csv'
    with open(fichier_csv, 'w', newline='', encoding="utf-8") as csvfile:
        csv_writer = csv.writer(csvfile, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        for dummy_ligne in range(ligne_donnees, feuille_xls_active.nrows):
            nouvelle_ligne = []
            for dummy_colonne in range(feuille_xls_active.ncols):
                cellule = feuille_xls_active.cell(dummy_ligne, dummy_colonne).value
                nouvelle_ligne.append(cellule.replace('\n',' _ '))
            csv_writer.writerow(nouvelle_ligne)

extraction_file = sys.argv[1]
xls2csv(extraction_file, ligne_donnees = 1)

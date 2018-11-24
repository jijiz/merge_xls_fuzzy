import xlrd
import xlwt
from xlutils.copy import copy
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

nom_fic_adh       = "adherents 2017"
nom_fic_retraites = "retraites_22_mai_2017"

book_adh = xlrd.open_workbook(nom_fic_adh+".xls")
sheet_adh = book_adh.sheet_by_index(0)

book_cible = xlrd.open_workbook(nom_fic_retraites+".xls")
sheet_cible = book_cible.sheet_by_index(0)

wb = copy(book_cible)
w_sheet = wb.get_sheet(0)

nom_adh          = []
prenom_adh       = []
indice_ligne_adh = []
for rx in range(sheet_adh.nrows):
    nom_adh.append(str(sheet_adh.cell_value(rx,1)).upper().strip())
    prenom_adh.append(str(sheet_adh.cell_value(rx,2)).upper().strip())
    indice_ligne_adh.append(rx)


for rx2 in range(5,sheet_cible.nrows):
    nom_cible = str(sheet_cible.cell_value(rx2,1)).upper().strip()
    prenom_cible = str(sheet_cible.cell_value(rx2,2)).upper().strip()
    for x in range(1,len(nom_adh)):
        if ((((nom_cible == nom_adh[x]) and (fuzz.partial_ratio(prenom_cible, prenom_adh[x]) > 80)) or ((prenom_cible == prenom_adh[x]) and (fuzz.partial_ratio(nom_cible, nom_adh[x]) > 80)))):
            w_sheet.write(rx2,14, 'ADHERENT')
            if not((nom_cible == nom_adh[x]) and (prenom_cible == prenom_adh[x])):
                w_sheet.write(rx2,15, 'A VERIFIER')
                for rx3 in range(1,sheet_adh.ncols):
                    w_sheet.write(rx2,15+rx3, sheet_adh.cell_value(indice_ligne_adh[x],rx3))
wb.save("D:\\parser xls\\"+nom_fic_retraites+"_adh.xls")
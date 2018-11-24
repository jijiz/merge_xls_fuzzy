import xlrd
import xlwt
from xlutils.copy import copy
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

def col_to_num(col_str):
    expn = 0
    col_num = 0
    for char in reversed(col_str):
        col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1

    return col_num
	
str_col_nom_adh = 'B'
str_col_prenom_adh = 'C'
str_col_date_naiss_adh = 'M'

str_col_nom_certif = 'C'
str_col_prenom_certif = 'D'
str_col_date_naiss_certif = 'Z'

str_col_bool_adh = 'AL'


num_col_nom_adh = col_to_num(str_col_nom_adh)-1
num_col_prenom_adh = col_to_num(str_col_prenom_adh)-1
num_col_date_naiss_adh = col_to_num(str_col_date_naiss_adh)-1

num_col_nom_certif = col_to_num(str_col_nom_certif)-1
num_col_prenom_certif = col_to_num(str_col_prenom_certif)-1
num_col_date_naiss_certif = col_to_num(str_col_date_naiss_certif)-1

num_col_bool_adh = col_to_num(str_col_bool_adh)-1

nom_fic_adherents = "adherents 2017"
nom_fic_certif = "certifies_ta"
book_adh = xlrd.open_workbook(nom_fic_adherents+".xls")
sheet_adh = book_adh.sheet_by_index(0)

book_cible = xlrd.open_workbook(nom_fic_certif+".xls")
sheet_cible = book_cible.sheet_by_index(0)

wb = copy(book_cible)
w_sheet = wb.get_sheet(0)

nom_adh        = []
prenom_adh     = []
date_naiss_adh = []
for rx in range(sheet_adh.nrows):
    nom_adh.append(str(sheet_adh.cell_value(rx,num_col_nom_adh)).upper().strip())
    prenom_adh.append(str(sheet_adh.cell_value(rx,num_col_prenom_adh)).upper().strip())
    date_naiss_adh.append(str(sheet_adh.cell_value(rx,num_col_date_naiss_adh)).upper().strip())


for rx2 in range(1,sheet_cible.nrows):
    nom_cible 		     = str(sheet_cible.cell_value(rx2,num_col_nom_certif)).upper().strip()
    prenom_cible 		 = str(sheet_cible.cell_value(rx2,num_col_prenom_certif)).upper().strip()
    date_naissance_cible = str(sheet_cible.cell_value(rx2,num_col_date_naiss_certif)).upper().strip()
    print(rx2)
    for x in range(1,sheet_adh.nrows):
        i_fuzz_nom       = fuzz.partial_ratio(nom_cible, nom_adh[x])
        bool_fuzz_nom    = i_fuzz_nom > 80
        i_fuzz_prenom    = fuzz.partial_ratio(prenom_cible, prenom_adh[x])
        bool_fuzz_prenom = i_fuzz_prenom > 80
        bool_date_naissance_identique = date_naissance_cible == date_naiss_adh[x]
        
        if (((nom_cible == nom_adh[x] and bool_fuzz_prenom and bool_date_naissance_identique) or (prenom_cible == prenom_adh[x]) and bool_fuzz_nom and bool_date_naissance_identique)):
            w_sheet.write(rx2,num_col_bool_adh, 'ADHERENT')
            if (nom_cible != nom_adh[x]) or (prenom_cible != prenom_adh[x]):
                w_sheet.write(rx2,num_col_bool_adh+1, 'A VERIFIER')
                w_sheet.write(rx2,num_col_bool_adh+2, str(i_fuzz_nom))
                w_sheet.write(rx2,num_col_bool_adh+3, str(i_fuzz_prenom))
			
wb.save("D:\\parser xls\\"+nom_fic_certif+"_adh.xls")


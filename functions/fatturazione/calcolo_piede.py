from utils.variabili import alfabeto

def calcolo_piede(writer, sheet_name, dir_fisso, ldv, istat, fuel, last_col, last_row):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    first_foot_row = last_row + 1

    last_let = alfabeto[last_col]

    # Formattazione
    thick_top = workbook.add_format({'top': 2,})
    thick_bottom = workbook.add_format({'bottom': 2})
    center_format = workbook.add_format({'align': 'center'})
    percent_format = workbook.add_format({'num_format': '0.00%'})

    # Somme delle colonne NOLO e TOTALE
    worksheet.write_formula((first_foot_row), 6, f'=SUM(G2:G{first_foot_row})', thick_top)  # Colonna NOLO (G)
    worksheet.write_formula((first_foot_row), (last_col), f'=SUM({last_let}2:{last_let}{first_foot_row})', thick_top)  # Colonna TOTALE

    # Riga vuota (seconda riga sotto la tabella)
    worksheet.write(first_foot_row + 1, 0, "")

    # Riga "Diritto fisso"
    worksheet.merge_range(first_foot_row + 2, last_col - 7, first_foot_row + 2, last_col - 6, "Diritto fisso")
    worksheet.write(first_foot_row + 2, last_col - 4, dir_fisso, center_format)
    worksheet.write(first_foot_row + 2, last_col - 3, "X", center_format)
    worksheet.write(first_foot_row + 2, last_col - 2, last_row, center_format)
    worksheet.write(first_foot_row + 2, last_col - 1, f"'=", center_format)
    worksheet.write_formula(first_foot_row + 2, last_col, f'={alfabeto[last_col - 4]}{first_foot_row + 3}*{alfabeto[last_col - 2]}{first_foot_row + 3}', workbook.add_format({'num_format': '#,##0.00'}))

    # Riga "Copia LDV"
    worksheet.merge_range(first_foot_row + 3, last_col - 7, first_foot_row + 3, last_col - 6, "Copia LDV")
    worksheet.write(first_foot_row + 3, last_col - 4, ldv, center_format)
    worksheet.write(first_foot_row + 3, last_col - 3, "X", center_format)
    worksheet.write(first_foot_row + 3, last_col - 2, last_row, center_format)
    worksheet.write(first_foot_row + 3, last_col - 1, f"'=", center_format)
    worksheet.write_formula(first_foot_row + 3, last_col, f'={alfabeto[last_col - 4]}{first_foot_row + 4}*{alfabeto[last_col - 2]}{first_foot_row + 4}', workbook.add_format({'num_format': '#,##0.00'}))

    # Riga vuota (quinta riga sotto la tabella)
    worksheet.write(first_foot_row + 4, 0, "")

    # Riga "Variazione ISTAT 2024"
    worksheet.merge_range(first_foot_row + 5, last_col - 7, first_foot_row + 5, last_col - 5, "Variazione ISTAT 2024")
    worksheet.write(first_foot_row + 5, last_col - 3, istat/100, percent_format)
    worksheet.write_formula(first_foot_row + 5, last_col, f'=(SUM({last_let}{first_foot_row + 1}:{last_let}{first_foot_row + 4}))*{alfabeto[last_col - 3]}{first_foot_row + 6}', workbook.add_format({'num_format': '#,##0.00'}))

    # Riga "Supplemento carburante"
    worksheet.merge_range(first_foot_row + 6, last_col - 7, first_foot_row + 6, last_col - 5, "Supplemento carburante")
    worksheet.write(first_foot_row + 6, last_col - 3, fuel/100, percent_format)
    worksheet.write_formula(first_foot_row + 6, last_col, f'=G{first_foot_row + 1}*{alfabeto[last_col - 3]}{first_foot_row + 7}', workbook.add_format({'bottom': 2, 'num_format': '#,##0.00', }))

    # Riga vuota (ottava riga sotto la tabella)
    worksheet.write(first_foot_row + 7, 0, "")

    # Riga totale finale
    worksheet.write(first_foot_row + 8, last_col - 2, "TOTALE", workbook.add_format({'bold': True}))
    worksheet.write(first_foot_row + 8, last_col - 1, "'=", workbook.add_format({'bold': True, 'align': 'center'}))
    worksheet.write_formula(first_foot_row + 8, last_col, f'=SUM({last_let}{first_foot_row + 1}:{last_let}{first_foot_row + 7})', workbook.add_format({'bold': True}))
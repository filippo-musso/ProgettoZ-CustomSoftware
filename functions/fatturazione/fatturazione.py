import pandas as pd
import locale
from utils.caricamento_file import *
from calcolo_piede import calcolo_piede
from utils.variabili import *

def fatturazione():
    locale.setlocale(locale.LC_TIME, 'it_IT.UTF-8')

    # Caricamento dei dati
    disagiate_arco = carica_loc_arco(file_paths["Disagiate Arco"])
    disagiate_susa = carica_dis_susa(file_paths["Disagiate Susa"])
    facchinaggio = carica_loc_arco(file_paths["Facchinaggio"])
    balneari = carica_loc_arco(file_paths["Balneari"])
    impervie = carica_loc_arco(file_paths["Impervie"])
    alta_urb = carica_alta_urb_susa(file_paths["Alta Urbanizzazione"])
    dir_fisso, ldv, istat, fuel = carica_variabili(file_paths["Variabili"])

    # Caricamento del primo file Excel
    df_tariffe = pd.read_excel(file_paths["Tariffe"])
    df_fatt = pd.read_excel(file_paths["Tariffe Aggiunta"], parse_dates=[11], dtype={3: str, 7:str})

    # Inizializzazione della lista per i dati da scrivere nel file di output
    output_data = []

    last_col_index = 7

    tass = False
    espr = False
    tel = False
    dis = False
    fac = False
    bal = False
    imp = False
    urb = False
    mez_pic = False

    def mezzo_piccolo(notes):
        if notes and isinstance(notes, str):
            note = notes.lower()

            if "mezzo piccolo" in note:
                return True
            else: 
                return False

    for index, row in df_tariffe.iterrows():
        num_doc = f"{row.iloc[0]}/{row.iloc[1]}"  # Formattazione numero/lettera

        if row['TASS.'] != 0 and tass == False: 
            tass = True
        if row['ESPR.'] != 0 and espr == False: 
            espr = True
        if row['TEL.'] != 0 and tel == False:
            tel = True

        for _, riga in df_fatt.iterrows():
            data = riga["tm_datadoc"]
            if num_doc == f"{riga.iloc[0]}/{riga.iloc[1]}":  # Confronto con il secondo file
                if riga['tm_coddest'] == 0:
                    cap = riga['an_cap']
                    loc = riga['an_citta']
                    prov = riga['an_prov']
                else:
                    cap = riga['dd_capdest']
                    loc = riga['dd_locdest']
                    prov = riga['dd_prodest']

                if riga['tm_vettor'] == 3:
                    if (cap, loc, prov) in disagiate_arco and dis == False:
                        dis = True
                    if (cap, loc, prov) in facchinaggio and fac == False:
                        fac = True
                    if data.month >= 6 and data.month <= 8:
                        if (cap, loc, prov) in balneari and bal == False: 
                            bal = True
                    if (cap, loc, prov) in impervie and imp == False:
                        imp = True
                    if mez_pic == False:
                        mez_pic = mezzo_piccolo(riga['tm_note'])
                if riga['tm_vettor'] == 946:
                    if (cap, loc, prov) in disagiate_susa and dis == False:
                        dis = True
                    if (cap, loc, prov) in alta_urb and urb == False:
                        urb = True
                break
            
    line = 2

    if tass == True:
        last_col_index += 1
    if espr == True: 
        last_col_index += 1
    if tel == True: 
        last_col_index += 1
    if dis == True:
        last_col_index += 1
    if fac == True:
        last_col_index += 1
    if bal == True:
        last_col_index += 1
    if imp == True:
        last_col_index += 1
    if urb == True:
        last_col_index += 1
    if mez_pic == True:
        last_col_index += 1

    for index, row in df_tariffe.iterrows():
        num_doc = f"{row.iloc[0]}/{row.iloc[1]}"  # Formattazione numero/lettera

        nolo = row.iloc[7]
        arr = row.iloc[5]  # ARR.
        
        data_output = {
            "DOC": num_doc,
            "CLIENTE": row.iloc[2],  # Colonna 2
            "REGIONE": row.iloc[3],  # Colonna 3
            "PESO": row.iloc[4],     # Colonna 4
            "ARR.": row.iloc[5],     # Colonna 5
            "TAR.": row.iloc[6],     # Colonna 6
            "NOLO": round(row.iloc[7], 2)     # Colonna 7
        }

        if tass == True:
            data_output["TASS."] = row["TASS."]
        if espr == True: 
            data_output["ESPR."] = row["ESPR."]
        if tel == True: 
            data_output["TEL."] = row["TEL."]
        if dis == True:
            data_output["DIS."] = 0
        if fac == True:
            data_output["FAC."] = 0
        if bal == True:
            data_output["BAL."] = 0
        if imp == True:
            data_output["IMP."] = 0
        if urb == True:
            data_output["URB."] = 0
        if mez_pic == True: 
            data_output["MEZ.PICC."] = 0

        # Calcolo Addebiti.
        for _, riga in df_fatt.iterrows():
            data  = riga["tm_datadoc"]

            if num_doc == f"{riga.iloc[0]}/{riga.iloc[1]}":  # Confronto con il secondo file

                if riga['tm_coddest'] == 0:
                    cap = riga['an_cap']
                    loc = riga['an_citta']
                    prov = riga['an_prov']
                else:
                    cap = riga['dd_capdest']
                    loc = riga['dd_locdest']
                    prov = riga['dd_prodest']

                if riga['tm_vettor'] == 3:
                    if (cap, loc, prov) in disagiate_arco:
                        data_output["DIS."] = ((arr + 100 - 1) // 100) * 4.50  # Calcolo
                    if (cap, loc, prov) in facchinaggio and fac == True:
                        data_output["FAC."] = ((arr + 100 - 1) // 100) * 6
                    if data.month >= 6 and data.month <= 8:
                        if (cap, loc, prov) in balneari and bal == True:
                            data_output["BAL."] = round((nolo * 0.15), 2)
                    if (cap, loc, prov) in impervie and imp == True:
                        data_output["IMP."] = "IMP"
                    if mezzo_piccolo(riga["tm_note"]) == True:
                        data_output["MEZ.PICC."] = ((arr + 100 - 1) // 100) * 20                

                elif riga['tm_vettor'] == 946:
                    if (cap, loc, prov) in disagiate_susa:
                        data_output["DIS."] = ((arr + 100 - 1) // 100) * 4.50  # Calcolo
                    if (cap, loc, prov) in alta_urb and urb == True: 
                        data_output["URB."] = ((arr + 100 - 1) // 100) * 1.2

                data_output["TOTALE"] = f'=SUM(G{line}:{alfabeto[last_col_index-1]}{line})'

                line += 1

                break  # Esci dal ciclo una volta trovato il documento

        output_data.append(data_output)


    # Creazione del DataFrame per il file di output
    df_output = pd.DataFrame(output_data)

    # Ricava il mese dalla data del documenta della prima riga delfile
    data = df_fatt["tm_datadoc"].iloc[0]
    mese_fatturazione = data.strftime("%B")

    # Aggiunge il nome cliente in basa al deposito, preso dalla prima riga del file
    tm_magaz = df_fatt["tm_magaz"].iloc[0]
    nome_cliente = nomi_clienti[tm_magaz]

    # Salvataggio del file di output
    nome_file_output = f"Trasporto {nome_cliente} {mese_fatturazione}.xlsx"

    last_col = (len(output_data[0]) - 1)
    last_row = len(df_output)  # +1 perché l'indice Excel parte da 1

    # Creazione file di output
    with pd.ExcelWriter(nome_file_output, engine='xlsxwriter') as writer:
        df_output.to_excel(writer, sheet_name='Dettaglio Trasporto', index=False)
        calcolo_piede(writer, 'Dettaglio Trasporto', dir_fisso, ldv, istat, fuel, last_col, last_row)

    input(f"File di output creato: {nome_file_output}")


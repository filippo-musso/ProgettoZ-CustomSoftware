from customtkinter import *
from datetime import datetime

# Configura il tema di customtkinter
set_appearance_mode("light")  # Modalità scura
set_default_color_theme("blue")  # Tema blu predefinito

# Crea la finestra principale
root = CTk()
root.geometry('1200x800')
root.title('Progetto Z - Custom Software')

# Variabili globali
font = ("Arial", 20)
blue = "#2e8ccf"

# Funzione per gestire la visibilità dei frame
def switch_frame(selected_frame):
    # Nasconde tutti i frame (li rimuove dallo schermo)
    for frame in frames.values():
        frame.grid_forget()
    
    # Mostra il frame selezionato
    frames[selected_frame].grid(row=0, column=0, sticky="nsew")

# Configura layout principale
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=7)
root.rowconfigure(0, weight=1)
root.rowconfigure(1, weight=20)

# Barra superiore
top_bar = CTkFrame(root, fg_color=blue, height=30)
top_bar.grid(row=0, column=0, columnspan=2, sticky="nsew")

# Barra a sinistra
left_frame = CTkFrame(root, fg_color="white", width=200)
left_frame.grid(row=1, column=0, sticky="nsew")

# Configura larghezza fissa per la colonna sinistra
left_frame.grid_propagate(False)  # Impedisce l'adattamento automatico

# Frame a destra
right_container = CTkFrame(root)
right_container.grid(row=1, column=1, sticky="nsew")
right_container.columnconfigure(0, weight=1)
right_container.rowconfigure(0, weight=1)

# Dizionario per contenere i frame della parte destra
frames = {
    "home": CTkFrame(right_container, fg_color="white"),
    "Fatturazione": CTkFrame(right_container, fg_color="lightblue"),
    "frame2": CTkFrame(right_container, fg_color="lightgreen"),
    "frame3": CTkFrame(right_container, fg_color="lightpink"),
}

# Contenuti dei frame di destra
CTkLabel(frames["frame2"], text="Contenuto del Frame 2", font=font).pack(pady=20)
CTkLabel(frames["frame3"], text="Contenuto del Frame 3", font=font).pack(pady=20)

# Mostra inizialmente il frame "home"
frames["home"].grid(row=0, column=0, sticky="nsew")

# Barra di navigazione superiore
CTkButton(top_bar, text='Home', fg_color='white', text_color=blue, command=lambda: switch_frame("home")).pack(pady=2, padx=5, side='right')

# Opzioni nella barra sinistra
CTkButton(left_frame, text="Fatturazione", command=lambda: switch_frame("Fatturazione")).pack(pady=10, padx=10, fill="x")
CTkButton(left_frame, text="Mostra Frame 2", command=lambda: switch_frame("frame2")).pack(pady=10, padx=10, fill="x")
CTkButton(left_frame, text="Mostra Frame 3", command=lambda: switch_frame("frame3")).pack(pady=10, padx=10, fill="x")

# Pulsante Esci nella barra sinistra
CTkButton(left_frame, text='Esci', command=root.destroy).pack(side='bottom', pady=10)

# Contenuto del frame Fatturazione
CTkLabel(frames["Fatturazione"], text="FATTURAZIONE PROGETTO Z", font=font).pack(pady=20)

# Funzione universale per formattare la data in un Entry come data
def format_date(event, entry):
    current_text = entry.get()

    # Aggiungi automaticamente i separatori se la lunghezza è corretta
    if len(current_text) == 2 or len(current_text) == 5:
        entry.insert(len(current_text), "/")  # Inserisce il carattere "/"

# Funzione per ottenere la data in formato oggetto date
def get_date(entry):
    date_string = entry.get()
    try:
        # Converti la stringa in un oggetto date
        formatted_date = datetime.strptime(date_string, "%d/%m/%Y").date()
        return formatted_date
    except ValueError:
        print("Formato data non valido. Usa il formato DD/MM/YYYY.")

def fatturazione():
    iniz = get_date(data_iniziale)
    fin = get_date(data_finale)
    print(iniz, fin)

data_iniziale = CTkEntry(frames["Fatturazione"], placeholder_text="Data Iniziale", width=250)
data_finale = CTkEntry(frames["Fatturazione"], placeholder_text="Data Finale", width=250)
data_iniziale.pack(pady=20)
data_finale.pack(pady=20)

# Lega la funzione di formattazione alla digitazione
data_iniziale.bind("<KeyRelease>", lambda event: format_date(event, data_iniziale))
data_finale.bind("<KeyRelease>", lambda event: format_date(event, data_finale))

avvio_fatturazione = CTkButton(frames["Fatturazione"], text="Avvio Generazione Files", command=fatturazione)
avvio_fatturazione.pack(pady=20)

# Avvio del ciclo principale
root.mainloop()


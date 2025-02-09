import xlwings as xw
from datetime import datetime, timedelta
import random
import os
import sys
import tkinter as tk
from tkinter import filedialog

def seleziona_file():
    root = tk.Tk()
    root.withdraw() 
    
    input_file = filedialog.askopenfilename(
        title='Seleziona il file Excel da elaborare',
        filetypes=[('File Excel', '*.xlsx'), ('Tutti i file', '*.*')]
    )
    
    if input_file:
        cartella = os.path.dirname(input_file)
        nome_file = os.path.basename(input_file)
        nome, ext = os.path.splitext(nome_file)
        file_output = os.path.join(cartella, f"{nome}_elaborato{ext}")
        
        return input_file, file_output
    
    return None, None

def genera_orario_fine(data):

    ORA_INIZIO = 9  # PERSONALIZZA
    ORA_FINE = 18    # PERSONALIZZA
    minuti_random = random.randint(0, 59)
    return datetime.combine(
        data,
        datetime.strptime(f"{ORA_INIZIO}:{minuti_random:02d}", "%H:%M").time()
    )

def calcola_durata_minuti(ora_inizio, ora_fine):

    delta = ora_fine - ora_inizio
    return int(delta.total_seconds() / 60)

def modifica_file_prenotazioni(file_input: str, file_output: str) -> None:

    NOME_FOGLIO = "NOME_TUO_FOGLIO"        # PERSONALIZZA
    STATO_INIZIALE = "STATO_INIZIALE"      # PERSONALIZZA
    NUOVO_STATO = "NUOVO_STATO"            # PERSONALIZZA
    ORARIO_INIZIO = "08:30"                # PERSONALIZZA

    app = None
    wb = None
    
    try:
        print("1. Avvio Excel...")
        app = xw.App(visible=False)
        app.display_alerts = False
        
        wb = app.books.open(file_input)
        foglio = wb.sheets[NOME_FOGLIO]
        
        intestazioni = {}
        range_intestazioni = foglio.range('A1').expand('right')
        for cella in range_intestazioni:
            intestazioni[cella.value] = cella.column
        
        # PERSONALIZZA 
        colonne_richieste = [
            'COLONNA_STATO',
            'COLONNA_ORA_INIZIO',
            'COLONNA_ORA_FINE',
            'COLONNA_ARRIVO_PREVISTO',
            'COLONNA_DURATA'
        ]
        
        for nome_colonna in colonne_richieste:
            if nome_colonna not in intestazioni:
                raise ValueError(f"Formato file non valido!")
        
        ultima_riga = foglio.range('A' + str(foglio.cells.last_cell.row)).end('up').row
        righe_elaborate = 0
        
        for riga in range(2, ultima_riga + 1):
            try:
                stato = foglio.cells(riga, intestazioni['COLONNA_STATO']).value
                
                if stato == STATO_INIZIALE:
                    ora_inizio_orig = foglio.cells(riga, intestazioni['COLONNA_ORA_INIZIO']).value
                    
                    if isinstance(ora_inizio_orig, datetime):
                        nuovo_orario_inizio = datetime.combine(
                            ora_inizio_orig.date(),
                            datetime.strptime(ORARIO_INIZIO, "%H:%M").time()
                        )
                        
                        nuovo_orario_fine = genera_orario_fine(ora_inizio_orig.date())
                        durata = calcola_durata_minuti(nuovo_orario_inizio, nuovo_orario_fine)
                        
                        # Aggiorna dati riga
                        foglio.cells(riga, intestazioni['COLONNA_STATO']).value = NUOVO_STATO
                        foglio.cells(riga, intestazioni['COLONNA_ORA_INIZIO']).value = nuovo_orario_inizio
                        foglio.cells(riga, intestazioni['COLONNA_ORA_FINE']).value = nuovo_orario_fine
                        foglio.cells(riga, intestazioni['COLONNA_ARRIVO_PREVISTO']).value = nuovo_orario_fine
                        foglio.cells(riga, intestazioni['COLONNA_DURATA']).value = durata
                        
                        righe_elaborate += 1
            
            except Exception as e:
                print(f"Errore nell'elaborazione della riga {riga}: {str(e)}")
                continue
        
        wb.save(file_output)
        print(f"Elaborate {righe_elaborate} righe")
        
    except Exception as e:
        print(f"Errore critico: {str(e)}")
        raise
    
    finally:
        if wb: wb.close()
        if app: app.quit()

if __name__ == "__main__":
    try:
        file_input, file_output = seleziona_file()
        
        if file_input and file_output:
            modifica_file_prenotazioni(file_input, file_output)
        else:
            print("Nessun file selezionato. Operazione annullata.")
            
    except Exception as e:
        print(f"Errore durante l'esecuzione: {str(e)}")

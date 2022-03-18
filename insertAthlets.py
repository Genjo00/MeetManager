# Codice di Gianluca Figini
# Questo programma è altamente sensibile alle specifiche del computer su cui viene eseguito così come al programma gara.
# Creato per interagire con il programma tecnico del 51mo Meeting di chiasso (11-12.06.22)

import pandas as pd
import subprocess
import time
from directKeys import click, rightClick, PressKey, ReleaseKey, CTRL, N, V

currentLine = 8;
excelFile = pd.read_excel("C:\\Users\\Gianluca Figini\\Documents\\Meeting_Chiasso\\Iscrizioni.xlsx", sheet_name = "Foglio1")
mask = pd.isna(excelFile)
nrLines = excelFile.shape[0] #Cambiare a 3 se si vuole testare il funzionamento 

FIRST_NAME = 1
LAST_NAME = 0
GENDER = 2
DATE_OF_BIRTH = 3
EVENT_NR = 5
ENTRY_TIME = 7
LICENCE_NR = 4


def copy2clip(txt):
    cmd='echo '+txt.strip()+'|clip'
    return subprocess.check_call(cmd, shell=True)

def newAthlet() :
    #Apri la finestra nuovo atleta
    PressKey(CTRL)
    PressKey(N)
    ReleaseKey(CTRL)
    ReleaseKey(N)
    time.sleep(0.5)
    #Copia da excel e incolla il cognome nel campo
    copy2clip(excelFile.iloc[currentLine, LAST_NAME])
    click(260, 108)
    click(260, 108)
    PressKey(CTRL)
    PressKey(V)
    ReleaseKey(CTRL)
    ReleaseKey(V)
    time.sleep(0.5)
    #Copia da excel e incolla il nome nel campo
    copy2clip(excelFile.iloc[currentLine, FIRST_NAME])
    click(260, 132)
    click(260, 132)
    PressKey(CTRL)
    PressKey(V)
    ReleaseKey(CTRL)
    ReleaseKey(V)
    time.sleep(0.5)
    #Seleziona il genere
    click(260, 175)
    click(260, 175)
    time.sleep(0.5)
    if excelFile.iloc[currentLine, GENDER] == "M":
        click(260, 197)
    elif excelFile.iloc[currentLine, GENDER] == "F":
        click(260, 211)
    else:
        print("{}, {}, genere sconosciuto, inserire gare manualmente".format(excelFile.iloc[currentLine, LAST_NAME], excelFile.iloc[currentLine, FIRST_NAME]))
    time.sleep(0.5)
    #Copia da excel e incolla la data di nascita nel campo
    copy2clip(str(excelFile.iloc[currentLine, DATE_OF_BIRTH]))
    click(260, 193)
    click(260, 193)
    PressKey(CTRL)
    PressKey(V)
    ReleaseKey(CTRL)
    ReleaseKey(V)
    time.sleep(0.5)
    #Copia da excel e incolla il numero di licenza nel campo
    if not mask.iloc[currentLine, LICENCE_NR]:
        copy2clip(str(excelFile.iloc[currentLine, LICENCE_NR]))
        click(260, 341)
        click(260, 341)
        PressKey(CTRL)
        PressKey(V)
        ReleaseKey(CTRL)
        ReleaseKey(V)
        time.sleep(0.5)
    #Click su close
    click(97, 540)

def newEntry(line):
    #Localizza la posizione della gara sullo schermo
    gara = int(excelFile.iloc[currentLine, EVENT_NR])
    yob = int(str(excelFile.iloc[line, DATE_OF_BIRTH])[-4:])
    verShift = (gara -1)/2
    if gara > 20:
        print("{} {}, numero gara {} non valido".format(excelFile.iloc[line, LAST_NAME], excelFile.iloc[line, FIRST_NAME], excelFile.iloc[currentLine, EVENT_NR]))
        return
    if yob > 2008:
        verShift -= 1
    if gara > 12 and yob < 2009:
        verShift -= 2
    #Controllo correttezza
    if (gara % 2) ^ (excelFile.iloc[line, GENDER] =="M"):
        print("{} {}, gara {}, discrepanza genere".format(excelFile.iloc[line, LAST_NAME], excelFile.iloc[line, FIRST_NAME], excelFile.iloc[currentLine, EVENT_NR]))
    #Click destro sulla gara
    rightClick(186, int(690 + 20 * verShift))
    time.sleep(0.5)
    #Click su add entry
    if verShift == 0:
        click(254, 821)
    else:
        click(254, int(578 + 20 * (verShift-1)))
    time.sleep(0.5)
    #Tempo di partenza (se non NT)
    if not mask.iloc[currentLine, ENTRY_TIME]:
        click(1375, 702)
        copy2clip(str(excelFile.iloc[currentLine, ENTRY_TIME]))
        PressKey(CTRL)
        PressKey(V)
        ReleaseKey(CTRL)
        ReleaseKey(V)

    
    


if __name__ == "__main__":
    time.sleep(5)
    athletes = 0
    entries = 0
    #finchè c'è qualcosa da leggere nella tabella
    while currentLine < nrLines:
        #Crea nuovo atleta
        newAthlet()
        athletes += 1
        firstRun = True
        lastFullLine = currentLine
        #se la colonna cognome è vuota oppure questa è la prima iscrizione di un atleta
        while currentLine < nrLines and (mask.iloc[currentLine, LAST_NAME] or firstRun):
            #Inserisci nuova iscrizione
            newEntry(lastFullLine)
            entries += 1
            firstRun = False
            currentLine += 1
            time.sleep(0.5)
        #Salta linee vuote
            while currentLine < nrLines and mask.iloc[currentLine, EVENT_NR]:
                currentLine += 1
        while currentLine < nrLines and mask.iloc[currentLine, EVENT_NR]:
            currentLine += 1
    print("Atleti inseriti: {}\nIscrizioni inserite: {}".format(athletes, entries))
  
     

# Attestati Generator

## Descrizione
**Attestati Generator** è una web app e un plugin per WordPress che permette di generare massivamente attestati, certificati o documenti personalizzati partendo da un template Word (`.docx`) e un file di dati (`.csv`).

All'interno della cartella PLUGINWP trovi il plugin pronto da installare per WP.
Il plugin offre una semplice interfaccia di amministrazione dove è possibile caricare il file dati e il template. Il sistema genererà automaticamente un file per ogni riga del CSV, sostituendo i segnaposto nel template con i dati corrispondenti.

## Funzionalità Principali
- **Upload CSV e Template**: Caricamento semplice tramite interfaccia Drag & Drop o selezione file.
- **Supporto Segnaposto**: Sostituzione automatica dei tag nel file Word (es. `${NOME}`, `${COGNOME}`) con i valori delle colonne del CSV.
- **Output Flessibile**:
  - **DOCX**: Genera file Word modificabili. Include automaticamente script (PowerShell e Batch) nel pacchetto ZIP per convertire massivamente i file in PDF in locale.
  - **PDF**: Generazione diretta in PDF (tramite librerie PHP o conversione HTML fallback).
- **Archivio ZIP**: Tutti i file generati vengono scaricati in un unico comodo archivio ZIP.
- **Gestione CSV Intelligente**: Rilevamento automatico del delimitatore (virgola o punto e virgola) e pulizia BOM.

## Requisiti
- WordPress 5.0 o superiore
- PHP 7.4 o superiore
- Estensioni PHP: `zip`, `xml`, `gd` (per la gestione immagini/pdf)

## Installazione per WP
1. Scarica il pacchetto del plugin o clona la repository nella cartella `wp-content/plugins/`.
2. Assicurati di eseguire `composer install` nella root del plugin per installare le dipendenze necessarie (`phpoffice/phpword`, `dompdf/dompdf`, ecc.).
3. Attiva il plugin dal pannello di amministrazione di WordPress.
4. Troverai la voce **"Attestati"** nel menu laterale della dashboard.

## Utilizzo

### 1. Preparazione del Template (.docx)
Crea un file Word e inserisci i segnaposto corrispondenti alle intestazioni del tuo file CSV.
Esempio:
> "Si certifica che **${NOME} ${COGNOME}** ha frequentato il corso..."

I segnaposto devono corrispondere esattamente ai nomi delle colonne nel CSV (case-insensitive).

### 2. Preparazione del CSV
Il file CSV deve contenere una riga di intestazione con i nomi dei campi.
Esempio:
```csv
NOME;COGNOME;DATA;CORSO
Mario;Rossi;01/01/2024;Corso Base
Luigi;Verdi;02/01/2024;Corso Avanzato
```

### 3. Generazione
1. Vai su **Attestati** nel menu admin.
2. Carica il file `.csv`.
3. Carica il file template `.docx`.
4. Seleziona il formato di output desiderato (DOCX o PDF).
5. Clicca su **Genera Attestati**.
6. Attendi il processo e scarica il file ZIP.

## Conversione Locale (Opzionale)
Se scegli di generare file in formato **DOCX**, all'interno dello ZIP troverai due script di utilità:
- `converti-docx-in-pdf.ps1` (PowerShell)
- `CONVERTI_IN_PDF.bat` (Batch)

Questi script permettono di convertire rapidamente tutti i file Word in PDF direttamente sul tuo computer Windows, sfruttando Microsoft Word installato per garantire la massima fedeltà di layout.

**Come usare lo script:**
1. Estrai lo zip in una cartella.
2. Fai doppio clic su `CONVERTI_IN_PDF.bat`.
3. Lo script convertirà tutti i `.docx` presenti nella cartella in `.pdf`.

## Autore
**Dario Molino**

# Attestati Generator - Plugin WordPress

## Requisiti
- WordPress 5.0+
- PHP 7.4+
- Estensione PHP ZipArchive

## Installazione

### 1. Installa le dipendenze Composer
Apri il terminale nella cartella del plugin ed esegui:

```bash
cd wp-content/plugins/attestati-generator
composer install --no-dev
```

> **Nota:** Il `composer.json` include `"platform": {"php": "7.4.33"}` che forza la risoluzione delle dipendenze per PHP 7.4, anche se stai usando una versione più recente.

### 2. Attiva il plugin
Vai in **WordPress Admin → Plugin** e attiva **Attestati Generator**.

### 3. Usa il plugin
Trovi la voce **"Attestati"** (icona 🏆) nel menu laterale dell'admin.

## Utilizzo
1. Carica un file **CSV** (separato da `;`) con le colonne: `COGNOME`, `NOME`, `DATA`, `TITOLOEVENTO`
2. Carica un template **Word (.docx)** con i placeholder: `${NOME}`, `${COGNOME}`, `${DATA}`, `${TITOLOEVENTO}`
3. Scegli il formato di output (DOCX o PDF)
4. Clicca "Genera Attestati" e scarica lo ZIP

## Conversione DOCX → PDF
Lo ZIP generato in formato DOCX include lo script `CONVERTI_IN_PDF.bat` per convertire automaticamente tutti i DOCX in PDF usando Microsoft Word.

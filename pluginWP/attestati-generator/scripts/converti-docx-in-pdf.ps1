# ============================================
# Converti tutti i DOCX in PDF usando Word
# ============================================
# Uso: 
#   Trascinare i file DOCX sullo script, oppure eseguire:
#   .\converti-docx-in-pdf.ps1 "C:\cartella\con\docx"
#   .\converti-docx-in-pdf.ps1  (usa la stessa cartella dello script)
# ============================================

param(
    [string]$InputPath = ""
)

# Determina la cartella di input
if ($InputPath -eq "") {
    $InputPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}

# Se è stato passato un file singolo, usa la sua cartella
if (Test-Path $InputPath -PathType Leaf) {
    $InputPath = Split-Path -Parent $InputPath
}

# Trova tutti i file DOCX
$docxFiles = Get-ChildItem -Path $InputPath -Filter "*.docx" -File
if ($docxFiles.Count -eq 0) {
    Write-Host ""
    Write-Host "  Nessun file .docx trovato in: $InputPath" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Premi INVIO per chiudere"
    exit
}

Write-Host ""
Write-Host "  =========================================" -ForegroundColor Cyan
Write-Host "  Convertitore DOCX -> PDF" -ForegroundColor Cyan
Write-Host "  =========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Cartella: $InputPath" -ForegroundColor Gray
Write-Host "  File trovati: $($docxFiles.Count)" -ForegroundColor Gray
Write-Host ""

# Avvia Word
Write-Host "  Avvio Microsoft Word..." -ForegroundColor Yellow
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
}
catch {
    Write-Host "  ERRORE: Microsoft Word non trovato!" -ForegroundColor Red
    Write-Host "  Assicurati di avere Word installato." -ForegroundColor Red
    Write-Host ""
    Read-Host "Premi INVIO per chiudere"
    exit 1
}

$converted = 0
$errors = 0

foreach ($docx in $docxFiles) {
    $pdfPath = [System.IO.Path]::ChangeExtension($docx.FullName, ".pdf")
    $fileName = $docx.Name
    
    Write-Host "  [$($converted + $errors + 1)/$($docxFiles.Count)] Conversione: $fileName" -NoNewline
    
    try {
        $doc = $word.Documents.Open($docx.FullName)
        # WdSaveFormat: wdFormatPDF = 17
        $doc.SaveAs([ref]$pdfPath, [ref]17)
        $doc.Close([ref]$false)
        $converted++
        Write-Host " -> OK" -ForegroundColor Green
    }
    catch {
        $errors++
        Write-Host " -> ERRORE" -ForegroundColor Red
        Write-Host "     $($_.Exception.Message)" -ForegroundColor DarkRed
    }
}

# Chiudi Word
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

Write-Host ""
Write-Host "  =========================================" -ForegroundColor Cyan
Write-Host "  Completato!" -ForegroundColor Green
Write-Host "  Convertiti: $converted   Errori: $errors" -ForegroundColor $(if ($errors -gt 0) { "Yellow" } else { "Green" })
Write-Host "  =========================================" -ForegroundColor Cyan
Write-Host ""
Read-Host "Premi INVIO per chiudere"

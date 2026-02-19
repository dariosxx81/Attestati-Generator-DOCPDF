<?php
/**
 * Generatore Attestati - Logica Core (PHP)
 * 
 * Supporta template HTML (→ PDF via dompdf) e DOCX (→ DOCX o PDF via PHPWord)
 */

require_once __DIR__ . '/vendor/autoload.php';

use Dompdf\Dompdf;
use Dompdf\Options;
use PhpOffice\PhpWord\TemplateProcessor;

/**
 * Legge un file CSV e restituisce un array di record
 */
function parseCsv($csvPath)
{
    $records = [];

    if (!file_exists($csvPath)) {
        throw new Exception("File CSV non trovato: $csvPath");
    }

    $handle = fopen($csvPath, 'r');
    if ($handle === false) {
        throw new Exception("Impossibile aprire il file CSV");
    }

    // Rileva il delimitatore (punto e virgola o virgola)
    $firstLine = fgets($handle);
    rewind($handle);
    $delimiter = (substr_count($firstLine, ';') > substr_count($firstLine, ',')) ? ';' : ',';

    // Leggi l'header
    $header = fgetcsv($handle, 0, $delimiter);
    if ($header === false) {
        fclose($handle);
        throw new Exception("Il file CSV è vuoto");
    }

    // Normalizza le chiavi dell'header (trim + uppercase)
    $header = array_map(function ($col) {
        return strtoupper(trim($col));
    }, $header);

    // Rimuovi BOM se presente
    if (isset($header[0])) {
        $header[0] = preg_replace('/^\xEF\xBB\xBF/', '', $header[0]);
    }

    // Leggi i record  
    while (($row = fgetcsv($handle, 0, $delimiter)) !== false) {
        if (count($row) === count($header)) {
            $record = [];
            foreach ($header as $i => $key) {
                $record[$key] = trim($row[$i]);
            }
            $records[] = $record;
        }
    }

    fclose($handle);
    return $records;
}

/**
 * Sostituisce i placeholder nel template HTML con i dati
 */
function processHtmlTemplate($templatePath, $data)
{
    $html = file_get_contents($templatePath);

    if ($html === false) {
        throw new Exception("Impossibile leggere il template: $templatePath");
    }

    // Sostituisci i placeholder {{KEY}} con i valori
    foreach ($data as $key => $value) {
        $html = str_replace('{{' . $key . '}}', htmlspecialchars($value), $html);
    }

    // Sostituisci anche {KEY} (singole parentesi)
    foreach ($data as $key => $value) {
        $html = preg_replace('/(?<!\{)\{' . preg_quote($key, '/') . '\}(?!\})/', htmlspecialchars($value), $html);
    }

    return $html;
}

/**
 * Genera un DOCX da un template Word usando PHPWord TemplateProcessor
 * Nel template Word i placeholder devono essere scritti come ${NOME}, ${COGNOME}, ecc.
 */
function processDocxTemplate($templatePath, $data, $outputPath)
{
    $templateProcessor = new TemplateProcessor($templatePath);

    // Sostituisci i placeholder ${KEY}
    foreach ($data as $key => $value) {
        $templateProcessor->setValue($key, $value);
        $templateProcessor->setValue(strtolower($key), $value);
    }

    $templateProcessor->saveAs($outputPath);
    return $outputPath;
}

/**
 * Genera un PDF da HTML usando dompdf
 */
function generatePdf($html, $outputPath)
{
    $options = new Options();
    $options->set('isHtml5ParserEnabled', true);
    $options->set('isRemoteEnabled', true);
    $options->set('defaultFont', 'sans-serif');
    $options->set('isPhpEnabled', false);

    $dompdf = new Dompdf($options);
    $dompdf->loadHtml($html);
    $dompdf->setPaper('A4', 'landscape');
    $dompdf->render();

    $pdfContent = $dompdf->output();
    file_put_contents($outputPath, $pdfContent);

    return $outputPath;
}

/**
 * Converte un DOCX in PDF usando PHPWord + dompdf
 * Nota: la formattazione potrebbe non essere identica al DOCX originale
 */
function convertDocxToPdf($docxPath, $pdfPath)
{

    // ---- METODO 2: Fallback PHPWord → dompdf ----
    \PhpOffice\PhpWord\Settings::setPdfRendererName('DomPDF');
    \PhpOffice\PhpWord\Settings::setPdfRendererPath(realpath(__DIR__ . '/vendor/dompdf/dompdf'));

    $phpWord = \PhpOffice\PhpWord\IOFactory::load($docxPath);

    foreach ($phpWord->getSections() as $section) {
        $section->getStyle()->setOrientation('landscape');
        $section->getStyle()->setPaperSize('A4');
    }

    // Sospendi error handler - dompdf genera warning DOMXPath recuperabili
    set_error_handler(function ($errno, $errstr) {
        if (strpos($errstr, 'DOMXPath') !== false || $errno === E_DEPRECATED || $errno === E_WARNING) {
            return true;
        }
        return false;
    });

    try {
        $pdfWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'PDF');
        $pdfWriter->save($pdfPath);
    } catch (\Exception $e) {
        $htmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'HTML');
        ob_start();
        $htmlWriter->save('php://output');
        $html = ob_get_clean();

        if (strpos($html, '<meta charset') === false) {
            $html = '<meta charset="UTF-8">' . $html;
        }
        $html = preg_replace('/mso-[^;]+;?/', '', $html);
        $html = preg_replace('/<o:p>.*?<\/o:p>/s', '', $html);

        generatePdf($html, $pdfPath);
    }

    restore_error_handler();

    return $pdfPath;
}

/**
 * Genera il nome del file attestato
 */
function generateFileName($data, $extension = '.pdf')
{
    $nome = isset($data['NOME']) ? trim($data['NOME']) : '';
    $cognome = isset($data['COGNOME']) ? trim($data['COGNOME']) : '';

    if (empty($nome) && empty($cognome)) {
        return 'Attestato_' . time() . $extension;
    }

    $fileName = 'Attestato_' . $nome . '_' . $cognome;
    $fileName = preg_replace('/\s+/', '_', $fileName);
    return $fileName . $extension;
}

/**
 * Crea un file ZIP con tutti i file generati
 */
function createZip($files, $zipPath)
{
    $zip = new ZipArchive();

    if ($zip->open($zipPath, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
        throw new Exception("Impossibile creare il file ZIP");
    }

    foreach ($files as $file) {
        if (file_exists($file['path'])) {
            $zip->addFile($file['path'], $file['name']);
        }
    }

    $zip->close();
    return $zipPath;
}

/**
 * Funzione principale: genera tutti gli attestati
 * 
 * @param string $csvPath    Percorso del file CSV
 * @param string $templatePath  Percorso del template (HTML o DOCX)
 * @param string $format     Formato output desiderato: 'pdf' o 'docx'
 */
function generateCertificates($csvPath, $templatePath, $format = 'pdf')
{
    $outputDir = __DIR__ . '/output';

    // Assicurati che la directory output esista
    if (!is_dir($outputDir)) {
        mkdir($outputDir, 0755, true);
    }

    // 1. Leggi i dati dal CSV
    $records = parseCsv($csvPath);

    if (empty($records)) {
        throw new Exception("Il file CSV è vuoto o non contiene dati validi");
    }

    $generatedFiles = [];
    $ext = strtolower(pathinfo($templatePath, PATHINFO_EXTENSION));
    $isHtmlTemplate = ($ext === 'html' || $ext === 'htm');
    $isDocxTemplate = ($ext === 'docx');

    // Normalizza formato richiesto
    $outputFormat = strtolower($format);

    // 2. Genera un attestato per ogni record
    foreach ($records as $i => $record) {
        $fileName = '';
        $outputPath = '';

        if ($isHtmlTemplate) {
            // Template HTML → produce sempre PDF
            $html = processHtmlTemplate($templatePath, $record);
            $fileName = generateFileName($record, '.pdf');
            $outputPath = $outputDir . '/' . $fileName;
            generatePdf($html, $outputPath);

        } elseif ($isDocxTemplate && $outputFormat === 'docx') {
            // DOCX template → output DOCX 
            $fileName = generateFileName($record, '.docx');
            $outputPath = $outputDir . '/' . $fileName;
            processDocxTemplate($templatePath, $record, $outputPath);

        } elseif ($isDocxTemplate && $outputFormat === 'pdf') {
            // DOCX template → output PDF (PHPWord → HTML → dompdf)
            $tempDocx = $outputDir . '/temp_' . uniqid() . '.docx';
            processDocxTemplate($templatePath, $record, $tempDocx);

            $fileName = generateFileName($record, '.pdf');
            $outputPath = $outputDir . '/' . $fileName;
            convertDocxToPdf($tempDocx, $outputPath);

            // Elimina DOCX temporaneo
            if (file_exists($tempDocx)) {
                unlink($tempDocx);
            }

        } else {
            throw new Exception("Formato template non supportato: .$ext. Usa .html o .docx");
        }

        $generatedFiles[] = [
            'path' => $outputPath,
            'name' => $fileName
        ];
    }

    // 3. Crea un file ZIP con tutti gli attestati
    $timestamp = date('Y-m-d_H-i-s');
    $zipFileName = "Attestati_{$timestamp}.zip";
    $zipPath = $outputDir . '/' . $zipFileName;

    // Se output è DOCX, includi lo script per convertire in PDF
    if ($isDocxTemplate && $outputFormat === 'docx') {
        $scriptPath = __DIR__ . '/converti-docx-in-pdf.ps1';
        $batPath = __DIR__ . '/CONVERTI_IN_PDF.bat';
        if (file_exists($scriptPath)) {
            $generatedFiles[] = [
                'path' => $scriptPath,
                'name' => 'converti-docx-in-pdf.ps1',
                'keep' => true
            ];
        }
        if (file_exists($batPath)) {
            $generatedFiles[] = [
                'path' => $batPath,
                'name' => 'CONVERTI_IN_PDF.bat',
                'keep' => true
            ];
        }
    }

    createZip($generatedFiles, $zipPath);

    // 4. Elimina i file singoli dopo aver creato lo ZIP (tranne lo script)
    foreach ($generatedFiles as $file) {
        if (!empty($file['keep']))
            continue;
        if (file_exists($file['path'])) {
            unlink($file['path']);
        }
    }

    // Formato effettivo dell'output
    $actualFormat = (!empty($generatedFiles) && pathinfo($generatedFiles[0]['name'], PATHINFO_EXTENSION) === 'pdf') ? 'PDF' : 'DOCX';

    return [
        'success' => true,
        'count' => count($records),
        'zipFile' => $zipFileName,
        'certificates' => array_map(function ($f) {
            return $f['name'];
        }, $generatedFiles),
        'format' => $actualFormat
    ];
}

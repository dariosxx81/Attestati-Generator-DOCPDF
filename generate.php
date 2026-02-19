<?php
/**
 * Endpoint per la generazione degli attestati
 * Riceve CSV + template via POST, genera attestati e restituisce JSON
 */

header('Content-Type: application/json; charset=utf-8');

// Gestione errori
set_error_handler(function ($errno, $errstr, $errfile, $errline) {
    throw new ErrorException($errstr, 0, $errno, $errfile, $errline);
});

try {
    require_once __DIR__ . '/generator.php';

    // Verifica che i file siano stati caricati
    if (!isset($_FILES['csv']) || !isset($_FILES['template'])) {
        throw new Exception('È necessario caricare sia il file CSV che il template');
    }

    // Verifica errori upload
    if ($_FILES['csv']['error'] !== UPLOAD_ERR_OK) {
        throw new Exception('Errore nel caricamento del file CSV');
    }
    if ($_FILES['template']['error'] !== UPLOAD_ERR_OK) {
        throw new Exception('Errore nel caricamento del template');
    }

    // Prepara directory uploads
    $uploadDir = __DIR__ . '/uploads';
    if (!is_dir($uploadDir)) {
        mkdir($uploadDir, 0755, true);
    }

    // Salva i file caricati
    $csvPath = $uploadDir . '/' . uniqid() . '_' . basename($_FILES['csv']['name']);
    $templatePath = $uploadDir . '/' . uniqid() . '_' . basename($_FILES['template']['name']);

    move_uploaded_file($_FILES['csv']['tmp_name'], $csvPath);
    move_uploaded_file($_FILES['template']['tmp_name'], $templatePath);

    // Leggi formato selezionato
    $format = isset($_POST['format']) ? $_POST['format'] : 'pdf';

    // Genera attestati
    $result = generateCertificates($csvPath, $templatePath, $format);

    // Pulisci file temporanei
    if (file_exists($csvPath))
        unlink($csvPath);
    if (file_exists($templatePath))
        unlink($templatePath);

    // Rispondi con JSON
    echo json_encode($result);

} catch (Exception $e) {
    http_response_code(500);
    echo json_encode([
        'error' => $e->getMessage()
    ]);
}

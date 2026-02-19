<?php
/**
 * Endpoint per il download dei file generati
 */

$filename = isset($_GET['file']) ? basename($_GET['file']) : '';
$from = isset($_GET['from']) ? $_GET['from'] : 'output';

if (empty($filename)) {
    http_response_code(400);
    echo 'Nome file mancante';
    exit;
}

// Determina la directory sorgente
if ($from === 'examples') {
    $filePath = __DIR__ . '/examples/' . $filename;
} else {
    $filePath = __DIR__ . '/output/' . $filename;
}

if (!file_exists($filePath)) {
    http_response_code(404);
    echo 'File non trovato';
    exit;
}

// Determina il content type
$ext = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
$contentTypes = [
    'zip' => 'application/zip',
    'pdf' => 'application/pdf',
    'docx' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'html' => 'text/html'
];

$contentType = isset($contentTypes[$ext]) ? $contentTypes[$ext] : 'application/octet-stream';

header('Content-Type: ' . $contentType);
header('Content-Disposition: attachment; filename="' . $filename . '"');
header('Content-Length: ' . filesize($filePath));
header('Cache-Control: no-cache, must-revalidate');

readfile($filePath);
exit;

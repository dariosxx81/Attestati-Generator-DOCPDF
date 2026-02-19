<?php
/**
 * Classe core per la generazione degli attestati
 */

if (!defined('ABSPATH')) {
    exit;
}

use PhpOffice\PhpWord\TemplateProcessor;

class AG_Certificate_Generator
{
    private $upload_dir;
    private $output_dir;

    public function __construct()
    {
        $this->upload_dir = AG_PLUGIN_DIR . 'uploads';
        $this->output_dir = AG_PLUGIN_DIR . 'output';

        if (!is_dir($this->upload_dir)) {
            wp_mkdir_p($this->upload_dir);
        }
        if (!is_dir($this->output_dir)) {
            wp_mkdir_p($this->output_dir);
        }
    }

    public function get_upload_dir()
    {
        return $this->upload_dir;
    }

    public function get_output_dir()
    {
        return $this->output_dir;
    }

    /**
     * Genera tutti gli attestati
     */
    public function generate($csv_path, $template_path, $format = 'docx')
    {
        $records = $this->parse_csv($csv_path);

        if (empty($records)) {
            throw new Exception('Il file CSV è vuoto o non contiene dati validi');
        }

        $ext = strtolower(pathinfo($template_path, PATHINFO_EXTENSION));

        if ($ext !== 'docx') {
            throw new Exception('Formato template non supportato. Usa un file .docx');
        }

        $generated_files = [];
        $output_format = strtolower($format);

        foreach ($records as $record) {
            if ($output_format === 'docx') {
                $file_name = $this->generate_file_name($record, '.docx');
                $output_path = $this->output_dir . '/' . $file_name;
                $this->process_docx_template($template_path, $record, $output_path);
            } else {
                // DOCX → PDF via PHPWord + dompdf
                $temp_docx = $this->output_dir . '/temp_' . uniqid() . '.docx';
                $this->process_docx_template($template_path, $record, $temp_docx);

                $file_name = $this->generate_file_name($record, '.pdf');
                $output_path = $this->output_dir . '/' . $file_name;
                $this->convert_docx_to_pdf($temp_docx, $output_path);

                @unlink($temp_docx);
            }

            $generated_files[] = [
                'path' => $output_path,
                'name' => $file_name,
            ];
        }

        // Crea ZIP
        $timestamp = date('Y-m-d_H-i-s');
        $zip_file_name = "Attestati_{$timestamp}.zip";
        $zip_path = $this->output_dir . '/' . $zip_file_name;

        // Se DOCX, includi script di conversione PDF
        if ($output_format === 'docx') {
            $script_path = AG_PLUGIN_DIR . 'scripts/converti-docx-in-pdf.ps1';
            $bat_path = AG_PLUGIN_DIR . 'scripts/CONVERTI_IN_PDF.bat';

            if (file_exists($script_path)) {
                $generated_files[] = [
                    'path' => $script_path,
                    'name' => 'converti-docx-in-pdf.ps1',
                    'keep' => true,
                ];
            }
            if (file_exists($bat_path)) {
                $generated_files[] = [
                    'path' => $bat_path,
                    'name' => 'CONVERTI_IN_PDF.bat',
                    'keep' => true,
                ];
            }
        }

        $this->create_zip($generated_files, $zip_path);

        // Pulizia file singoli (tranne script)
        foreach ($generated_files as $file) {
            if (!empty($file['keep']))
                continue;
            @unlink($file['path']);
        }

        $actual_format = (!empty($generated_files) && pathinfo($generated_files[0]['name'], PATHINFO_EXTENSION) === 'pdf') ? 'PDF' : 'DOCX';

        return [
            'count' => count($records),
            'zipFile' => $zip_file_name,
            'certificates' => array_map(fn($f) => $f['name'], $generated_files),
            'format' => $actual_format,
        ];
    }

    /**
     * Parsa il CSV
     */
    private function parse_csv($csv_path)
    {
        $records = [];

        if (!file_exists($csv_path)) {
            throw new Exception("File CSV non trovato: $csv_path");
        }

        $handle = fopen($csv_path, 'r');
        if ($handle === false) {
            throw new Exception("Impossibile aprire il file CSV");
        }

        // Rileva il delimitatore
        $first_line = fgets($handle);
        rewind($handle);
        $delimiter = (substr_count($first_line, ';') > substr_count($first_line, ',')) ? ';' : ',';

        // Header
        $header = fgetcsv($handle, 0, $delimiter);
        if ($header === false) {
            fclose($handle);
            throw new Exception("Il file CSV è vuoto");
        }

        // Normalizza header
        $header = array_map(function ($col) {
            return strtoupper(trim($col));
        }, $header);

        // Rimuovi BOM
        if (isset($header[0])) {
            $header[0] = preg_replace('/^\xEF\xBB\xBF/', '', $header[0]);
        }

        // Leggi record
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
     * Processa template DOCX
     */
    private function process_docx_template($template_path, $data, $output_path)
    {
        $template = new TemplateProcessor($template_path);

        foreach ($data as $key => $value) {
            $template->setValue($key, $value);
            $template->setValue(strtolower($key), $value);
        }

        $template->saveAs($output_path);
    }

    /**
     * Converte DOCX in PDF
     */
    private function convert_docx_to_pdf($docx_path, $pdf_path)
    {
        \PhpOffice\PhpWord\Settings::setPdfRendererName('DomPDF');
        \PhpOffice\PhpWord\Settings::setPdfRendererPath(AG_PLUGIN_DIR . 'vendor/dompdf/dompdf');

        $phpWord = \PhpOffice\PhpWord\IOFactory::load($docx_path);

        foreach ($phpWord->getSections() as $section) {
            $section->getStyle()->setOrientation('landscape');
            $section->getStyle()->setPaperSize('A4');
        }

        set_error_handler(function ($errno, $errstr) {
            if (strpos($errstr, 'DOMXPath') !== false || $errno === E_DEPRECATED || $errno === E_WARNING) {
                return true;
            }
            return false;
        });

        try {
            $writer = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'PDF');
            $writer->save($pdf_path);
        } catch (\Exception $e) {
            // Fallback: DOCX → HTML → PDF
            $htmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'HTML');
            ob_start();
            $htmlWriter->save('php://output');
            $html = ob_get_clean();

            if (strpos($html, '<meta charset') === false) {
                $html = '<meta charset="UTF-8">' . $html;
            }
            $html = preg_replace('/mso-[^;]+;?/', '', $html);
            $html = preg_replace('/<o:p>.*?<\/o:p>/s', '', $html);

            $this->generate_pdf_from_html($html, $pdf_path);
        }

        restore_error_handler();
    }

    /**
     * Genera PDF da HTML con dompdf
     */
    private function generate_pdf_from_html($html, $output_path)
    {
        $options = new \Dompdf\Options();
        $options->set('isHtml5ParserEnabled', true);
        $options->set('isRemoteEnabled', true);
        $options->set('defaultFont', 'sans-serif');
        $options->set('isPhpEnabled', false);

        $dompdf = new \Dompdf\Dompdf($options);
        $dompdf->loadHtml($html);
        $dompdf->setPaper('A4', 'landscape');
        $dompdf->render();

        file_put_contents($output_path, $dompdf->output());
    }

    /**
     * Genera nome file
     */
    private function generate_file_name($data, $extension = '.pdf')
    {
        $nome = isset($data['NOME']) ? trim($data['NOME']) : '';
        $cognome = isset($data['COGNOME']) ? trim($data['COGNOME']) : '';

        if (empty($nome) && empty($cognome)) {
            return 'Attestato_' . time() . $extension;
        }

        $file_name = 'Attestato_' . $nome . '_' . $cognome;
        $file_name = preg_replace('/\s+/', '_', $file_name);
        return $file_name . $extension;
    }

    /**
     * Crea ZIP
     */
    private function create_zip($files, $zip_path)
    {
        $zip = new ZipArchive();

        if ($zip->open($zip_path, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
            throw new Exception("Impossibile creare il file ZIP");
        }

        foreach ($files as $file) {
            if (file_exists($file['path'])) {
                $zip->addFile($file['path'], $file['name']);
            }
        }

        $zip->close();
    }
}

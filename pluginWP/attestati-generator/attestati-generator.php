<?php
/**
 * Plugin Name: Attestati Generator
 * Plugin URI:  https://example.com/attestati-generator
 * Description: Genera attestati personalizzati da file CSV e template Word (.docx).
 * Version:     1.0
 * Author:      Dario Molino
 * Text Domain: attestati-generator
 * Requires PHP: 7.4
 */

if (!defined('ABSPATH')) {
    exit;
}

define('AG_PLUGIN_DIR', plugin_dir_path(__FILE__));
define('AG_PLUGIN_URL', plugin_dir_url(__FILE__));
define('AG_VERSION', '1.0');

// Autoload Composer dependencies
$autoload = AG_PLUGIN_DIR . 'vendor/autoload.php';
if (file_exists($autoload)) {
    require_once $autoload;
}

// Include generator class
require_once AG_PLUGIN_DIR . 'includes/class-generator.php';

/**
 * Class Attestati_Generator_Plugin
 */
class Attestati_Generator_Plugin
{
    private static $instance = null;

    public static function instance()
    {
        if (is_null(self::$instance)) {
            self::$instance = new self();
        }
        return self::$instance;
    }

    private function __construct()
    {
        add_action('admin_menu', [$this, 'add_admin_menu']);
        add_action('admin_enqueue_scripts', [$this, 'enqueue_assets']);
        add_action('wp_ajax_ag_generate', [$this, 'ajax_generate']);
        add_action('wp_ajax_ag_download', [$this, 'ajax_download']);
    }

    /**
     * Aggiunge voce nel menu admin
     */
    public function add_admin_menu()
    {
        add_menu_page(
            'Attestati Generator',
            'Attestati',
            'manage_options',
            'attestati-generator',
            [$this, 'render_admin_page'],
            'dashicons-awards',
            30
        );
    }

    /**
     * Enqueue CSS e JS solo nella pagina del plugin
     */
    public function enqueue_assets($hook)
    {
        if ($hook !== 'toplevel_page_attestati-generator') {
            return;
        }

        wp_enqueue_style(
            'ag-admin',
            AG_PLUGIN_URL . 'assets/admin.css',
            [],
            AG_VERSION
        );

        wp_enqueue_script(
            'ag-admin',
            AG_PLUGIN_URL . 'assets/admin.js',
            ['jquery'],
            AG_VERSION,
            true
        );

        wp_localize_script('ag-admin', 'agAjax', [
            'ajaxUrl' => admin_url('admin-ajax.php'),
            'nonce' => wp_create_nonce('ag_nonce'),
        ]);
    }

    /**
     * Renderizza la pagina admin
     */
    public function render_admin_page()
    {
        if (!current_user_can('manage_options')) {
            return;
        }
        include AG_PLUGIN_DIR . 'templates/admin-page.php';
    }

    /**
     * AJAX: genera attestati
     */
    public function ajax_generate()
    {
        check_ajax_referer('ag_nonce', 'nonce');

        if (!current_user_can('manage_options')) {
            wp_send_json_error(['message' => 'Permessi insufficienti']);
        }

        // Verifica file caricati
        if (empty($_FILES['csv']) || empty($_FILES['template'])) {
            wp_send_json_error(['message' => 'Carica sia il file CSV che il template DOCX']);
        }

        if ($_FILES['csv']['error'] !== UPLOAD_ERR_OK) {
            wp_send_json_error(['message' => 'Errore nel caricamento del CSV']);
        }

        if ($_FILES['template']['error'] !== UPLOAD_ERR_OK) {
            wp_send_json_error(['message' => 'Errore nel caricamento del template']);
        }

        try {
            $generator = new AG_Certificate_Generator();

            $upload_dir = $generator->get_upload_dir();

            // Salva file caricati
            $csv_path = $upload_dir . '/' . uniqid() . '_' . sanitize_file_name($_FILES['csv']['name']);
            $tpl_path = $upload_dir . '/' . uniqid() . '_' . sanitize_file_name($_FILES['template']['name']);

            move_uploaded_file($_FILES['csv']['tmp_name'], $csv_path);
            move_uploaded_file($_FILES['template']['tmp_name'], $tpl_path);

            $format = isset($_POST['format']) ? sanitize_text_field($_POST['format']) : 'docx';

            // Genera
            $result = $generator->generate($csv_path, $tpl_path, $format);

            // Pulizia upload
            @unlink($csv_path);
            @unlink($tpl_path);

            wp_send_json_success($result);

        } catch (Exception $e) {
            wp_send_json_error(['message' => $e->getMessage()]);
        }
    }

    /**
     * AJAX: download file generato
     */
    public function ajax_download()
    {
        check_ajax_referer('ag_nonce', 'nonce');

        if (!current_user_can('manage_options')) {
            wp_die('Permessi insufficienti');
        }

        $filename = isset($_GET['file']) ? sanitize_file_name($_GET['file']) : '';

        if (empty($filename)) {
            wp_die('Nome file mancante');
        }

        $generator = new AG_Certificate_Generator();
        $file_path = $generator->get_output_dir() . '/' . $filename;

        if (!file_exists($file_path)) {
            wp_die('File non trovato');
        }

        $ext = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
        $content_types = [
            'zip' => 'application/zip',
            'pdf' => 'application/pdf',
            'docx' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        ];
        $content_type = $content_types[$ext] ?? 'application/octet-stream';

        header('Content-Type: ' . $content_type);
        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header('Content-Length: ' . filesize($file_path));
        header('Cache-Control: no-cache, must-revalidate');

        readfile($file_path);
        exit;
    }
}

// Avvia il plugin
Attestati_Generator_Plugin::instance();

// Attivazione: crea directory di output
register_activation_hook(__FILE__, function () {
    $dirs = [
        AG_PLUGIN_DIR . 'uploads',
        AG_PLUGIN_DIR . 'output',
    ];
    foreach ($dirs as $dir) {
        if (!is_dir($dir)) {
            wp_mkdir_p($dir);
        }
    }
});

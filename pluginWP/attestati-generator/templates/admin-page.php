<?php if (!defined('ABSPATH'))
    exit; ?>

<div class="wrap ag-wrap">
    <div class="ag-container">

        <!-- Header -->
        <div class="ag-header">
            <span class="ag-header-icon dashicons dashicons-awards"></span>
            <h1>Generatore Attestati</h1>
            <p class="ag-subtitle">Genera attestati personalizzati da CSV e template Word</p>
        </div>

        <!-- Upload Section -->
        <div class="ag-section ag-upload-section" id="agUploadSection">
            <div class="ag-card">
                <h2>📁 Carica i tuoi file</h2>

                <form id="agForm" enctype="multipart/form-data">

                    <!-- CSV -->
                    <div class="ag-file-group">
                        <label for="agCsvFile" class="ag-file-label">
                            <span class="dashicons dashicons-media-spreadsheet"></span>
                            <span>File CSV con i dati</span>
                        </label>
                        <input type="file" id="agCsvFile" name="csv" accept=".csv" required class="ag-file-input">
                        <div class="ag-file-name" id="agCsvName">Nessun file selezionato</div>
                    </div>

                    <!-- Template -->
                    <div class="ag-file-group">
                        <label for="agTemplateFile" class="ag-file-label">
                            <span class="dashicons dashicons-media-document"></span>
                            <span>Template Word (.docx)</span>
                        </label>
                        <input type="file" id="agTemplateFile" name="template" accept=".docx" required
                            class="ag-file-input">
                        <div class="ag-file-name" id="agTemplateName">Nessun file selezionato</div>
                    </div>

                    <input type="hidden" name="format" value="docx">

                    <button type="submit" class="button button-primary button-hero ag-btn-generate" id="agGenerateBtn">
                        ✨ Genera Attestati
                    </button>
                </form>
            </div>

            <!-- Info -->
            <div class="ag-card ag-info-card">
                <h3>ℹ️ Come funziona</h3>
                <ol>
                    <li>Carica un file <strong>CSV</strong> con le colonne: <code>NOME</code>, <code>COGNOME</code>,
                        <code>DATA</code>, <code>TITOLOEVENTO</code> (separato da <code>;</code>)
                    </li>
                    <li>Carica un template <strong>Word (.docx)</strong> con i placeholder: <code>${NOME}</code>,
                        <code>${COGNOME}</code>, <code>${DATA}</code>, <code>${TITOLOEVENTO}</code>
                    </li>
                    <li>Gli attestati vengono generati in formato <strong>DOCX</strong></li>
                    <li>Clicca su "Genera Attestati" e scarica lo <strong>ZIP</strong></li>
                </ol>
            </div>
        </div>

        <!-- Progress -->
        <div class="ag-section ag-progress-section" id="agProgressSection" style="display:none;">
            <div class="ag-card">
                <h2>⚙️ Generazione in corso...</h2>
                <div class="ag-progress-bar-wrap">
                    <div class="ag-progress-bar" id="agProgressBar"></div>
                </div>
                <p class="ag-progress-text" id="agProgressText">Elaborazione dei file...</p>
            </div>
        </div>

        <!-- Results -->
        <div class="ag-section ag-results-section" id="agResultsSection" style="display:none;">
            <div class="ag-card ag-success-card">
                <div class="ag-success-icon">✅</div>
                <h2>Attestati generati con successo!</h2>
                <p class="ag-success-msg" id="agSuccessMsg"></p>

                <button class="button button-primary button-hero ag-btn-download" id="agDownloadBtn">
                    📥 Scarica ZIP
                </button>

                <button class="button button-secondary ag-btn-reset" id="agResetBtn">
                    🔄 Genera altri attestati
                </button>

                <div class="ag-cert-list" id="agCertList"></div>
            </div>
        </div>

        <!-- Error -->
        <div class="ag-section ag-error-section" id="agErrorSection" style="display:none;">
            <div class="ag-card ag-error-card">
                <div class="ag-error-icon">❌</div>
                <h2>Si è verificato un errore</h2>
                <p class="ag-error-msg" id="agErrorMsg"></p>

                <button class="button button-secondary ag-btn-reset" id="agErrorResetBtn">
                    🔄 Riprova
                </button>
            </div>
        </div>

    </div>
</div>
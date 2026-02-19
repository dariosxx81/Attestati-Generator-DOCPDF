<!DOCTYPE html>
<html lang="it">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="Generatore automatico di attestati da file CSV e template HTML">
    <title>Generatore Attestati Automatico</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="assets/styles.css">
</head>

<body>
    <div class="container">
        <!-- Header -->
        <header class="header">
            <div class="header-icon">🎓</div>
            <h1 class="header-title">Generatore Attestati</h1>
            <p class="header-subtitle">Genera attestati personalizzati in automatico da CSV e template Word</p>
        </header>

        <!-- Main Content -->
        <main class="main-content">
            <!-- Upload Section -->
            <section class="upload-section">
                <div class="card">
                    <h2 class="card-title">📁 Carica i tuoi file</h2>

                    <form id="uploadForm" enctype="multipart/form-data">
                        <!-- CSV Upload -->
                        <div class="file-input-wrapper">
                            <label for="csvFile" class="file-label">
                                <span class="file-label-icon">📊</span>
                                <span class="file-label-text">File CSV con i dati</span>
                            </label>
                            <input type="file" id="csvFile" name="csv" accept=".csv" required class="file-input">
                            <div class="file-name" id="csvFileName">Nessun file selezionato</div>
                        </div>

                        <!-- Template Upload -->
                        <div class="file-input-wrapper">
                            <label for="templateFile" class="file-label">
                                <span class="file-label-icon">📄</span>
                                <span class="file-label-text">Template Word (.docx)</span>
                            </label>
                            <input type="file" id="templateFile" name="template" accept=".docx" required
                                class="file-input">
                            <div class="file-name" id="templateFileName">Nessun file selezionato</div>
                        </div>



                        <!-- Generate Button -->
                        <button type="submit" class="btn-generate" id="generateBtn">
                            <span class="btn-icon">✨</span>
                            <span class="btn-text">Genera Attestati</span>
                        </button>
                    </form>
                </div>

                <!-- Info Card -->
                <div class="card info-card">
                    <h3 class="info-title">ℹ️ Come funziona</h3>
                    <ol class="info-list">
                        <li>Carica un file <strong>CSV</strong> con le colonne: <code>NOME</code>, <code>COGNOME</code>,
                            <code>DATA</code>, <code>TITOLOEVENTO</code> (separato da <code>;</code>)
                        </li>
                        <li>Carica un template <strong>Word (.docx)</strong> con i placeholder: <code>${NOME}</code>,
                            <code>${COGNOME}</code>, <code>${DATA}</code>, <code>${TITOLOEVENTO}</code>
                        </li>
                        <li>Clicca su "Genera Attestati" e scarica lo <strong>ZIP</strong></li>
                        <li>Nello ZIP trovi anche lo script <strong>converti-docx-in-pdf.ps1</strong> per convertire in
                            PDF</li>
                    </ol>
                </div>
            </section>

            <!-- Progress Section -->
            <section class="progress-section" id="progressSection" style="display: none;">
                <div class="card">
                    <h2 class="card-title">⚙️ Generazione in corso...</h2>
                    <div class="progress-bar-container">
                        <div class="progress-bar" id="progressBar"></div>
                    </div>
                    <p class="progress-text" id="progressText">Elaborazione dei file...</p>
                </div>
            </section>

            <!-- Results Section -->
            <section class="results-section" id="resultsSection" style="display: none;">
                <div class="card success-card">
                    <div class="success-icon">✅</div>
                    <h2 class="success-title">Attestati generati con successo!</h2>
                    <p class="success-message" id="successMessage"></p>

                    <button class="btn-download" id="downloadBtn">
                        <span class="btn-icon">📥</span>
                        <span class="btn-text">Scarica ZIP</span>
                    </button>

                    <button class="btn-reset" id="resetBtn">
                        <span class="btn-icon">🔄</span>
                        <span class="btn-text">Genera altri attestati</span>
                    </button>

                    <!-- Certificate List -->
                    <div class="certificate-list" id="certificateList"></div>
                </div>
            </section>

            <!-- Error Section -->
            <section class="error-section" id="errorSection" style="display: none;">
                <div class="card error-card">
                    <div class="error-icon">❌</div>
                    <h2 class="error-title">Si è verificato un errore</h2>
                    <p class="error-message" id="errorMessage"></p>

                    <button class="btn-reset" id="errorResetBtn">
                        <span class="btn-icon">🔄</span>
                        <span class="btn-text">Riprova</span>
                    </button>
                </div>
            </section>
        </main>

        <!-- Footer -->
        <footer class="footer">
            <p>Creato con ❤️ per semplificare la generazione di attestati</p>
        </footer>
    </div>

    <script src="assets/script.js"></script>
</body>

</html>
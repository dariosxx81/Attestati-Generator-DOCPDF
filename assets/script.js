// Elementi DOM
const uploadForm = document.getElementById('uploadForm');
const csvFile = document.getElementById('csvFile');
const templateFile = document.getElementById('templateFile');
const csvFileName = document.getElementById('csvFileName');
const templateFileName = document.getElementById('templateFileName');
const generateBtn = document.getElementById('generateBtn');

const progressSection = document.getElementById('progressSection');
const progressBar = document.getElementById('progressBar');
const progressText = document.getElementById('progressText');

const resultsSection = document.getElementById('resultsSection');
const successMessage = document.getElementById('successMessage');
const downloadBtn = document.getElementById('downloadBtn');
const resetBtn = document.getElementById('resetBtn');
const certificateList = document.getElementById('certificateList');

const errorSection = document.getElementById('errorSection');
const errorMessage = document.getElementById('errorMessage');
const errorResetBtn = document.getElementById('errorResetBtn');

let currentZipFile = null;

// Event Listeners per aggiornare nome file
csvFile.addEventListener('change', (e) => {
    const fileName = e.target.files[0]?.name || 'Nessun file selezionato';
    csvFileName.textContent = fileName;
    csvFileName.classList.toggle('selected', e.target.files.length > 0);
});

templateFile.addEventListener('change', (e) => {
    const fileName = e.target.files[0]?.name || 'Nessun file selezionato';
    templateFileName.textContent = fileName;
    templateFileName.classList.toggle('selected', e.target.files.length > 0);
});

// Submit del form
uploadForm.addEventListener('submit', async (e) => {
    e.preventDefault();

    // Validazione
    if (!csvFile.files[0] || !templateFile.files[0]) {
        showError('Seleziona sia il file CSV che il template HTML');
        return;
    }

    // Prepara i dati
    const formData = new FormData();
    formData.append('csv', csvFile.files[0]);
    formData.append('template', templateFile.files[0]);

    // Formato di output: sempre DOCX
    formData.append('format', 'docx');

    // Mostra progress
    showProgress();

    try {
        // Invia richiesta al server PHP
        const response = await fetch('generate.php', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        if (!response.ok) {
            throw new Error(data.error || 'Errore durante la generazione');
        }

        // Mostra risultati
        showResults(data);

    } catch (error) {
        console.error('Errore:', error);
        showError(error.message);
    }
});

// Download ZIP
downloadBtn.addEventListener('click', () => {
    if (currentZipFile) {
        window.location.href = `download.php?file=${encodeURIComponent(currentZipFile)}`;
    }
});

// Reset
resetBtn.addEventListener('click', resetForm);
errorResetBtn.addEventListener('click', resetForm);

// Funzioni di visualizzazione
function showProgress() {
    document.querySelector('.upload-section').style.display = 'none';
    resultsSection.style.display = 'none';
    errorSection.style.display = 'none';

    progressSection.style.display = 'block';

    let progress = 0;
    const interval = setInterval(() => {
        progress += Math.random() * 15;
        if (progress > 90) progress = 90;
        progressBar.style.width = progress + '%';
    }, 200);

    progressSection.dataset.interval = interval;
}

function showResults(data) {
    const interval = progressSection.dataset.interval;
    if (interval) clearInterval(interval);

    progressSection.style.display = 'none';

    currentZipFile = data.zipFile;

    resultsSection.style.display = 'block';
    successMessage.textContent = `Sono stati generati ${data.count} attestati in formato ${data.format}!`;

    if (data.certificates && data.certificates.length > 0) {
        certificateList.innerHTML = '<h3 style="margin-bottom: 16px; color: var(--text-primary);">📋 Attestati generati:</h3>';
        data.certificates.forEach(cert => {
            const item = document.createElement('div');
            item.className = 'certificate-item';
            item.textContent = cert;
            certificateList.appendChild(item);
        });
    }
}

function showError(message) {
    const interval = progressSection.dataset.interval;
    if (interval) clearInterval(interval);

    progressSection.style.display = 'none';
    resultsSection.style.display = 'none';
    document.querySelector('.upload-section').style.display = 'none';

    errorSection.style.display = 'block';
    errorMessage.textContent = message;
}

function resetForm() {
    uploadForm.reset();
    csvFileName.textContent = 'Nessun file selezionato';
    templateFileName.textContent = 'Nessun file selezionato';
    csvFileName.classList.remove('selected');
    templateFileName.classList.remove('selected');

    currentZipFile = null;

    document.querySelector('.upload-section').style.display = 'grid';
    progressSection.style.display = 'none';
    resultsSection.style.display = 'none';
    errorSection.style.display = 'none';

    progressBar.style.width = '0%';
}

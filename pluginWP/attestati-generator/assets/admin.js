(function ($) {
    'use strict';

    var currentZipFile = null;
    var progressInterval = null;

    // File name display
    $('#agCsvFile').on('change', function () {
        var name = this.files[0] ? this.files[0].name : 'Nessun file selezionato';
        $('#agCsvName').text(name).toggleClass('selected', this.files.length > 0);
    });

    $('#agTemplateFile').on('change', function () {
        var name = this.files[0] ? this.files[0].name : 'Nessun file selezionato';
        $('#agTemplateName').text(name).toggleClass('selected', this.files.length > 0);
    });

    // Form submit
    $('#agForm').on('submit', function (e) {
        e.preventDefault();

        var csvFile = $('#agCsvFile')[0].files[0];
        var templateFile = $('#agTemplateFile')[0].files[0];

        if (!csvFile || !templateFile) {
            alert('Seleziona sia il file CSV che il template DOCX');
            return;
        }

        var formData = new FormData();
        formData.append('action', 'ag_generate');
        formData.append('nonce', agAjax.nonce);
        formData.append('csv', csvFile);
        formData.append('template', templateFile);
        formData.append('format', 'docx');

        showProgress();

        $.ajax({
            url: agAjax.ajaxUrl,
            type: 'POST',
            data: formData,
            processData: false,
            contentType: false,
            success: function (response) {
                if (response.success) {
                    showResults(response.data);
                } else {
                    showError(response.data.message || 'Errore sconosciuto');
                }
            },
            error: function (xhr) {
                var msg = 'Errore di comunicazione con il server';
                try {
                    var resp = JSON.parse(xhr.responseText);
                    if (resp.data && resp.data.message) {
                        msg = resp.data.message;
                    }
                } catch (e) { }
                showError(msg);
            }
        });
    });

    // Download
    $('#agDownloadBtn').on('click', function () {
        if (currentZipFile) {
            var url = agAjax.ajaxUrl +
                '?action=ag_download' +
                '&nonce=' + encodeURIComponent(agAjax.nonce) +
                '&file=' + encodeURIComponent(currentZipFile);
            window.location.href = url;
        }
    });

    // Reset
    $('#agResetBtn, #agErrorResetBtn').on('click', resetForm);

    function showProgress() {
        $('#agUploadSection').hide();
        $('#agResultsSection').hide();
        $('#agErrorSection').hide();
        $('#agProgressSection').show();

        var progress = 0;
        progressInterval = setInterval(function () {
            progress += Math.random() * 15;
            if (progress > 90) progress = 90;
            $('#agProgressBar').css('width', progress + '%');
        }, 200);
    }

    function showResults(data) {
        clearInterval(progressInterval);
        $('#agProgressSection').hide();

        currentZipFile = data.zipFile;

        $('#agSuccessMsg').text('Sono stati generati ' + data.count + ' attestati in formato ' + data.format + '!');

        var $list = $('#agCertList').empty();
        if (data.certificates && data.certificates.length > 0) {
            $list.append('<h3>📋 Attestati generati:</h3>');
            data.certificates.forEach(function (cert) {
                $list.append('<div class="ag-cert-item">' + cert + '</div>');
            });
        }

        $('#agResultsSection').show();
    }

    function showError(message) {
        clearInterval(progressInterval);
        $('#agProgressSection').hide();
        $('#agResultsSection').hide();
        $('#agUploadSection').hide();

        $('#agErrorMsg').text(message);
        $('#agErrorSection').show();
    }

    function resetForm() {
        $('#agForm')[0].reset();
        $('#agCsvName').text('Nessun file selezionato').removeClass('selected');
        $('#agTemplateName').text('Nessun file selezionato').removeClass('selected');

        currentZipFile = null;

        $('#agUploadSection').show();
        $('#agProgressSection').hide();
        $('#agResultsSection').hide();
        $('#agErrorSection').hide();

        $('#agProgressBar').css('width', '0%');
    }

})(jQuery);

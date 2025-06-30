$(document).ready(function () {
    const body = $('body');
    const darkModeToggle = $('#darkModeToggle');
    const dropArea = $('#drop-area');
    const fileElem = $('#fileElem');
    const fileNameDisplay = $('#fileName');
    const analyzeButton = $('#analyzeButton');
    const uploadMessage = $('#uploadMessage');
    const jsonResults = $('#jsonResults');
    const resultsSection = $('.analysis-results');
    const progressContainer = $('#progressContainer');
    const progressBar = $('#progressBar');
    const summaryResults = $('#summaryResults');
    const downloadPdfButton = $('#downloadPdfButton');
    const downloadXlsxButton = $('#downloadXlsxButton');

    // Elementos do Modal
    const modelModal = $('#modelModal');
    const modelInfoButton = $('#modelInfoButton');
    const helpButton = $('#helpButton');
    const closeButton = $('.close-button');

    // Progress bars
    const progressBarProcess = $('#progressBarProcess');
    const progressBarProcessInner = $('#progressBarProcessInner');
    const progressText = $('#progressText');

    let uploadedFile = null;
    let currentReport = null;
    let currentProcessingTime = 0;
    let currentAgrupamentosXlsx = null;
    let currentNaoAgrupadosXlsx = null;
    let dropTimeout = null;

    // Modo Escuro
    function applyDarkMode(isDark) {
        if (isDark) {
            body.addClass('dark-mode');
            darkModeToggle.html('<i class="fas fa-sun"></i>');
            localStorage.setItem('theme', 'dark');
        } else {
            body.removeClass('dark-mode');
            darkModeToggle.html('<i class="fas fa-moon"></i>');
            localStorage.setItem('theme', 'light');
        }
    }

    const savedTheme = localStorage.getItem('theme');
    if (savedTheme === 'dark') {
        applyDarkMode(true);
    } else {
        applyDarkMode(false);
    }

    darkModeToggle.on('click', function () {
        applyDarkMode(!body.hasClass('dark-mode'));
    });

    // Modal
    modelInfoButton.on('click', function () {
        modelModal.css('display', 'block');
    });

    helpButton.on('click', function () {
        modelModal.css('display', 'block');
    });

    closeButton.on('click', function () {
        modelModal.css('display', 'none');
    });

    $(window).on('click', function (event) {
        if ($(event.target).is(modelModal)) {
            modelModal.css('display', 'none');
        }
    });

    // Drag & Drop
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.on(eventName, preventDefaults);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    // Efeito ao entrar na área de drop
    dropArea.on('dragenter dragover', function (e) {
        clearTimeout(dropTimeout);
        dropArea.addClass('highlight');
        dropArea.css('background-color', body.hasClass('dark-mode') ? '#3a3a3a' : '#e1f0fa');
        dropArea.css('border-color', '#007bff');
        dropArea.css('box-shadow', '0 0 15px rgba(0, 123, 255, 0.5)');
    });

    // Efeito ao sair da área de drop
    dropArea.on('dragleave', function (e) {
        const relatedTarget = e.relatedTarget || e.originalEvent.explicitOriginalTarget;
        if (!dropArea.is(relatedTarget)) {
            resetDropArea();
        }
    });

    // Efeito ao soltar o arquivo
    dropArea.on('drop', function (e) {
        resetDropArea();
        let dt = e.originalEvent.dataTransfer;
        let files = dt.files;
        handleFiles(files);
    });

    function resetDropArea() {
        dropArea.removeClass('highlight');
        dropArea.css('background-color', body.hasClass('dark-mode') ? '#444' : '#fff');
        dropArea.css('border-color', '#007bff');
        dropArea.css('box-shadow', 'none');
    }

    fileElem.on('change', function () {
        handleFiles(this.files);
    });

    function handleFiles(files) {
        uploadedFile = null;
        uploadMessage.removeClass('success error').text('');
        currentReport = null;
        currentAgrupamentosXlsx = null;
        currentNaoAgrupadosXlsx = null;

        if (files.length > 0) {
            const file = files[0];
            const fileExtension = file.name.split('.').pop().toLowerCase();

            if (fileExtension === 'csv' || fileExtension === 'xlsx') {
                uploadedFile = file;
                fileNameDisplay.text(`Arquivo selecionado: ${file.name}`);
                analyzeButton.prop('disabled', false);

                // Efeito visual de confirmação
                dropArea.css('background-color', body.hasClass('dark-mode') ? '#2d5a2d' : '#e6f7e6');
                dropArea.css('border-color', '#28a745');
                dropTimeout = setTimeout(() => {
                    resetDropArea();
                }, 2000);
            } else {
                fileNameDisplay.text('');
                uploadMessage.addClass('error').text('Formato de arquivo inválido. Por favor, use .csv ou .xlsx.');
                analyzeButton.prop('disabled', true);

                // Efeito visual de erro
                dropArea.css('background-color', body.hasClass('dark-mode') ? '#5a2d2d' : '#f7e6e6');
                dropArea.css('border-color', '#dc3545');
                dropTimeout = setTimeout(() => {
                    resetDropArea();
                }, 2000);
            }
        } else {
            fileNameDisplay.text('');
            analyzeButton.prop('disabled', true);
        }
    }

    // Função para atualizar o progresso
    function updateProgress(progress) {
        progressBarProcessInner.css('width', progress + '%');
        progressText.text('Processando... ' + progress + '%');
    }

    // Análise do arquivo
    analyzeButton.on('click', function () {
        if (!uploadedFile) {
            uploadMessage.addClass('error').text('Nenhum arquivo selecionado para análise.');
            return;
        }

        uploadMessage.removeClass('success error').text('Enviando arquivo para análise...');
        analyzeButton.prop('disabled', true).text('Analisando...');
        progressContainer.show();
        progressBar.css('width', '0%');

        // Mostra a barra de processamento
        progressBarProcess.show();
        progressBarProcessInner.css('width', '0%');
        progressText.text('Processando... 0%');

        jsonResults.text('Processando dados...');
        resultsSection.hide();

        const formData = new FormData();
        formData.append('file', uploadedFile);

        $.ajax({
            url: '/analyze_file',
            type: 'POST',
            data: formData,
            processData: false,
            contentType: false,
            xhr: function () {
                const xhr = new window.XMLHttpRequest();

                // Progresso do upload
                xhr.upload.addEventListener('progress', function (e) {
                    if (e.lengthComputable) {
                        const percentComplete = Math.round((e.loaded / e.total) * 50);
                        progressBar.css('width', percentComplete + '%');
                    }
                });

                // Progresso do processamento (simulado)
                let progress = 50;
                const progressInterval = setInterval(function () {
                    progress += Math.random() * 5;
                    if (progress >= 100) {
                        clearInterval(progressInterval);
                        progress = 100;
                    }
                    updateProgress(progress);
                }, 500);

                return xhr;
            },
            success: function (response) {
                progressContainer.hide();
                progressBarProcess.hide();
                analyzeButton.prop('disabled', false).text('Analisar Arquivo');

                if (response.error) {
                    uploadMessage.addClass('error').text(response.error);
                    jsonResults.text('Erro na análise: ' + response.error);
                    resultsSection.show();
                } else if (response.success) {
                    uploadMessage.addClass('success').text('Análise concluída com sucesso!');
                    currentReport = response.relatorio;
                    currentAgrupamentosXlsx = response.relatorio_agrupamentos || [];
                    currentNaoAgrupadosXlsx = response.relatorio_nao_agrupados || [];
                    currentProcessingTime = response.resumo.tempo_processamento;

                    // Exibe o resumo
                    const resumo = response.resumo;
                    summaryResults.html(`
                        <p>Total de pacientes: ${resumo.total_pacientes}</p>
                        <p>Total de solicitações com status PENDENTE: ${resumo.total_solicitacoes}</p>
                        <p>Pacientes em agrupamentos: ${resumo.pacientes_agrupados}</p>
                        <p>Agrupamentos encontrados: ${resumo.agrupamentos_encontrados}</p>
                        <p>Tempo total de processamento: ${resumo.tempo_processamento} segundos</p>
                        <details>
                            <summary>Detalhes do tempo</summary>
                            <p>Leitura do arquivo: ${resumo.tempos_parciais.leitura} segundos</p>
                            <p>Formatação dos dados: ${resumo.tempos_parciais.formatacao} segundos</p>
                            <p>Análise dos dados: ${resumo.tempos_parciais.analise} segundos</p>
                        </details>
                    `);

                    // Exibe o relatório completo
                    jsonResults.text(response.relatorio.join('\n'));
                    resultsSection.show();
                } else if (response.message) {
                    uploadMessage.addClass('error').text(response.message);
                    jsonResults.text('Detalhes: ' + JSON.stringify(response.details || {}, null, 2));
                    resultsSection.show();
                }
            },
            error: function (jqXHR, textStatus, errorThrown) {
                progressContainer.hide();
                progressBarProcess.hide();
                analyzeButton.prop('disabled', false).text('Analisar Arquivo');
                uploadMessage.addClass('error').text('Erro ao enviar arquivo: ' + textStatus + '. Detalhes: ' + errorThrown);
                jsonResults.text('Falha na comunicação com o servidor.');
                resultsSection.show();
            }
        });
    });

    // Download do PDF
    downloadPdfButton.on('click', function () {
        if (!currentReport) {
            uploadMessage.addClass('error').text('Nenhum relatório disponível para download.');
            return;
        }

        showProcessingModal(true);

        $.ajax({
            url: '/download_pdf',
            type: 'POST',
            contentType: 'application/json',
            data: JSON.stringify({
                relatorio: currentReport,
                tempo_processamento: currentProcessingTime
            }),
            success: function (response, status, xhr) {
                // Cria um link temporário para forçar o download
                const blob = new Blob([response], { type: 'application/pdf' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'relatorio_oci.pdf';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);

                showProcessingModal(false);
            },
            error: function (jqXHR, textStatus, errorThrown) {
                showProcessingModal(false);
                uploadMessage.addClass('error').text('Erro ao gerar PDF: ' + errorThrown);
            }
        });
    });

    // Download do XLSX
    downloadXlsxButton.on('click', function () {
        if (!currentAgrupamentosXlsx || currentAgrupamentosXlsx.length === 0) {
            uploadMessage.addClass('error').text('Nenhum relatório disponível para download.');
            return;
        }

        showProcessingModal(true);

        $.ajax({
            url: '/download_xlsx',
            type: 'POST',
            contentType: 'application/json',
            data: JSON.stringify({ 
                relatorio_agrupamentos: currentAgrupamentosXlsx,
                relatorio_nao_agrupados: currentNaoAgrupadosXlsx
            }),
            xhrFields: {
                responseType: 'blob'
            },
            success: function (response, status, xhr) {
                // Cria um link temporário para forçar o download
                const blob = new Blob([response], { 
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
                });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `relatorio_oci_${new Date().toISOString().slice(0,19).replace(/:/g,'-')}.xlsx`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);

                showProcessingModal(false);
            },
            error: function (jqXHR, textStatus, errorThrown) {
                showProcessingModal(false);
                uploadMessage.addClass('error').text('Erro ao gerar XLSX: ' + errorThrown);
            }
        });
    });

    // Adiciona o modal de processamento
    const processingModal = `
<div id="processingModal" class="modal" style="display: none;">
  <div class="modal-content" style="text-align: center; max-width: 400px;">
    <h3>Processando Download</h3>
    <div id="processingMessage">
      <p><i class="fas fa-spinner fa-spin"></i> Preparando arquivo para download...</p>
      <p>Este arquivo pode demorar um pouco para ser processado. Aguarde...</p>
    </div>
  </div>
</div>`;

    $('body').append(processingModal);

    // Função para mostrar/ocultar o modal de processamento
    function showProcessingModal(show) {
        if (show) {
            $('#processingModal').css('display', 'block');
        } else {
            $('#processingModal').css('display', 'none');
        }
    }

    // Dowload do arquivo modelo
    document.getElementById("modelInfoButton").addEventListener("click", function () {
        window.location.href = "/download-modelo";
    });
});
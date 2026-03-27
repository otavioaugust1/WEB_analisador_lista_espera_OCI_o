// Admin Shared JavaScript - Funções comuns

class AdminUtils {
    constructor() {
        this.body = document.querySelector('body');
        this.darkModeToggle = document.getElementById('darkModeToggle');
        this.modelInfoButton = document.getElementById('modelInfoButton');
        this.helpButton = document.getElementById('helpButton');
        this.initializeTheme();
        this.initializeHeaderButtons();
    }

    // Tema escuro
    initializeTheme() {
        const savedTheme = localStorage.getItem('theme');
        if (savedTheme === 'dark') {
            this.applyDarkMode(true);
        } else {
            this.applyDarkMode(false);
        }

        if (this.darkModeToggle) {
            this.darkModeToggle.addEventListener('click', () => {
                this.applyDarkMode(!this.body.classList.contains('dark-mode'));
            });
        }
    }

    applyDarkMode(isDark) {
        if (isDark) {
            this.body.classList.add('dark-mode');
            if (this.darkModeToggle) this.darkModeToggle.innerHTML = '<i class="fas fa-sun"></i>';
            localStorage.setItem('theme', 'dark');
        } else {
            this.body.classList.remove('dark-mode');
            if (this.darkModeToggle) this.darkModeToggle.innerHTML = '<i class="fas fa-moon"></i>';
            localStorage.setItem('theme', 'light');
        }
    }

    // Modal de Modelo
    initializeHeaderButtons() {
        if (this.modelInfoButton || this.helpButton) {
            this.setupModelModal();
        }
    }

    setupModelModal() {
        let modelModal = document.getElementById('modelModal');
        
        if (!modelModal) {
            modelModal = this.createModelModal();
            document.body.appendChild(modelModal);
        }

        const closeButton = modelModal.querySelector('.close-button');

        if (this.modelInfoButton) {
            this.modelInfoButton.addEventListener('click', () => {
                modelModal.style.display = 'block';
            });
        }

        if (this.helpButton) {
            this.helpButton.addEventListener('click', () => {
                modelModal.style.display = 'block';
            });
        }

        if (closeButton) {
            closeButton.addEventListener('click', () => {
                modelModal.style.display = 'none';
            });
        }

        window.addEventListener('click', (event) => {
            if (event.target === modelModal) {
                modelModal.style.display = 'none';
            }
        });
    }

    createModelModal() {
        const modal = document.createElement('div');
        modal.id = 'modelModal';
        modal.className = 'modal';
        modal.innerHTML = `
            <div class="modal-content">
                <span class="close-button">&times;</span>
                <h2>Modelo de Colunas para o Arquivo de Filas</h2>
                <p>Para que a análise seja precisa, seu arquivo CSV ou XLSX deve conter as colunas obrigatórias.</p>
                <h3>Colunas Obrigatórias:</h3>
                <ul>
                    <li><strong>IDENTIFICADOR_LOCAL</strong> - ID único local</li>
                    <li><strong>DOCUMENTO_PACIENTE</strong> - CPF ou CNS</li>
                    <li><strong>DATA_SOLICITACAO</strong> - Data da solicitação</li>
                    <li><strong>CNES_SOLICITANTE</strong> - CNES da unidade</li>
                    <li><strong>CODIGO_SIGTAP</strong> - Código do procedimento</li>
                    <li><strong>STATUS</strong> - Status (1 = em espera)</li>
                </ul>
                <p><a href="/db/arquivo_modelo.xlsx" download>Baixar Modelo Excel</a></p>
            </div>
        `;
        return modal;
    }

    // Mostrar alerta
    mostrarAlerta(mensagem, tipo = 'success') {
        const alerta = document.getElementById('alerta');
        if (alerta) {
            alerta.textContent = mensagem;
            alerta.className = `alert show alert-${tipo}`;
            setTimeout(() => alerta.classList.remove('show'), 3000);
        }
    }

    // Fechar modal
    fecharModal(modal) {
        if (modal) {
            modal.classList.remove('show');
            const form = modal.querySelector('form');
            if (form) form.reset();
        }
    }

    // Abrir modal
    abrirModal(modal) {
        if (modal) {
            modal.classList.add('show');
        }
    }

    // Confirmar ação
    confirmarOpcao(mensagem) {
        return confirm(mensagem);
    }
}

// Inicializar utils ao carregar página
let adminUtils = null;
document.addEventListener('DOMContentLoaded', () => {
    adminUtils = new AdminUtils();
});

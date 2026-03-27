// Admin Agrupamentos JavaScript

document.addEventListener('DOMContentLoaded', function() {
    let codigoEditando = null;

    // Elements
    const alerta = document.getElementById('alerta');
    const modal = document.getElementById('modalAgrupamento');
    const form = document.getElementById('formAgrupamento');

    // Abrir modal novo
    document.querySelector('.btn-novo').addEventListener('click', function() {
        codigoEditando = null;
        document.getElementById('modalTitulo').textContent = 'Novo Agrupamento';
        document.getElementById('codigo').disabled = false;
        form.reset();
        adminUtils.abrirModal(modal);
    });

    // Abrir modal editar
    document.querySelectorAll('.btn-editar').forEach(btn => {
        btn.addEventListener('click', function() {
            const codigo = this.dataset.codigo;
            codigoEditando = codigo;
            
            fetch(`/api/agrupamentos/${codigo}`)
                .then(r => r.json())
                .then(data => {
                    document.getElementById('modalTitulo').textContent = 'Editar Agrupamento';
                    document.getElementById('codigo').value = data.codigo;
                    document.getElementById('nome').value = data.nome;
                    document.getElementById('descricao').value = data.descricao || '';
                    document.getElementById('codigo').disabled = true;
                    adminUtils.abrirModal(modal);
                })
                .catch(() => adminUtils.mostrarAlerta('Erro ao carregar', 'error'));
        });
    });

    // Fechar modal
    const closeBtn = document.querySelector('.close-btn');
    if (closeBtn) closeBtn.addEventListener('click', () => adminUtils.fecharModal(modal));
    if (modal) modal.addEventListener('click', function(e) {
        if (e.target === this) adminUtils.fecharModal(modal);
    });

    const cancelBtn = document.querySelector('.btn-cancel');
    if (cancelBtn) cancelBtn.addEventListener('click', () => adminUtils.fecharModal(modal));

    // Salvar agrupamento
    form.addEventListener('submit', function(e) {
        e.preventDefault();

        const dados = {
            codigo: document.getElementById('codigo').value,
            nome: document.getElementById('nome').value,
            descricao: document.getElementById('descricao').value,
            itens_obrigatorios: [],
            itens_facultativos: []
        };

        const metodo = codigoEditando ? 'PUT' : 'POST';
        const url = codigoEditando ? `/api/agrupamentos/${codigoEditando}` : '/api/agrupamentos';

        fetch(url, {
            method: metodo,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(dados)
        })
        .then(r => r.json())
        .then(() => {
            adminUtils.mostrarAlerta(codigoEditando ? 'Atualizado!' : 'Criado!');
            adminUtils.fecharModal(modal);
            setTimeout(() => location.reload(), 1500);
        })
        .catch(() => adminUtils.mostrarAlerta('Erro', 'error'));
    });

    // Visualizar agrupamento
    document.querySelectorAll('.btn-visualizar').forEach(btn => {
        btn.addEventListener('click', function() {
            window.location.href = `/admin/visualizar/${this.dataset.codigo}`;
        });
    });

    // Deletar agrupamento
    document.querySelectorAll('.btn-deletar').forEach(btn => {
        btn.addEventListener('click', function() {
            const codigo = this.dataset.codigo;
            if (adminUtils.confirmarOpcao(`Deletar "${codigo}"?`)) {
                fetch(`/api/agrupamentos/${codigo}`, { method: 'DELETE' })
                    .then(r => r.json())
                    .then(() => {
                        adminUtils.mostrarAlerta('Deletado!');
                        setTimeout(() => location.reload(), 1500);
                    })
                    .catch(() => adminUtils.mostrarAlerta('Erro', 'error'));
            }
        });
    });

    // Exportar JSON
    document.querySelector('.btn-exportar[data-format="json"]').addEventListener('click', function() {
        window.location.href = '/api/exportar/json';
    });

    // Exportar CSV
    document.querySelector('.btn-exportar[data-format="csv"]').addEventListener('click', function() {
        window.location.href = '/api/exportar/csv';
    });
});


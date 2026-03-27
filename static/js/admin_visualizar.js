// Admin Visualizar JavaScript

document.addEventListener('DOMContentLoaded', function() {
    const alerta = document.getElementById('alerta');
    const modal = document.getElementById('modalProcedimento');
    const form = document.getElementById('formProcedimento');
    const agrupamentoId = document.body.dataset.agrupamentoId;
    
    let tipoProcedimentoAtual = null;
    let procedimentoIdEditando = null;

    // Abrir modal novo procedimento
    document.querySelectorAll('.btn-adicionar').forEach(btn => {
        btn.addEventListener('click', function() {
            const tipo = this.dataset.tipo;
            tipoProcedimentoAtual = tipo;
            procedimentoIdEditando = null;
            
            const titulo = tipo === 'obrigatorio' 
                ? 'Novo Procedimento Obrigatório' 
                : 'Novo Procedimento Facultativo';
            document.getElementById('modalProcTitulo').textContent = titulo;
            form.reset();
            adminUtils.abrirModal(modal);
        });
    });

    // Editar procedimento
    document.querySelectorAll('.btn-table-editar').forEach(btn => {
        btn.addEventListener('click', function() {
            const id = this.dataset.id;
            const tipo = this.dataset.tipo;
            const codigo = this.dataset.codigo;
            const descricao = this.dataset.descricao;

            tipoProcedimentoAtual = tipo;
            procedimentoIdEditando = id;

            const titulo = tipo === 'obrigatorio'
                ? 'Editar Procedimento Obrigatório'
                : 'Editar Procedimento Facultativo';
            
            document.getElementById('modalProcTitulo').textContent = titulo;
            document.getElementById('procCodigo').value = codigo;
            document.getElementById('procDescricao').value = descricao;
            adminUtils.abrirModal(modal);
        });
    });

    // Deletar procedimento
    document.querySelectorAll('.btn-table-deletar').forEach(btn => {
        btn.addEventListener('click', function() {
            const id = this.dataset.id;
            const tipo = this.dataset.tipo;

            if (adminUtils.confirmarOpcao('Deletar procedimento?')) {
                const endpoint = tipo === 'obrigatorio'
                    ? `/api/procedimentos-obrigatorios/${id}`
                    : `/api/procedimentos-facultativos/${id}`;

                fetch(endpoint, { method: 'DELETE' })
                    .then(r => r.json())
                    .then(() => {
                        adminUtils.mostrarAlerta('Deletado!');
                        setTimeout(() => location.reload(), 1500);
                    })
                    .catch(() => adminUtils.mostrarAlerta('Erro', 'error'));
            }
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

    // Salvar procedimento
    form.addEventListener('submit', function(e) {
        e.preventDefault();

        const codigo = document.getElementById('procCodigo').value;
        const descricao = document.getElementById('procDescricao').value;

        if (procedimentoIdEditando) {
            // Editar
            const endpoint = tipoProcedimentoAtual === 'obrigatorio'
                ? `/api/procedimentos-obrigatorios/${procedimentoIdEditando}`
                : `/api/procedimentos-facultativos/${procedimentoIdEditando}`;

            fetch(endpoint, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ codigo, descricao })
            })
            .then(r => r.json())
            .then(() => {
                adminUtils.mostrarAlerta('Atualizado!');
                adminUtils.fecharModal(modal);
                setTimeout(() => location.reload(), 1500);
            })
            .catch(() => adminUtils.mostrarAlerta('Erro', 'error'));
        } else {
            // Criar
            const endpoint = tipoProcedimentoAtual === 'obrigatorio'
                ? `/api/agrupamentos/${agrupamentoId}/procedimentos-obrigatorios`
                : `/api/agrupamentos/${agrupamentoId}/procedimentos-facultativos`;

            fetch(endpoint, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ codigo, descricao })
            })
            .then(r => r.json())
            .then(() => {
                adminUtils.mostrarAlerta('Criado!');
                adminUtils.fecharModal(modal);
                setTimeout(() => location.reload(), 1500);
            })
            .catch(() => adminUtils.mostrarAlerta('Erro', 'error'));
        }
    });
});


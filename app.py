"""
Analisador de Lista de Espera para OCI (Oncologia Clínica Integrada)
Sistema web para análise e processamento de listas de espera com Flask.

Autor: Otávio August
Versão: 2.0.0 - Refatorada com SQLite
"""

import io
import os
import re
import sys
from datetime import datetime
from pathlib import Path
from time import time

import pandas as pd
from flask import (
    Flask,
    jsonify,
    render_template,
    request,
    send_file,
    send_from_directory,
)
from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas
from werkzeug.utils import secure_filename

import database

# ==================== CONFIGURAÇÃO DA APLICAÇÃO ====================

# Detectar se está rodando como bundle PyInstaller
if getattr(sys, 'frozen', False):
    # Rodando como executável compilado
    _bundle_dir = sys._MEIPASS          # recursos bundled (templates, static)
    _data_dir = Path(sys.executable).parent  # diretório do .exe (db, uploads)
else:
    _bundle_dir = Path(__file__).parent
    _data_dir = Path(__file__).parent

app = Flask(
    __name__,
    template_folder=str(Path(_bundle_dir) / 'templates'),
    static_folder=str(Path(_bundle_dir) / 'static'),
)

app.config['UPLOAD_FOLDER'] = str(_data_dir / 'uploads')
app.config['ALLOWED_EXTENSIONS'] = {'csv', 'xlsx'}
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB
app.config['JSON_SORT_KEYS'] = False

# Criar pastas se não existirem
Path(app.config['UPLOAD_FOLDER']).mkdir(exist_ok=True)
(_data_dir / 'db').mkdir(exist_ok=True)

# Configurar caminho do banco de dados
database.DATABASE_FILE = str(_data_dir / 'db' / 'agrupamentos.db')

# Inicializar banco de dados
database.init_database()

# ==================== CONFIGURAÇÕES ====================

REQUIRED_COLUMNS = [
    'IDENTIFICADOR_LOCAL',
    'DOCUMENTO_PACIENTE',
    'DATA_SOLICITACAO',
    'CNES_SOLICITANTE',
    'CNES_REGULADOR',
    'CODIGO_SIGTAP',
    'CBO',
    'CID10',
    'CODIGO_MODALIDADE_ASSISTENCIAL',
    'CODIGO_CARTER_SOLICITACAO',
    'STATUS',
    'DATA_AUTORIZACAO',
    'DATA_EXECUCAO',
    'CNES_EXECUTANTE',
]

# Tamanho de chunk para processar arquivos grandes (em MB)
CHUNK_SIZE_MB = 10

# ==================== FUNÇÕES UTILITÁRIAS ====================


def allowed_file(filename: str) -> bool:
    """Valida extensão do arquivo."""
    if '.' not in filename:
        return False
    ext = filename.rsplit('.', 1)[1].lower()
    return ext in app.config['ALLOWED_EXTENSIONS']


def get_file_size_mb(file_obj) -> float:
    """Calcula tamanho do arquivo em MB."""
    file_obj.seek(0, os.SEEK_END)
    size = file_obj.tell()
    file_obj.seek(0)
    return size / (1024 * 1024)


def ler_arquivo_em_chunks(
    file_obj, file_extension: str, chunk_size_mb: int = CHUNK_SIZE_MB
):
    """
    Lê arquivo em chunks para não sobrecarregar memória.
    Funciona com CSV e XLSX.
    """
    if file_extension.lower() == 'csv':
        # Para CSV, lê por chunks definidos
        chunks = []
        try:
            for chunk in pd.read_csv(
                file_obj,
                encoding='utf-8',
                sep=';',
                dtype=str,
                chunksize=5000,  # 5000 linhas por chunk
            ):
                chunks.append(chunk)

            if chunks:
                df = pd.concat(chunks, ignore_index=True)
                return df
            else:
                return pd.DataFrame()
        except Exception as e:
            raise ValueError(f'Erro ao ler arquivo CSV: {str(e)}')

    elif file_extension.lower() == 'xlsx':
        # Para XLSX, tenta ler normalmente primeiro
        try:
            file_obj.seek(0)
            df = pd.read_excel(file_obj, dtype=str)
            return df
        except Exception as e:
            raise ValueError(f'Erro ao ler arquivo XLSX: {str(e)}')

    else:
        raise ValueError(f'Formato de arquivo não suportado: {file_extension}')


def formatar_dados(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica formatações necessárias nos dados."""

    def aplicar_zfill(serie, tamanho):
        """Aplica padding apenas em valores numéricos."""
        serie = serie.fillna('').astype(str)
        return serie.where(
            ~serie.str.match(r'^\d+$'), serie.str.zfill(tamanho)
        )

    # Formatar códigos numéricos
    cnes_cols = ['CNES_SOLICITANTE', 'CNES_REGULADOR', 'CNES_EXECUTANTE']
    for col in cnes_cols:
        if col in df.columns:
            df[col] = aplicar_zfill(df[col], 7)

    if 'CODIGO_SIGTAP' in df.columns:
        df['CODIGO_SIGTAP'] = aplicar_zfill(df['CODIGO_SIGTAP'], 10)

    # Formatar datas
    date_cols = ['DATA_SOLICITACAO', 'DATA_AUTORIZACAO', 'DATA_EXECUCAO']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(
                df[col], errors='coerce', infer_datetime_format=True
            )
            df[col] = df[col].dt.strftime('%d/%m/%Y')
            df[col] = df[col].replace('NaT', '')

    return df


def analisar_dados(df: pd.DataFrame) -> dict:
    """Analisa dados e identifica agrupamentos OCI."""

    # Obter agrupamentos do banco de dados
    agrupamentos = database.obter_agrupamentos()

    # Obter tabela de descrições SIGTAP
    descricoes_sigtap = database.obter_todas_descricoes_sigtap()

    try:
        # Filtrar apenas STATUS == 1 (Em Espera)
        df = df[df['STATUS'] == '1'].copy()
        df = df.fillna('')

        total_solicitacoes = len(df)
        total_pacientes = len(df['DOCUMENTO_PACIENTE'].unique())

        relatorio = []
        relatorio_agrupamentos = []
        relatorio_nao_agrupados = []

        relatorio.append(
            "*********************    FORAM ENCONTRADOS {} CONJUNTOS DE OCI'S    ***********************\n"
        )

        pacientes_em_agrupamentos = set()
        agrupamentos_encontrados = 0

        # Analisar cada agrupamento
        for codigo, agrupamento in agrupamentos.items():
            codigos_obrigatorios = {
                item['codigo'] for item in agrupamento['itens_obrigatorios']
            }
            codigos_facultativos = {
                item['codigo'] for item in agrupamento['itens_facultativos']
            }

            pacientes_agrupados = []

            # Procurar por pacientes que têm os itens obrigatórios
            for paciente, dados_paciente in df.groupby('DOCUMENTO_PACIENTE'):
                codigos_paciente = set(
                    dados_paciente['CODIGO_SIGTAP'].unique()
                )

                # Verifica se tem TODOS os códigos obrigatórios
                if codigos_obrigatorios.issubset(codigos_paciente):
                    pacientes_agrupados.append(paciente)
                    pacientes_em_agrupamentos.add(paciente)

            if pacientes_agrupados:
                agrupamentos_encontrados += 1
                relatorio.append(
                    '_________________________________________________________________________________________'
                )
                codigo_formatado = (
                    f'{codigo[:2]}.{codigo[2:4]}.{codigo[4:6]}.{codigo[6:]}'
                )
                relatorio.append(
                    f"{codigo_formatado} - {agrupamento['nome']}\n"
                )

                for paciente in sorted(pacientes_agrupados):
                    relatorio.append(f'--- {paciente}')

                    # Processar itens obrigatórios
                    for item in agrupamento['itens_obrigatorios']:
                        registros = df[
                            (df['DOCUMENTO_PACIENTE'] == paciente)
                            & (df['CODIGO_SIGTAP'] == item['codigo'])
                        ]

                        for _, registro in registros.iterrows():
                            relatorio.append(
                                f"-------- OBG\tCNES_SOLC {registro['CNES_SOLICITANTE']}\tCID-{registro['CID10']}\tDT_SOLC-{registro['DATA_SOLICITACAO']}\t{item['codigo']} - {item['descricao']}"
                            )
                            relatorio_agrupamentos.append(
                                {
                                    'AGRUPAMENTO_OCI': codigo_formatado,
                                    'DESCRICAO_OCI': agrupamento['nome'],
                                    'DOCUMENTO_PACIENTE': paciente,
                                    'DATA_SOLICITACAO': registro[
                                        'DATA_SOLICITACAO'
                                    ],
                                    'CNES_SOLICITANTE': registro[
                                        'CNES_SOLICITANTE'
                                    ],
                                    'ITEM OBG/FAC (X)': 'OBG',
                                    'CID10': registro['CID10'],
                                    'CODIGO_SIGTAP': item['codigo'],
                                    'DESCRICAO_SIGTAP': item['descricao'],
                                    'CBO': registro['CBO'],
                                }
                            )

                    # Processar itens facultativos
                    for item in agrupamento['itens_facultativos']:
                        registros = df[
                            (df['DOCUMENTO_PACIENTE'] == paciente)
                            & (df['CODIGO_SIGTAP'] == item['codigo'])
                        ]

                        for _, registro in registros.iterrows():
                            relatorio.append(
                                f"-------- FAC\tCNES_SOLC {registro['CNES_SOLICITANTE']}\tCID-{registro['CID10']}\tDT_SOLC-{registro['DATA_SOLICITACAO']}\t{item['codigo']} - {item['descricao']}"
                            )
                            relatorio_agrupamentos.append(
                                {
                                    'AGRUPAMENTO_OCI': codigo_formatado,
                                    'DESCRICAO_OCI': agrupamento['nome'],
                                    'DOCUMENTO_PACIENTE': paciente,
                                    'DATA_SOLICITACAO': registro[
                                        'DATA_SOLICITACAO'
                                    ],
                                    'CNES_SOLICITANTE': registro[
                                        'CNES_SOLICITANTE'
                                    ],
                                    'ITEM OBG/FAC (X)': 'FAC',
                                    'CID10': registro['CID10'],
                                    'CODIGO_SIGTAP': item['codigo'],
                                    'DESCRICAO_SIGTAP': item['descricao'],
                                    'CBO': registro['CBO'],
                                }
                            )

        # Pacientes não agrupados
        pacientes_restantes = df[
            ~df['DOCUMENTO_PACIENTE'].isin(pacientes_em_agrupamentos)
        ]

        relatorio.append(
            '\n********************    PACIENTES QUE NÃO ESTÃO EM NENHUM CONJUNTO  ***********************'
        )

        for _, linha in pacientes_restantes.iterrows():
            descricao = descricoes_sigtap.get(
                linha['CODIGO_SIGTAP'],
                'Código não faz parte de um item de OCI',
            )
            relatorio.append(
                f"- CNES_SOLC {linha['CNES_SOLICITANTE']}\tCID {linha['CID10']}\tCNS/CPF_PAC {linha['DOCUMENTO_PACIENTE']}\tDT_SOLC {linha['DATA_SOLICITACAO']}\t{linha['CODIGO_SIGTAP']} - {descricao}"
            )
            relatorio_nao_agrupados.append(
                {
                    'DOCUMENTO_PACIENTE': linha['DOCUMENTO_PACIENTE'],
                    'DATA_SOLICITACAO': linha['DATA_SOLICITACAO'],
                    'CNES_SOLICITANTE': linha['CNES_SOLICITANTE'],
                    'CID10': linha['CID10'],
                    'CODIGO_SIGTAP': linha['CODIGO_SIGTAP'],
                    'DESCRICAO_SIGTAP': descricao,
                    'CBO': linha['CBO'],
                }
            )

        # Atualizar cabeçalho com número real de conjuntos
        relatorio[0] = relatorio[0].format(agrupamentos_encontrados)

        # Rodapé com data e hora
        relatorio.append('\n{:%d/%m/%Y %H:%M:%S}'.format(datetime.now()))

        return {
            'relatorio': relatorio,
            'relatorio_agrupamentos': relatorio_agrupamentos,
            'relatorio_nao_agrupados': relatorio_nao_agrupados,
            'total_pacientes': total_pacientes,
            'pacientes_agrupados': len(pacientes_em_agrupamentos),
            'agrupamentos_encontrados': agrupamentos_encontrados,
            'total_solicitacoes': total_solicitacoes,
        }

    except Exception as e:
        app.logger.error(f'Erro na análise de dados: {str(e)}')
        raise e


def gerar_pdf(relatorio: list, tempo_processamento: float) -> io.BytesIO:
    """Gera PDF do relatório."""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(letter))
    width, height = landscape(letter)

    margem_esquerda = 30
    linha_altura = 12
    max_linhas_por_pagina = 50
    pagina_numero = 1
    data_hora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    # Filtrar apenas linhas de agrupamentos
    linhas_agrupamentos = []
    for linha in relatorio:
        if linha.startswith(
            '********************    PACIENTES QUE NÃO ESTÃO EM NENHUM CONJUNTO'
        ):
            break
        linhas_agrupamentos.append(linha)

    def desenhar_cabecalho(pagina):
        """Desenha cabeçalho do PDF."""
        try:
            c.drawImage(
                'static/img/agora-tem-especialistas.png',
                margem_esquerda,
                height - 80,
                width=100,
                height=50,
            )
        except:
            pass  # Ignora se imagem não existir

        c.setFont('Helvetica-Bold', 16)
        c.drawString(
            margem_esquerda + 120,
            height - 60,
            'Relatório de Análise de Filas OCI',
        )
        c.setFont('Helvetica', 10)
        c.drawString(
            margem_esquerda + 120, height - 80, f'Gerado em: {data_hora}'
        )
        c.drawString(
            margem_esquerda + 120,
            height - 95,
            f'Tempo de processamento: {tempo_processamento} segundos',
        )
        c.drawRightString(
            width - margem_esquerda, height - 95, f'Página {pagina}'
        )
        c.line(
            margem_esquerda,
            height - 105,
            width - margem_esquerda,
            height - 105,
        )

    linha_index = 0
    total_linhas = len(linhas_agrupamentos)

    while linha_index < total_linhas:
        c.setPageSize(landscape(letter))
        desenhar_cabecalho(pagina_numero)
        y = height - 120

        linhas_na_pagina = 0
        while (
            linhas_na_pagina < max_linhas_por_pagina
            and linha_index < total_linhas
        ):
            linha = linhas_agrupamentos[linha_index]

            if linha.startswith('---') or linha.startswith('--------'):
                c.setFont('Courier', 8)
            else:
                c.setFont('Helvetica-Bold', 9)

            c.drawString(margem_esquerda, y, linha[:200])
            y -= linha_altura
            linha_index += 1
            linhas_na_pagina += 1

            if y < 50:
                break

        c.showPage()
        pagina_numero += 1

    c.save()
    buffer.seek(0)
    return buffer


# ==================== ROTAS ====================


@app.route('/')
def index():
    """Página inicial."""
    return render_template('index.html')


@app.route('/analyze_file', methods=['POST'])
def analyze_file():
    """Processa arquivo enviado e realiza análise."""

    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'Nome de arquivo vazio'}), 400

    # Validação melhorada da extensão
    if not allowed_file(file.filename):
        return (
            jsonify(
                {
                    'error': 'Tipo de arquivo não permitido',
                    'allowed': list(app.config['ALLOWED_EXTENSIONS']),
                    'received': file.filename.rsplit('.', 1)[1].lower()
                    if '.' in file.filename
                    else 'sem extensão',
                }
            ),
            400,
        )

    try:
        tempo_inicio = time()

        # Validar tamanho do arquivo
        file_size_mb = get_file_size_mb(file)
        if file_size_mb > 500:
            return (
                jsonify(
                    {
                        'error': f'Arquivo muito grande ({file_size_mb:.2f}MB). Máximo permitido: 500MB'
                    }
                ),
                413,
            )

        app.logger.info(
            f'Processando arquivo: {file.filename} ({file_size_mb:.2f}MB)'
        )

        # Determinar extensão
        file_extension = file.filename.rsplit('.', 1)[1].lower()

        # Ler dados com tratamento para arquivos grandes
        df = ler_arquivo_em_chunks(file, file_extension)

        if df.empty:
            return (
                jsonify({'error': 'Arquivo vazio ou sem dados válidos'}),
                400,
            )

        df = df.fillna('')
        tempo_leitura = time()

        # Aplicar formatações
        df = formatar_dados(df)
        tempo_formatacao = time()

        # Validar colunas obrigatórias
        missing_columns = [
            col for col in REQUIRED_COLUMNS if col not in df.columns
        ]
        if missing_columns:
            return (
                jsonify(
                    {
                        'message': 'Colunas obrigatórias faltando no arquivo',
                        'details': {
                            'missing_columns': missing_columns,
                            'required_columns': REQUIRED_COLUMNS,
                            'available_columns': list(df.columns),
                        },
                    }
                ),
                400,
            )

        # Realizar análise
        resultado = analisar_dados(df)
        tempo_analise = time()

        tempo_total = tempo_analise - tempo_inicio

        return jsonify(
            {
                'success': True,
                'relatorio': resultado['relatorio'],
                'relatorio_agrupamentos': resultado['relatorio_agrupamentos'],
                'relatorio_nao_agrupados': resultado[
                    'relatorio_nao_agrupados'
                ],
                'resumo': {
                    'total_pacientes': resultado['total_pacientes'],
                    'total_solicitacoes': resultado['total_solicitacoes'],
                    'pacientes_agrupados': resultado['pacientes_agrupados'],
                    'agrupamentos_encontrados': resultado[
                        'agrupamentos_encontrados'
                    ],
                    'tempo_processamento': round(tempo_total, 2),
                    'tamanho_arquivo_mb': round(file_size_mb, 2),
                    'tempos_parciais': {
                        'leitura': round(tempo_leitura - tempo_inicio, 2),
                        'formatacao': round(
                            tempo_formatacao - tempo_leitura, 2
                        ),
                        'analise': round(tempo_analise - tempo_formatacao, 2),
                    },
                },
            }
        )

    except Exception as e:
        app.logger.error(f'Erro ao processar arquivo: {str(e)}')
        return jsonify({'error': f'Erro ao processar arquivo: {str(e)}'}), 500


@app.route('/download_pdf', methods=['POST'])
def download_pdf():
    """Gera e baixa PDF do relatório."""
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'Dados não fornecidos'}), 400

        relatorio = data.get('relatorio')
        tempo_processamento = data.get('tempo_processamento')

        if not relatorio or tempo_processamento is None:
            return jsonify({'error': 'Parâmetros faltando'}), 400

        pdf_buffer = gerar_pdf(relatorio, tempo_processamento)
        return send_file(
            pdf_buffer,
            as_attachment=True,
            download_name='relatorio_agrupamentos_oci.pdf',
            mimetype='application/pdf',
        )
    except Exception as e:
        app.logger.error(f'Erro ao gerar PDF: {str(e)}')
        return jsonify({'error': f'Erro ao gerar PDF: {str(e)}'}), 500


@app.route('/download_xlsx', methods=['POST'])
def download_xlsx():
    """Gera e baixa XLSX com os relatórios."""
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'Dados não fornecidos'}), 400

        relatorio_agrupamentos = data.get('relatorio_agrupamentos', [])
        relatorio_nao_agrupados = data.get('relatorio_nao_agrupados', [])

        if not relatorio_agrupamentos and not relatorio_nao_agrupados:
            return jsonify({'error': 'Nenhum relatório fornecido'}), 400

        ordem_agrupamentos = [
            'AGRUPAMENTO_OCI',
            'DESCRICAO_OCI',
            'DOCUMENTO_PACIENTE',
            'DATA_SOLICITACAO',
            'CNES_SOLICITANTE',
            'ITEM OBG/FAC (X)',
            'CID10',
            'CODIGO_SIGTAP',
            'DESCRICAO_SIGTAP',
            'CBO',
        ]

        ordem_nao_agrupados = [
            'DOCUMENTO_PACIENTE',
            'DATA_SOLICITACAO',
            'CNES_SOLICITANTE',
            'CID10',
            'CODIGO_SIGTAP',
            'DESCRICAO_SIGTAP',
            'CBO',
        ]

        df_agrupamentos = (
            pd.DataFrame(relatorio_agrupamentos)[ordem_agrupamentos]
            if relatorio_agrupamentos
            else pd.DataFrame()
        )
        df_nao_agrupados = (
            pd.DataFrame(relatorio_nao_agrupados)[ordem_nao_agrupados]
            if relatorio_nao_agrupados
            else pd.DataFrame()
        )

        if not df_agrupamentos.empty:
            df_agrupamentos = df_agrupamentos.sort_values(
                by=[
                    'AGRUPAMENTO_OCI',
                    'DOCUMENTO_PACIENTE',
                    'ITEM OBG/FAC (X)',
                    'CODIGO_SIGTAP',
                ]
            )

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Aba de agrupamentos
            if not df_agrupamentos.empty:
                df_agrupamentos.to_excel(
                    writer, sheet_name='Agrupamentos', index=False
                )
                worksheet = writer.sheets['Agrupamentos']

                header_format = writer.book.add_format(
                    {
                        'bold': True,
                        'text_wrap': True,
                        'valign': 'top',
                        'fg_color': '#4472C4',
                        'font_color': 'white',
                        'border': 1,
                    }
                )

                for col_num, value in enumerate(df_agrupamentos.columns):
                    worksheet.write(0, col_num, value, header_format)

                for col_num, column in enumerate(df_agrupamentos.columns):
                    max_len = max(
                        df_agrupamentos[column].astype(str).map(len).max(),
                        len(column),
                    )
                    worksheet.set_column(
                        col_num, col_num, min(max_len + 2, 50)
                    )

                worksheet.freeze_panes(1, 0)

            # Aba de não agrupados
            if not df_nao_agrupados.empty:
                df_nao_agrupados.to_excel(
                    writer, sheet_name='Não Agrupados', index=False
                )
                worksheet = writer.sheets['Não Agrupados']

                header_format = writer.book.add_format(
                    {
                        'bold': True,
                        'text_wrap': True,
                        'valign': 'top',
                        'fg_color': '#4472C4',
                        'font_color': 'white',
                        'border': 1,
                    }
                )

                for col_num, value in enumerate(df_nao_agrupados.columns):
                    worksheet.write(0, col_num, value, header_format)

                for col_num, column in enumerate(df_nao_agrupados.columns):
                    max_len = max(
                        df_nao_agrupados[column].astype(str).map(len).max(),
                        len(column),
                    )
                    worksheet.set_column(
                        col_num, col_num, min(max_len + 2, 50)
                    )

                worksheet.freeze_panes(1, 0)

        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name=f'relatorio_oci_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
    except Exception as e:
        app.logger.error(f'Erro ao gerar XLSX: {str(e)}')
        return jsonify({'error': f'Erro ao gerar XLSX: {str(e)}'}), 500


@app.route('/download-modelo')
def download_modelo():
    """Baixa arquivo modelo."""
    try:
        return send_from_directory(
            'db', 'arquivo_modelo.xlsx', as_attachment=True
        )
    except FileNotFoundError:
        return jsonify({'error': 'Arquivo modelo não encontrado'}), 404


@app.route('/info')
def info():
    """Retorna informações sobre a aplicação."""
    return jsonify(
        {
            'nome': 'Analisador de Lista de Espera OCI',
            'versao': '2.0.0',
            'autoridades_suportadas': list(app.config['ALLOWED_EXTENSIONS']),
            'tamanho_maximo_mb': app.config['MAX_CONTENT_LENGTH']
            // (1024 * 1024),
            'agrupamentos_total': database.contar_agrupamentos(),
        }
    )


# ==================== ROTAS DE GERENCIAMENTO ====================


@app.route('/admin/agrupamentos')
def admin_agrupamentos():
    """Página de gerenciamento de agrupamentos."""
    agrupamentos = database.listar_agrupamentos_detalhado()
    return render_template(
        'admin_agrupamentos.html', agrupamentos=agrupamentos
    )


@app.route('/admin/visualizar/<codigo>')
def admin_visualizar(codigo):
    """Visualiza detalhes de um agrupamento."""
    agrupamento = database.obter_agrupamento_detalhado(codigo)
    if not agrupamento:
        return jsonify({'error': 'Agrupamento não encontrado'}), 404
    return render_template('admin_visualizar.html', agrupamento=agrupamento)


@app.route('/api/agrupamentos', methods=['GET'])
def api_list_agrupamentos():
    """Retorna lista de agrupamentos em JSON."""
    try:
        agrupamentos = database.listar_agrupamentos_detalhado()
        return jsonify(agrupamentos)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/agrupamentos', methods=['POST'])
def api_criar_agrupamento():
    """Cria um novo agrupamento."""
    try:
        data = request.json

        if not data or not all(k in data for k in ['codigo', 'nome']):
            return jsonify({'error': 'Campos obrigatórios: codigo, nome'}), 400

        # Criar agrupamento
        agrupamento_id = database.adicionar_agrupamento(
            codigo=data['codigo'],
            nome=data['nome'],
            descricao=data.get('descricao', ''),
            itens_obrigatorios=data.get('itens_obrigatorios', []),
            itens_facultativos=data.get('itens_facultativos', []),
        )

        return (
            jsonify(
                {
                    'id': agrupamento_id,
                    'message': 'Agrupamento criado com sucesso',
                }
            ),
            201,
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/agrupamentos/<codigo>', methods=['GET'])
def api_obter_agrupamento(codigo):
    """Obtém um agrupamento específico."""
    try:
        agrupamento = database.obter_agrupamento_detalhado(codigo)
        if not agrupamento:
            return jsonify({'error': 'Agrupamento não encontrado'}), 404
        return jsonify(agrupamento)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/agrupamentos/<codigo>', methods=['PUT'])
def api_atualizar_agrupamento(codigo):
    """Atualiza um agrupamento."""
    try:
        data = request.json

        sucesso = database.atualizar_agrupamento(
            codigo=codigo,
            nome=data['nome'],
            descricao=data.get('descricao', ''),
            itens_obrigatorios=data.get('itens_obrigatorios', []),
            itens_facultativos=data.get('itens_facultativos', []),
        )

        if not sucesso:
            return jsonify({'error': 'Agrupamento não encontrado'}), 404

        return jsonify({'message': 'Agrupamento atualizado com sucesso'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/agrupamentos/<codigo>', methods=['DELETE'])
def api_deletar_agrupamento(codigo):
    """Deleta um agrupamento."""
    try:
        sucesso = database.deletar_agrupamento(codigo)

        if not sucesso:
            return jsonify({'error': 'Agrupamento não encontrado'}), 404

        return jsonify({'message': 'Agrupamento deletado com sucesso'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ==================== ROTAS DE PROCEDIMENTOS ====================


@app.route(
    '/api/agrupamentos/<int:agrupamento_id>/procedimentos-obrigatorios',
    methods=['POST'],
)
def api_criar_procedimento_obrigatorio(agrupamento_id):
    """Cria um procedimento obrigatório."""
    try:
        data = request.json

        if not data or not all(k in data for k in ['codigo', 'descricao']):
            return (
                jsonify({'error': 'Campos obrigatórios: codigo, descricao'}),
                400,
            )

        procedimento_id = database.adicionar_procedimento_obrigatorio(
            agrupamento_id=agrupamento_id,
            codigo=data['codigo'],
            descricao=data['descricao'],
            ordem=data.get('ordem'),
        )

        return (
            jsonify(
                {
                    'id': procedimento_id,
                    'message': 'Procedimento obrigatório criado com sucesso',
                }
            ),
            201,
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route(
    '/api/agrupamentos/<int:agrupamento_id>/procedimentos-facultativos',
    methods=['POST'],
)
def api_criar_procedimento_facultativo(agrupamento_id):
    """Cria um procedimento facultativo."""
    try:
        data = request.json

        if not data or not all(k in data for k in ['codigo', 'descricao']):
            return (
                jsonify({'error': 'Campos obrigatórios: codigo, descricao'}),
                400,
            )

        procedimento_id = database.adicionar_procedimento_facultativo(
            agrupamento_id=agrupamento_id,
            codigo=data['codigo'],
            descricao=data['descricao'],
            ordem=data.get('ordem'),
        )

        return (
            jsonify(
                {
                    'id': procedimento_id,
                    'message': 'Procedimento facultativo criado com sucesso',
                }
            ),
            201,
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route(
    '/api/procedimentos-obrigatorios/<int:procedimento_id>', methods=['PUT']
)
def api_atualizar_procedimento_obrigatorio(procedimento_id):
    """Atualiza um procedimento obrigatório."""
    try:
        data = request.json

        sucesso = database.atualizar_procedimento_obrigatorio(
            procedimento_id=procedimento_id,
            codigo=data['codigo'],
            descricao=data['descricao'],
        )

        if not sucesso:
            return jsonify({'error': 'Procedimento não encontrado'}), 404

        return jsonify(
            {'message': 'Procedimento obrigatório atualizado com sucesso'}
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route(
    '/api/procedimentos-facultativos/<int:procedimento_id>', methods=['PUT']
)
def api_atualizar_procedimento_facultativo(procedimento_id):
    """Atualiza um procedimento facultativo."""
    try:
        data = request.json

        sucesso = database.atualizar_procedimento_facultativo(
            procedimento_id=procedimento_id,
            codigo=data['codigo'],
            descricao=data['descricao'],
        )

        if not sucesso:
            return jsonify({'error': 'Procedimento não encontrado'}), 404

        return jsonify(
            {'message': 'Procedimento facultativo atualizado com sucesso'}
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route(
    '/api/procedimentos-obrigatorios/<int:procedimento_id>', methods=['DELETE']
)
def api_deletar_procedimento_obrigatorio(procedimento_id):
    """Deleta um procedimento obrigatório."""
    try:
        sucesso = database.deletar_procedimento_obrigatorio(procedimento_id)

        if not sucesso:
            return jsonify({'error': 'Procedimento não encontrado'}), 404

        return jsonify(
            {'message': 'Procedimento obrigatório deletado com sucesso'}
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route(
    '/api/procedimentos-facultativos/<int:procedimento_id>', methods=['DELETE']
)
def api_deletar_procedimento_facultativo(procedimento_id):
    """Deleta um procedimento facultativo."""
    try:
        sucesso = database.deletar_procedimento_facultativo(procedimento_id)

        if not sucesso:
            return jsonify({'error': 'Procedimento não encontrado'}), 404

        return jsonify(
            {'message': 'Procedimento facultativo deletado com sucesso'}
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ==================== ROTAS DE EXPORTAÇÃO ====================


@app.route('/api/exportar/json', methods=['GET'])
def api_exportar_json():
    """Exporta agrupamentos em JSON."""
    try:
        conteudo = database.exportar_agrupamentos_json()

        return send_file(
            io.BytesIO(conteudo.encode('utf-8')),
            mimetype='application/json',
            as_attachment=True,
            download_name=f'agrupamentos_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json',
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/exportar/csv', methods=['GET'])
def api_exportar_csv():
    """Exporta agrupamentos em CSV."""
    try:
        conteudo = database.exportar_agrupamentos_csv()

        return send_file(
            io.BytesIO(conteudo.encode('utf-8')),
            mimetype='text/csv',
            as_attachment=True,
            download_name=f'agrupamentos_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ==================== INICIALIZAÇÃO ====================

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs('db', exist_ok=True)
    app.run(debug=True, host='0.0.0.0')

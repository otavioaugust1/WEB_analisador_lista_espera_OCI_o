"""
Módulo de gerenciamento do banco de dados SQLite para agrupamentos OCI.
Fornece funções para inicializar, migrar e consultar dados de agrupamentos.
"""

import sqlite3
from pathlib import Path
from typing import Dict, List, Optional, Tuple

DATABASE_FILE = 'db/agrupamentos.db'


def get_connection():
    """Retorna uma conexão com o banco de dados SQLite."""
    Path('db').mkdir(exist_ok=True)
    conn = sqlite3.connect(DATABASE_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def init_database():
    """Inicializa o banco de dados criando as tabelas necessárias."""
    conn = get_connection()
    cursor = conn.cursor()

    # Tabela de agrupamentos
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS agrupamentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT UNIQUE NOT NULL,
            nome TEXT NOT NULL,
            descricao TEXT,
            criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            atualizado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """
    )

    # Tabela de itens obrigatórios
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS itens_obrigatorios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            agrupamento_id INTEGER NOT NULL,
            codigo TEXT NOT NULL,
            descricao TEXT NOT NULL,
            ordem INTEGER DEFAULT 0,
            FOREIGN KEY (agrupamento_id) REFERENCES agrupamentos(id)
        )
    """
    )

    # Tabela de itens facultativos
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS itens_facultativos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            agrupamento_id INTEGER NOT NULL,
            codigo TEXT NOT NULL,
            descricao TEXT NOT NULL,
            ordem INTEGER DEFAULT 0,
            FOREIGN KEY (agrupamento_id) REFERENCES agrupamentos(id)
        )
    """
    )

    conn.commit()
    conn.close()


def obter_agrupamentos() -> Dict:
    """Retorna todos os agrupamentos em formato compatível com o app anterior."""
    conn = get_connection()
    cursor = conn.cursor()

    agrupamentos = {}

    cursor.execute('SELECT id, codigo, nome FROM agrupamentos ORDER BY codigo')
    for row in cursor.fetchall():
        agrupamento_id = row['id']
        codigo = row['codigo']
        nome = row['nome']

        # Obter itens obrigatórios
        cursor.execute(
            'SELECT codigo, descricao FROM itens_obrigatorios WHERE agrupamento_id = ? ORDER BY ordem',
            (agrupamento_id,),
        )
        itens_obrigatorios = [
            {'codigo': r['codigo'], 'descricao': r['descricao']}
            for r in cursor.fetchall()
        ]

        # Obter itens facultativos
        cursor.execute(
            'SELECT codigo, descricao FROM itens_facultativos WHERE agrupamento_id = ? ORDER BY ordem',
            (agrupamento_id,),
        )
        itens_facultativos = [
            {'codigo': r['codigo'], 'descricao': r['descricao']}
            for r in cursor.fetchall()
        ]

        agrupamentos[codigo] = {
            'nome': nome,
            'itens_obrigatorios': itens_obrigatorios,
            'itens_facultativos': itens_facultativos,
        }

    conn.close()
    return agrupamentos


def obter_agrupamento(codigo: str) -> Optional[Dict]:
    """Retorna um agrupamento específico pelo código."""
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        'SELECT id, codigo, nome FROM agrupamentos WHERE codigo = ?', (codigo,)
    )
    row = cursor.fetchone()

    if not row:
        return None

    agrupamento_id = row['id']

    # Obter itens obrigatórios
    cursor.execute(
        'SELECT codigo, descricao FROM itens_obrigatorios WHERE agrupamento_id = ? ORDER BY ordem',
        (agrupamento_id,),
    )
    itens_obrigatorios = [
        {'codigo': r['codigo'], 'descricao': r['descricao']}
        for r in cursor.fetchall()
    ]

    # Obter itens facultativos
    cursor.execute(
        'SELECT codigo, descricao FROM itens_facultativos WHERE agrupamento_id = ? ORDER BY ordem',
        (agrupamento_id,),
    )
    itens_facultativos = [
        {'codigo': r['codigo'], 'descricao': r['descricao']}
        for r in cursor.fetchall()
    ]

    conn.close()

    return {
        'codigo': codigo,
        'nome': row['nome'],
        'itens_obrigatorios': itens_obrigatorios,
        'itens_facultativos': itens_facultativos,
    }


def adicionar_agrupamento(
    codigo: str,
    nome: str,
    itens_obrigatorios: List[Dict],
    itens_facultativos: List[Dict],
    descricao: str = None,
) -> int:
    """Adiciona um novo agrupamento ao banco de dados."""
    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute(
            'INSERT INTO agrupamentos (codigo, nome, descricao) VALUES (?, ?, ?)',
            (codigo, nome, descricao),
        )
        agrupamento_id = cursor.lastrowid

        # Inserir itens obrigatórios
        for ordem, item in enumerate(itens_obrigatorios):
            cursor.execute(
                'INSERT INTO itens_obrigatorios (agrupamento_id, codigo, descricao, ordem) VALUES (?, ?, ?, ?)',
                (agrupamento_id, item['codigo'], item['descricao'], ordem),
            )

        # Inserir itens facultativos
        for ordem, item in enumerate(itens_facultativos):
            cursor.execute(
                'INSERT INTO itens_facultativos (agrupamento_id, codigo, descricao, ordem) VALUES (?, ?, ?, ?)',
                (agrupamento_id, item['codigo'], item['descricao'], ordem),
            )

        conn.commit()
        return agrupamento_id

    finally:
        conn.close()


def atualizar_agrupamento(
    codigo: str,
    nome: str,
    itens_obrigatorios: List[Dict],
    itens_facultativos: List[Dict],
    descricao: str = None,
) -> bool:
    """Atualiza um agrupamento existente."""
    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute(
            'SELECT id FROM agrupamentos WHERE codigo = ?', (codigo,)
        )
        row = cursor.fetchone()

        if not row:
            return False

        agrupamento_id = row['id']

        # Atualizar agrupamento
        cursor.execute(
            'UPDATE agrupamentos SET nome = ?, descricao = ?, atualizado_em = CURRENT_TIMESTAMP WHERE id = ?',
            (nome, descricao, agrupamento_id),
        )

        # Remover itens antigos
        cursor.execute(
            'DELETE FROM itens_obrigatorios WHERE agrupamento_id = ?',
            (agrupamento_id,),
        )
        cursor.execute(
            'DELETE FROM itens_facultativos WHERE agrupamento_id = ?',
            (agrupamento_id,),
        )

        # Inserir novos itens
        for ordem, item in enumerate(itens_obrigatorios):
            cursor.execute(
                'INSERT INTO itens_obrigatorios (agrupamento_id, codigo, descricao, ordem) VALUES (?, ?, ?, ?)',
                (agrupamento_id, item['codigo'], item['descricao'], ordem),
            )

        for ordem, item in enumerate(itens_facultativos):
            cursor.execute(
                'INSERT INTO itens_facultativos (agrupamento_id, codigo, descricao, ordem) VALUES (?, ?, ?, ?)',
                (agrupamento_id, item['codigo'], item['descricao'], ordem),
            )

        conn.commit()
        return True

    finally:
        conn.close()


def deletar_agrupamento(codigo: str) -> bool:
    """Deleta um agrupamento do banco de dados."""
    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute(
            'SELECT id FROM agrupamentos WHERE codigo = ?', (codigo,)
        )
        row = cursor.fetchone()

        if not row:
            return False

        agrupamento_id = row['id']

        # Deletar itens relacionados
        cursor.execute(
            'DELETE FROM itens_obrigatorios WHERE agrupamento_id = ?',
            (agrupamento_id,),
        )
        cursor.execute(
            'DELETE FROM itens_facultativos WHERE agrupamento_id = ?',
            (agrupamento_id,),
        )

        # Deletar agrupamento
        cursor.execute(
            'DELETE FROM agrupamentos WHERE id = ?', (agrupamento_id,)
        )

        conn.commit()
        return True

    finally:
        conn.close()


def obter_todas_descricoes_sigtap() -> Dict[str, str]:
    """Retorna um dicionário com todos os códigos SIGTAP e suas descrições."""
    conn = get_connection()
    cursor = conn.cursor()

    descricoes = {}

    # Coletando de itens obrigatórios
    cursor.execute('SELECT DISTINCT codigo, descricao FROM itens_obrigatorios')
    for row in cursor.fetchall():
        descricoes[row['codigo']] = row['descricao']

    # Coletando de itens facultativos (se não estiver duplicado)
    cursor.execute('SELECT DISTINCT codigo, descricao FROM itens_facultativos')
    for row in cursor.fetchall():
        if row['codigo'] not in descricoes:
            descricoes[row['codigo']] = row['descricao']

    conn.close()
    return descricoes


def contar_agrupamentos() -> int:
    """Retorna a quantidade de agrupamentos."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT COUNT(*) as count FROM agrupamentos')
    count = cursor.fetchone()['count']
    conn.close()
    return count


# ==================== FUNÇÕES ADICIONAIS PARA GERENCIAMENTO ====================


def listar_agrupamentos_detalhado() -> List[Dict]:
    """Retorna lista de todos os agrupamentos com contagem de procedimentos."""
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT 
            a.id,
            a.codigo,
            a.nome,
            a.descricao,
            a.criado_em,
            a.atualizado_em,
            (SELECT COUNT(*) FROM itens_obrigatorios WHERE agrupamento_id = a.id) as qtd_obrigatorios,
            (SELECT COUNT(*) FROM itens_facultativos WHERE agrupamento_id = a.id) as qtd_facultativos
        FROM agrupamentos a
        ORDER BY a.codigo
    """
    )

    agrupamentos = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return agrupamentos


def obter_agrupamento_detalhado(codigo: str) -> Optional[Dict]:
    """Retorna agrupamento com detalhes completos de procedimentos."""
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        'SELECT id, codigo, nome, descricao FROM agrupamentos WHERE codigo = ?',
        (codigo,),
    )
    row = cursor.fetchone()

    if not row:
        return None

    agrupamento_id = row['id']

    # Obter itens obrigatórios com ID
    cursor.execute(
        """
        SELECT id, codigo, descricao, ordem 
        FROM itens_obrigatorios 
        WHERE agrupamento_id = ? 
        ORDER BY ordem
    """,
        (agrupamento_id,),
    )
    itens_obrigatorios = [dict(r) for r in cursor.fetchall()]

    # Obter itens facultativos com ID
    cursor.execute(
        """
        SELECT id, codigo, descricao, ordem 
        FROM itens_facultativos 
        WHERE agrupamento_id = ? 
        ORDER BY ordem
    """,
        (agrupamento_id,),
    )
    itens_facultativos = [dict(r) for r in cursor.fetchall()]

    conn.close()

    return {
        'id': agrupamento_id,
        'codigo': row['codigo'],
        'nome': row['nome'],
        'descricao': row['descricao'],
        'itens_obrigatorios': itens_obrigatorios,
        'itens_facultativos': itens_facultativos,
    }


# ==================== FUNÇÕES DE PROCEDIMENTOS ====================


def adicionar_procedimento_obrigatorio(
    agrupamento_id: int, codigo: str, descricao: str, ordem: int = None
) -> int:
    """Adiciona um procedimento obrigatório a um agrupamento."""
    conn = get_connection()
    cursor = conn.cursor()

    try:
        if ordem is None:
            cursor.execute(
                'SELECT MAX(ordem) as max_ordem FROM itens_obrigatorios WHERE agrupamento_id = ?',
                (agrupamento_id,),
            )
            result = cursor.fetchone()
            ordem = (result['max_ordem'] or 0) + 1

        cursor.execute(
            """
            INSERT INTO itens_obrigatorios (agrupamento_id, codigo, descricao, ordem)
            VALUES (?, ?, ?, ?)
        """,
            (agrupamento_id, codigo, descricao, ordem),
        )

        conn.commit()
        return cursor.lastrowid
    finally:
        conn.close()


def adicionar_procedimento_facultativo(
    agrupamento_id: int, codigo: str, descricao: str, ordem: int = None
) -> int:
    """Adiciona um procedimento facultativo a um agrupamento."""
    conn = get_connection()
    cursor = conn.cursor()

    try:
        if ordem is None:
            cursor.execute(
                'SELECT MAX(ordem) as max_ordem FROM itens_facultativos WHERE agrupamento_id = ?',
                (agrupamento_id,),
            )
            result = cursor.fetchone()
            ordem = (result['max_ordem'] or 0) + 1

        cursor.execute(
            """
            INSERT INTO itens_facultativos (agrupamento_id, codigo, descricao, ordem)
            VALUES (?, ?, ?, ?)
        """,
            (agrupamento_id, codigo, descricao, ordem),
        )

        conn.commit()
        return cursor.lastrowid
    finally:
        conn.close()


def atualizar_procedimento_obrigatorio(
    procedimento_id: int, codigo: str, descricao: str
) -> bool:
    """Atualiza um procedimento obrigatório."""
    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute(
            """
            UPDATE itens_obrigatorios 
            SET codigo = ?, descricao = ?
            WHERE id = ?
        """,
            (codigo, descricao, procedimento_id),
        )

        conn.commit()
        return cursor.rowcount > 0
    finally:
        conn.close()


def atualizar_procedimento_facultativo(
    procedimento_id: int, codigo: str, descricao: str
) -> bool:
    """Atualiza um procedimento facultativo."""
    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute(
            """
            UPDATE itens_facultativos 
            SET codigo = ?, descricao = ?
            WHERE id = ?
        """,
            (codigo, descricao, procedimento_id),
        )

        conn.commit()
        return cursor.rowcount > 0
    finally:
        conn.close()


def deletar_procedimento_obrigatorio(procedimento_id: int) -> bool:
    """Deleta um procedimento obrigatório."""
    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute(
            'DELETE FROM itens_obrigatorios WHERE id = ?', (procedimento_id,)
        )
        conn.commit()
        return cursor.rowcount > 0
    finally:
        conn.close()


def deletar_procedimento_facultativo(procedimento_id: int) -> bool:
    """Deleta um procedimento facultativo."""
    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute(
            'DELETE FROM itens_facultativos WHERE id = ?', (procedimento_id,)
        )
        conn.commit()
        return cursor.rowcount > 0
    finally:
        conn.close()


def exportar_agrupamentos_json() -> str:
    """Exporta todos os agrupamentos em formato JSON."""
    import json

    agrupamentos = obter_agrupamentos()
    return json.dumps(agrupamentos, ensure_ascii=False, indent=2)


def exportar_agrupamentos_csv() -> str:
    """Exporta agrupamentos em formato CSV."""
    import csv
    import io

    saida = io.StringIO()
    writer = csv.writer(saida)

    # Cabeçalho
    writer.writerow(
        ['CÓDIGO', 'NOME', 'TIPO', 'CÓDIGO SIGTAP', 'DESCRIÇÃO SIGTAP']
    )

    agrupamentos = obter_agrupamentos()

    for codigo, dados in agrupamentos.items():
        nome = dados['nome']

        for item in dados['itens_obrigatorios']:
            writer.writerow(
                [
                    codigo,
                    nome,
                    'OBRIGATÓRIO',
                    item['codigo'],
                    item['descricao'],
                ]
            )

        for item in dados['itens_facultativos']:
            writer.writerow(
                [
                    codigo,
                    nome,
                    'FACULTATIVO',
                    item['codigo'],
                    item['descricao'],
                ]
            )

    return saida.getvalue()

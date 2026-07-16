# -*- coding: utf-8 -*-
"""
Analisador de Lista de Espera - Versão Desktop (PyWebView)
Para usar em Windows ou outros sistemas desktop

Instalação:
    pip install pywebview

Execução:
    python app_desktop.py
"""

import os
import shutil
import sys
import threading
from pathlib import Path

# Definir diretório de dados antes de importar a app
if getattr(sys, 'frozen', False):
    # Executável PyInstaller: usar pasta ao lado do .exe
    _app_dir = Path(sys.executable).parent
    _bundle_dir = Path(sys._MEIPASS)
else:
    _app_dir = Path(__file__).parent
    _bundle_dir = _app_dir

os.chdir(str(_app_dir))

# Se rodando como exe e o banco de dados não existe ainda,
# copiar o banco seed que foi bundled no executável
_db_dest = _app_dir / 'db' / 'agrupamentos.db'
_db_seed = _bundle_dir / 'db' / 'agrupamentos.db'
if not _db_dest.exists() and _db_seed.exists():
    _db_dest.parent.mkdir(exist_ok=True)
    shutil.copy2(str(_db_seed), str(_db_dest))

# Copiar arquivo modelo xlsx se não existir ao lado do exe
_modelo_dest = _app_dir / 'db' / 'arquivo_modelo.xlsx'
_modelo_seed = _bundle_dir / 'db' / 'arquivo_modelo.xlsx'
if not _modelo_dest.exists() and _modelo_seed.exists():
    _modelo_dest.parent.mkdir(exist_ok=True)
    shutil.copy2(str(_modelo_seed), str(_modelo_dest))

import webview

# Importar a aplicação Flask
from app import app


def run_flask():
    """Executa a aplicação Flask em thread separada."""
    app.run(
        debug=False,
        host='127.0.0.1',
        port=5000,
        use_reloader=False,
        use_debugger=False,
        threaded=True,
    )


def main():
    """Inicializa a aplicação desktop com PyWebView."""

    # Criar pastas necessárias
    Path('db').mkdir(exist_ok=True)
    Path('uploads').mkdir(exist_ok=True)

    print('=' * 60)
    print('Analisador de Lista de Espera para OCI - Desktop v2.0.0')
    print('=' * 60)
    print()
    print('Iniciando servidor Flask...')

    # Iniciar Flask em thread separada
    flask_thread = threading.Thread(target=run_flask, daemon=True)
    flask_thread.start()

    print('[OK] Servidor Flask iniciado')
    print()
    print('Abrindo janela da aplicacao...')

    # Criar janela PyWebView
    window = webview.create_window(
        title='Analisador de Lista de Espera OCI',
        url='http://127.0.0.1:5000',
        width=1400,
        height=900,
        min_size=(1000, 700),
        resizable=True,
        fullscreen=False,
    )

    # Iniciar aplicação
    try:
        webview.start()
    except KeyboardInterrupt:
        print('\nAplicacao encerrada pelo usuario')
        sys.exit(0)


if __name__ == '__main__':
    # Verificar se PyWebView está instalado
    try:
        import webview
    except ImportError:
        print('Erro: PyWebView não está instalado!')
        print()
        print('Para usar a versão desktop, instale com:')
        print('  pip install pywebview')
        print()
        print('Ou use a versão web normal:')
        print('  python app.py')
        sys.exit(1)

    main()

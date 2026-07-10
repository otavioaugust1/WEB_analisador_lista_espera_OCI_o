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
import sys
import threading
from pathlib import Path

# Definir diretório de dados antes de importar a app
if getattr(sys, 'frozen', False):
    # Executável PyInstaller: usar pasta ao lado do .exe
    _app_dir = Path(sys.executable).parent
else:
    _app_dir = Path(__file__).parent

os.chdir(str(_app_dir))

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

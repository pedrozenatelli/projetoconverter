import os
import subprocess
import sys
import threading

import webview


def run_server():
    # Ajuste o comando abaixo para rodar com suas configurações
    os.environ['DJANGO_SETTINGS_MODULE'] = 'core.settings'
    subprocess.call([sys.executable, 'manage.py', 'runserver'])


# Roda o servidor em uma thread separada
t = threading.Thread(target=run_server)
t.daemon = True
t.start()

# Crie uma janela usando pywebview
webview.create_window("My Django App", "http://127.0.0.1:8000/")
webview.start()

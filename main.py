import os
import sys
import json

# Ajouter le venv local au sys.path pour garantir l'accès aux packages
# (utile si l'app est lancée avec le Python système plutôt que le venv)
_base = os.path.dirname(os.path.abspath(__file__))
_venv_site = os.path.join(_base, 'venv', 'Lib', 'site-packages')
if os.path.isdir(_venv_site) and _venv_site not in sys.path:
    sys.path.insert(0, _venv_site)

import webview
from api import FileAPI


if __name__ == '__main__':
    args = sys.argv[1:]
    debug = 'debug=True' in args
    initial_file = next((a for a in args if a != 'debug=True' and os.path.isfile(a)), None)

    file_api = FileAPI()
    html_path = os.path.join(_base, 'frontend', 'index.html')
    window = webview.create_window(
        'Éditeur Markdown',
        f'file://{html_path}',
        js_api=file_api,
        width=1280,
        height=800,
        min_size=(800, 600),
    )

    if initial_file:
        def on_loaded():
            path = os.path.abspath(initial_file)
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    content = f.read()
                data = json.dumps({'path': path, 'content': content, 'name': os.path.basename(path)})
                window.evaluate_js(f'openInitialFile({data})')
            except Exception as e:
                print(f'Erreur ouverture fichier initial : {e}')
        window.events.loaded += on_loaded

    webview.start(debug=debug)

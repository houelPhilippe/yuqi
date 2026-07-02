import os
import json
import shutil
import webview

try:
    from spellchecker import SpellChecker as _SpellChecker
except ImportError:
    _SpellChecker = None

_config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
_settings_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'settings.json')
_readme_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'README.md')

_MAX_RECENT = 10

# ── Correcteur orthographique ──────────────────────────────────────────────────
_spell_checkers = {}   # {lang: SpellChecker}
_ignored_words  = set()

def _get_spell_checker(lang):
    if lang not in _spell_checkers:
        if _SpellChecker is None:
            _spell_checkers[lang] = None
        else:
            try:
                _spell_checkers[lang] = _SpellChecker(language=lang)
            except Exception as e:
                print(f'[spell] _get_spell_checker({lang!r}) failed: {e}')
                _spell_checkers[lang] = None
    return _spell_checkers[lang]


class FileAPI:
    def __init__(self):
        self._project_dir = None   # répertoire projet actif (défini à l'exécution)

    def _load_config(self):
        try:
            with open(_config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    def _save_config(self, cfg):
        try:
            with open(_config_path, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ── Gestion du projet ──────────────────────────────────────────────────────

    def _project_settings_path(self):
        if self._project_dir and os.path.isdir(self._project_dir):
            return os.path.join(self._project_dir, 'settings.json')
        return _settings_path

    def get_current_project(self):
        """Retourne le projet actif depuis config.json."""
        cfg = self._load_config()
        last = cfg.get('lastProjectDir', '')
        recent = cfg.get('recentProjects', [])
        if last and os.path.isdir(last):
            self._project_dir = last
            return {'dir': last, 'exists': True, 'recent': recent}
        self._project_dir = None
        return {'dir': last, 'exists': False, 'recent': recent}

    def _activate_project(self, path):
        """Définit le répertoire projet actif et met à jour config.json."""
        self._project_dir = path
        cfg = self._load_config()
        cfg['lastProjectDir'] = path
        recent = [r for r in cfg.get('recentProjects', []) if r != path]
        recent.insert(0, path)
        cfg['recentProjects'] = recent[:_MAX_RECENT]
        self._save_config(cfg)
        return {
            'dir': path,
            'settings': self.load_settings(),
            'recent': cfg['recentProjects'],
        }

    def choose_project_dir(self):
        """Ouvre un sélecteur de dossier pour choisir le répertoire projet."""
        try:
            dialog_type = webview.FileDialog.FOLDER
        except AttributeError:
            dialog_type = webview.FOLDER_DIALOG
        result = webview.windows[0].create_file_dialog(dialog_type)
        if not result:
            return None
        path = result[0] if isinstance(result, (list, tuple)) else result
        return self._activate_project(path)

    def open_recent_project(self, path):
        """Ouvre un projet récent par chemin."""
        if not os.path.isdir(path):
            return {'error': f'Répertoire introuvable : {path}'}
        return self._activate_project(path)

    # ── Paramètres ─────────────────────────────────────────────────────────────

    def save_settings(self, data):
        try:
            with open(self._project_settings_path(), 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return {'success': True}
        except Exception as e:
            return {'error': str(e)}

    def load_settings(self):
        try:
            with open(self._project_settings_path(), 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    def compile_project(self, format, project_dir):
        """Compile le projet global Quarto (quarto render --to html/pdf ...)."""
        import subprocess, threading, json, re
        ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')

        def run():
            try:
                profile = 'pdf' if format == 'pdf' else ('word' if format == 'docx' else 'html')
                cmd = ['quarto', 'render', '--to', format,
                       '--profile', profile, '--no-clean']
                proc = subprocess.Popen(
                    cmd, cwd=project_dir,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding='utf-8', errors='replace'
                )
                for line in proc.stdout:
                    clean = ansi_escape.sub('', line.rstrip('\r\n'))
                    webview.windows[0].evaluate_js(
                        f'window.appendCompileOutput({json.dumps(clean)})'
                    )
                proc.wait()
                webview.windows[0].evaluate_js(f'window.compileFinished({proc.returncode})')
            except Exception as e:
                webview.windows[0].evaluate_js(
                    f'window.compileFinished(-1, {json.dumps(str(e))})'
                )

        threading.Thread(target=run, daemon=True).start()
        return {'started': True}

    def compile_quarto(self, file_path, project_dir, cmd_template=''):
        import subprocess, threading, json, re, shlex
        ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')

        def run():
            try:
                if project_dir and os.path.isdir(project_dir):
                    cwd = project_dir
                    rel = os.path.relpath(file_path, project_dir).replace('\\', '/')
                else:
                    cwd = os.path.dirname(file_path)
                    rel = os.path.basename(file_path)
                if cmd_template:
                    cmd = shlex.split(cmd_template.replace('{file}', rel))
                else:
                    cmd = ['quarto', 'render', rel,
                           '--to', 'html', '--profile', 'html', '--no-clean']
                proc = subprocess.Popen(
                    cmd, cwd=cwd,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding='utf-8', errors='replace'
                )
                for line in proc.stdout:
                    clean = ansi_escape.sub('', line.rstrip('\r\n'))
                    webview.windows[0].evaluate_js(
                        f'window.appendCompileOutput({json.dumps(clean)})'
                    )
                proc.wait()
                webview.windows[0].evaluate_js(f'window.compileFinished({proc.returncode})')
            except Exception as e:
                webview.windows[0].evaluate_js(
                    f'window.compileFinished(-1, {json.dumps(str(e))})'
                )

        threading.Thread(target=run, daemon=True).start()
        return {'started': True}

    def check_quarto_yml(self, file_path):
        """Vérifie si _quarto.yml existe dans le répertoire du fichier."""
        dir_path = os.path.dirname(file_path)
        exists = os.path.isfile(os.path.join(dir_path, '_quarto.yml'))
        return {'exists': exists}

    def compile_quarto_pdf(self, file_path, project_dir=None, cmd_template=''):
        """Lance quarto render [fichier] --to pdf depuis le répertoire du fichier (ou du projet)."""
        import subprocess, threading, json, re, shlex
        ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')

        def run():
            try:
                if project_dir and os.path.isdir(project_dir):
                    cwd = project_dir
                    rel = os.path.relpath(file_path, project_dir).replace('\\', '/')
                else:
                    cwd = os.path.dirname(file_path)
                    rel = os.path.basename(file_path)
                if cmd_template:
                    cmd = shlex.split(cmd_template.replace('{file}', rel))
                else:
                    cmd = ['quarto', 'render', rel, '--to', 'pdf']
                proc = subprocess.Popen(
                    cmd, cwd=cwd,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding='utf-8', errors='replace'
                )
                for line in proc.stdout:
                    clean = ansi_escape.sub('', line.rstrip('\r\n'))
                    webview.windows[0].evaluate_js(
                        f'window.appendCompileOutput({json.dumps(clean)})'
                    )
                proc.wait()
                webview.windows[0].evaluate_js(f'window.compileFinished({proc.returncode})')
            except Exception as e:
                webview.windows[0].evaluate_js(
                    f'window.compileFinished(-1, {json.dumps(str(e))})'
                )

        threading.Thread(target=run, daemon=True).start()
        return {'started': True}

    def compile_directory(self, dir_path, parent_dir):
        """Compile tous les .qmd d'un répertoire : quarto render .\\dirname\\*.qmd --to html --profile html --no-clean"""
        import subprocess, threading, json, re
        ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')

        def run():
            try:
                dir_name = os.path.basename(dir_path)
                cwd = parent_dir if parent_dir and os.path.isdir(parent_dir) else os.path.dirname(dir_path)
                glob_pattern = os.path.join('.', dir_name, '*.qmd')
                cmd = ['quarto', 'render', glob_pattern,
                       '--to', 'html', '--profile', 'html', '--no-clean']
                proc = subprocess.Popen(
                    cmd, cwd=cwd,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding='utf-8', errors='replace'
                )
                for line in proc.stdout:
                    clean = ansi_escape.sub('', line.rstrip('\r\n'))
                    webview.windows[0].evaluate_js(
                        f'window.appendCompileOutput({json.dumps(clean)})'
                    )
                proc.wait()
                webview.windows[0].evaluate_js(f'window.compileFinished({proc.returncode})')
            except Exception as e:
                webview.windows[0].evaluate_js(
                    f'window.compileFinished(-1, {json.dumps(str(e))})'
                )

        threading.Thread(target=run, daemon=True).start()
        return {'started': True}

    def _find_plantuml_cmd(self):
        """Retourne (cmd, diagnostics) — cmd est None si introuvable."""
        import shutil
        diag = []

        # 1. Commande 'plantuml' dans le PATH
        p = shutil.which('plantuml')
        if p:
            diag.append(f'✓ Commande plantuml trouvée : {p}')
            return ['plantuml'], diag
        diag.append('✗ Commande plantuml absente du PATH')

        # 2. JAR configuré dans les paramètres
        settings = self.load_settings()
        jar = settings.get('plantumlJar', '').strip()
        if jar:
            if os.path.isfile(jar):
                java = shutil.which('java')
                if java:
                    diag.append(f'✓ JAR configuré : {jar}')
                    diag.append(f'✓ java trouvé : {java}')
                    return [java, '-jar', jar], diag
                else:
                    diag.append(f'✗ JAR configuré ({jar}) mais java introuvable dans le PATH')
            else:
                diag.append(f'✗ JAR configuré mais fichier absent : {jar}')
        else:
            diag.append('✗ Aucun JAR configuré dans les paramètres')

        # 3. Emplacements courants Windows
        common_jars = [
            r'C:\ProgramData\chocolatey\lib\plantuml\tools\plantuml.jar',
            r'C:\tools\plantuml\plantuml.jar',
            os.path.expanduser(r'~\scoop\apps\plantuml\current\plantuml.jar'),
            os.path.expanduser(r'~\plantuml\plantuml.jar'),
        ]
        java = shutil.which('java')
        for candidate in common_jars:
            if os.path.isfile(candidate):
                if java:
                    diag.append(f'✓ JAR trouvé automatiquement : {candidate}')
                    return [java, '-jar', candidate], diag
                else:
                    diag.append(f'✗ JAR trouvé ({candidate}) mais java introuvable')
                    break

        diag.append('')
        diag.append('Pour résoudre :')
        diag.append('  • Option 1 — Installer PlantUML via Chocolatey :')
        diag.append('      choco install plantuml')
        diag.append('  • Option 2 — Télécharger plantuml.jar sur plantuml.com')
        diag.append('      puis le configurer dans Paramètres → PlantUML')
        diag.append('  • Option 3 — Installer Java (JRE 8+) si absent du PATH')
        return None, diag

    def compile_puml(self, file_path, fmt):
        """Compile un fichier .puml vers SVG et/ou PNG via PlantUML.
        fmt : 'svg' | 'png' | 'both'
        """
        import subprocess, threading, json, re
        ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')

        def run():
            cmd_base, diag = self._find_plantuml_cmd()
            # Afficher le diagnostic dans tous les cas
            for line in diag:
                webview.windows[0].evaluate_js(
                    f'window.appendCompileOutput({json.dumps(line)})'
                )
            if not cmd_base:
                webview.windows[0].evaluate_js('window.compileFinished(-1, "PlantUML introuvable")')
                return

            formats = ['svg', 'png'] if fmt == 'both' else [fmt]
            out_dir  = os.path.dirname(file_path)
            last_rc  = 0

            for f in formats:
                cmd = cmd_base + [f'-t{f}', '-o', out_dir, file_path]
                try:
                    proc = subprocess.Popen(
                        cmd, cwd=out_dir,
                        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                        text=True, encoding='utf-8', errors='replace'
                    )
                    for line in proc.stdout:
                        clean = ansi_escape.sub('', line.rstrip('\r\n'))
                        if clean:
                            webview.windows[0].evaluate_js(
                                f'window.appendCompileOutput({json.dumps(clean)})'
                            )
                    proc.wait()
                    last_rc = proc.returncode
                    status = f'✓ {f.upper()} généré' if proc.returncode == 0 else f'✗ Erreur {f.upper()}'
                    webview.windows[0].evaluate_js(
                        f'window.appendCompileOutput({json.dumps(status)})'
                    )
                except Exception as e:
                    last_rc = -1
                    webview.windows[0].evaluate_js(
                        f'window.appendCompileOutput({json.dumps("Erreur : " + str(e))})'
                    )

            webview.windows[0].evaluate_js(f'window.compileFinished({last_rc})')

        threading.Thread(target=run, daemon=True).start()
        return {'started': True}

    def browse_plantuml_jar(self):
        """Ouvre un sélecteur de fichier pour choisir plantuml.jar."""
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=('JAR files (*.jar)', 'All files (*.*)'),
        )
        if not result:
            return None
        return result[0] if isinstance(result, (list, tuple)) else result

    def browse_bib_file(self):
        """Ouvre un sélecteur de fichier pour choisir un fichier BibTeX (.bib)."""
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=('BibTeX files (*.bib)', 'All files (*.*)'),
        )
        if not result:
            return None
        return result[0] if isinstance(result, (list, tuple)) else result

    def parse_bib_file(self, path):
        """Parse un fichier BibTeX et retourne la liste des entrées."""
        import re
        try:
            with open(path, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read()
            entries = []
            for m in re.finditer(
                r'@(\w+)\s*\{\s*([^,\s]+)\s*,([^@]*?)(?=@\w|\Z)',
                content, re.DOTALL
            ):
                etype = m.group(1).lower()
                if etype in ('comment', 'string', 'preamble'):
                    continue
                key = m.group(2).strip()
                body = m.group(3)
                fields = {}
                for fm in re.finditer(
                    r'(\w+)\s*=\s*(?:\{((?:[^{}]|\{[^{}]*\})*)\}|"([^"]*)")',
                    body, re.DOTALL
                ):
                    fname = fm.group(1).lower()
                    fval = (fm.group(2) or fm.group(3) or '').strip()
                    fval = re.sub(r'\{([^{}]*)\}', r'\1', fval)
                    fval = re.sub(r'\s+', ' ', fval).strip()
                    fields[fname] = fval
                entries.append({
                    'key':       key,
                    'type':      etype,
                    'title':     fields.get('title',     ''),
                    'author':    fields.get('author',    ''),
                    'year':      fields.get('year',      ''),
                    'journal':   fields.get('journal',   ''),
                    'booktitle': fields.get('booktitle', ''),
                    'publisher': fields.get('publisher', ''),
                    'volume':    fields.get('volume',    ''),
                    'number':    fields.get('number',    ''),
                    'pages':     fields.get('pages',     ''),
                    'doi':       fields.get('doi',       ''),
                    'url':       fields.get('url',       ''),
                    'note':      fields.get('note',      ''),
                })
            return entries
        except Exception as e:
            return {'error': str(e)}

    def get_default_dir(self):
        cfg = self._load_config()
        path = cfg.get('default_dir', os.path.expanduser('~'))
        if not os.path.isdir(path):
            path = os.path.expanduser('~')
        return path

    def set_default_dir(self, path):
        cfg = self._load_config()
        cfg['default_dir'] = path
        self._save_config(cfg)
        return {'success': True}

    def choose_directory(self):
        try:
            dialog_type = webview.FileDialog.FOLDER
        except AttributeError:
            dialog_type = webview.FOLDER_DIALOG
        result = webview.windows[0].create_file_dialog(dialog_type)
        if not result:
            return None
        return result[0] if isinstance(result, (list, tuple)) else result

    def list_directory(self, path):
        try:
            items = []
            for e in os.scandir(path):
                if e.name.startswith('.'):
                    continue
                is_dir = e.is_dir()
                ext = os.path.splitext(e.name)[1].lower() if not is_dir else ''
                try:
                    mtime = e.stat().st_mtime
                except OSError:
                    mtime = 0
                items.append({
                    'name': e.name,
                    'path': e.path.replace('\\', '/'),
                    'is_dir': is_dir,
                    'ext': ext,
                    'mtime': mtime,
                })
            items.sort(key=lambda x: (not x['is_dir'], x['name'].lower()))
            parent = os.path.dirname(path).replace('\\', '/')
            return {'path': path.replace('\\', '/'), 'parent': parent, 'items': items}
        except Exception as e:
            return {'error': str(e)}

    def open_file_by_path(self, path):
        try:
            _BINARY_EXTS = {'.pdf', '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.svg', '.ico'}
            if os.path.splitext(path)[1].lower() in _BINARY_EXTS:
                return {'path': path, 'content': '', 'name': os.path.basename(path)}
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()
            return {'path': path, 'content': content, 'name': os.path.basename(path)}
        except Exception as e:
            return {'error': str(e)}

    def open_file(self):
        file_types = (
            'Fichiers supportés (*.md;*.qmd;*.html;*.htm;*.pdf;*.yml;*.yaml;*.css;*.scss;*.less;*.lua;*.json;*.xml;*.log;*.tex;*.puml;*.plantuml;*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp;*.svg;*.ico)',
            'All files (*.*)'
        )
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=file_types,
        )
        if not result:
            return None
        path = result[0] if isinstance(result, (list, tuple)) else result
        try:
            _BINARY_EXTS = {'.pdf', '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.svg', '.ico'}
            if os.path.splitext(path)[1].lower() in _BINARY_EXTS:
                return {'path': path, 'content': '', 'name': os.path.basename(path)}
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()
            return {'path': path, 'content': content, 'name': os.path.basename(path)}
        except Exception as e:
            return {'error': str(e)}

    def pick_include_file(self, start_dir, base_dir):
        """Ouvre un sélecteur de fichier .md/.qmd.
        start_dir : répertoire d'ouverture du dialogue (répertoire projet).
        base_dir  : répertoire de base pour le calcul du chemin relatif (répertoire du fichier en cours).
        """
        file_types = ('Markdown Files (*.md;*.qmd)', 'All files (*.*)')
        kwargs = dict(
            dialog_type=webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=file_types,
        )
        if start_dir and os.path.isdir(start_dir):
            kwargs['directory'] = start_dir
        result = webview.windows[0].create_file_dialog(**kwargs)
        if not result:
            return None
        path = result[0] if isinstance(result, (list, tuple)) else result
        try:
            ref = base_dir if (base_dir and os.path.isdir(base_dir)) else start_dir
            rel = os.path.relpath(path, ref).replace('\\', '/') if ref else os.path.basename(path)
        except ValueError:
            rel = os.path.basename(path)
        return {'path': path, 'rel': rel}

    def get_css_classes(self, css_path):
        """Extrait les noms de classes CSS définies dans un fichier .css."""
        import re
        try:
            with open(css_path, 'r', encoding='utf-8') as f:
                content = f.read()
            # Retire les commentaires /* ... */
            content = re.sub(r'/\*.*?\*/', '', content, flags=re.DOTALL)
            # Collecte les .classname dans les sélecteurs (avant le premier {)
            classes = set()
            for selector_block in re.split(r'\{[^}]*\}', content):
                for cls in re.findall(r'\.([a-zA-Z][a-zA-Z0-9_-]*)', selector_block):
                    classes.add(cls)
            return sorted(classes)
        except Exception:
            return []

    def search_in_files(self, directory, query, max_matches_per_file=10):
        """Recherche une chaîne dans tous les fichiers texte du répertoire (récursif)."""
        import os
        TEXT_EXTS = {'.md', '.qmd', '.yaml', '.yml', '.log', '.txt',
                     '.html', '.htm', '.css', '.lua', '.js', '.py', '.json', '.xml', '.csv', '.tex'}
        results = []
        if not query or not os.path.isdir(directory):
            return results
        ql = query.lower()
        try:
            for root, dirs, files in os.walk(directory):
                dirs[:] = sorted(d for d in dirs if not d.startswith('.'))
                for filename in sorted(files):
                    if filename.startswith('.'):
                        continue
                    ext = os.path.splitext(filename)[1].lower()
                    if ext not in TEXT_EXTS:
                        continue
                    filepath = os.path.join(root, filename)
                    rel = os.path.relpath(filepath, directory).replace('\\', '/')
                    try:
                        with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
                            lines = f.readlines()
                        matches = []
                        for i, line in enumerate(lines):
                            if ql in line.lower():
                                matches.append({'line': i + 1, 'text': line.rstrip('\n')})
                                if len(matches) >= max_matches_per_file:
                                    break
                        if matches:
                            results.append({
                                'path': filepath.replace('\\', '/'),
                                'name': filename,
                                'rel':  rel,
                                'matches': matches,
                            })
                    except Exception:
                        pass
        except Exception:
            pass
        return results

    def search_files_by_name(self, directory, query, max_results=200):
        """Recherche les fichiers dont le nom contient la chaîne donnée (récursif)."""
        results = []
        if not query or not os.path.isdir(directory):
            return results
        ql = query.lower()
        try:
            for root, dirs, files in os.walk(directory):
                dirs[:] = sorted(d for d in dirs if not d.startswith('.'))
                for filename in sorted(files):
                    if filename.startswith('.'):
                        continue
                    if ql not in filename.lower():
                        continue
                    filepath = os.path.join(root, filename)
                    rel = os.path.relpath(filepath, directory).replace('\\', '/')
                    results.append({
                        'path': filepath.replace('\\', '/'),
                        'name': filename,
                        'rel':  rel,
                    })
                    if len(results) >= max_results:
                        return results
        except Exception:
            pass
        return results

    def rename_file(self, old_path, new_name):
        try:
            dir_path = os.path.dirname(old_path)
            new_path = os.path.join(dir_path, new_name)
            os.rename(old_path, new_path)
            return {'success': True, 'new_path': new_path.replace('\\', '/')}
        except Exception as e:
            return {'error': str(e)}

    def save_file(self, path, content):
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(content)
            return {'success': True}
        except Exception as e:
            return {'error': str(e)}

    def open_image(self):
        file_types = ('Images (*.png;*.jpg;*.jpeg;*.gif;*.webp;*.svg;*.bmp)', 'All files (*.*)')
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=file_types,
        )
        if not result:
            return None
        path = result[0] if isinstance(result, (list, tuple)) else result
        return {'path': path, 'name': os.path.basename(path)}

    def open_url(self, url):
        """Ouvre une URL dans le navigateur par défaut du système."""
        import subprocess, sys
        try:
            if sys.platform == 'win32':
                os.startfile(url)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', url])
            else:
                subprocess.Popen(['xdg-open', url])
            return {'success': True}
        except Exception as e:
            return {'error': str(e)}

    def duplicate_file(self, path):
        try:
            base, ext = os.path.splitext(path)
            new_path = base + '_copie' + ext
            counter = 1
            while os.path.exists(new_path):
                new_path = base + f'_copie{counter}' + ext
                counter += 1
            shutil.copy2(path, new_path)
            return {'path': new_path.replace('\\', '/'), 'name': os.path.basename(new_path)}
        except Exception as e:
            return {'error': str(e)}

    def delete_file(self, path):
        try:
            os.remove(path)
            return {'success': True}
        except Exception as e:
            return {'error': str(e)}

    def delete_dir(self, path):
        try:
            shutil.rmtree(path)
            return {'success': True}
        except Exception as e:
            return {'error': str(e)}

    def create_directory(self, parent_path, name):
        try:
            new_path = os.path.join(parent_path, name)
            os.makedirs(new_path, exist_ok=False)
            return {'success': True, 'path': new_path.replace('\\', '/')}
        except FileExistsError:
            return {'error': 'Un répertoire portant ce nom existe déjà.'}
        except Exception as e:
            return {'error': str(e)}

    def get_spell_suggestions(self, word, lang='fr'):
        clean = word.strip(".,;:!?\"'«»()[]{}…–—").lower()
        if not clean or clean in _ignored_words:
            return {'correct': True, 'suggestions': []}
        spell = _get_spell_checker(lang)
        if spell is None:
            # Retenter une fois (évite le cache None d'une erreur transitoire)
            _spell_checkers.pop(lang, None)
            spell = _get_spell_checker(lang)
        if spell is None:
            return {'correct': True, 'suggestions': [], 'error': 'pyspellchecker non disponible pour ' + lang}
        try:
            misspelled = spell.unknown([clean])
            if not misspelled:
                return {'correct': True, 'suggestions': []}
            candidates = spell.candidates(clean)
            suggestions = sorted(candidates - {clean})[:8] if candidates else []
            # Restaurer la casse si le mot original commence par une majuscule
            if word and word[0].isupper():
                suggestions = [s.capitalize() for s in suggestions]
            return {'correct': False, 'suggestions': suggestions}
        except Exception as e:
            return {'correct': True, 'suggestions': [], 'error': str(e)}

    def ignore_word(self, word):
        clean = word.lower()
        _ignored_words.add(clean)
        # Enseigner le mot à tous les correcteurs actifs
        for spell in _spell_checkers.values():
            if spell is not None:
                spell.word_frequency.load_words([clean])
        return {'success': True}

    def scan_todos(self, directory):
        """Scanne récursivement les *.md et *.qmd pour les <!-- TODO: ... -->.
        Fusionne avec todo_tracking.json existant pour préserver l'état 'done'.
        Sauvegarde le résultat dans todo_tracking.json."""
        import re
        TODO_RE = re.compile(r'<!--\s*(TODO\s*:?\s*[\s\S]*?)-->', re.IGNORECASE)
        json_path = os.path.join(directory, 'todo_tracking.json')

        # Charger le suivi existant {rel||text: done}
        tracking = {}
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                existing = json.load(f)
            for item in existing.get('todos', []):
                key = item.get('rel', '') + '||' + item.get('text', '')
                tracking[key] = bool(item.get('done', False))
        except Exception:
            pass

        todos = []
        skip_dirs = {'.git', 'venv', 'node_modules', '__pycache__', '.quarto'}
        try:
            for root, dirs, files in os.walk(directory):
                dirs[:] = sorted(d for d in dirs
                                 if not d.startswith('.') and d not in skip_dirs)
                for filename in sorted(files):
                    ext = os.path.splitext(filename)[1].lower()
                    if ext not in ('.md', '.qmd'):
                        continue
                    filepath = os.path.join(root, filename)
                    rel = os.path.relpath(filepath, directory).replace('\\', '/')
                    try:
                        with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
                            content = f.read()
                        for m in TODO_RE.finditer(content):
                            raw = m.group(1).strip()
                            if not re.match(r'^TODO\s*:?', raw, re.IGNORECASE):
                                continue
                            text = re.sub(r'^TODO\s*:?\s*', '', raw,
                                          flags=re.IGNORECASE).strip() or '(sans texte)'
                            key = rel + '||' + text
                            todos.append({
                                'path': filepath.replace('\\', '/'),
                                'rel':  rel,
                                'name': filename,
                                'text': text,
                                'done': tracking.get(key, False),
                            })
                    except Exception:
                        pass
        except Exception as e:
            return {'error': str(e)}

        result = {
            'dir':      directory.replace('\\', '/'),
            'todos':    todos,
        }
        try:
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

        result['json_path'] = json_path.replace('\\', '/')
        return result

    def load_todo_json(self, directory):
        """Charge todo_tracking.json depuis le répertoire donné sans re-scanner."""
        json_path = os.path.join(directory, 'todo_tracking.json')
        if not os.path.isfile(json_path):
            return None
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            data['json_path'] = json_path.replace('\\', '/')
            data['dir'] = data.get('dir', directory.replace('\\', '/'))
            return data
        except Exception as e:
            return {'error': str(e)}

    def save_todo_tracking(self, json_path, todos):
        """Persiste la liste de todos (avec leur état done) dans le fichier JSON."""
        try:
            data = {}
            try:
                with open(json_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            except Exception:
                pass
            data['todos'] = todos
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return {'success': True}
        except Exception as e:
            return {'error': str(e)}

    def read_readme(self):
        try:
            with open(_readme_path, 'r', encoding='utf-8') as f:
                return {'content': f.read()}
        except Exception as e:
            return {'error': str(e)}

    # ── Conversion XML → JSON ──────────────────────────────────────────────────

    def browse_xsd_for_xml(self, initial_dir=''):
        """Ouvre un sélecteur de fichier pour choisir un fichier XSD."""
        kwargs = dict(
            dialog_type=webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=('XSD files (*.xsd)', 'All files (*.*)'),
        )
        if initial_dir and os.path.isdir(initial_dir):
            kwargs['directory'] = initial_dir
        result = webview.windows[0].create_file_dialog(**kwargs)
        if not result:
            return None
        return result[0] if isinstance(result, (list, tuple)) else result

    def run_xml_to_json(self, xml_path, xsd_path=''):
        """Convertit un fichier XML en JSON (JSON sauvegardé dans le même répertoire que le XML).
        Streame la sortie vers window.xmlConvertLog et window.xmlConvertFinished."""
        import threading
        from convert_xml_to_json import run_conversion

        def run():
            try:
                output_dir = os.path.dirname(os.path.abspath(xml_path))
                json_name = os.path.basename(xml_path)
                xsd = xsd_path.strip() if xsd_path else None

                md_path = None
                if xsd:
                    md_name = os.path.splitext(json_name)[0] + "_ordre.md"
                    md_path = os.path.join(output_dir, md_name)

                def log(msg):
                    webview.windows[0].evaluate_js(
                        f'window.xmlConvertLog({json.dumps(msg)})'
                    )

                ok = run_conversion(
                    xml_path, json_name, xsd,
                    output_dir=output_dir, log=log,
                )
                md_exists = md_path and os.path.isfile(md_path)
                webview.windows[0].evaluate_js(
                    f'window.xmlConvertFinished({json.dumps(ok)}, {json.dumps(md_path if md_exists else None)})'
                )
            except Exception as exc:
                webview.windows[0].evaluate_js(
                    f'window.xmlConvertLog({json.dumps("ERREUR : " + str(exc))})'
                )
                webview.windows[0].evaluate_js('window.xmlConvertFinished(false, null)')

        threading.Thread(target=run, daemon=True).start()
        return {'started': True}

    def minimize(self):
        webview.windows[0].minimize()

    def toggle_fullscreen(self):
        webview.windows[0].toggle_fullscreen()

    def save_pasted_image(self, data_base64, ext, initial_dir=''):
        import base64 as _b64
        ext = ext.replace('jpeg', 'jpg') or 'png'
        file_types = (f'Images (*.{ext})', 'All files (*.*)')
        result = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG,
            directory=initial_dir if initial_dir and os.path.isdir(initial_dir) else '',
            save_filename=f'image.{ext}',
            file_types=file_types,
        )
        if not result:
            return None
        path = result[0] if isinstance(result, (list, tuple)) else result
        if not path.lower().endswith(f'.{ext}'):
            path += f'.{ext}'
        try:
            data = _b64.b64decode(data_base64)
            with open(path, 'wb') as f:
                f.write(data)
            return {'path': path, 'name': os.path.basename(path)}
        except Exception as e:
            return {'error': str(e)}

    def read_image_base64(self, path):
        """Lit un fichier image sur le disque et le renvoie encodé en base64.
        Utilisé pour copier une image dans le presse-papiers depuis le JS,
        car fetch()/canvas sur une URL file:// n'est pas fiable selon le
        moteur webview (restrictions cross-origin variables)."""
        import base64 as _b64
        import mimetypes
        try:
            if not os.path.isfile(path):
                return {'error': 'Fichier introuvable'}
            mime, _ = mimetypes.guess_type(path)
            with open(path, 'rb') as f:
                data = f.read()
            return {'data': _b64.b64encode(data).decode('ascii'), 'mime': mime or 'application/octet-stream'}
        except Exception as e:
            return {'error': str(e)}

    def quit(self):
        webview.windows[0].destroy()

    def save_file_as(self, content, current_name='untitled.md'):
        result = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename=current_name,
        )
        if not result:
            return None
        path = result[0] if isinstance(result, (list, tuple)) else result
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(content)
            return {'path': path, 'name': os.path.basename(path)}
        except Exception as e:
            return {'error': str(e)}

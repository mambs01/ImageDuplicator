"""
Web GUI for DuplicateFinder.py.

Usage:
    python src/app.py
    Opens automatically in your browser.

The CLI is still fully usable:
    python src/DuplicateFinder.py -p <photos> -c <chat> -o <output> [-s <start> -e <end>]
"""

import json
import os
import platform
import re
import socket
import subprocess
import sys
import threading
import webbrowser

from flask import Flask, Response, jsonify, render_template, request

app = Flask(__name__)

ANSI_ESCAPE = re.compile(r'\x1b\[[0-9;]*m')  # matches ANSI terminal color codes, e.g. \x1b[32m (green) or \x1b[0m (reset)


# ── Native file / folder pickers ─────────────────────────────────────────────

def _mac_pick(script):
    """Run an AppleScript picker; return POSIX path string or '' on cancel."""
    result = subprocess.run(['osascript', '-e', script],
                            capture_output=True, text=True, timeout=60)
    return result.stdout.strip() if result.returncode == 0 else ''


def _zenity(*args):
    """Try zenity; return path string, '' on cancel, or None if not installed."""
    try:
        result = subprocess.run(['zenity', *args],
                                capture_output=True, text=True, timeout=120)
        return result.stdout.strip() if result.returncode == 0 else ''
    except FileNotFoundError:
        return None
    except Exception:
        return ''


def _tk_directory(title):
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    path = filedialog.askdirectory(title=title)
    root.destroy()
    return path or ''


def _tk_file(title):
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    path = filedialog.askopenfilename(title=title)
    root.destroy()
    return path or ''


def pick_directory(title='Select Folder'):
    if platform.system() == 'Darwin':
        return _mac_pick(f'POSIX path of (choose folder with prompt "{title}")')
    result = _zenity('--file-selection', '--directory', f'--title={title}')
    return result if result is not None else _tk_directory(title)


def pick_file(title='Select File'):
    if platform.system() == 'Darwin':
        return _mac_pick(f'POSIX path of (choose file with prompt "{title}")')
    result = _zenity('--file-selection', f'--title={title}')
    return result if result is not None else _tk_file(title)


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/browse/directory')
def api_browse_directory():
    title = request.args.get('title', 'Select Folder')
    try:
        return jsonify({'path': pick_directory(title)})
    except Exception as exc:
        app.logger.error('browse/directory error: %s', exc, exc_info=True)
        return jsonify({'path': '', 'error': str(exc)}), 500


@app.route('/browse/file')
def api_browse_file():
    title = request.args.get('title', 'Select File')
    try:
        return jsonify({'path': pick_file(title)})
    except Exception as exc:
        app.logger.error('browse/file error: %s', exc, exc_info=True)
        return jsonify({'path': '', 'error': str(exc)}), 500


@app.route('/run')
def run_script():
    photos = request.args.get('photos', '')
    chat   = request.args.get('chat',   '')
    outdir = request.args.get('outdir', '')
    output = request.args.get('output', '') or 'duplicates.xlsx'
    start  = request.args.get('start',  '')
    end    = request.args.get('end',    '')

    spreadsheet = os.path.join(outdir, output) if outdir else output

    script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               'DuplicateFinder.py')
    cmd = [sys.executable, script_path,
           '-p', photos, '-c', chat, '-o', spreadsheet]
    if start and end:
        cmd += ['-s', start, '-e', end]

    def stream():
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
            )
            for raw in proc.stdout:
                line = ANSI_ESCAPE.sub('', raw).rstrip()
                if line:
                    yield f'data: {json.dumps({"line": line})}\n\n'
            proc.wait()
            yield f'data: {json.dumps({"done": True, "exit_code": proc.returncode})}\n\n'
        except Exception as exc:
            payload = {"line": f"[X] Failed to start: {exc}", "done": True, "exit_code": 1}
            yield f'data: {json.dumps(payload)}\n\n'

    return Response(
        stream(),
        mimetype='text/event-stream',
        headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'},
    )


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == '__main__':
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.bind(('127.0.0.1', 0))
    port = sock.getsockname()[1]
    sock.close()

    url = f'http://127.0.0.1:{port}'
    print(f'ImageDuplicator GUI → {url}  (Ctrl+C to quit)')
    threading.Thread(target=lambda: (
        __import__('time').sleep(0.8), webbrowser.open(url)
    ), daemon=True).start()

    app.run(host='127.0.0.1', port=port, debug=False, threaded=True)

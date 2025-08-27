import os
import sys
import subprocess
from pathlib import Path
from flask import Flask, Response, send_file, request, render_template_string

DEFAULT_INICIO = os.environ.get("VR_INICIO", "2025-04-15")
DEFAULT_FIM = os.environ.get("VR_FIM", "2025-05-15")
DEFAULT_COMPETENCIA = os.environ.get("VR_COMPETENCIA", "2025-05")
ENV_PROJECT_ROOT = os.environ.get("VR_PROJECT_ROOT", "")

HERE = Path(__file__).resolve().parent
PROJECT_ROOT = Path(ENV_PROJECT_ROOT).resolve() if ENV_PROJECT_ROOT else HERE

def find_project_root():
    for base in [PROJECT_ROOT, PROJECT_ROOT.parent, PROJECT_ROOT.parent.parent, PROJECT_ROOT.parent.parent.parent]:
        if (base / "src" / "main.py").exists() and (base / "data").exists():
            return base
    return PROJECT_ROOT

PROJECT_ROOT = find_project_root()
MAIN_PY = PROJECT_ROOT / "src" / "main.py"
DATA_DIR = PROJECT_ROOT / "data"
OUT_DIR = PROJECT_ROOT / "out"

app = Flask(__name__)

INDEX_HTML = """
<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>VR Mensal - Gerador</title>
  <style>
    :root { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; }
    body { display: grid; place-items: center; min-height: 100dvh; margin: 0; background: #0b132b; color: #eee; }
    .card { background: #161b33; padding: 28px; border-radius: 16px; box-shadow: 0 10px 30px rgba(0,0,0,.35); width: min(460px, 92vw); }
    h1 { margin: 0 0 12px; font-size: 22px; letter-spacing: .3px; }
    p.sub { margin: 0 18px 18px 0; color: #b9c2ff; font-size: 13px; opacity: .9; }
    button { appearance: none; border: 0; padding: 14px 18px; border-radius: 12px; background: #5bc0be; color: #0b132b; font-weight: 700; cursor: pointer; width: 100%; }
    .hint { font-size: 12px; color: #b3b3b3; margin-top: 12px; }
    code { background: #0f1533; padding: 2px 6px; border-radius: 6px; }
    form { margin-top: 10px; }
    .kv { font-size: 12px; color: #9fb0ff; margin-top: 8px; }
  </style>
</head>
<body>
  <div class="card">
    <h1>VR Mensal</h1>
    <p class="sub">Coloque as planilhas em <code>data/</code> e clique no botão para gerar e baixar o CSV.</p>
    <form action="/generate" method="get">
      <input type="hidden" name="inicio" value="{{inicio}}">
      <input type="hidden" name="fim" value="{{fim}}">
      <input type="hidden" name="competencia" value="{{competencia}}">
      <button type="submit">Gerar & baixar CSV</button>
    </form>
    <div class="kv">Projeto: <code>{{project}}</code></div>
    <div class="kv">main.py: <code>{{main_status}}</code> — data/: <code>{{data_status}}</code> — out/: <code>{{out_status}}</code></div>
    <p class="hint">Para mudar a competência: <code>/generate?inicio=2025-05-15&fim=2025-06-15&competencia=2025-06</code></p>
  </div>
</body>
</html>
"""

@app.get("/")
def index():
    return render_template_string(
        INDEX_HTML,
        inicio=DEFAULT_INICIO,
        fim=DEFAULT_FIM,
        competencia=DEFAULT_COMPETENCIA,
        project=str(PROJECT_ROOT),
        main_status="OK" if MAIN_PY.exists() else "NÃO ENCONTRADO",
        data_status="OK" if DATA_DIR.exists() else "NÃO ENCONTRADO",
        out_status="OK" if OUT_DIR.exists() else "NÃO ENCONTRADO",
    )

@app.get("/generate")
def generate():
    inicio = request.args.get("inicio", DEFAULT_INICIO)
    fim = request.args.get("fim", DEFAULT_FIM)
    competencia = request.args.get("competencia", DEFAULT_COMPETENCIA)

    if not MAIN_PY.exists():
        return Response(f"main.py não encontrado em {MAIN_PY}\nColoque este app.py ao lado de src/, data/, out/ (ou use VR_PROJECT_ROOT).", status=500, mimetype="text/plain; charset=utf-8")
    if not DATA_DIR.exists():
        return Response(f"Pasta data/ não encontrada em {DATA_DIR}", status=500, mimetype="text/plain; charset=utf-8")

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    comp_num = competencia.replace("-", "")
    out_csv = OUT_DIR / f"VR_MENSAL_{comp_num}.csv"

    cmd = [
        sys.executable, str(MAIN_PY),
        "--inicio", inicio, "--fim", fim, "--competencia", competencia,
        "--data_dir", str(DATA_DIR), "--out_dir", str(OUT_DIR),
    ]
    run = subprocess.run(cmd, capture_output=True, text=True)
    if run.returncode != 0:
        err = (run.stderr or "") + "\n" + (run.stdout or "")
        return Response("Falha ao executar main.py\n\n" + err, status=500, mimetype="text/plain; charset=utf-8")

    if not out_csv.exists():
        return Response("CSV não foi gerado. Verifique as planilhas .xlsx em data/ e tente novamente.", status=500, mimetype="text/plain; charset=utf-8")

    return send_file(out_csv, as_attachment=True, download_name=out_csv.name, mimetype="text/csv", max_age=0)

if __name__ == "__main__":
    app.run(debug=True, port=5000)

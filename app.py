import os, io, zipfile, tempfile, subprocess
from pathlib import Path
from flask import Flask, request, send_file, render_template, jsonify
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_CONTENT_LENGTH", 256 * 1024 * 1024))

HERE = Path(__file__).resolve().parent
MAIN_PY = HERE / "src" / "main.py"
BUNDLED_DATA = HERE / "data"
DEFAULT_OUT = HERE / "out"

DEFAULT_INICIO = os.environ.get("VR_INICIO", "2025-04-15")
DEFAULT_FIM = os.environ.get("VR_FIM", "2025-05-15")
DEFAULT_COMPETENCIA = os.environ.get("VR_COMPETENCIA", "2025-05")

@app.get("/")
def index():
    return render_template("index.html",
                           default_inicio=DEFAULT_INICIO,
                           default_fim=DEFAULT_FIM,
                           default_competencia=DEFAULT_COMPETENCIA)

def _run_generator(data_dir: Path, out_dir: Path, inicio: str, fim: str, competencia: str):
    comp_num = competencia.replace("-", "")
    cmd = [
        os.environ.get("PYTHON_BIN", "python"),
        str(MAIN_PY),
        "--inicio", inicio,
        "--fim", fim,
        "--competencia", competencia,
        "--data_dir", str(data_dir),
        "--out_dir", str(out_dir),
    ]
    run = subprocess.run(cmd, capture_output=True, text=True)
    memzip = io.BytesIO()
    with zipfile.ZipFile(memzip, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("exec_cmd.txt", " ".join(cmd))
        z.writestr("stdout.txt", run.stdout or "")
        z.writestr("stderr.txt", run.stderr or "")
        csv_path = out_dir / f"VR_MENSAL_{comp_num}.csv"
        val_path = out_dir / f"VALIDACOES_{comp_num}.xlsx"
        if csv_path.exists():
            z.write(csv_path, csv_path.name)
        if val_path.exists():
            z.write(val_path, val_path.name)
        # Include a listing of data files used
        listing = "\n".join(sorted(p.name for p in data_dir.glob("*.xlsx")))
        z.writestr("data_files_used.txt", listing)
    memzip.seek(0)
    return run.returncode, memzip, comp_num

@app.post("/api/generate")
def generate():
    inicio = request.form.get("inicio", DEFAULT_INICIO)
    fim = request.form.get("fim", DEFAULT_FIM)
    competencia = request.form.get("competencia", DEFAULT_COMPETENCIA)
    mode = request.form.get("mode", "bundled")  # bundled | upload

    if not MAIN_PY.exists():
        return jsonify({"error": "main.py não encontrado."}), 500

    if mode == "bundled":
        # Use the bundled data folder directly
        DEFAULT_OUT.mkdir(parents=True, exist_ok=True)
        code, memzip, comp_num = _run_generator(BUNDLED_DATA, DEFAULT_OUT, inicio, fim, competencia)
        filename = ("ERRO_" if code != 0 else "") + f"resultados_{comp_num}.zip"
        return send_file(memzip, as_attachment=True, download_name=filename, mimetype="application/zip", max_age=0)
    else:
        # Upload mode
        files = request.files.getlist("files")
        if not files:
            return jsonify({"error": "Envie um ZIP com as planilhas ou selecione os .xlsx"}), 400
        with tempfile.TemporaryDirectory(prefix="vrm_") as tmpd:
            tmp = Path(tmpd)
            data_dir = tmp / "data"; out_dir = tmp / "out"
            data_dir.mkdir(parents=True, exist_ok=True)
            out_dir.mkdir(parents=True, exist_ok=True)
            # ZIP or multiple XLSX
            if len(files) == 1 and files[0].filename.lower().endswith(".zip"):
                zf = zipfile.ZipFile(files[0].stream)
                for info in zf.infolist():
                    if info.is_dir(): continue
                    if not info.filename.lower().endswith(".xlsx"): continue
                    safe = secure_filename(Path(info.filename).name)
                    with zf.open(info) as zsrc, open(data_dir / safe, "wb") as dst:
                        dst.write(zsrc.read())
            else:
                for f in files:
                    if not f.filename.lower().endswith(".xlsx"): continue
                    safe = secure_filename(Path(f.filename).name)
                    f.save(data_dir / safe)
            if not any(p.suffix.lower()==".xlsx" for p in data_dir.glob("*")):
                return jsonify({"error": "Nenhum .xlsx encontrado no upload."}), 400
            code, memzip, comp_num = _run_generator(data_dir, out_dir, inicio, fim, competencia)
            filename = ("ERRO_" if code != 0 else "") + f"resultados_{comp_num}.zip"
            return send_file(memzip, as_attachment=True, download_name=filename, mimetype="application/zip", max_age=0)

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True, use_reloader=False)

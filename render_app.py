from __future__ import annotations

import json
import os
import secrets
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any

from flask import Flask, jsonify, redirect, render_template_string, request, send_file, url_for
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from werkzeug.utils import secure_filename

from telkom_batch_check import (
    CrdbClient,
    LookupResult,
    load_cache,
    normalize_number,
    parse_result,
    process_workbook,
    save_cache,
)


DATA_DIR = Path(os.environ.get("DATA_DIR", "/tmp/telkom-render-checker"))
UPLOAD_DIR = DATA_DIR / "uploads"
REPORT_DIR = DATA_DIR / "reports"
CACHE_PATH = DATA_DIR / "telkom_lookup_cache.json"
STATE_PATH = DATA_DIR / "state.json"
APP_PASSWORD = os.environ.get("APP_PASSWORD", "").strip()

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(16))


class CloudState:
    def __init__(self) -> None:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        REPORT_DIR.mkdir(parents=True, exist_ok=True)
        self.client = CrdbClient(timeout=45)
        self.cache = load_cache(CACHE_PATH)
        self.pending_captcha: dict[str, dict[str, str]] = {}
        self.files: list[dict[str, str]] = []
        self.rows: list[dict[str, Any]] = []
        self.load_state()

    def load_state(self) -> None:
        if STATE_PATH.exists():
            raw = json.loads(STATE_PATH.read_text(encoding="utf-8"))
            self.files = raw.get("files", [])
        self.rows = self.load_rows()

    def save_state(self) -> None:
        STATE_PATH.write_text(json.dumps({"files": self.files}, indent=2), encoding="utf-8")

    def reset(self) -> None:
        self.files = []
        self.rows = []
        self.pending_captcha = {}
        for path in UPLOAD_DIR.glob("*.xlsx"):
            path.unlink(missing_ok=True)
        for path in REPORT_DIR.glob("*.xlsx"):
            path.unlink(missing_ok=True)
        STATE_PATH.unlink(missing_ok=True)

    def add_upload(self, file_storage) -> None:
        filename = secure_filename(file_storage.filename or "")
        if not filename.lower().endswith(".xlsx"):
            raise ValueError("Please upload .xlsx files only.")
        target = UPLOAD_DIR / filename
        file_storage.save(target)
        label = target.stem.replace("_", " ").replace("-", " ").title()
        self.files = [item for item in self.files if item["path"] != str(target)]
        self.files.append({"label": label, "path": str(target)})
        self.save_state()
        self.rows = self.load_rows()

    def load_rows(self) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        for item in self.files:
            path = Path(item["path"])
            if not path.exists():
                continue
            wb = load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
            number_col = self.find_number_column(headers)
            company_col = 1
            for row_num in range(2, ws.max_row + 1):
                company = ws.cell(row_num, company_col).value
                original = ws.cell(row_num, number_col).value
                clean = normalize_number(original)
                rows.append(
                    {
                        "area": item["label"],
                        "source": str(path),
                        "row": row_num,
                        "company": company or "",
                        "original_number": original or "",
                        "clean_number": clean,
                    }
                )
            wb.close()
        return rows

    @staticmethod
    def find_number_column(headers: list[Any]) -> int:
        for idx, header in enumerate(headers, start=1):
            text = str(header or "").strip().lower()
            if text == "contact number" or ("contact" in text and "number" in text):
                return idx
            if text in {"phone", "phone number", "telephone", "number", "msisdn"}:
                return idx
        return 2

    def summary(self) -> dict[str, Any]:
        checked_numbers = {number for number, result in self.cache.items() if result.lookup_status == "Found"}
        checked_rows = [row for row in self.rows if row["clean_number"] in checked_numbers]
        telkom_rows = [row for row in checked_rows if self.cache[row["clean_number"]].telkom == "Yes"]
        non_telkom_rows = [row for row in checked_rows if self.cache[row["clean_number"]].telkom == "No"]
        return {
            "files": len(self.files),
            "total_rows": len(self.rows),
            "checked_rows": len(checked_rows),
            "remaining_rows": len(self.rows) - len(checked_rows),
            "unique_checked_numbers": len(checked_numbers),
            "telkom_rows": len(telkom_rows),
            "non_telkom_rows": len(non_telkom_rows),
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

    def next_row(self) -> dict[str, Any] | None:
        for row in self.rows:
            clean = row["clean_number"]
            if clean and clean not in self.cache:
                return row
        return None

    def checked_rows(self, limit: int = 30) -> list[dict[str, Any]]:
        items: list[dict[str, Any]] = []
        for row in self.rows:
            result = self.cache.get(row["clean_number"])
            if not result:
                continue
            items.append(
                {
                    **row,
                    "status": result.lookup_status,
                    "provider": result.current_provider,
                    "telkom": result.telkom,
                    "raw": result.raw_result,
                }
            )
        return items[-limit:]

    def save_result(self, result: LookupResult) -> None:
        if result.lookup_status == "Found":
            self.cache[result.clean_number] = result
            save_cache(CACHE_PATH, self.cache)

    def check_number(self, clean_number: str) -> dict[str, Any]:
        if clean_number in self.cache:
            return {"kind": "found", "result": asdict(self.cache[clean_number])}
        response = self.client.submit_number(clean_number)
        result = parse_result(clean_number, response)
        if result.lookup_status == "Captcha required":
            return self.new_captcha(clean_number)
        self.save_result(result)
        return {"kind": "found", "result": asdict(result)}

    def new_captcha(self, clean_number: str) -> dict[str, Any]:
        response = self.client.request("GET", "captcha/captcha-gen")
        image_data = str(response.get("imageData") or "")
        string_data = str(response.get("stringData") or "")
        self.pending_captcha[clean_number] = {"captchaEncrypt": string_data}
        return {
            "kind": "captcha",
            "number": clean_number,
            "imageData": image_data,
            "message": "Please type the captcha shown in the image.",
        }

    def submit_captcha(self, clean_number: str, code: str) -> dict[str, Any]:
        pending = self.pending_captcha.get(clean_number)
        if not pending:
            return self.new_captcha(clean_number)
        payload = {
            "number": clean_number,
            "captcha": code.strip(),
            "captchaEncrypt": pending["captchaEncrypt"],
            "puid": self.client.puid,
        }
        response = self.client.request("POST", "publicInquiry/submitRequest", payload)
        result = parse_result(clean_number, response)
        if result.lookup_status == "Captcha required":
            return self.new_captcha(clean_number)
        self.pending_captcha.pop(clean_number, None)
        self.save_result(result)
        return {"kind": "found", "result": asdict(result)}

    def export_reports(self) -> dict[str, Any]:
        outputs = []
        for item in self.files:
            source = Path(item["path"])
            output = REPORT_DIR / f"{source.stem}_manual_checked.xlsx"
            process_workbook(source, output, self.cache, CACHE_PATH, delay_seconds=0, cache_only=True)
            outputs.append({"label": output.name, "url": url_for("download_report", filename=output.name)})
        telkom = self.export_telkom_only()
        outputs.append({"label": telkom.name, "url": url_for("download_report", filename=telkom.name)})
        return {"outputs": outputs}

    def export_telkom_only(self) -> Path:
        output = REPORT_DIR / "telkom_companies_only.xlsx"
        out_wb = Workbook()
        out_ws = out_wb.active
        out_ws.title = "Telkom Companies"
        out_row = 1
        wrote_header = False

        for item in self.files:
            source = Path(item["path"])
            wb = load_workbook(source, data_only=True)
            ws = wb.active
            headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
            number_col = self.find_number_column(headers)
            if not wrote_header:
                for col, value in enumerate(["Area"] + headers + ["Clean Number", "Current Provider", "Raw Result"], start=1):
                    cell = out_ws.cell(out_row, col, value)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="1F4E78")
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                wrote_header = True
                out_row += 1
            for row_num in range(2, ws.max_row + 1):
                clean = normalize_number(ws.cell(row_num, number_col).value)
                result = self.cache.get(clean)
                if not result or result.telkom != "Yes":
                    continue
                out_ws.cell(out_row, 1, item["label"])
                for col in range(1, ws.max_column + 1):
                    out_ws.cell(out_row, col + 1, ws.cell(row_num, col).value)
                start = ws.max_column + 2
                out_ws.cell(out_row, start, result.clean_number)
                out_ws.cell(out_row, start + 1, result.current_provider)
                out_ws.cell(out_row, start + 2, result.raw_result)
                out_row += 1
            wb.close()

        for row in range(2, out_ws.max_row + 1):
            for col in range(1, out_ws.max_column + 1):
                out_ws.cell(row, col).fill = PatternFill("solid", fgColor="C6EFCE")
                out_ws.cell(row, col).alignment = Alignment(vertical="top", wrap_text=True)
        out_ws.freeze_panes = "A2"
        out_ws.auto_filter.ref = out_ws.dimensions
        for col in range(1, out_ws.max_column + 1):
            max_len = max(len(str(out_ws.cell(row, col).value or "")) for row in range(1, out_ws.max_row + 1))
            out_ws.column_dimensions[get_column_letter(col)].width = min(max(max_len + 2, 10), 52)
        out_wb.save(output)
        return output


STATE = CloudState()


def authorized() -> bool:
    if not APP_PASSWORD:
        return True
    supplied = request.args.get("key") or request.headers.get("X-App-Password") or ""
    return secrets.compare_digest(supplied, APP_PASSWORD)


@app.before_request
def require_password():
    if request.path.startswith("/static"):
        return None
    if not authorized():
        return jsonify({"error": "Unauthorized. Add ?key=your-password to the URL."}), 401
    return None


PAGE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Telkom Cloud Checker</title>
  <style>
    body { margin: 0; font-family: Arial, sans-serif; background: #f6f7f9; color: #17202a; }
    main { max-width: 1120px; margin: 0 auto; padding: 28px; }
    header { display: flex; justify-content: space-between; gap: 16px; align-items: flex-start; }
    h1 { margin: 0; font-size: 28px; }
    .grid { display: grid; grid-template-columns: repeat(6, 1fr); gap: 10px; margin: 18px 0; }
    .metric, .panel { background: #fff; border: 1px solid #d9dee7; border-radius: 8px; padding: 14px; }
    .metric span { display: block; font-size: 12px; color: #657080; }
    .metric strong { display: block; font-size: 24px; margin-top: 4px; }
    .panel { margin-top: 14px; }
    .row-title { display: grid; grid-template-columns: 1.2fr .8fr .8fr; gap: 12px; }
    .label { font-size: 12px; color: #657080; margin-bottom: 4px; }
    .value { font-size: 18px; font-weight: 700; word-break: break-word; }
    button { border: 0; border-radius: 6px; background: #174ea6; color: white; padding: 11px 14px; font-weight: 700; cursor: pointer; }
    button.secondary { background: #4b5563; }
    button.danger { background: #9f1239; }
    button:disabled { background: #9aa4b2; cursor: wait; }
    input { padding: 10px; border: 1px solid #c9d1dc; border-radius: 6px; font-size: 16px; }
    .actions { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 14px; align-items: center; }
    .result { padding: 12px; border-radius: 6px; margin-top: 14px; background: #eef3ff; white-space: pre-wrap; }
    .yes { background: #d9ead3; } .no { background: #f4cccc; } .unknown { background: #ffe699; }
    .captcha { display: none; margin-top: 14px; padding: 14px; border: 1px solid #d7b945; background: #fff7d6; border-radius: 8px; }
    .captcha img { display: block; max-width: 260px; border: 1px solid #d0d0d0; background: white; margin-bottom: 10px; }
    table { width: 100%; border-collapse: collapse; font-size: 14px; }
    th, td { text-align: left; border-bottom: 1px solid #e5e7eb; padding: 8px; vertical-align: top; }
    th { background: #eef2f7; }
    a.download { display: inline-block; margin: 6px 8px 0 0; color: #174ea6; font-weight: 700; }
    @media (max-width: 850px) { .grid { grid-template-columns: 1fr 1fr; } .row-title, header { display: block; } }
  </style>
</head>
<body>
<main>
  <header>
    <div>
      <h1>Telkom Cloud Checker</h1>
      <p>Upload Excel files, check each number, type captcha when requested, then export reports.</p>
    </div>
    <button class="secondary" id="exportBtn">Export Reports</button>
  </header>

  <section class="panel">
    <form action="/upload{% if key %}?key={{ key }}{% endif %}" method="post" enctype="multipart/form-data">
      <div class="label">Upload .xlsx files</div>
      <input type="file" name="files" multiple accept=".xlsx">
      <button type="submit">Upload</button>
      <button class="danger" type="button" id="resetBtn">Reset Uploaded Files</button>
    </form>
  </section>

  <section class="grid" id="summary"></section>

  <section class="panel">
    <div class="row-title">
      <div><div class="label">Company</div><div class="value" id="company">Loading...</div></div>
      <div><div class="label">Number</div><div class="value" id="number"></div></div>
      <div><div class="label">File</div><div class="value" id="area"></div></div>
    </div>
    <div class="actions">
      <button id="checkBtn">Check This Number</button>
      <button class="secondary" id="skipBtn">Skip For Now</button>
    </div>
    <div class="captcha" id="captchaBox">
      <div class="label">Captcha</div>
      <img id="captchaImg" alt="Captcha">
      <div class="actions">
        <input id="captchaInput" autocomplete="off" placeholder="Type captcha code">
        <button id="captchaBtn">Submit Captcha</button>
      </div>
    </div>
    <div id="result" class="result"></div>
    <div id="downloads"></div>
  </section>

  <section class="panel">
    <h2>Recent Checked Rows</h2>
    <table>
      <thead><tr><th>File</th><th>Company</th><th>Number</th><th>Provider</th><th>Telkom?</th></tr></thead>
      <tbody id="recent"></tbody>
    </table>
  </section>
</main>
<script>
const key = new URLSearchParams(location.search).get('key') || '';
const suffix = key ? '?key=' + encodeURIComponent(key) : '';
let current = null;
let busy = false;
async function api(path, body) {
  const response = await fetch(path + suffix, body ? { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) } : {});
  return response.json();
}
function setBusy(value) { busy = value; document.querySelectorAll('button').forEach(btn => btn.disabled = value); }
function metric(label, value) { return `<div class="metric"><span>${label}</span><strong>${value}</strong></div>`; }
function renderSummary(summary) {
  document.getElementById('summary').innerHTML = [
    metric('Files', summary.files), metric('Total rows', summary.total_rows), metric('Checked', summary.checked_rows),
    metric('Remaining', summary.remaining_rows), metric('Telkom', summary.telkom_rows), metric('Non-Telkom', summary.non_telkom_rows)
  ].join('');
}
function renderCurrent(row) {
  current = row; document.getElementById('captchaBox').style.display = 'none'; document.getElementById('captchaInput').value = '';
  if (!row) {
    document.getElementById('company').textContent = 'No unchecked row available';
    document.getElementById('number').textContent = ''; document.getElementById('area').textContent = '';
    document.getElementById('result').textContent = 'Upload files or export reports if checking is complete.'; return;
  }
  document.getElementById('company').textContent = row.company;
  document.getElementById('number').textContent = row.clean_number;
  document.getElementById('area').textContent = row.area;
}
function renderRecent(rows) {
  document.getElementById('recent').innerHTML = rows.slice().reverse().map(row => `
    <tr><td>${row.area}</td><td>${row.company}</td><td>${row.clean_number}</td><td>${row.provider}</td><td>${row.telkom}</td></tr>`).join('');
}
function showResult(result) {
  const box = document.getElementById('result');
  box.className = 'result ' + (result.telkom === 'Yes' ? 'yes' : result.telkom === 'No' ? 'no' : 'unknown');
  box.textContent = `${result.lookup_status}\\nProvider: ${result.current_provider || 'Unknown'}\\nTelkom: ${result.telkom}\\n${result.raw_result}`;
}
async function refresh() {
  const data = await api('/api/state');
  renderSummary(data.summary); renderCurrent(data.next); renderRecent(data.recent);
}
async function checkCurrent() {
  if (!current || busy) return; setBusy(true);
  try {
    const data = await api('/api/check', { number: current.clean_number });
    if (data.kind === 'captcha') {
      document.getElementById('captchaBox').style.display = 'block';
      document.getElementById('captchaImg').src = 'data:image/jpeg;base64,' + data.imageData;
      document.getElementById('result').textContent = data.message;
      document.getElementById('captchaInput').focus();
    } else { showResult(data.result); await refresh(); }
  } finally { setBusy(false); }
}
async function submitCaptcha() {
  if (!current || busy) return;
  const code = document.getElementById('captchaInput').value.trim();
  if (!code) return; setBusy(true);
  try {
    const data = await api('/api/captcha', { number: current.clean_number, code });
    if (data.kind === 'captcha') {
      document.getElementById('captchaImg').src = 'data:image/jpeg;base64,' + data.imageData;
      document.getElementById('captchaInput').value = '';
      document.getElementById('result').textContent = 'That code was not accepted. Try the new captcha.';
      document.getElementById('captchaInput').focus();
    } else { document.getElementById('captchaBox').style.display = 'none'; showResult(data.result); await refresh(); }
  } finally { setBusy(false); }
}
document.getElementById('checkBtn').addEventListener('click', checkCurrent);
document.getElementById('captchaBtn').addEventListener('click', submitCaptcha);
document.getElementById('captchaInput').addEventListener('keydown', e => { if (e.key === 'Enter') submitCaptcha(); });
document.getElementById('skipBtn').addEventListener('click', refresh);
document.getElementById('resetBtn').addEventListener('click', async () => { if (confirm('Remove uploaded files from this cloud app?')) { await api('/api/reset', {}); await refresh(); } });
document.getElementById('exportBtn').addEventListener('click', async () => {
  setBusy(true);
  try {
    const data = await api('/api/export', {});
    document.getElementById('downloads').innerHTML = data.outputs.map(item => `<a class="download" href="${item.url}${key ? '&key=' + encodeURIComponent(key) : ''}">${item.label}</a>`).join('');
    await refresh();
  } finally { setBusy(false); }
});
refresh();
</script>
</body>
</html>
"""


@app.get("/")
def index():
    return render_template_string(PAGE, key=request.args.get("key", ""))


@app.post("/upload")
def upload():
    files = request.files.getlist("files")
    for uploaded in files:
        if uploaded and uploaded.filename:
            STATE.add_upload(uploaded)
    return redirect("/" + (f"?key={request.args.get('key')}" if request.args.get("key") else ""))


@app.get("/api/state")
def api_state():
    return jsonify({"summary": STATE.summary(), "next": STATE.next_row(), "recent": STATE.checked_rows()})


@app.post("/api/check")
def api_check():
    data = request.get_json(force=True)
    return jsonify(STATE.check_number(str(data.get("number") or "")))


@app.post("/api/captcha")
def api_captcha():
    data = request.get_json(force=True)
    return jsonify(STATE.submit_captcha(str(data.get("number") or ""), str(data.get("code") or "")))


@app.post("/api/export")
def api_export():
    return jsonify(STATE.export_reports())


@app.post("/api/reset")
def api_reset():
    STATE.reset()
    return jsonify({"ok": True})


@app.get("/download/<path:filename>")
def download_report(filename: str):
    path = REPORT_DIR / secure_filename(filename)
    if not path.exists():
        return jsonify({"error": "Report not found. Export reports first."}), 404
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8765"))
    app.run(host="0.0.0.0", port=port)

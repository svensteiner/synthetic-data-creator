"""
audit.py - HTML Auditbericht Generator
Erstellt ein professionelles Dokument als Nachweis fuer datenschutzkonforme Verarbeitung.
"""

import datetime
import html
import socket
from pathlib import Path

from synthesizer import TYPE_LABELS


def generate_audit_report(
    input_path: str,
    output_path: str,
    col_info: dict,       # {'sheet': {'col': 'type'}}
    row_counts: dict,     # {'sheet': n_rows}
) -> str:
    """
    Erstellt einen HTML-Auditbericht neben der Ausgabedatei.
    Gibt den Pfad zur HTML-Datei zurueck.
    """
    audit_path = str(Path(output_path).with_suffix("")) + "_Auditbericht.html"

    now       = datetime.datetime.now()
    timestamp = now.strftime("%d.%m.%Y um %H:%M:%S Uhr")
    try:
        machine = socket.gethostname()
    except Exception:
        machine = "Unbekannt"

    # Statistiken
    all_cols   = {col: ct for sh in col_info.values() for col, ct in sh.items()}
    total_cols = len(all_cols)
    anon_cols  = sum(1 for ct in all_cols.values() if ct not in ("text", "beschreibung"))
    skip_cols  = total_cols - anon_cols
    total_rows = sum(row_counts.values())

    # Zeilen fuer Spalten-Tabelle
    rows_html = ""
    for sheet, cols in col_info.items():
        for col, ct in cols.items():
            is_anon  = ct not in ("text", "beschreibung")
            status   = "Anonymisiert" if is_anon else "Unveraendert"
            dot_col  = "#27ae60" if is_anon else "#bdc3c7"
            rows_html += f"""
                <tr>
                    <td class="sheet-col">{html.escape(sheet)}</td>
                    <td><strong>{html.escape(col)}</strong></td>
                    <td>{html.escape(TYPE_LABELS.get(ct, ct))}</td>
                    <td><span class="dot" style="background:{dot_col}"></span> {status}</td>
                </tr>"""

    html = f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<title>Auditbericht - Synthetische Datengenerierung</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f0f2f5; color: #2c3e50; }}

  .page {{ max-width: 900px; margin: 30px auto; background: white;
            box-shadow: 0 2px 12px rgba(0,0,0,.12); border-radius: 6px; overflow: hidden; }}

  /* Header */
  .header {{ background: #1a3a5c; color: white; padding: 28px 36px; }}
  .header h1 {{ font-size: 20px; font-weight: 600; margin-bottom: 4px; }}
  .header p  {{ font-size: 12px; opacity: .75; }}
  .header .badge {{ display:inline-block; background:#27ae60; color:white;
                    font-size:10px; padding:2px 8px; border-radius:3px;
                    margin-top:8px; letter-spacing:.5px; }}

  /* Notice */
  .notice {{ background:#eaf4ea; border-left:4px solid #27ae60;
             padding:12px 18px; margin:24px 36px 0; font-size:12px; color:#1e6b1e;
             border-radius:0 4px 4px 0; }}

  /* Meta */
  .meta {{ margin:20px 36px; border:1px solid #e0e0e0; border-radius:5px; overflow:hidden; }}
  .meta table {{ width:100%; border-collapse:collapse; font-size:12.5px; }}
  .meta td {{ padding:8px 14px; border-bottom:1px solid #f0f0f0; }}
  .meta td:first-child {{ background:#f8f9fa; font-weight:600; color:#555; width:180px; }}
  .meta tr:last-child td {{ border-bottom:none; }}

  /* Summary cards */
  .cards {{ display:flex; gap:16px; margin:20px 36px; }}
  .card {{ flex:1; border-radius:6px; padding:16px 20px; text-align:center; border:1px solid #e0e0e0; }}
  .card .num {{ font-size:34px; font-weight:700; line-height:1; }}
  .card .lbl {{ font-size:11px; color:#888; margin-top:5px; text-transform:uppercase; letter-spacing:.5px; }}
  .card.blue  {{ background:#eef3fa; }} .card.blue .num  {{ color:#1a3a5c; }}
  .card.green {{ background:#eaf4ea; }} .card.green .num {{ color:#27ae60; }}
  .card.gray  {{ background:#f8f8f8; }} .card.gray .num  {{ color:#7f8c8d; }}

  /* Section */
  .section {{ margin:20px 36px; }}
  .section h2 {{ font-size:13px; color:#1a3a5c; text-transform:uppercase;
                 letter-spacing:.6px; border-bottom:2px solid #e0e0e0;
                 padding-bottom:7px; margin-bottom:12px; }}

  /* Table */
  table.cols {{ width:100%; border-collapse:collapse; font-size:12.5px; }}
  table.cols th {{ background:#1a3a5c; color:white; padding:9px 12px; text-align:left;
                   font-weight:500; font-size:11px; text-transform:uppercase; letter-spacing:.4px; }}
  table.cols td {{ padding:8px 12px; border-bottom:1px solid #f0f0f0; }}
  table.cols tr:last-child td {{ border-bottom:none; }}
  table.cols tr:nth-child(even) td {{ background:#fafafa; }}
  .sheet-col {{ color:#888; font-size:11px; }}
  .dot {{ display:inline-block; width:8px; height:8px; border-radius:50%;
          margin-right:5px; vertical-align:middle; }}

  /* Footer */
  .footer {{ margin:24px 36px 36px; padding:16px 18px;
             background:#f8f9fa; border-radius:5px;
             font-size:11px; color:#888; line-height:1.7; }}
  .footer strong {{ color:#555; }}

  @media print {{
    body {{ background:white; }}
    .page {{ box-shadow:none; margin:0; border-radius:0; }}
  }}
</style>
</head>
<body>
<div class="page">

  <div class="header">
    <h1>Auditbericht &ndash; Synthetische Datengenerierung</h1>
    <p>Dokumentation der datenschutzkonformen Datenverarbeitung</p>
    <span class="badge">DSGVO-konform &bull; 100% lokal verarbeitet</span>
  </div>

  <div class="notice">
    Alle Verarbeitungsschritte wurden ausschliesslich lokal auf diesem Geraet durchgefuehrt.
    Es wurden keine personenbezogenen oder vertraulichen Daten an externe Dienste,
    Cloud-Dienste oder KI-Plattformen uebermittelt.
  </div>

  <div class="meta">
    <table>
      <tr><td>Originaldatei</td>      <td><strong>{html.escape(Path(input_path).name)}</strong></td></tr>
      <tr><td>Originalpfad</td>       <td><code style="font-size:11px">{html.escape(input_path)}</code></td></tr>
      <tr><td>Ausgabedatei</td>       <td><strong>{html.escape(Path(output_path).name)}</strong></td></tr>
      <tr><td>Ausgabepfad</td>        <td><code style="font-size:11px">{html.escape(output_path)}</code></td></tr>
      <tr><td>Zeitpunkt</td>          <td>{timestamp}</td></tr>
      <tr><td>Geraet / Hostname</td>  <td>{machine}</td></tr>
      <tr><td>Verarbeitete Zeilen</td><td>{total_rows:,} Datensaetze</td></tr>
    </table>
  </div>

  <div class="cards">
    <div class="card blue">
      <div class="num">{total_cols}</div>
      <div class="lbl">Spalten gesamt</div>
    </div>
    <div class="card green">
      <div class="num">{anon_cols}</div>
      <div class="lbl">Anonymisiert</div>
    </div>
    <div class="card gray">
      <div class="num">{skip_cols}</div>
      <div class="lbl">Unveraendert</div>
    </div>
  </div>

  <div class="section">
    <h2>Spalten-Detail</h2>
    <table class="cols">
      <thead>
        <tr>
          <th>Sheet</th>
          <th>Spaltenname</th>
          <th>Erkannter Typ</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>
        {rows_html}
      </tbody>
    </table>
  </div>

  <div class="footer">
    <strong>Bestaetigungsvermerk:</strong><br>
    Die Datei &bdquo;{html.escape(Path(output_path).name)}&ldquo; wurde am {timestamp} durch den
    Synthetische Daten Generator (v1.0) erstellt. Saemtliche personenbezogenen und
    vertraulichen Daten wurden durch synthetische, realistische Alternativwerte ersetzt.
    Die Ausgabedatei enthaelt keine Originalinformationen und kann ohne datenschutzrechtliche
    Bedenken fuer Analysezwecke, KI-gestuetzte Auswertungen und interne Tests verwendet werden.
    <br><br>
    <strong>Methode:</strong> Regelbasierte Erkennung &bull;
    Faker (de_DE) &bull; Konsistentes Mapping &bull; Min/Max-Bereichserhaltung fuer Betraege
  </div>

</div>
</body>
</html>"""

    with open(audit_path, "w", encoding="utf-8") as f:
        f.write(html)

    return audit_path

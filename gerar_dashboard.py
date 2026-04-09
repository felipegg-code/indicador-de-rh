"""
TechSolutions — Gerador de Dashboard RH
========================================
Lê os 3 arquivos xlsx e gera o dashboard HTML automaticamente.

Uso:
    python scripts/gerar_dashboard.py

Arquivos esperados em /dados:
    - BANCO_DE_HORAS_58_ANALISADO.xlsx  (aba: GERAL)
    - TRATAMENTO_PONTO_GERAL.xlsx       (aba: EXCECOES_JORNADA)
"""

import pandas as pd
import re
import json
import os
from pathlib import Path
from datetime import datetime

# ── CAMINHOS ──────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent.parent
DADOS_DIR  = BASE_DIR / "dados"
OUTPUT_DIR = BASE_DIR / "docs"
OUTPUT_DIR.mkdir(exist_ok=True)

BH_FILE  = DADOS_DIR / "BANCO_DE_HORAS_58_ANALISADO.xlsx"
EXC_FILE = DADOS_DIR / "TRATAMENTO_PONTO_GERAL.xlsx"

# ── HELPERS ───────────────────────────────────────────────────
def parse_bh_min(s):
    if not isinstance(s, str): return 0
    s = s.strip()
    neg = s.startswith('-')
    m = re.match(r'(\d+):(\d+)', s.lstrip('-'))
    if not m: return 0
    v = int(m.group(1)) * 60 + int(m.group(2))
    return -v if neg else v

def fmt_bh_saldo(s):
    if not s: return '+00:00'
    s = str(s).strip()
    return s if s.startswith('-') else '+' + s

def fmt_bh_h(minutos):
    neg = minutos < 0
    ab  = abs(minutos)
    h   = ab // 60
    m   = ab % 60
    return f"{'−' if neg else '+'}{h:02d}:{m:02d}"

def fmt_date(val):
    try:
        d = pd.to_datetime(val)
        return d.strftime('%d/%m')
    except:
        return str(val)[:5]

def get_week_label(val):
    try:
        d = pd.to_datetime(val)
        week = d.isocalendar().week
        return f"Sem {week} ({d.strftime('%d/%m')})"
    except:
        return '—'

def parse_intra_detail(s):
    s = str(s or '')
    mi = re.search(r'INTERVALO = (\d+:\d+)', s)
    mj = re.search(r'JORNADA = (\d+:\d+)', s)
    mr = re.search(r'MINIMO (\d+:\d+)', s)
    if not (mi and mj and mr):
        return {'intervalo': '—', 'jornada': '—', 'deficit': '—'}
    def to_m(t):
        h, m = t.split(':')
        return int(h) * 60 + int(m)
    diff = to_m(mr.group(1)) - to_m(mi.group(1))
    deficit = f"-{diff//60:02d}:{diff%60:02d}" if diff > 0 else '00:00'
    return {'intervalo': mi.group(1), 'jornada': mj.group(1), 'deficit': deficit}

def parse_inter_detail(s):
    s = str(s or '')
    m = re.search(r'INTERJORNADA = (\d+:\d+)', s)
    if not m:
        return {'inter': '—', 'deficit': '—'}
    def to_m(t):
        h, mm = t.split(':')
        return int(h) * 60 + int(mm)
    diff = 660 - to_m(m.group(1))
    deficit = f"-{diff//60:02d}:{diff%60:02d}" if diff > 0 else '00:00'
    return {'inter': m.group(1), 'deficit': deficit}

# ── LER DADOS ─────────────────────────────────────────────────
print("📂 Lendo arquivos...")

if not BH_FILE.exists():
    raise FileNotFoundError(f"Arquivo não encontrado: {BH_FILE}")
if not EXC_FILE.exists():
    raise FileNotFoundError(f"Arquivo não encontrado: {EXC_FILE}")

bh_df  = pd.read_excel(BH_FILE,  sheet_name='GERAL')
exc_df = pd.read_excel(EXC_FILE, sheet_name='EXCECOES_JORNADA')

# ── PROCESSAR BANCO DE HORAS ──────────────────────────────────
print("⚙️  Processando banco de horas...")

bh_df['bh_min'] = bh_df['TOTALGERAL'].apply(parse_bh_min)
bh_df['bh_h']   = (bh_df['bh_min'] / 60).round(2)

bh_colab = []
for _, r in bh_df.iterrows():
    bh_colab.append({
        'nome':  str(r.get('NOME',  '')),
        'secao': str(r.get('SECAO', '')),
        'funcao':str(r.get('FUNCAO','')),
        'saldo': fmt_bh_saldo(r.get('TOTALGERAL', '')),
        'bh_h':  float(r['bh_h']),
        'faixa': str(r.get('FAIXA_BANCO_HORAS', 'NORMAL')),
        'acao':  str(r.get('ACAO_BANCO_HORAS',  'SEM ACAO')),
    })

# Seções
sec_map = {}
for d in bh_colab:
    s = d['secao']
    if s not in sec_map:
        sec_map[s] = {'secao': s, 'n': 0, 'total': 0.0, 'maxPos': 0.0, 'maxNeg': 0.0}
    sec_map[s]['n']     += 1
    sec_map[s]['total'] += d['bh_h']
    if d['bh_h'] > sec_map[s]['maxPos']: sec_map[s]['maxPos'] = d['bh_h']
    if d['bh_h'] < sec_map[s]['maxNeg']: sec_map[s]['maxNeg'] = d['bh_h']

sec_arr = sorted(sec_map.values(), key=lambda x: -x['total'])

# KPIs banco de horas
kpi_bh_total   = len(bh_colab)
kpi_bh_normal  = sum(1 for d in bh_colab if d['faixa'] == 'NORMAL')
kpi_bh_atencao = sum(1 for d in bh_colab if d['faixa'] != 'NORMAL')
kpi_bh_critico = sum(1 for d in bh_colab if 'CRITICO' in d['faixa'])

# Faixas para donut
faixa_order = [
    {'k': 'CRITICO_POSITIVO_20', 'lbl': 'Crítico +20h',  'c': '#c0392b'},
    {'k': 'ALERTA_POSITIVO_15',  'lbl': 'Alerta +15h',   'c': '#f47c20'},
    {'k': 'ATENCAO_POSITIVO_10', 'lbl': 'Atenção +10h',  'c': '#c97b00'},
    {'k': 'NORMAL',              'lbl': 'Normal',         'c': '#1e9e5e'},
    {'k': 'ATENCAO_NEGATIVO_10', 'lbl': 'Atenção -10h',  'c': '#8b6fa0'},
    {'k': 'ALERTA_NEGATIVO_15',  'lbl': 'Alerta -15h',   'c': '#6a4fa0'},
]
faixa_counts = {}
for d in bh_colab:
    faixa_counts[d['faixa']] = faixa_counts.get(d['faixa'], 0) + 1
faixas = [f for f in faixa_order if faixa_counts.get(f['k'], 0) > 0]
for f in faixas:
    f['n'] = faixa_counts.get(f['k'], 0)

# ── PROCESSAR EXCEÇÕES DE JORNADA ─────────────────────────────
print("⚙️  Processando exceções de jornada...")

exc_df['DATA'] = pd.to_datetime(exc_df['DATA'], errors='coerce')
intra_rows = exc_df[exc_df['TIPO_OCORRENCIA'].str.contains('INTRAJORNADA', na=False)]
inter_rows = exc_df[exc_df['TIPO_OCORRENCIA'].str.contains('INTERJORNADA', na=False)]

def build_intra(rows):
    result = []
    for _, r in rows.iterrows():
        d = parse_intra_detail(r.get('DETALHE_OCORRENCIA', ''))
        result.append({
            'nome':   str(r.get('NOME',    '')),
            'secao':  str(r.get('SECAO',   '')),
            'data':   fmt_date(r['DATA']),
            'semana': get_week_label(r['DATA']),
            'b1': str(r.get('BATIDA1', '—')),
            'b2': str(r.get('BATIDA2', '—')),
            'b3': str(r.get('BATIDA3', '—')),
            'b4': str(r.get('BATIDA4', '—')),
            'htrab':    d['jornada'],
            'intra':    d['intervalo'],
            'deficit':  d['deficit'],
        })
    return result

def build_inter(rows):
    result = []
    for _, r in rows.iterrows():
        d = parse_inter_detail(r.get('DETALHE_OCORRENCIA', ''))
        result.append({
            'nome':   str(r.get('NOME',  '')),
            'secao':  str(r.get('SECAO', '')),
            'data':   fmt_date(r['DATA']),
            'semana': get_week_label(r['DATA']),
            'b1': str(r.get('BATIDA1', '—')),
            'b2': str(r.get('BATIDA2', '—')),
            'b3': str(r.get('BATIDA3', '—')),
            'b4': str(r.get('BATIDA4', '—')),
            'inter':   d['inter'],
            'deficit': d['deficit'],
        })
    return result

intra_data = build_intra(intra_rows)
inter_data = build_inter(inter_rows)

# Agregações intra
intra_sec_map  = {}
intra_week_map = {}
for d in intra_data:
    intra_sec_map[d['secao']]  = intra_sec_map.get(d['secao'],  0) + 1
    intra_week_map[d['semana']]= intra_week_map.get(d['semana'],0) + 1

# Agregações inter
inter_sec_map  = {}
inter_week_map = {}
for d in inter_data:
    inter_sec_map[d['secao']]  = inter_sec_map.get(d['secao'],  0) + 1
    inter_week_map[d['semana']]= inter_week_map.get(d['semana'],0) + 1

# KPIs intra
intra_sec_sorted  = sorted(intra_sec_map.items(),  key=lambda x: -x[1])
intra_week_sorted = sorted(intra_week_map.items(), key=lambda x: -x[1])
kpi_intra_total = len(intra_data)
kpi_intra_sec   = len(intra_sec_map)
kpi_intra_top   = intra_sec_sorted[0][0]  if intra_sec_sorted  else '—'
kpi_intra_top_n = intra_sec_sorted[0][1]  if intra_sec_sorted  else 0
kpi_intra_pico  = intra_week_sorted[0][1] if intra_week_sorted else 0
kpi_intra_pico_sem = intra_week_sorted[0][0] if intra_week_sorted else '—'

# KPIs inter
inter_sec_sorted  = sorted(inter_sec_map.items(),  key=lambda x: -x[1])
inter_week_sorted = sorted(inter_week_map.items(), key=lambda x: -x[1])
kpi_inter_total = len(inter_data)
kpi_inter_sec   = len(inter_sec_map)
kpi_inter_top   = inter_sec_sorted[0][0]  if inter_sec_sorted  else '—'
kpi_inter_top_n = inter_sec_sorted[0][1]  if inter_sec_sorted  else 0
kpi_inter_pico  = inter_week_sorted[0][1] if inter_week_sorted else 0
kpi_inter_pico_sem = inter_week_sorted[0][0] if inter_week_sorted else '—'

# Período
datas = exc_df['DATA'].dropna()
periodo = f"{datas.min().strftime('%d/%m/%Y')} a {datas.max().strftime('%d/%m/%Y')}" if len(datas) else '—'
gerado_em = datetime.now().strftime('%d/%m/%Y às %H:%M')

# ── EMBUTIR DADOS NO HTML ─────────────────────────────────────
print("🎨  Gerando HTML...")

# Serializar para JS
def js(obj):
    return json.dumps(obj, ensure_ascii=False)

# Seções positivas e negativas
sec_pos = [s for s in sec_arr if s['total'] > 0][:8]
sec_neg = [s for s in sec_arr if s['total'] < 0]

def sec_rows_html(rows, cor):
    if not rows:
        return '<tr><td colspan="6" style="text-align:center;color:var(--muted);padding:16px;">Nenhuma seção nesta categoria</td></tr>'
    html = ''
    for i, d in enumerate(rows, 1):
        total_fmt = fmt_bh_h(round(d['total'] * 60))
        max_fmt   = fmt_bh_h(round((d['maxPos'] if cor == 'blue' else d['maxNeg']) * 60))
        risk = ('CRÍTICO' if abs(d['total']) >= 15 else
                'ATENÇÃO'  if abs(d['total']) >= 10 else
                'MONITORAR' if abs(d['total']) >= 5 else 'NORMAL')
        pill_cls = ('p-red' if risk == 'CRÍTICO' else
                    'p-amber' if risk == 'ATENÇÃO' else
                    'p-blue' if risk == 'MONITORAR' else 'p-green')
        html += f'''<tr>
          <td class="mono">{i}</td>
          <td style="font-weight:500;">{d["secao"]}</td>
          <td style="text-align:center;">{d["n"]}</td>
          <td class="mono" style="color:var(--{cor});font-weight:700;">{total_fmt}</td>
          <td class="mono" style="color:var(--{cor});">{max_fmt}</td>
          <td><span class="pill {pill_cls}">{risk}</span></td>
        </tr>'''
    return html

def faixa_pill(f):
    if 'CRITICO_POS'  in f: return '<span class="pill p-red">CRÍTICO +20h</span>'
    if 'ALERTA_POS'   in f: return '<span class="pill p-orange">ALERTA +15h</span>'
    if 'ATENCAO_POS'  in f: return '<span class="pill p-amber">ATENÇÃO +10h</span>'
    if 'ALERTA_NEG'   in f: return '<span class="pill p-red">ALERTA −15h</span>'
    if 'ATENCAO_NEG'  in f: return '<span class="pill p-amber">ATENÇÃO −10h</span>'
    return '<span class="pill p-green">NORMAL</span>'

def acao_pill(a):
    if a == 'FAZER COMPENSACAO': return '<span class="pill p-red">COMPENSAÇÃO</span>'
    if a == 'MONITORAR SALDO':   return '<span class="pill p-amber">MONITORAR</span>'
    return '<span class="pill p-green">SEM AÇÃO</span>'

def bh_rows_html(rows):
    html = ''
    for i, d in enumerate(rows, 1):
        cor = 'blue' if d['bh_h'] >= 0 else 'red'
        html += f'''<tr>
          <td class="mono">{i}</td>
          <td style="font-weight:500;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="{d["nome"]}">{d["nome"]}</td>
          <td style="color:var(--muted);font-size:11px;">{d["secao"]}</td>
          <td style="font-size:11px;color:var(--muted);">{d["funcao"]}</td>
          <td class="mono" style="font-weight:700;color:var(--{cor});">{d["saldo"]}</td>
          <td>{faixa_pill(d["faixa"])}</td>
          <td>{acao_pill(d["acao"])}</td>
        </tr>'''
    return html

def intra_rows_html(rows):
    html = ''
    for i, d in enumerate(rows, 1):
        html += f'''<tr>
          <td class="mono">{i}</td>
          <td style="font-weight:500;max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="{d["nome"]}">{d["nome"]}</td>
          <td style="color:var(--muted);font-size:11px;white-space:nowrap;">{d["secao"]}</td>
          <td class="mono">{d["data"]}</td>
          <td class="mono" style="color:var(--green);">{d["b1"]}</td>
          <td class="mono" style="color:var(--muted);">{d["b2"]}</td>
          <td class="mono" style="color:var(--muted);">{d["b3"]}</td>
          <td class="mono" style="color:var(--green);">{d["b4"]}</td>
          <td class="mono">{d["htrab"]}</td>
          <td class="mono" style="color:var(--amber);font-weight:600;">{d["intra"]}</td>
          <td class="mono" style="color:var(--red);font-weight:600;">{d["deficit"]}</td>
          <td><span class="pill p-red">GRAVE</span></td>
        </tr>'''
    return html

def inter_rows_html(rows):
    html = ''
    for i, d in enumerate(rows, 1):
        html += f'''<tr>
          <td class="mono">{i}</td>
          <td style="font-weight:500;max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="{d["nome"]}">{d["nome"]}</td>
          <td style="color:var(--muted);font-size:11px;white-space:nowrap;">{d["secao"]}</td>
          <td class="mono">{d["data"]}</td>
          <td class="mono" style="color:var(--green);">{d["b1"]}</td>
          <td class="mono" style="color:var(--muted);">{d["b2"]}</td>
          <td class="mono" style="color:var(--muted);">{d["b3"]}</td>
          <td class="mono" style="color:var(--green);">{d["b4"]}</td>
          <td class="mono" style="color:var(--amber);font-weight:600;">{d["inter"]}</td>
          <td class="mono" style="color:var(--red);font-weight:600;">{d["deficit"]}</td>
          <td><span class="pill p-red">GRAVE</span></td>
        </tr>'''
    return html

def faixa_legend_html(faixas):
    html = ''
    for f in faixas:
        html += f'<div class="leg"><div class="leg-sq" style="background:{f["c"]}"></div>{f["lbl"]}: <b>{f["n"]}</b></div>'
    return html

# Dados para gráficos JS
chart_bh_sec_labels = json.dumps([s['secao'] for s in sec_arr], ensure_ascii=False)
chart_bh_sec_data   = json.dumps([round(s['total'], 1) for s in sec_arr])
chart_bh_sec_colors = json.dumps(['rgba(37,99,184,.75)' if s['total'] >= 0 else 'rgba(192,57,43,.75)' for s in sec_arr])

chart_faixa_labels = json.dumps([f['lbl'] for f in faixas], ensure_ascii=False)
chart_faixa_data   = json.dumps([f['n']   for f in faixas])
chart_faixa_colors = json.dumps([f['c']   for f in faixas])

chart_intra_sec_labels = json.dumps([k for k, _ in intra_sec_sorted], ensure_ascii=False)
chart_intra_sec_data   = json.dumps([v for _, v in intra_sec_sorted])
chart_intra_week_labels= json.dumps(sorted(intra_week_map.keys()), ensure_ascii=False)
chart_intra_week_data  = json.dumps([intra_week_map[k] for k in sorted(intra_week_map.keys())])

chart_inter_sec_labels = json.dumps([k for k, _ in inter_sec_sorted], ensure_ascii=False)
chart_inter_sec_data   = json.dumps([v for _, v in inter_sec_sorted])
chart_inter_week_labels= json.dumps(sorted(inter_week_map.keys()), ensure_ascii=False)
chart_inter_week_data  = json.dumps([inter_week_map[k] for k in sorted(inter_week_map.keys())])

# Seções únicas para filtro
secoes_unicas = json.dumps(sorted(set(d['secao'] for d in bh_colab)), ensure_ascii=False)

# ── TEMPLATE HTML ─────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>TechSolutions — Dashboard RH</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<style>
:root{{
  --bg:#f4f5f7;--white:#fff;--border:#e3e6ed;--text:#1e2235;--muted:#7b83a0;
  --orange:#f47c20;--orange-l:#fff4ec;--orange-bd:#ffd4ae;
  --blue:#2563b8;--blue-l:#edf3fc;--blue-bd:#bad0f7;
  --green:#1e9e5e;--green-l:#eaf7f1;--green-bd:#a8e4c8;
  --amber:#c97b00;--amber-l:#fef8ec;--amber-bd:#f5d98a;
  --red:#c0392b;--red-l:#fdf1f0;--red-bd:#f5bcb7;
  --font:'Inter',sans-serif;--r:10px;--r-sm:7px;--sh:0 1px 3px rgba(0,0,0,0.07);
}}
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0;}}
body{{background:var(--bg);color:var(--text);font-family:var(--font);font-size:13px;line-height:1.5;}}
.topbar{{background:var(--orange);padding:14px 28px 12px;}}
.topbar-title{{font-size:17px;font-weight:700;color:#fff;}}
.topbar-sub{{font-size:11px;color:rgba(255,255,255,.82);margin-top:2px;}}
.nav{{background:var(--white);border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;padding:0 24px;flex-wrap:wrap;gap:8px;}}
.tabs{{display:flex;}}
.tab{{background:none;border:none;border-bottom:3px solid transparent;padding:13px 16px;font-size:12px;font-weight:500;color:var(--muted);cursor:pointer;font-family:var(--font);transition:all .15s;white-space:nowrap;}}
.tab:hover{{color:var(--orange);}}
.tab.active{{color:var(--orange);border-bottom-color:var(--orange);}}
.nav-chips{{display:flex;gap:6px;padding:8px 0;}}
.chip{{background:var(--bg);border:1px solid var(--border);border-radius:20px;padding:4px 11px;font-size:10px;color:var(--muted);}}
.wrap{{max-width:1300px;margin:0 auto;padding:20px 20px 48px;}}
.panel{{display:none;}}
.panel.active{{display:block;animation:up .22s ease both;}}
@keyframes up{{from{{opacity:0;transform:translateY(8px)}}to{{opacity:1;transform:none}}}}
.kpi-strip{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px;}}
.kpi{{background:var(--white);border:1px solid var(--border);border-radius:var(--r);padding:15px 18px;box-shadow:var(--sh);border-top:3px solid var(--kc,var(--border));}}
.kpi-lbl{{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:var(--kc,var(--muted));margin-bottom:7px;}}
.kpi-val{{font-size:28px;font-weight:700;color:var(--kc,var(--text));line-height:1;}}
.kpi-sub{{font-size:10px;color:var(--muted);margin-top:5px;}}
.g2{{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px;}}
.g3{{display:grid;grid-template-columns:2fr 1fr;gap:14px;margin-bottom:14px;}}
.card{{background:var(--white);border:1px solid var(--border);border-radius:var(--r);padding:18px 20px;box-shadow:var(--sh);}}
.card-hdr{{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:14px;gap:8px;}}
.card-title{{font-size:11px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;display:flex;align-items:center;gap:7px;}}
.card-title::before{{content:'';width:3px;height:13px;background:var(--orange);border-radius:2px;flex-shrink:0;}}
.card-tag{{font-size:9px;font-weight:600;padding:3px 9px;border-radius:10px;background:var(--amber-l);color:var(--amber);border:1px solid var(--amber-bd);white-space:nowrap;}}
.chart-box{{position:relative;width:100%;}}
.tbl-wrap{{overflow-x:auto;}}
table{{width:100%;border-collapse:collapse;}}
thead th{{font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;padding:7px 10px;border-bottom:2px solid var(--border);text-align:left;white-space:nowrap;background:var(--bg);}}
tbody td{{padding:7px 10px;border-bottom:1px solid var(--border);font-size:12px;color:var(--text);vertical-align:middle;}}
tbody tr:last-child td{{border-bottom:none;}}
tbody tr:hover td{{background:#f9fafc;}}
.mono{{font-family:'Courier New',monospace;font-size:11px;}}
.pill{{display:inline-block;font-size:9px;font-weight:600;padding:2px 8px;border-radius:10px;white-space:nowrap;}}
.p-red{{background:var(--red-l);color:var(--red);border:1px solid var(--red-bd);}}
.p-amber{{background:var(--amber-l);color:var(--amber);border:1px solid var(--amber-bd);}}
.p-blue{{background:var(--blue-l);color:var(--blue);border:1px solid var(--blue-bd);}}
.p-green{{background:var(--green-l);color:var(--green);border:1px solid var(--green-bd);}}
.p-orange{{background:var(--orange-l);color:var(--orange);border:1px solid var(--orange-bd);}}
.p-gray{{background:var(--bg);color:var(--muted);border:1px solid var(--border);}}
.legend{{display:flex;flex-wrap:wrap;gap:12px;margin-top:10px;}}
.leg{{display:flex;align-items:center;gap:5px;font-size:10px;color:var(--muted);}}
.leg-sq{{width:9px;height:9px;border-radius:2px;flex-shrink:0;}}
.donut-wrap{{position:relative;display:flex;align-items:center;justify-content:center;}}
.donut-center{{position:absolute;text-align:center;pointer-events:none;}}
.donut-num{{font-size:22px;font-weight:700;}}
.donut-lbl{{font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;}}
.filter-row{{display:flex;align-items:center;gap:8px;margin-bottom:14px;flex-wrap:wrap;}}
.filter-lbl{{font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;}}
.filter-sel{{background:var(--white);border:1px solid var(--border);border-radius:var(--r-sm);padding:5px 11px;font-size:11px;color:var(--text);cursor:pointer;font-family:var(--font);outline:none;transition:border-color .15s;}}
.filter-sel:focus,.filter-sel:hover{{border-color:var(--orange);}}
.footer{{border-top:1px solid var(--border);padding-top:14px;margin-top:4px;display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px;}}
.footer-txt{{font-size:10px;color:var(--muted);}}
@media(max-width:900px){{.kpi-strip{{grid-template-columns:repeat(2,1fr);}}.g2,.g3{{grid-template-columns:1fr;}}}}
</style>
</head>
<body>
<div class="topbar">
  <div class="topbar-title">TECHSOLUTIONS — DASHBOARD RH</div>
  <div class="topbar-sub">Filial 58 &nbsp;·&nbsp; Período: {periodo} &nbsp;·&nbsp; Gerado em {gerado_em}</div>
</div>
<div class="nav">
  <div class="tabs">
    <button class="tab active" onclick="showPanel('bh')">Banco de Horas</button>
    <button class="tab" onclick="showPanel('intra')">Intrajornada</button>
    <button class="tab" onclick="showPanel('inter')">Interjornada</button>
  </div>
  <div class="nav-chips">
    <div class="chip">{kpi_bh_total} colaboradores (BH)</div>
    <div class="chip">{periodo}</div>
  </div>
</div>
<div class="wrap">

<!-- ════ BANCO DE HORAS ════ -->
<div id="panel-bh" class="panel active">
  <div class="kpi-strip">
    <div class="kpi" style="--kc:var(--blue)"><div class="kpi-lbl">Total analisado</div><div class="kpi-val">{kpi_bh_total}</div><div class="kpi-sub">colaboradores</div></div>
    <div class="kpi" style="--kc:var(--green)"><div class="kpi-lbl">Dentro do controle</div><div class="kpi-val">{kpi_bh_normal}</div><div class="kpi-sub">faixa normal</div></div>
    <div class="kpi" style="--kc:var(--amber)"><div class="kpi-lbl">Atenção / Alerta</div><div class="kpi-val">{kpi_bh_atencao}</div><div class="kpi-sub">monitorar saldo</div></div>
    <div class="kpi" style="--kc:var(--red)"><div class="kpi-lbl">Crítico</div><div class="kpi-val">{kpi_bh_critico}</div><div class="kpi-sub">compensação necessária</div></div>
  </div>
  <div class="g2">
    <div class="card">
      <div class="card-hdr"><div class="card-title">Saldo por seção</div><div class="card-tag">TOTALGERAL</div></div>
      <div class="chart-box" style="height:300px"><canvas id="cBhSec" role="img" aria-label="Saldo de banco de horas por seção."></canvas></div>
    </div>
    <div class="card">
      <div class="card-hdr"><div class="card-title">Distribuição por faixa</div><div class="card-tag">{kpi_bh_total} colab.</div></div>
      <div class="donut-wrap" style="height:180px;margin-bottom:10px;">
        <div class="chart-box" style="height:180px;"><canvas id="cBhFaixa" role="img" aria-label="Distribuição por faixa."></canvas></div>
        <div class="donut-center"><div class="donut-num" style="color:var(--orange)">{kpi_bh_total}</div><div class="donut-lbl">total</div></div>
      </div>
      <div class="legend" style="justify-content:center;">{faixa_legend_html(faixas)}</div>
    </div>
  </div>
  <div class="g2" style="margin-bottom:14px;">
    <div class="card">
      <div class="card-hdr"><div class="card-title">Seções com maior saldo positivo</div><div class="card-tag" style="background:var(--blue-l);color:var(--blue);border-color:var(--blue-bd);">Banco acumulado</div></div>
      <div class="tbl-wrap"><table><thead><tr><th>#</th><th>Seção</th><th>Colab.</th><th>Saldo total</th><th>Maior saldo</th><th>Situação</th></tr></thead>
      <tbody>{sec_rows_html(sec_pos, 'blue')}</tbody></table></div>
    </div>
    <div class="card">
      <div class="card-hdr"><div class="card-title">Seções com saldo negativo</div><div class="card-tag" style="background:var(--red-l);color:var(--red);border-color:var(--red-bd);">Atenção</div></div>
      <div class="tbl-wrap"><table><thead><tr><th>#</th><th>Seção</th><th>Colab.</th><th>Saldo total</th><th>Maior negativo</th><th>Situação</th></tr></thead>
      <tbody>{sec_rows_html(sec_neg, 'red')}</tbody></table></div>
    </div>
  </div>
  <div class="card">
    <div class="card-hdr"><div class="card-title">Ranking completo</div><div class="card-tag" id="bhColabTag">{kpi_bh_total} colaboradores</div></div>
    <div class="filter-row">
      <span class="filter-lbl">Seção:</span>
      <select class="filter-sel" id="bhSecFil" onchange="filterBh()"><option value="">Todas</option></select>
      <span class="filter-lbl" style="margin-left:8px;">Ação:</span>
      <select class="filter-sel" id="bhAcaoFil" onchange="filterBh()">
        <option value="">Todas</option><option>FAZER COMPENSACAO</option><option>MONITORAR SALDO</option><option>SEM ACAO</option>
      </select>
      <span class="filter-lbl" style="margin-left:8px;">Ordenar:</span>
      <select class="filter-sel" id="bhSort" onchange="filterBh()">
        <option value="abs">Maior saldo (abs)</option><option value="pos">Maior positivo</option><option value="neg">Maior negativo</option>
      </select>
    </div>
    <div class="tbl-wrap" style="max-height:360px;overflow-y:auto;">
      <table><thead><tr><th>#</th><th>Nome</th><th>Seção</th><th>Função</th><th>Saldo</th><th>Faixa</th><th>Ação</th></tr></thead>
      <tbody id="tbBh">{bh_rows_html(bh_colab)}</tbody></table>
    </div>
  </div>
</div>

<!-- ════ INTRAJORNADA ════ -->
<div id="panel-intra" class="panel">
  <div class="kpi-strip">
    <div class="kpi" style="--kc:var(--red)"><div class="kpi-lbl">Ocorrências</div><div class="kpi-val">{kpi_intra_total}</div><div class="kpi-sub">intrajornada irregular</div></div>
    <div class="kpi" style="--kc:var(--amber)"><div class="kpi-lbl">Seções afetadas</div><div class="kpi-val">{kpi_intra_sec}</div><div class="kpi-sub">com violações</div></div>
    <div class="kpi" style="--kc:var(--orange)"><div class="kpi-lbl">Seção crítica</div><div class="kpi-val" style="font-size:15px;padding-top:4px;">{kpi_intra_top}</div><div class="kpi-sub">{kpi_intra_top_n} ocorrência(s)</div></div>
    <div class="kpi" style="--kc:var(--blue)"><div class="kpi-lbl">Pico semanal</div><div class="kpi-val">{kpi_intra_pico}</div><div class="kpi-sub">{kpi_intra_pico_sem}</div></div>
  </div>
  <div class="g3">
    <div class="card">
      <div class="card-hdr"><div class="card-title">Ocorrências por seção</div><div class="card-tag">Intrajornada</div></div>
      <div class="chart-box" style="height:260px"><canvas id="cIntraSec" role="img" aria-label="Violações intrajornada por seção."></canvas></div>
    </div>
    <div class="card">
      <div class="card-hdr"><div class="card-title">Evolução semanal</div><div class="card-tag">Por semana</div></div>
      <div class="chart-box" style="height:260px"><canvas id="cIntraWeek" role="img" aria-label="Evolução semanal intrajornada."></canvas></div>
    </div>
  </div>
  <div class="card">
    <div class="card-hdr"><div class="card-title">Detalhe completo das ocorrências</div><div class="card-tag">Intervalo irregular</div></div>
    <div class="tbl-wrap">
      <table><thead><tr><th>#</th><th>Colaborador</th><th>Seção</th><th>Data</th><th>Entrada</th><th>Saída almoço</th><th>Retorno</th><th>Saída</th><th>Jornada</th><th>Intervalo real</th><th>Déficit</th><th>Severidade</th></tr></thead>
      <tbody>{intra_rows_html(intra_data)}</tbody></table>
    </div>
  </div>
</div>

<!-- ════ INTERJORNADA ════ -->
<div id="panel-inter" class="panel">
  <div class="kpi-strip">
    <div class="kpi" style="--kc:var(--red)"><div class="kpi-lbl">Ocorrências</div><div class="kpi-val">{kpi_inter_total}</div><div class="kpi-sub">descanso &lt; 11h</div></div>
    <div class="kpi" style="--kc:var(--amber)"><div class="kpi-lbl">Seções afetadas</div><div class="kpi-val">{kpi_inter_sec}</div><div class="kpi-sub">com violações</div></div>
    <div class="kpi" style="--kc:var(--orange)"><div class="kpi-lbl">Seção crítica</div><div class="kpi-val" style="font-size:15px;padding-top:4px;">{kpi_inter_top}</div><div class="kpi-sub">{kpi_inter_top_n} ocorrência(s)</div></div>
    <div class="kpi" style="--kc:var(--blue)"><div class="kpi-lbl">Pico semanal</div><div class="kpi-val">{kpi_inter_pico}</div><div class="kpi-sub">{kpi_inter_pico_sem}</div></div>
  </div>
  <div class="g3">
    <div class="card">
      <div class="card-hdr"><div class="card-title">Ocorrências por seção</div><div class="card-tag">Interjornada</div></div>
      <div class="chart-box" style="height:260px"><canvas id="cInterSec" role="img" aria-label="Violações interjornada por seção."></canvas></div>
    </div>
    <div class="card">
      <div class="card-hdr"><div class="card-title">Evolução semanal</div><div class="card-tag">Por semana</div></div>
      <div class="chart-box" style="height:260px"><canvas id="cInterWeek" role="img" aria-label="Evolução semanal interjornada."></canvas></div>
    </div>
  </div>
  <div class="card">
    <div class="card-hdr"><div class="card-title">Detalhe completo das ocorrências</div><div class="card-tag">Descanso &lt; 11h</div></div>
    <div class="tbl-wrap">
      <table><thead><tr><th>#</th><th>Colaborador</th><th>Seção</th><th>Data</th><th>Entrada</th><th>Saída almoço</th><th>Retorno</th><th>Saída</th><th>Interjornada real</th><th>Déficit</th><th>Severidade</th></tr></thead>
      <tbody>{inter_rows_html(inter_data)}</tbody></table>
    </div>
  </div>
</div>

<div class="footer">
  <div class="footer-txt">TechSolutions Gestão de Pessoas · Filial 58</div>
  <div class="footer-txt">Gerado automaticamente em {gerado_em} · {periodo}</div>
</div>
</div>

<script>
const C={{blue:'#2563b8',green:'#1e9e5e',amber:'#c97b00',red:'#c0392b',orange:'#f47c20',muted:'#9ba3bb'}};
const GRID='rgba(0,0,0,0.05)',TXT='#9ba3bb',FONT={{family:"'Inter',sans-serif",size:11}};
const TT={{backgroundColor:'#fff',titleColor:'#1e2235',bodyColor:'#7b83a0',borderColor:'#e3e6ed',borderWidth:1,padding:10}};
const charts={{}};
function mkChart(id,cfg){{if(charts[id])charts[id].destroy();charts[id]=new Chart(document.getElementById(id),cfg);}}

// Banco de horas — gráfico seção
mkChart('cBhSec',{{type:'bar',data:{{labels:{chart_bh_sec_labels},datasets:[{{label:'Saldo BH',data:{chart_bh_sec_data},backgroundColor:{chart_bh_sec_colors},borderRadius:4,borderSkipped:false}}]}},
  options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{...TT,callbacks:{{label:t=>`${{t.parsed.x>=0?'+':''}}${{t.parsed.x.toFixed(1)}}h`}}}}}},
  scales:{{x:{{ticks:{{color:TXT,font:FONT,callback:v=>`${{v}}h`}},grid:{{color:GRID}},border:{{display:false}}}},y:{{ticks:{{color:TXT,font:FONT}},grid:{{display:false}}}}}}}}}});

// Banco de horas — donut faixas
mkChart('cBhFaixa',{{type:'doughnut',data:{{labels:{chart_faixa_labels},datasets:[{{data:{chart_faixa_data},backgroundColor:{chart_faixa_colors},borderWidth:3,borderColor:'#fff',hoverOffset:5}}]}},
  options:{{responsive:true,maintainAspectRatio:false,cutout:'64%',plugins:{{legend:{{display:false}},tooltip:{{...TT,callbacks:{{label:t=>`${{t.label}}: ${{t.parsed}}`}}}}}}}}}});

// Intrajornada — seção
mkChart('cIntraSec',{{type:'bar',data:{{labels:{chart_intra_sec_labels},datasets:[{{label:'Ocorrências',data:{chart_intra_sec_data},backgroundColor:'rgba(192,57,43,.75)',borderRadius:4,borderSkipped:false}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{...TT}}}},
  scales:{{x:{{ticks:{{color:TXT,font:FONT,maxRotation:30}},grid:{{display:false}}}},y:{{ticks:{{color:TXT,font:FONT,stepSize:1}},grid:{{color:GRID}},border:{{display:false}}}}}}}}}});

// Intrajornada — semanal
mkChart('cIntraWeek',{{type:'line',data:{{labels:{chart_intra_week_labels},datasets:[{{label:'Ocorrências',data:{chart_intra_week_data},borderColor:C.red,backgroundColor:'rgba(192,57,43,.07)',pointRadius:6,pointBackgroundColor:C.red,borderWidth:2,tension:.3,fill:true}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{...TT}}}},
  scales:{{x:{{ticks:{{color:TXT,font:FONT}},grid:{{display:false}}}},y:{{ticks:{{color:TXT,font:FONT,stepSize:1}},grid:{{color:GRID}},border:{{display:false}},min:0}}}}}}}});

// Interjornada — seção
mkChart('cInterSec',{{type:'bar',data:{{labels:{chart_inter_sec_labels},datasets:[{{label:'Ocorrências',data:{chart_inter_sec_data},backgroundColor:'rgba(192,57,43,.75)',borderRadius:4,borderSkipped:false}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{...TT}}}},
  scales:{{x:{{ticks:{{color:TXT,font:FONT,maxRotation:30}},grid:{{display:false}}}},y:{{ticks:{{color:TXT,font:FONT,stepSize:1}},grid:{{color:GRID}},border:{{display:false}}}}}}}}}});

// Interjornada — semanal
mkChart('cInterWeek',{{type:'line',data:{{labels:{chart_inter_week_labels},datasets:[{{label:'Ocorrências',data:{chart_inter_week_data},borderColor:C.orange,backgroundColor:'rgba(244,124,32,.07)',pointRadius:6,pointBackgroundColor:C.orange,borderWidth:2,tension:.3,fill:true}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{...TT}}}},
  scales:{{x:{{ticks:{{color:TXT,font:FONT}},grid:{{display:false}}}},y:{{ticks:{{color:TXT,font:FONT,stepSize:1}},grid:{{color:GRID}},border:{{display:false}},min:0}}}}}}}});

// Navegação
const PANELS=['bh','intra','inter'];
let activePanel='bh';
function showPanel(name){{
  document.getElementById('panel-'+activePanel).classList.remove('active');
  document.querySelectorAll('.tab').forEach((t,i)=>t.classList.toggle('active',PANELS[i]===name));
  activePanel=name;
  document.getElementById('panel-'+name).classList.add('active');
}}

// Filtro banco de horas
const BH_ALL = {js(bh_colab)};
const SECOES  = {secoes_unicas};
const sel = document.getElementById('bhSecFil');
SECOES.forEach(s=>{{const o=document.createElement('option');o.value=s;o.textContent=s;sel.appendChild(o);}});

function faixaPill(f){{
  if(f.includes('CRITICO_POS'))  return'<span class="pill p-red">CRÍTICO +20h</span>';
  if(f.includes('ALERTA_POS'))   return'<span class="pill p-orange">ALERTA +15h</span>';
  if(f.includes('ATENCAO_POS'))  return'<span class="pill p-amber">ATENÇÃO +10h</span>';
  if(f.includes('ALERTA_NEG'))   return'<span class="pill p-red">ALERTA −15h</span>';
  if(f.includes('ATENCAO_NEG'))  return'<span class="pill p-amber">ATENÇÃO −10h</span>';
  return'<span class="pill p-green">NORMAL</span>';
}}
function acaoPill(a){{
  if(a==='FAZER COMPENSACAO') return'<span class="pill p-red">COMPENSAÇÃO</span>';
  if(a==='MONITORAR SALDO')   return'<span class="pill p-amber">MONITORAR</span>';
  return'<span class="pill p-green">SEM AÇÃO</span>';
}}
function filterBh(){{
  const sec  = document.getElementById('bhSecFil').value;
  const acao = document.getElementById('bhAcaoFil').value;
  const sort = document.getElementById('bhSort').value;
  let data = [...BH_ALL];
  if(sec)  data = data.filter(d=>d.secao===sec);
  if(acao) data = data.filter(d=>d.acao===acao);
  if(sort==='abs') data.sort((a,b)=>Math.abs(b.bh_h)-Math.abs(a.bh_h));
  if(sort==='pos') data.sort((a,b)=>b.bh_h-a.bh_h);
  if(sort==='neg') data.sort((a,b)=>a.bh_h-b.bh_h);
  document.getElementById('bhColabTag').textContent = data.length+' colaboradores';
  const tb = document.getElementById('tbBh');
  tb.innerHTML='';
  data.forEach((d,i)=>{{
    const c = d.bh_h>=0?C.blue:C.red;
    tb.innerHTML+=`<tr>
      <td class="mono">${{i+1}}</td>
      <td style="font-weight:500;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${{d.nome}}">${{d.nome}}</td>
      <td style="color:var(--muted);font-size:11px;">${{d.secao}}</td>
      <td style="font-size:11px;color:var(--muted);">${{d.funcao}}</td>
      <td class="mono" style="font-weight:700;color:${{c}};">${{d.saldo}}</td>
      <td>${{faixaPill(d.faixa)}}</td>
      <td>${{acaoPill(d.acao)}}</td>
    </tr>`;
  }});
}}
</script>
</body>
</html>"""

output_path = OUTPUT_DIR / "index.html"
with open(output_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"✅  Dashboard gerado em: {output_path}")
print(f"   Colaboradores BH : {kpi_bh_total}")
print(f"   Intrajornada     : {kpi_intra_total} ocorrências")
print(f"   Interjornada     : {kpi_inter_total} ocorrências")
print(f"   Período          : {periodo}")

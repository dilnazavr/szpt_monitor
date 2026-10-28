#!/usr/bin/env python3
"""
SZPT Monitor — веб-интерфейс
Запуск: python app.py
Открыть: http://127.0.0.1:5000
"""
from __future__ import annotations

import json
import math
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template_string,
    request,
    send_from_directory,
    url_for,
)

from config import (
    DEFAULT_THRESHOLD_LETTER,
    DEFAULT_THRESHOLD_MONITOR,
    ETALON_DIR,
    LETTERS_DIR,
    OUTPUT_DIR,
    OWS_TOKEN,
)

app = Flask(__name__)
app.secret_key = "szpt-monitor-aktobe-2026"

RESULTS_JSON = OUTPUT_DIR / "last_results.json"

# Текстовые поля — NaN из Excel нельзя резать срезом [:80]
_TEXT_KEYS = (
    "year", "date", "etalon_period", "product", "lot_name",
    "customer", "customer_bin", "supplier_name", "supplier_biin",
    "lot_number", "number_anno", "contract_number", "status",
    "lot_url", "lot_id",
)
_NUM_KEYS = (
    "unit_price", "etalon", "percent", "amount", "amount_szpt",
    "difference", "qty",
)


def _is_nan(v: Any) -> bool:
    if v is None:
        return True
    try:
        if isinstance(v, float) and math.isnan(v):
            return True
    except (TypeError, ValueError):
        pass
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return False


def _clean_row(r: Dict[str, Any]) -> Dict[str, Any]:
    """NaN/None → '' для текста, None для чисел — чтобы Jinja не падала на [:N]."""
    out = dict(r)
    for k in _TEXT_KEYS:
        if k not in out:
            continue
        v = out[k]
        if _is_nan(v):
            out[k] = ""
        else:
            out[k] = str(v)
    for k in _NUM_KEYS:
        if k not in out:
            continue
        v = out[k]
        if _is_nan(v):
            out[k] = None
        else:
            try:
                out[k] = float(v)
            except (TypeError, ValueError):
                out[k] = None
    return out


def _clean_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return [_clean_row(r) for r in rows if isinstance(r, dict)]

BASE = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _list_files(folder: Path, pattern: str = "*") -> List[str]:
    if not folder.exists():
        return []
    return sorted([p.name for p in folder.glob(pattern)], reverse=True)[:40]


def _load_results() -> List[Dict[str, Any]]:
    if RESULTS_JSON.exists():
        try:
            with open(RESULTS_JSON, encoding="utf-8") as f:
                data = json.load(f)
            return _clean_rows(data.get("rows") or [])
        except Exception:
            pass
    # fallback: последний xlsx
    xlsx_files = sorted(OUTPUT_DIR.glob("zavysheniya_*.xlsx"), reverse=True)
    if not xlsx_files:
        xlsx_files = sorted(OUTPUT_DIR.glob("*.xlsx"), reverse=True)
    if not xlsx_files:
        return []
    try:
        df = pd.read_excel(xlsx_files[0])
        colmap = {
            "Год": "year", "Дата": "date", "Период эталона": "etalon_period",
            "Товар": "product",
            "Предмет лота": "lot_name", "Цена за ед.": "unit_price",
            "Эталон Актобе": "etalon", "% отклонения": "percent",
            "Сумма": "amount", "Сумма по договору": "amount",
            "Сумма по цене СЗПТ": "amount_szpt", "Разница": "difference",
            "Количество": "qty",
            "Заказчик": "customer", "БИН заказчика": "customer_bin",
            "Поставщик": "supplier_name", "БИН поставщика": "supplier_biin",
            "Номер лота": "lot_number", "Номер объявления": "number_anno",
            "Номер договора": "contract_number", "Статус": "status",
            "ID лота": "lot_id", "Ссылка": "lot_url",
        }
        df = df.rename(columns={k: v for k, v in colmap.items() if k in df.columns})
        df = df.where(pd.notnull(df), None)
        return _clean_rows(df.to_dict(orient="records"))
    except Exception:
        return []


def _save_results(rows: List[Dict[str, Any]], meta: Optional[Dict] = None) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "count": len(rows),
        "meta": meta or {},
        "rows": rows,
    }
    with open(RESULTS_JSON, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def _results_meta() -> Dict[str, Any]:
    if RESULTS_JSON.exists():
        try:
            with open(RESULTS_JSON, encoding="utf-8") as f:
                data = json.load(f)
            return {
                "updated_at": data.get("updated_at"),
                "count": data.get("count", 0),
                "meta": data.get("meta") or {},
            }
        except Exception:
            pass
    return {"updated_at": None, "count": 0, "meta": {}}


def _etalon_count() -> int:
    return len(list(ETALON_DIR.glob("*.xlsx"))) + len(list(ETALON_DIR.glob("*.xls")))


# ---------------------------------------------------------------------------
# Layout / CSS
# ---------------------------------------------------------------------------

LAYOUT = r"""
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{% block title %}SZPT Monitor — Актобе{% endblock %}</title>
  <style>
    :root {
      --bg: #0b1220;
      --bg2: #111827;
      --card: #1a2332;
      --card2: #1f2a3d;
      --border: #2d3a4f;
      --text: #e8eef7;
      --muted: #8b9cb3;
      --accent: #3b82f6;
      --accent2: #60a5fa;
      --ok: #34d399;
      --warn: #fbbf24;
      --err: #f87171;
      --danger: #ef4444;
      --radius: 14px;
      --shadow: 0 8px 32px rgba(0,0,0,.35);
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: "Segoe UI", system-ui, -apple-system, sans-serif;
      background: var(--bg);
      color: var(--text);
      min-height: 100vh;
      line-height: 1.55;
    }
    /* subtle gradient bg */
    body::before {
      content: "";
      position: fixed;
      inset: 0;
      background:
        radial-gradient(ellipse 80% 50% at 20% -10%, rgba(59,130,246,.18), transparent),
        radial-gradient(ellipse 60% 40% at 90% 10%, rgba(16,185,129,.1), transparent);
      pointer-events: none;
      z-index: 0;
    }
    .shell { position: relative; z-index: 1; max-width: 1200px; margin: 0 auto; padding: 0 1.25rem 3rem; }

    /* NAV */
    nav {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 1.1rem 0;
      border-bottom: 1px solid var(--border);
      margin-bottom: 1.75rem;
      flex-wrap: wrap;
      gap: .75rem;
    }
    .brand {
      display: flex;
      align-items: center;
      gap: .65rem;
      text-decoration: none;
      color: var(--text);
      font-weight: 700;
      font-size: 1.15rem;
    }
    .brand .logo {
      width: 36px; height: 36px;
      background: linear-gradient(135deg, #3b82f6, #10b981);
      border-radius: 10px;
      display: grid; place-items: center;
      font-size: 1rem;
    }
    .nav-links { display: flex; gap: .35rem; flex-wrap: wrap; }
    .nav-links a {
      color: var(--muted);
      text-decoration: none;
      padding: .45rem .85rem;
      border-radius: 8px;
      font-size: .9rem;
      font-weight: 500;
      transition: .15s;
    }
    .nav-links a:hover, .nav-links a.active {
      background: var(--card);
      color: var(--text);
    }
    .nav-links a.active { color: var(--accent2); }

    /* STATUS PILLS */
    .pills { display: flex; flex-wrap: wrap; gap: .5rem; margin-bottom: 1.5rem; }
    .pill {
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 999px;
      padding: .35rem .9rem;
      font-size: .8rem;
      color: var(--muted);
    }
    .pill strong { color: var(--text); }
    .pill.ok strong { color: var(--ok); }
    .pill.warn strong { color: var(--warn); }

    /* CARDS */
    .grid { display: grid; gap: 1.15rem; }
    @media (min-width: 800px) {
      .grid-2 { grid-template-columns: 1fr 1fr; }
      .grid-3 { grid-template-columns: repeat(3, 1fr); }
    }
    .card {
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 1.35rem 1.5rem;
      box-shadow: var(--shadow);
    }
    .card h2 {
      font-size: 1.05rem;
      font-weight: 650;
      margin-bottom: 1rem;
      display: flex;
      align-items: center;
      gap: .55rem;
    }
    .card h2 .num {
      width: 26px; height: 26px;
      background: linear-gradient(135deg, #3b82f6, #2563eb);
      border-radius: 8px;
      display: grid; place-items: center;
      font-size: .8rem;
      font-weight: 700;
      color: #fff;
    }

    label {
      display: block;
      font-size: .8rem;
      color: var(--muted);
      margin: .7rem 0 .25rem;
      font-weight: 500;
    }
    input, select {
      width: 100%;
      padding: .55rem .75rem;
      border-radius: 9px;
      border: 1px solid var(--border);
      background: var(--bg2);
      color: var(--text);
      font-size: .95rem;
    }
    input:focus, select:focus {
      outline: none;
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(59,130,246,.22);
    }
    .row { display: flex; gap: .85rem; flex-wrap: wrap; }
    .row > * { flex: 1; min-width: 110px; }

    button, .btn {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: .4rem;
      margin-top: 1rem;
      width: 100%;
      padding: .7rem 1rem;
      border: none;
      border-radius: 9px;
      background: linear-gradient(135deg, #3b82f6, #2563eb);
      color: #fff;
      font-size: .95rem;
      font-weight: 600;
      cursor: pointer;
      text-decoration: none;
      transition: .15s;
    }
    button:hover, .btn:hover { filter: brightness(1.08); transform: translateY(-1px); }
    .btn-sm { width: auto; margin-top: 0; padding: .4rem .75rem; font-size: .82rem; }
    .btn-ghost {
      background: transparent;
      border: 1px solid var(--border);
      color: var(--muted);
    }
    .btn-ghost:hover { border-color: var(--accent); color: var(--text); filter: none; }
    .btn-ok { background: linear-gradient(135deg, #10b981, #059669); }
    .btn-warn { background: linear-gradient(135deg, #f59e0b, #d97706); }

    /* FLASH */
    .flash {
      padding: .85rem 1.1rem;
      border-radius: 10px;
      margin-bottom: 1.15rem;
      font-size: .9rem;
      white-space: pre-wrap;
    }
    .flash.ok { background: rgba(52,211,153,.12); border: 1px solid rgba(52,211,153,.35); color: #6ee7b7; }
    .flash.err { background: rgba(248,113,113,.12); border: 1px solid rgba(248,113,113,.35); color: #fca5a5; }

    /* TABLE */
    .table-wrap {
      overflow-x: auto;
      border-radius: 12px;
      border: 1px solid var(--border);
      background: var(--card);
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: .85rem;
    }
    th {
      text-align: left;
      padding: .75rem 1rem;
      background: var(--card2);
      color: var(--muted);
      font-weight: 600;
      font-size: .75rem;
      text-transform: uppercase;
      letter-spacing: .03em;
      white-space: nowrap;
      position: sticky;
      top: 0;
    }
    td {
      padding: .7rem 1rem;
      border-top: 1px solid var(--border);
      vertical-align: top;
    }
    tr:hover td { background: rgba(59,130,246,.06); }
    .pct {
      font-weight: 700;
      padding: .15rem .45rem;
      border-radius: 6px;
      font-size: .82rem;
    }
    .pct.high { background: rgba(239,68,68,.2); color: #fca5a5; }
    .pct.mid { background: rgba(251,191,36,.18); color: #fcd34d; }
    .pct.low { background: rgba(52,211,153,.15); color: #6ee7b7; }
    .mono { font-family: ui-monospace, monospace; font-size: .8rem; color: var(--muted); }
    .lot-name { max-width: 280px; }
    .customer { max-width: 200px; }

    /* FILTERS */
    .filters {
      display: flex;
      flex-wrap: wrap;
      gap: .75rem;
      margin-bottom: 1rem;
      align-items: flex-end;
    }
    .filters .field { min-width: 140px; flex: 1; }
    .filters label { margin-top: 0; }
    .filters button { margin-top: 0; width: auto; padding: .55rem 1.1rem; }

    /* STATS */
    .stat-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
      gap: .85rem;
      margin-bottom: 1.35rem;
    }
    .stat {
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 1rem 1.15rem;
    }
    .stat .val { font-size: 1.5rem; font-weight: 700; line-height: 1.2; }
    .stat .lbl { font-size: .75rem; color: var(--muted); margin-top: .2rem; }
    .stat.danger .val { color: var(--err); }
    .stat.ok .val { color: var(--ok); }
    .stat.warn .val { color: var(--warn); }

    /* FILES */
    .file-list { list-style: none; }
    .file-list li {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 1rem;
      padding: .55rem 0;
      border-bottom: 1px solid var(--border);
      font-size: .9rem;
    }
    .file-list a { color: var(--accent2); text-decoration: none; }
    .file-list a:hover { text-decoration: underline; }
    .empty { color: var(--muted); font-size: .9rem; padding: .5rem 0; }

    /* DETAIL MODAL-ish panel */
    .detail-panel {
      background: var(--card2);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 1.25rem;
      margin-top: 1rem;
    }
    .detail-panel h3 { margin-bottom: .75rem; font-size: 1rem; }
    .detail-grid {
      display: grid;
      grid-template-columns: 140px 1fr;
      gap: .4rem .75rem;
      font-size: .9rem;
    }
    .detail-grid dt { color: var(--muted); }
    .detail-grid dd { word-break: break-word; }

    footer {
      margin-top: 2.5rem;
      text-align: center;
      color: var(--muted);
      font-size: .78rem;
    }
    .page-title {
      font-size: 1.35rem;
      font-weight: 700;
      margin-bottom: .35rem;
    }
    .page-sub { color: var(--muted); font-size: .9rem; margin-bottom: 1.25rem; }

    /* SEARCH highlight */
    .search-box { position: relative; flex: 2; min-width: 200px; }
    .search-box input { padding-left: 2.2rem; }
    .search-box::before {
      content: "🔍";
      position: absolute;
      left: .7rem;
      top: 50%;
      transform: translateY(-50%);
      font-size: .85rem;
      opacity: .6;
    }
  </style>
</head>
<body>
  <div class="shell">
    <nav>
      <a class="brand" href="{{ url_for('index') }}">
        <span class="logo">₸</span>
        SZPT Monitor
      </a>
      <div class="nav-links">
        <a href="{{ url_for('index') }}" class="{{ 'active' if request.endpoint == 'index' else '' }}">Главная</a>
        <a href="{{ url_for('lots') }}" class="{{ 'active' if request.endpoint in ('lots','lot_detail') else '' }}">Объявления</a>
        <a href="{{ url_for('files_page') }}" class="{{ 'active' if request.endpoint == 'files_page' else '' }}">Файлы</a>
      </div>
    </nav>

    {% with messages = get_flashed_messages(with_categories=true) %}
      {% if messages %}
        {% for cat, m in messages %}
          <div class="flash {{ cat }}">{{ m }}</div>
        {% endfor %}
      {% endif %}
    {% endwith %}

    {% block content %}{% endblock %}

    <footer>
      SZPT Monitor · Актобе · API ows.goszakup.gov.kz · эталон Stat.gov.kz · письма не отправляются по сети
    </footer>
  </div>
</body>
</html>
"""

# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------

INDEX_HTML = LAYOUT.replace("{% block content %}{% endblock %}", r"""
{% block content %}
<div class="pills">
  <div class="pill {{ 'ok' if has_token else 'warn' }}">
    Токен OWS: <strong>{{ 'подключён' if has_token else 'нет — укажите в .env' }}</strong>
  </div>
  <div class="pill {{ 'ok' if etalon_count else 'warn' }}">
    Эталон: <strong>{{ etalon_count }} файл(ов)</strong>
  </div>
  <div class="pill {{ 'ok' if results_count else 'warn' }}">
    В базе: <strong>{{ results_count }} объявлений</strong>
    {% if updated_at %}· {{ updated_at }}{% endif %}
  </div>
</div>

<div class="grid grid-2">
  <div class="card">
    <h2><span class="num">1</span> Сканирование завышений</h2>
    <p style="color:var(--muted);font-size:.85rem;margin-bottom:.5rem;">
      Запрос лотов через API, сравнение с эталоном Актобе, сохранение в Excel и в базу сайта.
    </p>
    <form method="post" action="{{ url_for('scan') }}">
      <div class="row">
        <div>
          <label>Порог отклонения, %</label>
          <input type="number" name="threshold" value="{{ th_mon }}" step="0.1" min="0">
        </div>
        <div>
          <label>Режим</label>
          <select name="mode">
            <option value="full" selected>Полный (нормальный, дольше)</option>
            <option value="fast">Быстрый (может пропустить лоты)</option>
          </select>
        </div>
      </div>
      <div class="row">
        <div>
          <label>Дата с</label>
          <input type="date" name="date_from" value="{{ date_from_default }}">
        </div>
        <div>
          <label>Дата по</label>
          <input type="date" name="date_to" value="{{ date_to_default }}">
        </div>
      </div>
      <div class="row">
        <div>
          <label>Макс. страниц API (0 = все, без лимита)</label>
          <input type="number" name="max_pages" value="0" min="0" max="5000">
        </div>
      </div>
      <p style="color:var(--muted);font-size:.8rem;margin:.5rem 0 0;">
        Полный режим + 0 страниц = нормальная проверка (может занять 10–40 мин).
      </p>
      <div id="progress-box" style="display:none;margin-top:.75rem;">
        <div style="background:#0f172a;border-radius:8px;height:10px;overflow:hidden;border:1px solid var(--border);">
          <div id="progress-bar" style="height:100%;width:0%;background:linear-gradient(90deg,#3b82f6,#10b981);transition:width .3s;"></div>
        </div>
        <p id="progress-text" style="color:var(--muted);font-size:.8rem;margin-top:.4rem;">Ожидание…</p>
      </div>
      <button type="submit" id="btn-scan">▶ Запустить скан</button>
      <script>
        (function(){
          var form = document.querySelector('form[action*="scan"]');
          if(!form) return;
          form.addEventListener('submit', function(){
            var box = document.getElementById('progress-box');
            var btn = document.getElementById('btn-scan');
            if(box) box.style.display = 'block';
            if(btn){ btn.disabled = true; btn.textContent = 'Сканирование…'; }
            setInterval(function(){
              fetch('/api/progress').then(function(r){ return r.json(); }).then(function(d){
                var bar = document.getElementById('progress-bar');
                var txt = document.getElementById('progress-text');
                var pct = d.total ? Math.min(100, Math.round(100*d.current/d.total)) : (d.current ? 40 : 8);
                if(bar) bar.style.width = pct + '%';
                if(txt) txt.textContent = (d.stage||'') + ' — ' + (d.current||0) + (d.total ? '/'+d.total : '');
              }).catch(function(){});
            }, 1500);
          });
        })();
      </script>
    </form>
  </div>

  <div class="card">
    <h2><span class="num">2</span> Письма ДЭР</h2>
    <p style="color:var(--muted);font-size:.85rem;margin-bottom:.5rem;">
      Официальные письма RU/KZ в УЗ и УО + таблица-приложение (≥15%).
    </p>
    <form method="post" action="{{ url_for('letters') }}">
      <div class="row">
        <div>
          <label>Порог для писем, %</label>
          <input type="number" name="threshold" value="{{ th_let }}" step="0.1" min="0">
        </div>
        <div>
          <label>Адресаты</label>
          <select name="to">
            <option value="uz,uo">УЗ + УО</option>
            <option value="uz">только УЗ</option>
            <option value="uo">только УО</option>
          </select>
        </div>
      </div>
      <label>Период, дней</label>
      <input type="number" name="days" value="365" min="1" max="730">
      <label>Адресат УЗ (ФИО, можно пусто)</label>
      <input type="text" name="addr_uz" value="" placeholder="например: Султангереев Е.Б.">
      <label>Адресат УО (ФИО, можно пусто)</label>
      <input type="text" name="addr_uo" value="" placeholder="например: Кузембаева Б.Б.">
      <button type="submit" class="btn-ok">✉ Сформировать письма</button>
    </form>
  </div>
</div>

{% if results_count %}
<div class="card" style="margin-top:1.15rem;">
  <h2>Последние результаты</h2>
  <div class="stat-grid">
    <div class="stat danger">
      <div class="val">{{ results_count }}</div>
      <div class="lbl">завышений</div>
    </div>
    <div class="stat warn">
      <div class="val">{{ max_pct }}%</div>
      <div class="lbl">макс. отклонение</div>
    </div>
    <div class="stat">
      <div class="val">{{ products_n }}</div>
      <div class="lbl">товаров СЗПТ</div>
    </div>
  </div>
  <a class="btn btn-sm" href="{{ url_for('lots') }}" style="width:auto;">Смотреть все объявления →</a>
</div>
{% endif %}
{% endblock %}
""")

LOTS_HTML = LAYOUT.replace("{% block content %}{% endblock %}", r"""
{% block content %}
<div class="page-title">Объявления с завышением</div>
<div class="page-sub">
  Данные последнего скана{% if updated_at %}: {{ updated_at }}{% endif %}.
  Клик по строке — подробности.
</div>

<div class="stat-grid">
  <div class="stat danger">
    <div class="val">{{ total }}</div>
    <div class="lbl">всего записей</div>
  </div>
  <div class="stat warn">
    <div class="val">{{ shown }}</div>
    <div class="lbl">после фильтра</div>
  </div>
  <div class="stat">
    <div class="val">{{ avg_pct }}%</div>
    <div class="lbl">среднее %</div>
  </div>
</div>

<form class="filters card" method="get" action="{{ url_for('lots') }}" style="padding:1rem 1.25rem;">
  <div class="field search-box">
    <label>Поиск</label>
    <input type="text" name="q" value="{{ q }}" placeholder="товар, заказчик, лот, БИН...">
  </div>
  <div class="field">
    <label>Товар</label>
    <select name="product">
      <option value="">Все</option>
      {% for p in products %}
        <option value="{{ p }}" {{ 'selected' if p == product else '' }}>{{ p }}</option>
      {% endfor %}
    </select>
  </div>
  <div class="field">
    <label>Мин. %</label>
    <input type="number" name="min_pct" value="{{ min_pct }}" step="0.1" min="0">
  </div>
  <div class="field">
    <label>Сортировка</label>
    <select name="sort">
      <option value="percent" {{ 'selected' if sort == 'percent' else '' }}>% убыв.</option>
      <option value="amount" {{ 'selected' if sort == 'amount' else '' }}>Сумма убыв.</option>
      <option value="date" {{ 'selected' if sort == 'date' else '' }}>Дата</option>
      <option value="product" {{ 'selected' if sort == 'product' else '' }}>Товар</option>
    </select>
  </div>
  <button type="submit">Фильтр</button>
</form>

{% if rows %}
<div class="table-wrap">
  <table>
    <thead>
      <tr>
        <th>%</th>
        <th>Товар</th>
        <th>Предмет лота</th>
        <th>Цена / эталон</th>
        <th>Сумма / разница</th>
        <th>Заказчик</th>
        <th>Поставщик</th>
        <th>Лот</th>
        <th></th>
      </tr>
    </thead>
    <tbody>
      {% for r in rows %}
      <tr>
        <td>
          {% set pct = r.percent|float %}
          <span class="pct {{ 'high' if pct >= 50 else ('mid' if pct >= 25 else 'low') }}">
            +{{ '%.1f'|format(pct) }}%
          </span>
        </td>
        <td><strong>{{ r.product or '—' }}</strong></td>
        <td class="lot-name">{{ r.lot_name or '—' }}</td>
        <td>
          <div>{{ r.unit_price if r.unit_price is not none else '—' }} ₸</div>
          <div class="mono">эт. {{ r.etalon if r.etalon is not none else '—' }} ₸ {% if r.etalon_period %}({{ r.etalon_period }}){% endif %}</div>
        </td>
        <td>
          <div>{{ r.amount if r.amount is not none else '—' }}</div>
          <div class="mono">разн. {{ r.difference if r.difference is not none else '—' }}</div>
        </td>
        <td class="customer">
          <div>{{ r.customer or '—' }}</div>
          <div class="mono">{{ r.customer_bin or '' }}</div>
        </td>
        <td class="customer">
          <div>{{ r.supplier_name or '—' }}</div>
          <div class="mono">{{ r.supplier_biin or '' }}</div>
        </td>
        <td class="mono">
          {{ r.lot_number or r.number_anno or r.lot_id or '—' }}
          {% if r.lot_url %}<div><a href="{{ r.lot_url }}" target="_blank" rel="noopener" style="color:var(--accent2);font-size:.75rem;">открыть</a></div>{% endif %}
        </td>
        <td>
          <a class="btn btn-sm btn-ghost" href="{{ url_for('lot_detail', idx=r._idx) }}">Подробнее</a>
        </td>
      </tr>
      {% endfor %}
    </tbody>
  </table>
</div>
{% else %}
<div class="card">
  <p class="empty">
    {% if total == 0 %}
      Нет данных. Сначала запустите <a href="{{ url_for('index') }}" style="color:var(--accent2)">сканирование</a>.
    {% else %}
      По выбранным фильтрам ничего не найдено.
    {% endif %}
  </p>
</div>
{% endif %}
{% endblock %}
""")

# Fix Chinese "详情" -> Russian
LOTS_HTML = LOTS_HTML.replace(">详情<", ">Подробнее<")

DETAIL_HTML = LAYOUT.replace("{% block content %}{% endblock %}", r"""
{% block content %}
<div class="page-title">Карточка объявления</div>
<div class="page-sub"><a href="{{ url_for('lots') }}" style="color:var(--accent2);text-decoration:none;">← Назад к списку</a></div>

<div class="card">
  <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:1rem;flex-wrap:wrap;">
    <div>
      <h2 style="margin-bottom:.35rem;">{{ row.product or 'Товар' }}</h2>
      <p style="color:var(--muted);max-width:600px;">{{ row.lot_name or '—' }}</p>
    </div>
    {% set pct = (row.percent or 0)|float %}
    <span class="pct {{ 'high' if pct >= 50 else ('mid' if pct >= 25 else 'low') }}" style="font-size:1.1rem;padding:.4rem .8rem;">
      +{{ '%.1f'|format(pct) }}%
    </span>
  </div>

  <div class="detail-panel">
    <h3>Цены</h3>
    <dl class="detail-grid">
      <dt>Цена за единицу</dt><dd><strong>{{ row.unit_price }} ₸</strong></dd>
      <dt>Эталон Актобе</dt><dd>{{ row.etalon }} ₸</dd>
      <dt>Отклонение</dt><dd>+{{ row.percent }}%</dd>
      <dt>Сумма по договору</dt><dd>{{ row.amount }} ₸</dd>
      <dt>Сумма по цене СЗПТ</dt><dd>{{ row.amount_szpt or '—' }} ₸</dd>
      <dt>Разница</dt><dd><strong>{{ row.difference if row.difference is not none else '—' }} ₸</strong></dd>
      <dt>Количество</dt><dd>{{ row.qty }}</dd>
      <dt>Период эталона</dt><dd>{{ row.etalon_period or '—' }}</dd>
    </dl>
  </div>

  <div class="detail-panel">
    <h3>Заказчик и поставщик</h3>
    <dl class="detail-grid">
      <dt>Заказчик</dt><dd>{{ row.customer or '—' }}</dd>
      <dt>БИН заказчика</dt><dd class="mono">{{ row.customer_bin or '—' }}</dd>
      <dt>Поставщик</dt><dd>{{ row.supplier_name or '—' }}</dd>
      <dt>БИН поставщика</dt><dd class="mono">{{ row.supplier_biin or '—' }}</dd>
      <dt>Номер лота</dt><dd class="mono">{{ row.lot_number or '—' }}</dd>
      <dt>Номер объявления</dt><dd class="mono">{{ row.number_anno or '—' }}</dd>
      <dt>Номер договора</dt><dd class="mono">{{ row.contract_number or '—' }}</dd>
      <dt>ID лота</dt><dd class="mono">{{ row.lot_id or '—' }}</dd>
      <dt>Статус</dt><dd>{{ row.status or '—' }}</dd>
      <dt>Дата</dt><dd>{{ row.date or '—' }}</dd>
      <dt>Год</dt><dd>{{ row.year or '—' }}</dd>
    </dl>
  </div>

  {% if row.lot_url or row.number_anno %}
  <div style="margin-top:1.25rem;">
    <a class="btn btn-sm" style="width:auto;margin-right:.5rem;"
       href="{{ row.lot_url or ('https://goszakup.gov.kz/ru/announce/index/' ~ row.number_anno) }}"
       target="_blank" rel="noopener">
      Открыть лот на goszakup.gov.kz
    </a>
  </div>
  {% endif %}
</div>
{% endblock %}
""")

FILES_HTML = LAYOUT.replace("{% block content %}{% endblock %}", r"""
{% block content %}
<div class="page-title">Файлы</div>
<div class="page-sub">Отчёты Excel и сгенерированные письма</div>

<div class="grid grid-2">
  <div class="card">
    <h2>output/ — Excel отчёты</h2>
    {% if out_files %}
      <ul class="file-list">
        {% for f in out_files %}
          <li>
            <span>{{ f }}</span>
            <a href="{{ url_for('download_output', filename=f) }}">Скачать</a>
          </li>
        {% endfor %}
      </ul>
    {% else %}
      <p class="empty">Пока нет файлов</p>
    {% endif %}
  </div>
  <div class="card">
    <h2>output/letters/ — письма</h2>
    {% if letter_files %}
      <ul class="file-list">
        {% for f in letter_files %}
          <li>
            <span>{{ f }}</span>
            <a href="{{ url_for('download_letters', filename=f) }}">Скачать</a>
          </li>
        {% endfor %}
      </ul>
    {% else %}
      <p class="empty">Пока нет писем</p>
    {% endif %}
  </div>
</div>
{% endblock %}
""")


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    rows = _load_results()
    meta = _results_meta()
    max_pct = 0.0
    products = set()
    for r in rows:
        try:
            max_pct = max(max_pct, float(r.get("percent") or 0))
        except (TypeError, ValueError):
            pass
        if r.get("product"):
            products.add(r["product"])

    from datetime import timedelta as _td
    _today = datetime.now().date()
    _from = _today - _td(days=90)
    return render_template_string(
        INDEX_HTML,
        has_token=bool(OWS_TOKEN),
        etalon_count=_etalon_count(),
        results_count=len(rows),
        updated_at=meta.get("updated_at"),
        th_mon=DEFAULT_THRESHOLD_MONITOR,
        th_let=DEFAULT_THRESHOLD_LETTER,
        max_pct=round(max_pct, 1),
        products_n=len(products),
        date_from_default=_from.isoformat(),
        date_to_default=_today.isoformat(),
    )


@app.route("/lots")
def lots():
    all_rows = _clean_rows(_load_results())
    meta = _results_meta()

    q = (request.args.get("q") or "").strip().lower()
    product = (request.args.get("product") or "").strip()
    min_pct = request.args.get("min_pct", type=float)
    if min_pct is None:
        min_pct = 0.0
    sort = request.args.get("sort") or "percent"

    products = sorted({str(r.get("product")) for r in all_rows if r.get("product")})

    filtered = []
    for i, r in enumerate(all_rows):
        r = _clean_row(dict(r))
        r["_idx"] = i
        try:
            pct = float(r.get("percent") or 0)
        except (TypeError, ValueError):
            pct = 0.0
        if pct < min_pct:
            continue
        if product and str(r.get("product") or "") != product:
            continue
        if q:
            blob = " ".join(str(r.get(k) or "") for k in (
                "product", "lot_name", "customer", "customer_bin",
                "supplier_name", "supplier_biin",
                "lot_number", "number_anno", "contract_number", "status",
            )).lower()
            if q not in blob:
                continue
        filtered.append(r)

    reverse = sort in ("percent", "amount")
    def sort_key(x):
        v = x.get(sort)
        try:
            return float(v)
        except (TypeError, ValueError):
            return str(v or "")
    filtered.sort(key=sort_key, reverse=reverse)

    avg = 0.0
    if filtered:
        s = 0.0
        n = 0
        for r in filtered:
            try:
                s += float(r.get("percent") or 0)
                n += 1
            except (TypeError, ValueError):
                pass
        avg = round(s / n, 1) if n else 0.0

    return render_template_string(
        LOTS_HTML,
        rows=filtered,
        total=len(all_rows),
        shown=len(filtered),
        avg_pct=avg,
        products=products,
        q=request.args.get("q") or "",
        product=product,
        min_pct=min_pct,
        sort=sort,
        updated_at=meta.get("updated_at"),
    )


@app.route("/lots/<int:idx>")
def lot_detail(idx: int):
    rows = _clean_rows(_load_results())
    if idx < 0 or idx >= len(rows):
        flash("Объявление не найдено", "err")
        return redirect(url_for("lots"))
    return render_template_string(DETAIL_HTML, row=_clean_row(rows[idx]))


@app.route("/files")
def files_page():
    return render_template_string(
        FILES_HTML,
        out_files=_list_files(OUTPUT_DIR, "*.xlsx"),
        letter_files=_list_files(LETTERS_DIR),
    )


@app.route("/scan", methods=["POST"])
def scan():
    th = request.form.get("threshold", DEFAULT_THRESHOLD_MONITOR)
    date_from = (request.form.get("date_from") or "").strip()
    date_to = (request.form.get("date_to") or "").strip()
    mode = request.form.get("mode") or "full"
    max_pages = request.form.get("max_pages") or "0"
    cmd = [
        sys.executable, str(BASE / "run_scan.py"),
        "--threshold", str(th),
        "--mode", str(mode),
        "--max-pages", str(max_pages),
    ]
    if date_from and date_to:
        cmd.extend(["--date-from", date_from, "--date-to", date_to])
    else:
        cmd.extend(["--days", "90"])
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=900, cwd=str(BASE),
            encoding="utf-8", errors="replace",
        )
        # После скана подтянуть Excel в JSON для сайта
        xlsx_files = sorted(OUTPUT_DIR.glob("zavysheniya_*.xlsx"), reverse=True)
        if not xlsx_files:
            xlsx_files = sorted(OUTPUT_DIR.glob("*.xlsx"), reverse=True)
        if xlsx_files:
            try:
                df = pd.read_excel(xlsx_files[0])
                colmap = {
                    "Год": "year", "Дата": "date", "Период эталона": "etalon_period",
                    "Товар": "product",
                    "Предмет лота": "lot_name", "Цена за ед.": "unit_price",
                    "Эталон Актобе": "etalon", "% отклонения": "percent",
                    "Сумма": "amount", "Сумма по договору": "amount",
                    "Сумма по цене СЗПТ": "amount_szpt", "Разница": "difference",
                    "Количество": "qty",
                    "Заказчик": "customer", "БИН заказчика": "customer_bin",
                    "Поставщик": "supplier_name", "БИН поставщика": "supplier_biin",
                    "Номер лота": "lot_number", "Номер объявления": "number_anno",
                    "Номер договора": "contract_number", "Статус": "status",
                    "ID лота": "lot_id",
                }
                df = df.rename(columns={k: v for k, v in colmap.items() if k in df.columns})
                df = df.where(pd.notnull(df), None)
                rows = _clean_rows(df.to_dict(orient="records"))
                _save_results(rows, meta={"threshold": th, "days": days, "file": xlsx_files[0].name})
            except Exception as e:
                flash(f"Скан выполнен, но не удалось загрузить таблицу: {e}", "err")

        if result.returncode == 0:
            n = len(_load_results())
            flash(f"Скан завершён. Найдено записей: {n}. Смотрите раздел «Объявления».", "ok")
        else:
            msg = (result.stderr or result.stdout or "Неизвестная ошибка").strip()
            flash(f"Ошибка скана:\n{msg[:1500]}", "err")
    except subprocess.TimeoutExpired:
        flash("Скан превысил время ожидания (15 мин).", "err")
    except Exception as e:
        flash(f"Исключение: {e}", "err")
    return redirect(url_for("index"))


@app.route("/letters", methods=["POST"])
def letters():
    th = request.form.get("threshold", DEFAULT_THRESHOLD_LETTER)
    to = request.form.get("to", "uz,uo")
    days = request.form.get("days", 365)
    addr_uz = (request.form.get("addr_uz") or "").strip()
    addr_uo = (request.form.get("addr_uo") or "").strip()
    cmd = [
        sys.executable, str(BASE / "generate_letters.py"),
        "--threshold", str(th),
        "--to", to,
        "--days", str(days),
    ]
    if addr_uz:
        cmd.extend(["--addr-uz-ru", addr_uz, "--addr-uz-kz", addr_uz])
    if addr_uo:
        cmd.extend(["--addr-uo-ru", addr_uo, "--addr-uo-kz", addr_uo])
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=900, cwd=str(BASE),
            encoding="utf-8", errors="replace",
        )
        if result.returncode == 0:
            flash("Письма и приложение сформированы. Раздел «Файлы».", "ok")
        else:
            msg = (result.stderr or result.stdout or "Неизвестная ошибка").strip()
            flash(f"Ошибка генерации:\n{msg[:1500]}", "err")
    except subprocess.TimeoutExpired:
        flash("Генерация превысила время ожидания (15 мин).", "err")
    except Exception as e:
        flash(f"Исключение: {e}", "err")
    return redirect(url_for("index"))


@app.route("/download/output/<path:filename>")
def download_output(filename):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


@app.route("/download/letters/<path:filename>")
def download_letters(filename):
    return send_from_directory(LETTERS_DIR, filename, as_attachment=True)


@app.route("/api/results")
def api_results():
    """JSON API для всех результатов (для отладки / интеграций)."""
    return jsonify(_results_meta() | {"rows": _load_results()})


@app.route("/api/progress")
def api_progress():
    path = OUTPUT_DIR / "scan_progress.json"
    if path.exists():
        try:
            with open(path, encoding="utf-8") as f:
                return jsonify(json.load(f))
        except Exception:
            pass
    return jsonify({"stage": "ожидание", "current": 0, "total": 0, "extra": ""})


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    print("=" * 50)
    print("  SZPT Monitor — http://127.0.0.1:5000")
    print("=" * 50)
    app.run(debug=True, host="0.0.0.0", port=5000)

"""
Naver Sellout Dashboard Generator
엑셀 파일을 읽어서 index.html 대시보드를 자동 생성합니다.
"""
import pandas as pd
import numpy as np
import json
import sys
from pathlib import Path
from datetime import datetime

# ── 파일 경로 ──────────────────────────────────────────────────
EXCEL_FILE = Path("data/sellout.xlsm")   # 업로드할 파일 경로
OUTPUT_FILE = Path("index.html")

def find_excel():
    """data/ 폴더에서 엑셀 파일 자동 탐색"""
    data_dir = Path("data")
    for ext in ["*.xlsm", "*.xlsx", "*.xls"]:
        files = list(data_dir.glob(ext))
        if files:
            return files[0]
    raise FileNotFoundError("data/ 폴더에 엑셀 파일이 없습니다.")

# ── 데이터 파싱 ────────────────────────────────────────────────
def parse_revenue(xl):
    """배송유형 데일리 시트에서 일별 매출 파싱"""
    df = xl.parse("배송유형 데일리", header=None)
    rev = df.iloc[5:, [0, 1, 6, 11]].copy()
    rev.columns = ["date", "philips", "sonicare", "total"]
    rev["date"] = pd.to_datetime(rev["date"], errors="coerce")
    rev = rev.dropna(subset=["date"])
    for c in ["philips", "sonicare", "total"]:
        rev[c] = pd.to_numeric(rev[c], errors="coerce").fillna(0)
    rev = rev[rev["total"] > 0].copy()
    return rev

def parse_weekly(xl):
    """Weekly vs 2025 시트에서 주간 YTD 데이터 파싱"""
    df = xl.parse("Weekly vs 2025", header=None)
    weekly = df.iloc[9:, [2, 3, 5, 6]].copy()
    weekly.columns = ["week", "sales_2025", "sales_2026", "csg"]
    weekly = weekly[weekly["week"].astype(str).str.match(r"W\d+")]
    weekly["sales_2025"] = pd.to_numeric(weekly["sales_2025"], errors="coerce")
    weekly["sales_2026"] = pd.to_numeric(weekly["sales_2026"], errors="coerce")
    weekly["csg"] = pd.to_numeric(weekly["csg"], errors="coerce")

    # YTD 누적 계산
    weekly["ytd_2025"] = weekly["sales_2025"].cumsum()
    weekly["ytd_2026"] = weekly["sales_2026"].cumsum()
    # 2026 데이터가 있는 주까지만
    weekly = weekly[weekly["sales_2026"].notna()]
    return weekly

def parse_monthly(xl):
    """Daily 시트에서 월별 매출 파싱"""
    df = xl.parse("Daily", header=None)
    monthly = df.iloc[5:10, [1, 2, 3]].copy()
    monthly.columns = ["month", "sales_2025", "sales_2026"]
    monthly["month"] = pd.to_numeric(monthly["month"], errors="coerce")
    monthly["sales_2025"] = pd.to_numeric(monthly["sales_2025"], errors="coerce")
    monthly["sales_2026"] = pd.to_numeric(monthly["sales_2026"], errors="coerce")
    monthly = monthly.dropna(subset=["month"])
    return monthly

def parse_daily_so(xl):
    """Daily SO 시트에서 SKU별 일별 수량 파싱"""
    df = xl.parse("Daily SO", header=None)
    df_filled = df.copy()
    df_filled.iloc[8:, 0] = df.iloc[8:, 0].ffill()
    df_filled.iloc[8:, 1] = df.iloc[8:, 1].ffill()

    # 카테고리 요약 행
    data = df_filled.iloc[8:]
    summary_mask = data.iloc[:, 0].isin(["MG 요약", "OHC", "BT", "총합계"]) & data.iloc[:, 4].isna()
    cat_summary = data[summary_mask].copy()

    # 2026년 현재 달 컬럼 찾기 (col 5~25 = 현재 월 1일~21일)
    month_row = df.iloc[6, :]
    day_row   = df.iloc[7, :]
    current_month_cols = []
    cur_month = None
    for col_idx in range(5, df.shape[1]):
        m = month_row.iloc[col_idx]
        d = day_row.iloc[col_idx]
        if not pd.isna(m) and m != "SKU":
            cur_month = int(float(m))
        if cur_month and not pd.isna(d) and d != "SKU":
            try:
                current_month_cols.append((col_idx, int(float(d))))
            except:
                pass
        if len(current_month_cols) >= 31:
            break

    # SKU 데이터 파싱
    sku_rows = data[data.iloc[:, 4].notna()].copy()
    sku_rows = sku_rows[~sku_rows.iloc[:, 4].astype(str).str.contains("요약|총합", na=True)]
    sku_rows["sku"] = sku_rows.iloc[:, 4].astype(str)
    sku_rows["cat"] = sku_rows.iloc[:, 0]

    # 전일(마지막 데이터 날) & MTD
    if current_month_cols:
        col_indices = [c[0] for c in current_month_cols]
        sku_rows["mtd"] = sku_rows.iloc[:, col_indices].apply(pd.to_numeric, errors="coerce").sum(axis=1)
        # 마지막 날 컬럼 (0이 아닌 가장 최근)
        last_col_idx = col_indices[-1]
        for ci in reversed(col_indices):
            col_vals = sku_rows.iloc[:, ci].apply(pd.to_numeric, errors="coerce")
            if col_vals.sum() > 0:
                last_col_idx = ci
                break
        sku_rows["last_day"] = sku_rows.iloc[:, last_col_idx].apply(pd.to_numeric, errors="coerce").fillna(0)
        last_day_num = day_row.iloc[last_col_idx]
    else:
        sku_rows["mtd"] = 0
        sku_rows["last_day"] = 0
        last_day_num = "?"

    return sku_rows, cat_summary, current_month_cols, last_day_num

# ── KPI 계산 ───────────────────────────────────────────────────
def calc_kpi(rev, monthly):
    latest = rev["date"].max()
    prev   = rev[rev["date"] < latest]["date"].max()

    today_rev = float(rev[rev["date"] == latest]["total"].values[0])
    yest_rev  = float(rev[rev["date"] == prev]["total"].values[0]) if prev else 0
    dod = (today_rev - yest_rev) / yest_rev if yest_rev else 0

    cur_m = latest.month
    cur_m_data = monthly[monthly["month"] == cur_m]
    prev_m_data = monthly[monthly["month"] == cur_m - 1] if cur_m > 1 else None

    mtd_2026 = float(cur_m_data["sales_2026"].values[0]) if len(cur_m_data) else 0
    mtd_2025 = float(cur_m_data["sales_2025"].values[0]) if len(cur_m_data) else 0
    prev_2026 = float(prev_m_data["sales_2026"].values[0]) if (prev_m_data is not None and len(prev_m_data)) else 0

    mom = (mtd_2026 - prev_2026) / prev_2026 if prev_2026 else 0
    yoy = (mtd_2026 - mtd_2025) / mtd_2025 if mtd_2025 else 0

    ytd_2026 = rev[rev["date"].dt.year == latest.year]["total"].sum()
    ytd_2025 = float(monthly["sales_2025"].sum())

    return {
        "latest_date": latest.strftime("%Y.%m.%d"),
        "latest_month": cur_m,
        "today_rev": today_rev,
        "yest_rev": yest_rev,
        "dod": dod,
        "mom": mom,
        "yoy": yoy,
        "mtd_2026": mtd_2026,
        "mtd_2025": mtd_2025,
        "ytd_2026": ytd_2026,
        "ytd_2025": ytd_2025,
    }

# ── HTML 생성 ──────────────────────────────────────────────────
def badge(val, fmt=".1%", invert=False):
    positive = val > 0
    if invert:
        positive = not positive
    cls = "pos" if positive else "neg"
    arrow = "▲" if positive else "▼"
    return f'<span class="kpi-badge {cls}">{arrow} {val:{fmt}}</span>'

def fmt_억(val):
    if abs(val) >= 1e8:
        return f"{val/1e8:.1f}억"
    return f"{val/1e4:.0f}만"

def generate_html(kpi, rev, weekly, sku_rows, cat_summary, month_cols, last_day, generated_at):
    # ── JS 데이터 준비 ──
    # 현재 달 일별 매출
    cur_year = pd.to_datetime(kpi["latest_date"]).year
    cur_m    = kpi["latest_month"]
    may_rev = rev[(rev["date"].dt.year == cur_year) & (rev["date"].dt.month == cur_m)].copy()
    may_rev = may_rev.sort_values("date")

    daily_labels   = [d.strftime("%-m/%-d") for d in may_rev["date"]]
    daily_philips  = [round(v/10000, 1) for v in may_rev["philips"]]
    daily_sonicare = [round(v/10000, 1) for v in may_rev["sonicare"]]

    # 카테고리 주간 (3주 블록)
    cat_names = {"MG 요약": "MG", "OHC": "OHC", "BT": "BT"}
    cat_weekly = {}
    if month_cols:
        blocks = [month_cols[:7], month_cols[7:14], month_cols[14:]]
        w_labels = [f"W1 ({month_cols[0][1]}일~{month_cols[min(6,len(month_cols)-1)][1]}일)",
                    f"W2 ({month_cols[7][1]}일~{month_cols[min(13,len(month_cols)-1)][1]}일)" if len(month_cols)>7 else "",
                    f"W3 ({month_cols[14][1]}일~{month_cols[-1][1]}일)" if len(month_cols)>14 else ""]
        w_labels = [l for l in w_labels if l]
        for raw_name, label in cat_names.items():
            row = cat_summary[cat_summary.iloc[:, 0] == raw_name]
            if len(row) == 0:
                cat_weekly[label] = [0] * len(w_labels)
                continue
            row = row.iloc[0]
            vals = []
            for blk in blocks[:len(w_labels)]:
                s = sum(pd.to_numeric(row.iloc[c], errors="coerce") or 0 for c, _ in blk)
                vals.append(int(s))
            cat_weekly[label] = vals
    else:
        w_labels = []
        for label in cat_names.values():
            cat_weekly[label] = []

    # 주간 YTD
    wk_labels   = weekly["week"].tolist()
    ytd2026_list = [round(v/1e8, 1) for v in weekly["ytd_2026"]]
    ytd2025_list = [round(v/1e8, 1) for v in weekly["ytd_2025"].fillna(0)]
    csg_list    = [round(v*100, 1) for v in weekly["csg"].fillna(0)]

    # SKU TOP 10
    top10_day = sku_rows.nlargest(10, "last_day")[["sku","cat","last_day","mtd"]].to_dict("records")
    top10_mtd = sku_rows.nlargest(10, "mtd")[["sku","cat","last_day","mtd"]].to_dict("records")

    j = lambda x: json.dumps(x, ensure_ascii=False)

    # KPI values
    k = kpi
    mom_badge  = badge(k["mom"])
    yoy_badge  = badge(k["yoy"])
    dod_badge  = badge(k["dod"])
    ytd_yoy    = (k["ytd_2026"] - k["ytd_2025"]) / k["ytd_2025"] if k["ytd_2025"] else 0

    def top10_rows_day(data):
        mx = max((d["last_day"] for d in data), default=1) or 1
        ranks = ["🥇","🥈","🥉"] + [str(i) for i in range(4,11)]
        rows = ""
        for i, d in enumerate(data):
            pct = d["last_day"] / mx * 100
            rows += f"""<tr>
              <td class="rank-num">{ranks[i]}</td>
              <td><span class="sku-code">{d['sku']}</span></td>
              <td><span class="cat-pill cat-{d['cat'].lower().replace(' ','')}">{d['cat']}</span></td>
              <td class="bar-cell">
                <div class="mini-bar-wrap">
                  <div class="mini-bar-bg"><div class="mini-bar-fill" style="width:{pct:.0f}%"></div></div>
                  <span class="qty-num">{int(d['last_day'])}</span>
                </div>
              </td>
              <td class="qty-num" style="color:var(--text-muted)">{int(d['mtd']):,}</td>
            </tr>"""
        return rows

    def top10_rows_mtd(data):
        mx = max((d["mtd"] for d in data), default=1) or 1
        ranks = ["🥇","🥈","🥉"] + [str(i) for i in range(4,11)]
        rows = ""
        for i, d in enumerate(data):
            pct = d["mtd"] / mx * 100
            rows += f"""<tr>
              <td class="rank-num">{ranks[i]}</td>
              <td><span class="sku-code">{d['sku']}</span></td>
              <td><span class="cat-pill cat-{d['cat'].lower().replace(' ','')}">{d['cat']}</span></td>
              <td class="bar-cell">
                <div class="mini-bar-wrap">
                  <div class="mini-bar-bg"><div class="mini-bar-fill" style="width:{pct:.0f}%"></div></div>
                  <span class="qty-num">{int(d['mtd']):,}</span>
                </div>
              </td>
              <td class="qty-num" style="color:var(--text-muted)">{int(d['last_day'])}</td>
            </tr>"""
        return rows

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>네이버 셀아웃 대시보드 | {k['latest_date']}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700;900&family=DM+Mono:wght@400;500&family=Bebas+Neue&display=swap');
  :root {{
    --bg:#0a0c10;--surface:#12151c;--surface2:#1a1f2b;--border:#232a38;
    --accent:#3b82f6;--accent2:#22d3ee;--accent3:#a78bfa;
    --positive:#10b981;--negative:#f43f5e;--neutral:#f59e0b;
    --text:#e2e8f0;--text-muted:#64748b;--text-dim:#94a3b8;
    --mg:#3b82f6;--ohc:#22d3ee;--bt:#a78bfa;
  }}
  *{{margin:0;padding:0;box-sizing:border-box}}
  body{{font-family:'Noto Sans KR',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden}}
  header{{padding:20px 32px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;background:linear-gradient(135deg,#0d1117,#0a0c10);position:sticky;top:0;z-index:100}}
  .brand{{display:flex;align-items:center;gap:14px}}
  .brand-icon{{width:38px;height:38px;background:linear-gradient(135deg,var(--accent),var(--accent2));border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:18px;box-shadow:0 0 20px rgba(59,130,246,.35)}}
  .brand-text h1{{font-family:'Bebas Neue',sans-serif;font-size:22px;letter-spacing:2px;line-height:1}}
  .brand-text p{{font-size:11px;color:var(--text-muted);letter-spacing:.5px;font-weight:300}}
  .header-date .date-main{{font-family:'DM Mono',monospace;font-size:16px;color:var(--accent2);font-weight:500}}
  .header-date .date-sub{{font-size:11px;color:var(--text-muted);margin-top:2px;text-align:right}}
  main{{padding:24px 32px;max-width:1600px;margin:0 auto}}
  .section-label{{font-size:10px;letter-spacing:3px;text-transform:uppercase;color:var(--text-muted);margin-bottom:14px;display:flex;align-items:center;gap:8px}}
  .section-label::before{{content:'';display:block;width:3px;height:12px;border-radius:2px;background:var(--accent)}}
  .kpi-grid{{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin-bottom:28px}}
  .kpi-card{{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:20px;position:relative;overflow:hidden;transition:border-color .2s,transform .2s}}
  .kpi-card:hover{{border-color:var(--accent);transform:translateY(-2px)}}
  .kpi-card::before{{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,var(--accent),var(--accent2));opacity:0;transition:opacity .2s}}
  .kpi-card:hover::before,.kpi-card.highlight::before{{opacity:1}}
  .kpi-card.highlight{{border-color:rgba(59,130,246,.4);background:linear-gradient(135deg,#12151c,#151c2a)}}
  .kpi-label{{font-size:11px;color:var(--text-muted);margin-bottom:8px}}
  .kpi-value{{font-family:'DM Mono',monospace;font-size:22px;font-weight:500;line-height:1.1;margin-bottom:6px}}
  .kpi-value .unit{{font-size:13px;color:var(--text-muted);font-weight:400}}
  .kpi-badge{{display:inline-flex;align-items:center;gap:4px;font-size:12px;font-weight:600;padding:3px 8px;border-radius:6px;font-family:'DM Mono',monospace}}
  .kpi-badge.pos{{background:rgba(16,185,129,.15);color:var(--positive)}}
  .kpi-badge.neg{{background:rgba(244,63,94,.15);color:var(--negative)}}
  .kpi-sub{{font-size:10px;color:var(--text-muted);margin-top:4px}}
  .charts-row{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:28px}}
  .chart-card,.weekly-card{{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:22px}}
  .weekly-card{{margin-bottom:28px}}
  .chart-header{{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:20px}}
  .chart-title{{font-size:14px;font-weight:700;margin-bottom:3px}}
  .chart-subtitle{{font-size:11px;color:var(--text-muted)}}
  .legend-row{{display:flex;gap:14px;align-items:center;flex-wrap:wrap}}
  .legend-item{{display:flex;align-items:center;gap:5px;font-size:11px;color:var(--text-dim)}}
  .legend-dot{{width:8px;height:8px;border-radius:50%;flex-shrink:0}}
  .chart-wrap{{position:relative;height:240px}}
  .weekly-wrap{{position:relative;height:280px}}
  .tables-row{{display:grid;grid-template-columns:1fr 1fr;gap:16px}}
  .table-card{{background:var(--surface);border:1px solid var(--border);border-radius:14px;overflow:hidden}}
  .table-header{{padding:18px 20px 14px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;background:var(--surface2)}}
  .table-title{{font-size:14px;font-weight:700}}
  .table-date{{font-size:11px;color:var(--text-muted);font-family:'DM Mono',monospace}}
  table{{width:100%;border-collapse:collapse}}
  thead th{{padding:10px 16px;text-align:left;font-size:10px;letter-spacing:1px;text-transform:uppercase;color:var(--text-muted);background:var(--surface2);border-bottom:1px solid var(--border);font-weight:500}}
  thead th:last-child,tbody td:last-child{{text-align:right}}
  tbody tr{{border-bottom:1px solid rgba(35,42,56,.6);transition:background .15s}}
  tbody tr:hover{{background:rgba(59,130,246,.05)}}
  tbody tr:last-child{{border-bottom:none}}
  tbody td{{padding:11px 16px;font-size:13px;color:var(--text-dim)}}
  .rank-num{{font-family:'DM Mono',monospace;font-size:12px;color:var(--text-muted);width:24px;text-align:center}}
  .sku-code{{font-family:'DM Mono',monospace;font-size:13px;color:var(--text);font-weight:500}}
  .bar-cell{{min-width:80px}}
  .mini-bar-wrap{{display:flex;align-items:center;gap:8px}}
  .mini-bar-bg{{flex:1;height:4px;background:var(--border);border-radius:2px;overflow:hidden}}
  .mini-bar-fill{{height:100%;border-radius:2px;background:linear-gradient(90deg,var(--accent),var(--accent2))}}
  .qty-num{{font-family:'DM Mono',monospace;font-size:13px;color:var(--text);text-align:right;font-weight:500;min-width:36px}}
  .cat-pill{{display:inline-block;font-size:10px;padding:2px 6px;border-radius:4px;font-weight:600;letter-spacing:.5px}}
  .cat-mg{{background:rgba(59,130,246,.15);color:var(--mg)}}
  .cat-ohc{{background:rgba(34,211,238,.15);color:var(--ohc)}}
  .cat-bt{{background:rgba(167,139,250,.15);color:var(--bt)}}
  .live-dot{{width:8px;height:8px;border-radius:50%;background:var(--positive);box-shadow:0 0 8px var(--positive);animation:pulse 2s infinite;display:inline-block;margin-right:6px}}
  @keyframes pulse{{0%,100%{{opacity:1}}50%{{opacity:.4}}}}
  @keyframes fadeUp{{from{{opacity:0;transform:translateY(16px)}}to{{opacity:1;transform:translateY(0)}}}}
  .kpi-card{{animation:fadeUp .4s ease both}}
  .kpi-card:nth-child(1){{animation-delay:.05s}}.kpi-card:nth-child(2){{animation-delay:.1s}}
  .kpi-card:nth-child(3){{animation-delay:.15s}}.kpi-card:nth-child(4){{animation-delay:.2s}}
  .kpi-card:nth-child(5){{animation-delay:.25s}}
  .chart-card,.weekly-card,.table-card{{animation:fadeUp .5s ease both;animation-delay:.3s}}
  .update-stamp{{font-size:10px;color:var(--text-muted);text-align:center;padding:20px;opacity:.5}}
</style>
</head>
<body>
<header>
  <div class="brand">
    <div class="brand-icon">📊</div>
    <div class="brand-text">
      <h1>NAVER SELLOUT</h1>
      <p><span class="live-dot"></span>Daily Performance Dashboard · {cur_year}</p>
    </div>
  </div>
  <div class="header-date">
    <div class="date-main">{k['latest_date']}</div>
    <div class="date-sub">자동 생성: {generated_at}</div>
  </div>
</header>
<main>
  <div class="section-label">섹션 01 &nbsp;|&nbsp; 핵심 지표</div>
  <div class="kpi-grid">
    <div class="kpi-card highlight">
      <div class="kpi-label">전일 매출 ({k['latest_date']})</div>
      <div class="kpi-value">{k['today_rev']/1e4:,.0f}<span class="unit">만원</span></div>
      {dod_badge}
      <div class="kpi-sub">vs 전전일 {k['yest_rev']/1e4:,.0f}만원</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">MOM (월 대비)</div>
      <div class="kpi-value">{k['mom']*100:+.1f}<span class="unit">%</span></div>
      {mom_badge}
      <div class="kpi-sub">전월 {k['mtd_2026']/1e8:.1f}억 → 이번달 {k['mtd_2026']/1e8:.1f}억</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">YOY (연 대비)</div>
      <div class="kpi-value">{k['yoy']*100:+.1f}<span class="unit">%</span></div>
      {yoy_badge}
      <div class="kpi-sub">25년 {k['mtd_2025']/1e8:.1f}억 → 26년 {k['mtd_2026']/1e8:.1f}억</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">MTD ({cur_m}월 누적)</div>
      <div class="kpi-value">{k['mtd_2026']/1e8:.1f}<span class="unit">억</span></div>
      {badge(k['yoy'])}
      <div class="kpi-sub">2025년 동월 {k['mtd_2025']/1e8:.1f}억</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">YTD (연간 누적)</div>
      <div class="kpi-value">{k['ytd_2026']/1e8:.1f}<span class="unit">억</span></div>
      {badge(ytd_yoy)}
      <div class="kpi-sub">2025년 {k['ytd_2025']/1e8:.1f}억</div>
    </div>
  </div>

  <div class="section-label">섹션 02 &nbsp;|&nbsp; 일별 트렌드 & 카테고리 믹스</div>
  <div class="charts-row">
    <div class="chart-card">
      <div class="chart-header">
        <div>
          <div class="chart-title">일별 매출 트렌드 ({cur_m}월)</div>
          <div class="chart-subtitle">단위: 만원</div>
        </div>
        <div class="legend-row">
          <div class="legend-item"><div class="legend-dot" style="background:#3b82f6"></div>Philips</div>
          <div class="legend-item"><div class="legend-dot" style="background:#22d3ee"></div>Sonicare</div>
        </div>
      </div>
      <div class="chart-wrap"><canvas id="dailyChart"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-header">
        <div>
          <div class="chart-title">카테고리별 수량 믹스 ({cur_m}월)</div>
          <div class="chart-subtitle">MG / OHC / BT 주간 판매 수량</div>
        </div>
        <div class="legend-row">
          <div class="legend-item"><div class="legend-dot" style="background:var(--mg)"></div>MG</div>
          <div class="legend-item"><div class="legend-dot" style="background:var(--ohc)"></div>OHC</div>
          <div class="legend-item"><div class="legend-dot" style="background:var(--bt)"></div>BT</div>
        </div>
      </div>
      <div class="chart-wrap"><canvas id="catChart"></canvas></div>
    </div>
  </div>

  <div class="section-label">섹션 03 &nbsp;|&nbsp; Weekly Sell-Out 트렌드 (2026 vs 2025 누적)</div>
  <div class="weekly-card">
    <div class="chart-header">
      <div>
        <div class="chart-title">주간 셀아웃 누적 비교</div>
        <div class="chart-subtitle">2025년 vs 2026년 YTD 누적 매출 · 단위: 억원</div>
      </div>
      <div class="legend-row">
        <div class="legend-item"><div class="legend-dot" style="background:#3b82f6"></div>2026 누적</div>
        <div class="legend-item"><div class="legend-dot" style="background:#64748b"></div>2025 누적</div>
        <div class="legend-item"><div class="legend-dot" style="background:var(--positive)"></div>YoY CSG%</div>
      </div>
    </div>
    <div class="weekly-wrap"><canvas id="weeklyChart"></canvas></div>
  </div>

  <div class="section-label">섹션 04 &nbsp;|&nbsp; SKU 성과 순위</div>
  <div class="tables-row">
    <div class="table-card">
      <div class="table-header">
        <div class="table-title">🏆 전일 TOP 10 SKU</div>
        <div class="table-date">기준: {k['latest_date']} (수량)</div>
      </div>
      <table>
        <thead><tr><th style="width:36px">#</th><th>SKU</th><th>Cat</th><th>수량</th><th>MTD</th></tr></thead>
        <tbody>{top10_rows_day(top10_day)}</tbody>
      </table>
    </div>
    <div class="table-card">
      <div class="table-header">
        <div class="table-title">📅 월 누적 TOP 10 SKU (MTD)</div>
        <div class="table-date">기준: {cur_m}월 누적</div>
      </div>
      <table>
        <thead><tr><th style="width:36px">#</th><th>SKU</th><th>Cat</th><th>MTD 수량</th><th>전일</th></tr></thead>
        <tbody>{top10_rows_mtd(top10_mtd)}</tbody>
      </table>
    </div>
  </div>
  <div class="update-stamp">자동 생성: {generated_at} · GitHub Actions · parse_dashboard.py</div>
</main>
<script>
const gridColor='rgba(35,42,56,0.7)',tickColor='#64748b';
const fontBase={{family:"'DM Mono',monospace",size:11}};
Chart.defaults.color=tickColor; Chart.defaults.font=fontBase;

// 1) 일별 매출
const ctx1=document.getElementById('dailyChart').getContext('2d');
const g1=ctx1.createLinearGradient(0,0,0,240);
g1.addColorStop(0,'rgba(59,130,246,0.3)'); g1.addColorStop(1,'rgba(59,130,246,0)');
const g2=ctx1.createLinearGradient(0,0,0,240);
g2.addColorStop(0,'rgba(34,211,238,0.2)'); g2.addColorStop(1,'rgba(34,211,238,0)');
new Chart(ctx1,{{type:'line',data:{{labels:{j(daily_labels)},datasets:[
  {{label:'Philips(만원)',data:{j(daily_philips)},borderColor:'#3b82f6',backgroundColor:g1,fill:true,tension:0.35,pointRadius:2,pointHoverRadius:5,borderWidth:2}},
  {{label:'Sonicare(만원)',data:{j(daily_sonicare)},borderColor:'#22d3ee',backgroundColor:g2,fill:true,tension:0.35,pointRadius:2,pointHoverRadius:5,borderWidth:2}}
]}},options:{{responsive:true,maintainAspectRatio:false,interaction:{{mode:'index',intersect:false}},
plugins:{{legend:{{display:false}},tooltip:{{backgroundColor:'#1a1f2b',borderColor:'#232a38',borderWidth:1,
callbacks:{{label:ctx=>` ${{ctx.dataset.label}}: ${{ctx.parsed.y.toLocaleString()}}만원`}}}}}},
scales:{{x:{{grid:{{color:gridColor}},ticks:{{maxTicksLimit:12}}}},y:{{grid:{{color:gridColor}},ticks:{{callback:v=>v>=1000?`${{(v/100).toFixed(0)}}억`:`${{v}}만`}}}}}}}}
}});

// 2) 카테고리 주간 바
new Chart(document.getElementById('catChart'),{{type:'bar',data:{{labels:{j(w_labels)},datasets:[
  {{label:'MG',data:{j(cat_weekly.get('MG',[]))},backgroundColor:'rgba(59,130,246,0.75)',borderRadius:4,borderSkipped:false}},
  {{label:'OHC',data:{j(cat_weekly.get('OHC',[]))},backgroundColor:'rgba(34,211,238,0.75)',borderRadius:4,borderSkipped:false}},
  {{label:'BT',data:{j(cat_weekly.get('BT',[]))},backgroundColor:'rgba(167,139,250,0.75)',borderRadius:4,borderSkipped:false}}
]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},
tooltip:{{backgroundColor:'#1a1f2b',borderColor:'#232a38',borderWidth:1,callbacks:{{label:ctx=>` ${{ctx.dataset.label}}: ${{ctx.parsed.y.toLocaleString()}}개`}}}}}},
scales:{{x:{{grid:{{display:false}}}},y:{{grid:{{color:gridColor}}}}}}
}});

// 3) 주간 YTD
const ctx3=document.getElementById('weeklyChart').getContext('2d');
const g3=ctx3.createLinearGradient(0,0,0,280);
g3.addColorStop(0,'rgba(59,130,246,0.25)'); g3.addColorStop(1,'rgba(59,130,246,0.01)');
const g4=ctx3.createLinearGradient(0,0,0,280);
g4.addColorStop(0,'rgba(100,116,139,0.2)'); g4.addColorStop(1,'rgba(100,116,139,0.01)');
new Chart(ctx3,{{type:'line',data:{{labels:{j(wk_labels)},datasets:[
  {{label:'2026 YTD(억원)',data:{j(ytd2026_list)},borderColor:'#3b82f6',backgroundColor:g3,fill:true,tension:0.3,pointRadius:3,pointHoverRadius:6,borderWidth:2.5,yAxisID:'y',order:2}},
  {{label:'2025 YTD(억원)',data:{j(ytd2025_list)},borderColor:'#64748b',backgroundColor:g4,fill:true,tension:0.3,pointRadius:2,pointHoverRadius:5,borderWidth:1.5,borderDash:[4,3],yAxisID:'y',order:3}},
  {{label:'YoY CSG%',data:{j(csg_list)},borderColor:'#10b981',backgroundColor:'transparent',tension:0.3,pointRadius:3,pointHoverRadius:6,borderWidth:2,yAxisID:'y1',order:1}}
]}},options:{{responsive:true,maintainAspectRatio:false,interaction:{{mode:'index',intersect:false}},
plugins:{{legend:{{display:false}},tooltip:{{backgroundColor:'#1a1f2b',borderColor:'#232a38',borderWidth:1,
callbacks:{{label:ctx=>ctx.datasetIndex===2?` CSG: ${{ctx.parsed.y}}%`:` ${{ctx.dataset.label}}: ${{ctx.parsed.y.toFixed(1)}}억원`}}}}}},
scales:{{x:{{grid:{{color:gridColor}}}},y:{{grid:{{color:gridColor}},ticks:{{callback:v=>`${{v.toFixed(0)}}억`}},position:'left'}},
y1:{{grid:{{display:false}},ticks:{{callback:v=>`${{v}}%`}},position:'right'}}}}
}});
</script>
</body>
</html>"""
    return html

# ── MAIN ───────────────────────────────────────────────────────
def main():
    print("📂 엑셀 파일 탐색 중...")
    excel_path = find_excel()
    print(f"✅ 파일 발견: {excel_path}")

    print("📊 데이터 파싱 중...")
    xl = pd.ExcelFile(excel_path, engine="openpyxl")

    rev        = parse_revenue(xl)
    weekly     = parse_weekly(xl)
    monthly    = parse_monthly(xl)
    sku_rows, cat_summary, month_cols, last_day = parse_daily_so(xl)
    kpi        = calc_kpi(rev, monthly)

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M KST")
    print(f"📅 기준일: {kpi['latest_date']}")

    print("🎨 HTML 생성 중...")
    html = generate_html(kpi, rev, weekly, sku_rows, cat_summary, month_cols, last_day, generated_at)

    OUTPUT_FILE.write_text(html, encoding="utf-8")
    print(f"✅ 완료: {OUTPUT_FILE} 생성됨")

if __name__ == "__main__":
    main()

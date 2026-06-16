from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parent
OUTPUT_FILE = ROOT / "index.html"
PUBLIC_DIR = ROOT / "public"
PUBLIC_OUTPUT_FILE = PUBLIC_DIR / "index.html"
SOURCE_NAME_MARKER = "2025-2026.Q1"
RMB_PER_YI = 100_000_000
TOP_N = 15


def find_source_workbook() -> Path:
    exact = Path.home() / "Desktop" / "海关数据" / "海关出口数据2025-2026.Q1.xlsx"
    if exact.exists():
        return exact

    desktop = Path.home() / "Desktop"
    candidates = [
        path
        for path in desktop.rglob("*.xlsx")
        if SOURCE_NAME_MARKER in path.name and not path.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError(
            f"Could not find an .xlsx file containing {SOURCE_NAME_MARKER!r} under {desktop}"
        )
    return sorted(candidates, key=lambda item: item.stat().st_mtime, reverse=True)[0]


def to_yi(value: float | int) -> float:
    return round(float(value) / RMB_PER_YI, 4)


def period_sort_key(label: str) -> tuple[int, int]:
    if "Q" in label:
        year, quarter = label.split("Q", 1)
        return int(year), int(quarter) * 3
    year, month = label.split("-", 1)
    return int(year), int(month)


def delta(current: int, prior: int | None) -> int | None:
    if prior is None:
        return None
    return int(current) - int(prior)


def ranked_rows(df: pd.DataFrame, period_col: str, period: str, *, top_n: int = TOP_N) -> list[dict]:
    period_df = df[df[period_col] == period].copy()
    ranked = (
        period_df.groupby("贸易伙伴名称", as_index=False)["人民币"]
        .sum()
        .sort_values("人民币", ascending=False)
        .head(top_n)
    )
    return [
        {"country": row["贸易伙伴名称"], "amountYi": to_yi(row["人民币"])}
        for _, row in ranked.iterrows()
    ]


def tidy_series(df: pd.DataFrame, period_col: str, periods: list[str], countries: list[str]) -> dict[str, list[float]]:
    grouped = (
        df[df["贸易伙伴名称"].isin(countries)]
        .groupby(["贸易伙伴名称", period_col], as_index=False)["人民币"]
        .sum()
    )
    pivot = grouped.pivot(index="贸易伙伴名称", columns=period_col, values="人民币").fillna(0)
    series: dict[str, list[float]] = {}
    for country in countries:
        if country not in pivot.index:
            series[country] = [0 for _ in periods]
            continue
        series[country] = [to_yi(pivot.at[country, period]) if period in pivot.columns else 0 for period in periods]
    return series


def load_dashboard_data() -> dict:
    source = find_source_workbook()
    df = pd.read_excel(source, sheet_name="数据导出")

    required = ["年", "月", "贸易伙伴名称", "人民币"]
    missing = [column for column in required if column not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    df = df[required].copy()
    if df[required].isna().any().any():
        raise ValueError("Required amount columns contain blank values")

    df["年"] = df["年"].astype(int)
    df["月"] = df["月"].astype(int)
    df["人民币"] = pd.to_numeric(df["人民币"], errors="raise")
    df["年月"] = df["年"].astype(str) + "-" + df["月"].astype(str).str.zfill(2)
    df["季度"] = df["年"].astype(str) + "Q" + (((df["月"] - 1) // 3) + 1).astype(str)

    monthly_total = (
        df.groupby("年月", as_index=False)["人民币"].sum().sort_values("年月")
    )
    quarterly_total = (
        df.groupby("季度", as_index=False)["人民币"].sum()
    )
    quarterly_total["sort"] = quarterly_total["季度"].map(period_sort_key)
    quarterly_total = quarterly_total.sort_values("sort").drop(columns=["sort"])

    months = monthly_total["年月"].tolist()
    quarters = quarterly_total["季度"].tolist()
    latest_month = months[-1]
    latest_quarter = quarters[-1]
    previous_month = months[-2] if len(months) > 1 else None
    previous_quarter = quarters[-2] if len(quarters) > 1 else None
    latest_year = int(latest_month.split("-")[0])
    latest_month_no = latest_month.split("-")[1]
    yoy_month = f"{latest_year - 1}-{latest_month_no}"
    yoy_quarter = f"{latest_year - 1}Q{latest_quarter.split('Q')[1]}"

    monthly_map = dict(zip(monthly_total["年月"], monthly_total["人民币"]))
    quarterly_map = dict(zip(quarterly_total["季度"], quarterly_total["人民币"]))
    latest_month_amount = int(monthly_map[latest_month])
    latest_quarter_amount = int(quarterly_map[latest_quarter])

    top_latest_month = ranked_rows(df, "年月", latest_month)
    top_latest_quarter = ranked_rows(df, "季度", latest_quarter)
    selector_countries = [row["country"] for row in top_latest_quarter]
    default_country = selector_countries[0] if selector_countries else ""

    monthly_all = [
        {"period": row["年月"], "amountYi": to_yi(row["人民币"])}
        for _, row in monthly_total.iterrows()
    ]
    quarterly_all = [
        {"period": row["季度"], "amountYi": to_yi(row["人民币"])}
        for _, row in quarterly_total.iterrows()
    ]

    return {
        "meta": {
            "sourceFile": source.name,
            "sourcePath": str(source),
            "sheet": "数据导出",
            "latestMonth": latest_month,
            "latestQuarter": latest_quarter,
            "previousMonth": previous_month,
            "previousQuarter": previous_quarter,
            "yoyMonth": yoy_month if yoy_month in monthly_map else None,
            "yoyQuarter": yoy_quarter if yoy_quarter in quarterly_map else None,
            "latestMonthAmountYi": to_yi(latest_month_amount),
            "latestMonthDeltaYi": to_yi(delta(latest_month_amount, monthly_map.get(previous_month)) or 0)
            if previous_month
            else None,
            "latestQuarterAmountYi": to_yi(latest_quarter_amount),
            "latestQuarterDeltaYi": to_yi(delta(latest_quarter_amount, quarterly_map.get(yoy_quarter)) or 0)
            if yoy_quarter in quarterly_map
            else None,
            "unit": "亿元",
        },
        "all": {
            "month": monthly_all,
            "quarter": quarterly_all,
        },
        "countries": {
            "topLatestMonth": top_latest_month,
            "topLatestQuarter": top_latest_quarter,
            "selector": selector_countries,
            "defaultCountry": default_country,
            "monthPeriods": months,
            "quarterPeriods": quarters,
            "monthSeries": tidy_series(df, "年月", months, selector_countries),
            "quarterSeries": tidy_series(df, "季度", quarters, selector_countries),
        },
    }


HTML_TEMPLATE = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>海关出口金额 BI 看板</title>
  <script src="https://cdn.jsdelivr.net/npm/echarts@5.5.1/dist/echarts.min.js"></script>
  <style>
    :root {
      --bg: #f6f8fb;
      --surface: #ffffff;
      --surface-2: #f0f4f8;
      --surface-3: #e8eef6;
      --text: #172033;
      --muted: #667085;
      --line: #d9e1ec;
      --blue: #2563eb;
      --green: #0f766e;
      --amber: #b7791f;
      --red: #c2410c;
      --purple: #6d5dfc;
      --shadow: 0 14px 34px rgba(23, 32, 51, 0.08);
      font-family: Inter, "Microsoft YaHei", "PingFang SC", "Noto Sans SC", Arial, sans-serif;
    }

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      min-width: 320px;
      background: var(--bg);
      color: var(--text);
    }

    button,
    input {
      font: inherit;
    }

    .app {
      min-height: 100vh;
      padding: 22px;
    }

    .shell {
      width: min(1520px, 100%);
      margin: 0 auto;
    }

    .topbar {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 18px;
      align-items: end;
      margin-bottom: 16px;
    }

    .title-group h1 {
      margin: 0;
      font-size: 25px;
      line-height: 1.22;
      font-weight: 760;
    }

    .title-group p {
      margin: 7px 0 0;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.5;
    }

    .status-strip {
      display: flex;
      flex-wrap: wrap;
      justify-content: flex-end;
      gap: 8px;
      color: var(--muted);
      font-size: 12px;
    }

    .badge {
      display: inline-flex;
      align-items: center;
      min-height: 28px;
      padding: 5px 10px;
      border: 1px solid var(--line);
      border-radius: 999px;
      background: rgba(255, 255, 255, 0.78);
      white-space: nowrap;
    }

    .kpis {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
      margin-bottom: 14px;
    }

    .kpi,
    .panel,
    .side-panel {
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }

    .kpi {
      min-height: 118px;
      padding: 16px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
    }

    .kpi-label {
      color: var(--muted);
      font-size: 13px;
      line-height: 1.4;
    }

    .kpi-value {
      margin-top: 12px;
      font-size: 30px;
      line-height: 1;
      font-weight: 760;
      word-break: break-word;
    }

    .kpi-sub {
      margin-top: 10px;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.45;
    }

    .positive {
      color: var(--green);
    }

    .negative {
      color: var(--red);
    }

    .grid {
      display: grid;
      grid-template-columns: minmax(0, 1fr) 360px;
      gap: 16px;
      align-items: start;
    }

    .main-stack {
      display: grid;
      gap: 16px;
      min-width: 0;
    }

    .side-stack {
      display: grid;
      gap: 16px;
      position: sticky;
      top: 18px;
    }

    .panel {
      min-width: 0;
      padding: 18px;
    }

    .panel-head {
      display: flex;
      justify-content: space-between;
      gap: 16px;
      align-items: flex-start;
      margin-bottom: 10px;
    }

    .panel-title h2,
    .side-panel h2 {
      margin: 0;
      font-size: 18px;
      line-height: 1.3;
      font-weight: 730;
    }

    .panel-title p {
      margin: 6px 0 0;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.45;
    }

    .segmented {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 4px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface-2);
      flex: 0 0 auto;
    }

    .segmented button {
      min-width: 54px;
      height: 30px;
      border: 0;
      border-radius: 6px;
      background: transparent;
      color: #344054;
      cursor: pointer;
      padding: 0 10px;
      font-size: 12px;
    }

    .segmented button.active {
      background: var(--surface);
      color: var(--blue);
      box-shadow: 0 2px 8px rgba(23, 32, 51, 0.08);
      font-weight: 700;
    }

    .chart {
      width: 100%;
      height: 520px;
    }

    .region-cards {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
    }

    .region-card {
      min-height: 104px;
      padding: 14px;
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }

    .region-card-label {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.4;
    }

    .region-card-value {
      margin-top: 10px;
      color: var(--text);
      font-size: 23px;
      line-height: 1;
      font-weight: 760;
      word-break: break-word;
    }

    .region-card-sub {
      margin-top: 9px;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.45;
    }

    #countryChart {
      height: 610px;
    }

    .country-tools {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 12px;
      align-items: start;
      margin-bottom: 12px;
    }

    .select-box {
      position: relative;
      min-width: 0;
    }

    .select-trigger {
      width: 100%;
      min-height: 42px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
      color: var(--text);
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 8px 12px;
      cursor: pointer;
    }

    .trigger-text {
      min-width: 0;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      text-align: left;
    }

    .chevron {
      width: 9px;
      height: 9px;
      border-right: 2px solid var(--muted);
      border-bottom: 2px solid var(--muted);
      transform: rotate(45deg);
      margin-bottom: 4px;
      flex: 0 0 auto;
    }

    .dropdown {
      display: none;
      position: absolute;
      z-index: 30;
      top: calc(100% + 6px);
      left: 0;
      right: 0;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
      box-shadow: 0 18px 48px rgba(23, 32, 51, 0.18);
      padding: 10px;
    }

    .dropdown.open {
      display: block;
    }

    .search {
      width: 100%;
      height: 38px;
      border: 1px solid var(--line);
      border-radius: 7px;
      padding: 0 10px;
      outline: none;
      margin-bottom: 8px;
    }

    .actions {
      display: flex;
      gap: 8px;
      margin-bottom: 8px;
    }

    .action-btn {
      flex: 1;
      height: 34px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: var(--surface);
      color: #344054;
      cursor: pointer;
    }

    .action-btn.primary {
      background: var(--blue);
      border-color: var(--blue);
      color: #fff;
    }

    .country-list {
      max-height: 330px;
      overflow: auto;
      border-top: 1px solid var(--line);
      padding-top: 6px;
    }

    .country-option {
      display: flex;
      align-items: center;
      gap: 9px;
      min-height: 36px;
      padding: 6px 4px;
      border-radius: 6px;
      cursor: pointer;
    }

    .country-option:hover {
      background: var(--surface-2);
    }

    .country-option input {
      width: 16px;
      height: 16px;
      accent-color: var(--blue);
      flex: 0 0 auto;
    }

    .country-name {
      min-width: 0;
      flex: 1;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      font-size: 13px;
    }

    .selected-list {
      display: flex;
      flex-wrap: wrap;
      gap: 7px;
      min-height: 30px;
      margin-bottom: 12px;
    }

    .tag {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      max-width: 100%;
      min-height: 28px;
      padding: 5px 9px;
      border-radius: 999px;
      background: var(--surface-2);
      color: #344054;
      font-size: 12px;
    }

    .tag span {
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .tag button {
      width: 18px;
      height: 18px;
      border: 0;
      border-radius: 999px;
      background: rgba(102, 112, 133, 0.18);
      cursor: pointer;
      padding: 0;
      line-height: 18px;
      color: #344054;
    }

    .side-panel {
      padding: 16px;
    }

    .side-panel h2 {
      font-size: 16px;
      margin-bottom: 4px;
    }

    .side-panel p {
      margin: 0 0 12px;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.45;
    }

    .rank-list {
      display: grid;
      gap: 6px;
    }

    .rank-row {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      align-items: center;
      gap: 12px;
      min-height: 30px;
      color: #344054;
      font-size: 12px;
    }

    .rank-name {
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .rank-value {
      color: var(--muted);
      white-space: nowrap;
    }

    .source-note {
      margin-top: 16px;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.55;
    }

    .empty-state {
      padding: 18px 4px;
      color: var(--muted);
      font-size: 13px;
      text-align: center;
    }

    @media (max-width: 1180px) {
      .grid {
        grid-template-columns: 1fr;
      }

      .side-stack {
        position: static;
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }

      .kpis {
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }

      .region-cards {
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }
    }

    @media (max-width: 760px) {
      .app {
        padding: 14px;
      }

      .topbar,
      .panel-head,
      .country-tools {
        grid-template-columns: 1fr;
      }

      .panel-head {
        display: grid;
      }

      .status-strip {
        justify-content: flex-start;
      }

      .kpis,
      .region-cards,
      .side-stack {
        grid-template-columns: 1fr;
      }

      .kpi-value {
        font-size: 26px;
      }

      .panel {
        padding: 14px;
      }

      .chart {
        height: 450px;
      }

      #countryChart {
        height: 560px;
      }
    }
  </style>
</head>
<body>
  <div class="app">
    <div class="shell">
      <header class="topbar">
        <div class="title-group">
          <h1>海关出口金额 BI 看板</h1>
          <p>只展示人民币出口金额，统一换算为亿元</p>
        </div>
        <div class="status-strip">
          <span class="badge" id="periodBadge"></span>
          <span class="badge">金额单位：亿元</span>
          <span class="badge" id="sourceBadge"></span>
        </div>
      </header>

      <section class="kpis" aria-label="关键金额变化">
        <article class="kpi">
          <div>
            <div class="kpi-label">最新月出口额</div>
            <div class="kpi-value" id="latestMonthAmount">--</div>
          </div>
          <div class="kpi-sub" id="latestMonthLabel">--</div>
        </article>
        <article class="kpi">
          <div>
            <div class="kpi-label">较上月变化</div>
            <div class="kpi-value" id="monthDelta">--</div>
          </div>
          <div class="kpi-sub" id="monthDeltaLabel">--</div>
        </article>
        <article class="kpi">
          <div>
            <div class="kpi-label">最新季度出口额</div>
            <div class="kpi-value" id="latestQuarterAmount">--</div>
          </div>
          <div class="kpi-sub" id="latestQuarterLabel">--</div>
        </article>
        <article class="kpi">
          <div>
            <div class="kpi-label">较去年同季变化</div>
            <div class="kpi-value" id="quarterDelta">--</div>
          </div>
          <div class="kpi-sub" id="quarterDeltaLabel">--</div>
        </article>
      </section>

      <div class="grid">
        <main class="main-stack">
          <section class="panel" aria-label="所有国家合计出口金额">
            <div class="panel-head">
              <div class="panel-title">
                <h2>所有国家合计出口额</h2>
                <p id="allSubtitle">默认按月展示，可切换季度；柱状图显示每个周期的完整金额。</p>
              </div>
              <div class="segmented" aria-label="合计图周期切换">
                <button id="allMonthBtn" class="active" type="button">月</button>
                <button id="allQuarterBtn" type="button">季度</button>
              </div>
            </div>
            <div id="allChart" class="chart" role="img" aria-label="所有国家合计出口额柱状图"></div>
          </section>

          <section class="region-cards" aria-label="区域国家金额变化">
            <article class="region-card">
              <div class="region-card-label">当月环比</div>
              <div class="region-card-value" id="regionMonthMom">--</div>
              <div class="region-card-sub" id="regionMonthMomSub">--</div>
            </article>
            <article class="region-card">
              <div class="region-card-label">当月同比</div>
              <div class="region-card-value" id="regionMonthYoy">--</div>
              <div class="region-card-sub" id="regionMonthYoySub">--</div>
            </article>
            <article class="region-card">
              <div class="region-card-label">当季环比</div>
              <div class="region-card-value" id="regionQuarterQoq">--</div>
              <div class="region-card-sub" id="regionQuarterQoqSub">--</div>
            </article>
            <article class="region-card">
              <div class="region-card-label">当季同比</div>
              <div class="region-card-value" id="regionQuarterYoy">--</div>
              <div class="region-card-sub" id="regionQuarterYoySub">--</div>
            </article>
          </section>

          <section class="panel" aria-label="区域国家出口金额趋势">
            <div class="panel-head">
              <div class="panel-title">
                <h2>区域国家出口趋势</h2>
                <p id="countrySubtitle">Top 15 贸易伙伴范围 · 默认展示最新季度第一名国家。</p>
              </div>
              <div class="segmented" aria-label="国家图周期切换">
                <button id="countryMonthBtn" class="active" type="button">月</button>
                <button id="countryQuarterBtn" type="button">季度</button>
              </div>
            </div>

            <div class="country-tools">
              <div class="select-box" id="selectBox">
                <button class="select-trigger" id="selectTrigger" type="button" aria-expanded="false">
                  <span class="trigger-text" id="triggerText">--</span>
                  <span class="chevron" aria-hidden="true"></span>
                </button>
                <div class="dropdown" id="dropdown">
                  <input class="search" id="countrySearch" type="search" placeholder="搜索区域国家" autocomplete="off">
                  <div class="actions">
                    <button class="action-btn primary" id="selectDefault" type="button">恢复默认</button>
                    <button class="action-btn" id="selectAll" type="button">选择全部</button>
                    <button class="action-btn" id="clearSelected" type="button">清空</button>
                  </div>
                  <div class="country-list" id="countryList"></div>
                </div>
              </div>
            </div>
            <div class="selected-list" id="selectedList"></div>
            <div id="countryChart" class="chart" role="img" aria-label="区域国家出口额曲线图"></div>
          </section>
        </main>

        <aside class="side-stack" aria-label="Top 15 国家">
          <section class="side-panel">
            <h2 id="latestMonthTopTitle">当月出口额 Top 15</h2>
            <p>仅显示贸易伙伴名称和出口额。</p>
            <div class="rank-list" id="latestMonthTop"></div>
          </section>

          <section class="side-panel">
            <h2 id="latestQuarterTopTitle">Q1 出口额 Top 15</h2>
            <p>仅显示贸易伙伴名称和出口额。</p>
            <div class="rank-list" id="latestQuarterTop"></div>
          </section>

          <section class="side-panel">
            <h2>数据口径</h2>
            <p>金额字段取“人民币”，页面内所有金额均为亿元；贸易伙伴仅显示名称。</p>
            <div class="source-note" id="sourceNote"></div>
          </section>
        </aside>
      </div>
    </div>
  </div>

  <script>
    const BI_DATA = __BI_DATA__;
    const amountFormatter = new Intl.NumberFormat("zh-CN", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
    const compactFormatter = new Intl.NumberFormat("zh-CN", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });

    const palette = [
      "#2563eb", "#0f766e", "#b7791f", "#c2410c", "#6d5dfc",
      "#0891b2", "#9333ea", "#ca8a04", "#dc2626", "#16a34a",
      "#475569", "#be185d", "#64748b", "#0284c7", "#7c3aed"
    ];

    const state = {
      allMode: "month",
      countryMode: "month",
      allHighlight: null,
      countryHighlight: null,
      selectedCountries: new Set([BI_DATA.countries.defaultCountry])
    };

    const dom = {
      periodBadge: document.getElementById("periodBadge"),
      sourceBadge: document.getElementById("sourceBadge"),
      latestMonthAmount: document.getElementById("latestMonthAmount"),
      latestMonthLabel: document.getElementById("latestMonthLabel"),
      monthDelta: document.getElementById("monthDelta"),
      monthDeltaLabel: document.getElementById("monthDeltaLabel"),
      latestQuarterAmount: document.getElementById("latestQuarterAmount"),
      latestQuarterLabel: document.getElementById("latestQuarterLabel"),
      quarterDelta: document.getElementById("quarterDelta"),
      quarterDeltaLabel: document.getElementById("quarterDeltaLabel"),
      allSubtitle: document.getElementById("allSubtitle"),
      allMonthBtn: document.getElementById("allMonthBtn"),
      allQuarterBtn: document.getElementById("allQuarterBtn"),
      countrySubtitle: document.getElementById("countrySubtitle"),
      regionMonthMom: document.getElementById("regionMonthMom"),
      regionMonthMomSub: document.getElementById("regionMonthMomSub"),
      regionMonthYoy: document.getElementById("regionMonthYoy"),
      regionMonthYoySub: document.getElementById("regionMonthYoySub"),
      regionQuarterQoq: document.getElementById("regionQuarterQoq"),
      regionQuarterQoqSub: document.getElementById("regionQuarterQoqSub"),
      regionQuarterYoy: document.getElementById("regionQuarterYoy"),
      regionQuarterYoySub: document.getElementById("regionQuarterYoySub"),
      countryMonthBtn: document.getElementById("countryMonthBtn"),
      countryQuarterBtn: document.getElementById("countryQuarterBtn"),
      selectBox: document.getElementById("selectBox"),
      selectTrigger: document.getElementById("selectTrigger"),
      triggerText: document.getElementById("triggerText"),
      dropdown: document.getElementById("dropdown"),
      countrySearch: document.getElementById("countrySearch"),
      countryList: document.getElementById("countryList"),
      selectDefault: document.getElementById("selectDefault"),
      selectAll: document.getElementById("selectAll"),
      clearSelected: document.getElementById("clearSelected"),
      selectedList: document.getElementById("selectedList"),
      latestMonthTopTitle: document.getElementById("latestMonthTopTitle"),
      latestQuarterTopTitle: document.getElementById("latestQuarterTopTitle"),
      latestMonthTop: document.getElementById("latestMonthTop"),
      latestQuarterTop: document.getElementById("latestQuarterTop"),
      sourceNote: document.getElementById("sourceNote")
    };

    const allChart = echarts.init(document.getElementById("allChart"), null, { renderer: "canvas" });
    const countryChart = echarts.init(document.getElementById("countryChart"), null, { renderer: "canvas" });

    function formatAmount(value) {
      return `${amountFormatter.format(Number(value || 0))} 亿元`;
    }

    function formatCompact(value) {
      return `${compactFormatter.format(Number(value || 0))} 亿元`;
    }

    function signedClass(value) {
      if (value === null || value === undefined || Number(value) === 0) return "";
      return Number(value) > 0 ? "positive" : "negative";
    }

    function formatSignedAmount(value) {
      if (value === null || value === undefined) return "--";
      const numeric = Number(value);
      const sign = numeric > 0 ? "+" : "";
      return `${sign}${formatAmount(numeric)}`;
    }

    function formatPct(value) {
      if (value === null || value === undefined || Number.isNaN(Number(value))) return "--";
      const numeric = Number(value);
      const sign = numeric > 0 ? "+" : "";
      return `${sign}${(numeric * 100).toFixed(2)}%`;
    }

    function calcDelta(current, prior) {
      if (prior === null || prior === undefined) return null;
      return Number(current || 0) - Number(prior || 0);
    }

    function calcPct(current, prior) {
      if (!prior) return null;
      return (Number(current || 0) - Number(prior)) / Number(prior);
    }

    function setChangeCard(valueNode, subNode, deltaValue, pctValue, currentPeriod, priorPeriod, currentAmount) {
      valueNode.textContent = formatSignedAmount(deltaValue);
      valueNode.className = `region-card-value ${signedClass(deltaValue)}`.trim();
      subNode.textContent = priorPeriod
        ? `${formatPct(pctValue)} · ${currentPeriod} 对比 ${priorPeriod} · 当前 ${formatAmount(currentAmount)}`
        : "--";
    }

    function setMode(buttons, activeMode) {
      buttons.month.classList.toggle("active", activeMode === "month");
      buttons.quarter.classList.toggle("active", activeMode === "quarter");
    }

    function updateKpis() {
      const meta = BI_DATA.meta;
      dom.periodBadge.textContent = `${BI_DATA.all.month[0].period} 至 ${meta.latestMonth}`;
      dom.sourceBadge.textContent = meta.sourceFile;
      dom.latestMonthAmount.textContent = formatAmount(meta.latestMonthAmountYi);
      dom.latestMonthLabel.textContent = meta.latestMonth;
      dom.monthDelta.textContent = formatSignedAmount(meta.latestMonthDeltaYi);
      dom.monthDelta.className = `kpi-value ${signedClass(meta.latestMonthDeltaYi)}`.trim();
      dom.monthDeltaLabel.textContent = meta.previousMonth ? `对比 ${meta.previousMonth}` : "--";
      dom.latestQuarterAmount.textContent = formatAmount(meta.latestQuarterAmountYi);
      dom.latestQuarterLabel.textContent = meta.latestQuarter;
      dom.quarterDelta.textContent = formatSignedAmount(meta.latestQuarterDeltaYi);
      dom.quarterDelta.className = `kpi-value ${signedClass(meta.latestQuarterDeltaYi)}`.trim();
      dom.quarterDeltaLabel.textContent = meta.yoyQuarter ? `对比 ${meta.yoyQuarter}` : "--";
      dom.latestMonthTopTitle.textContent = `${meta.latestMonth} 出口额 Top 15`;
      dom.latestQuarterTopTitle.textContent = `${meta.latestQuarter} 出口额 Top 15`;
      dom.sourceNote.textContent = `${meta.sourceFile} / ${meta.sheet}`;
    }

    function allTooltip(params) {
      const item = Array.isArray(params) ? params[0] : params;
      return `<strong>${item.name}</strong><br>出口额：${formatAmount(item.value)}`;
    }

    function renderAllChart() {
      const mode = state.allMode;
      const rows = BI_DATA.all[mode];
      const labels = rows.map((row) => row.period);
      const values = rows.map((row) => {
        const active = !state.allHighlight || state.allHighlight === row.period;
        return {
          name: row.period,
          value: row.amountYi,
          itemStyle: {
            color: active ? "#2563eb" : "#cbd5e1",
            opacity: active ? 1 : 0.34,
            borderColor: active && state.allHighlight ? "#172033" : "transparent",
            borderWidth: active && state.allHighlight ? 1 : 0
          },
          label: {
            color: active ? "#344054" : "#98a2b3"
          }
        };
      });
      const isQuarter = mode === "quarter";

      setMode({ month: dom.allMonthBtn, quarter: dom.allQuarterBtn }, mode);
      dom.allSubtitle.textContent = isQuarter
        ? "所有国家按季度加总，金额单位为亿元。"
        : "所有国家按月加总，金额单位为亿元。";

      allChart.setOption({
        animationDuration: 350,
        grid: { top: 42, right: 22, bottom: isQuarter ? 54 : 78, left: 70, containLabel: true },
        tooltip: { trigger: "axis", axisPointer: { type: "shadow" }, confine: true, formatter: allTooltip },
        xAxis: {
          type: "category",
          data: labels,
          axisTick: { alignWithLabel: true },
          axisLabel: {
            interval: 0,
            rotate: isQuarter ? 0 : 35,
            color: "#667085",
            fontSize: 12
          },
          axisLine: { lineStyle: { color: "#d9e1ec" } }
        },
        yAxis: {
          type: "value",
          name: "亿元",
          nameTextStyle: { color: "#667085" },
          axisLabel: { color: "#667085", formatter: (value) => compactFormatter.format(value) },
          splitLine: { lineStyle: { color: "#edf2f7" } }
        },
        series: [{
          name: "所有国家合计",
          type: "bar",
          barMaxWidth: 42,
          data: values,
          itemStyle: { color: "#2563eb", borderRadius: [4, 4, 0, 0] },
          label: {
            show: true,
            position: "top",
            formatter: (item) => amountFormatter.format(item.value),
            color: "#344054",
            fontSize: 11
          }
        }]
      }, true);
    }

    function countryTooltip(params) {
      const rows = Array.isArray(params) ? params : [params];
      const lines = rows
        .map((item) => `<div style="display:flex;gap:18px;justify-content:space-between;min-width:220px;">
          <span>${item.marker}${item.seriesName}</span><strong>${formatAmount(item.value)}</strong>
        </div>`)
        .join("");
      return `<strong>${rows[0].axisValue}</strong>${lines}`;
    }

    function activeCountryList() {
      return Array.from(state.selectedCountries).filter(Boolean);
    }

    function sumSelectedAt(mode, period) {
      if (!period) return null;
      const selected = activeCountryList();
      if (!selected.length) return 0;
      const periods = mode === "month" ? BI_DATA.countries.monthPeriods : BI_DATA.countries.quarterPeriods;
      const seriesMap = mode === "month" ? BI_DATA.countries.monthSeries : BI_DATA.countries.quarterSeries;
      const periodIndex = periods.indexOf(period);
      if (periodIndex < 0) return null;
      return selected.reduce((sum, country) => sum + Number(seriesMap[country]?.[periodIndex] || 0), 0);
    }

    function updateRegionCards() {
      const meta = BI_DATA.meta;
      const monthCurrent = sumSelectedAt("month", meta.latestMonth);
      const monthPrevious = sumSelectedAt("month", meta.previousMonth);
      const monthYoy = sumSelectedAt("month", meta.yoyMonth);
      const quarterCurrent = sumSelectedAt("quarter", meta.latestQuarter);
      const quarterPrevious = sumSelectedAt("quarter", meta.previousQuarter);
      const quarterYoy = sumSelectedAt("quarter", meta.yoyQuarter);

      setChangeCard(
        dom.regionMonthMom,
        dom.regionMonthMomSub,
        calcDelta(monthCurrent, monthPrevious),
        calcPct(monthCurrent, monthPrevious),
        meta.latestMonth,
        meta.previousMonth,
        monthCurrent
      );
      setChangeCard(
        dom.regionMonthYoy,
        dom.regionMonthYoySub,
        calcDelta(monthCurrent, monthYoy),
        calcPct(monthCurrent, monthYoy),
        meta.latestMonth,
        meta.yoyMonth,
        monthCurrent
      );
      setChangeCard(
        dom.regionQuarterQoq,
        dom.regionQuarterQoqSub,
        calcDelta(quarterCurrent, quarterPrevious),
        calcPct(quarterCurrent, quarterPrevious),
        meta.latestQuarter,
        meta.previousQuarter,
        quarterCurrent
      );
      setChangeCard(
        dom.regionQuarterYoy,
        dom.regionQuarterYoySub,
        calcDelta(quarterCurrent, quarterYoy),
        calcPct(quarterCurrent, quarterYoy),
        meta.latestQuarter,
        meta.yoyQuarter,
        quarterCurrent
      );
    }

    function renderCountryChart() {
      const mode = state.countryMode;
      const periods = mode === "month" ? BI_DATA.countries.monthPeriods : BI_DATA.countries.quarterPeriods;
      const seriesMap = mode === "month" ? BI_DATA.countries.monthSeries : BI_DATA.countries.quarterSeries;
      const selected = activeCountryList();
      const isQuarter = mode === "quarter";

      setMode({ month: dom.countryMonthBtn, quarter: dom.countryQuarterBtn }, mode);
      dom.countrySubtitle.textContent = `${selected.length || 0} 个国家 · ${isQuarter ? "季度" : "月度"}出口额趋势，金额单位为亿元。`;

      countryChart.setOption({
        animationDuration: 350,
        color: palette,
        grid: { top: selected.length > 1 ? 92 : 54, right: 26, bottom: isQuarter ? 60 : 82, left: 70, containLabel: true },
        legend: {
          show: selected.length > 1,
          type: "scroll",
          top: 4,
          left: 8,
          right: 8,
          height: 56,
          textStyle: { color: "#344054", fontSize: 12 }
        },
        tooltip: { trigger: "axis", confine: true, formatter: countryTooltip },
        xAxis: {
          type: "category",
          data: periods,
          boundaryGap: false,
          axisLabel: { interval: 0, rotate: isQuarter ? 0 : 35, color: "#667085", fontSize: 12 },
          axisLine: { lineStyle: { color: "#d9e1ec" } }
        },
        yAxis: {
          type: "value",
          name: "亿元",
          nameTextStyle: { color: "#667085" },
          axisLabel: { color: "#667085", formatter: (value) => compactFormatter.format(value) },
          splitLine: { lineStyle: { color: "#edf2f7" } }
        },
        series: selected.map((country, index) => {
          const color = palette[index % palette.length];
          const data = (seriesMap[country] || periods.map(() => 0)).map((value, periodIndex) => {
            const period = periods[periodIndex];
            const active = !state.countryHighlight || state.countryHighlight === period;
            return {
              name: period,
              value,
              symbolSize: active && state.countryHighlight ? 11 : 7,
              itemStyle: {
                color: active ? color : "#cbd5e1",
                opacity: active ? 1 : 0.28,
                borderColor: active && state.countryHighlight ? "#172033" : color,
                borderWidth: active && state.countryHighlight ? 1.5 : 0
              },
              label: {
                color: active ? "#344054" : "#98a2b3"
              }
            };
          });
          const markLine = state.countryHighlight && index === 0
            ? {
                symbol: ["none", "none"],
                label: { show: false },
                lineStyle: { color: "#172033", width: 1, opacity: 0.34 },
                data: [{ xAxis: state.countryHighlight }]
              }
            : undefined;

          return {
            name: country,
            type: "line",
            smooth: true,
            symbol: "circle",
            symbolSize: 7,
            connectNulls: true,
            data,
            lineStyle: { width: 3, opacity: state.countryHighlight ? 0.58 : 1 },
            itemStyle: { color },
            label: {
              show: true,
              position: index % 2 === 0 ? "top" : "bottom",
              formatter: (item) => amountFormatter.format(item.value),
              color: "#344054",
              fontSize: 10
            },
            labelLayout: { hideOverlap: false },
            emphasis: { focus: "series" },
            markLine
          };
        })
      }, true);
    }

    function renderRanking(container, rows) {
      container.innerHTML = "";
      rows.forEach((row) => {
        const div = document.createElement("div");
        div.className = "rank-row";
        const name = document.createElement("span");
        name.className = "rank-name";
        name.title = row.country;
        name.textContent = row.country;
        const amount = document.createElement("span");
        amount.className = "rank-value";
        amount.textContent = formatCompact(row.amountYi);
        div.append(name, amount);
        container.appendChild(div);
      });
    }

    function updateTrigger() {
      const selected = activeCountryList();
      if (!selected.length) {
        dom.triggerText.textContent = "未选择国家";
        return;
      }
      const preview = selected.slice(0, 3).join("、");
      dom.triggerText.textContent = selected.length > 3 ? `${preview} 等 ${selected.length} 个` : preview;
    }

    function updateSelectedTags() {
      dom.selectedList.innerHTML = "";
      const selected = activeCountryList();
      if (!selected.length) {
        const tag = document.createElement("div");
        tag.className = "tag";
        tag.innerHTML = "<span>未选择国家</span>";
        dom.selectedList.appendChild(tag);
        return;
      }
      selected.forEach((country) => {
        const tag = document.createElement("div");
        tag.className = "tag";
        const label = document.createElement("span");
        label.textContent = country;
        const remove = document.createElement("button");
        remove.type = "button";
        remove.setAttribute("aria-label", `移除 ${country}`);
        remove.textContent = "×";
        remove.addEventListener("click", () => {
          state.selectedCountries.delete(country);
          refreshCountry();
        });
        tag.append(label, remove);
        dom.selectedList.appendChild(tag);
      });
    }

    function renderCountryList() {
      const query = dom.countrySearch.value.trim().toLowerCase();
      const countries = BI_DATA.countries.selector.filter((country) => country.toLowerCase().includes(query));
      dom.countryList.innerHTML = "";

      if (!countries.length) {
        const empty = document.createElement("div");
        empty.className = "empty-state";
        empty.textContent = "无匹配国家";
        dom.countryList.appendChild(empty);
        return;
      }

      countries.forEach((country) => {
        const label = document.createElement("label");
        label.className = "country-option";
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.checked = state.selectedCountries.has(country);
        checkbox.addEventListener("change", () => {
          if (checkbox.checked) {
            state.selectedCountries.add(country);
          } else {
            state.selectedCountries.delete(country);
          }
          refreshCountry();
        });
        const name = document.createElement("span");
        name.className = "country-name";
        name.textContent = country;
        label.append(checkbox, name);
        dom.countryList.appendChild(label);
      });
    }

    function toggleDropdown(forceOpen) {
      const open = forceOpen ?? !dom.dropdown.classList.contains("open");
      dom.dropdown.classList.toggle("open", open);
      dom.selectTrigger.setAttribute("aria-expanded", String(open));
      if (open) dom.countrySearch.focus();
    }

    function refreshCountry() {
      updateTrigger();
      updateSelectedTags();
      renderCountryList();
      updateRegionCards();
      renderCountryChart();
    }

    dom.allMonthBtn.addEventListener("click", () => {
      state.allMode = "month";
      state.allHighlight = null;
      renderAllChart();
    });
    dom.allQuarterBtn.addEventListener("click", () => {
      state.allMode = "quarter";
      state.allHighlight = null;
      renderAllChart();
    });
    dom.countryMonthBtn.addEventListener("click", () => {
      state.countryMode = "month";
      state.countryHighlight = null;
      renderCountryChart();
    });
    dom.countryQuarterBtn.addEventListener("click", () => {
      state.countryMode = "quarter";
      state.countryHighlight = null;
      renderCountryChart();
    });
    dom.selectTrigger.addEventListener("click", () => toggleDropdown());
    dom.countrySearch.addEventListener("input", renderCountryList);
    dom.selectDefault.addEventListener("click", () => {
      state.selectedCountries = new Set([BI_DATA.countries.defaultCountry]);
      refreshCountry();
    });
    dom.selectAll.addEventListener("click", () => {
      state.selectedCountries = new Set(BI_DATA.countries.selector);
      refreshCountry();
    });
    dom.clearSelected.addEventListener("click", () => {
      state.selectedCountries.clear();
      refreshCountry();
    });
    document.addEventListener("click", (event) => {
      if (!dom.selectBox.contains(event.target)) toggleDropdown(false);
    });
    window.addEventListener("resize", () => {
      allChart.resize();
      countryChart.resize();
    });

    allChart.on("click", (params) => {
      const period = params?.name || params?.axisValue;
      if (!period) return;
      state.allHighlight = state.allHighlight === period ? null : period;
      renderAllChart();
    });

    countryChart.on("click", (params) => {
      const period = params?.name || params?.axisValue;
      if (!period) return;
      state.countryHighlight = state.countryHighlight === period ? null : period;
      renderCountryChart();
    });

    updateKpis();
    renderRanking(dom.latestMonthTop, BI_DATA.countries.topLatestMonth);
    renderRanking(dom.latestQuarterTop, BI_DATA.countries.topLatestQuarter);
    renderAllChart();
    refreshCountry();
  </script>
</body>
</html>
"""


def build_html(data: dict) -> str:
    data_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    return HTML_TEMPLATE.replace("__BI_DATA__", data_json)


def main() -> None:
    data = load_dashboard_data()
    html = build_html(data)
    PUBLIC_DIR.mkdir(exist_ok=True)
    OUTPUT_FILE.write_text(html, encoding="utf-8")
    PUBLIC_OUTPUT_FILE.write_text(html, encoding="utf-8")
    print(json.dumps(data["meta"], ensure_ascii=False, indent=2))
    print(f"Wrote {OUTPUT_FILE}")
    print(f"Wrote {PUBLIC_OUTPUT_FILE}")


if __name__ == "__main__":
    main()

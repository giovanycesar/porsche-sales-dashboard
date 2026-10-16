# build_dashboard.py
"""
Script to build an executive, elegant, Porsche-branded HTML dashboard
from porsche_database_cleaned.xlsx (or _dashboard_data.json).
"""

import os
import json
import pandas as pd

# Locate directory
target_dir = r"C:\Users\gvnce\OneDrive\Área de Trabalho\Dados\porsche_project"
json_path = os.path.join(target_dir, "_dashboard_data.json")

if not os.path.exists(json_path):
    excel_path = os.path.join(target_dir, "porsche_database_cleaned.xlsx")
    df = pd.read_excel(excel_path)
    records = df.to_dict(orient="records")
else:
    with open(json_path, "r", encoding="utf-8") as f:
        records = json.load(f)

json_data_str = json.dumps(records, ensure_ascii=False)

html_content = f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PORSCHE | Sales Intelligence Dashboard</title>
  
  <!-- Google Fonts: Inter & Montserrat for Porsche luxury feel -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Montserrat:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
  
  <!-- Font Awesome -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  
  <!-- Chart.js -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>

  <style>
    :root {{
      --bg-main: #0a0b0d;
      --bg-card: #121418;
      --bg-card-hover: #181b21;
      --bg-input: #1a1d24;
      --porsche-red: #d5001c;
      --porsche-red-hover: #b30017;
      --porsche-red-glow: rgba(213, 0, 28, 0.25);
      --porsche-gold: #c6a15b;
      --border-color: rgba(255, 255, 255, 0.08);
      --border-focus: rgba(213, 0, 28, 0.5);
      --text-main: #ffffff;
      --text-secondary: #9ea3ae;
      --text-muted: #646b79;
      --badge-green: #10b981;
      --badge-blue: #3b82f6;
      --badge-amber: #f59e0b;
      --shadow-card: 0 12px 30px rgba(0, 0, 0, 0.6);
      --transition-smooth: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
    }}

    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }}

    body {{
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
      background-color: var(--bg-main);
      color: var(--text-main);
      min-height: 100vh;
      line-height: 1.5;
      overflow-x: hidden;
      background-image: 
        radial-gradient(circle at 50% 0%, rgba(35, 40, 50, 0.5) 0%, transparent 60%),
        radial-gradient(circle at 90% 20%, rgba(213, 0, 28, 0.04) 0%, transparent 40%);
    }}

    /* Custom Scrollbar */
    ::-webkit-scrollbar {{
      width: 8px;
      height: 8px;
    }}
    ::-webkit-scrollbar-track {{
      background: #0a0b0d;
    }}
    ::-webkit-scrollbar-thumb {{
      background: #252830;
      border-radius: 4px;
    }}
    ::-webkit-scrollbar-thumb:hover {{
      background: var(--porsche-red);
    }}

    /* Container */
    .container {{
      max-width: 1540px;
      margin: 0 auto;
      padding: 0 2rem 4rem 2rem;
    }}

    /* Porsche Header */
    header.porsche-header {{
      border-bottom: 1px solid var(--border-color);
      background: rgba(10, 11, 13, 0.85);
      backdrop-filter: blur(16px);
      position: sticky;
      top: 0;
      z-index: 100;
      padding: 1.25rem 0;
    }}

    .header-inner {{
      max-width: 1540px;
      margin: 0 auto;
      padding: 0 2rem;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 1.5rem;
    }}

    .brand-section {{
      display: flex;
      align-items: center;
      gap: 1.75rem;
    }}

    .porsche-wordmark {{
      font-family: 'Montserrat', sans-serif;
      font-size: 1.6rem;
      font-weight: 900;
      letter-spacing: 0.45em;
      text-transform: uppercase;
      color: #ffffff;
      position: relative;
      display: inline-block;
      padding-bottom: 4px;
    }}

    .porsche-wordmark::after {{
      content: '';
      position: absolute;
      left: 0;
      bottom: 0;
      width: 38px;
      height: 2.5px;
      background: var(--porsche-red);
      border-radius: 2px;
    }}

    .brand-divider {{
      height: 28px;
      width: 1px;
      background: var(--border-color);
    }}

    .platform-title {{
      display: flex;
      flex-direction: column;
    }}

    .platform-title span.title {{
      font-size: 0.82rem;
      font-weight: 700;
      letter-spacing: 0.18em;
      text-transform: uppercase;
      color: #e5e7eb;
    }}

    .platform-title span.subtitle {{
      font-size: 0.72rem;
      color: var(--text-muted);
      letter-spacing: 0.08em;
    }}

    .header-status {{
      display: flex;
      align-items: center;
      gap: 1rem;
    }}

    .status-pill {{
      display: flex;
      align-items: center;
      gap: 0.5rem;
      padding: 0.4rem 0.9rem;
      background: rgba(16, 185, 129, 0.08);
      border: 1px solid rgba(16, 185, 129, 0.3);
      border-radius: 9999px;
      font-size: 0.72rem;
      font-weight: 600;
      letter-spacing: 0.06em;
      color: #34d399;
      text-transform: uppercase;
    }}

    .status-dot {{
      width: 7px;
      height: 7px;
      border-radius: 50%;
      background: #10b981;
      box-shadow: 0 0 10px #10b981;
      animation: pulse-dot 2s infinite;
    }}

    @keyframes pulse-dot {{
      0%, 100% {{ opacity: 1; transform: scale(1); }}
      50% {{ opacity: 0.4; transform: scale(0.85); }}
    }}

    .btn-export {{
      background: transparent;
      border: 1px solid var(--border-color);
      color: var(--text-main);
      padding: 0.5rem 1rem;
      border-radius: 6px;
      font-size: 0.75rem;
      font-weight: 600;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 0.5rem;
      transition: var(--transition-smooth);
    }}

    .btn-export:hover {{
      background: rgba(255, 255, 255, 0.06);
      border-color: rgba(255, 255, 255, 0.25);
    }}

    /* Hero / Intro */
    .hero-banner {{
      padding: 2.25rem 0 1.5rem 0;
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
      border-bottom: 1px solid var(--border-color);
      margin-bottom: 2rem;
    }}

    .hero-text h1 {{
      font-family: 'Montserrat', sans-serif;
      font-size: 1.85rem;
      font-weight: 700;
      letter-spacing: -0.02em;
      margin-bottom: 0.35rem;
      color: #ffffff;
    }}

    .hero-text p {{
      color: var(--text-secondary);
      font-size: 0.92rem;
      max-width: 780px;
    }}

    .hero-tag {{
      display: inline-block;
      font-size: 0.68rem;
      font-weight: 700;
      letter-spacing: 0.18em;
      text-transform: uppercase;
      color: var(--porsche-red);
      margin-bottom: 0.25rem;
    }}

    /* Filters Section */
    .filter-panel {{
      background: var(--bg-card);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 1.5rem;
      margin-bottom: 2rem;
      box-shadow: var(--shadow-card);
      position: relative;
    }}

    .filter-panel::before {{
      content: '';
      position: absolute;
      top: 0;
      left: 2rem;
      right: 2rem;
      height: 1px;
      background: linear-gradient(90deg, transparent, rgba(213, 0, 28, 0.4), transparent);
    }}

    .filter-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 1.25rem;
      padding-bottom: 0.75rem;
      border-bottom: 1px solid rgba(255, 255, 255, 0.05);
    }}

    .filter-header-left {{
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }}

    .filter-header-left i {{
      color: var(--porsche-red);
      font-size: 1rem;
    }}

    .filter-header-left h3 {{
      font-size: 0.85rem;
      font-weight: 700;
      letter-spacing: 0.12em;
      text-transform: uppercase;
      color: #ffffff;
    }}

    .filter-count-badge {{
      background: rgba(255, 255, 255, 0.06);
      border: 1px solid var(--border-color);
      padding: 0.25rem 0.75rem;
      border-radius: 20px;
      font-size: 0.75rem;
      color: var(--text-secondary);
    }}

    .filter-count-badge span {{
      color: #ffffff;
      font-weight: 700;
    }}

    .btn-reset-filters {{
      background: transparent;
      border: 1px solid rgba(213, 0, 28, 0.35);
      color: #ff6b6b;
      padding: 0.4rem 0.85rem;
      border-radius: 6px;
      font-size: 0.75rem;
      font-weight: 600;
      letter-spacing: 0.05em;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 0.4rem;
      transition: var(--transition-smooth);
    }}

    .btn-reset-filters:hover {{
      background: rgba(213, 0, 28, 0.12);
      border-color: var(--porsche-red);
      color: #ffffff;
    }}

    /* Model Family Quick Filter Pills */
    .family-pills-label {{
      font-size: 0.72rem;
      font-weight: 600;
      letter-spacing: 0.1em;
      text-transform: uppercase;
      color: var(--text-muted);
      margin-bottom: 0.6rem;
    }}

    .family-pills-group {{
      display: flex;
      flex-wrap: wrap;
      gap: 0.5rem;
      margin-bottom: 1.25rem;
    }}

    .family-pill {{
      background: rgba(255, 255, 255, 0.03);
      border: 1px solid var(--border-color);
      color: var(--text-secondary);
      padding: 0.4rem 0.9rem;
      border-radius: 9999px;
      font-size: 0.75rem;
      font-weight: 600;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      cursor: pointer;
      transition: var(--transition-smooth);
    }}

    .family-pill:hover {{
      border-color: rgba(255, 255, 255, 0.3);
      color: #ffffff;
      background: rgba(255, 255, 255, 0.06);
    }}

    .family-pill.active {{
      background: var(--porsche-red);
      border-color: var(--porsche-red);
      color: #ffffff;
      box-shadow: 0 4px 14px var(--porsche-red-glow);
    }}

    /* Dropdown Filters Grid */
    .filter-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 1rem;
    }}

    .filter-item {{
      display: flex;
      flex-direction: column;
      gap: 0.4rem;
    }}

    .filter-item label {{
      font-size: 0.72rem;
      font-weight: 600;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      color: var(--text-muted);
    }}

    .filter-select, .filter-input {{
      background: var(--bg-input);
      border: 1px solid var(--border-color);
      color: var(--text-main);
      padding: 0.65rem 0.9rem;
      border-radius: 6px;
      font-size: 0.85rem;
      outline: none;
      transition: var(--transition-smooth);
      font-family: inherit;
      width: 100%;
    }}

    .filter-select:focus, .filter-input:focus {{
      border-color: var(--border-focus);
      box-shadow: 0 0 0 2px var(--porsche-red-glow);
    }}

    .filter-select option {{
      background: #181a20;
      color: #ffffff;
    }}

    /* KPI Cards Grid */
    .kpi-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
      gap: 1.25rem;
      margin-bottom: 2.25rem;
    }}

    .kpi-card {{
      background: var(--bg-card);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 1.5rem;
      box-shadow: var(--shadow-card);
      position: relative;
      overflow: hidden;
      transition: var(--transition-smooth);
    }}

    .kpi-card:hover {{
      transform: translateY(-3px);
      border-color: rgba(255, 255, 255, 0.16);
      box-shadow: 0 16px 36px rgba(0, 0, 0, 0.7);
    }}

    .kpi-card.hero-kpi {{
      border-left: 3px solid var(--porsche-red);
    }}

    .kpi-card.gold-kpi {{
      border-left: 3px solid var(--porsche-gold);
    }}

    .kpi-top {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 0.75rem;
    }}

    .kpi-label {{
      font-size: 0.72rem;
      font-weight: 700;
      letter-spacing: 0.12em;
      text-transform: uppercase;
      color: var(--text-muted);
    }}

    .kpi-icon {{
      width: 36px;
      height: 36px;
      border-radius: 8px;
      background: rgba(255, 255, 255, 0.04);
      display: flex;
      align-items: center;
      justify-content: center;
      color: var(--text-secondary);
      font-size: 0.95rem;
    }}

    .hero-kpi .kpi-icon {{
      background: rgba(213, 0, 28, 0.1);
      color: var(--porsche-red);
    }}

    .gold-kpi .kpi-icon {{
      background: rgba(198, 161, 91, 0.1);
      color: var(--porsche-gold);
    }}

    .kpi-value {{
      font-family: 'Montserrat', sans-serif;
      font-size: 1.85rem;
      font-weight: 800;
      letter-spacing: -0.02em;
      color: #ffffff;
      margin-bottom: 0.35rem;
    }}

    .kpi-subtext {{
      font-size: 0.75rem;
      color: var(--text-secondary);
      display: flex;
      align-items: center;
      gap: 0.4rem;
    }}

    .kpi-highlight {{
      color: #34d399;
      font-weight: 600;
    }}

    /* Business Questions Special Section */
    .business-questions-container {{
      margin-bottom: 2.5rem;
    }}

    .section-header {{
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
      margin-bottom: 1.25rem;
      padding-bottom: 0.75rem;
      border-bottom: 1px solid var(--border-color);
    }}

    .section-title-wrap h2 {{
      font-family: 'Montserrat', sans-serif;
      font-size: 1.25rem;
      font-weight: 700;
      letter-spacing: 0.04em;
      text-transform: uppercase;
      color: #ffffff;
      display: flex;
      align-items: center;
      gap: 0.6rem;
    }}

    .section-title-wrap h2::before {{
      content: '';
      display: inline-block;
      width: 4px;
      height: 18px;
      background: var(--porsche-red);
      border-radius: 2px;
    }}

    .section-title-wrap p {{
      color: var(--text-muted);
      font-size: 0.8rem;
      margin-top: 0.2rem;
    }}

    .bq-grid {{
      display: grid;
      grid-template-columns: repeat(12, 1fr);
      gap: 1.5rem;
    }}

    /* Card Question 1: Modelos por Cidade */
    .bq-card-city {{
      grid-column: span 7;
      background: var(--bg-card);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 1.75rem;
      box-shadow: var(--shadow-card);
      display: flex;
      flex-direction: column;
    }}

    /* Card Question 2: Ano Campeão */
    .bq-card-year {{
      grid-column: span 5;
      background: var(--bg-card);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 1.75rem;
      box-shadow: var(--shadow-card);
      display: flex;
      flex-direction: column;
    }}

    /* Full-width Question 3: Insights de Carros Populares por Cidade */
    .bq-card-insights {{
      grid-column: span 12;
      background: var(--bg-card);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 1.75rem;
      box-shadow: var(--shadow-card);
    }}

    .bq-header {{
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      margin-bottom: 1.25rem;
    }}

    .bq-tag {{
      display: inline-block;
      font-size: 0.68rem;
      font-weight: 700;
      letter-spacing: 0.15em;
      text-transform: uppercase;
      color: var(--porsche-red);
      background: rgba(213, 0, 28, 0.08);
      padding: 0.2rem 0.6rem;
      border-radius: 4px;
      margin-bottom: 0.4rem;
    }}

    .bq-title {{
      font-size: 1.05rem;
      font-weight: 700;
      color: #ffffff;
      line-height: 1.35;
    }}

    /* City Model Explorer Widget inside BQ Card */
    .city-explorer-controls {{
      display: flex;
      gap: 1rem;
      margin-bottom: 1.25rem;
    }}

    .city-select-box {{
      flex: 1;
    }}

    .city-stats-summary {{
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 0.75rem;
      margin-bottom: 1.25rem;
    }}

    .city-stat-box {{
      background: rgba(255, 255, 255, 0.02);
      border: 1px solid rgba(255, 255, 255, 0.05);
      border-radius: 8px;
      padding: 0.75rem;
    }}

    .city-stat-label {{
      font-size: 0.68rem;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.06em;
    }}

    .city-stat-val {{
      font-size: 1.15rem;
      font-weight: 700;
      color: #ffffff;
      margin-top: 0.2rem;
    }}

    .city-model-list {{
      display: flex;
      flex-direction: column;
      gap: 0.6rem;
      max-height: 240px;
      overflow-y: auto;
      padding-right: 0.5rem;
    }}

    .city-model-row {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 0.65rem 0.9rem;
      background: rgba(255, 255, 255, 0.02);
      border: 1px solid rgba(255, 255, 255, 0.04);
      border-radius: 6px;
      transition: var(--transition-smooth);
    }}

    .city-model-row:hover {{
      background: rgba(255, 255, 255, 0.05);
      border-color: rgba(255, 255, 255, 0.1);
    }}

    .city-model-name {{
      font-weight: 600;
      font-size: 0.85rem;
      color: #ffffff;
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }}

    .city-model-name i {{
      color: var(--porsche-red);
      font-size: 0.8rem;
    }}

    .city-model-qty {{
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }}

    .qty-badge {{
      background: rgba(213, 0, 28, 0.12);
      border: 1px solid rgba(213, 0, 28, 0.3);
      color: #ff6b6b;
      font-weight: 700;
      font-size: 0.75rem;
      padding: 0.15rem 0.5rem;
      border-radius: 4px;
    }}

    .price-badge {{
      font-size: 0.8rem;
      color: var(--text-secondary);
      font-weight: 500;
    }}

    /* Year Winner Widget */
    .year-winner-highlight {{
      display: flex;
      align-items: center;
      gap: 1.25rem;
      background: linear-gradient(135deg, rgba(213, 0, 28, 0.12) 0%, rgba(255, 255, 255, 0.02) 100%);
      border: 1px solid rgba(213, 0, 28, 0.3);
      border-radius: 10px;
      padding: 1.25rem;
      margin-bottom: 1.25rem;
    }}

    .year-badge-big {{
      font-family: 'Montserrat', sans-serif;
      font-size: 2.6rem;
      font-weight: 900;
      line-height: 1;
      color: #ffffff;
      text-shadow: 0 0 20px var(--porsche-red-glow);
    }}

    .year-winner-info h4 {{
      font-size: 0.95rem;
      font-weight: 700;
      color: #ffffff;
      margin-bottom: 0.2rem;
    }}

    .year-winner-info p {{
      font-size: 0.78rem;
      color: var(--text-secondary);
      line-height: 1.4;
    }}

    .year-distribution-bars {{
      display: flex;
      flex-direction: column;
      gap: 0.65rem;
    }}

    .year-bar-item {{
      display: flex;
      flex-direction: column;
      gap: 0.25rem;
    }}

    .year-bar-labels {{
      display: flex;
      justify-content: space-between;
      font-size: 0.75rem;
    }}

    .year-bar-labels .yr {{
      font-weight: 700;
      color: #ffffff;
    }}

    .year-bar-labels .val {{
      color: var(--text-muted);
    }}

    .progress-track {{
      height: 7px;
      background: rgba(255, 255, 255, 0.05);
      border-radius: 4px;
      overflow: hidden;
    }}

    .progress-fill {{
      height: 100%;
      background: linear-gradient(90deg, #ff4d4f, var(--porsche-red));
      border-radius: 4px;
      transition: width 0.6s ease;
    }}

    /* Insights Cards Grid */
    .insights-subgrid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
      gap: 1.25rem;
      margin-top: 1rem;
    }}

    .insight-box {{
      background: rgba(255, 255, 255, 0.02);
      border: 1px solid rgba(255, 255, 255, 0.06);
      border-radius: 10px;
      padding: 1.25rem;
      position: relative;
      transition: var(--transition-smooth);
    }}

    .insight-box:hover {{
      background: rgba(255, 255, 255, 0.035);
      border-color: rgba(255, 255, 255, 0.12);
      transform: translateY(-2px);
    }}

    .insight-icon-tag {{
      display: flex;
      align-items: center;
      gap: 0.5rem;
      margin-bottom: 0.65rem;
    }}

    .insight-icon-tag i {{
      color: var(--porsche-red);
      font-size: 0.9rem;
    }}

    .insight-icon-tag span {{
      font-size: 0.72rem;
      font-weight: 700;
      letter-spacing: 0.1em;
      text-transform: uppercase;
      color: var(--text-muted);
    }}

    .insight-box h4 {{
      font-size: 0.92rem;
      font-weight: 700;
      color: #ffffff;
      margin-bottom: 0.4rem;
    }}

    .insight-box p {{
      font-size: 0.8rem;
      color: var(--text-secondary);
      line-height: 1.45;
    }}

    /* Main Charts Section */
    .charts-grid {{
      display: grid;
      grid-template-columns: repeat(12, 1fr);
      gap: 1.5rem;
      margin-bottom: 2.5rem;
    }}

    .chart-card {{
      background: var(--bg-card);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 1.5rem;
      box-shadow: var(--shadow-card);
      display: flex;
      flex-direction: column;
    }}

    .chart-card.col-8 {{
      grid-column: span 8;
    }}

    .chart-card.col-4 {{
      grid-column: span 4;
    }}

    .chart-card.col-6 {{
      grid-column: span 6;
    }}

    .chart-card-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 1.25rem;
    }}

    .chart-card-header h3 {{
      font-size: 0.95rem;
      font-weight: 700;
      letter-spacing: 0.04em;
      color: #ffffff;
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }}

    .chart-card-header h3 i {{
      color: var(--porsche-red);
      font-size: 0.85rem;
    }}

    .chart-canvas-container {{
      position: relative;
      flex: 1;
      min-height: 280px;
    }}

    /* Data Table Section */
    .table-section {{
      background: var(--bg-card);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 1.75rem;
      box-shadow: var(--shadow-card);
      margin-bottom: 2.5rem;
    }}

    .table-controls {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 1.25rem;
      flex-wrap: wrap;
      gap: 1rem;
    }}

    .table-search-box {{
      position: relative;
      width: 320px;
    }}

    .table-search-box i {{
      position: absolute;
      left: 1rem;
      top: 50%;
      transform: translateY(-50%);
      color: var(--text-muted);
      font-size: 0.85rem;
    }}

    .table-search-input {{
      width: 100%;
      background: var(--bg-input);
      border: 1px solid var(--border-color);
      color: #ffffff;
      padding: 0.6rem 1rem 0.6rem 2.4rem;
      border-radius: 6px;
      font-size: 0.82rem;
      outline: none;
      transition: var(--transition-smooth);
    }}

    .table-search-input:focus {{
      border-color: var(--border-focus);
      box-shadow: 0 0 0 2px var(--porsche-red-glow);
    }}

    .table-wrapper {{
      overflow-x: auto;
      border: 1px solid var(--border-color);
      border-radius: 8px;
    }}

    table.porsche-table {{
      width: 100%;
      border-collapse: collapse;
      text-align: left;
      font-size: 0.82rem;
    }}

    table.porsche-table th {{
      background: #181b22;
      color: var(--text-secondary);
      font-weight: 700;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      font-size: 0.72rem;
      padding: 0.85rem 1rem;
      border-bottom: 1px solid var(--border-color);
      white-space: nowrap;
      cursor: pointer;
      user-select: none;
      transition: color 0.2s;
    }}

    table.porsche-table th:hover {{
      color: #ffffff;
    }}

    table.porsche-table th i {{
      font-size: 0.65rem;
      margin-left: 0.35rem;
      color: var(--text-muted);
    }}

    table.porsche-table td {{
      padding: 0.8rem 1rem;
      border-bottom: 1px solid rgba(255, 255, 255, 0.04);
      color: #e5e7eb;
      white-space: nowrap;
    }}

    table.porsche-table tr:hover td {{
      background: rgba(255, 255, 255, 0.03);
    }}

    /* Table Badges */
    .badge-pill {{
      display: inline-block;
      padding: 0.2rem 0.55rem;
      border-radius: 4px;
      font-size: 0.72rem;
      font-weight: 600;
      letter-spacing: 0.04em;
    }}

    .badge-delivered {{
      background: rgba(16, 185, 129, 0.12);
      color: #34d399;
      border: 1px solid rgba(16, 185, 129, 0.3);
    }}

    .badge-intransit {{
      background: rgba(59, 130, 246, 0.12);
      color: #60a5fa;
      border: 1px solid rgba(59, 130, 246, 0.3);
    }}

    .badge-pending {{
      background: rgba(245, 158, 11, 0.12);
      color: #fbbf24;
      border: 1px solid rgba(245, 158, 11, 0.3);
    }}

    .badge-cancelled {{
      background: rgba(239, 68, 68, 0.12);
      color: #f87171;
      border: 1px solid rgba(239, 68, 68, 0.3);
    }}

    .badge-invalid {{
      background: rgba(239, 68, 68, 0.15);
      color: #ff5e5e;
      border: 1px solid rgba(239, 68, 68, 0.4);
      font-weight: 700;
    }}

    .badge-model {{
      font-weight: 700;
      color: #ffffff;
    }}

    /* Pagination */
    .table-pagination {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-top: 1.25rem;
      font-size: 0.8rem;
      color: var(--text-muted);
    }}

    .pagination-buttons {{
      display: flex;
      gap: 0.5rem;
    }}

    .page-btn {{
      background: var(--bg-input);
      border: 1px solid var(--border-color);
      color: var(--text-main);
      padding: 0.35rem 0.75rem;
      border-radius: 4px;
      cursor: pointer;
      font-size: 0.75rem;
      transition: var(--transition-smooth);
    }}

    .page-btn:hover:not(:disabled) {{
      border-color: rgba(255, 255, 255, 0.3);
      background: rgba(255, 255, 255, 0.08);
    }}

    .page-btn:disabled {{
      opacity: 0.4;
      cursor: not-allowed;
    }}

    /* Footer */
    footer.porsche-footer {{
      border-top: 1px solid var(--border-color);
      padding: 2.5rem 0;
      text-align: center;
      color: var(--text-muted);
      font-size: 0.78rem;
    }}

    footer.porsche-footer p {{
      margin-bottom: 0.5rem;
    }}

    .footer-links {{
      display: flex;
      justify-content: center;
      gap: 1.5rem;
      margin-top: 1rem;
    }}

    .footer-links a {{
      color: var(--text-secondary);
      text-decoration: none;
      transition: color 0.2s;
    }}

    .footer-links a:hover {{
      color: var(--porsche-red);
    }}

    /* Responsive */
    @media (max-width: 1200px) {{
      .bq-card-city, .bq-card-year {{
        grid-column: span 12;
      }}
      .chart-card.col-8, .chart-card.col-4, .chart-card.col-6 {{
        grid-column: span 12;
      }}
    }}

    @media (max-width: 768px) {{
      .header-inner {{
        flex-direction: column;
        align-items: flex-start;
      }}
      .brand-section {{
        flex-wrap: wrap;
        gap: 1rem;
      }}
      .hero-banner {{
        flex-direction: column;
        align-items: flex-start;
        gap: 1rem;
      }}
      .filter-grid {{
        grid-template-columns: 1fr;
      }}
      .city-stats-summary {{
        grid-template-columns: 1fr;
      }}
    }}
  </style>
</head>
<body>

  <!-- Top Porsche Header -->
  <header class="porsche-header">
    <div class="header-inner">
      <div class="brand-section">
        <div class="porsche-wordmark">PORSCHE</div>
        <div class="brand-divider"></div>
        <div class="platform-title">
          <span class="title">Intelligence & Sales Analytics</span>
          <span class="subtitle">Porsche Sales Sanitization Platform • Schema v1.0</span>
        </div>
      </div>
      <div class="header-status">
        <div class="status-pill">
          <span class="status-dot"></span>
          Base Higienizada
        </div>
        <button class="btn-export" onclick="exportFilteredCSV()">
          <i class="fa-solid fa-file-arrow-down"></i> Exportar CSV
        </button>
      </div>
    </div>
  </header>

  <div class="container">

    <!-- Hero Introduction -->
    <div class="hero-banner">
      <div class="hero-text">
        <span class="hero-tag">Executive Performance Overview</span>
        <h1>Painel de Inteligência de Vendas</h1>
        <p>Análise estratégica de comercialização de veículos Porsche, com dados tratados e auditados pelo <code>porsche_sales_sanitization_schema.md</code>. Métricas de faturamento, popularidade de modelos por praça e tendências anuais.</p>
      </div>
    </div>

    <!-- Interactive Filters Panel -->
    <section class="filter-panel">
      <div class="filter-header">
        <div class="filter-header-left">
          <i class="fa-solid fa-sliders"></i>
          <h3>Filtros Dinâmicos de Inteligência</h3>
          <div class="filter-count-badge">
            Exibindo <span id="filtered-count">100</span> de <span id="total-count">100</span> registros
          </div>
        </div>
        <button class="btn-reset-filters" onclick="resetFilters()">
          <i class="fa-solid fa-rotate-left"></i> Redefinir Filtros
        </button>
      </div>

      <!-- Quick Model Family Pills -->
      <div class="family-pills-label">Família de Veículos (Seleção Rápida):</div>
      <div class="family-pills-group" id="family-pills-container">
        <button class="family-pill active" onclick="setFamilyFilter('ALL')">Todos os Modelos</button>
        <button class="family-pill" onclick="setFamilyFilter('911')">911</button>
        <button class="family-pill" onclick="setFamilyFilter('718')">718 (Cayman / Boxster)</button>
        <button class="family-pill" onclick="setFamilyFilter('Cayenne')">Cayenne</button>
        <button class="family-pill" onclick="setFamilyFilter('Macan')">Macan</button>
        <button class="family-pill" onclick="setFamilyFilter('Panamera')">Panamera</button>
        <button class="family-pill" onclick="setFamilyFilter('Taycan')">Taycan (Elétrico)</button>
      </div>

      <!-- Detail Select Dropdowns -->
      <div class="filter-grid">
        <div class="filter-item">
          <label><i class="fa-solid fa-car-side"></i> Modelo Específico</label>
          <select id="filter-model" class="filter-select" onchange="applyFilters()">
            <option value="ALL">Todos os modelos específicos</option>
          </select>
        </div>

        <div class="filter-item">
          <label><i class="fa-solid fa-location-dot"></i> Cidade</label>
          <select id="filter-city" class="filter-select" onchange="applyFilters()">
            <option value="ALL">Todas as cidades</option>
          </select>
        </div>

        <div class="filter-item">
          <label><i class="fa-solid fa-calendar"></i> Ano do Modelo</label>
          <select id="filter-year" class="filter-select" onchange="applyFilters()">
            <option value="ALL">Todos os anos</option>
          </select>
        </div>

        <div class="filter-item">
          <label><i class="fa-solid fa-credit-card"></i> Método de Pagamento</label>
          <select id="filter-payment" class="filter-select" onchange="applyFilters()">
            <option value="ALL">Todos os métodos</option>
          </select>
        </div>
      </div>
    </section>

    <!-- Top Executive KPIs -->
    <div class="kpi-grid">
      <div class="kpi-card hero-kpi">
        <div class="kpi-top">
          <span class="kpi-label">Faturamento Total</span>
          <div class="kpi-icon"><i class="fa-solid fa-dollar-sign"></i></div>
        </div>
        <div class="kpi-value" id="kpi-revenue">$0</div>
        <div class="kpi-subtext">
          Ticket Médio: <span class="kpi-highlight" id="kpi-avg-price">$0</span>
        </div>
      </div>

      <div class="kpi-card">
        <div class="kpi-top">
          <span class="kpi-label">Volume de Vendas</span>
          <div class="kpi-icon"><i class="fa-solid fa-car"></i></div>
        </div>
        <div class="kpi-value" id="kpi-units">0</div>
        <div class="kpi-subtext">
          Representação: <span class="kpi-highlight" id="kpi-units-pct">100%</span> da frota total
        </div>
      </div>

      <div class="kpi-card gold-kpi">
        <div class="kpi-top">
          <span class="kpi-label">Ano Campeão de Vendas</span>
          <div class="kpi-icon"><i class="fa-solid fa-trophy"></i></div>
        </div>
        <div class="kpi-value" id="kpi-top-year">-</div>
        <div class="kpi-subtext">
          Liderança com <span class="kpi-highlight" id="kpi-top-year-count">0</span> veículos vendidos
        </div>
      </div>

      <div class="kpi-card">
        <div class="kpi-top">
          <span class="kpi-label">Modelo Líder em Volume</span>
          <div class="kpi-icon"><i class="fa-solid fa-crown"></i></div>
        </div>
        <div class="kpi-value" style="font-size: 1.35rem; line-height: 1.25;" id="kpi-top-model">-</div>
        <div class="kpi-subtext">
          Total de <span class="kpi-highlight" id="kpi-top-model-units">0</span> un. no filtro
        </div>
      </div>

      <div class="kpi-card">
        <div class="kpi-top">
          <span class="kpi-label">Quilometragem Média</span>
          <div class="kpi-icon"><i class="fa-solid fa-gauge-high"></i></div>
        </div>
        <div class="kpi-value" id="kpi-avg-mileage">0 mi</div>
        <div class="kpi-subtext">
          Mix de seminovos e 0km higienizados
        </div>
      </div>
    </div>

    <!-- Strategic Business Questions Answered Section -->
    <section class="business-questions-container">
      <div class="section-header">
        <div class="section-title-wrap">
          <h2>Perguntas de Negócios & Respostas Estratégicas</h2>
          <p>Módulos analíticos dedicados a responder diretamente às perguntas centrais da diretoria executiva.</p>
        </div>
      </div>

      <div class="bq-grid">

        <!-- Question 1: Principais Modelos por Cidade -->
        <div class="bq-card-city">
          <div class="bq-header">
            <div>
              <span class="bq-tag"><i class="fa-solid fa-city"></i> Pergunta 1</span>
              <h3 class="bq-title">Quais os principais modelos de carros vendidos por cidade?</h3>
            </div>
          </div>
          
          <div class="city-explorer-controls">
            <div class="city-select-box">
              <label style="font-size:0.72rem; text-transform:uppercase; color:var(--text-muted); font-weight:600; display:block; margin-bottom:0.3rem;">
                Selecione a Praça / Cidade:
              </label>
              <select id="bq-city-selector" class="filter-select" onchange="renderCityExplorer()">
                <!-- Populated dynamically -->
              </select>
            </div>
          </div>

          <div class="city-stats-summary" id="city-stats-summary">
            <!-- Dynamically populated -->
          </div>

          <div style="font-size:0.72rem; text-transform:uppercase; color:var(--text-muted); font-weight:700; margin-bottom:0.5rem;">
            Modelos Vendidos Nesta Cidade:
          </div>
          <div class="city-model-list" id="city-model-list">
            <!-- Dynamically populated -->
          </div>
        </div>

        <!-- Question 2: Qual o ano de modelo de carro que mais saiu no período? -->
        <div class="bq-card-year">
          <div class="bq-header">
            <div>
              <span class="bq-tag"><i class="fa-solid fa-calendar-check"></i> Pergunta 2</span>
              <h3 class="bq-title">Qual o ano de modelo de carro que mais saiu em um período?</h3>
            </div>
          </div>

          <div class="year-winner-highlight" id="year-winner-highlight">
            <div class="year-badge-big" id="winner-year-display">2024</div>
            <div class="year-winner-info">
              <h4 id="winner-year-title">Modelo Ano 2024 é o Grande Líder</h4>
              <p id="winner-year-desc">
                Com a maior concentração de emplacamentos e vendas recentes, consolidando a preferência por tecnologia de ponta e modelos do ano corrente.
              </p>
            </div>
          </div>

          <div style="font-size:0.72rem; text-transform:uppercase; color:var(--text-muted); font-weight:700; margin-bottom:0.6rem;">
            Distribuição de Volume por Ano de Fabricação / Modelo:
          </div>
          <div class="year-distribution-bars" id="year-distribution-bars">
            <!-- Dynamically populated -->
          </div>
        </div>

        <!-- Question 3: Insights de Carros Populares por Cidade -->
        <div class="bq-card-insights">
          <div class="bq-header">
            <div>
              <span class="bq-tag"><i class="fa-solid fa-lightbulb"></i> Pergunta 3 & Inteligência de Mercado</span>
              <h3 class="bq-title">Insights de Carros Populares de Vendas com Base nos Dados em Cada Cidade</h3>
            </div>
          </div>

          <div class="insights-subgrid">
            <div class="insight-box">
              <div class="insight-icon-tag">
                <i class="fa-solid fa-mountain-sun"></i>
                <span>SUVs Urbanos & Versatilidade</span>
              </div>
              <h4>Dominância de Macan e Cayenne em Polos Urbanos</h4>
              <p>
                Em cidades com tráfego intenso e perfis corporativos como <strong>Atlanta, Phoenix, Denver e Austin</strong>, a linha SUV (especialmente <em>Macan T, Macan Electric e Cayenne Coupe</em>) lidera o volume, oferecendo versatilidade premium com alto valor residual.
              </p>
            </div>

            <div class="insight-box">
              <div class="insight-icon-tag">
                <i class="fa-solid fa-bolt"></i>
                <span>Eletrificação & Vanguarda</span>
              </div>
              <h4>Adoção Acelerada de Taycan & Macan Electric</h4>
              <p>
                Praças da Costa Oeste e centros tecnológicos como <strong>Los Angeles, San Diego, Seattle e San Jose</strong> concentram forte saída de modelos 100% elétricos (<em>Taycan 4S e Macan Electric</em>), fortemente impulsionados por métodos de <em>Lease</em> e <em>Wire Transfer</em>.
              </p>
            </div>

            <div class="insight-box">
              <div class="insight-icon-tag">
                <i class="fa-solid fa-flag-checkered"></i>
                <span>Herança Esportiva & Exclusividade</span>
              </div>
              <h4>Ícones 911 e 718 Lideram os Maiores Tickets</h4>
              <p>
                Em metrópoles consolidadas como <strong>Boston, New York, Dallas e Chicago</strong>, as versões <em>911 Turbo S, 911 Carrera GTS e 718 Cayman GT4 RS</em> atingem os maiores tíquetes médios (acima de <strong>$200.000</strong>), comercializados com pagamento à vista ou transferência direta.
              </p>
            </div>

            <div class="insight-box">
              <div class="insight-icon-tag">
                <i class="fa-solid fa-wallet"></i>
                <span>Modalidade de Fechamento</span>
              </div>
              <h4>Perfil Financeiro: Wire Transfer & Financing</h4>
              <p>
                Mais de <strong>50%</strong> do faturamento é liquidado via <em>Wire Transfer</em> (transferência bancária internacional/direta) e <em>Financing</em> estruturado, evidenciando alto poder aquisitivo e negociações de alto calibre.
              </p>
            </div>
          </div>
        </div>

      </div>
    </section>

    <!-- Main Analytics Charts -->
    <div class="charts-grid">
      <!-- Chart 1: Vendas por Ano do Modelo -->
      <div class="chart-card col-6">
        <div class="chart-card-header">
          <h3><i class="fa-solid fa-chart-column"></i> Volume por Ano de Modelo (Model Year)</h3>
          <span style="font-size:0.72rem; color:var(--text-muted); text-transform:uppercase;">Volume de Vendas</span>
        </div>
        <div class="chart-canvas-container">
          <canvas id="chart-years"></canvas>
        </div>
      </div>

      <!-- Chart 2: Top Modelos mais Vendidos -->
      <div class="chart-card col-6">
        <div class="chart-card-header">
          <h3><i class="fa-solid fa-ranking-star"></i> Top 8 Modelos Mais Vendidos (Volume)</h3>
          <span style="font-size:0.72rem; color:var(--text-muted); text-transform:uppercase;">Ranking Unidades</span>
        </div>
        <div class="chart-canvas-container">
          <canvas id="chart-models"></canvas>
        </div>
      </div>

      <!-- Chart 3: Métodos de Pagamento -->
      <div class="chart-card col-6">
        <div class="chart-card-header">
          <h3><i class="fa-solid fa-pie-chart"></i> Distribuição por Método de Pagamento</h3>
          <span style="font-size:0.72rem; color:var(--text-muted); text-transform:uppercase;">Fatia de Mercado</span>
        </div>
        <div class="chart-canvas-container">
          <canvas id="chart-payments"></canvas>
        </div>
      </div>

      <!-- Chart 4: Top Cidades por Volume -->
      <div class="chart-card col-6">
        <div class="chart-card-header">
          <h3><i class="fa-solid fa-map-location-dot"></i> Principais Cidades por Volume de Vendas</h3>
          <span style="font-size:0.72rem; color:var(--text-muted); text-transform:uppercase;">Top Mercados</span>
        </div>
        <div class="chart-canvas-container">
          <canvas id="chart-cities"></canvas>
        </div>
      </div>
    </div>

    <!-- Transactional Table Section -->
    <section class="table-section">
      <div class="table-controls">
        <div class="section-title-wrap">
          <h2 style="font-size:1.15rem;">Detalhamento de Vendas Higienizadas</h2>
          <p>Tabela analítica com todos os registros após o pipeline de padronização.</p>
        </div>
        <div class="table-search-box">
          <i class="fa-solid fa-magnifying-glass"></i>
          <input type="text" id="table-search" class="table-search-input" placeholder="Buscar cliente, modelo, cidade..." oninput="onSearchInput()">
        </div>
      </div>

      <div class="table-wrapper">
        <table class="porsche-table">
          <thead>
            <tr>
              <th onclick="sortTable('sale_id')">ID <i class="fa-solid fa-sort"></i></th>
              <th onclick="sortTable('SaleDateSanitized')">Data Normalizada <i class="fa-solid fa-sort"></i></th>
              <th onclick="sortTable('customer_name')">Cliente <i class="fa-solid fa-sort"></i></th>
              <th onclick="sortTable('PorscheModelSanitized')">Modelo Porsche <i class="fa-solid fa-sort"></i></th>
              <th onclick="sortTable('ModelYearSanitized')">Ano <i class="fa-solid fa-sort"></i></th>
              <th onclick="sortTable('SalesPriceSanitized')">Preço (USD) <i class="fa-solid fa-sort"></i></th>
              <th onclick="sortTable('VehicleMileageSanitized')">Milhas <i class="fa-solid fa-sort"></i></th>
              <th onclick="sortTable('PayMethodSanitized')">Pagamento <i class="fa-solid fa-sort"></i></th>
              <th onclick="sortTable('CitySanitized')">Cidade <i class="fa-solid fa-sort"></i></th>
              <th onclick="sortTable('StateSanitized')">UF <i class="fa-solid fa-sort"></i></th>
              <th onclick="sortTable('DeliveryStatusSanitized')">Status Entrega <i class="fa-solid fa-sort"></i></th>
            </tr>
          </thead>
          <tbody id="table-body">
            <!-- Injected via JavaScript -->
          </tbody>
        </table>
      </div>

      <div class="table-pagination">
        <div>
          Mostrando página <span id="current-page-display">1</span> de <span id="total-pages-display">1</span>
        </div>
        <div class="pagination-buttons">
          <button class="page-btn" id="btn-prev-page" onclick="changePage(-1)"><i class="fa-solid fa-chevron-left"></i> Anterior</button>
          <button class="page-btn" id="btn-next-page" onclick="changePage(1)">Próximo <i class="fa-solid fa-chevron-right"></i></button>
        </div>
      </div>
    </section>

  </div>

  <!-- Porsche Footer -->
  <footer class="porsche-footer">
    <div class="porsche-wordmark" style="font-size:1.1rem; margin-bottom:0.75rem;">PORSCHE</div>
    <p>© 2026 Dr. Ing. h.c. F. Porsche AG. Dados operacionais sanitizados sob conformidade do schema de vendas.</p>
    <div class="footer-links">
      <a href="https://www.porsche.com/brazil/pt/" target="_blank">Porsche Brasil Oficial</a>
      <a href="#top" onclick="window.scrollTo({{top:0, behavior:'smooth'}})">Voltar ao Topo</a>
    </div>
  </footer>

  <!-- EMBEDDED RAW DATA & APPLICATION LOGIC -->
  <script>
    // RAW SANITIZED DATASET
    const RAW_DATA = {json_data_str};

    // State
    let filteredData = [...RAW_DATA];
    let selectedFamily = 'ALL';
    let searchQuery = '';
    let currentPage = 1;
    const pageSize = 12;
    let sortColumn = 'sale_id';
    let sortAsc = true;

    // Charts instances
    let chartYears = null;
    let chartModels = null;
    let chartPayments = null;
    let chartCities = null;

    // Initialize Dashboard on load
    window.addEventListener('DOMContentLoaded', () => {{
      populateDropdownOptions();
      initCitySelector();
      applyFilters();
    }});

    // Populate Dropdown Menus
    function populateDropdownOptions() {{
      const models = [...new Set(RAW_DATA.map(d => d.PorscheModelSanitized))].sort();
      const cities = [...new Set(RAW_DATA.map(d => d.CitySanitized))].sort();
      const years = [...new Set(RAW_DATA.map(d => d.ModelYearSanitized))].sort((a,b) => a - b);
      const payments = [...new Set(RAW_DATA.map(d => d.PayMethodSanitized))].sort();

      const modelSelect = document.getElementById('filter-model');
      models.forEach(m => {{
        const opt = document.createElement('option');
        opt.value = m;
        opt.textContent = m;
        modelSelect.appendChild(opt);
      }});

      const citySelect = document.getElementById('filter-city');
      cities.forEach(c => {{
        const opt = document.createElement('option');
        opt.value = c;
        opt.textContent = c;
        citySelect.appendChild(opt);
      }});

      const yearSelect = document.getElementById('filter-year');
      years.forEach(y => {{
        const opt = document.createElement('option');
        opt.value = y;
        opt.textContent = y;
        yearSelect.appendChild(opt);
      }});

      const paySelect = document.getElementById('filter-payment');
      payments.forEach(p => {{
        const opt = document.createElement('option');
        opt.value = p;
        opt.textContent = p;
        paySelect.appendChild(opt);
      }});
    }}

    // Init City Selector for Question 1
    function initCitySelector() {{
      const citySelector = document.getElementById('bq-city-selector');
      citySelector.innerHTML = '';
      
      // Calculate top cities in full dataset first
      const cityCounts = {{}};
      RAW_DATA.forEach(d => {{
        cityCounts[d.CitySanitized] = (cityCounts[d.CitySanitized] || 0) + 1;
      }});
      
      const sortedCities = Object.keys(cityCounts).sort((a,b) => cityCounts[b] - cityCounts[a] || a.localeCompare(b));

      sortedCities.forEach(city => {{
        const opt = document.createElement('option');
        opt.value = city;
        opt.textContent = `${{city}} (${{cityCounts[city]}} vendas)`;
        citySelector.appendChild(opt);
      }});

      renderCityExplorer();
    }}

    // Family quick filter
    function setFamilyFilter(family) {{
      selectedFamily = family;
      
      // Update pills active class
      const pills = document.querySelectorAll('.family-pill');
      pills.forEach(p => p.classList.remove('active'));
      
      event.target.classList.add('active');
      applyFilters();
    }}

    // Apply Filters Logic
    function applyFilters() {{
      const selModel = document.getElementById('filter-model').value;
      const selCity = document.getElementById('filter-city').value;
      const selYear = document.getElementById('filter-year').value;
      const selPayment = document.getElementById('filter-payment').value;

      filteredData = RAW_DATA.filter(item => {{
        // Family match
        if (selectedFamily !== 'ALL') {{
          if (selectedFamily === '718') {{
            if (!item.PorscheModelSanitized.includes('718') && !item.PorscheModelSanitized.includes('Cayman') && !item.PorscheModelSanitized.includes('Boxster') && !item.PorscheModelSanitized.includes('Spyder')) {{
              return false;
            }}
          }} else if (!item.PorscheModelSanitized.includes(selectedFamily)) {{
            return false;
          }}
        }}

        // Specific Model
        if (selModel !== 'ALL' && item.PorscheModelSanitized !== selModel) {{
          return false;
        }}

        // City
        if (selCity !== 'ALL' && item.CitySanitized !== selCity) {{
          return false;
        }}

        // Year
        if (selYear !== 'ALL' && String(item.ModelYearSanitized) !== String(selYear)) {{
          return false;
        }}

        // Payment
        if (selPayment !== 'ALL' && item.PayMethodSanitized !== selPayment) {{
          return false;
        }}

        // Search text
        if (searchQuery.trim() !== '') {{
          const q = searchQuery.toLowerCase();
          const match = 
            (item.customer_name && item.customer_name.toLowerCase().includes(q)) ||
            (item.PorscheModelSanitized && item.PorscheModelSanitized.toLowerCase().includes(q)) ||
            (item.CitySanitized && item.CitySanitized.toLowerCase().includes(q)) ||
            (item.salesperson && item.salesperson.toLowerCase().includes(q)) ||
            (String(item.sale_id).includes(q));
          if (!match) return false;
        }}

        return true;
      }});

      // Update counters
      document.getElementById('filtered-count').textContent = filteredData.length;
      document.getElementById('total-count').textContent = RAW_DATA.length;

      // Update KPIs
      updateKPIs();

      // Update Business Question components
      renderCityExplorer();
      renderYearQuestion();

      // Update Charts
      updateCharts();

      // Reset table page and render
      currentPage = 1;
      renderTable();
    }}

    // Reset All Filters
    function resetFilters() {{
      selectedFamily = 'ALL';
      document.querySelectorAll('.family-pill').forEach((p, idx) => {{
        if (idx === 0) p.classList.add('active');
        else p.classList.remove('active');
      }});

      document.getElementById('filter-model').value = 'ALL';
      document.getElementById('filter-city').value = 'ALL';
      document.getElementById('filter-year').value = 'ALL';
      document.getElementById('filter-payment').value = 'ALL';
      document.getElementById('table-search').value = '';
      searchQuery = '';

      applyFilters();
    }}

    // Format Currency
    function formatUSD(num) {{
      return new Intl.NumberFormat('en-US', {{ style: 'currency', currency: 'USD', maximumFractionDigits: 0 }}).format(num);
    }}

    // Update KPI Cards
    function updateKPIs() {{
      const totalRev = filteredData.reduce((acc, cur) => acc + (Number(cur.SalesPriceSanitized) || 0), 0);
      const units = filteredData.length;
      const avgPrice = units > 0 ? (totalRev / units) : 0;
      
      const totalMiles = filteredData.reduce((acc, cur) => acc + (Number(cur.VehicleMileageSanitized) || 0), 0);
      const avgMiles = units > 0 ? Math.round(totalMiles / units) : 0;

      // Year frequency
      const yearCounts = {{}};
      const modelCounts = {{}};
      filteredData.forEach(d => {{
        yearCounts[d.ModelYearSanitized] = (yearCounts[d.ModelYearSanitized] || 0) + 1;
        modelCounts[d.PorscheModelSanitized] = (modelCounts[d.PorscheModelSanitized] || 0) + 1;
      }});

      let topYear = '-';
      let topYearCount = 0;
      for (const [y, cnt] of Object.entries(yearCounts)) {{
        if (cnt > topYearCount) {{
          topYearCount = cnt;
          topYear = y;
        }}
      }}

      let topModel = '-';
      let topModelCount = 0;
      for (const [m, cnt] of Object.entries(modelCounts)) {{
        if (cnt > topModelCount) {{
          topModelCount = cnt;
          topModel = m;
        }}
      }}

      document.getElementById('kpi-revenue').textContent = formatUSD(totalRev);
      document.getElementById('kpi-avg-price').textContent = formatUSD(avgPrice);
      document.getElementById('kpi-units').textContent = units;
      document.getElementById('kpi-units-pct').textContent = RAW_DATA.length > 0 ? Math.round((units / RAW_DATA.length) * 100) + '%' : '0%';
      document.getElementById('kpi-top-year').textContent = topYear;
      document.getElementById('kpi-top-year-count').textContent = topYearCount;
      document.getElementById('kpi-top-model').textContent = topModel;
      document.getElementById('kpi-top-model-units').textContent = topModelCount;
      document.getElementById('kpi-avg-mileage').textContent = avgMiles.toLocaleString('en-US') + ' mi';
    }}

    // Render Question 1: City Explorer Widget
    function renderCityExplorer() {{
      const citySelector = document.getElementById('bq-city-selector');
      if (!citySelector) return;
      
      const city = citySelector.value;
      const citySales = RAW_DATA.filter(d => d.CitySanitized === city);

      const totalCityRev = citySales.reduce((acc, cur) => acc + (Number(cur.SalesPriceSanitized) || 0), 0);
      const cityUnits = citySales.length;
      const avgCityPrice = cityUnits > 0 ? (totalCityRev / cityUnits) : 0;

      // Stats summary boxes
      const statsContainer = document.getElementById('city-stats-summary');
      statsContainer.innerHTML = `
        <div class="city-stat-box">
          <div class="city-stat-label">Vendas na Cidade</div>
          <div class="city-stat-val" style="color:var(--porsche-red);">${{cityUnits}} un.</div>
        </div>
        <div class="city-stat-box">
          <div class="city-stat-label">Faturamento Total</div>
          <div class="city-stat-val">${{formatUSD(totalCityRev)}}</div>
        </div>
        <div class="city-stat-box">
          <div class="city-stat-label">Ticket Médio</div>
          <div class="city-stat-val">${{formatUSD(avgCityPrice)}}</div>
        </div>
      `;

      // Group models in this city
      const modelCounts = {{}};
      const modelPrices = {{}};
      citySales.forEach(d => {{
        const m = d.PorscheModelSanitized;
        modelCounts[m] = (modelCounts[m] || 0) + 1;
        modelPrices[m] = (modelPrices[m] || 0) + Number(d.SalesPriceSanitized);
      }});

      const sortedModels = Object.keys(modelCounts).sort((a,b) => modelCounts[b] - modelCounts[a]);

      const listContainer = document.getElementById('city-model-list');
      listContainer.innerHTML = '';

      if (sortedModels.length === 0) {{
        listContainer.innerHTML = '<div style="color:var(--text-muted); font-size:0.8rem; padding:1rem;">Nenhum registro para esta cidade.</div>';
        return;
      }}

      sortedModels.forEach(m => {{
        const qty = modelCounts[m];
        const avg = modelPrices[m] / qty;
        const row = document.createElement('div');
        row.className = 'city-model-row';
        row.innerHTML = `
          <div class="city-model-name">
            <i class="fa-solid fa-chevron-right"></i>
            <span>${{m}}</span>
          </div>
          <div class="city-model-qty">
            <span class="price-badge">${{formatUSD(avg)}}</span>
            <span class="qty-badge">${{qty}} ${{qty > 1 ? 'vendas' : 'venda'}}</span>
          </div>
        `;
        listContainer.appendChild(row);
      }});
    }}

    // Render Question 2: Year Winner and Distribution
    function renderYearQuestion() {{
      const yearCounts = {{}};
      filteredData.forEach(d => {{
        yearCounts[d.ModelYearSanitized] = (yearCounts[d.ModelYearSanitized] || 0) + 1;
      }});

      const sortedYears = Object.keys(yearCounts).sort((a,b) => Number(a) - Number(b));
      let maxYear = '2024';
      let maxCount = 0;
      for (const [y, cnt] of Object.entries(yearCounts)) {{
        if (cnt > maxCount) {{
          maxCount = cnt;
          maxYear = y;
        }}
      }}

      document.getElementById('winner-year-display').textContent = maxYear;
      document.getElementById('winner-year-title').textContent = `Ano ${{maxYear}} Lidera com ${{maxCount}} Unidades`;

      const barsContainer = document.getElementById('year-distribution-bars');
      barsContainer.innerHTML = '';

      const totalInFilter = filteredData.length || 1;

      sortedYears.forEach(yr => {{
        const cnt = yearCounts[yr];
        const pct = Math.round((cnt / totalInFilter) * 100);
        const item = document.createElement('div');
        item.className = 'year-bar-item';
        item.innerHTML = `
          <div class="year-bar-labels">
            <span class="yr">Ano Modelo ${{yr}}</span>
            <span class="val">${{cnt}} vendas (${{pct}}%)</span>
          </div>
          <div class="progress-track">
            <div class="progress-fill" style="width: ${{pct}}%;"></div>
          </div>
        `;
        barsContainer.appendChild(item);
      }});
    }}

    // Update Interactive Charts
    function updateCharts() {{
      // 1. Chart Years
      const yearMap = {{}};
      filteredData.forEach(d => {{
        yearMap[d.ModelYearSanitized] = (yearMap[d.ModelYearSanitized] || 0) + 1;
      }});
      const yearLabels = Object.keys(yearMap).sort((a,b) => a - b);
      const yearVals = yearLabels.map(y => yearMap[y]);

      if (chartYears) chartYears.destroy();
      const ctxYears = document.getElementById('chart-years').getContext('2d');
      chartYears = new Chart(ctxYears, {{
        type: 'bar',
        data: {{
          labels: yearLabels,
          datasets: [{{
            label: 'Veículos Vendidos',
            data: yearVals,
            backgroundColor: yearLabels.map(y => y === '2024' ? '#d5001c' : 'rgba(255, 255, 255, 0.2)'),
            borderRadius: 6,
            hoverBackgroundColor: '#d5001c'
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{
            legend: {{ display: false }},
            tooltip: {{
              backgroundColor: '#181b21',
              titleColor: '#fff',
              bodyColor: '#e5e7eb',
              borderColor: 'rgba(255,255,255,0.1)',
              borderWidth: 1
            }}
          }},
          scales: {{
            x: {{
              grid: {{ display: false }},
              ticks: {{ color: '#9ca3af' }}
            }},
            y: {{
              grid: {{ color: 'rgba(255, 255, 255, 0.05)' }},
              ticks: {{ color: '#9ca3af', stepSize: 5 }}
            }}
          }}
        }}
      }});

      // 2. Chart Top Models
      const modelMap = {{}};
      filteredData.forEach(d => {{
        modelMap[d.PorscheModelSanitized] = (modelMap[d.PorscheModelSanitized] || 0) + 1;
      }});
      const sortedModels = Object.entries(modelMap).sort((a,b) => b[1] - a[1]).slice(0, 8);
      const modelLabels = sortedModels.map(x => x[0]);
      const modelVals = sortedModels.map(x => x[1]);

      if (chartModels) chartModels.destroy();
      const ctxModels = document.getElementById('chart-models').getContext('2d');
      chartModels = new Chart(ctxModels, {{
        type: 'bar',
        data: {{
          labels: modelLabels,
          datasets: [{{
            label: 'Vendas',
            data: modelVals,
            backgroundColor: 'rgba(213, 0, 28, 0.75)',
            borderRadius: 4,
            hoverBackgroundColor: '#d5001c'
          }}]
        }},
        options: {{
          indexAxis: 'y',
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{
            legend: {{ display: false }},
            tooltip: {{
              backgroundColor: '#181b21',
              borderColor: 'rgba(255,255,255,0.1)',
              borderWidth: 1
            }}
          }},
          scales: {{
            x: {{
              grid: {{ color: 'rgba(255, 255, 255, 0.05)' }},
              ticks: {{ color: '#9ca3af', stepSize: 1 }}
            }},
            y: {{
              grid: {{ display: false }},
              ticks: {{ color: '#e5e7eb', font: {{ size: 11 }} }}
            }}
          }}
        }}
      }});

      // 3. Chart Payments
      const payMap = {{}};
      filteredData.forEach(d => {{
        payMap[d.PayMethodSanitized] = (payMap[d.PayMethodSanitized] || 0) + 1;
      }});
      const payLabels = Object.keys(payMap);
      const payVals = payLabels.map(p => payMap[p]);

      const porscheColors = [
        '#d5001c', '#c6a15b', '#3b82f6', '#10b981', '#8b5cf6', 
        '#ec4899', '#f97316', '#64748b', '#06b6d4'
      ];

      if (chartPayments) chartPayments.destroy();
      const ctxPay = document.getElementById('chart-payments').getContext('2d');
      chartPayments = new Chart(ctxPay, {{
        type: 'doughnut',
        data: {{
          labels: payLabels,
          datasets: [{{
            data: payVals,
            backgroundColor: porscheColors.slice(0, payLabels.length),
            borderColor: '#121418',
            borderWidth: 3
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          cutout: '68%',
          plugins: {{
            legend: {{
              position: 'right',
              labels: {{ color: '#d1d5db', font: {{ size: 11 }}, boxWidth: 12, padding: 12 }}
            }}
          }}
        }}
      }});

      // 4. Chart Cities
      const cityMap = {{}};
      filteredData.forEach(d => {{
        cityMap[d.CitySanitized] = (cityMap[d.CitySanitized] || 0) + 1;
      }});
      const topCities = Object.entries(cityMap).sort((a,b) => b[1] - a[1]).slice(0, 8);
      const cityLabels = topCities.map(x => x[0]);
      const cityVals = topCities.map(x => x[1]);

      if (chartCities) chartCities.destroy();
      const ctxCities = document.getElementById('chart-cities').getContext('2d');
      chartCities = new Chart(ctxCities, {{
        type: 'bar',
        data: {{
          labels: cityLabels,
          datasets: [{{
            label: 'Vendas Registradas',
            data: cityVals,
            backgroundColor: 'rgba(198, 161, 91, 0.75)',
            borderRadius: 6,
            hoverBackgroundColor: '#c6a15b'
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{
            legend: {{ display: false }}
          }},
          scales: {{
            x: {{
              grid: {{ display: false }},
              ticks: {{ color: '#9ca3af' }}
            }},
            y: {{
              grid: {{ color: 'rgba(255, 255, 255, 0.05)' }},
              ticks: {{ color: '#9ca3af', stepSize: 1 }}
            }}
          }}
        }}
      }});
    }}

    // Table Rendering & Pagination
    function renderTable() {{
      const tbody = document.getElementById('table-body');
      tbody.innerHTML = '';

      // Sort
      const sorted = [...filteredData].sort((a, b) => {{
        let vA = a[sortColumn];
        let vB = b[sortColumn];
        if (typeof vA === 'string') vA = vA.toLowerCase();
        if (typeof vB === 'string') vB = vB.toLowerCase();
        if (vA < vB) return sortAsc ? -1 : 1;
        if (vA > vB) return sortAsc ? 1 : -1;
        return 0;
      }});

      const totalPages = Math.ceil(sorted.length / pageSize) || 1;
      if (currentPage > totalPages) currentPage = totalPages;

      document.getElementById('current-page-display').textContent = currentPage;
      document.getElementById('total-pages-display').textContent = totalPages;
      document.getElementById('btn-prev-page').disabled = (currentPage === 1);
      document.getElementById('btn-next-page').disabled = (currentPage === totalPages);

      const start = (currentPage - 1) * pageSize;
      const end = start + pageSize;
      const pageSlice = sorted.slice(start, end);

      if (pageSlice.length === 0) {{
        tbody.innerHTML = '<tr><td colspan="11" style="text-align:center; padding:2rem; color:var(--text-muted);">Nenhum registro encontrado para estes filtros.</td></tr>';
        return;
      }}

      pageSlice.forEach(row => {{
        const tr = document.createElement('tr');
        
        // Status Badge class
        let statusClass = 'badge-pending';
        const st = (row.DeliveryStatusSanitized || '').toLowerCase();
        if (st.includes('delivered')) statusClass = 'badge-delivered';
        else if (st.includes('transit')) statusClass = 'badge-intransit';
        else if (st.includes('cancel')) statusClass = 'badge-cancelled';

        // Date display
        const dateHtml = (row.SaleDateSanitized === 'INVALID') 
          ? '<span class="badge-pill badge-invalid">INVALID</span>' 
          : row.SaleDateSanitized;

        tr.innerHTML = `
          <td><strong>#${{row.sale_id}}</strong></td>
          <td>${{dateHtml}}</td>
          <td>${{row.customer_name || '-'}}</td>
          <td><span class="badge-model">${{row.PorscheModelSanitized}}</span></td>
          <td>${{row.ModelYearSanitized}}</td>
          <td style="font-weight:600; color:#ffffff;">${{formatUSD(row.SalesPriceSanitized)}}</td>
          <td>${{Number(row.VehicleMileageSanitized).toLocaleString('en-US')}} mi</td>
          <td>${{row.PayMethodSanitized}}</td>
          <td>${{row.CitySanitized}}</td>
          <td><strong>${{row.StateSanitized}}</strong></td>
          <td><span class="badge-pill ${{statusClass}}">${{row.DeliveryStatusSanitized}}</span></td>
        `;
        tbody.appendChild(tr);
      }});
    }}

    function changePage(delta) {{
      currentPage += delta;
      renderTable();
    }}

    function sortTable(col) {{
      if (sortColumn === col) {{
        sortAsc = !sortAsc;
      }} else {{
        sortColumn = col;
        sortAsc = true;
      }}
      renderTable();
    }}

    function onSearchInput() {{
      searchQuery = document.getElementById('table-search').value;
      applyFilters();
    }}

    // Export Filtered View to CSV
    function exportFilteredCSV() {{
      if (!filteredData || filteredData.length === 0) {{
        alert('Nenhum dado para exportar.');
        return;
      }}

      const headers = [
        'Sale ID', 'Data Normalizada', 'Cliente', 'Modelo Porsche', 
        'Ano Modelo', 'Preço (USD)', 'Quilometragem (mi)', 
        'Método Pagamento', 'Cidade', 'Estado', 'Vendedor', 'Status Entrega'
      ];

      const csvRows = [headers.join(',')];

      filteredData.forEach(d => {{
        const row = [
          d.sale_id,
          `"${{d.SaleDateSanitized}}"`,
          `"${{d.customer_name || ''}}"`,
          `"${{d.PorscheModelSanitized}}"`,
          d.ModelYearSanitized,
          d.SalesPriceSanitized,
          d.VehicleMileageSanitized,
          `"${{d.PayMethodSanitized}}"`,
          `"${{d.CitySanitized}}"`,
          `"${{d.StateSanitized}}"`,
          `"${{d.salesperson || ''}}"`,
          `"${{d.DeliveryStatusSanitized}}"`
        ];
        csvRows.push(row.join(','));
      }});

      const csvContent = 'data:text/csv;charset=utf-8,' + encodeURIComponent(csvRows.join('\\n'));
      const link = document.createElement('a');
      link.setAttribute('href', csvContent);
      link.setAttribute('download', 'porsche_vendas_filtradas.csv');
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }}
  </script>
</body>
</html>
'''

output_path = os.path.join(target_dir, "dashboard.html")
with open(output_path, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"SUCCESS: Dashboard HTML generated at: {output_path}")
print(f"File size: {len(html_content)} bytes")

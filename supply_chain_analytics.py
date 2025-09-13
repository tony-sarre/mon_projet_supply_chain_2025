# app.py
# ========================= PREMIUM SUPPLY CHAIN DASHBOARD =========================
# Requirements:
# pip install dash==2.17.1 dash-bootstrap-components==1.6.0 plotly==5.22.0
# pip install pandas scikit-learn flask-caching numpy
# pip install reportlab
# Optional: pip install openai

import os
import re
import io
import json
import base64
import threading
from pathlib import Path
from datetime import datetime
import numpy as np
import pandas as pd

from flask_caching import Cache
import dash
from dash import Dash, html, dcc, Input, Output, State, dash_table, no_update
import dash_bootstrap_components as dbc
import plotly.express as px
import warnings
warnings.filterwarnings("ignore", message="Parsing dates.*ambiguous", category=DeprecationWarning)

# ------------- OpenAI client (clé hardcodée à ta demande) ----------------
OPENAI_API_KEY_HARDCODED = "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"
openai_client = None
try:
    from openai import OpenAI
    api_key = OPENAI_API_KEY_HARDCODED or os.getenv("OPENAI_API_KEY")
    if api_key:
        openai_client = OpenAI(api_key=api_key)
except Exception:
    openai_client = None
# -------------------------------------------------------------------------

# ----------------------------- Brand & Meta --------------------------------------
APP_TITLE = "Supply Chain Command Center"
APP_BRAND = "Maad SaS"
AUTHOR = "Tony SARRE"
THEME = dbc.themes.CYBORG  # sobre & premium

# ------------------------------ Company / Branding -------------------------------
COMPANY_NAME = "Maad SaS"
COMPANY_CAPITAL = os.getenv("COMPANY_CAPITAL", "")           # ex "100 000 000 XOF"
COMPANY_RCS = os.getenv("COMPANY_RCS", "")                   # ex "RCS Dakar B 123 456"
COMPANY_ADDRESS = os.getenv("COMPANY_ADDRESS", "")
COMPANY_PHONE = os.getenv("COMPANY_PHONE", "")
COMPANY_EMAIL = os.getenv("COMPANY_EMAIL", "")
DEFAULT_TVA_RATE = float(os.getenv("COMPANY_TVA_RATE", "0.18"))  # 18% par défaut

def _get_logo_data_uri():
    """Charge le logo depuis /mnt/data/logo_maad.jpg si présent."""
    try:
        logo_path = Path("logo_maad.jpg")
        if logo_path.exists():
            b64 = base64.b64encode(logo_path.read_bytes()).decode("utf-8")
            return f"data:image/jpeg;base64,{b64}"
    except Exception:
        pass
    return None

LOGO_DATA_URI = _get_logo_data_uri()

def get_logo_for_reportlab():
    """
    Retourne une source compatible ReportLab pour Image():
    - chemin de fichier si logo_maad.jpg présent
    - sinon, un ImageReader à partir du base64 de LOGO_DATA_URI
    - sinon, None
    """
    try:
        from reportlab.lib.utils import ImageReader
        p = Path("logo_maad.jpg")
        if p.exists():
            return str(p)  # ReportLab accepte le chemin fichier

        if LOGO_DATA_URI and "," in LOGO_DATA_URI:
            # extraire la partie base64
            b64 = LOGO_DATA_URI.split(",", 1)[1] if LOGO_DATA_URI.startswith("data:") else LOGO_DATA_URI
            raw = base64.b64decode(b64)
            return ImageReader(io.BytesIO(raw))  # compatible Image()
    except Exception:
        pass
    return None


# ------------------------------ PO Numbering (sequential) ------------------------
PO_COUNTER_PATH = Path("./po_counter.json")  # changer le dossier si nécessaire
PO_LOCK = threading.Lock()

def _load_po_state():
    try:
        if PO_COUNTER_PATH.exists():
            import json as _json
            return _json.loads(PO_COUNTER_PATH.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {"date": None, "seq": 0}

def _save_po_state(state: dict):
    try:
        import json as _json
        PO_COUNTER_PATH.write_text(_json.dumps(state, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass

def get_next_po_number() -> str:
    """PO-YYYYMMDD-###, séquentiel, reset quotidien."""
    today = datetime.now().strftime("%Y%m%d")
    with PO_LOCK:
        st = _load_po_state()
        if st.get("date") != today:
            st = {"date": today, "seq": 1}
        else:
            st["seq"] = int(st.get("seq", 0)) + 1
        _save_po_state(st)
        return f"PO-{today}-{st['seq']:03d}"

# ------------------------------ Data loading -------------------------------------

def load_supply_data() -> pd.DataFrame:
    from sklearn.ensemble import RandomForestRegressor
    from sklearn.preprocessing import OneHotEncoder
    from sklearn.compose import ColumnTransformer
    from sklearn.pipeline import Pipeline
    from sklearn.impute import SimpleImputer

    # URLs (inchangés + delisting)
    SUPPLIERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRTyAxh6v8o0FXV0r7f6ALPDgmeJNkjTZITjrEoKBHo2gs_f3iyV8sFk8fOzcAsUSkJMXBJCpJnhQKi/pub?gid=1015760114&single=true&output=csv"
    url_leadtime = "https://data.heroku.com/dataclips/lgzfzobmggbjgowhffqskezhsxhq.csv"
    BUFFER_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT-503n7w2ixefop7XeUnda3B76ui1jQRZKJHehhJ0WtumnpSzUVYjnvGv-_tFQ6jXayAcjJEAryQMv/pub?gid=1011110883&single=true&output=csv"
    inventory_pikine_staging = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT-503n7w2ixefop7XeUnda3B76ui1jQRZKJHehhJ0WtumnpSzUVYjnvGv-_tFQ6jXayAcjJEAryQMv/pub?gid=1876150276&single=true&output=csv"
    sales_pikine = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT-503n7w2ixefop7XeUnda3B76ui1jQRZKJHehhJ0WtumnpSzUVYjnvGv-_tFQ6jXayAcjJEAryQMv/pub?gid=1493123930&single=true&output=csv"
    Tbh_7dsales = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT-503n7w2ixefop7XeUnda3B76ui1jQRZKJHehhJ0WtumnpSzUVYjnvGv-_tFQ6jXayAcjJEAryQMv/pub?gid=1080970598&single=true&output=csv"
    Tbh_30dsales_products = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT-503n7w2ixefop7XeUnda3B76ui1jQRZKJHehhJ0WtumnpSzUVYjnvGv-_tFQ6jXayAcjJEAryQMv/pub?gid=1655420642&single=true&output=csv"
    Product_category = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQbJqHsr6Kifee7I91YD-7-sCZDWgM5GvxCeN0OqUvZhok0j-kDywguqe5I61y97b-uBhHbWraTIrux/pub?gid=803048228&single=true&output=csv"
    DELISTING_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSugb2blp3F6gZhyxjyRjExj8f4I6xpp8J2KFQCIfNW77VzG8WqFOw0PjEBrjzxf00mjEWaViFuhgMf/pub?gid=1681543945&single=true&output=csv"

    # Load
    suppliers_df = pd.read_csv(SUPPLIERS_URL)
    df_leadtime = pd.read_csv(url_leadtime)
    buffer_df = pd.read_csv(BUFFER_URL)
    inventory_pikine_staging_df = pd.read_csv(inventory_pikine_staging, skiprows=1)
    sales_pikine_df = pd.read_csv(sales_pikine, header=1, low_memory=False)
    Tbh_7dsales_df = pd.read_csv(Tbh_7dsales)
    Tbh_30dsales_products_df = pd.read_csv(Tbh_30dsales_products)
    Product_category_df = pd.read_csv(Product_category)

    # Clean names
    if 'product_name' in inventory_pikine_staging_df.columns:
        inventory_pikine_staging_df['product_name'] = inventory_pikine_staging_df['product_name'].str.lower().str.strip()
    if 'product_name' in Product_category_df.columns:
        Product_category_df['product_name'] = Product_category_df['product_name'].str.lower().str.strip()
    if '2' in Tbh_7dsales_df.columns:
        Tbh_7dsales_df['2'] = Tbh_7dsales_df['2'].astype(str).str.lower().str.strip()
    if '2' in Tbh_30dsales_products_df.columns:
        Tbh_30dsales_products_df['2'] = Tbh_30dsales_products_df['2'].astype(str).str.lower().str.strip()
    if 'product_name' in sales_pikine_df.columns:
        sales_pikine_df['product_name'] = sales_pikine_df['product_name'].astype(str).str.lower().str.strip()

    # ADS 7 & 30
    if {'2','9'}.issubset(Tbh_7dsales_df.columns):
        ads7 = Tbh_7dsales_df[['2','9']].copy().rename(columns={'2':'product_name','9':'Average Daily Sales (7d)'})
        ads7['Average Daily Sales (7d)'] = pd.to_numeric(ads7['Average Daily Sales (7d)'], errors='coerce')
    else:
        ads7 = pd.DataFrame(columns=['product_name','Average Daily Sales (7d)'])
    if {'2','9'}.issubset(Tbh_30dsales_products_df.columns):
        ads30 = Tbh_30dsales_products_df[['2','9']].copy().rename(columns={'2':'product_name','9':'Average Daily Sales (30d)'})
        ads30['Average Daily Sales (30d)'] = pd.to_numeric(ads30['Average Daily Sales (30d)'], errors='coerce')
    else:
        ads30 = pd.DataFrame(columns=['product_name','Average Daily Sales (30d)'])

    # Stock total
    if {'product_id','Supplier','total_stock','product_name'}.issubset(inventory_pikine_staging_df.columns):
        total_stock_df = inventory_pikine_staging_df.groupby(['product_name','Supplier'])['total_stock'].sum().reset_index()
    else:
        total_stock_df = pd.DataFrame(columns=['product_name','Supplier','total_stock'])

    # Merge ventes
    merged_sales_df = pd.merge(ads7, ads30, on='product_name', how='left')
    if 'Average Daily Sales (7d)' in merged_sales_df.columns:
        merged_sales_df['Average Daily Sales (7d)'] = merged_sales_df['Average Daily Sales (7d)'] + 0.1
    if 'Average Daily Sales (30d)' in merged_sales_df.columns:
        merged_sales_df['Average Daily Sales (30d)'] = merged_sales_df['Average Daily Sales (30d)'] + 0.1
    final_stock_sales_df = pd.merge(total_stock_df, merged_sales_df, on='product_name', how='left')

    # Coverage (7/30)
    final_stock_sales_df['Coverage Day (7d)'] = final_stock_sales_df['total_stock'] / final_stock_sales_df['Average Daily Sales (7d)'].replace(0, pd.NA)
    final_stock_sales_df['Coverage Day (30d)'] = final_stock_sales_df['total_stock'] / final_stock_sales_df['Average Daily Sales (30d)'].replace(0, pd.NA)

    # Max Daily Sales Pikine
    if {'product_name','sum'}.issubset(sales_pikine_df.columns):
        max_sales_pikine = sales_pikine_df.groupby('product_name')['sum'].max().reset_index()
        max_sales_pikine.rename(columns={'sum':'Max Daily Sales (Pikine)'}, inplace=True)
        max_sales_pikine['Max Daily Sales (Pikine)'] = pd.to_numeric(max_sales_pikine['Max Daily Sales (Pikine)'], errors='coerce')
    else:
        max_sales_pikine = pd.DataFrame(columns=['product_name','Max Daily Sales (Pikine)'])
    final_stock_sales_df = pd.merge(final_stock_sales_df, max_sales_pikine, on='product_name', how='left')

    # Catégorie
    if {'product_name','CATEGORY ABC-XYZ'}.issubset(Product_category_df.columns):
        cat = Product_category_df[['product_name','CATEGORY ABC-XYZ']].copy().rename(columns={'CATEGORY ABC-XYZ':'Product Category'})
        final_stock_sales_df = pd.merge(final_stock_sales_df, cat, on='product_name', how='left')
        final_stock_sales_df['Product Category'] = final_stock_sales_df['Product Category'].fillna('CX')
    else:
        final_stock_sales_df['Product Category'] = 'CX'

    # Buffer
    buffer_df_subset = pd.read_csv(BUFFER_URL, usecols=[0,1,2])
    if 'buffer In Stock' in buffer_df_subset.columns:
        first_col = buffer_df_subset.columns[0]
        buf = buffer_df_subset[[first_col,'buffer In Stock']].copy().rename(columns={first_col:'Product Category','buffer In Stock':'Buffer Value'})
        buf['Buffer Value'] = pd.to_numeric(buf['Buffer Value'], errors='coerce').fillna(0)
        final_stock_sales_df = pd.merge(final_stock_sales_df, buf, on='Product Category', how='left')
    else:
        final_stock_sales_df['Buffer Value'] = 0

    # Optimal Stock (reorder point — inchangé)
    if {'Max Daily Sales (Pikine)','Average Daily Sales (30d)','Buffer Value'}.issubset(final_stock_sales_df.columns):
        final_stock_sales_df['Optimal Stock (Reorder Point)'] = final_stock_sales_df.apply(
            lambda row: max(row['Max Daily Sales (Pikine)'],
                            (row['Max Daily Sales (Pikine)']/2.0) + (row['Buffer Value'] * row['Average Daily Sales (30d)'])),
            axis=1
        )
    else:
        final_stock_sales_df['Optimal Stock (Reorder Point)'] = pd.NA

    # Lead time
    if {'supplier','max_leadtime'}.issubset(df_leadtime.columns):
        ltd = df_leadtime[['supplier','max_leadtime']].copy().rename(columns={'supplier':'Supplier','max_leadtime':'Max Lead Time'})
        final_stock_sales_df = pd.merge(final_stock_sales_df, ltd, on='Supplier', how='left')
    else:
        final_stock_sales_df['Max Lead Time'] = pd.NA

    # OOS rate 30d
    if {'2','28'}.issubset(Tbh_30dsales_products_df.columns):
        oos = Tbh_30dsales_products_df[['2','28']].copy().rename(columns={'2':'product_name','28':'Daily OOS Rate (30d)'})
        oos['Daily OOS Rate (30d)'] = oos['Daily OOS Rate (30d)'].astype(str).str.replace('%','',regex=False)
        oos['Daily OOS Rate (30d)'] = pd.to_numeric(oos['Daily OOS Rate (30d)'], errors='coerce').fillna(0)/100.0
        final_stock_sales_df = pd.merge(final_stock_sales_df, oos, on='product_name', how='left')
    else:
        final_stock_sales_df['Daily OOS Rate (30d)'] = pd.NA

    # >>> Consolidation demandée
    # Max Avg Daily Sales & Max Coverage Day
    if {'Average Daily Sales (7d)','Average Daily Sales (30d)'}.issubset(final_stock_sales_df.columns):
        final_stock_sales_df['Max Avg Daily Sales'] = final_stock_sales_df[['Average Daily Sales (7d)','Average Daily Sales (30d)']].max(axis=1)
    else:
        final_stock_sales_df['Max Avg Daily Sales'] = pd.NA
    if {'Coverage Day (7d)','Coverage Day (30d)'}.issubset(final_stock_sales_df.columns):
        final_stock_sales_df['Max Coverage Day'] = final_stock_sales_df[['Coverage Day (7d)','Coverage Day (30d)']].max(axis=1)
    else:
        final_stock_sales_df['Max Coverage Day'] = pd.NA

    # Supprimer colonnes intermédiaires
    for c in ['Average Daily Sales (7d)','Average Daily Sales (30d)','Coverage Day (7d)','Coverage Day (30d)']:
        if c in final_stock_sales_df.columns:
            final_stock_sales_df.drop(columns=[c], inplace=True)

    # Delisting (nouvelle colonne)
    try:
        delisting_df = pd.read_csv(DELISTING_URL)
        if delisting_df.shape[1] > 3:
            dsub = delisting_df.iloc[:, [0,3]].copy()
            dsub.columns = ['product_name','delisting']
            dsub['product_name'] = dsub['product_name'].astype(str).str.lower().str.strip()
            final_stock_sales_df = pd.merge(final_stock_sales_df, dsub, on='product_name', how='left')
            final_stock_sales_df['delisting'] = final_stock_sales_df['delisting'].fillna('Not Delisted')
        else:
            final_stock_sales_df['delisting'] = 'Not Delisted'
    except Exception:
        final_stock_sales_df['delisting'] = 'Not Delisted'

    # --------- Imputation ML (identique esprit)
    missing_values_summary = final_stock_sales_df.isnull().sum()
    imputation_candidates = missing_values_summary[missing_values_summary > 0].index.tolist()
    features_for_imputation = {}
    all_cols = final_stock_sales_df.columns.tolist()
    for col in imputation_candidates:
        features_for_imputation[col] = [f for f in all_cols if f != col and final_stock_sales_df[f].isnull().sum() < len(final_stock_sales_df)*0.5 and f in [
            'product_name','Supplier','total_stock','Product Category','Buffer Value','Max Lead Time',
            'Max Daily Sales (Pikine)','Max Avg Daily Sales','Max Coverage Day'
        ]]
    for target_col in imputation_candidates:
        features = features_for_imputation.get(target_col, [])
        if not features:
            final_stock_sales_df[target_col] = final_stock_sales_df[target_col].fillna(0)
            continue
        valid_features = [f for f in features if f in final_stock_sales_df.columns]
        if not valid_features:
            final_stock_sales_df[target_col] = final_stock_sales_df[target_col].fillna(0)
            continue
        subset_df = final_stock_sales_df[valid_features + [target_col]].copy()
        train_data = subset_df.dropna(subset=[target_col])
        predict_data = subset_df[subset_df[target_col].isnull()]

        from sklearn.preprocessing import OneHotEncoder
        from sklearn.compose import ColumnTransformer
        from sklearn.pipeline import Pipeline
        from sklearn.impute import SimpleImputer
        from sklearn.ensemble import RandomForestRegressor

        X_train = train_data[valid_features]
        y_train = train_data[target_col]
        X_predict = predict_data[valid_features]

        categorical_features = X_train.select_dtypes(include=['object','category']).columns
        numerical_features = X_train.select_dtypes(include=['number']).columns

        numerical_pipeline = Pipeline([('imputer', SimpleImputer(strategy='mean'))])
        categorical_pipeline = Pipeline([
            ('imputer', SimpleImputer(strategy='constant', fill_value='missing')),
            ('onehot', OneHotEncoder(handle_unknown='ignore'))
        ])
        preprocessor = ColumnTransformer(
            transformers=[
                ('num', numerical_pipeline, numerical_features),
                ('cat', categorical_pipeline, categorical_features)
            ],
            remainder='passthrough'
        )
        model_pipeline = Pipeline([
            ('preprocessor', preprocessor),
            ('regressor', RandomForestRegressor(n_estimators=100, random_state=42, n_jobs=-1))
        ])
        try:
            model_pipeline.fit(X_train, y_train)
            predicted = model_pipeline.predict(X_predict)
            missing_idx = final_stock_sales_df[final_stock_sales_df[target_col].isnull()].index
            if len(predicted) == len(missing_idx):
                final_stock_sales_df.loc[missing_idx, target_col] = predicted
            else:
                final_stock_sales_df[target_col] = final_stock_sales_df[target_col].fillna(0)
        except Exception:
            final_stock_sales_df[target_col] = final_stock_sales_df[target_col].fillna(0)

    # Flags & outputs (inchangés)
    if {'total_stock','Optimal Stock (Reorder Point)'}.issubset(final_stock_sales_df.columns):
        final_stock_sales_df['Predicted Stockout'] = final_stock_sales_df['total_stock'] <= final_stock_sales_df['Optimal Stock (Reorder Point)']
    else:
        final_stock_sales_df['Predicted Stockout'] = False

    def get_stock_status(row):
        stock = pd.to_numeric(row.get('total_stock', 0), errors='coerce')
        reorder_point = pd.to_numeric(row.get('Optimal Stock (Reorder Point)', 0), errors='coerce')
        max_lead_time = pd.to_numeric(row.get('Max Lead Time', 0), errors='coerce')
        max_avg_daily_sales = pd.to_numeric(row.get('Max Avg Daily Sales', 0), errors='coerce')
        if pd.isna(stock) or stock <= 0:
            return 'Out of Stock'
        elif pd.isna(reorder_point) or stock <= reorder_point:
            return 'Predicted Stockout Soon'
        elif pd.isna(max_lead_time) or pd.isna(max_avg_daily_sales) or stock <= reorder_point + (max_lead_time * max_avg_daily_sales):
            return 'Order Soon'
        else:
            return 'Stock OK'
    final_stock_sales_df['Stock Status'] = final_stock_sales_df.apply(get_stock_status, axis=1)

    if {'total_stock','Optimal Stock (Reorder Point)'}.issubset(final_stock_sales_df.columns):
        final_stock_sales_df['Predicted Order Quantity'] = final_stock_sales_df.apply(
            lambda row: max(0, (row['Optimal Stock (Reorder Point)'] or 0) - (row['total_stock'] or 0)), axis=1
        )
    else:
        final_stock_sales_df['Predicted Order Quantity'] = pd.NA

    if 'Daily OOS Rate (30d)' in final_stock_sales_df.columns:
        final_stock_sales_df['Daily OOS Rate (30d)'] = pd.to_numeric(final_stock_sales_df['Daily OOS Rate (30d)'], errors='coerce')

    # Order d’affichage — ajoute delisting
    preferred_order = [
        'product_name','Supplier','Product Category','delisting',
        'total_stock','Optimal Stock (Reorder Point)',
        'Max Daily Sales (Pikine)','Max Lead Time','Max Avg Daily Sales','Max Coverage Day',
        'Daily OOS Rate (30d)','Predicted Stockout','Stock Status','Predicted Order Quantity',
        'Unit Price HT','Discount'
    ]
    cols = [c for c in preferred_order if c in final_stock_sales_df.columns] + \
           [c for c in final_stock_sales_df.columns if c not in preferred_order]
    final_stock_sales_df = final_stock_sales_df[cols]
    return final_stock_sales_df

# Utility: add Actions columns
def add_action_cols(df: pd.DataFrame) -> pd.DataFrame:
    df2 = df.copy()
    df2["✏️ Edit"] = "✏️"
    df2["🗑️ Delete"] = "🗑️"
    return df2

# ----------------------------- App & Cache ---------------------------------------
app = Dash(__name__, title=APP_TITLE, external_stylesheets=[THEME], suppress_callback_exceptions=True)
server = app.server
cache = Cache(app.server, config={"CACHE_TYPE": "SimpleCache", "CACHE_DEFAULT_TIMEOUT": 3600})

# ------------------------------ Custom CSS & JS ----------------------------------
app.index_string = """
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <style>
            :root {
                --brand-accent:#22d3ee;
                --badge-danger:#ef4444; --badge-warning:#f59e0b; --badge-ok:#10b981;
            }
            .sidebar { position: fixed; top:0; bottom:0; left:0; width: 320px;
                padding: 16px 14px; background: #0b1220; border-right:1px solid #1f2937; overflow-y:auto; }
            .content { margin-left: 320px; padding: 18px 18px 120px 18px; }
            .brand { font-weight:800; font-size:20px; letter-spacing:.6px; color:#fff; }
            .muted { color:#9ca3af; font-size:13px; }
            .pill { border:1px solid #243244; padding:10px 12px; border-radius:12px; background:#0f1828; }
            .kpi { border-radius:16px; padding:18px; border:1px solid #1f2937; background:linear-gradient(180deg,#0b1220,#0f1625);
                   box-shadow:0 10px 24px rgba(0,0,0,.25); }
            .banner-risk { border-left:4px solid var(--badge-danger); background: rgba(239,68,68,.08);
                padding:10px 14px; border-radius:10px; margin-bottom:10px; }
            .badge{ padding:2px 8px; border-radius:10px; font-size:11px; border:1px solid #374151 }
            .badge-danger{ background:rgba(239,68,68,.15); color:#fecaca; border-color:#7f1d1d; }
            .badge-warn{ background:rgba(245,158,11,.15); color:#fde68a; border-color:#78350f; }
            .badge-ok{ background:rgba(16,185,129,.15); color:#a7f3d0; border-color:#064e3b; }
            .btn-primary { background: var(--brand-accent); color:#001018; font-weight:700; border:none; }
            .search-input input { background:#0a1320; color:#e5e7eb; border:1px solid #1f2937; border-radius:10px; }
            .section-title { font-weight:700; font-size:18px; margin-bottom:10px; }
            .soft-card { border:1px solid #1f2937; border-radius:16px; padding:14px; background:#0b1220; }

            /* --- Floating chat bubble (LIGHT THEME) --- */
            .chat-fab {
                position: fixed; right: 24px; bottom: 24px; z-index: 10000;
                border-radius: 9999px; padding: 12px 16px; border:none;
                background: var(--brand-accent); color:#001018; font-weight:800;
                box-shadow: 0 12px 24px rgba(0,0,0,.35);
                display:flex; align-items:center; gap:8px; cursor:pointer;
            }
            .chat-window {
                position: fixed; right: 24px; bottom: 92px; width: 520px; max-width: 96vw;
                background: #ffffff; border:1px solid #e5e7eb; border-radius:16px;
                z-index: 10000; box-shadow: 0 24px 48px rgba(0,0,0,.15);
                display:flex; flex-direction:column; overflow:hidden; color:#111827;
            }
            .chat-header {
                padding:10px 12px; display:flex; align-items:center; justify-content:space-between;
                background:#f3f4f6; border-bottom:1px solid #e5e7eb; color:#111827;
            }
            .chat-body {
                padding:12px; max-height:56vh; overflow:auto; display:flex; flex-direction:column; gap:10px;
                background:#fafafa;
            }
            .chat-input-wrap {
                padding:10px; background:#f9fafb; border-top:1px solid #e5e7eb; display:grid;
                grid-template-columns: 1fr auto; gap:10px;
            }
            .chat-textarea { background:#ffffff; color:#111827; border:1px solid #d1d5db; border-radius:10px; }
            .chat-avatar {
                min-width:32px; height:32px; border-radius:50%; display:flex; align-items:center; justify-content:center;
                font-size:16px; color:#fff; background:#6366f1; box-shadow:0 2px 6px rgba(0,0,0,.1);
            }
            .chat-avatar.user { background:#3b82f6; }
            .chat-bubble { display:flex; gap:10px; }
            .chat-bubble.user { flex-direction: row; }
            .chat-bubble.bot { flex-direction: row-reverse; }
            .chat-msg { background:#ffffff; border:1px solid #e5e7eb; border-radius:12px; padding:10px 12px; max-width: 90%; }
            .upload-box {
                border: 2px dashed #9ca3af; border-radius: 10px; padding: 10px; text-align:center; color:#6b7280; background:#ffffff;
            }
        </style>
        <script>
            // Enter to send / Shift+Enter newline
            const attachEnterSend = () => {
                const ta = document.getElementById('chat-input');
                const btn = document.getElementById('chat-send');
                if(!ta || !btn) return;
                if(ta._boundEnter) return; // avoid double bind
                ta._boundEnter = true;
                ta.addEventListener('keydown', (e) => {
                    if(e.key === 'Enter' && !e.shiftKey) {
                        e.preventDefault();
                        btn.click();
                    }
                });
            };
            const mo = new MutationObserver(() => attachEnterSend());
            window.addEventListener('load', () => {
                attachEnterSend();
                const app = document.getElementById('_dash-app-content');
                if(app) mo.observe(app, {childList: true, subtree: true});
            });

            // Simple debounce available if needed
            window.debounce = (func, wait=280) => {
                let timeout; 
                return (...args) => { clearTimeout(timeout); timeout = setTimeout(()=>func.apply(this,args), wait); }
            };
        </script>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
"""

# ------------------------------ Data cache layer ---------------------------------
@cache.memoize()
def get_df_cached():
    return load_supply_data()

# ------------------------------ Sidebar ------------------------------------------
def make_sidebar():
    df = get_df_cached()
    suppliers = sorted([s for s in df['Supplier'].dropna().unique().tolist() if s != '']) if 'Supplier' in df.columns else []
    cats = sorted([c for c in df['Product Category'].dropna().unique().tolist() if c != '']) if 'Product Category' in df.columns else []
    status_vals = df['Stock Status'].dropna().unique().tolist() if 'Stock Status' in df.columns else []

    return html.Div(className="sidebar", children=[
        html.Div([
            html.Div([
                html.Img(src=LOGO_DATA_URI, style={"height": "42px", "marginRight": "8px"}) if LOGO_DATA_URI else html.Div(),
                html.Div(APP_BRAND, className="brand"),
            ], style={"display": "flex", "alignItems": "center", "gap": "10px"}),
            html.Div("Supply Chain Command Center", className="muted")
        ]),
        html.Hr(),
        html.Div(className="banner-risk", id="risk-banner", children="Chargement..."),
        html.Div(className="section-title", children="Recherche"),
        dbc.InputGroup(className="search-input", children=[
            dbc.Input(id="search-input", placeholder="Rechercher un produit...", type="text", debounce=True)
        ]),
        html.Br(),
        html.Div(className="section-title", children="Filtres"),
        html.Div(className="pill", children=[
            html.Small("Supplier"),
            dcc.Dropdown(id="filter-supplier", options=[{"label": s, "value": s} for s in suppliers],
                         multi=True, placeholder="Tous", persistence=True),
            html.Br(),
            html.Small("Statut"),
            dcc.Dropdown(id="filter-status", options=[{"label": s, "value": s} for s in status_vals],
                         multi=True, placeholder="Tous", persistence=True),
            html.Br(),
            html.Small("Catégorie"),
            dcc.Dropdown(id="filter-category", options=[{"label": c, "value": c} for c in cats],
                         multi=True, placeholder="Toutes", persistence=True),
            html.Br(),
            dbc.Checklist(
                options=[{"label": "Afficher uniquement risques de rupture", "value": "risk"}],
                value=[], id="toggle-risk-only", switch=True
            )
        ]),
        html.Br(),
        html.Div([
            dbc.Button("Exporter CSV (filtré)", id="btn-export", className="btn-primary", size="sm"),
            html.Span(" "),
            dbc.Button("Bon de commande PDF", id="btn-po-pdf", className="btn-primary", size="sm", disabled=True),
            dcc.Download(id="download-data"),
            dcc.Download(id="download-po"),
        ]),
        html.Br(),
        html.Div([
            dbc.Nav([
                dbc.NavLink("Overview", href="/", id="nav-overview", active="exact"),
                dbc.NavLink("Analyses", href="/analytics", id="nav-analytics", active="exact"),
                dbc.NavLink("Prédictions", href="/predictions", id="nav-pred", active="exact"),
                dbc.NavLink("About", href="/about", id="nav-about", active="exact"),
            ], vertical=True, pills=True)
        ]),
        html.Br(),
        html.Small(f"© {datetime.now().year} • {AUTHOR}")
    ])

# ------------------------------ Pages --------------------------------------------
def make_kpis(df: pd.DataFrame):
    def fmt(n):
        if pd.isna(n): return "-"
        if isinstance(n, (int, float)):
            try:
                if abs(n) >= 1000: return f"{n:,.0f}".replace(",", " ")
                return f"{n:,.0f}"
            except Exception:
                return str(n)
        return str(n)

    total_skus = df['product_name'].nunique() if 'product_name' in df.columns else len(df)
    at_risk = int((df['Stock Status'].isin(['Out of Stock', 'Predicted Stockout Soon']).sum())) if 'Stock Status' in df.columns else 0
    order_sum = df['Predicted Order Quantity'].sum() if 'Predicted Order Quantity' in df.columns else 0
    suppliers = df['Supplier'].nunique() if 'Supplier' in df.columns else 0

    cards = dbc.Row([
        dbc.Col(html.Div(className="kpi", children=[html.Small("SKUs"), html.H3(fmt(total_skus))]), md=3),
        dbc.Col(html.Div(className="kpi", children=[html.Small("À risque de rupture"),
            html.H3([fmt(at_risk), html.Span("  ", className="mx-1"), html.Span("•", className="badge badge-danger")])]), md=3),
        dbc.Col(html.Div(className="kpi", children=[html.Small("Qté réassort prédite (somme)"), html.H3(fmt(order_sum))]), md=3),
        dbc.Col(html.Div(className="kpi", children=[html.Small("Fournisseurs actifs"), html.H3(fmt(suppliers))]), md=3),
    ], className="gy-3")
    return cards

def page_overview(master_df: pd.DataFrame = None):
    df = master_df if master_df is not None else get_df_cached()
    df = add_action_cols(df)
    kpi_cards = make_kpis(df)

    table = dash_table.DataTable(
        id="main-table",
        columns=[{"name": c, "id": c, "deletable": False} for c in df.columns] +
                [{"name": "Actions", "id": "✏️ Edit"}, {"name": " ", "id": "🗑️ Delete"}],
        data=df.to_dict("records"),
        page_size=15,
        filter_action="native",
        sort_action="native", sort_mode="multi",
        column_selectable="single",
        editable=True,
        row_selectable="single",
        selected_rows=[],
        style_table={"overflowX": "auto"},
        style_header={"backgroundColor": "#0f1625", "border": "1px solid #1f2937", "fontWeight": "700"},
        style_cell={"backgroundColor": "#0b1220", "color": "#e5e7eb", "border": "1px solid #1f2937", "fontSize": 12},
        style_data_conditional=[
            {"if": {"filter_query": "{Stock Status} = 'Out of Stock'"}, "backgroundColor": "rgba(239,68,68,.15)", "color": "#fecaca"},
            {"if": {"filter_query": "{Stock Status} = 'Predicted Stockout Soon'"}, "backgroundColor": "rgba(245,158,11,.15)", "color": "#fde68a"},
            {"if": {"filter_query": "{Stock Status} = 'Order Soon'"}, "backgroundColor": "rgba(59,130,246,.12)", "color": "#bfdbfe"},
            {"if": {"filter_query": "{Stock Status} = 'Stock OK'"}, "backgroundColor": "rgba(16,185,129,.1)", "color": "#a7f3d0"},
        ],
        export_format="none",
        persistence=True,
        persisted_props=["filter_query", "sort_by", "page_current", "selected_rows", "selected_columns", "hidden_columns"],
    )

    add_btn = dbc.Button("➕ Ajouter", id="btn-add-row", className="btn-primary", size="sm")

    modal = dbc.Modal([
        dbc.ModalHeader(dbc.ModalTitle("Éditer / Ajouter un produit")),
        dbc.ModalBody([
            dbc.Row([
                dbc.Col([html.Small("Nom produit"), dbc.Input(id="edit-product", type="text")], md=6),
                dbc.Col([html.Small("Fournisseur"), dbc.Input(id="edit-supplier", type="text")], md=6),
            ]),
            html.Br(),
            dbc.Row([
                dbc.Col([html.Small("Catégorie"), dbc.Input(id="edit-category", type="text")], md=6),
                dbc.Col([html.Small("Stock total"), dbc.Input(id="edit-stock", type="number")], md=6),
            ]),
        ]),
        dbc.ModalFooter([
            dbc.Button("Annuler", id="edit-cancel", className="btn-secondary", n_clicks=0),
            dbc.Button("Enregistrer", id="edit-save", className="btn-primary", n_clicks=0),
        ]),
    ], id="edit-modal", is_open=False, backdrop="static")

    return html.Div(className="content", children=[
        html.H2("Overview"),
        html.Div(kpi_cards),
        html.Br(),
        html.Div(className="soft-card", children=[
            html.Div(dbc.Row([dbc.Col(html.Div("Détails Produits (non modifié : calculs d’origine)", className="section-title"), md=9),
                              dbc.Col(html.Div(add_btn, style={"textAlign":"right"}), md=3)])),
            table
        ]),
        modal
    ])

def page_analytics():
    df = get_df_cached()
    top_po = df.sort_values("Predicted Order Quantity", ascending=False).head(20) if 'Predicted Order Quantity' in df.columns else df.head(20)
    fig_po = px.bar(top_po, x="product_name", y="Predicted Order Quantity", color="Supplier",
                    title="Top 20 — Besoin de réappro (Predicted Order Quantity)") if 'Predicted Order Quantity' in top_po.columns else px.bar()
    if not fig_po.data: fig_po.update_layout(title="Top 20 — Besoin de réappro (Données indisponibles)")

    cat_share = df.groupby("Product Category", dropna=False)['total_stock'].sum().reset_index() if {'Product Category','total_stock'}.issubset(df.columns) else pd.DataFrame(columns=['Product Category','total_stock'])
    fig_cat = px.pie(cat_share, names="Product Category", values="total_stock", title="Répartition du stock par catégorie") if not cat_share.empty else px.pie()

    status_count = df['Stock Status'].value_counts(dropna=False).reset_index() if 'Stock Status' in df.columns else pd.DataFrame(columns=['Stock Status','count'])
    if not status_count.empty: status_count.columns = ['Stock Status','count']
    fig_status = px.bar(status_count, x="Stock Status", y="count", title="Distribution des statuts") if not status_count.empty else px.bar()

    sup_map = df.groupby(['Supplier','Product Category'], dropna=False)['Predicted Order Quantity'].sum().reset_index() \
        if {'Supplier','Product Category','Predicted Order Quantity'}.issubset(df.columns) else pd.DataFrame(columns=['Supplier','Product Category','Predicted Order Quantity'])
    fig_treemap = px.treemap(sup_map, path=['Supplier','Product Category'], values='Predicted Order Quantity',
                             title="Carte réassort — Supplier > Category") if not sup_map.empty else px.treemap()

    for fig in [fig_po, fig_cat, fig_status, fig_treemap]:
        fig.update_layout(template="plotly_dark")

    return html.Div(className="content", children=[
        html.H2("Analyses"),
        dbc.Row([dbc.Col(dcc.Graph(figure=fig_po), md=12)], className="gy-3"),
        html.Br(),
        dbc.Row([dbc.Col(dcc.Graph(figure=fig_cat), md=6), dbc.Col(dcc.Graph(figure=fig_status), md=6)], className="gy-3"),
        html.Br(),
        dbc.Row([dbc.Col(dcc.Graph(figure=fig_treemap), md=12)], className="gy-3"),
    ])

def page_predictions():
    table = dash_table.DataTable(
        id="pred-table",
        columns=[{"name": c, "id": c} for c in ["product_name","Supplier","total_stock",
                                                "Optimal Stock (Reorder Point)","Max Lead Time",
                                                "Max Avg Daily Sales","Buffer Value","Suggested Order Qty"]],
        data=[], page_size=15, sort_action="native",
        style_table={"overflowX": "auto"},
        style_header={"backgroundColor":"#0f1625","border":"1px solid #1f2937","fontWeight":"700"},
        style_cell={"backgroundColor":"#0b1220","color":"#e5e7eb","border":"1px solid #1f2937","fontSize":12},
    )
    controls = html.Div(className="soft-card", children=[
        html.Div("Scénarios What-If (non intrusif)", className="section-title"),
        dbc.Row([
            dbc.Col([html.Small("Multiplicateur Buffer (x)"), dcc.Slider(0.5, 2.0, 0.1, value=1.0, id="buffer-mult")], md=6),
            dbc.Col([html.Small("Multiplicateur Lead Time (x)"), dcc.Slider(0.5, 2.0, 0.1, value=1.0, id="lead-mult")], md=6),
        ]),
        html.Br(), html.Small("La suggestion ci-dessous est recalculée à la volée pour l’exploration, sans changer les résultats source.")
    ])
    return html.Div(className="content", children=[
        html.H2("Prédictions"), controls, html.Br(),
        html.Div(className="soft-card", children=[html.Div("Plan de Réassort — Suggestions", className="section-title"), table])
    ])

def page_about():
    return html.Div(className="content", children=[
        html.H2("About"),
        html.Div(className="soft-card", children=[
            html.H4(APP_BRAND),
            html.P("Plateforme premium de pilotage Supply Chain : visibilité temps quasi-réel des stocks, risques de rupture, analyses graphiques, et recommandations d’approvisionnement."),
            html.Hr(),
            html.H5("Auteur"),
            html.P(f"{AUTHOR} — Data Scientist / Ph.D / Supply Chain — Dashboard construit avec Dash/Plotly, stylé via Bootstrap (thème {THEME.rsplit('.',1)[-1]}) et CSS custom."),
        ])
    ])

# ------------------------------ Chatbot helpers ----------------------------------
def _detect_lang(text: str) -> str:
    if not text: return "fr"
    t = text.lower()
    fr_markers = ["bonjour","salut","stock","commande","fournisseur","rupture","couverture","jours","crédit","delisting"]
    if any(w in t for w in fr_markers) or re.search(r"[àâçéèêëîïôùûüœ]", t): return "fr"
    return "en"

def _clean_df_for_advice(df: pd.DataFrame) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame) or df.empty: return pd.DataFrame()
    d = pd.DataFrame()
    d["product_name_display"] = df.get("product_name","")
    d["supplier_name"] = df.get("Supplier","")
    d["abc_class"] = df.get("Product Category","")
    d["xyz_class"] = ""
    d["current_stock"] = pd.to_numeric(df.get("total_stock",0), errors="coerce").fillna(0).clip(lower=0)
    d["avg_daily_sales"] = pd.to_numeric(df.get("Max Avg Daily Sales",0), errors="coerce").fillna(0).clip(lower=0)
    cov = np.where(d["avg_daily_sales"]>0, d["current_stock"]/d["avg_daily_sales"], np.nan)
    d["coverage_days"] = pd.to_numeric(df.get("Max Coverage Day", cov), errors="coerce")
    d["coverage_days"] = pd.Series(d["coverage_days"]).fillna(0).clip(lower=0, upper=365)
    d["leadtime_days"] = pd.to_numeric(df.get("Max Lead Time",7), errors="coerce").fillna(7).clip(lower=0, upper=180)
    d["credit_days"] = 14
    if "Predicted Stockout" in df.columns:
        d["rupture_ml"] = np.where(df["Predicted Stockout"], "OUI", "NON")
    else:
        st = df.get("Stock Status","").astype(str)
        d["rupture_ml"] = np.where(st.isin(["Out of Stock","Predicted Stockout Soon"]), "OUI", "NON")
    d["delisting_product"] = "NON"
    bad = d["product_name_display"].astype(str).str.contains(r'\b(CFA|CASH|CFA\s*-\s*CASH|ESPECES|CAISSE)\b', case=False, na=False)
    d = d[~bad]
    d["risk"] = (d["rupture_ml"].astype(str).str.upper()=="OUI").astype(int)
    return d

def _bubble(role: str, text: str):
    is_user = (role == "user")
    return html.Div(className=f"chat-bubble {'user' if is_user else 'bot'}", children=[
        html.Div("🧑" if is_user else "🤖", className=f"chat-avatar {'user' if is_user else ''}"),
        html.Div(dcc.Markdown(text or "", link_target="_blank", style={"whiteSpace":"pre-wrap","wordBreak":"break-word"}), className="chat-msg")
    ])

def _render_messages(msgs: list):
    return [_bubble(m.get("role","assistant"), m.get("text","")) for m in msgs or []]

def _chatbot_reply(user_text: str, df: pd.DataFrame, history_messages: list) -> str:
    try:
        if not user_text or not str(user_text).strip():
            return "Je peux analyser tes stocks, les risques de rupture et suggérer des quantités à commander."
        lang = _detect_lang(user_text)
        sample = _clean_df_for_advice(df if isinstance(df, pd.DataFrame) else pd.DataFrame())
        fields = ['product_name_display','supplier_name','abc_class','xyz_class','current_stock','avg_daily_sales',
                  'coverage_days','leadtime_days','credit_days','rupture_ml','delisting_product']
        view = sample[[c for c in fields if c in sample.columns]].copy() if not sample.empty else pd.DataFrame()

        cov_med = None; rows_digest = []
        try:
            if not view.empty:
                cov_med = float(pd.to_numeric(view.get('coverage_days', pd.Series(dtype=float)), errors='coerce').median())
                at_risk = view[view.get('rupture_ml','NON').astype(str).str.upper().eq('OUI')] if not view.empty else pd.DataFrame()
                top_list = at_risk.sort_values('coverage_days', ascending=True).head(6) if not at_risk.empty else view.sort_values('coverage_days', ascending=True).head(6)
                for _, r in (top_list if isinstance(top_list, pd.DataFrame) else pd.DataFrame()).iterrows():
                    rows_digest.append({
                        "name": str(r.get('product_name_display','')),
                        "sup": str(r.get('supplier_name','')),
                        "cov": float(pd.to_numeric(r.get('coverage_days',0), errors='coerce') or 0),
                        "lt": int(pd.to_numeric(r.get('leadtime_days',0), errors='coerce') or 0),
                        "abc": str(r.get('abc_class','')), "xyz": str(r.get('xyz_class','')),
                        "cred": int(pd.to_numeric(r.get('credit_days',0), errors='coerce') or 0),
                    })
        except Exception:
            pass

        table_json = []
        try: table_json = view.head(80).to_dict(orient="records")
        except Exception: pass

        if openai_client is None:
            return _chatbot_fallback(user_text, view)

        if lang == "fr":
            system_msg = (
                "Tu es Tony, assistant Supply Chain senior. Réponds brièvement et orienté action. "
                "Appuie tes recommandations sur le tableau fourni (produits). Cite des SKU si pertinent. "
                "Propose des quantités, priorités fournisseurs, et leviers (crédit, promo déstockage, ABC/XYZ)."
            )
            ctx = (f"Contexte (JSON extrait <=80 lignes): {json.dumps(table_json, ensure_ascii=False)[:12000]}\n" +
                   (f"KPI: couverture médiane ≈ {cov_med:.1f} j\n" if cov_med is not None else ""))
            if rows_digest:
                ctx += "Priorités (extrait): " + "; ".join(
                    [f"{d['name']} (sup {d['sup']}) cov {d['cov']:.1f}j, LT {d['lt']}j, crédit {d['cred']}j, {d['abc']}{d['xyz']}" for d in rows_digest]
                ) + "\n"
            user_q = f"Question: {user_text.strip()}"
        else:
            system_msg = (
                "You are Tony, a senior Supply Chain assistant. Be concise and action-oriented. "
                "Ground recommendations in the provided product table. Mention concrete SKUs when helpful."
            )
            ctx = (f"Context (JSON excerpt <=80 rows): {json.dumps(table_json, ensure_ascii=False)[:12000]}\n" +
                   (f"KPI: median coverage ≈ {cov_med:.1f} d\n" if cov_med is not None else ""))
            if rows_digest:
                ctx += "Priorities (excerpt): " + "; ".join(
                    [f"{d['name']} (sup {d['sup']}) cov {d['cov']:.1f}d, LT {d['lt']}d, credit {d['cred']}d, {d['abc']}{d['xyz']}" for d in rows_digest]
                ) + "\n"
            user_q = f"User question: {user_text.strip()}"

        hist = [{"role": "assistant" if m.get("role")=="assistant" else "user", "content": m.get("text","")} for m in (history_messages or [])]
        messages = [{"role":"system","content":system_msg}] + hist + [{"role":"user","content": ctx + "\n" + user_q}]
        try:
            resp = openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=messages, temperature=0.4, max_tokens=600
            )
            out = (resp.choices[0].message.content or "").strip()
            return out if out else _chatbot_fallback(user_text, view)
        except Exception as api_err:
            return _chatbot_fallback(user_text, view, prefix=f"⚠️ IA: {type(api_err).__name__} ")
    except Exception as e:
        return f"⚠️ Erreur inattendue: {type(e).__name__}: {e}"

def _chatbot_fallback(user_text: str, df: pd.DataFrame, prefix: str = "") -> str:
    lang = _detect_lang(user_text)
    d = _clean_df_for_advice(df)
    if d.empty:
        return prefix + ("Données insuffisantes. Recharge les sources." if lang=="fr" else "Insufficient data. Please reload sources.")
    try:
        top_risk = d.sort_values(['risk','coverage_days'], ascending=[False, True]).head(5)
        cnt_risk = int(d['risk'].sum())
        cov_med = float(pd.to_numeric(d['coverage_days'], errors='coerce').median())
        lines = [f"{prefix}" + (f"Risque: {cnt_risk} prod. • Couverture médiane ~ {cov_med:.1f} j." if lang=="fr"
                                else f"Risk: {cnt_risk} SKUs • Median coverage ~ {cov_med:.1f}d.")]
        for _, r in top_risk.iterrows():
            if lang=="fr":
                lines.append(f"- {r.get('product_name_display','?')} · cov {float(r.get('coverage_days',0)):.1f}j · LT {int(r.get('leadtime_days',0))}j · crédit {int(r.get('credit_days',0))}j")
            else:
                lines.append(f"- {r.get('product_name_display','?')} · cov {float(r.get('coverage_days',0)):.1f}d · LT {int(r.get('leadtime_days',0))}d · credit {int(r.get('credit_days',0))}d")
        lines += [("Cibles: A/AX en Y/Z <7j; crédit>14j si LT>20j; promos classe C." if lang=="fr"
                   else "Focus A/A+ in Y/Z <7d; credit>14d if LT>20d; promo bundles for class C.")]
        return "\n".join(lines)
    except Exception:
        return prefix + ("Recommandation impossible avec les données disponibles." if lang=="fr" else "Unable to compute recommendation with available data.")

# ------------------------------ Layout root (with floating chat) ------------------
initial_df = get_df_cached()
app.layout = html.Div([
    dcc.Location(id="url"),
    dcc.Store(id="master-data", data=initial_df.to_json(orient="records")),
    dcc.Store(id="filtered-data"),
    dcc.Store(id="uploaded-csv"),
    dcc.Store(id="chat-store", data=[]),
    dcc.Store(id="chat-open", data=False),

    make_sidebar(),
    html.Div(id="page-container", children=page_overview(initial_df)),

    html.Button(id="chat-fab", className="chat-fab", children=[html.Span("Assistant"), html.Span("💬")]),
    html.Div(id="chat-window", className="chat-window", style={"display":"none"}, children=[
        html.Div(className="chat-header", children=[
            html.Strong("Assistant "),
            html.Div([dbc.Button("Minimiser", id="chat-close", size="sm", className="btn-primary")])
        ]),
        html.Div(id="chat-messages", className="chat-body"),
        html.Div(style={"padding":"10px"}, children=[
            dcc.Upload(id="chat-upload",
                       children=html.Div(["📤 Glisser-déposer un CSV ici ou ", html.B("cliquer pour sélectionner")]),
                       multiple=False, className="upload-box"),
            html.Small(id="upload-status", style={"color":"#6b7280"})
        ]),
        html.Div(className="chat-input-wrap", children=[
            dbc.Textarea(id="chat-input", className="chat-textarea",
                         placeholder="Écrire une question… (Enter = envoyer, Shift+Enter = nouvelle ligne)", rows=2),
            dbc.Button("Envoyer", id="chat-send", className="btn-primary", n_clicks=0)
        ])
    ]),
])

# Validation layout
app.validation_layout = html.Div([
    dcc.Location(id="url"),
    dcc.Store(id="master-data"),
    dcc.Store(id="filtered-data"),
    dcc.Store(id="uploaded-csv"),
    dcc.Store(id="chat-store"),
    dcc.Store(id="chat-open"),
    make_sidebar(),
    page_overview(initial_df),
    page_analytics(),
    page_predictions(),
    page_about(),
    html.Div(id="page-container"),
    html.Button(id="chat-fab"),
    html.Div(id="chat-window"),
    html.Div(id="chat-messages"),
    dbc.Textarea(id="chat-input"),
    dcc.Upload(id="chat-upload"),
    html.Small(id="upload-status"),
    dbc.Button(id="chat-send"),
    dbc.Button(id="chat-close"),
    dcc.Download(id="download-data"),
    dcc.Download(id="download-po"),
    dbc.Button(id="btn-add-row"),
    dbc.Modal(id="edit-modal"),
    dbc.Input(id="edit-product"),
    dbc.Input(id="edit-supplier"),
    dbc.Input(id="edit-category"),
    dbc.Input(id="edit-stock"),
])

# ------------------------------ Routing ------------------------------------------
@app.callback(
    Output("page-container", "children"),
    Input("url", "pathname"),
    State("master-data","data"),
    prevent_initial_call=False
)
def render_page(path, master_json):
    base = pd.DataFrame(json.loads(master_json)) if master_json else get_df_cached()
    if path == "/analytics":
        return page_analytics()
    elif path == "/predictions":
        return page_predictions()
    elif path == "/about":
        return page_about()
    return page_overview(base)

# ------------------------------ Filtering logic ----------------------------------
def filter_dataframe(df: pd.DataFrame, query: str, suppliers: list, statuses: list, cats: list, risk_only: bool):
    out = df.copy()
    if query:
        q = str(query).strip().lower()
        if 'product_name' in out.columns:
            out = out[out['product_name'].astype(str).str.contains(q, na=False)]
    if suppliers:
        out = out[out['Supplier'].isin(suppliers)] if 'Supplier' in out.columns else out
    if statuses:
        out = out[out['Stock Status'].isin(statuses)] if 'Stock Status' in out.columns else out
    if cats:
        out = out[out['Product Category'].isin(cats)] if 'Product Category' in out.columns else out
    if risk_only and 'Stock Status' in out.columns:
        out = out[out['Stock Status'].isin(['Out of Stock','Predicted Stockout Soon'])]
    return out

@app.callback(
    [Output("filtered-data","data"), Output("main-table","data"), Output("risk-banner","children")],
    [Input("search-input","value"), Input("filter-supplier","value"), Input("filter-status","value"),
     Input("filter-category","value"), Input("toggle-risk-only","value")],
    State("master-data","data"),
    prevent_initial_call=False
)
def apply_filters(search, sup, stat, cat, risk_toggle, master_json):
    base = pd.DataFrame(json.loads(master_json)) if master_json else get_df_cached()
    sup = sup or []; stat = stat or []; cat = cat or []
    risk_only = ('risk' in (risk_toggle or []))
    fdf = filter_dataframe(base, search, sup, stat, cat, risk_only)
    risk_count = int((fdf['Stock Status'].isin(['Out of Stock','Predicted Stockout Soon']).sum())) if 'Stock Status' in fdf.columns else 0
    banner = [html.B("Alerte Rupture : "), f"{risk_count} SKU(s) à risque dans la vue filtrée — ",
              html.Span("OOS", className="badge badge-danger"), " / ", html.Span("Rupture imminente", className="badge-warn")]
    fdf_actions = add_action_cols(fdf)
    return fdf_actions.to_json(orient="records"), fdf_actions.to_dict("records"), banner

# ------------------------------ Export CSV ---------------------------------------
@app.callback(
    Output("download-data","data"),
    Input("btn-export","n_clicks"),
    State("filtered-data","data"),
    prevent_initial_call=True
)
def export_csv(n, data_json):
    if not n or not data_json: return no_update
    df = pd.DataFrame(json.loads(data_json))
    for c in ["✏️ Edit","🗑️ Delete"]:
        if c in df.columns: df.drop(columns=[c], inplace=True)
    return dcc.send_data_frame(df.to_csv, f"supply_filtered_{datetime.now().strftime('%Y%m%d_%H%M')}.csv", index=False)

# ------------------------------ Export Purchase Order PDF (1 produit sélectionné) -----
@app.callback(
    Output("download-po", "data"),
    Input("btn-po-pdf", "n_clicks"),
    State("main-table", "active_cell"),
    State("main-table", "selected_rows"),
    State("main-table", "data"),
    prevent_initial_call=True
)
def export_po_pdf(n, active_cell, selected_rows, table_data):
    if not n or not table_data:
        return no_update

    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.platypus import (SimpleDocTemplate, Paragraph, Table,
                                        TableStyle, Spacer, Image)
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib.units import mm

        # 1) Déterminer la ligne choisie (active_cell ou sélection de ligne)
        row_idx = None
        if active_cell and isinstance(active_cell, dict):
            row_idx = active_cell.get("row")
        if (row_idx is None) and selected_rows:
            row_idx = selected_rows[0]
        if row_idx is None or row_idx < 0 or row_idx >= len(table_data):
            return no_update

        r = table_data[row_idx]

        # Champs attendus
        prod_name = str(r.get("product_name", "")).strip()
        supplier  = str(r.get("Supplier", "")).strip()
        qty       = r.get("Predicted Order Quantity", 0)
        unit_ht   = r.get("Unit Price HT", 0.0)
        remise    = r.get("Discount", 0.0)

        if not prod_name or not supplier:
            return no_update

        # Normalisations numériques
        try: qty = int(pd.to_numeric(qty, errors="coerce") or 0)
        except Exception: qty = 0
        try: unit_ht = float(pd.to_numeric(unit_ht, errors="coerce") or 0.0)
        except Exception: unit_ht = 0.0
        try: remise = float(pd.to_numeric(remise, errors="coerce") or 0.0)
        except Exception: remise = 0.0
        if qty <= 0:
            qty = 1  # on force 1 pour éviter un BC vide

        # 2) Calculs
        taux_tva  = DEFAULT_TVA_RATE
        total_ht  = (qty * unit_ht) * (1 - remise/100.0)
        total_ttc = total_ht * (1 + taux_tva)

        # ref code simple depuis le nom
        def ref_from_name(name: str) -> str:
            if not name: return ""
            parts = [p for p in str(name).split() if p]
            if not parts: return str(name)[:6].upper()
            left = (parts[0][:3] if len(parts[0])>=3 else parts[0]).upper()
            right = (parts[1][:3] if len(parts)>1 and len(parts[1])>=3 else (parts[0][3:6] if len(parts[0])>3 else "")).upper()
            return "-".join([left, right]) if right else left

        ref_code = ref_from_name(prod_name)

        # 3) PDF
        buf = io.BytesIO()
        po_number = get_next_po_number()
        fname = f"{po_number}_{ref_code}.pdf"

        doc = SimpleDocTemplate(buf, pagesize=A4, title="Bon de commande")
        styles = getSampleStyleSheet()
        story = []

        # Header : logo + entreprise
        header_row = []
        logo_src = get_logo_for_reportlab()
        if logo_src:
            logo_img = Image(logo_src)
            logo_img.drawHeight = 18 * mm
            logo_img.drawWidth = 18 * mm
            header_row.append(logo_img)
        else:
            header_row.append(Paragraph("", styles["Normal"]))

        company_lines = [f"<b>{COMPANY_NAME}</b>"]
        if COMPANY_CAPITAL: company_lines.append(f"Montant du capital social : {COMPANY_CAPITAL}")
        if COMPANY_RCS:     company_lines.append(f"N° et lieu RCS : {COMPANY_RCS}")
        if COMPANY_ADDRESS: company_lines.append(f"Adresse du siège : {COMPANY_ADDRESS}")
        if COMPANY_PHONE:   company_lines.append(f"Téléphone : {COMPANY_PHONE}")
        if COMPANY_EMAIL:   company_lines.append(f"Email : {COMPANY_EMAIL}")
        header_row.append(Paragraph("<br/>".join(company_lines), styles["Normal"]))

        header_tbl = Table([header_row], colWidths=[25*mm, 150*mm])
        header_tbl.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ]))
        story.append(header_tbl)
        story.append(Spacer(1, 8))

        # Titre + méta
        story.append(Paragraph("<b>BON DE COMMANDE</b>", styles["Title"]))
        story.append(Spacer(1, 6))
        meta_left = [
            f"Bon de commande N° : <b>{po_number}</b>",
            f"Date : {datetime.now().strftime('%d/%m/%Y')}",
        ]
        meta_right = ["Fournisseur :", f"<b>{supplier}</b>"] if supplier else []
        meta_tbl = Table([[Paragraph("<br/>".join(meta_left), styles["Normal"]),
                           Paragraph("<br/>".join(meta_right), styles["Normal"])]],
                         colWidths=[100*mm, 75*mm])
        meta_tbl.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP")]))
        story.append(meta_tbl)
        story.append(Spacer(1, 10))

        # Tableau lignes
        headers = ["REF","DESCRIPTION","QUANTITÉ","PU HT","REMISE","TOTAL HT","TAUX TVA","TOTAL TTC"]
        data_tbl = [headers, [
            ref_code, prod_name, qty, f"{unit_ht:.2f}", f"{remise:.1f}%",
            f"{total_ht:.2f}", f"{int(taux_tva*100)}%", f"{total_ttc:.2f}"
        ], ["","TOTAL", qty, "", "", f"{total_ht:.2f}", "", f"{total_ttc:.2f}"]]

        tbl = Table(data_tbl, hAlign="LEFT",
                    colWidths=[25*mm,65*mm,20*mm,20*mm,20*mm,25*mm,20*mm,25*mm])
        tbl.setStyle(TableStyle([
            ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (2,1), (2,-1), "RIGHT"),
            ("ALIGN", (3,1), (-1,-1), "RIGHT"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("BOTTOMPADDING", (0,0), (-1,0), 6),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 12))

        # Conditions
        for title, content in [
            ("Conditions de livraison :", "À préciser (lieu, délai, incoterm)."),
            ("Conditions de règlement :", "À préciser (échéance, mode, pénalités)."),
            ("Délai de rétractation :", "Selon conditions générales applicables."),
        ]:
            story.append(Paragraph(f"<b>{title}</b>", styles["Normal"]))
            story.append(Paragraph(content, styles["Normal"]))
            story.append(Spacer(1, 6))

        doc.build(story)
        buf.seek(0)
        return dcc.send_bytes(lambda b: b.write(buf.getvalue()), filename=fname)

    except ModuleNotFoundError:
        print("ReportLab non installé: pip install reportlab")
        return no_update
    except Exception as e:
        print("PDF error:", repr(e))
        return no_update

# Activer le bouton PO si une cellule OU une ligne est sélectionnée et contient product_name + Supplier
@app.callback(
    Output("btn-po-pdf", "disabled"),
    [Input("main-table", "active_cell"),
     Input("main-table", "selected_rows"),
     Input("main-table", "data")],
    prevent_initial_call=False
)
def toggle_po_button(active_cell, selected_rows, data):
    if not data:
        return True

    # déterminer la ligne sélectionnée
    row_idx = None
    if active_cell and isinstance(active_cell, dict):
        row_idx = active_cell.get("row")
    if (row_idx is None) and selected_rows:
        row_idx = selected_rows[0]

    if row_idx is None or row_idx < 0 or row_idx >= len(data):
        return True

    r = data[row_idx] or {}
    prod = str(r.get("product_name", "")).strip()
    sup  = str(r.get("Supplier", "")).strip()
    return not (prod and sup)

# ------------------------------ Predictions What-If -------------------------------
@app.callback(
    Output("pred-table","data"),
    [Input("buffer-mult","value"), Input("lead-mult","value"), State("filtered-data","data")],
    prevent_initial_call=False
)
def recompute_suggestions(buffer_mult, lead_mult, filtered_json):
    df = get_df_cached()
    base = pd.DataFrame(json.loads(filtered_json)) if filtered_json else add_action_cols(df)
    out = base.copy()
    for c in ["✏️ Edit","🗑️ Delete"]:
        if c in out.columns: out.drop(columns=[c], inplace=True)
    for col in ['Max Daily Sales (Pikine)','Buffer Value','Max Avg Daily Sales','total_stock','Optimal Stock (Reorder Point)','Max Lead Time']:
        if col not in out.columns: out[col] = 0
    new_reorder = np.maximum(
        out['Max Daily Sales (Pikine)'].fillna(0),
        (out['Max Daily Sales (Pikine)'].fillna(0)/2.0) + (out['Buffer Value'].fillna(0) * float(buffer_mult)) * out.get('Max Avg Daily Sales',0).fillna(0)
    )
    anticip = (out['Max Lead Time'].fillna(0) * out['Max Avg Daily Sales'].fillna(0)) * (float(lead_mult) - 1.0)
    suggested = (new_reorder - out['total_stock'].fillna(0) + anticip).clip(lower=0)
    res = out[['product_name','Supplier','total_stock','Optimal Stock (Reorder Point)','Max Lead Time','Max Avg Daily Sales','Buffer Value']].copy()
    res['Suggested Order Qty'] = suggested.round(0).astype(int)
    return res.to_dict("records")

# ------------------------------ Table Actions (Edit/Delete/Add) -------------------
@app.callback(
    Output("edit-modal","is_open"),
    Output("edit-product","value"),
    Output("edit-supplier","value"),
    Output("edit-category","value"),
    Output("edit-stock","value"),
    Output("main-table","active_cell"),
    Input("btn-add-row","n_clicks"),
    Input("main-table","active_cell"),
    State("main-table","data"),
    prevent_initial_call=True
)
def open_edit_modal(n_add, active_cell, data):
    ctx = dash.callback_context
    if not ctx.triggered:
        return False, None, None, None, None, None
    trig = ctx.triggered[0]["prop_id"].split(".")[0]
    if trig == "btn-add-row":
        return True, "", "", "", 0, None
    if trig == "main-table" and active_cell:
        col = active_cell.get("column_id")
        row = active_cell.get("row")
        if col == "✏️ Edit" and data and 0 <= row < len(data):
            r = data[row]
            return True, r.get("product_name",""), r.get("Supplier",""), r.get("Product Category",""), r.get("total_stock",0), None
    return dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update

@app.callback(
    Output("master-data","data", allow_duplicate=True),
    Output("filtered-data","data", allow_duplicate=True),
    Output("main-table","data", allow_duplicate=True),
    Output("risk-banner","children", allow_duplicate=True),
    Output("edit-modal","is_open", allow_duplicate=True),
    Input("edit-save","n_clicks"),
    State("edit-product","value"),
    State("edit-supplier","value"),
    State("edit-category","value"),
    State("edit-stock","value"),
    State("main-table","active_cell"),
    State("main-table","data"),
    State("search-input","value"),
    State("filter-supplier","value"),
    State("filter-status","value"),
    State("filter-category","value"),
    State("toggle-risk-only","value"),
    State("master-data","data"),
    prevent_initial_call=True
)
def save_edit(n, prod, sup, cat, stock, active_cell, table_data, q, fs, fst, fc, rt, master_json):
    base = pd.DataFrame(json.loads(master_json)) if master_json else get_df_cached()
    df = base.copy()

    if active_cell and active_cell.get("column_id") == "✏️ Edit" and active_cell.get("row") is not None and table_data:
        row = active_cell["row"]
        r = table_data[row]
        key_p = r.get("product_name")
        key_s = r.get("Supplier")
        idx = df[(df["product_name"].astype(str)==str(key_p)) & (df["Supplier"].astype(str)==str(key_s))].index
        if len(idx)>0:
            i = idx[0]
            if prod is not None: df.at[i,"product_name"] = prod
            if sup is not None: df.at[i,"Supplier"] = sup
            if cat is not None: df.at[i,"Product Category"] = cat
            if stock is not None:
                try: df.at[i,"total_stock"] = float(stock)
                except: pass
    else:
        new_row = {c: np.nan for c in df.columns}  # np.nan au lieu de None
        new_row["product_name"] = prod or ""
        new_row["Supplier"] = sup or ""
        new_row["Product Category"] = cat or ""
        try:
            new_row["total_stock"] = float(stock or 0)
        except:
            new_row["total_stock"] = 0.0
        # initialise proprement quelques colonnes numériques si elles existent
        for c in ["Optimal Stock (Reorder Point)", "Max Lead Time", "Max Avg Daily Sales", "Max Coverage Day",
                  "Daily OOS Rate (30d)", "Predicted Order Quantity"]:
            if c in df.columns and pd.isna(new_row.get(c)):
                new_row[c] = 0.0
        if "Predicted Stockout" in df.columns and pd.isna(new_row.get("Predicted Stockout")):
            new_row["Predicted Stockout"] = False
        if "Stock Status" in df.columns and pd.isna(new_row.get("Stock Status")):
            new_row["Stock Status"] = "Order Soon" if new_row["total_stock"] else "Out of Stock"

        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    # Refiltrer et renvoyer
    sup_list = fs or []; stat_list = fst or []; cat_list = fc or []
    risk_only = ('risk' in (rt or []))
    fdf = filter_dataframe(df, q, sup_list, stat_list, cat_list, risk_only)
    risk_count = int((fdf['Stock Status'].isin(['Out of Stock','Predicted Stockout Soon']).sum())) if 'Stock Status' in fdf.columns else 0
    banner = [html.B("Alerte Rupture : "), f"{risk_count} SKU(s) à risque dans la vue filtrée — ",
              html.Span("OOS", className="badge badge-danger"), " / ", html.Span("Rupture imminente", className="badge-warn")]
    fdf_actions = add_action_cols(fdf)
    return df.to_json(orient="records"), fdf_actions.to_json(orient="records"), fdf_actions.to_dict("records"), banner, False

@app.callback(
    Output("master-data","data", allow_duplicate=True),
    Output("filtered-data","data", allow_duplicate=True),
    Output("main-table","data", allow_duplicate=True),
    Output("risk-banner","children", allow_duplicate=True),
    Input("main-table","active_cell"),
    State("main-table","data"),
    State("search-input","value"),
    State("filter-supplier","value"),
    State("filter-status","value"),
    State("filter-category","value"),
    State("toggle-risk-only","value"),
    State("master-data","data"),
    prevent_initial_call=True
)
def delete_row(active_cell, table_data, q, fs, fst, fc, rt, master_json):
    if not active_cell or active_cell.get("column_id") != "🗑️ Delete" or not table_data:
        raise dash.exceptions.PreventUpdate
    base = pd.DataFrame(json.loads(master_json)) if master_json else get_df_cached()
    row = active_cell["row"]
    r = table_data[row]
    key_p = r.get("product_name")
    key_s = r.get("Supplier")
    df = base[~((base["product_name"].astype(str)==str(key_p)) & (base["Supplier"].astype(str)==str(key_s)))].copy()

    sup_list = fs or []; stat_list = fst or []; cat_list = fc or []
    risk_only = ('risk' in (rt or []))
    fdf = filter_dataframe(df, q, sup_list, stat_list, cat_list, risk_only)
    risk_count = int((fdf['Stock Status'].isin(['Out of Stock','Predicted Stockout Soon']).sum())) if 'Stock Status' in fdf.columns else 0
    banner = [html.B("Alerte Rupture : "), f"{risk_count} SKU(s) à risque dans la vue filtrée — ",
              html.Span("OOS", className="badge badge-danger"), " / ", html.Span("Rupture imminente", className="badge-warn")]
    fdf_actions = add_action_cols(fdf)
    return df.to_json(orient="records"), fdf_actions.to_json(orient="records"), fdf_actions.to_dict("records"), banner

# ------------------------------ Floating Chat callbacks ---------------------------
@app.callback(
    Output("chat-open","data"),
    [Input("chat-fab","n_clicks"), Input("chat-close","n_clicks")],
    State("chat-open","data"),
    prevent_initial_call=False
)
def toggle_chat(n_fab, n_close, is_open):
    is_open = bool(is_open)
    ctx = dash.callback_context
    if not ctx.triggered:
        return is_open
    trig = ctx.triggered[0]["prop_id"].split(".")[0]
    if trig == "chat-close": return False
    if trig == "chat-fab": return not is_open
    return is_open

@app.callback(
    Output("chat-window","style"),
    Input("chat-open","data")
)
def show_hide_chat(is_open):
    return {"display": "flex" if is_open else "none"}

def parse_contents(content):
    import base64
    content_type, content_string = content.split(',')
    decoded = base64.b64decode(content_string)
    try:
        df = pd.read_csv(io.BytesIO(decoded)); return df
    except Exception:
        try:
            df = pd.read_excel(io.BytesIO(decoded)); return df
        except Exception:
            return pd.DataFrame()

@app.callback(
    Output("uploaded-csv","data"),
    Output("upload-status","children"),
    Output("chat-messages","children", allow_duplicate=True),
    Output("chat-store","data", allow_duplicate=True),
    Input("chat-upload","contents"),
    State("chat-upload","filename"),
    State("chat-store","data"),
    State("chat-messages","children"),
    prevent_initial_call=True
)
def on_upload(contents, filename, history, rendered):
    if not contents:
        raise dash.exceptions.PreventUpdate
    df_u = parse_contents(contents)
    if df_u.empty:
        status = f"⚠️ Échec de lecture du fichier {filename or ''}."
        return no_update, status, rendered, history
    history = history or []
    msg = f"📥 Fichier chargé : **{filename}** — {len(df_u)} lignes détectées. L’assistant l’utilisera pour ses conseils."
    history.append({"role":"assistant","text":msg,"ts":datetime.now().isoformat()})
    return df_u.to_json(orient="records"), f"✅ {filename} importé.", _render_messages(history), history

@app.callback(
    Output("chat-messages","children", allow_duplicate=True),
    Output("chat-store","data", allow_duplicate=True),
    Input("chat-send","n_clicks"),
    State("chat-input","value"),
    State("chat-store","data"),
    State("uploaded-csv","data"),
    State("master-data","data"),
    prevent_initial_call=True
)
def on_chat(n_clicks, user_text, history, uploaded_json, master_json):
    history = history or []
    user_text = (user_text or "").strip()
    if not user_text:
        return _render_messages(history), history
    df_up = pd.DataFrame(json.loads(uploaded_json)) if uploaded_json else pd.DataFrame()
    use_uploaded = ('product_name' in df_up.columns) and len(df_up)>0
    df_base = pd.DataFrame(json.loads(master_json)) if master_json else get_df_cached()
    df = df_up if use_uploaded else df_base

    history.append({"role":"user","text":user_text,"ts":datetime.now().isoformat()})
    reply = _chatbot_reply(user_text, df, history)
    history.append({"role":"assistant","text":reply,"ts":datetime.now().isoformat()})
    return _render_messages(history), history

# ------------------------------ Run ----------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8050))
    app.run(debug=True, host="0.0.0.0", port=port)

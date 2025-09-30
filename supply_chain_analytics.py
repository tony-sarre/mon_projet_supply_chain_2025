# ========================= PREMIUM SUPPLY CHAIN DASHBOARD =========================
# Requirements:
# pip install dash==2.17.1 dash-bootstrap-components==1.6.0 plotly==5.22.0
# pip install pandas scikit-learn flask-caching numpy
# pip install reportlab
# Optional: pip install openai
from dash import Dash
import dash_bootstrap_components as dbc
import os

app = Dash(__name__, suppress_callback_exceptions=True, external_stylesheets=[dbc.themes.BOOTSTRAP])

# ⚠️ Très important pour Render/Gunicorn
server = app.server
import re
import io
import base64
from datetime import datetime
import json
import threading
from pathlib import Path
import numpy as np
import pandas as pd

from flask_caching import Cache
import dash
from dash import Dash, html, dcc, Input, Output, State, dash_table, no_update
from dash.dependencies import Input, Output, State, ALL
import dash_bootstrap_components as dbc
import plotly.express as px
import warnings



from openai import OpenAI

warnings.filterwarnings("ignore", message="Parsing dates.*ambiguous", category=DeprecationWarning)

# ------------- OpenAI client (clé hardcodée à ta demande) ----------------
#OPENAI_API_KEY_HARDCODED = "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"
api_key = os.getenv("OPENAI_API_KEY_HARDCODED")

openai_client = None
try:
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
COMPANY_CAPITAL = os.getenv("COMPANY_CAPITAL", "")
COMPANY_RCS = os.getenv("COMPANY_RCS", "")
COMPANY_ADDRESS = os.getenv("COMPANY_ADDRESS", "")
COMPANY_PHONE = os.getenv("COMPANY_PHONE", "")
COMPANY_EMAIL = os.getenv("COMPANY_EMAIL", "")
DEFAULT_TVA_RATE = float(os.getenv("COMPANY_TVA_RATE", "0.18"))  # 18% par défaut
def _get_logo_data_uri():
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
    try:
        from reportlab.lib.utils import ImageReader

        # Priorité 1 : Fichier local
        p = Path("logo_maad.jpg")
        if p.exists() and p.stat().st_size > 0:
            print(f"✅ Logo found: {p.absolute()}")
            return str(p)

        # Priorité 2 : Base64
        if LOGO_DATA_URI and "," in LOGO_DATA_URI:
            try:
                b64 = LOGO_DATA_URI.split(",", 1)[1] if LOGO_DATA_URI.startswith("data:") else LOGO_DATA_URI
                raw = base64.b64decode(b64)
                print(f"✅ Logo loaded from base64 ({len(raw)} bytes)")
                return ImageReader(io.BytesIO(raw))
            except Exception as e:
                print(f"⚠️ Base64 logo decode failed: {e}")

        print("⚠️ No logo available, continuing without")
        return None

    except Exception as e:
        print(f"❌ Logo loading error: {e}")
        return None

# ------------------------------ PO Numbering -------------------------------------
PO_COUNTER_PATH = Path("./po_counter.json")
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
    today = datetime.now().strftime("%Y%m%d")
    with PO_LOCK:
        st = _load_po_state()
        if st.get("date") != today:
            st = {"date": today, "seq": 1}
        else:
            st["seq"] = int(st.get("seq", 0)) + 1
        _save_po_state(st)
        return f"PO-{today}-{st['seq']:03d}"
def train_stockout_model_for_df(df_in: pd.DataFrame, threshold: float = 0.5):
    from sklearn.ensemble import RandomForestClassifier
    from sklearn.model_selection import train_test_split
    from sklearn.metrics import roc_auc_score
    from sklearn.preprocessing import OneHotEncoder
    from sklearn.compose import ColumnTransformer
    from sklearn.pipeline import Pipeline
    from sklearn.impute import SimpleImputer

    df = df_in.copy()
    if 'Stock Status' not in df.columns:
        df['Stockout Probability'] = np.nan
        return df, None

    y = (df['Stock Status'].astype(str) == "Out of Stock").astype(int)

    num_features = ['total_stock', 'Avg Daily Sales', 'Max Lead Time',
                    'Max Coverage Day', 'credit_days', 'Daily OOS Rate (7d)', 'Daily OOS Rate (30d)']
    cat_features = ['Supplier', 'Product Category']
    for c in num_features:
        if c not in df.columns: df[c] = 0
    for c in cat_features:
        if c not in df.columns: df[c] = ""

    X = df[num_features + cat_features].copy()
    if y.nunique() < 2 or len(df) < 40:
        if 'Optimal Stock (Reorder Point)' in df.columns:
            prob = (df['total_stock'] <= df['Optimal Stock (Reorder Point)']).astype(float)
        else:
            prob = pd.Series(0.0, index=df.index)
        df['Stockout Probability'] = prob
        df['Predicted Stockout'] = (df['Stockout Probability'] > threshold).astype(bool)
        return df, None

    preproc = ColumnTransformer([
        ("num", SimpleImputer(strategy="median"), num_features),
        ("cat", OneHotEncoder(handle_unknown="ignore"), cat_features)
    ])
    clf = RandomForestClassifier(
        n_estimators=600, max_depth=10, min_samples_leaf=2,
        class_weight="balanced", random_state=42, n_jobs=-1
    )
    model = Pipeline([("prep", preproc), ("clf", clf)])

    X_train, X_test, y_train, y_test = train_test_split(
        X, y, stratify=y, test_size=0.2, random_state=42
    )
    try:
        model.fit(X_train, y_train)
        proba_test = model.predict_proba(X_test)[:, 1]
        try:
            auc = roc_auc_score(y_test, proba_test)
            print(f"[Stockout RF] AUC hold-out = {auc:.3f} (seuil={threshold})")
        except Exception:
            pass
        df['Stockout Probability'] = model.predict_proba(X)[:, 1]
        df['Predicted Stockout'] = (df['Stockout Probability'] > threshold).astype(bool)
        return df, model
    except Exception as e:
        print("[Stockout RF] Train error:", repr(e))
        if 'Optimal Stock (Reorder Point)' in df.columns:
            prob = (df['total_stock'] <= df['Optimal Stock (Reorder Point)']).astype(float)
        else:
            prob = pd.Series(0.0, index=df.index)
        df['Stockout Probability'] = prob
        df['Predicted Stockout'] = (df['Stockout Probability'] > threshold).astype(bool)
        return df, None


import pandas as pd
import numpy as np
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import OneHotEncoder
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.impute import SimpleImputer
import warnings

warnings.filterwarnings('ignore')


def load_supply_data() -> pd.DataFrame:
    import numpy as np
    import pandas as pd
    from sklearn.ensemble import RandomForestClassifier
    from sklearn.model_selection import train_test_split
    from sklearn.metrics import precision_score

    # =========================
    # Helpers
    # =========================
    def clean_text_column(df, colname):
        if colname in df.columns:
            df[colname] = (
                df[colname]
                .astype(str)
                .fillna("")
                .str.lower()
                .str.strip()
            )
        return df

    def safe_numeric(s, default=0):
        return pd.to_numeric(s, errors="coerce").fillna(default)

    # =========================
    # URLs
    # =========================
    SUPPLIERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRTyAxh6v8o0FXV0r7f6ALPDgmeJNkjTZITjrEoKBHo2gs_f3iyV8sFk8fOzcAsUSkJMXBJCpJnhQKi/pub?gid=1015760114&single=true&output=csv"
    url_leadtime = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT0FO0s7-V6uloIHLDB8Nm5TiH-W7q7zJOaA_jnzQtTgMUp-WOOX6CQP33__djc4shJym4r0PSAQF6t/pub?gid=1347877260&single=true&output=csv"
    inventory_pikine_staging = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT-503n7w2ixefop7XeUnda3B76ui1jQRZKJHehhJ0WtumnpSzUVYjnvGv-_tFQ6jXayAcjJEAryQMv/pub?gid=1876150276&single=true&output=csv"
    sales_pikine = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT-503n7w2ixefop7XeUnda3B76ui1jQRZKJHehhJ0WtumnpSzUVYjnvGv-_tFQ6jXayAcjJEAryQMv/pub?gid=1493123930&single=true&output=csv"
    Tbh_7dsales = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT-503n7w2ixefop7XeUnda3B76ui1jQRZKJHehhJ0WtumnpSzUVYjnvGv-_tFQ6jXayAcjJEAryQMv/pub?gid=1080970598&single=true&output=csv"
    Tbh_30dsales_products = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT-503n7w2ixefop7XeUnda3B76ui1jQRZKJHehhJ0WtumnpSzUVYjnvGv-_tFQ6jXayAcjJEAryQMv/pub?gid=1655420642&single=true&output=csv"
    Product_category = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQbJqHsr6Kifee7I91YD-7-sCZDWgM5GvxCeN0OqUvZhok0j-kDywguqe5I61y97b-uBhHbWraTIrux/pub?gid=803048228&single=true&output=csv"
    DELISTING_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQax3ZQW2QhDLE-waDewtdD8x_Q5tpn2FWzVJftr9egik4_JF3s2ytSYJmXh55aUnp79vmF-XtkaTmN/pub?gid=1681543945&single=true&output=csv"
    PARAMETRES_REPLENISH_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRxM2QokFGadTdTRDE2pInLKP57QkwMdSDQS8L5nXoYL0YRu9HSHoFvsnQs_MHjcwXUVUm5puexguy8/pub?gid=1011110883&single=true&output=csv"
    SUPPLIER_CATEGORIZATION_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQax3ZQW2QhDLE-waDewtdD8x_Q5tpn2FWzVJftr9egik4_JF3s2ytSYJmXh55aUnp79vmF-XtkaTmN/pub?gid=1938047484&single=true&output=csv"

    # =========================
    # Load CSV
    # =========================
    suppliers_df = pd.read_csv(SUPPLIERS_URL)
    df_leadtime = pd.read_csv(url_leadtime)
    inventory_pikine_staging_df = pd.read_csv(inventory_pikine_staging, skiprows=1)
    sales_pikine_df = pd.read_csv(sales_pikine, header=1, low_memory=False)
    Tbh_7dsales_df = pd.read_csv(Tbh_7dsales)
    Tbh_30dsales_products_df = pd.read_csv(Tbh_30dsales_products)
    Product_category_df = pd.read_csv(Product_category)

    try:
        delisting_df = pd.read_csv(DELISTING_URL, skiprows=3, usecols=[1, 3])
        delisting_df.columns = ["product_name", "delisting_status"]
    except:
        delisting_df = pd.DataFrame(columns=["product_name", "delisting_status"])

    try:
        parametres_replenish_df = pd.read_csv(PARAMETRES_REPLENISH_URL)
        print("'Parametres Replenish' data loaded successfully.")
    except:
        parametres_replenish_df = pd.DataFrame()

    try:
        supplier_categorization_df = pd.read_csv(SUPPLIER_CATEGORIZATION_URL)
        print("'Supplier Categorization' data loaded successfully.")
    except Exception as e:
        print(f"Error loading 'Supplier Categorization' data: {e}")
        supplier_categorization_df = pd.DataFrame()

    # =========================
    # Harmonisation inventaire
    # =========================
    inv = inventory_pikine_staging_df.copy()
    inv.columns = (
        inv.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace("-", "_")
    )
    col_map = {
        "product_name": ["product_name", "produit", "nom_produit", "name"],
        "Supplier": ["supplier", "supplier.1", "fournisseur", "vendor"],
        "total_stock": ["total_stock", "stock_total", "qte_stock", "current_stock"],
    }
    for target, aliases in col_map.items():
        if target not in inv.columns:
            for alias in aliases:
                if alias in inv.columns:
                    inv.rename(columns={alias: target}, inplace=True)
                    break
        if target not in inv.columns:
            inv[target] = 0

    if "supplier" in inv.columns and "Supplier" not in inv.columns:
        inv.rename(columns={"supplier": "Supplier"}, inplace=True)

    inventory_pikine_staging_df = inv
    print("Colonnes inventaire après harmonisation:", inventory_pikine_staging_df.columns.tolist())

    # =========================
    # Clean text
    # =========================
    inventory_pikine_staging_df = clean_text_column(inventory_pikine_staging_df, "product_name")
    inventory_pikine_staging_df = clean_text_column(inventory_pikine_staging_df, "Supplier")
    Product_category_df = clean_text_column(Product_category_df, "product_name")
    sales_pikine_df = clean_text_column(sales_pikine_df, "product_name")
    delisting_df = clean_text_column(delisting_df, "product_name")
    if "2" in Tbh_7dsales_df.columns:
        Tbh_7dsales_df["2"] = Tbh_7dsales_df["2"].astype(str).str.lower().str.strip()
    if "2" in Tbh_30dsales_products_df.columns:
        Tbh_30dsales_products_df["2"] = Tbh_30dsales_products_df["2"].astype(str).str.lower().str.strip()

    # =========================
    # Total stock
    # =========================
    total_stock_df = (
        inventory_pikine_staging_df.groupby(["product_name", "Supplier"])["total_stock"]
        .sum()
        .reset_index()
    )
    print(f"Shape of total_stock_df before merge: {total_stock_df.shape}")

    # =========================
    # Max sales (unique par produit)
    # =========================
    if "product_name" in sales_pikine_df.columns and "sum" in sales_pikine_df.columns:
        max_sales_pikine = (
            sales_pikine_df.groupby("product_name", as_index=False)["sum"].max()
            .rename(columns={"sum": "Max Daily Sales (Pikine)"})
            .drop_duplicates(subset=["product_name"])
        )
    else:
        max_sales_pikine = pd.DataFrame(columns=["product_name", "Max Daily Sales (Pikine)"])

    final_stock_sales_df = total_stock_df.merge(
        max_sales_pikine, on="product_name", how="left", validate="m:1"
    )
    print(f"Shape of final_stock_sales_df after merging max sales: {final_stock_sales_df.shape}")

    # =========================
    # Product Category (unique par produit)
    # =========================
    if "product_name" in Product_category_df.columns and "CATEGORY ABC-XYZ" in Product_category_df.columns:
        product_category_subset = (
            Product_category_df[["product_name", "CATEGORY ABC-XYZ"]]
            .dropna(subset=["product_name"])
            .drop_duplicates(subset=["product_name"])
            .rename(columns={"CATEGORY ABC-XYZ": "Product Category"})
        )
        final_stock_sales_df = final_stock_sales_df.merge(
            product_category_subset, on="product_name", how="left", validate="m:1"
        )
        final_stock_sales_df["Product Category"] = final_stock_sales_df["Product Category"].fillna("CX")
    else:
        final_stock_sales_df["Product Category"] = "CX"
    print(f"Shape of final_stock_sales_df after merging product category: {final_stock_sales_df.shape}")

    # =========================
    # ADS 7d / ADS 30d + OOS 7d / OOS 30d - MERGE UNIQUE
    # =========================
    if {"2", "9"}.issubset(Tbh_7dsales_df.columns):
        ads7 = (
            Tbh_7dsales_df[["2", "9"]]
            .rename(columns={"2": "product_name", "9": "Average Daily Sales (7d)"})
            .drop_duplicates(subset=["product_name"])
        )
        ads7["Average Daily Sales (7d)"] = safe_numeric(ads7["Average Daily Sales (7d)"], 0)
    else:
        ads7 = pd.DataFrame(columns=["product_name", "Average Daily Sales (7d)"])

    if {"2", "9"}.issubset(Tbh_30dsales_products_df.columns):
        ads30 = (
            Tbh_30dsales_products_df[["2", "9"]]
            .rename(columns={"2": "product_name", "9": "Average Daily Sales (30d)"})
            .drop_duplicates(subset=["product_name"])
        )
        ads30["Average Daily Sales (30d)"] = safe_numeric(ads30["Average Daily Sales (30d)"], 0)
    else:
        ads30 = pd.DataFrame(columns=["product_name", "Average Daily Sales (30d)"])

    if {"2", "12"}.issubset(Tbh_7dsales_df.columns):
        oos7 = (
            Tbh_7dsales_df[["2", "12"]]
            .rename(columns={"2": "product_name", "12": "Daily OOS Rate (7d)"})
            .drop_duplicates(subset=["product_name"])
        )
        oos7["Daily OOS Rate (7d)"] = safe_numeric(
            oos7["Daily OOS Rate (7d)"].astype(str).str.replace("%", "", regex=False), 100
        ) / 100.0
    else:
        oos7 = pd.DataFrame(columns=["product_name", "Daily OOS Rate (7d)"])

    if {"2", "28"}.issubset(Tbh_30dsales_products_df.columns):
        oos30 = (
            Tbh_30dsales_products_df[["2", "28"]]
            .rename(columns={"2": "product_name", "28": "Daily OOS Rate (30d)"})
            .drop_duplicates(subset=["product_name"])
        )
        oos30["Daily OOS Rate (30d)"] = safe_numeric(
            oos30["Daily OOS Rate (30d)"].astype(str).str.replace("%", "", regex=False), 0
        ) / 100.0
    else:
        oos30 = pd.DataFrame(columns=["product_name", "Daily OOS Rate (30d)"])

    # MERGE UNIQUE
    final_stock_sales_df = final_stock_sales_df.merge(ads7, on="product_name", how="left", validate="m:1")
    final_stock_sales_df = final_stock_sales_df.merge(ads30, on="product_name", how="left", validate="m:1")
    final_stock_sales_df = final_stock_sales_df.merge(oos7, on="product_name", how="left", validate="m:1")
    final_stock_sales_df = final_stock_sales_df.merge(oos30, on="product_name", how="left", validate="m:1")

    # =========================
    # CALCUL UNIQUE DE RECALCULATED AVERAGE DAILY SALES
    # =========================
    def calc_recalculated_ads(row):
        oos_7d = row.get("Daily OOS Rate (7d)", 1.0)
        oos_30d = row.get("Daily OOS Rate (30d)", 1.0)
        sales_7d = row.get("Average Daily Sales (7d)", np.nan)
        sales_30d = row.get("Average Daily Sales (30d)", np.nan)

        if pd.isna(sales_7d) and pd.isna(sales_30d):
            return 0.1

        if pd.isna(sales_7d) or pd.isna(oos_7d) or (oos_7d >= 0.6):
            if pd.notna(sales_30d):
                if pd.notna(oos_30d) and (oos_30d >= 0.6):
                    if pd.notna(sales_7d):
                        return (sales_7d + sales_30d) / 2.0 + 0.1
                    else:
                        return sales_30d + 0.1
                else:
                    return sales_30d + 0.1
            else:
                return (sales_7d + 0.1) if pd.notna(sales_7d) else 0.1

        return sales_7d + 0.1

    final_stock_sales_df["Recalculated Average Daily Sales"] = (
        final_stock_sales_df.apply(calc_recalculated_ads, axis=1)
    )

    print(f"✅ Recalculated ADS - Stats:")
    print(f"   Min: {final_stock_sales_df['Recalculated Average Daily Sales'].min():.2f}")
    print(f"   Max: {final_stock_sales_df['Recalculated Average Daily Sales'].max():.2f}")
    print(f"   Médiane: {final_stock_sales_df['Recalculated Average Daily Sales'].median():.2f}")
    print(f"   Moyenne: {final_stock_sales_df['Recalculated Average Daily Sales'].mean():.2f}")

    # Max Coverage Day
    final_stock_sales_df["Max Coverage Day"] = np.minimum(
        safe_numeric(final_stock_sales_df["total_stock"], 0) /
        np.maximum(safe_numeric(final_stock_sales_df["Recalculated Average Daily Sales"], 0.01), 0.01),
        365
    )

    # =========================
    # Lead time (Supplier unique)
    # =========================
    if {"supplier", "avg_leadtime"}.issubset(df_leadtime.columns):
        lt = (
            df_leadtime[["supplier", "avg_leadtime"]]
            .rename(columns={"supplier": "Supplier", "avg_leadtime": "Avg Lead Time"})
        )
        lt["Supplier"] = lt["Supplier"].astype(str).str.lower().str.strip()
        lt = lt.drop_duplicates(subset=["Supplier"])
        final_stock_sales_df = final_stock_sales_df.merge(
            lt, on="Supplier", how="left", validate="m:1"
        )
    else:
        final_stock_sales_df["Avg Lead Time"] = np.nan

    # =========================
    # Supplier credit info (AVANT supplier categorization)
    # =========================
    if {"Supplier name", "credit_days", "Credit_cumulable"}.issubset(suppliers_df.columns):
        sc = suppliers_df[["Supplier name", "credit_days", "Credit_cumulable"]].copy()
        sc = sc.rename(columns={"Supplier name": "Supplier"})
        sc["Supplier"] = sc["Supplier"].astype(str).str.lower().str.strip()
        sc["credit_days"] = safe_numeric(sc["credit_days"], 0)
        sc["Credit_cumulable"] = sc["Credit_cumulable"].fillna("Non")
        sc = sc.drop_duplicates(subset=["Supplier"])
        final_stock_sales_df = final_stock_sales_df.merge(sc, on="Supplier", how="left", validate="m:1")
        final_stock_sales_df["credit_days"] = final_stock_sales_df["credit_days"].fillna(0)
        final_stock_sales_df["Credit_cumulable"] = final_stock_sales_df["Credit_cumulable"].fillna("Non")
        print("'credit_days' and 'Credit_cumulable' columns merged.")
    else:
        final_stock_sales_df["credit_days"] = 0
        final_stock_sales_df["Credit_cumulable"] = "Non"

    # =========================
    # SUPPLIER CATEGORIZATION → UPDATE PRODUCT CATEGORY
    # =========================
    # =========================
    # SUPPLIER CATEGORIZATION → UPDATE PRODUCT CATEGORY
    # =========================
    supplier_categorization_lookup_df = pd.DataFrame()
    if not supplier_categorization_df.empty and supplier_categorization_df.shape[1] > 7:
        supplier_categorization_lookup_df = supplier_categorization_df.iloc[:, [0, 7]].copy()
        supplier_categorization_lookup_df.columns = ['Supplier_Name_Lookup', 'Supplier_Categorization_Lookup']
        supplier_categorization_lookup_df['Supplier_Name_Lookup'] = (
            supplier_categorization_lookup_df['Supplier_Name_Lookup'].astype(str).str.lower().str.strip()
        )
        supplier_categorization_lookup_df = supplier_categorization_lookup_df.drop_duplicates(
            subset=['Supplier_Name_Lookup'])
        print(f"Supplier categorization loaded: {len(supplier_categorization_lookup_df)} suppliers")

    if not supplier_categorization_lookup_df.empty and 'Supplier' in final_stock_sales_df.columns:
        final_stock_sales_df['Supplier'] = final_stock_sales_df['Supplier'].astype(str).str.lower().str.strip()

        final_stock_sales_df = pd.merge(
            final_stock_sales_df,
            supplier_categorization_lookup_df,
            left_on='Supplier',
            right_on='Supplier_Name_Lookup',
            how='left'
        )

        if 'Supplier_Name_Lookup' in final_stock_sales_df.columns:
            final_stock_sales_df.drop(columns=['Supplier_Name_Lookup'], inplace=True)

        final_stock_sales_df['Supplier_Categorization_Lookup'] = (
            final_stock_sales_df['Supplier_Categorization_Lookup'].fillna('Unknown')
        )

        # RÈGLE MÉTIER CRITIQUE: Si credit_days > 0, utiliser supplier categorization
        if all(col in final_stock_sales_df.columns for col in
               ['credit_days', 'Product Category', 'Supplier_Categorization_Lookup']):

            final_stock_sales_df['credit_days'] = pd.to_numeric(
                final_stock_sales_df['credit_days'], errors='coerce'
            ).fillna(0)

            # AVANT modification
            original_categories = final_stock_sales_df['Product Category'].copy()

            # APPLICATION DE LA RÈGLE
            final_stock_sales_df['Product Category'] = final_stock_sales_df.apply(
                lambda row: row['Supplier_Categorization_Lookup'] if row['credit_days'] > 0 else row[
                    'Product Category'],
                axis=1
            )

            # LOGGING : compter les modifications
            modified_count = (final_stock_sales_df['Product Category'] != original_categories).sum()
            print(f"'Product Category' updated: {modified_count} products modified based on credit_days > 0")

            # Afficher échantillon
            if modified_count > 0:
                sample = final_stock_sales_df[final_stock_sales_df['Product Category'] != original_categories][
                    ['product_name', 'Supplier', 'credit_days', 'Product Category']
                ].head(5)
                print("Sample of modified products:")
                print(sample.to_string(index=False))

        if 'Supplier_Categorization_Lookup' in final_stock_sales_df.columns:
            final_stock_sales_df.drop(columns=['Supplier_Categorization_Lookup'], inplace=True)
    else:
        print("Warning: Supplier categorization data unavailable or incomplete.")
    # =========================
    # Parametres Replenish (Buffer lookup) - APRÈS update Product Category
    # =========================
    if not parametres_replenish_df.empty and parametres_replenish_df.shape[1] > 1:
        pr_buf = parametres_replenish_df.iloc[:, [0, 1]].copy()
        pr_buf.columns = ["Param_Product_Category", "Param_Buffer_Value_Lookup"]
        pr_buf["Param_Product_Category"] = pr_buf["Param_Product_Category"].astype(str).str.lower().str.strip()
        pr_buf["Param_Buffer_Value_Lookup"] = safe_numeric(pr_buf["Param_Buffer_Value_Lookup"], 0)
        final_stock_sales_df["Product Category"] = final_stock_sales_df["Product Category"].astype(
            str).str.lower().str.strip()
        final_stock_sales_df = final_stock_sales_df.merge(
            pr_buf.drop_duplicates(subset=["Param_Product_Category"]),
            left_on="Product Category",
            right_on="Param_Product_Category",
            how="left",
            validate="m:1",
        )
        if "Param_Product_Category" in final_stock_sales_df.columns:
            final_stock_sales_df.drop(columns=["Param_Product_Category"], inplace=True)
        final_stock_sales_df["Param_Buffer_Value_Lookup"] = final_stock_sales_df["Param_Buffer_Value_Lookup"].fillna(0)
        print("'Param_Buffer_Value_Lookup' merged for Optimal Stock calculation.")
    else:
        final_stock_sales_df["Param_Buffer_Value_Lookup"] = 0.0

    # Optimal Stock
    final_stock_sales_df["Max Daily Sales (Pikine)"] = safe_numeric(final_stock_sales_df["Max Daily Sales (Pikine)"], 0)
    final_stock_sales_df["Recalculated Average Daily Sales"] = safe_numeric(
        final_stock_sales_df["Recalculated Average Daily Sales"], 0.1)
    final_stock_sales_df["Param_Buffer_Value_Lookup"] = safe_numeric(final_stock_sales_df["Param_Buffer_Value_Lookup"],
                                                                     0)
    final_stock_sales_df["Optimal Stock (Reorder Point)"] = final_stock_sales_df.apply(
        lambda r: max(
            r["Max Daily Sales (Pikine)"],
            (r["Max Daily Sales (Pikine)"] / 2.0) + (
                        r["Param_Buffer_Value_Lookup"] * r["Recalculated Average Daily Sales"]),
        ),
        axis=1,
    )
    print("'Optimal Stock (Reorder Point)' calculated successfully.")

    # =========================
    # Delisting
    # =========================
    if {"product_name", "delisting_status"}.issubset(delisting_df.columns):
        delis = delisting_df[["product_name", "delisting_status"]].drop_duplicates(subset=["product_name"])
        final_stock_sales_df = final_stock_sales_df.merge(
            delis, on="product_name", how="left", validate="m:1"
        )
        final_stock_sales_df["delisting_status"] = final_stock_sales_df["delisting_status"].fillna("Not Delisted")
        print("'delisting_status' column merged.")
    else:
        final_stock_sales_df["delisting_status"] = "Not Delisted"

    # =========================
    # Dédupes
    # =========================
    final_stock_sales_df = final_stock_sales_df.drop_duplicates(subset=["product_name"], keep="first")
    print(f"Shape after removing duplicates: {final_stock_sales_df.shape}")

    # =========================
    # Métriques métier (suite exacte de votre code)
    # =========================
    final_stock_sales_df["Avg Lead Time"] = safe_numeric(final_stock_sales_df.get("Avg Lead Time", 0), 0)
    final_stock_sales_df["Credit_cumulable"] = final_stock_sales_df.get("Credit_cumulable", "Non").fillna("Non")

    def calc_adjusted_leadtime(row):
        return row["Avg Lead Time"] if str(row.get("Credit_cumulable", "")).lower() == "oui" else row[
                                                                                                      "Avg Lead Time"] + 3

    final_stock_sales_df["ADJUSTED_LEADTIME"] = final_stock_sales_df.apply(calc_adjusted_leadtime, axis=1)

    final_stock_sales_df["credit_days"] = safe_numeric(final_stock_sales_df["credit_days"], 0)

    def calc_max_credit_buffer(row):
        cd = row["credit_days"]
        cc = str(row.get("Credit_cumulable", "")).lower()
        return min(cd, 20) if cc == "oui" else cd

    final_stock_sales_df["MAX_CREDIT_BUFFER"] = final_stock_sales_df.apply(calc_max_credit_buffer, axis=1)

    final_stock_sales_df["AJUSTER_BUFFER"] = np.maximum(
        safe_numeric(final_stock_sales_df["Param_Buffer_Value_Lookup"], 0),
        safe_numeric(final_stock_sales_df["MAX_CREDIT_BUFFER"], 0),
    )

    final_stock_sales_df["MOQ MAAD"] = (safe_numeric(final_stock_sales_df["ADJUSTED_LEADTIME"], 0) + 3) * safe_numeric(
        final_stock_sales_df["Recalculated Average Daily Sales"], 0.1)

    if "Param_Supplier_Factor" not in final_stock_sales_df.columns:
        final_stock_sales_df["Param_Supplier_Factor"] = 0.0

    # Purchase Need
    req_cols = ["total_stock", "Max Daily Sales (Pikine)", "Recalculated Average Daily Sales",
                "Optimal Stock (Reorder Point)", "Param_Buffer_Value_Lookup", "Param_Supplier_Factor"]
    if all(c in final_stock_sales_df.columns for c in req_cols):
        for c in req_cols:
            final_stock_sales_df[c] = safe_numeric(final_stock_sales_df[c], 0)

        def calculate_purchase_need(row):
            total_stock = row["total_stock"]
            max_daily_sales = row["Max Daily Sales (Pikine)"]
            avg_daily_sales = row["Recalculated Average Daily Sales"]
            optimal_stock = row["Optimal Stock (Reorder Point)"]
            param_buffer_value = row["Param_Buffer_Value_Lookup"]
            param_supplier_factor = row["Param_Supplier_Factor"]
            if total_stock <= 0:
                return max_daily_sales + (param_buffer_value * avg_daily_sales)
            elif total_stock < optimal_stock + (param_supplier_factor * avg_daily_sales):
                return (optimal_stock - total_stock) + (param_supplier_factor * avg_daily_sales)
            else:
                return 0

        final_stock_sales_df["purchase_need"] = final_stock_sales_df.apply(calculate_purchase_need, axis=1)
        print("'purchase_need' calculated successfully.")
    else:
        final_stock_sales_df["purchase_need"] = np.nan

    final_stock_sales_df["QAC"] = np.maximum(
        safe_numeric(final_stock_sales_df["MOQ MAAD"], 0),
        safe_numeric(final_stock_sales_df["purchase_need"], 0),
    )

    # Predicted Order Quantity
    need_cols = ["total_stock", "Optimal Stock (Reorder Point)", "Recalculated Average Daily Sales", "credit_days",
                 "ADJUSTED_LEADTIME"]
    if all(c in final_stock_sales_df.columns for c in need_cols):
        for c in need_cols:
            final_stock_sales_df[c] = safe_numeric(final_stock_sales_df[c], 0)

        def calc_poq(row):
            stock = row["total_stock"]
            rp = row["Optimal Stock (Reorder Point)"]
            ads = row["Recalculated Average Daily Sales"]
            credit_days = row["credit_days"]
            alt = row["ADJUSTED_LEADTIME"]
            demand_period = alt + credit_days
            demand_during_period = ads * demand_period
            return max(0.0, (rp + demand_during_period) - stock)

        final_stock_sales_df["Predicted Order Quantity"] = final_stock_sales_df.apply(calc_poq, axis=1)
        print("'Predicted Order Quantity' calculated successfully.")
    else:
        final_stock_sales_df["Predicted Order Quantity"] = np.nan

        # Ajusted_total_need

    def calc_adjusted_total_need(row):
        if str(row.get("delisting_status", "")).lower() == "delisted":
            return "NO NEED"
        mcd = float(row.get("Max Coverage Day", 0))
        alt = float(row.get("ADJUSTED_LEADTIME", 0))
        opt = float(row.get("Optimal Stock (Reorder Point)", 0))
        ads = float(row.get("Recalculated Average Daily Sales", 0.1))
        optimal_days = (opt / ads) if ads > 0 else 0
        if mcd <= alt + 3:
            return "ORDER NOW"
        elif mcd < alt + optimal_days:
            return "ORDER NOT URGENT"
        else:
            return "NO NEED"

    final_stock_sales_df["Ajusted_total_need"] = final_stock_sales_df.apply(calc_adjusted_total_need, axis=1)

    # Predicted Stockout
    final_stock_sales_df["Predicted Stockout"] = safe_numeric(final_stock_sales_df["total_stock"], 0) <= safe_numeric(final_stock_sales_df["Optimal Stock (Reorder Point)"], 0)

    # Stock Status
    def get_stock_status(row):
        stock = float(row.get("total_stock", 0))
        reorder_point = float(row.get("Optimal Stock (Reorder Point)", 0))
        alt = float(row.get("ADJUSTED_LEADTIME", 0))
        ads = float(row.get("Recalculated Average Daily Sales", 0))
        if stock <= 0:
            return "Out of Stock"
        elif stock <= reorder_point:
            return "Predicted Stockout Soon"
        elif stock <= reorder_point + (alt * ads):
            return "Order Soon"
        else:
            return "Stock OK"
    final_stock_sales_df["Stock Status"] = final_stock_sales_df.apply(get_stock_status, axis=1)

    # Credit Adequacy
    for c in ["total_stock","Recalculated Average Daily Sales","ADJUSTED_LEADTIME","credit_days","Predicted Order Quantity"]:
        if c not in final_stock_sales_df.columns:
            final_stock_sales_df[c] = 0.0
        final_stock_sales_df[c] = safe_numeric(final_stock_sales_df[c], 0)

    cds = final_stock_sales_df["credit_days"].clip(lower=0)
    ads = final_stock_sales_df["Recalculated Average Daily Sales"].clip(lower=0)
    alt = final_stock_sales_df["ADJUSTED_LEADTIME"].clip(lower=0)
    s0  = final_stock_sales_df["total_stock"].clip(lower=0)
    q   = final_stock_sales_df["Predicted Order Quantity"].clip(lower=0)

    target_low  = ads * cds
    target_high = ads * (cds + alt)
    s_post = s0 + q
    gap_low = (target_low - s_post).clip(lower=0)
    excess  = (s_post - target_high).clip(lower=0)
    denom = target_low.where(target_low > 0, 1.0)
    err = (gap_low + excess) / denom
    credit_score = np.exp(-1.5 * err)

    final_stock_sales_df["Credit Adequacy Score"] = credit_score
    final_stock_sales_df["Credit Adequacy Risk"] = (err > 0)

    # Neutralisation si pas de crédit ou delisted
    no_credit_mask = (final_stock_sales_df["credit_days"] <= 0)
    final_stock_sales_df.loc[no_credit_mask, "Credit Adequacy Risk"] = False
    final_stock_sales_df.loc[no_credit_mask, "Credit Adequacy Score"] = np.where(
        ads.loc[no_credit_mask] > 0, 1.0, 0.0
    )
    dmask = final_stock_sales_df["delisting_status"].astype(str).str.lower().eq("delisted")
    final_stock_sales_df.loc[dmask, ["Predicted Order Quantity","Predicted Stockout","purchase_need","QAC","MOQ MAAD"]] = [0, False, 0, 0, 0]
    final_stock_sales_df.loc[dmask, "Ajusted_total_need"] = "NO NEED"
    final_stock_sales_df.loc[dmask, "Credit Adequacy Risk"] = False
    final_stock_sales_df.loc[dmask, "Credit Adequacy Score"] = 1.0

    # =========================
    # Nettoyage final
    # =========================
    for cdrop in ["Max Lead Time", "Buffer Value", "Param_Supplier_Factor", "Param_Buffer_Value_Lookup"]:
        if cdrop in final_stock_sales_df.columns:
            final_stock_sales_df.drop(columns=[cdrop], inplace=True, errors="ignore")

    print("Final Stock and Sales Analysis with all calculated columns:")
    print(f"NaNs total: {int(final_stock_sales_df.isnull().sum().sum())}")
    print("\nSTATISTIQUES DE CONTRÔLE:")
    print(f"• Total produits analysés: {len(final_stock_sales_df)}")
    print(f"• Produits avec stock > 0: {(final_stock_sales_df['total_stock'] > 0).sum()}")
    print(f"• Produits Out of Stock: {(final_stock_sales_df['total_stock'] <= 0).sum()}")
    print(f"• Predicted Stockouts: {int(final_stock_sales_df['Predicted Stockout'].sum())}")
    print(f"• Besoin d'achat total: {final_stock_sales_df['purchase_need'].sum():,.0f} unités")
    print(f"• Couverture moyenne: {final_stock_sales_df['Max Coverage Day'].mean():.1f} jours")
    print(f"• Produits delisted: {(final_stock_sales_df['delisting_status'].str.lower() == 'delisted').sum()}")

    # =========================
    # ML: Stockout Probability
    # =========================
    try:
        feature_cols = [
            "total_stock", "Recalculated Average Daily Sales",
            "Max Daily Sales (Pikine)", "Optimal Stock (Reorder Point)",
            "ADJUSTED_LEADTIME", "MAX_CREDIT_BUFFER", "AJUSTER_BUFFER"
        ]
        feature_cols = [c for c in feature_cols if c in final_stock_sales_df.columns]
        if feature_cols and "Predicted Stockout" in final_stock_sales_df.columns:
            df_train = final_stock_sales_df.dropna(subset=feature_cols + ["Predicted Stockout"]).copy()
            X = df_train[feature_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
            y = df_train["Predicted Stockout"].astype(int)
            if len(df_train) > 50 and y.nunique() > 1:
                X_train, X_test, y_train, y_test = train_test_split(
                    X, y, test_size=0.2, random_state=42, stratify=y
                )
                model = RandomForestClassifier(
                    n_estimators=200, max_depth=8, class_weight="balanced", random_state=42, n_jobs=-1
                )
                model.fit(X_train, y_train)
                y_pred = model.predict(X_test)
                prec = precision_score(y_test, y_pred)
                print(f"[ML Stockout] Modèle entraîné, précision test = {prec:.2f}")
                X_all = final_stock_sales_df[feature_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
                probs = model.predict_proba(X_all)[:, 1]
                final_stock_sales_df["Stockout Probability"] = probs
            else:
                final_stock_sales_df["Stockout Probability"] = 0.0
                print("[ML Stockout] Pas assez de données variées pour entraîner le modèle.")
        else:
            final_stock_sales_df["Stockout Probability"] = 0.0
            print("[ML Stockout] Colonnes nécessaires absentes ou Predicted Stockout manquant.")
    except Exception as e:
        print(f"[ML Stockout] Erreur lors du calcul ML: {e}")
        final_stock_sales_df["Stockout Probability"] = 0.0

    # =========================
    # Garantie colonnes critiques pour UI
    # =========================
    def _ensure_not_empty_string(s: pd.Series, default_val: str) -> pd.Series:
        return (
            s.where(~s.isna(), None)
            .apply(lambda x: None if (isinstance(x, str) and x.strip() == "") else x)
            .fillna(default_val)
            .astype(str).str.strip()
        )

    if "Supplier" not in final_stock_sales_df.columns:
        final_stock_sales_df["Supplier"] = "unknown"
    else:
        final_stock_sales_df["Supplier"] = _ensure_not_empty_string(final_stock_sales_df["Supplier"], "unknown")

    if "Recalculated Average Daily Sales" not in final_stock_sales_df.columns:
        final_stock_sales_df["Recalculated Average Daily Sales"] = 0.1

    final_stock_sales_df["Recalculated Average Daily Sales"] = (
        pd.to_numeric(final_stock_sales_df["Recalculated Average Daily Sales"], errors="coerce")
        .fillna(0.1)
        .clip(lower=0.1)
    )

    # Colonnes en tête pour visibilité
    _front = ["product_name", "Supplier", "Recalculated Average Daily Sales"]
    final_stock_sales_df = final_stock_sales_df[
        [c for c in _front if c in final_stock_sales_df.columns] +
        [c for c in final_stock_sales_df.columns if c not in _front]
    ]

    # =========================
    # TEST FINAL DE VALIDATION
    # =========================
    print("\n========== VÉRIFICATION FINALE RECALCULATED ADS ==========")
    sample = final_stock_sales_df[["product_name", "Average Daily Sales (7d)",
                                    "Average Daily Sales (30d)", "Daily OOS Rate (7d)",
                                    "Recalculated Average Daily Sales"]].head(10)
    print("Échantillon (10 premiers produits) :")
    print(sample.to_string())

    # À la fin de load_supply_data(), avant le return
    print("\n========== COLONNES DISPONIBLES POUR FILTRAGE ==========")
    print(f"✓ Supplier présent : {'Supplier' in final_stock_sales_df.columns}")
    print(f"✓ Product Category présent : {'Product Category' in final_stock_sales_df.columns}")
    print(f"✓ Ajusted_total_need présent : {'Ajusted_total_need' in final_stock_sales_df.columns}")
    print(f"✓ product_name présent : {'product_name' in final_stock_sales_df.columns}")

    if 'Ajusted_total_need' in final_stock_sales_df.columns:
        print(f"\nValeurs uniques de Ajusted_total_need :")
        print(final_stock_sales_df['Ajusted_total_need'].value_counts())
    print("========================================================\n")

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
            const attachEnterSend = () => {
                const ta = document.getElementById('chat-input');
                const btn = document.getElementById('chat-send');
                if(!ta || !btn) return;
                if(ta._boundEnter) return;
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
            window.debounce = (func, wait=280) => { let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>func.apply(this,a), wait);} };
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

    # Extraction des valeurs uniques pour les dropdowns
    suppliers = sorted(
        [s for s in df['Supplier'].dropna().unique().tolist() if s != '']) if 'Supplier' in df.columns else []
    cats = sorted([c for c in df['Product Category'].dropna().unique().tolist() if
                   c != '']) if 'Product Category' in df.columns else []

    return html.Div(className="sidebar", children=[
        html.Div([
            html.Div([
                html.Img(src=LOGO_DATA_URI,
                         style={"height": "42px", "marginRight": "8px"}) if LOGO_DATA_URI else html.Div(),
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
            dcc.Dropdown(
                id="filter-supplier",
                options=[{"label": s, "value": s} for s in suppliers],
                multi=True,
                placeholder="Tous",
                persistence=True
            ),
            html.Br(),
            html.Small("Catégorie"),
            dcc.Dropdown(
                id="filter-category",
                options=[{"label": c, "value": c} for c in cats],
                multi=True,
                placeholder="Toutes",
                persistence=True
            ),
            html.Br(),
            html.Small("Besoin d'achat"),
            dcc.Dropdown(
                id="filter-need",
                options=[
                    {"label": "ORDER NOW", "value": "ORDER NOW"},
                    {"label": "ORDER NOT URGENT", "value": "ORDER NOT URGENT"},
                    {"label": "NO NEED", "value": "NO NEED"}
                ],
                multi=True,
                placeholder="Tous",
                persistence=True
            ),
            html.Br(),
            dbc.Checklist(
                options=[
                    {"label": "Regrouper par produit (dé-dup)", "value": "by_product"},
                ],
                value=[],
                id="toggle-options",
                switch=True
            )
        ]),

        html.Br(),
        html.Div([
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
        html.Small(f"© {datetime.now().year} • {AUTHOR}"),
        html.Div(id="debug-info", style={"marginTop": "20px", "fontSize": "10px", "color": "#6b7280"})  # Pour debug
    ])

# ------------------------------ Aggregation helpers ------------------------------
def _first_non_null(series):
    for v in series:
        if pd.notna(v) and v != "":
            return v
    return np.nan

def aggregate_by_product(df: pd.DataFrame) -> pd.DataFrame:
    """Vue dé-dupliquée : 1 ligne par product_name, en conservant toutes les colonnes clés."""
    if df.empty or 'product_name' not in df.columns:
        return df

    tmp = df.copy()

    # ✅ Harmoniser les noms attendus par l'UI
    if 'product_category' in tmp.columns and 'Product Category' not in tmp.columns:
        tmp['Product Category'] = tmp['product_category']
    if 'credit_cumulable' in tmp.columns and 'Credit_cumulable' not in tmp.columns:
        tmp['Credit_cumulable'] = tmp['credit_cumulable']

    # Sécuriser types numériques courants
    numeric_like = [
        'total_stock', 'Predicted Order Quantity', 'purchase_need', 'MOQ MAAD', 'QAC',
        'Optimal Stock (Reorder Point)', 'Max Lead Time', 'Avg Daily Sales', 'Buffer Value',
        'Max Daily Sales (Pikine)', 'Max Coverage Day', 'Daily OOS Rate (7d)',
        'Daily OOS Rate (30d)', 'ADJUSTED_LEADTIME', 'MAX_CREDIT_BUFFER', 'AJUSTER_BUFFER',
        'credit_days'
    ]
    for c in numeric_like:
        if c in tmp.columns:
            tmp[c] = pd.to_numeric(tmp[c], errors='coerce')

    if 'Supplier' not in tmp.columns:
        tmp['Supplier'] = ""

    # Fournisseur principal = plus grand stock
    if {'product_name','Supplier','total_stock'}.issubset(tmp.columns):
        idx_max = tmp.groupby('product_name')['total_stock'].idxmax()
        main_sup = tmp.loc[idx_max, ['product_name','Supplier']].rename(columns={'Supplier':'Main Supplier'})
    else:
        main_sup = pd.DataFrame(columns=['product_name','Main Supplier'])

    # Priorités pour quelques colonnes catégorielles
    status_priority = {'Out of Stock': 3, 'Predicted Stockout Soon': 2, 'Order Soon': 1, 'Stock OK': 0}
    def pick_status(series):
        s = series.dropna().astype(str)
        if s.empty: return np.nan
        return s.iloc[s.map(lambda v: status_priority.get(v, 0)).argmax()]

    need_priority = {'ORDER NOW': 2, 'ORDER NOT URGENT': 1, 'NO NEED': 0}
    def pick_need(series):
        s = series.dropna().astype(str)
        if s.empty: return np.nan
        return s.iloc[s.map(lambda v: need_priority.get(v, 0)).argmax()]

    def first_non_null(series):
        for v in series:
            if pd.notna(v) and str(v).strip() != "":
                return v
        return np.nan

    def join_suppliers(s):
        return ", ".join(sorted({str(x).strip() for x in s if pd.notna(x) and str(x).strip() != ""}))

    def pick_credit_cumulable(series):
        s = series.dropna().astype(str).str.lower()
        if (s == 'oui').any(): return 'Oui'
        if (s == 'non').any(): return 'Non'
        return first_non_null(series)

    # 🔧 Plan d’agrégation
    agg_map = {}

    # Sommes (conservatives pour des besoins agrégés)
    for c in ['total_stock','Predicted Order Quantity','purchase_need','MOQ MAAD','QAC']:
        if c in tmp.columns: agg_map[c] = 'sum'

    # Max (couverture, LT, OOS, buffers…)
    for c in [
        'Optimal Stock (Reorder Point)', 'Max Lead Time', 'Avg Daily Sales', 'Buffer Value',
        'Max Daily Sales (Pikine)', 'Max Coverage Day', 'Daily OOS Rate (7d)', 'Daily OOS Rate (30d)',
        'ADJUSTED_LEADTIME', 'MAX_CREDIT_BUFFER', 'AJUSTER_BUFFER', 'credit_days'
    ]:
        if c in tmp.columns: agg_map[c] = 'max'

    # Bool -> OR
    if 'Predicted Stockout' in tmp.columns:
        tmp['Predicted Stockout'] = tmp['Predicted Stockout'].astype(bool)
        agg_map['Predicted Stockout'] = 'max'

    # Catégorielles importantes
    if 'Ajusted_total_need' in tmp.columns:
        agg_map['Ajusted_total_need'] = pick_need
    if 'Stock Status' in tmp.columns:
        agg_map['Stock Status'] = pick_status
    # Dans aggregate_by_product(), ligne ~1182
    if 'Product Category' in tmp.columns:
        # Fonction qui priorise les catégories issues de la règle crédit
        def pick_category_with_credit_priority(series):
            # Si toutes les valeurs sont identiques, prendre n'importe laquelle
            unique_vals = series.dropna().unique()
            if len(unique_vals) <= 1:
                return first_non_null(series)

            # Sinon, prioriser les valeurs qui ne sont PAS "CX" (défaut)
            non_default = series[series != "CX"].dropna()
            if not non_default.empty:
                return non_default.iloc[0]

            return first_non_null(series)

        agg_map['Product Category'] = pick_category_with_credit_priority
    if 'delisting_status' in tmp.columns:
        agg_map['delisting_status'] = first_non_null
    if 'Credit_cumulable' in tmp.columns:
        agg_map['Credit_cumulable'] = pick_credit_cumulable

    # Liste des fournisseurs
    agg_map['Supplier'] = join_suppliers

    grouped = tmp.groupby('product_name', dropna=False).agg(agg_map).reset_index()

    # Rejoindre le fournisseur principal
    if not main_sup.empty:
        grouped = grouped.merge(main_sup, on='product_name', how='left')
        # ✅ GARDER "Supplier" comme colonne principale
        if 'Supplier' in grouped.columns:
            grouped.rename(columns={'Supplier': 'Suppliers (all)'}, inplace=True)
        # ✅ AJOUTER : renommer Main Supplier -> Supplier pour compatibilité UI
        if 'Main Supplier' in grouped.columns:
            grouped['Supplier'] = grouped['Main Supplier']
            # Optionnel : garder aussi Suppliers (all) pour référence

    return grouped


# ------------------------------ Pages --------------------------------------------
def make_kpis(df: pd.DataFrame):
    # Helper pour format
    def fmt(n):
        if pd.isna(n): return "-"
        if isinstance(n, (int, float)):
            try: return f"{n:,.0f}".replace(",", " ")
            except Exception: return str(n)
        return str(n)

    # Nettoyage : exclure delisted
    mask_valid = df['delisting_status'].str.lower().ne('delisted') if 'delisting_status' in df.columns else [True] * len(df)
    df_valid = df[mask_valid].copy()

    total_skus = df_valid['product_name'].nunique() if 'product_name' in df_valid.columns else len(df_valid)

    # Rupture réelle (constaté dans le tableau)
    out_of_stock = int((df_valid['Stock Status'] == 'Out of Stock').sum()) if 'Stock Status' in df_valid.columns else 0

    # Fournisseurs
    suppliers = df_valid['Suppliers (all)'].nunique() if 'Suppliers (all)' in df_valid.columns else (
        df_valid['Supplier'].nunique() if 'Supplier' in df_valid.columns else 0
    )

    # Risque de rupture (ML)
    risk_count = 0
    if 'Predicted Stockout' in df_valid.columns and 'Credit Adequacy Score' in df_valid.columns:
        risk_count = int((
            (df_valid['Predicted Stockout'] == True) &
            (df_valid['Credit Adequacy Score'] < 0.5)
        ).sum())

    # Cartes KPI
    cards = dbc.Row([
        dbc.Col(html.Div(className="kpi", children=[html.Small("SKUs"), html.H3(fmt(total_skus))]), md=4),
        dbc.Col(html.Div(className="kpi", children=[html.Small("Produits en rupture"), html.H3(fmt(out_of_stock))]), md=4),
        dbc.Col(html.Div(className="kpi", children=[html.Small("Fournisseurs actifs"), html.H3(fmt(suppliers))]), md=4),
    ], className="gy-3")

    # Cloche d’alerte basée uniquement sur ML
    bell = html.Div(
        className="pill",
        children=[
            html.Span("🔔", style={"fontSize": "16px", "marginRight": "8px"}),
            html.B("À risque de rupture (ML) : "),
            html.Span(f"{risk_count}", className="badge badge-warn", style={"marginLeft": "6px"})
        ],
        style={"display": "inline-block", "marginTop": "10px"}
    )

    return cards, bell

def page_overview(master_df: pd.DataFrame = None):
    # Charger les données
    df = master_df if master_df is not None else get_df_cached()
    df = df.copy()

    # === UI HARDENING (garantir présence + valeurs non vides) ===
    def _ui_harden(df):
        # Supplier non vide
        if "Supplier" not in df.columns:
            df["Supplier"] = "unknown"
        else:
            s = df["Supplier"]
            df["Supplier"] = (
                s.where(~s.isna(), None)
                .apply(lambda x: None if (isinstance(x, str) and str(x).strip() == "") else x)
                .fillna("unknown")
                .astype(str).str.strip()
            )

        # Recalculated ADS non vide et numérique
        if "Recalculated Average Daily Sales" not in df.columns:
            df["Recalculated Average Daily Sales"] = 0.1

        df["Recalculated Average Daily Sales"] = (
            pd.to_numeric(df["Recalculated Average Daily Sales"], errors="coerce")
            .fillna(0.1)
            .clip(lower=0.1)
        )
        return df

    df = _ui_harden(df)

    # ✅ Harmoniser quelques alias pour l’UI
    # Ici, on NE recopie PAS bêtement product_category
    if "supplier_categorization" in df.columns and "Product Category" not in df.columns:
        # Respecter la règle métier → prendre supplier_categorization comme référence
        df["Product Category"] = df["supplier_categorization"]

    elif "product_category" in df.columns and "Product Category" not in df.columns:
        # Fallback si la catégorisation fournisseur n’existe pas encore
        df["Product Category"] = df["product_category"]

    if "credit_cumulable" in df.columns and "Credit_cumulable" not in df.columns:
        df["Credit_cumulable"] = df["credit_cumulable"]

    if "supplier" in df.columns and "Supplier" not in df.columns:
        df["Supplier"] = df["supplier"]

    # Colonnes prioritaires dans l’ordre
    cols_priority = [
        "product_name", "Supplier", "Suppliers (all)", "Recalculated Average Daily Sales",
        "Product Category",
        "total_stock", "Avg Daily Sales", "Max Daily Sales (Pikine)",
        "Max Coverage Day", "ADJUSTED_LEADTIME",
        "Optimal Stock (Reorder Point)", "Predicted Stockout",
        "Predicted Order Quantity", "purchase_need",
        "MOQ MAAD", "QAC", "Ajusted_total_need",
        "credit_days", "Credit_cumulable", "MAX_CREDIT_BUFFER", "AJUSTER_BUFFER",
        "delisting_status"
    ]
    # Forcer ces colonnes à être prioritaires et visibles
    for must in ["Supplier", "Recalculated Average Daily Sales"]:
        if must not in cols_priority:
            # les mettre très haut (après product_name)
            insert_at = 1 if "product_name" in cols_priority else 0
            cols_priority.insert(insert_at, must)

    available_priority = [c for c in cols_priority if c in df.columns]
    extra_cols = [c for c in df.columns if c not in cols_priority]
    available_cols = available_priority + extra_cols

    # 🔒 Double sécurité : garantir Supplier et Recalculated ADS
    for col in ["Supplier", "Recalculated Average Daily Sales"]:
        if col not in available_cols:
            available_cols.insert(1, col)

            # ➕ Nouvelle colonne Actions
    # ➕ Colonne Actions avec vrais boutons Dash
    #df["Actions"] = [
     #   html.Div([
      #      html.Button("✏️", id={"type": "edit-btn", "index": i}, n_clicks=0,
       #                 className="btn btn-sm btn-warning", style={"marginRight": "4px"}),
        #    html.Button("🗑️", id={"type": "delete-btn", "index": i}, n_clicks=0,
         #               className="btn btn-sm btn-danger", style={"marginRight": "4px"}),
          #  html.Button("➕", id={"type": "add-btn", "index": i}, n_clicks=0,
           #             className="btn btn-sm btn-success")
        #], style={"display": "flex", "gap": "4px"})
        #for i in range(len(df))
    #]

    # KPIs
    kpi_cards, risk_bell = make_kpis(df)

    # En-tête
    header_row = dbc.Row([
        dbc.Col(html.H2("Overview"), md=8),
        dbc.Col(html.Div(risk_bell, style={"textAlign": "right"}), md=4),
    ])

    cols_to_hide = [
        "Credit Adequacy Score", "Credit Adequacy Risk", "Stock Status",
        "Daily OOS Rate (30d)_y", "Avg Lead Time",
        "Coverage Day (30d)", "Coverage Day (7d)",
        "Average Daily Sales (30d)", "Average Daily Sales (7d)",
        "Daily OOS Rate (30d)_x", "Stockout Probability"
    ]
    available_cols = [c for c in available_cols if c not in cols_to_hide]

    # Tableau principal
    table = dash_table.DataTable(
        id="main-table",
        columns=[
            {"name": c, "id": c, "deletable": False, "hideable": True,
             "presentation": "component" if c == "Actions" else "input"}
            for c in available_cols
        ],
        data=df[available_cols].to_dict("records"),
        page_size=15,
        filter_action="native",
        sort_action="native",
        sort_mode="multi",
        column_selectable="single",
        editable=True,
        row_selectable="single",  # ✅ UNE SEULE FOIS
        selected_rows=[],
        style_table={"overflowX": "auto", "maxWidth": "100%"},
        style_header={"backgroundColor": "#0f1625", "border": "1px solid #1f2937",
                      "fontWeight": "700", "textAlign": "center"},
        style_cell={"backgroundColor": "#0b1220", "color": "#e5e7eb",
                    "border": "1px solid #1f2937", "fontSize": 11,
                    "textAlign": "left", "padding": "6px"},
        style_data_conditional=[
            {"if": {"filter_query": "{Ajusted_total_need} = 'ORDER NOW'"},
             "backgroundColor": "rgba(239,68,68,.2)", "color": "#fee2e2"},
            {"if": {"filter_query": "{Ajusted_total_need} = 'ORDER NOT URGENT'"},
             "backgroundColor": "rgba(245,158,11,.2)", "color": "#fef3c7"},
            {"if": {"filter_query": "{Ajusted_total_need} = 'NO NEED'"},
             "backgroundColor": "rgba(16,185,129,.15)", "color": "#d1fae5"},
            {"if": {"state": "active"},
             "backgroundColor": "#1e293b", "color": "#f9fafb"},
            {"if": {"state": "selected"},
             "backgroundColor": "#334155", "color": "#f9fafb"},
        ],
        style_data={"whiteSpace": "normal", "height": "auto"},
        export_format="csv",
        export_headers="display",
        persistence=True,
        persisted_props=["filter_query", "sort_by", "page_current",
                         "selected_rows", "selected_columns", "hidden_columns"],
    )

    # Boutons d’actions globales
    action_buttons = dbc.ButtonGroup([
        #dbc.Button("📥 Export CSV", id="btn-export", className="btn-outline-primary", size="sm"),
        dbc.Button("🔄 Actualiser", id="btn-refresh", className="btn-outline-secondary", size="sm"),
        dbc.Button("➕ Ajouter une ligne", id="btn-add-row", className="btn-primary", size="sm"),
    ])

    # ✅ Modal pour édition
    # Dans page_overview(), remplacez le bloc edit_modal par :
    edit_modal = dbc.Modal(
        [
            dbc.ModalHeader(dbc.ModalTitle("Ajouter/Éditer produit")),
            dbc.ModalBody([
                dbc.Label("Nom du produit"),
                dbc.Input(id="edit-product", type="text", placeholder="Ex: Lait concentré 400g"),
                html.Br(),
                dbc.Label("Fournisseur"),
                dbc.Input(id="edit-supplier", type="text", placeholder="Ex: Carrefour"),
                html.Br(),
                dbc.Label("Catégorie"),
                dcc.Dropdown(
                    id="edit-category",
                    options=[
                        {"label": "AX", "value": "AX"},
                        {"label": "AY", "value": "AY"},
                        {"label": "AZ", "value": "AZ"},
                        {"label": "BX", "value": "BX"},
                        {"label": "BY", "value": "BY"},
                        {"label": "BZ", "value": "BZ"},
                        {"label": "CX", "value": "CX"},
                        {"label": "CY", "value": "CY"},
                        {"label": "CZ", "value": "CZ"},
                    ],
                    placeholder="Sélectionner une catégorie"
                ),
                html.Br(),
                dbc.Label("Stock initial"),
                dbc.Input(id="edit-stock", type="number", placeholder="0", min=0),
            ]),
            dbc.ModalFooter([
                dbc.Button("Annuler", id="edit-cancel", className="btn-secondary", n_clicks=0),
                dbc.Button("Enregistrer", id="edit-save", className="btn-primary", n_clicks=0),
            ]),
        ],
        id="edit-modal",
        is_open=False,
    )

@app.callback(
    Output("edit-modal", "is_open", allow_duplicate=True),
    Input("edit-cancel", "n_clicks"),
    prevent_initial_call=True
)
def close_modal(n_clicks):
    if n_clicks:
        return False
    return no_update

from dash import callback_context

# ✏️ Éditer une ligne
@app.callback(
    Output("main-table", "data", allow_duplicate=True),
    Input({"type": "edit-btn", "index": ALL}, "n_clicks"),
    State("main-table", "data"),
    prevent_initial_call=True
)
def edit_row(edit_clicks, rows):
    ctx = callback_context
    if not ctx.triggered:
        return rows
    triggered_id = ctx.triggered[0]["prop_id"].split(".")[0]
    triggered = json.loads(triggered_id)
    index = triggered["index"]

    if 0 <= index < len(rows):
        rows[index]["product_name"] = str(rows[index].get("product_name", "")) + " (✏️ édité)"
    return rows


# 🗑️ Supprimer une ligne
@app.callback(
    Output("main-table", "data", allow_duplicate=True),
    Input({"type": "delete-btn", "index": ALL}, "n_clicks"),
    State("main-table", "data"),
    prevent_initial_call=True
)
def delete_row(delete_clicks, rows):
    ctx = callback_context
    if not ctx.triggered:
        return rows
    triggered_id = ctx.triggered[0]["prop_id"].split(".")[0]
    triggered = json.loads(triggered_id)
    index = triggered["index"]

    if 0 <= index < len(rows):
        rows.pop(index)
    return rows


# ➕ Ajouter une ligne
@app.callback(
    Output("main-table", "data", allow_duplicate=True),
    Input({"type": "add-btn", "index": ALL}, "n_clicks"),
    State("main-table", "data"),
    State("main-table", "columns"),
    prevent_initial_call=True
)
def add_row(add_clicks, rows, columns):
    ctx = callback_context
    if not ctx.triggered:
        return rows
    triggered_id = ctx.triggered[0]["prop_id"].split(".")[0]
    triggered = json.loads(triggered_id)
    index = triggered["index"]

    # Crée une ligne vide
    new_row = {c["id"]: "" for c in columns}
    rows.insert(index + 1, new_row)
    return rows


def page_analytics(master_df: pd.DataFrame = None):
    df = master_df if master_df is not None else get_df_cached()

    # Colonnes nécessaires
    cols_to_keep = [
        "product_name", "Supplier", "Product Category",
        "total_stock", "Recalculated Average Daily Sales",
        "Max Coverage Day", "ADJUSTED_LEADTIME",
        "Ajusted_total_need", "purchase_need",
        "QAC", "Optimal Stock (Reorder Point)"
    ]
    cols_to_keep = [c for c in cols_to_keep if c in df.columns]
    analytics_df = df[cols_to_keep].dropna()

    # KPI cards
    kpi_cards, _ = make_kpis(df)

    # Scatter Lead Time vs Coverage
    fig_scatter = px.scatter(
        analytics_df,
        x="ADJUSTED_LEADTIME",
        y="Max Coverage Day",
        color="Ajusted_total_need",
        hover_data=["product_name", "Supplier", "purchase_need", "QAC"],
        labels={
            "ADJUSTED_LEADTIME": "Adjusted Lead Time (jours)",
            "Max Coverage Day": "Max Coverage (jours)"
        },
        title="Relation Lead Time ajusté vs Couverture stock"
    )

    # Histogramme distribution du stock
    fig_hist_stock = px.histogram(
        analytics_df,
        x="total_stock",
        nbins=40,
        title="Distribution du stock total",
        labels={"total_stock": "Stock total"}
    )

    # Boxplot par catégorie
    fig_box_category = px.box(
        analytics_df,
        x="Product Category",
        y="Recalculated Average Daily Sales",
        title="Distribution des ventes moyennes par catégorie",
        labels={"Recalculated Average Daily Sales": "Ventes moyennes recalculées"}
    )

    # Bar chart besoin d'achat par fournisseur
    purchase_by_supplier = (
        analytics_df.groupby("Supplier")["purchase_need"].sum()
        .reset_index().sort_values("purchase_need", ascending=False).head(10)
    )
    fig_purchase_supplier = px.bar(
        purchase_by_supplier,
        x="Supplier",
        y="purchase_need",
        title="Top 10 besoins d'achat par fournisseur",
        labels={"purchase_need": "Besoin d’achat"}
    )

    # Scatter QAC vs Optimal Stock
    fig_qac_vs_optimal = px.scatter(
        analytics_df,
        x="QAC",
        y="Optimal Stock (Reorder Point)",
        color="Product Category",
        title="Relation QAC vs Stock optimal",
        labels={"QAC": "Quantité ajustée commandée"}
    )

    return html.Div(className="content", children=[
        dbc.Row([dbc.Col(html.H2("Analyses avancées"), md=12)]),
        html.Div(kpi_cards),
        html.Br(),

        # Ligne 1
        dbc.Row([
            dbc.Col(dcc.Graph(figure=fig_scatter, id="analytics-scatter"), md=6),
            dbc.Col(dcc.Graph(figure=fig_hist_stock), md=6)
        ]),
        html.Br(),

        # Ligne 2
        dbc.Row([
            dbc.Col(dcc.Graph(figure=fig_box_category), md=6),
            dbc.Col(dcc.Graph(figure=fig_purchase_supplier), md=6),
        ]),
        html.Br(),

        # Ligne 3
        dbc.Row([
            dbc.Col(dcc.Graph(figure=fig_qac_vs_optimal), md=12)
        ])
    ])


# Callback : mettre à jour le scatter analytics avec filtres
#@app.callback(
 #   Output("analytics-scatter", "figure"),
  #  Input("analytics-filter-supplier", "value"),
   # Input("analytics-filter-category", "value"),
    #prevent_initial_call=False
#)
def update_analytics_scatter(supplier_value, category_value):
    df = get_df_cached()

    # Colonnes nécessaires
    cols = ["product_name", "Supplier", "Product Category",
            "ADJUSTED_LEADTIME", "Max Coverage Day",
            "Ajusted_total_need", "purchase_need", "QAC"]
    df = df[[c for c in cols if c in df.columns]].dropna()

    # Filtres
    if supplier_value and supplier_value != "Tous":
        df = df[df["Supplier"] == supplier_value]
    if category_value and category_value != "Toutes":
        df = df[df["Product Category"] == category_value]

    fig = px.scatter(
        df,
        x="ADJUSTED_LEADTIME",
        y="Max Coverage Day",
        color="Ajusted_total_need",
        hover_data=["product_name", "Supplier", "purchase_need", "QAC"],
        labels={
            "ADJUSTED_LEADTIME": "Adjusted Lead Time (jours)",
            "Max Coverage Day": "Max Coverage (jours)"
        },
        title="Relation Lead Time ajusté vs Couverture stock"
    )
    return fig

def page_predictive(master_df: pd.DataFrame = None):
    df = master_df if master_df is not None else get_df_cached()

    cols_to_keep = [
        "product_name", "Supplier", "Product Category",
        "total_stock", "Recalculated Average Daily Sales",
        "Predicted Stockout", "Predicted Order Quantity",
        "purchase_need", "QAC", "Ajusted_total_need",
        "ADJUSTED_LEADTIME", "Max Coverage Day"
    ]
    cols_to_keep = [c for c in cols_to_keep if c in df.columns]
    pred_df = df[cols_to_keep].copy()

    # KPI : % de stockout prédits
    if "Predicted Stockout" in pred_df.columns:
        stockout_rate = pred_df["Predicted Stockout"].mean() * 100
    else:
        stockout_rate = 0

    fig_bar = px.bar(
        pred_df.sort_values("Predicted Order Quantity", ascending=False).head(20),
        x="product_name",
        y="Predicted Order Quantity",
        color="Ajusted_total_need",
        hover_data=["Supplier", "QAC", "purchase_need"],
        labels={"Predicted Order Quantity": "Qté de commande prédite"},
        title="Top 20 produits par besoin de commande"
    )

    return html.Div(className="content", children=[
        dbc.Row([
            dbc.Col(html.H2("Prédictions"), md=8),
            dbc.Col(html.H5(f"Taux de stockout prédit : {stockout_rate:.1f}%", style={"textAlign": "right"}), md=4),
        ]),
        html.Br(),
        html.Div(className="soft-card", children=[
            dcc.Graph(figure=fig_bar, id="predictive-bar")
        ]),
        html.Br(),
        dash_table.DataTable(
            id="predictive-table",
            columns=[{"name": c, "id": c} for c in pred_df.columns],
            data=pred_df.to_dict("records"),
            page_size=15,
            filter_action="native",
            sort_action="native", sort_mode="multi",
            style_table={"overflowX": "auto"},
            style_header={"backgroundColor": "#0f1625", "border": "1px solid #1f2937", "fontWeight": "700"},
            style_cell={"backgroundColor": "#0b1220", "color": "#e5e7eb", "border": "1px solid #1f2937", "fontSize": 12},
        )
    ])


# Callback : mettre à jour le bar chart prédictif
@app.callback(
    Output("predictive-bar", "figure"),
    Input("predictive-filter-supplier", "value"),
    Input("predictive-filter-category", "value"),
    prevent_initial_call=False
)
def update_predictive_bar(supplier_value, category_value):
    df = get_df_cached()

    cols = ["product_name", "Supplier", "Product Category",
            "Predicted Stockout", "Predicted Order Quantity",
            "purchase_need", "QAC", "Ajusted_total_need"]
    df = df[[c for c in cols if c in df.columns]].copy()

    # Filtres
    if supplier_value and supplier_value != "Tous":
        df = df[df["Supplier"] == supplier_value]
    if category_value and category_value != "Toutes":
        df = df[df["Product Category"] == category_value]

    fig = px.bar(
        df.sort_values("Predicted Order Quantity", ascending=False).head(20),
        x="product_name",
        y="Predicted Order Quantity",
        color="Ajusted_total_need",
        hover_data=["Supplier", "QAC", "purchase_need"],
        labels={"Predicted Order Quantity": "Qté de commande prédite"},
        title="Top 20 produits par besoin de commande"
    )
    return fig


# Callback : mettre à jour le tableau prédictif
@app.callback(
    Output("predictive-table", "data"),
    Input("predictive-filter-supplier", "value"),
    Input("predictive-filter-category", "value"),
    prevent_initial_call=False
)
def update_predictive_table(supplier_value, category_value):
    df = get_df_cached()

    cols = ["product_name", "Supplier", "Product Category",
            "total_stock", "Recalculated Average Daily Sales",
            "Predicted Stockout", "Predicted Order Quantity",
            "purchase_need", "QAC", "Ajusted_total_need",
            "ADJUSTED_LEADTIME", "Max Coverage Day"]
    df = df[[c for c in cols if c in df.columns]].copy()

    # Filtres
    if supplier_value and supplier_value != "Tous":
        df = df[df["Supplier"] == supplier_value]
    if category_value and category_value != "Toutes":
        df = df[df["Product Category"] == category_value]

    return df.to_dict("records")

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
    if not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()

    d = pd.DataFrame()
    d["product_name_display"] = df.get("product_name", "")
    d["supplier_name"] = df.get("Supplier", "")
    d["abc_class"] = df.get("Product Category", "")
    d["xyz_class"] = ""
    d["current_stock"] = pd.to_numeric(df.get("total_stock", 0), errors="coerce").fillna(0).clip(lower=0)
    d["avg_daily_sales"] = pd.to_numeric(df.get("Recalculated Average Daily Sales", 0), errors="coerce").fillna(0).clip(
        lower=0)  # ✅ Changé de "Max Avg Daily Sales"

    # ✅ CORRECTION : Gestion robuste de coverage_days
    with np.errstate(divide='ignore', invalid='ignore'):
        cov = np.where(d["avg_daily_sales"] > 0, d["current_stock"] / d["avg_daily_sales"], 0)

    if "Max Coverage Day" in df.columns:
        # Si la colonne existe, l'utiliser
        d["coverage_days"] = pd.to_numeric(df["Max Coverage Day"], errors="coerce").fillna(0)
    else:
        # Sinon, utiliser le calcul
        d["coverage_days"] = cov

    # Convertir explicitement en Series avant clip
    d["coverage_days"] = pd.Series(d["coverage_days"], dtype=float).clip(lower=0, upper=365)

    # ✅ CORRECTION : Lead time avec fallback
    if "ADJUSTED_LEADTIME" in df.columns:
        d["leadtime_days"] = pd.to_numeric(df["ADJUSTED_LEADTIME"], errors="coerce").fillna(7)
    elif "Max Lead Time" in df.columns:
        d["leadtime_days"] = pd.to_numeric(df["Max Lead Time"], errors="coerce").fillna(7)
    else:
        d["leadtime_days"] = 7

    d["leadtime_days"] = pd.Series(d["leadtime_days"], dtype=float).clip(lower=0, upper=180)

    # ✅ CORRECTION : Crédit
    if "credit_days" in df.columns:
        d["credit_days"] = pd.to_numeric(df["credit_days"], errors="coerce").fillna(14)
    else:
        d["credit_days"] = 14

    # Rupture ML
    if "Predicted Stockout" in df.columns:
        d["rupture_ml"] = np.where(df["Predicted Stockout"], "OUI", "NON")
    else:
        st = df.get("Stock Status", "").astype(str)
        d["rupture_ml"] = np.where(st.isin(["Out of Stock", "Predicted Stockout Soon"]), "OUI", "NON")

    d["delisting_product"] = "NON"

    # Exclure produits non pertinents
    bad = d["product_name_display"].astype(str).str.contains(
        r'\b(CFA|CASH|CFA\s*-\s*CASH|ESPECES|CAISSE)\b',
        case=False,
        na=False
    )
    d = d[~bad].copy()  # ✅ Ajout de .copy() pour éviter SettingWithCopyWarning

    d["risk"] = (d["rupture_ml"].astype(str).str.upper() == "OUI").astype(int)

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
        fields = ['product_name_display', 'supplier_name', 'abc_class', 'xyz_class', 'current_stock', 'avg_daily_sales',
                  'coverage_days', 'leadtime_days', 'credit_days', 'rupture_ml', 'delisting_product', 'risk']
        view = sample[[c for c in fields if c in sample.columns]].copy() if not sample.empty else pd.DataFrame()

        cov_med = None;
        rows_digest = []
        try:
            if not view.empty:
                cov_med = float(view['coverage_days'].median()) if 'coverage_days' in view.columns else None
                at_risk = view[
                    view['rupture_ml'].str.upper() == 'OUI'] if 'rupture_ml' in view.columns else pd.DataFrame()
                top_list = at_risk.sort_values('coverage_days', ascending=True).head(
                    6) if not at_risk.empty else view.sort_values('coverage_days', ascending=True).head(6)

                for _, r in top_list.iterrows():
                    rows_digest.append({
                        "name": str(r.get('product_name_display', '')),
                        "sup": str(r.get('supplier_name', '')),
                        "cov": float(r.get('coverage_days', 0)),
                        "lt": int(r.get('leadtime_days', 0)),
                        "abc": str(r.get('abc_class', '')),
                        "xyz": str(r.get('xyz_class', '')),
                        "cred": int(r.get('credit_days', 0)),
                    })
        except Exception as e:
            print(f"[Chatbot] Erreur agrégation: {e}")

        table_json = []
        try:
            table_json = view.head(80).to_dict(orient="records")
        except Exception:
            pass

        if openai_client is None:
            return _chatbot_fallback(user_text, view)

        # ✅ DÉTECTION SI ANALYSE DEMANDÉE
        needs_analysis = any(keyword in user_text.lower() for keyword in [
            'analyse', 'recommande', 'conseil', 'rupture', 'commande', 'stock',
            'produit', 'quels', 'combien', 'urgent', 'priorité', 'fournisseur',
            'risque', 'order', 'achat', 'besoin', 'coverage', 'lead time'
        ])

        if lang == "fr":
            system_msg = (
                "Tu es maad_assistante, expert Supply Chain avec 15 ans d'expérience. Tu es conversationnel et réponds naturellement aux questions.\n\n"
                "RÈGLES DE CONVERSATION :\n"
                "- Si l'utilisateur te salue ou discute, réponds de manière amicale et naturelle\n"
                "- Si l'utilisateur pose une question générale sur la supply chain, explique clairement sans forcer une analyse de données\n"
                "- SEULEMENT si l'utilisateur demande explicitement une analyse, recommandation, ou conseil sur ses stocks, utilise les données fournies\n"
                "- Adapte ton niveau de détail à la question : simple question = réponse courte, analyse demandée = détails avec chiffres\n\n"
                "QUAND TU ANALYSES LES DONNÉES :\n"
                "- Cite des SKU concrets du tableau\n"
                "- Propose des quantités chiffrées\n"
                "- Mentionne les fournisseurs prioritaires\n"
                "- Suggère des leviers (crédit, promo déstockage, catégories ABC/XYZ)\n"
                "- Sois orienté action avec des recommandations précises"
            )

            # ✅ CONTEXTE RÉDUIT : Envoyer données SEULEMENT si nécessaire
            if needs_analysis and table_json:
                ctx = f"Données disponibles pour analyse :\n"
                ctx += f"- Extrait produits (20 premiers) : {json.dumps(table_json[:20], ensure_ascii=False)}\n"
                if cov_med is not None:
                    ctx += f"- KPI global : couverture médiane ≈ {cov_med:.1f} jours\n"
                if rows_digest:
                    ctx += "- Produits prioritaires : " + "; ".join(
                        [f"{d['name']} (fournisseur {d['sup']}, couverture {d['cov']:.1f}j, lead time {d['lt']}j)"
                         for d in rows_digest[:3]]
                    ) + "\n"
            else:
                ctx = "Conversation générale. Données disponibles si besoin d'analyse détaillée.\n"

            user_q = f"Question : {user_text.strip()}"

        else:  # English
            system_msg = (
                "You are maad_assistante, a Supply Chain expert with 15 years of experience. You're conversational and respond naturally to questions.\n\n"
                "CONVERSATION RULES:\n"
                "- If the user greets you or chats, respond in a friendly and natural way\n"
                "- If the user asks a general supply chain question, explain clearly without forcing data analysis\n"
                "- ONLY if the user explicitly requests analysis, recommendations, or advice on their inventory, use the provided data\n"
                "- Adapt your detail level to the question: simple question = short answer, analysis requested = details with numbers\n\n"
                "WHEN ANALYZING DATA:\n"
                "- Cite concrete SKUs from the table\n"
                "- Propose quantified amounts\n"
                "- Mention priority suppliers\n"
                "- Suggest levers (credit, promo clearance, ABC/XYZ categories)\n"
                "- Be action-oriented with precise recommendations"
            )

            if needs_analysis and table_json:
                ctx = f"Available data for analysis:\n"
                ctx += f"- Product excerpt (20 first): {json.dumps(table_json[:20], ensure_ascii=False)}\n"
                if cov_med is not None:
                    ctx += f"- Global KPI: median coverage ≈ {cov_med:.1f} days\n"
                if rows_digest:
                    ctx += "- Priority products: " + "; ".join(
                        [f"{d['name']} (supplier {d['sup']}, coverage {d['cov']:.1f}d, lead time {d['lt']}d)"
                         for d in rows_digest[:3]]
                    ) + "\n"
            else:
                ctx = "General conversation. Data available if detailed analysis needed.\n"

            user_q = f"Question: {user_text.strip()}"

        # Construction de l'historique complet
        hist = [{"role": "assistant" if m.get("role") == "assistant" else "user", "content": m.get("text", "")}
                for m in (history_messages or [])]

        messages = [
                       {"role": "system", "content": system_msg},
                       {"role": "system", "content": ctx}  # Contexte séparé pour clarté
                   ] + hist + [
                       {"role": "user", "content": user_q}
                   ]

        try:
            resp = openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=messages,
                temperature=0.7,  # ✅ Augmenté de 0.4 à 0.7 pour conversation plus naturelle
                max_tokens=800  # ✅ Augmenté de 600 à 800 pour réponses plus détaillées
            )
            out = (resp.choices[0].message.content or "").strip()
            return out if out else _chatbot_fallback(user_text, view)

        except Exception as api_err:
            print(f"[Chatbot] Erreur OpenAI: {type(api_err).__name__}: {api_err}")
            return _chatbot_fallback(user_text, view, prefix=f"⚠️ API indisponible. ")

    except Exception as e:
        import traceback
        print(f"[Chatbot] Erreur _chatbot_reply: {traceback.format_exc()}")
        return f"⚠️ Erreur technique : {type(e).__name__}"


def _chatbot_fallback(user_text: str, df: pd.DataFrame, prefix: str = "") -> str:
    """
    df ici est déjà le DataFrame nettoyé (view), pas besoin de re-nettoyer
    """
    lang = _detect_lang(user_text)

    # ✅ NE PAS appeler _clean_df_for_advice ici, df est déjà nettoyé
    if df.empty:
        return prefix + (
            "Données insuffisantes. Recharge les sources." if lang == "fr" else "Insufficient data. Please reload sources.")

    try:
        # Utiliser directement df (qui est view)
        top_risk = df.sort_values(['risk', 'coverage_days'], ascending=[False, True]).head(
            5) if 'risk' in df.columns else df.head(5)
        cnt_risk = int(df['risk'].sum()) if 'risk' in df.columns else 0
        cov_med = float(df['coverage_days'].median()) if 'coverage_days' in df.columns else 0

        lines = [f"{prefix}" + (f"Risque: {cnt_risk} prod. • Couverture médiane ~ {cov_med:.1f} j." if lang == "fr"
                                else f"Risk: {cnt_risk} SKUs • Median coverage ~ {cov_med:.1f}d.")]

        for _, r in top_risk.iterrows():
            if lang == "fr":
                lines.append(
                    f"- {r.get('product_name_display', '?')} · cov {float(r.get('coverage_days', 0)):.1f}j · LT {int(r.get('leadtime_days', 0))}j · crédit {int(r.get('credit_days', 0))}j")
            else:
                lines.append(
                    f"- {r.get('product_name_display', '?')} · cov {float(r.get('coverage_days', 0)):.1f}d · LT {int(r.get('leadtime_days', 0))}d · credit {int(r.get('credit_days', 0))}d")

        lines += [("Cibles: A/AX en Y/Z <7j; crédit>14j si LT>20j; promos classe C." if lang == "fr"
                   else "Focus A/A+ in Y/Z <7d; credit>14d if LT>20d; promo bundles for class C.")]
        return "\n".join(lines)

    except Exception as e:
        print(f"[Chatbot] Erreur fallback: {e}")
        return prefix + (
            "Recommandation impossible avec les données disponibles." if lang == "fr" else "Unable to compute recommendation with available data.")
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
    dcc.Dropdown(id="filter-supplier"),
    dcc.Dropdown(id="filter-category"),
    dcc.Dropdown(id="filter-need"),  # ✅ IMPORTANT
    dcc.Input(id="search-input"),
    dbc.Checklist(id="toggle-options"),
    dcc.Store(id="uploaded-csv"),
    dcc.Store(id="chat-store"),
    dcc.Store(id="chat-open"),
    make_sidebar(),
    page_overview(initial_df),
    page_analytics(),
    page_predictive(),
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
        return page_predictive()
    elif path == "/about":
        return page_about()
    return page_overview(base)

# ------------------------------ Filtering logic ----------------------------------
def filter_dataframe(df: pd.DataFrame, query: str, suppliers: list, statuses: list, cats: list, options: list):
    out = df.copy()

    # Filtre recherche par product_name
    if query:
        q = str(query).strip().lower()
        if 'product_name' in out.columns:
            out = out[out['product_name'].astype(str).str.lower().str.contains(q, na=False, regex=False)]

    # Filtre Supplier
    if suppliers and len(suppliers) > 0:
        if 'Supplier' in out.columns:
            mask = out['Supplier'].astype(str).str.lower().apply(
                lambda x: any(sup.lower() in x for sup in suppliers)
            )
            out = out[mask]

    # Filtre Product Category
    if cats and len(cats) > 0:
        if 'Product Category' in out.columns:
            out = out[out['Product Category'].astype(str).str.lower().isin([c.lower() for c in cats])]

    # Agrégation par produit
    options = options or []
    if 'by_product' in options:
        out = aggregate_by_product(out)

    return out

def validate_core_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Garantit que les colonnes critiques existent et sont valides"""
    df = df.copy()

    # 1. Supplier
    if "Supplier" not in df.columns:
        if "Suppliers (all)" in df.columns:
            # Extraire le premier fournisseur de la liste
            df["Supplier"] = df["Suppliers (all)"].str.split(",").str[0].str.strip()
        else:
            df["Supplier"] = "unknown"

    # Nettoyer les valeurs vides
    df["Supplier"] = df["Supplier"].fillna("unknown").replace("", "unknown")

    # 2. Recalculated Average Daily Sales
    if "Recalculated Average Daily Sales" not in df.columns:
        df["Recalculated Average Daily Sales"] = 0.1

    df["Recalculated Average Daily Sales"] = pd.to_numeric(
        df["Recalculated Average Daily Sales"],
        errors='coerce'
    ).fillna(0.1).clip(lower=0.1)

    return df

# ------------------------------ Callbacks: filtering / banner --------------------
# Callback 1 : Initialisation (sans allow_duplicate)
@app.callback(
    [Output("filtered-data", "data"),
     Output("main-table", "data"),
     Output("risk-banner", "children")],
    Input("master-data", "data"),
    prevent_initial_call=False  # ✅ S'exécute au chargement
)
def initialize_table(master_json):
    """Initialise le tableau au chargement sans filtres"""
    df = pd.DataFrame(json.loads(master_json)) if master_json else get_df_cached()
    df = validate_core_columns(df)
    banner = " "
    return df.to_json(orient="records"), df.to_dict("records"), banner


# Callback 2 : Filtrage (avec allow_duplicate)
@app.callback(
    [Output("filtered-data", "data", allow_duplicate=True),
     Output("main-table", "data", allow_duplicate=True),
     Output("risk-banner", "children", allow_duplicate=True)],
    [Input("search-input", "value"),
     Input("filter-supplier", "value"),
     Input("filter-category", "value"),
     Input("filter-need", "value"),
     Input("toggle-options", "value")],
    State("master-data", "data"),
    prevent_initial_call=True  # ✅ S'exécute sur interaction
)
def apply_filters(search, sup, cat, need, options, master_json):
    base = pd.DataFrame(json.loads(master_json)) if master_json else get_df_cached()
    base = validate_core_columns(base)

    sup = sup or []
    cat = cat or []
    need = need or []
    options = options or []

    fdf = filter_dataframe(base, search, sup, [], cat, options)

    if need and len(need) > 0 and 'Ajusted_total_need' in fdf.columns:
        fdf = fdf[fdf['Ajusted_total_need'].isin(need)]

    print(f"[apply_filters] Résultat: {len(fdf)} lignes")

    banner = " "
    fdf_actions = add_action_cols(fdf)
    return fdf_actions.to_json(orient="records"), fdf_actions.to_dict("records"), banner


@app.callback(
    [Output("master-data", "data", allow_duplicate=True),
     Output("filtered-data", "data", allow_duplicate=True),
     Output("main-table", "data", allow_duplicate=True),
     Output("risk-banner", "children", allow_duplicate=True)],
    Input("btn-refresh", "n_clicks"),
    prevent_initial_call=True
)
def refresh_data(n_clicks):
    """Recharge les données depuis les sources Google Sheets"""
    if not n_clicks:
        return no_update, no_update, no_update, no_update

    print("[Refresh] Rechargement des données...")

    # Forcer le rechargement en invalidant le cache
    cache.clear()

    # Recharger les données
    df = get_df_cached()
    df = validate_core_columns(df)

    print(f"[Refresh] {len(df)} produits rechargés")

    banner = " "
    return df.to_json(orient="records"), df.to_json(orient="records"), df.to_dict("records"), banner
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

# ------------------------------ Export Purchase Order PDF ------------------------
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

    # ============= BLOC 1 : Import ReportLab avec gestion d'erreur =============
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.platypus import (SimpleDocTemplate, Paragraph, Table,
                                        TableStyle, Spacer, Image)
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib.units import mm
    except ImportError as e:
        print(f"❌ ReportLab import error: {e}")
        print("→ Run: pip install reportlab")
        return no_update

    # ============= BLOC 2 : Logique de sélection améliorée =============
    row_idx = None

    # Priorité 1 : Ligne explicitement sélectionnée
    if selected_rows and len(selected_rows) > 0:
        row_idx = selected_rows[0]

    # Priorité 2 : Cellule active
    elif active_cell and isinstance(active_cell, dict):
        row_idx = active_cell.get("row")

    # Validation
    if row_idx is None or row_idx < 0 or row_idx >= len(table_data):
        print(f"❌ PDF Error: Invalid row_idx={row_idx}, table length={len(table_data)}")
        return no_update

    r = table_data[row_idx]
    print(f"✅ Generating PDF for row {row_idx}: {r.get('product_name')}")

    # ============= BLOC 3 : Extraction des données (VOTRE CODE ORIGINAL) =============
    prod_name = str(r.get("product_name", "")).strip()
    supplier = str(r.get("Supplier", "")).strip()
    qty = r.get("Predicted Order Quantity", 0)
    unit_ht = r.get("Unit Price HT", 0.0)
    remise = r.get("Discount", 0.0)

    if not prod_name or not supplier:
        print(f"❌ Missing required fields: product_name='{prod_name}', supplier='{supplier}'")
        return no_update

    try:
        qty = int(pd.to_numeric(qty, errors="coerce") or 0)
    except Exception:
        qty = 0
    try:
        unit_ht = float(pd.to_numeric(unit_ht, errors="coerce") or 0.0)
    except Exception:
        unit_ht = 0.0
    try:
        remise = float(pd.to_numeric(remise, errors="coerce") or 0.0)
    except Exception:
        remise = 0.0

    if qty <= 0:
        qty = 1

    taux_tva = DEFAULT_TVA_RATE
    total_ht = (qty * unit_ht) * (1 - remise / 100.0)
    total_ttc = total_ht * (1 + taux_tva)

    def ref_from_name(name: str) -> str:
        if not name: return ""
        parts = [p for p in str(name).split() if p]
        if not parts: return str(name)[:6].upper()
        left = (parts[0][:3] if len(parts[0]) >= 3 else parts[0]).upper()
        right = (parts[1][:3] if len(parts) > 1 and len(parts[1]) >= 3 else (
            parts[0][3:6] if len(parts[0]) > 3 else "")).upper()
        return "-".join([left, right]) if right else left

    ref_code = ref_from_name(prod_name)

    # ============= BLOC 4 : Génération PDF avec gestion d'erreur robuste =============
    try:
        buf = io.BytesIO()
        po_number = get_next_po_number()
        fname = f"{po_number}_{ref_code}.pdf"

        doc = SimpleDocTemplate(buf, pagesize=A4, title="Bon de commande")
        styles = getSampleStyleSheet()
        story = []

        # Header avec logo
        header_row = []
        logo_src = get_logo_for_reportlab()
        if logo_src:
            try:
                logo_img = Image(logo_src)
                logo_img.drawHeight = 18 * mm
                logo_img.drawWidth = 18 * mm
                header_row.append(logo_img)
            except Exception as logo_err:
                print(f"⚠️ Logo image creation failed: {logo_err}, skipping")
                header_row.append(Paragraph("", styles["Normal"]))
        else:
            header_row.append(Paragraph("", styles["Normal"]))

        # Company info
        company_lines = [f"<b>{COMPANY_NAME}</b>"]
        if COMPANY_CAPITAL: company_lines.append(f"Montant du capital social : {COMPANY_CAPITAL}")
        if COMPANY_RCS:     company_lines.append(f"N° et lieu RCS : {COMPANY_RCS}")
        if COMPANY_ADDRESS: company_lines.append(f"Adresse du siège : {COMPANY_ADDRESS}")
        if COMPANY_PHONE:   company_lines.append(f"Téléphone : {COMPANY_PHONE}")
        if COMPANY_EMAIL:   company_lines.append(f"Email : {COMPANY_EMAIL}")
        header_row.append(Paragraph("<br/>".join(company_lines), styles["Normal"]))

        header_tbl = Table([header_row], colWidths=[25 * mm, 150 * mm])
        header_tbl.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        story.append(header_tbl)
        story.append(Spacer(1, 8))

        # Title et metadata
        story.append(Paragraph("<b>BON DE COMMANDE</b>", styles["Title"]))
        story.append(Spacer(1, 6))

        meta_left = [
            f"Bon de commande N° : <b>{po_number}</b>",
            f"Date : {datetime.now().strftime('%d/%m/%Y')}",
        ]
        meta_right = ["Fournisseur :", f"<b>{supplier}</b>"] if supplier else []
        meta_tbl = Table([[Paragraph("<br/>".join(meta_left), styles["Normal"]),
                           Paragraph("<br/>".join(meta_right), styles["Normal"])]],
                         colWidths=[100 * mm, 75 * mm])
        meta_tbl.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
        story.append(meta_tbl)
        story.append(Spacer(1, 10))

        # Tableau produit
        headers = ["REF", "DESCRIPTION", "QUANTITÉ", "PU HT", "REMISE", "TOTAL HT", "TAUX TVA", "TOTAL TTC"]
        data_tbl = [
            headers,
            [ref_code, prod_name, qty, f"{unit_ht:.2f}", f"{remise:.1f}%",
             f"{total_ht:.2f}", f"{int(taux_tva * 100)}%", f"{total_ttc:.2f}"],
            ["", "TOTAL", qty, "", "", f"{total_ht:.2f}", "", f"{total_ttc:.2f}"]
        ]

        tbl = Table(data_tbl, hAlign="LEFT",
                    colWidths=[25 * mm, 65 * mm, 20 * mm, 20 * mm, 20 * mm, 25 * mm, 20 * mm, 25 * mm])
        tbl.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("ALIGN", (2, 1), (2, -1), "RIGHT"),
            ("ALIGN", (3, 1), (-1, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
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

        print(f"✅ PDF generated successfully: {fname}")
        return dcc.send_bytes(lambda b: b.write(buf.getvalue()), filename=fname)

    except Exception as e:
        import traceback
        print(f"❌ PDF generation failed:")
        print(f"   Error: {type(e).__name__}: {e}")
        print(f"   Row data: {r}")
        print(f"   Traceback:\n{traceback.format_exc()}")
        return no_update

# Activer le bouton PO si une cellule/ligne est sélectionnée et contient product_name + Supplier
@app.callback(
    Output("btn-po-pdf", "disabled"),
    [Input("main-table", "active_cell"),
     Input("main-table", "selected_rows"),
     Input("main-table", "data")],
    prevent_initial_call=True  # ✅ CRUCIAL : ne s'exécute que sur interaction
)
def toggle_po_button(active_cell, selected_rows, data):
    # Si pas de données, désactiver
    if not data or len(data) == 0:
        return True

    # Déterminer quelle ligne est sélectionnée
    row_idx = None

    # Priorité 1 : selected_rows (clic sur la checkbox)
    if selected_rows and len(selected_rows) > 0:
        row_idx = selected_rows[0]

    # Priorité 2 : active_cell (clic sur une cellule)
    if row_idx is None and active_cell and isinstance(active_cell, dict):
        row_idx = active_cell.get("row")

    # Vérifier que l'index est valide
    if row_idx is None or row_idx < 0 or row_idx >= len(data):
        return True  # Désactivé

    # Récupérer la ligne et vérifier les champs obligatoires
    try:
        r = data[row_idx]
        prod = str(r.get("product_name", "")).strip()
        sup = str(r.get("Supplier", "")).strip()

        # Activer seulement si les deux champs sont présents
        return not (prod and sup)  # False = activé, True = désactivé

    except (IndexError, KeyError, TypeError):
        return True  # Désactivé en cas d'erreur

@app.callback(
    Output("debug-info", "children"),  # Ajoutez <div id="debug-info"></div> dans la sidebar
    [Input("main-table", "active_cell"),
     Input("main-table", "selected_rows"),
     Input("main-table", "data")],
    prevent_initial_call=True
)
def debug_selection(active_cell, selected_rows, data):
    return html.Pre(f"""
🔍 DEBUG SÉLECTION:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- active_cell: {active_cell}
- selected_rows: {selected_rows}
- Nombre lignes data: {len(data) if data else 0}

{f"• Ligne sélectionnée: {selected_rows[0] if selected_rows else 'None'}" if selected_rows else ""}
{f"• Données ligne: {data[selected_rows[0]] if selected_rows and len(selected_rows) > 0 and len(data) > selected_rows[0] else 'N/A'}" if selected_rows else ""}
""")
# ------------------------------ Edit/Add/Delete rows ------------------------------
@app.callback(
    [Output("edit-modal", "is_open"),
     Output("edit-product", "value"),  # ✅ Corrigé de "edit-input"
     Output("edit-supplier", "value"),
     Output("edit-category", "value"),
     Output("edit-stock", "value")],
    Input("btn-add-row", "n_clicks"),
    prevent_initial_call=True
)
def open_edit_modal(n_add):
    """Ouvre le modal pour ajouter une nouvelle ligne"""
    if not n_add:
        return no_update, no_update, no_update, no_update, no_update

    # Ouvrir le modal avec des champs vides
    return True, "", "", "CX", 0

@app.callback(
    Output("edit-modal", "is_open", allow_duplicate=True),
    Input("edit-cancel", "n_clicks"),
    prevent_initial_call=True
)
def close_edit_modal(n_clicks):
    """Ferme le modal sans sauvegarder"""
    if n_clicks:
        return False
    return no_update


@app.callback(
    [Output("master-data", "data", allow_duplicate=True),
     Output("filtered-data", "data", allow_duplicate=True),
     Output("main-table", "data", allow_duplicate=True),
     Output("edit-modal", "is_open", allow_duplicate=True)],
    Input("edit-save", "n_clicks"),
    State("edit-product", "value"),  # ✅ Pas "edit-input"
    State("edit-supplier", "value"),
    State("edit-category", "value"),
    State("edit-stock", "value"),
    State("master-data", "data"),
    prevent_initial_call=True
)
def save_edit(n, prod, sup, cat, stock, master_json):
    if not n:
        return no_update, no_update, no_update, no_update

    base = pd.DataFrame(json.loads(master_json)) if master_json else get_df_cached()
    df = base.copy()

    # Créer une nouvelle ligne
    new_row = {}
    for c in df.columns:
        new_row[c] = np.nan

    new_row["product_name"] = prod or ""
    new_row["Supplier"] = sup or ""
    new_row["Product Category"] = cat or "CX"

    try:
        new_row["total_stock"] = float(stock or 0)
    except:
        new_row["total_stock"] = 0.0

    # Valeurs par défaut
    new_row["Recalculated Average Daily Sales"] = 0.1
    new_row["Optimal Stock (Reorder Point)"] = 0.0
    new_row["Max Coverage Day"] = 0.0
    new_row["Predicted Order Quantity"] = 0.0
    new_row["purchase_need"] = 0.0
    new_row["Predicted Stockout"] = False
    new_row["Stock Status"] = "Stock OK" if new_row["total_stock"] > 0 else "Out of Stock"
    new_row["Ajusted_total_need"] = "NO NEED"
    new_row["delisting_status"] = "Not Delisted"

    # Ajouter
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    print(f"[Edit] Produit ajouté : {prod}")

    return df.to_json(orient="records"), df.to_json(orient="records"), df.to_dict("records"), False
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
    State("toggle-options","value"),
    State("master-data","data"),
    prevent_initial_call=True
)
def delete_row(active_cell, table_data, q, fs, fst, fc, opts, master_json):
    if not active_cell or active_cell.get("column_id") != "🗑️ Delete" or not table_data:
        raise dash.exceptions.PreventUpdate
    base = pd.DataFrame(json.loads(master_json)) if master_json else get_df_cached()
    row = active_cell["row"]
    r = table_data[row]
    key_p = r.get("product_name")
    key_s = r.get("Supplier")
    df = base[~((base["product_name"].astype(str)==str(key_p)) & (base["Supplier"].astype(str)==str(key_s)))].copy()

    sup_list = fs or []; stat_list = fst or []; cat_list = fc or []; options = opts or []
    fdf = filter_dataframe(df, q, sup_list, stat_list, cat_list, options)
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
    [Output("chat-messages", "children", allow_duplicate=True),
     Output("chat-store", "data", allow_duplicate=True),
     Output("chat-input", "value")],  # ✅ Ajout pour vider l'input
    Input("chat-send", "n_clicks"),
    State("chat-input", "value"),
    State("chat-store", "data"),
    State("uploaded-csv", "data"),
    State("master-data", "data"),
    prevent_initial_call=True
)
def on_chat(n_clicks, user_text, history, uploaded_json, master_json):
    if not n_clicks:
        return no_update, no_update, no_update

    history = history or []
    user_text = (user_text or "").strip()

    if not user_text:
        return _render_messages(history), history, ""

    # Préparer les données
    df_up = pd.DataFrame(json.loads(uploaded_json)) if uploaded_json else pd.DataFrame()
    use_uploaded = ('product_name' in df_up.columns) and len(df_up) > 0
    df_base = pd.DataFrame(json.loads(master_json)) if master_json else get_df_cached()
    df = df_up if use_uploaded else df_base

    # Ajouter le message utilisateur
    history.append({"role": "user", "text": user_text, "ts": datetime.now().isoformat()})

    # Obtenir la réponse IA
    reply = _chatbot_reply(user_text, df, history[:-1])  # Exclut le dernier message pour éviter doublon

    # Ajouter la réponse
    history.append({"role": "assistant", "text": reply, "ts": datetime.now().isoformat()})

    return _render_messages(history), history, ""  # ✅ Vider l'input

# ------------------------------ Run ----------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8050))
    app.run(debug=True, host="0.0.0.0", port=port)

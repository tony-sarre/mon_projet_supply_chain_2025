import dash
from dash import dcc, html, dash_table, Input, Output, State
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import numpy as np
import requests
from io import BytesIO, StringIO
from datetime import datetime, timedelta
import warnings
import traceback

warnings.filterwarnings('ignore')


class OptimizedSupplyChainAnalytics:
    """
    Classe optimisée pour l'analyse supply chain CORRIGÉE
    Implémentation des 5 sheets + conditions d'optimalité exactes
    """

    def __init__(self, sales_url, suppliers_url):
        self.sales_url = sales_url
        self.suppliers_url = suppliers_url
        self.sales_sheets = {}
        self.suppliers_data = None

        # Leadtime mapping EXACT selon vos données
        self.leadtime_mapping = {
            'Finamark': 4, 'CMGA': 4, 'Condak (Pinthon)': 17, 'Café Touba Bakhdad': 4,
            'CSPE': 3, 'Patisen': 10, 'LSB': 4, 'EGNAB': 5, 'ARLA': 3, 'Sipa': 2,
            'FSP': 9, 'SENAME': 4, 'DAP': 4, 'SENFI SARL': 6, 'Sencom': 8,
            'NMA SANDERS': 29, 'GIE Palene Business': 3, 'Ets Meroueh': 2,
            'H&D Industries': 12, 'SIAGRO': 7, 'SOSAGRIN': 3, 'SENICO': 32,
            'ACCENT': 3, 'Sedipal': 3, 'Sogepal': 7, 'Novatis': 3, 'O\'Royal': 5,
            'VITFE': 24, 'AGROLINE': 2, 'COSEDI': 6, 'Wadic': 2, 'SODECA': 23,
            'Cayor': 3, 'SIAD': 24, 'NESTLE': 6, 'SODAGEM': 4, 'SOCADIS': 5,
            'ECOS Afrique': 5, 'Kebe & Frères': 7, 'IBS': 3, 'SOBOBA': 3,
            'Orange': 1, 'Maad': 7, 'MBD SUARL': 5
        }

        # CONDITIONS D'OPTIMALITÉ EXACTES selon votre tableau
        self.optimization_matrix = {
            'AX': {'buffer_days': 8, 'credit_optimization': 12},
            'AY': {'buffer_days': 8, 'credit_optimization': 12},
            'AZ': {'buffer_days': 10, 'credit_optimization': 14},
            'BX': {'buffer_days': 6, 'credit_optimization': 8},
            'BY': {'buffer_days': 6, 'credit_optimization': 8},
            'BZ': {'buffer_days': 8, 'credit_optimization': 14},
            'CX': {'buffer_days': 5, 'credit_optimization': 6},
            'CY': {'buffer_days': 5, 'credit_optimization': 6},
            'CZ': {'buffer_days': 5, 'credit_optimization': 10}
        }

    def load_all_data_sources(self):
        """Chargement des 5 SHEETS + fournisseurs"""
        try:
            sales_loaded = self._load_sales_data()
            suppliers_loaded = self._load_suppliers_data()
            return sales_loaded, suppliers_loaded
        except Exception as e:
            print(f"Erreur chargement: {e}")
            return False, False

    def _load_sales_data(self):
        """Charge TOUTES les données (5 sheets)"""
        try:
            response = requests.get(self.sales_url, timeout=45)
            response.raise_for_status()

            excel_file = BytesIO(response.content)
            self.sales_sheets = pd.read_excel(excel_file, sheet_name=None, engine='openpyxl')

            required_sheets = [
                'last-30days-sales-2025', 'inventory-warehouse-2025',
                'sales-warehouse', 'inventory-staging-2025', 'last-7days-sales-2025'
            ]

            missing_sheets = [sheet for sheet in required_sheets if sheet not in self.sales_sheets]
            if missing_sheets:
                print(f"Sheets manquantes: {missing_sheets}")

            print(f"✅ {len(self.sales_sheets)} sheets chargées")
            return True

        except Exception as e:
            print(f"❌ Erreur chargement ventes: {e}")
            return False

    def _load_suppliers_data(self):
        """Charge données fournisseurs"""
        try:
            export_url = self._convert_google_sheets_url(self.suppliers_url)
            headers = {'User-Agent': 'Mozilla/5.0 (compatible; SupplyChainBot/1.0)'}
            response = requests.get(export_url, headers=headers, timeout=30)
            response.raise_for_status()

            self.suppliers_data = pd.read_csv(StringIO(response.text))
            self.suppliers_data.columns = self.suppliers_data.columns.str.strip()

            # Mapping colonnes
            expected_columns = {
                'id': ['Nom 1', '1', 'id', 'ID', 'product_id'],
                'product_name': ['Nom 2', '2', 'product_name', 'Nom Produit'],
                'brand': ['Nom 4', '4', 'brand', 'Marque'],
                'supplier_name': ['Nom 6', '6', 'Supplier name', 'supplier_name', 'Fournisseur'],
                'credit_days': ['Days Credit', 'Credit Days', 'credit_days', 'Crédit Jours'],
                'purchase_price': ['Current Purchase Price', 'Purchase Price', 'Prix Achat']
            }

            for standard_col, possible_names in expected_columns.items():
                for name in possible_names:
                    if name in self.suppliers_data.columns:
                        self.suppliers_data[standard_col] = self.suppliers_data[name]
                        break

            # S'assurer que les colonnes numériques sont bien des nombres
            if 'credit_days' in self.suppliers_data.columns:
                self.suppliers_data['credit_days'] = pd.to_numeric(self.suppliers_data['credit_days'], errors='coerce')
            if 'purchase_price' in self.suppliers_data.columns:
                self.suppliers_data['purchase_price'] = pd.to_numeric(self.suppliers_data['purchase_price'],
                                                                      errors='coerce')

            print(f"✅ {len(self.suppliers_data)} fournisseurs chargés")
            return True

        except Exception as e:
            print(f"⚠️ Erreur fournisseurs: {e}")
            return False

    def _convert_google_sheets_url(self, url):
        """Conversion URL Google Sheets"""
        if '/edit' in url and '/d/' in url:
            sheet_id = url.split('/d/')[1].split('/edit')[0]
            if 'gid=' in url:
                gid = url.split('gid=')[1].split('#')[0].split('&')[0]
                return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
            else:
                return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
        return url

    def prepare_comprehensive_dataset_corrected(self):
        """
        Dataset complet CORRIGÉ utilisant les 5 sheets selon OOS%
        """
        try:
            # 1. Extraction ventes avec logique OOS CORRIGÉE
            sales_data = self._extract_sales_with_oos_logic()

            # 2. Inventaire warehouse + staging
            inventory_data = self._extract_complete_inventory()

            # 3. Merge principal
            master_df = sales_data.merge(
                inventory_data,
                left_on='id',
                right_on='product_id',
                how='left'
            ).drop_duplicates(subset=['id'], keep='first')

            # 4. Intégration fournisseurs
            master_df = self._integrate_suppliers(master_df)

            # 5. Calculs supply chain CORRIGÉS
            master_df = self._calculate_corrected_metrics(master_df)

            return master_df

        except Exception as e:
            print(f"❌ Erreur dataset: {e}")
            traceback.print_exc()
            return pd.DataFrame()

    def _extract_sales_with_oos_logic(self):
        """
        LOGIQUE CORRIGÉE : Source selon % OOS
        - OOS% < 20% → ventes 30j
        - OOS% ≥ 20% → ventes 7j (plus récentes)
        """
        try:
            sales_30d = self.sales_sheets['last-30days-sales-2025'].copy()
            sales_7d = self.sales_sheets['last-7days-sales-2025'].copy()

            # Nettoyage et standardisation des colonnes
            for df in [sales_30d, sales_7d]:
                # S'assurer que les colonnes essentielles existent
                if 'avg_quantity_per_day' not in df.columns and 'avg_daily_quantity' in df.columns:
                    df['avg_quantity_per_day'] = df['avg_daily_quantity']
                if 'max_quantity_per_day' not in df.columns and 'max_daily_quantity' in df.columns:
                    df['max_quantity_per_day'] = df['max_daily_quantity']

                numeric_cols = ['avg_quantity_per_day', 'max_quantity_per_day', 'total_amount']
                for col in numeric_cols:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                    else:
                        df[col] = 0

                df['perc_oos_duration'] = pd.to_numeric(df.get('perc_oos_duration', 0), errors='coerce').fillna(0)

                # S'assurer que les colonnes name/product_name existent
                if 'name' not in df.columns and 'product_name' in df.columns:
                    df['name'] = df['product_name']
                elif 'name' not in df.columns:
                    df['name'] = 'Produit Inconnu'

            # LOGIQUE CONDITIONNELLE OOS
            final_sales = []

            for _, row in sales_30d.iterrows():
                product_id = row['id']
                oos_pct = row['perc_oos_duration']

                if oos_pct >= 0.20:  # 20% ou plus → utiliser données 7j
                    match_7d = sales_7d[sales_7d['id'] == product_id]
                    if not match_7d.empty:
                        # Extrapoler 7j vers 30j pour cohérence
                        row_7d = match_7d.iloc[0].copy()
                        row_7d['avg_quantity_per_day'] = row_7d['avg_quantity_per_day']  # Garde tel quel
                        row_7d['max_quantity_per_day'] = row_7d['max_quantity_per_day']  # Garde tel quel
                        row_7d['total_amount'] = row_7d['total_amount'] * (30 / 7)  # Extrapole CA
                        row_7d['data_source'] = '7d_extrapolated'
                        final_sales.append(row_7d)
                    else:
                        row['data_source'] = '30d_default'
                        final_sales.append(row)
                else:  # OOS < 20% → utiliser données 30j
                    row['data_source'] = '30d_normal'
                    final_sales.append(row)

            result_df = pd.DataFrame(final_sales)
            print(f"✅ Sales avec logique OOS: {len(result_df)} produits")
            return result_df

        except Exception as e:
            print(f"❌ Erreur extraction ventes: {e}")
            traceback.print_exc()
            return self.sales_sheets.get('last-30days-sales-2025', pd.DataFrame())

    def _extract_complete_inventory(self):
        """Inventaire warehouse + staging combiné"""
        try:
            warehouse = self.sales_sheets['inventory-warehouse-2025'].copy()
            staging = self.sales_sheets.get('inventory-staging-2025', pd.DataFrame())

            # Nettoyage warehouse
            warehouse = warehouse[['product_id', 'product_name', 'total_stock']].copy()
            warehouse['warehouse_stock'] = pd.to_numeric(warehouse['total_stock'], errors='coerce').fillna(0)

            # Si staging disponible, l'ajouter
            if not staging.empty and 'product_id' in staging.columns:
                staging_clean = staging[['product_id', 'total_stock']].copy()
                staging_clean['staging_stock'] = pd.to_numeric(staging_clean['total_stock'], errors='coerce').fillna(0)

                # Merge warehouse + staging
                inventory_combined = warehouse.merge(
                    staging_clean, on='product_id', how='left'
                ).fillna(0)

                inventory_combined['total_stock'] = (
                        inventory_combined['warehouse_stock'] +
                        inventory_combined['staging_stock']
                )
            else:
                inventory_combined = warehouse.rename(columns={'total_stock': 'warehouse_stock'})
                inventory_combined['staging_stock'] = 0
                inventory_combined['total_stock'] = inventory_combined['warehouse_stock']

            print(f"✅ Inventaire combiné: {len(inventory_combined)} produits")
            return inventory_combined

        except Exception as e:
            print(f"❌ Erreur inventaire: {e}")
            traceback.print_exc()
            return pd.DataFrame()

    def _integrate_suppliers(self, df):
        """Intégration fournisseurs CORRIGÉE"""
        if self.suppliers_data is not None and not self.suppliers_data.empty:
            try:
                supplier_clean = self.suppliers_data.copy()
                supplier_clean['id'] = pd.to_numeric(supplier_clean['id'], errors='coerce')
                df['id_numeric'] = pd.to_numeric(df['id'], errors='coerce')

                merge_cols = ['id']
                # Ajouter seulement les colonnes qui existent
                available_supplier_cols = ['supplier_name', 'credit_days', 'purchase_price']
                for col in available_supplier_cols:
                    if col in supplier_clean.columns:
                        merge_cols.append(col)

                df_merged = df.merge(
                    supplier_clean[merge_cols],
                    left_on='id_numeric',
                    right_on='id',
                    how='left',
                    suffixes=('', '_supplier')
                )

                # CORRECTION : Valeurs par défaut avec gestion pandas Series
                # Supplier name
                if 'supplier_name' not in df_merged.columns:
                    df_merged['supplier_name'] = 'Fournisseur Inconnu'
                else:
                    df_merged['supplier_name'] = df_merged['supplier_name'].fillna('Fournisseur Inconnu')

                # Credit days
                if 'credit_days' not in df_merged.columns:
                    df_merged['credit_days'] = 15
                else:
                    df_merged['credit_days'] = pd.to_numeric(df_merged['credit_days'], errors='coerce').fillna(15)

                # Purchase price
                if 'purchase_price' not in df_merged.columns:
                    df_merged['purchase_price'] = 0
                else:
                    df_merged['purchase_price'] = pd.to_numeric(df_merged['purchase_price'], errors='coerce').fillna(0)

                print(f"✅ Merge fournisseurs réussi: {len(df_merged)} lignes")
                return df_merged

            except Exception as e:
                print(f"⚠️ Erreur merge fournisseurs: {e}")
                traceback.print_exc()

        # Fallback avec colonnes par défaut
        df['supplier_name'] = 'Fournisseur Inconnu'
        df['credit_days'] = 15
        df['purchase_price'] = 0
        return df

    def _calculate_corrected_metrics(self, df):
        """
        Calculs CORRIGÉS selon vos spécifications exactes
        """
        try:
            # === 1. MÉTRIQUES DE BASE D'ABORD ===
            df['current_stock'] = pd.to_numeric(df['total_stock'], errors='coerce').fillna(0).clip(lower=0)
            df['daily_demand_avg'] = pd.to_numeric(df['avg_quantity_per_day'], errors='coerce').fillna(0).clip(lower=0)
            df['daily_demand_max'] = pd.to_numeric(df['max_quantity_per_day'], errors='coerce').fillna(0).clip(lower=0)

            # === 2. CLASSIFICATION ABC-XYZ CORRIGÉE (après création des colonnes) ===
            df = self._calculate_abc_xyz_corrected(df)

            # === 3. LEADTIME MAPPING ===
            df['leadtime_days'] = df['supplier_name'].map(self.leadtime_mapping).fillna(10)

            # === 4. CONDITIONS D'OPTIMALITÉ APPLIQUÉES ===
            df = self._apply_optimization_conditions(df)

            # === 5. Coverage ===
            df['stock_coverage_days'] = np.where(
                df['daily_demand_avg'] > 0,
                df['current_stock'] / df['daily_demand_avg'],
                999
            )

            # === 6. CALCUL STOCK OPTIMAL CORRIGÉ ===
            # "Si tu as un crédit de 20 jours, autant prendre 20 jours de stock"
            df['stock_optimal'] = df['credit_days_optimized'] * df['daily_demand_avg']

            # === 7. REORDER POINT ===
            df['safety_stock'] = df['buffer_days'] * df['daily_demand_avg']
            df['reorder_point'] = df['leadtime_days'] * df['daily_demand_avg'] + df['safety_stock']

            # === 8. QUANTITÉ À COMMANDER ===
            df['quantity_to_order'] = np.maximum(
                df['stock_optimal'] - df['current_stock'], 0
            ).round(0)

            # === 9. STATUT ACHAT ===
            conditions = [
                df['current_stock'] <= (df['reorder_point'] * 0.5),
                df['current_stock'] <= df['reorder_point'],
                df['quantity_to_order'] > 0
            ]

            choices = ['CRITICAL', 'URGENT', 'RECOMMENDED']
            df['purchase_status'] = np.select(conditions, choices, default='SUFFICIENT')

            # === 10. PRIORITÉ BUSINESS ===
            df = self._calculate_priority_score(df)

            print(f"✅ Calculs métriques terminés: {len(df)} produits")
            return df

        except Exception as e:
            print(f"❌ Erreur calculs: {e}")
            traceback.print_exc()
            return df

    def _calculate_abc_xyz_corrected(self, df):
        """Classification ABC-XYZ selon logique métier correcte"""
        try:
            # Vérifier que les colonnes nécessaires existent
            if 'daily_demand_avg' not in df.columns or 'daily_demand_max' not in df.columns:
                print("⚠️ Colonnes daily_demand manquantes pour ABC-XYZ")
                df['ABC_Class'] = 'C'
                df['XYZ_Class'] = 'Z'
                df['ABC_XYZ_Combined'] = 'CZ'
                df['coefficient_variation'] = 1.0
                return df

            # ABC - Pareto sur CA
            df['total_amount'] = pd.to_numeric(df.get('total_amount', 0), errors='coerce').fillna(0)
            total_revenue = df['total_amount'].sum()

            if total_revenue > 0:
                df_sorted = df.sort_values('total_amount', ascending=False).reset_index()
                cumulative_pct = (df_sorted['total_amount'].cumsum() / total_revenue) * 100

                abc_map = pd.Series('C', index=df.index)
                abc_map.iloc[df_sorted[cumulative_pct <= 80].index] = 'A'
                abc_map.iloc[df_sorted[(cumulative_pct > 80) & (cumulative_pct <= 95)].index] = 'B'
            else:
                abc_map = pd.Series('C', index=df.index)

            # XYZ - Variabilité selon coefficient de variation
            cv = np.where(
                df['daily_demand_avg'] > 0,
                (df['daily_demand_max'] - df['daily_demand_avg']) / df['daily_demand_avg'],
                2.0
            )
            cv = np.clip(cv, 0, 5)  # Borner les valeurs aberrantes

            xyz_map = pd.Series('Z', index=df.index)
            xyz_map[cv <= 0.5] = 'X'  # Stable
            xyz_map[(cv > 0.5) & (cv <= 1.0)] = 'Y'  # Variable
            # Z reste pour cv > 1.0      # Très variable

            df['ABC_Class'] = abc_map
            df['XYZ_Class'] = xyz_map
            df['ABC_XYZ_Combined'] = abc_map + xyz_map
            df['coefficient_variation'] = cv

            print(f"✅ Classification ABC-XYZ terminée")
            return df

        except Exception as e:
            print(f"⚠️ Erreur ABC-XYZ: {e}")
            traceback.print_exc()
            df['ABC_Class'] = 'C'
            df['XYZ_Class'] = 'Z'
            df['ABC_XYZ_Combined'] = 'CZ'
            df['coefficient_variation'] = 1.0
            return df

    def _apply_optimization_conditions(self, df):
        """Application CONDITIONS D'OPTIMALITÉ exactes"""
        try:
            # S'assurer que la colonne ABC_XYZ_Combined existe
            if 'ABC_XYZ_Combined' not in df.columns:
                df['ABC_XYZ_Combined'] = 'CZ'

            # Buffer selon classification
            df['buffer_days'] = df['ABC_XYZ_Combined'].map(
                lambda x: self.optimization_matrix.get(x, {'buffer_days': 7})['buffer_days']
            ).fillna(7)

            # Crédit optimisé selon classification
            df['credit_days_optimized'] = df['ABC_XYZ_Combined'].map(
                lambda x: self.optimization_matrix.get(x, {'credit_optimization': 15})['credit_optimization']
            ).fillna(15)

            # Utiliser crédit fournisseur si disponible, sinon optimisé
            df['credit_days_final'] = np.maximum(
                pd.to_numeric(df.get('credit_days', 15), errors='coerce').fillna(15),
                df['credit_days_optimized']
            )

            return df

        except Exception as e:
            print(f"⚠️ Erreur conditions optimisation: {e}")
            df['buffer_days'] = 7
            df['credit_days_optimized'] = 15
            df['credit_days_final'] = 15
            return df

    def _calculate_priority_score(self, df):
        """Score priorité business (0-100)"""
        try:
            score = pd.Series(0.0, index=df.index)

            # 40% Impact CA
            if df['total_amount'].sum() > 0:
                score += (df['total_amount'] / df['total_amount'].max()) * 40

            # 30% Urgence
            urgency_map = {'CRITICAL': 30, 'URGENT': 20, 'RECOMMENDED': 10, 'SUFFICIENT': 0}
            score += df['purchase_status'].map(urgency_map).fillna(0)

            # 20% Classification
            class_map = {'A': 20, 'B': 12, 'C': 5}
            score += df['ABC_Class'].map(class_map).fillna(5)

            # 10% Risque OOS
            score += np.clip(df.get('perc_oos_duration', 0) * 100, 0, 10)

            df['priority_score'] = score.clip(upper=100).round(1)
            return df

        except Exception as e:
            print(f"⚠️ Erreur priorité: {e}")
            df['priority_score'] = 50
            return df


# === INTERFACE DASH avec prévention boucles infinies ===
app = dash.Dash(__name__)

app.layout = html.Div([
    dcc.Store(id='dataset-store'),
    dcc.Store(id='processing-state', data={'processing': False}),

    # Header
    html.Div([
        html.H1("🚀 Supply Chain Analytics", className='main-title'),
        html.P("Plateforme d'Intelligence Procurement & Optimisation Inventaire", className='subtitle')
    ], className='header'),

    # Configuration
    html.Div([
        html.Div([
            html.Label("URL Données Ventes & Inventaire:"),
            dcc.Input(
                id='sales-url',
                value="https://docs.google.com/spreadsheets/d/e/2PACX-1vQbJqHsr6Kifee7I91YD-7-sCZDWgM5GvxCeN0OqUvZhok0j-kDywguqe5I61y97b-uBhHbWraTIrux/pub?output=xlsx",
                type='url',
                style={'width': '100%'}
            )
        ], className='config-item'),

        html.Div([
            html.Label("URL Données Fournisseurs:"),
            dcc.Input(
                id='suppliers-url',
                value="https://docs.google.com/spreadsheets/d/e/2PACX-1vRTyAxh6v8o0FXV0r7f6ALPDgmeJNkjTZITjrEoKBHo2gs_f3iyV8sFk8fOzcAsUSkJMXBJCpJnhQKi/pub?gid=990028336&single=true&output=csv",
                type='url',
                style={'width': '100%'}
            )
        ], className='config-item'),

        html.Button("🔄 Analyser Supply Chain", id='analyze-btn', className='analyze-button')
    ], className='config-section'),

    # Résultats
    html.Div(id='results-container'),
    html.Div(id='loading-indicator')
])


@app.callback(
    [Output('dataset-store', 'data'),
     Output('results-container', 'children'),
     Output('processing-state', 'data'),
     Output('loading-indicator', 'children')],
    [Input('analyze-btn', 'n_clicks')],
    [State('sales-url', 'value'),
     State('suppliers-url', 'value'),
     State('processing-state', 'data')]
)
def analyze_supply_chain(n_clicks, sales_url, suppliers_url, processing_state):
    if n_clicks is None:
        return {}, html.Div("👆 Cliquez sur Analyser pour commencer"), {'processing': False}, ""

    # Prévention boucle infinie
    if processing_state.get('processing', False):
        return dash.no_update, dash.no_update, dash.no_update, html.Div("⏳ Traitement en cours...", className='loading')

    try:
        # Marquer comme en cours de traitement
        processing_state['processing'] = True

        # Initialisation
        analytics = OptimizedSupplyChainAnalytics(sales_url, suppliers_url)

        # Chargement
        sales_loaded, suppliers_loaded = analytics.load_all_data_sources()

        if not sales_loaded:
            return {}, html.Div("❌ Erreur chargement données ventes", className='error'), {'processing': False}, ""

        # Analyse
        dashboard_data = analytics.prepare_comprehensive_dataset_corrected()

        if dashboard_data.empty:
            return {}, html.Div("❌ Aucune donnée exploitable", className='error'), {'processing': False}, ""

        # Store pour callbacks
        dataset_dict = dashboard_data.to_dict('records')

        # KPIs
        total_products = len(dashboard_data)
        critical_orders = len(dashboard_data[dashboard_data['purchase_status'] == 'CRITICAL'])
        urgent_orders = len(dashboard_data[dashboard_data['purchase_status'] == 'URGENT'])
        recommended_orders = len(dashboard_data[dashboard_data['purchase_status'] == 'RECOMMENDED'])
        total_order_value = dashboard_data['quantity_to_order'].sum() * dashboard_data['purchase_price'].mean()

        # Dashboard components (le même code que précédemment pour la création des résultats)
        results = html.Div([
            # KPIs Executive
            html.Div([
                html.H2("📊 Executive Dashboard"),
                html.Div([
                    html.Div([
                        html.H3(f"{total_products:,}"),
                        html.P("Total SKUs")
                    ], className='kpi-card'),

                    html.Div([
                        html.H3(f"{critical_orders}", style={'color': '#d32f2f'}),
                        html.P("CRITICAL Orders")
                    ], className='kpi-card critical'),

                    html.Div([
                        html.H3(f"{urgent_orders}", style={'color': '#f57c00'}),
                        html.P("URGENT Orders")
                    ], className='kpi-card urgent'),

                    html.Div([
                        html.H3(f"{recommended_orders}", style={'color': '#388e3c'}),
                        html.P("RECOMMENDED")
                    ], className='kpi-card recommended'),

                    html.Div([
                        html.H3(f"{total_order_value:,.0f} €"),
                        html.P("Total Order Value")
                    ], className='kpi-card')
                ], className='kpi-row')
            ], className='dashboard-section'),

            # Table principale
            html.Div([
                html.H2("📋 Supply Chain Analytics Table"),
                dash_table.DataTable(
                    data=dashboard_data.head(200).to_dict('records'),
                    columns=[
                        {'name': 'Product ID', 'id': 'id'},
                        {'name': 'Product Name', 'id': 'name'},
                        {'name': 'Supplier', 'id': 'supplier_name'},
                        {'name': 'Class', 'id': 'ABC_XYZ_Combined'},
                        {'name': 'Current Stock', 'id': 'current_stock', 'type': 'numeric',
                         'format': {'specifier': ',.0f'}},
                        {'name': 'Daily Demand Avg', 'id': 'daily_demand_avg', 'type': 'numeric',
                         'format': {'specifier': ',.2f'}},
                        {'name': 'Coverage Days', 'id': 'stock_coverage_days', 'type': 'numeric',
                         'format': {'specifier': ',.1f'}},
                        {'name': 'Leadtime', 'id': 'leadtime_days', 'type': 'numeric', 'format': {'specifier': ',.0f'}},
                        {'name': 'Buffer Days', 'id': 'buffer_days', 'type': 'numeric',
                         'format': {'specifier': ',.0f'}},
                        {'name': 'Credit Days', 'id': 'credit_days_final', 'type': 'numeric',
                         'format': {'specifier': ',.0f'}},
                        {'name': 'Stock Optimal', 'id': 'stock_optimal', 'type': 'numeric',
                         'format': {'specifier': ',.0f'}},
                        {'name': 'Reorder Point', 'id': 'reorder_point', 'type': 'numeric',
                         'format': {'specifier': ',.0f'}},
                        {'name': 'Qty to Order', 'id': 'quantity_to_order', 'type': 'numeric',
                         'format': {'specifier': ',.0f'}},
                        {'name': 'Purchase Status', 'id': 'purchase_status'},
                        {'name': 'Priority Score', 'id': 'priority_score', 'type': 'numeric',
                         'format': {'specifier': ',.1f'}}
                    ],
                    style_cell={'textAlign': 'left', 'padding': '10px'},
                    style_header={'backgroundColor': '#1f77b4', 'color': 'white', 'fontWeight': 'bold'},
                    style_data_conditional=[
                        {
                            'if': {'filter_query': '{purchase_status} = CRITICAL'},
                            'backgroundColor': '#ffcdd2',
                            'color': '#b71c1c',
                            'fontWeight': 'bold'
                        },
                        {
                            'if': {'filter_query': '{purchase_status} = URGENT'},
                            'backgroundColor': '#ffe0b2',
                            'color': '#e65100',
                            'fontWeight': 'bold'
                        },
                        {
                            'if': {'filter_query': '{purchase_status} = RECOMMENDED'},
                            'backgroundColor': '#c8e6c9',
                            'color': '#2e7d32'
                        },
                        {
                            'if': {'filter_query': '{priority_score} >= 80'},
                            'backgroundColor': '#9c27b0',
                            'color': 'white'
                        }
                    ],
                    sort_action="native",
                    filter_action="native",
                    page_size=50,
                    style_table={'overflowX': 'auto'}
                )
            ], className='table-section'),

            # Graphiques d'analyse (même code que précédemment)
            html.Div([
                html.H2("📈 Analytics & Insights"),

                html.Div([
                    # Graphique 1: Distribution des statuts d'achat
                    html.Div([
                        dcc.Graph(
                            figure=px.pie(
                                values=dashboard_data['purchase_status'].value_counts().values,
                                names=dashboard_data['purchase_status'].value_counts().index,
                                title="Distribution des Statuts d'Achat",
                                color_discrete_map={
                                    'CRITICAL': '#d32f2f',
                                    'URGENT': '#f57c00',
                                    'RECOMMENDED': '#388e3c',
                                    'SUFFICIENT': '#9e9e9e'
                                }
                            )
                        )
                    ], className='chart-half'),

                    # Graphique 2: Classification ABC-XYZ
                    html.Div([
                        dcc.Graph(
                            figure=px.bar(
                                x=dashboard_data['ABC_XYZ_Combined'].value_counts().index,
                                y=dashboard_data['ABC_XYZ_Combined'].value_counts().values,
                                title="Distribution Classification ABC-XYZ",
                                labels={'x': 'Classification', 'y': 'Nombre de Produits'}
                            )
                        )
                    ], className='chart-half')
                ], className='charts-row')
            ], className='analytics-section'),

            # Insights Business (même code)
            html.Div([
                html.H2("💡 Business Insights & Recommandations"),

                html.Div([
                    html.Div([
                        html.H3("🚨 Alertes Critiques"),
                        html.Ul([
                                    html.Li(f"{critical_orders} commandes CRITIQUES nécessitent une action immédiate"),
                                    html.Li(f"{urgent_orders} commandes URGENTES à traiter sous 24h"),
                                    html.Li(f"Budget total requis: {total_order_value:,.0f} €")
                                ] if critical_orders > 0 or urgent_orders > 0 else [
                            html.Li("✅ Aucune alerte critique en cours", style={'color': '#4caf50'})
                        ])
                    ], className='insight-card'),

                    html.Div([
                        html.H3("📊 Optimisations Appliquées"),
                        html.Ul([
                            html.Li(f"✅ Logique OOS adaptative (5 sheets)"),
                            html.Li(f"✅ Conditions d'optimalité par classification"),
                            html.Li(f"✅ Leadtime mapping fournisseur"),
                            html.Li(f"✅ Stock optimal = Crédit × Demande quotidienne")
                        ])
                    ], className='insight-card')
                ], className='insights-row')
            ], className='insights-section')
        ])

        return dataset_dict, results, {'processing': False}, ""

    except Exception as e:
        error_msg = f"❌ Erreur analyse: {str(e)}"
        print(error_msg)
        traceback.print_exc()
        return {}, html.Div(error_msg, className='error'), {'processing': False}, ""


# CSS (même que précédemment)
app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <style>
            body { 
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
                margin: 0; 
                padding: 20px; 
                background-color: #f5f5f5;
            }

            .header {
                text-align: center;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
                padding: 2rem;
                border-radius: 10px;
                margin-bottom: 2rem;
            }

            .main-title {
                font-size: 2.5rem;
                margin: 0;
                font-weight: bold;
            }

            .subtitle {
                font-size: 1.1rem;
                margin: 0.5rem 0 0 0;
                opacity: 0.9;
            }

            .config-section {
                background: white;
                padding: 1.5rem;
                border-radius: 10px;
                margin-bottom: 2rem;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }

            .config-item {
                margin-bottom: 1rem;
            }

            .config-item label {
                display: block;
                margin-bottom: 0.5rem;
                font-weight: bold;
                color: #333;
            }

            .analyze-button {
                background: linear-gradient(135deg, #1f77b4 0%, #ff7f0e 100%);
                color: white;
                border: none;
                padding: 12px 24px;
                font-size: 1.1rem;
                border-radius: 5px;
                cursor: pointer;
                font-weight: bold;
                width: 100%;
                margin-top: 1rem;
            }

            .analyze-button:hover {
                transform: translateY(-2px);
                box-shadow: 0 4px 15px rgba(0,0,0,0.2);
            }

            .dashboard-section, .table-section, .analytics-section, .insights-section {
                background: white;
                padding: 1.5rem;
                border-radius: 10px;
                margin-bottom: 2rem;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }

            .kpi-row {
                display: flex;
                gap: 1rem;
                flex-wrap: wrap;
                margin-top: 1rem;
            }

            .kpi-card {
                flex: 1;
                min-width: 150px;
                background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
                padding: 1rem;
                border-radius: 8px;
                text-align: center;
                border-left: 4px solid #1976d2;
            }

            .kpi-card.critical {
                background: linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%);
                border-left-color: #d32f2f;
            }

            .kpi-card.urgent {
                background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%);
                border-left-color: #f57c00;
            }

            .kpi-card.recommended {
                background: linear-gradient(135deg, #e8f5e8 0%, #c8e6c9 100%);
                border-left-color: #388e3c;
            }

            .kpi-card h3 {
                margin: 0 0 0.5rem 0;
                font-size: 2rem;
                font-weight: bold;
            }

            .kpi-card p {
                margin: 0;
                font-size: 0.9rem;
                opacity: 0.8;
            }

            .charts-row {
                display: flex;
                gap: 1rem;
                margin-top: 1rem;
            }

            .chart-half {
                flex: 1;
            }

            .insights-row {
                display: flex;
                gap: 1rem;
                flex-wrap: wrap;
                margin-top: 1rem;
            }

            .insight-card {
                flex: 1;
                min-width: 300px;
                background: #f8f9fa;
                padding: 1rem;
                border-radius: 8px;
                border-left: 4px solid #17a2b8;
            }

            .insight-card h3 {
                margin: 0 0 1rem 0;
                color: #495057;
            }

            .insight-card ul {
                margin: 0;
                padding-left: 1.5rem;
            }

            .insight-card li {
                margin-bottom: 0.5rem;
                line-height: 1.4;
            }

            .error {
                background: #ffebee;
                color: #c62828;
                padding: 1rem;
                border-radius: 5px;
                border-left: 4px solid #d32f2f;
                font-weight: bold;
            }

            .loading {
                background: #e3f2fd;
                color: #1976d2;
                padding: 1rem;
                border-radius: 5px;
                border-left: 4px solid #1976d2;
                font-weight: bold;
                text-align: center;
            }

            h2 {
                color: #1976d2;
                border-bottom: 2px solid #e3f2fd;
                padding-bottom: 0.5rem;
            }
        </style>
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
'''

#if __name__ == '__main__':
#    app.run(debug=True, port=8050)
# Votre code existant...

# À LA FIN, modifier cette partie :
server = app.server  # Important pour Render

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 8050))
    app.run(
        debug=False,
        host='0.0.0.0',
        port=port
    )


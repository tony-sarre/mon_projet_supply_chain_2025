# ========================= PREMIUM SUPPLY CHAIN DASHBOARD =========================
# Requirements:
# pip install dash==2.17.1 dash-bootstrap-components==1.6.0 plotly==5.22.0
# pip install pandas scikit-learn flask-caching numpy
# pip install reportlab supabase
# Optional: pip install openai

# ============================================================
# 🚀 OPTIMISATIONS PERFORMANCE v2.0
# - Chargement parallèle des CSV (ThreadPoolExecutor)
# - Cache agressif en mémoire avec refresh en arrière-plan
# - Compression gzip des réponses
# - Connexion pooling (requests.Session)
# - PO counter en mémoire (thread-safe)
# ============================================================

import csv
import os, sys
import uuid
from functools import lru_cache
import hashlib  # ✅ NOUVEAU: Pour l'authentification
import gzip
from io import BytesIO

from dash.exceptions import PreventUpdate

# ============================================================
# 🔧 CONFIGURATION PERFORMANCE GLOBALE
# ============================================================
PERFORMANCE_CONFIG = {
    "PARALLEL_LOADING": True,           # Chargement CSV en parallèle
    "MAX_WORKERS": 6,                   # Threads pour chargement parallèle
    "CACHE_TTL_SECONDS": 600,           # Cache 10 minutes
    "BACKGROUND_REFRESH": True,         # Refresh cache en arrière-plan
    "COMPRESS_RESPONSES": True,         # Compression gzip
    "CONNECTION_TIMEOUT": 15,           # Timeout connexions HTTP
    "READ_TIMEOUT": 30,                 # Timeout lecture HTTP
    "CHUNK_SIZE": 50000,                # Taille des chunks pour gros DataFrames
    "DEBOUNCE_MS": 300,                 # Debounce pour callbacks UI
}

# ✅ NOUVEAU: Import Supabase pour tracking utilisateurs
try:
    from supabase import create_client, Client
    SUPABASE_AVAILABLE = True
except ImportError:
    SUPABASE_AVAILABLE = False
    print("⚠️ Supabase non installé. Exécuter: pip install supabase")

#import MATCH

print("CWD:", os.getcwd())
print("Dir files:", os.listdir("."))
print("sys.path[0]:", sys.path[0])

from dash import Dash
import dash_bootstrap_components as dbc
import os
import google.generativeai as genai
import orjson

ORJSON_OPTS = orjson.OPT_NON_STR_KEYS | orjson.OPT_SERIALIZE_NUMPY

app = Dash(__name__, suppress_callback_exceptions=True, external_stylesheets=[dbc.themes.BOOTSTRAP], prevent_initial_callbacks='initial_duplicate')

# ⚠️ Très important pour Render/Gunicorn
server = app.server

# ✅ NOUVEAU: Secret key pour les sessions Flask (authentification)
server.secret_key = os.getenv("SECRET_KEY", "maad-supply-chain-secret-key-change-in-production-2024")

# ============================================================
# 🚀 OPTIMISATIONS SERVEUR FLASK
# ============================================================

# Compression gzip des réponses
try:
    from flask_compress import Compress
    Compress(server)
    print("✅ Compression gzip activée")
except ImportError:
    print("⚠️ flask-compress non installé - compression désactivée")

# Configuration pour performance
server.config['SEND_FILE_MAX_AGE_DEFAULT'] = 31536000  # Cache statique 1 an
server.config['JSON_SORT_KEYS'] = False  # Plus rapide sans tri

# ✅ NOTE: L'authentification est intégrée directement dans ce fichier (voir section AUTH_USERS plus bas)
# Pas besoin d'importer depuis auth.py

import re
from dotenv import load_dotenv
import io
import base64
from datetime import datetime, timedelta
import json
import threading
from pathlib import Path
import numpy as np
import pandas as pd

from flask_caching import Cache
# from dash_extensions import Cache
import dash
from dash import Dash, html, dcc, Input, Output, State, dash_table, no_update, MATCH
from dash.dependencies import Input, Output, State, ALL
import dash_bootstrap_components as dbc
import plotly.express as px
import warnings
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.utils import ImageReader
from pathlib import Path
from reportlab.lib.units import mm
from reportlab.platypus import Image
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
# --- Word / python-docx (import paresseux & sûr) ---
DOCX_AVAILABLE = False
DOCX_IMPORT_ERROR = None
try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_AVAILABLE = True
except Exception as _e:
    DOCX_AVAILABLE = False
    DOCX_IMPORT_ERROR = f"{type(_e).__name__}: {_e}"
# Imports existants + ces nouveaux
#import anthropic
import json
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# from openai import OpenAI
import requests
# ✅ AJOUTER POUR WINDOWS
import threading
import plotly.graph_objects as go
import time
import threading
from queue import Queue
from functools import wraps
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeoutError
warnings.filterwarnings("ignore", message="Parsing dates.*ambiguous", category=DeprecationWarning)

# ============================================================
# 🔒 CONFIGURATION CACHE ULTRA-PERFORMANT MULTI-UTILISATEURS
# ============================================================

# ✅ CACHE EN MÉMOIRE GLOBAL (plus rapide que FileSystemCache)
# Structure: {"key": {"data": df, "timestamp": time, "refreshing": False}}
_GLOBAL_DATA_CACHE = {}
_GLOBAL_CACHE_LOCK = threading.Lock()

# Session HTTP avec connection pooling (réutilise les connexions)
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

def create_http_session():
    """Crée une session HTTP optimisée avec pooling et retry"""
    session = requests.Session()

    # Headers pour éviter les blocages Google Sheets
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/csv,text/plain,*/*',
        'Accept-Language': 'en-US,en;q=0.9,fr;q=0.8',
    })

    retry_strategy = Retry(
        total=3,
        backoff_factor=1.0,  # Augmenté pour Google Sheets
        status_forcelist=[429, 500, 502, 503, 504],
    )
    adapter = HTTPAdapter(
        max_retries=retry_strategy,
        pool_connections=10,
        pool_maxsize=20
    )
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session

# Session globale réutilisable
_HTTP_SESSION = create_http_session()

# Cache avec FileSystem comme backup
cache_config = {
    "CACHE_TYPE": "SimpleCache",  # Plus rapide que FileSystem sur Render
    "CACHE_DEFAULT_TIMEOUT": 600,  # 10 minutes
    "CACHE_THRESHOLD": 100  # Moins d'items mais plus importants
}

try:
    cache = Cache(app.server, config=cache_config)
    print("✅ Cache SimpleCache initialisé (optimisé pour performance)")
except Exception as e:
    print(f"⚠️ Cache failed: {e}")
    cache = Cache(app.server, config={"CACHE_TYPE": "NullCache"})

# Lock pour les opérations thread-safe
_data_lock = threading.Lock()

# Cache des données par utilisateur (évite les conflits)
_user_data_cache = {}
_user_cache_lock = threading.Lock()

# ============================================================
# 🚀 SYSTÈME DE CHARGEMENT PARALLÈLE DES CSV
# ============================================================

def load_csv_fast(url, session=None, skiprows=None, max_retries=2, **kwargs):
    """
    Charge un CSV rapidement avec session poolée et fallback.
    Google Sheets peut bloquer les requêtes parallèles - on a donc un fallback.
    """
    from io import StringIO

    # Essai 1: Avec session HTTP optimisée
    if session is None:
        session = _HTTP_SESSION

    for attempt in range(max_retries):
        try:
            response = session.get(
                url,
                timeout=(PERFORMANCE_CONFIG["CONNECTION_TIMEOUT"], PERFORMANCE_CONFIG["READ_TIMEOUT"])
            )
            response.raise_for_status()
            return pd.read_csv(StringIO(response.text), skiprows=skiprows, **kwargs)

        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 400 and attempt < max_retries - 1:
                # Google Sheets rate limiting - attendre et réessayer
                time.sleep(0.5 * (attempt + 1))
                continue
            elif attempt == max_retries - 1:
                # Fallback: utiliser pd.read_csv direct (plus lent mais plus fiable)
                try:
                    return pd.read_csv(url, skiprows=skiprows, **kwargs)
                except Exception:
                    pass
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(0.3)
                continue

    # Dernier recours: pd.read_csv direct
    try:
        return pd.read_csv(url, skiprows=skiprows, **kwargs)
    except Exception as e:
        print(f"⚠️ Erreur chargement {url[:60]}...: {e}")
        return pd.DataFrame()


def load_csvs_parallel(url_configs):
    """
    Charge plusieurs CSV en parallèle avec gestion du rate limiting Google Sheets.
    Utilise un nombre limité de workers pour éviter les blocages.
    """
    results = {}

    if not PERFORMANCE_CONFIG["PARALLEL_LOADING"]:
        # Fallback séquentiel
        for config in url_configs:
            cfg = config.copy()
            name = cfg.pop("name")
            url = cfg.pop("url")
            results[name] = load_csv_fast(url, **cfg)
        return results

    from concurrent.futures import ThreadPoolExecutor, as_completed

    def load_one(config):
        cfg = config.copy()
        name = cfg.pop("name")
        url = cfg.pop("url")
        # Petit délai aléatoire pour éviter les requêtes simultanées exactes
        time.sleep(0.1 * hash(name) % 10 / 10)
        return name, load_csv_fast(url, **cfg)

    # Limiter à 4 workers pour Google Sheets (évite rate limiting)
    max_workers = min(PERFORMANCE_CONFIG["MAX_WORKERS"], 4)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(load_one, cfg): cfg["name"] for cfg in url_configs}

        for future in as_completed(futures):
            try:
                name, df = future.result()
                results[name] = df
                if not df.empty:
                    print(f"   ✅ {name}: {len(df)} lignes")
            except Exception as e:
                name = futures[future]
                print(f"   ❌ {name}: {e}")
                results[name] = pd.DataFrame()

    return results


# ============================================================
# 🔄 CACHE INTELLIGENT AVEC REFRESH EN ARRIÈRE-PLAN
# ============================================================

def get_cached_data(cache_key, loader_func, ttl_seconds=None):
    """
    Récupère les données du cache ou les charge.
    Refresh en arrière-plan si données proches de l'expiration.
    """
    if ttl_seconds is None:
        ttl_seconds = PERFORMANCE_CONFIG["CACHE_TTL_SECONDS"]

    current_time = time.time()

    with _GLOBAL_CACHE_LOCK:
        cached = _GLOBAL_DATA_CACHE.get(cache_key)

        if cached:
            age = current_time - cached["timestamp"]

            # Données encore valides
            if age < ttl_seconds:
                # Si proche de l'expiration (>80%), refresh en arrière-plan
                if age > ttl_seconds * 0.8 and not cached.get("refreshing") and PERFORMANCE_CONFIG["BACKGROUND_REFRESH"]:
                    cached["refreshing"] = True
                    _async_executor.submit(_background_refresh, cache_key, loader_func)

                return cached["data"]

    # Pas en cache ou expiré - charger
    print(f"🔄 Chargement {cache_key}...")
    start = time.time()
    data = loader_func()
    elapsed = time.time() - start
    print(f"✅ {cache_key} chargé en {elapsed:.2f}s")

    with _GLOBAL_CACHE_LOCK:
        _GLOBAL_DATA_CACHE[cache_key] = {
            "data": data,
            "timestamp": current_time,
            "refreshing": False
        }

    return data


def _background_refresh(cache_key, loader_func):
    """Refresh le cache en arrière-plan"""
    try:
        data = loader_func()
        with _GLOBAL_CACHE_LOCK:
            _GLOBAL_DATA_CACHE[cache_key] = {
                "data": data,
                "timestamp": time.time(),
                "refreshing": False
            }
        print(f"🔄 Cache {cache_key} refreshed en arrière-plan")
    except Exception as e:
        print(f"⚠️ Background refresh failed for {cache_key}: {e}")
        with _GLOBAL_CACHE_LOCK:
            if cache_key in _GLOBAL_DATA_CACHE:
                _GLOBAL_DATA_CACHE[cache_key]["refreshing"] = False


def invalidate_cache(cache_key=None):
    """Invalide le cache (tout ou une clé spécifique)"""
    with _GLOBAL_CACHE_LOCK:
        if cache_key:
            _GLOBAL_DATA_CACHE.pop(cache_key, None)
        else:
            _GLOBAL_DATA_CACHE.clear()
    print(f"🧹 Cache invalidé: {cache_key or 'ALL'}")

# ============================================================
# 📦 CACHE GLOBAL POUR LES DONNÉES DE RÉCEPTION (ASYNC)
# ============================================================
_receptions_cache = {
    "data": None,
    "timestamp": 0,
    "loading": False,
    "error": None
}
_receptions_lock = threading.Lock()

# ThreadPool pour chargements asynchrones
_async_executor = ThreadPoolExecutor(max_workers=3)


def load_receptions_async(url: str, timeout: int = 8):
    """
    Charge les données de réception de manière asynchrone avec timeout.
    Ne bloque jamais le thread principal.
    """
    global _receptions_cache

    with _receptions_lock:
        # Vérifier si le cache est valide (10 minutes)
        if _receptions_cache["data"] is not None:
            age = time.time() - _receptions_cache["timestamp"]
            if age < 600:  # 10 minutes
                return _receptions_cache["data"]

        # Éviter les chargements multiples simultanés
        if _receptions_cache["loading"]:
            # Retourner les anciennes données ou DataFrame vide
            return _receptions_cache["data"] if _receptions_cache["data"] is not None else pd.DataFrame()

        _receptions_cache["loading"] = True

    def _fetch_receptions():
        """Fonction interne pour charger les réceptions"""
        try:
            import requests
            from io import StringIO

            response = requests.get(url, timeout=timeout)
            if response.status_code == 200:
                df = pd.read_csv(StringIO(response.text))

                # Harmoniser les colonnes
                df.columns = (
                    df.columns.astype(str)
                    .str.strip()
                    .str.lower()
                    .str.replace(" ", "_")
                )

                # Renommer les colonnes clés
                col_renames = {}
                for col in df.columns:
                    if 'product' in col and 'name' in col:
                        col_renames[col] = 'product_name'
                    elif col in ['product', 'produit', 'nom_produit']:
                        col_renames[col] = 'product_name'
                    elif 'quantity' in col or 'qty' in col or 'qte' in col:
                        col_renames[col] = 'last_reception_qty'
                    elif 'date' in col and 'reception' not in col_renames.values():
                        col_renames[col] = 'last_reception_date'

                if col_renames:
                    df.rename(columns=col_renames, inplace=True)

                return df
            else:
                print(f"⚠️ Réceptions HTTP {response.status_code}")
                return pd.DataFrame()

        except Exception as e:
            print(f"⚠️ Erreur chargement réceptions: {e}")
            return pd.DataFrame()

    try:
        # Soumettre la tâche au ThreadPool avec timeout
        future = _async_executor.submit(_fetch_receptions)
        result = future.result(timeout=timeout + 2)  # Timeout légèrement supérieur

        with _receptions_lock:
            if result is not None and not result.empty:
                _receptions_cache["data"] = result
                _receptions_cache["timestamp"] = time.time()
                _receptions_cache["error"] = None
                print(f"✅ Réceptions chargées: {len(result)} lignes")
            _receptions_cache["loading"] = False

        return result

    except FuturesTimeoutError:
        print(f"⏱️ Timeout chargement réceptions ({timeout}s)")
        with _receptions_lock:
            _receptions_cache["loading"] = False
            _receptions_cache["error"] = "timeout"
        return _receptions_cache["data"] if _receptions_cache["data"] is not None else pd.DataFrame()

    except Exception as e:
        print(f"⚠️ Erreur async réceptions: {e}")
        with _receptions_lock:
            _receptions_cache["loading"] = False
            _receptions_cache["error"] = str(e)
        return pd.DataFrame()


def get_cached_receptions():
    """Retourne les réceptions en cache (sans bloquer)"""
    with _receptions_lock:
        if _receptions_cache["data"] is not None:
            return _receptions_cache["data"].copy()
    return pd.DataFrame(columns=['product_name', 'last_reception_qty', 'last_reception_date'])


def preload_receptions_background(url: str):
    """Lance le chargement des réceptions en arrière-plan (non-bloquant)"""
    def _bg_load():
        load_receptions_async(url, timeout=15)

    _async_executor.submit(_bg_load)
    print("🔄 Chargement réceptions lancé en arrière-plan...")


def get_user_data(username: str = None):
    """Récupère les données pour un utilisateur spécifique (thread-safe)"""
    cache_key = username or "global"

    with _user_cache_lock:
        if cache_key in _user_data_cache:
            cached = _user_data_cache[cache_key]
            # Vérifier si le cache est encore valide (5 minutes)
            if time.time() - cached.get("timestamp", 0) < 300:
                return cached.get("data")

    # Charger les données fraîches
    df = load_supply_data()

    with _user_cache_lock:
        _user_data_cache[cache_key] = {
            "data": df,
            "timestamp": time.time()
        }

    return df


def invalidate_user_cache(username: str = None):
    """Invalide le cache pour un utilisateur"""
    cache_key = username or "global"
    with _user_cache_lock:
        if cache_key in _user_data_cache:
            del _user_data_cache[cache_key]


def get_df_cached():
    """
    🚀 VERSION OPTIMISÉE - Utilise le cache intelligent avec refresh en arrière-plan
    """
    return get_cached_data(
        cache_key="supply_data_main",
        loader_func=load_supply_data,
        ttl_seconds=PERFORMANCE_CONFIG["CACHE_TTL_SECONDS"]
    )


def clear_all_caches():
    """Nettoie tous les caches"""
    global _user_data_cache, _GLOBAL_DATA_CACHE

    # Cache utilisateur
    with _user_cache_lock:
        _user_data_cache = {}

    # Cache global
    invalidate_cache()

    # Cache Flask
    try:
        cache.clear()
    except:
        pass

    print("🧹 Tous les caches nettoyés")

RENDER_ENV = os.getenv("RENDER", False)
DEBUG_MODE = os.getenv("DEBUG", "False").lower() == "true"

# ============================================================
# 🗄️ CONFIGURATION SUPABASE
# ============================================================
SUPABASE_URL = os.getenv("SUPABASE_URL", "https://tvaxxzkilrsgfqjswtkt.supabase.co")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InR2YXh4emtpbHJzZ2ZxanN3dGt0Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjQ3Mzc1NzksImV4cCI6MjA4MDMxMzU3OX0.uIbLmOm2z9YlD5bxh8nXMY2JjqJr8Qj12hmBc7hBPy0")

# Initialiser le client Supabase
supabase_client = None
if SUPABASE_AVAILABLE:
    try:
        supabase_client = create_client(SUPABASE_URL, SUPABASE_KEY)
        print("✅ Supabase connecté avec succès")
    except Exception as e:
        print(f"⚠️ Erreur connexion Supabase: {e}")
        supabase_client = None
else:
    print("⚠️ Supabase non disponible - tracking désactivé")

# ============================================================
# 📝 SQL SUPABASE - TABLES REQUISES
# ============================================================
# Exécuter ce SQL dans Supabase pour créer les tables nécessaires:
#
# -- Table de tracking des commandes agents
# CREATE TABLE IF NOT EXISTS agent_orders (
#     id SERIAL PRIMARY KEY,
#     agent_username TEXT NOT NULL,
#     supplier TEXT NOT NULL,
#     order_total DECIMAL(15,2) DEFAULT 0,
#     products_count INTEGER DEFAULT 0,
#     po_number TEXT,
#     created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
# );
#
# -- Index pour les requêtes par agent
# CREATE INDEX idx_agent_orders_username ON agent_orders(agent_username);
# CREATE INDEX idx_agent_orders_created ON agent_orders(created_at);
#
# -- Table des notes produits (si pas encore créée)
# CREATE TABLE IF NOT EXISTS product_notes (
#     id SERIAL PRIMARY KEY,
#     product_name TEXT NOT NULL,
#     user_id TEXT,
#     note_text TEXT,
#     mentions JSONB DEFAULT '[]',
#     is_deleted BOOLEAN DEFAULT FALSE,
#     created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
# );
# ============================================================


# ============================================================
# 📊 FONCTIONS DE TRACKING SUPABASE - ASYNCHRONE
# ============================================================

# Queue pour le tracking asynchrone (ne bloque pas l'UI)
_tracking_queue = Queue()
_tracking_thread = None
_tracking_running = False


def _tracking_worker():
    """Worker thread pour traiter les événements de tracking en arrière-plan"""
    global _tracking_running
    while _tracking_running:
        try:
            # Attendre un événement (timeout 1 seconde)
            event = _tracking_queue.get(timeout=1)
            if event is None:
                continue

            event_type = event.get("type")
            data = event.get("data", {})

            if event_type == "activity":
                _do_track_activity(**data)
            elif event_type == "qac_edit":
                _do_track_qac_edit(**data)
            elif event_type == "po_generated":
                _do_track_po_generated(**data)
            elif event_type == "filter_used":
                _do_track_filter_used(**data)
            elif event_type == "selection":
                _do_track_selection(**data)

            _tracking_queue.task_done()
        except Exception:
            pass  # Timeout ou erreur, on continue


def start_tracking_worker():
    """Démarre le worker de tracking"""
    global _tracking_thread, _tracking_running
    if _tracking_thread is None or not _tracking_thread.is_alive():
        _tracking_running = True
        _tracking_thread = threading.Thread(target=_tracking_worker, daemon=True)
        _tracking_thread.start()
        print("✅ Worker de tracking Supabase démarré")


def stop_tracking_worker():
    """Arrête le worker de tracking"""
    global _tracking_running
    _tracking_running = False


# Démarrer le worker au lancement
start_tracking_worker()


def get_or_create_user(username: str) -> dict:
    """Récupère ou crée un utilisateur dans Supabase"""
    if not supabase_client or not username:
        return None

    try:
        username = username.lower().strip()

        # Chercher l'utilisateur existant
        result = supabase_client.table("users").select("*").eq("username", username).execute()

        if result.data and len(result.data) > 0:
            return result.data[0]

        # Créer l'utilisateur s'il n'existe pas
        if username in AUTH_USERS:
            user_info = AUTH_USERS[username]
            new_user = {
                "username": username,
                "name": user_info.get("name", username),
                "email": user_info.get("email", ""),
                "role": user_info.get("role", "user")
            }
            result = supabase_client.table("users").insert(new_user).execute()
            if result.data:
                print(f"✅ Utilisateur {username} créé dans Supabase")
                return result.data[0]

        return None
    except Exception as e:
        print(f"⚠️ Erreur get_or_create_user: {e}")
        return None


# Flag global pour désactiver le tracking si RLS bloque
SUPABASE_TRACKING_ENABLED = True

def create_session(user_id: str, ip_address: str = None, user_agent: str = None) -> str:
    """Crée une nouvelle session utilisateur"""
    global SUPABASE_TRACKING_ENABLED

    if not supabase_client or not user_id or not SUPABASE_TRACKING_ENABLED:
        return None

    try:
        session_data = {
            "user_id": user_id,
            "ip_address": ip_address,
            "user_agent": user_agent,
            "is_active": True
        }
        result = supabase_client.table("user_sessions").insert(session_data).execute()
        if result.data:
            print(f"✅ Session créée: {result.data[0]['id']}")
            return result.data[0]["id"]
        return None
    except Exception as e:
        error_str = str(e)
        if "row-level security policy" in error_str or "42501" in error_str:
            print(f"⚠️ RLS bloque create_session - Tracking désactivé (configurer les policies Supabase)")
            SUPABASE_TRACKING_ENABLED = False
        else:
            print(f"⚠️ Erreur create_session: {e}")
        return None


def end_session(session_id: str):
    """Termine une session utilisateur"""
    global SUPABASE_TRACKING_ENABLED

    if not supabase_client or not session_id or not SUPABASE_TRACKING_ENABLED:
        return

    try:
        supabase_client.table("user_sessions").update({
            "logout_at": "now()",
            "is_active": False
        }).eq("id", session_id).execute()
        print(f"✅ Session terminée: {session_id}")
    except Exception as e:
        error_str = str(e)
        if "row-level security policy" in error_str or "42501" in error_str:
            SUPABASE_TRACKING_ENABLED = False
        else:
            print(f"⚠️ Erreur end_session: {e}")


def track_activity(user_id: str, session_id: str, action_type: str, page: str = None, details: dict = None):
    """Enregistre une activité utilisateur (asynchrone via queue)"""
    if not supabase_client or not SUPABASE_TRACKING_ENABLED:
        return

    # Ajouter à la queue pour traitement asynchrone
    _tracking_queue.put({
        "type": "activity",
        "data": {
            "user_id": user_id,
            "session_id": session_id,
            "action_type": action_type,
            "page": page,
            "details": details
        }
    })


def track_stock_stats(total_skus: int, active_skus: int, out_of_stock: int, at_risk: int, suppliers: int):
    """
    🚀 Tracking des stats de stock dans Supabase (asynchrone)
    Appelé à chaque chargement de données pour historique
    """
    if not supabase_client or not SUPABASE_TRACKING_ENABLED:
        return

    details = {
        "total_skus": total_skus,
        "active_skus": active_skus,
        "out_of_stock": out_of_stock,
        "at_risk": at_risk,
        "suppliers": suppliers,
        "timestamp": datetime.now().isoformat()
    }

    _tracking_queue.put({
        "type": "activity",
        "data": {
            "user_id": None,
            "session_id": None,
            "action_type": "stock_stats_snapshot",
            "page": "overview",
            "details": details
        }
    })
    print(f"📊 Stock stats trackées: {active_skus} SKUs actifs, {out_of_stock} ruptures")


def _do_track_activity(user_id: str, session_id: str, action_type: str, page: str = None, details: dict = None):
    """Exécution réelle du tracking activité"""
    global SUPABASE_TRACKING_ENABLED

    if not supabase_client or not SUPABASE_TRACKING_ENABLED:
        return

    try:
        activity_data = {
            "user_id": user_id,
            "session_id": session_id,
            "action_type": action_type,
            "page": page,
            "action_details": details or {}
        }
        supabase_client.table("user_activities").insert(activity_data).execute()
        print(f"📊 Activité trackée: {action_type} sur {page}")
    except Exception as e:
        error_str = str(e)
        if "row-level security policy" in error_str or "42501" in error_str:
            SUPABASE_TRACKING_ENABLED = False
        else:
            print(f"⚠️ Erreur track_activity: {e}")


def track_qac_edit(user_id: str, session_id: str, product_name: str, old_value: int, new_value: int, supplier: str = None):
    """Enregistre une modification QAC (asynchrone)"""
    if not supabase_client or not SUPABASE_TRACKING_ENABLED:
        return

    _tracking_queue.put({
        "type": "qac_edit",
        "data": {
            "user_id": user_id,
            "session_id": session_id,
            "product_name": product_name,
            "old_value": old_value,
            "new_value": new_value,
            "supplier": supplier
        }
    })


def _do_track_qac_edit(user_id: str, session_id: str, product_name: str, old_value: int, new_value: int, supplier: str = None):
    """Exécution réelle du tracking QAC edit"""
    global SUPABASE_TRACKING_ENABLED

    if not supabase_client or not SUPABASE_TRACKING_ENABLED:
        return

    try:
        edit_data = {
            "user_id": user_id,
            "session_id": session_id,
            "product_name": product_name,
            "supplier": supplier,
            "old_qac_value": old_value,
            "new_qac_value": new_value
        }
        supabase_client.table("qac_edits_history").insert(edit_data).execute()
        print(f"📝 QAC edit tracké: {product_name} ({old_value} → {new_value})")
    except Exception as e:
        error_str = str(e)
        if "row-level security policy" in error_str or "42501" in error_str:
            SUPABASE_TRACKING_ENABLED = False
        else:
            print(f"⚠️ Erreur track_qac_edit: {e}")


# ============================================================
# 📊 NOUVELLES FONCTIONS DE TRACKING - ACTIONS AGENTS
# ============================================================

def track_po_generated(user_id: str, session_id: str, supplier: str, products_count: int, total_amount: float, po_type: str = "excel"):
    """Track la génération d'un bon de commande"""
    if not supabase_client or not SUPABASE_TRACKING_ENABLED:
        return

    _tracking_queue.put({
        "type": "po_generated",
        "data": {
            "user_id": user_id,
            "session_id": session_id,
            "supplier": supplier,
            "products_count": products_count,
            "total_amount": total_amount,
            "po_type": po_type
        }
    })


def _do_track_po_generated(user_id: str, session_id: str, supplier: str, products_count: int, total_amount: float, po_type: str = "excel"):
    """Exécution réelle du tracking PO"""
    global SUPABASE_TRACKING_ENABLED

    if not supabase_client or not SUPABASE_TRACKING_ENABLED:
        return

    try:
        # Tracker comme activité
        activity_data = {
            "user_id": user_id,
            "session_id": session_id,
            "action_type": "po_generated",
            "page": "overview",
            "action_details": {
                "supplier": supplier,
                "products_count": products_count,
                "total_amount": total_amount,
                "po_type": po_type
            }
        }
        supabase_client.table("user_activities").insert(activity_data).execute()
        print(f"📄 PO tracké: {supplier} - {products_count} produits - {total_amount:,.0f} FCFA")
    except Exception as e:
        error_str = str(e)
        if "row-level security policy" in error_str or "42501" in error_str:
            SUPABASE_TRACKING_ENABLED = False


def track_filter_used(user_id: str, session_id: str, filter_type: str, filter_value):
    """Track l'utilisation des filtres"""
    if not supabase_client or not SUPABASE_TRACKING_ENABLED:
        return

    _tracking_queue.put({
        "type": "filter_used",
        "data": {
            "user_id": user_id,
            "session_id": session_id,
            "filter_type": filter_type,
            "filter_value": filter_value
        }
    })


def _do_track_filter_used(user_id: str, session_id: str, filter_type: str, filter_value):
    """Exécution réelle du tracking filtre"""
    global SUPABASE_TRACKING_ENABLED

    if not supabase_client or not SUPABASE_TRACKING_ENABLED:
        return

    try:
        activity_data = {
            "user_id": user_id,
            "session_id": session_id,
            "action_type": "filter_used",
            "page": "overview",
            "action_details": {
                "filter_type": filter_type,
                "filter_value": str(filter_value) if filter_value else None
            }
        }
        supabase_client.table("user_activities").insert(activity_data).execute()
    except Exception as e:
        error_str = str(e)
        if "row-level security policy" in error_str or "42501" in error_str:
            SUPABASE_TRACKING_ENABLED = False


def track_selection(user_id: str, session_id: str, selected_count: int, action: str = "select"):
    """Track la sélection de produits"""
    if not supabase_client or not SUPABASE_TRACKING_ENABLED:
        return

    _tracking_queue.put({
        "type": "selection",
        "data": {
            "user_id": user_id,
            "session_id": session_id,
            "selected_count": selected_count,
            "action": action
        }
    })


def _do_track_selection(user_id: str, session_id: str, selected_count: int, action: str = "select"):
    """Exécution réelle du tracking sélection"""
    global SUPABASE_TRACKING_ENABLED

    if not supabase_client or not SUPABASE_TRACKING_ENABLED:
        return

    try:
        activity_data = {
            "user_id": user_id,
            "session_id": session_id,
            "action_type": f"products_{action}",
            "page": "overview",
            "action_details": {
                "selected_count": selected_count
            }
        }
        supabase_client.table("user_activities").insert(activity_data).execute()
    except Exception as e:
        error_str = str(e)
        if "row-level security policy" in error_str or "42501" in error_str:
            SUPABASE_TRACKING_ENABLED = False


def track_data_refresh(user_id: str, session_id: str, products_count: int):
    """Track le rafraîchissement des données"""
    track_activity(user_id, session_id, "data_refresh", "overview", {
        "products_count": products_count,
        "timestamp": datetime.now().isoformat()
    })


def track_export(user_id: str, session_id: str, export_type: str, records_count: int):
    """Track l'export de données"""
    track_activity(user_id, session_id, "data_export", "overview", {
        "export_type": export_type,
        "records_count": records_count
    })


# ==========================================
# 📊 DAILY STATS - Statistiques journalières
# ==========================================

def save_daily_stats(stats_data: dict):
    """Sauvegarde les statistiques journalières"""
    global SUPABASE_TRACKING_ENABLED

    if not supabase_client or not SUPABASE_TRACKING_ENABLED:
        return None

    try:
        # Ajouter la date du jour si pas présente
        if "stat_date" not in stats_data:
            stats_data["stat_date"] = datetime.now().strftime("%Y-%m-%d")

        result = supabase_client.table("daily_stats").insert(stats_data).execute()
        if result.data:
            print(f"📊 Stats journalières sauvegardées: {stats_data.get('stat_date')}")
            return result.data[0]
        return None
    except Exception as e:
        error_str = str(e)
        if "row-level security policy" in error_str or "42501" in error_str:
            SUPABASE_TRACKING_ENABLED = False
        else:
            print(f"⚠️ Erreur save_daily_stats: {e}")
        return None


def get_daily_stats(days: int = 30) -> list:
    """Récupère les statistiques des X derniers jours"""
    if not supabase_client:
        return []

    try:
        from datetime import timedelta
        start_date = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")

        result = supabase_client.table("daily_stats") \
            .select("*") \
            .gte("stat_date", start_date) \
            .order("stat_date", desc=True) \
            .execute()
        return result.data or []
    except Exception as e:
        print(f"⚠️ Erreur get_daily_stats: {e}")
        return []


def update_daily_stats_on_load(master_df):
    """Met à jour les stats journalières au chargement des données"""
    if master_df is None or master_df.empty:
        return

    try:
        today = datetime.now().strftime("%Y-%m-%d")

        # Calculer les stats
        total_skus = len(master_df)
        out_of_stock = int((master_df.get('total_stock', pd.Series([0])) <= 0).sum())
        total_suppliers = master_df.get('Supplier', pd.Series()).nunique()

        # Valeur totale stock
        if 'total_stock' in master_df.columns and 'Prix Achat' in master_df.columns:
            stock_val = master_df['total_stock'].fillna(0)
            price_val = pd.to_numeric(master_df['Prix Achat'], errors='coerce').fillna(0)
            total_stock_value = float((stock_val * price_val).sum())
        else:
            total_stock_value = 0

        stats = {
            "stat_date": today,
            "total_skus": total_skus,
            "out_of_stock_count": out_of_stock,
            "total_suppliers": total_suppliers,
            "total_stock_value": total_stock_value,
            "orders_generated": 0,
            "qac_edits_count": 0
        }

        save_daily_stats(stats)

    except Exception as e:
        print(f"⚠️ Erreur update_daily_stats_on_load: {e}")


# ==========================================
# 📋 RÉCUPÉRATION DE TOUTES LES DONNÉES
# ==========================================

def get_all_user_sessions(limit: int = 100) -> list:
    """Récupère toutes les sessions utilisateur"""
    if not supabase_client:
        return []

    try:
        result = supabase_client.table("user_sessions") \
            .select("*, users(username, name)") \
            .order("created_at", desc=True) \
            .limit(limit) \
            .execute()
        return result.data or []
    except Exception as e:
        print(f"⚠️ Erreur get_all_user_sessions: {e}")
        return []


def get_all_user_activities(limit: int = 500) -> list:
    """Récupère toutes les activités utilisateur"""
    if not supabase_client:
        return []

    try:
        result = supabase_client.table("user_activities") \
            .select("*, users(username, name)") \
            .order("created_at", desc=True) \
            .limit(limit) \
            .execute()
        return result.data or []
    except Exception as e:
        print(f"⚠️ Erreur get_all_user_activities: {e}")
        return []


def get_all_qac_edits(limit: int = 500) -> list:
    """Récupère tout l'historique des modifications QAC"""
    if not supabase_client:
        return []

    try:
        result = supabase_client.table("qac_edits_history") \
            .select("*, users(username, name)") \
            .order("created_at", desc=True) \
            .limit(limit) \
            .execute()
        return result.data or []
    except Exception as e:
        print(f"⚠️ Erreur get_all_qac_edits: {e}")
        return []


def get_all_product_notes(limit: int = 500) -> list:
    """Récupère toutes les notes produits"""
    if not supabase_client:
        return []

    try:
        result = supabase_client.table("product_notes") \
            .select("*, users(username, name)") \
            .eq("is_deleted", False) \
            .order("created_at", desc=True) \
            .limit(limit) \
            .execute()
        return result.data or []
    except Exception as e:
        print(f"⚠️ Erreur get_all_product_notes: {e}")
        return []


def get_all_users() -> list:
    """Récupère tous les utilisateurs"""
    if not supabase_client:
        return []

    try:
        result = supabase_client.table("users") \
            .select("id, username, name, email, role, is_active, created_at, last_login") \
            .order("created_at", desc=True) \
            .execute()
        return result.data or []
    except Exception as e:
        print(f"⚠️ Erreur get_all_users: {e}")
        return []


def get_dashboard_analytics() -> dict:
    """Récupère un résumé analytique pour le dashboard admin"""
    if not supabase_client:
        return {}

    try:
        analytics = {
            "total_users": 0,
            "active_sessions": 0,
            "total_activities": 0,
            "total_qac_edits": 0,
            "total_notes": 0,
            "recent_stats": []
        }

        # Compter les utilisateurs
        users = supabase_client.table("users").select("id", count="exact").execute()
        analytics["total_users"] = users.count if hasattr(users, 'count') else len(users.data or [])

        # Sessions actives
        sessions = supabase_client.table("user_sessions") \
            .select("id", count="exact") \
            .eq("is_active", True) \
            .execute()
        analytics["active_sessions"] = sessions.count if hasattr(sessions, 'count') else len(sessions.data or [])

        # Total activités (dernier mois)
        from datetime import timedelta
        month_ago = (datetime.now() - timedelta(days=30)).isoformat()
        activities = supabase_client.table("user_activities") \
            .select("id", count="exact") \
            .gte("created_at", month_ago) \
            .execute()
        analytics["total_activities"] = activities.count if hasattr(activities, 'count') else len(activities.data or [])

        # Total modifications QAC
        qac = supabase_client.table("qac_edits_history") \
            .select("id", count="exact") \
            .execute()
        analytics["total_qac_edits"] = qac.count if hasattr(qac, 'count') else len(qac.data or [])

        # Total notes
        notes = supabase_client.table("product_notes") \
            .select("id", count="exact") \
            .eq("is_deleted", False) \
            .execute()
        analytics["total_notes"] = notes.count if hasattr(notes, 'count') else len(notes.data or [])

        # Stats récentes
        analytics["recent_stats"] = get_daily_stats(7)

        print(f"📊 Analytics dashboard: {analytics['total_users']} users, {analytics['active_sessions']} sessions actives")
        return analytics

    except Exception as e:
        print(f"⚠️ Erreur get_dashboard_analytics: {e}")
        return {}


def save_product_note(user_id: str, product_name: str, note_text: str, mentions: list = None, supplier: str = None):
    """Sauvegarde une note produit"""
    if not supabase_client:
        return None

    try:
        note_data = {
            "user_id": user_id,
            "product_name": product_name,
            "supplier": supplier,
            "note_text": note_text,
            "mentions": mentions or []
        }
        result = supabase_client.table("product_notes").insert(note_data).execute()
        if result.data:
            print(f"📝 Note sauvegardée pour: {product_name}")
            return result.data[0]
        return None
    except Exception as e:
        print(f"⚠️ Erreur save_product_note: {e}")
        return None


def save_added_product(user_id: str, session_id: str, product_data: dict):
    """
    Sauvegarde un produit ajouté manuellement dans Supabase.
    Table: added_products
    """
    global SUPABASE_TRACKING_ENABLED

    if not supabase_client or not SUPABASE_TRACKING_ENABLED:
        return None

    try:
        record = {
            "user_id": user_id,
            "session_id": session_id,
            "product_name": product_data.get("product_name", ""),
            "supplier": product_data.get("supplier", ""),
            "category": product_data.get("category", ""),
            "initial_stock": product_data.get("stock", 0),
            "added_by": product_data.get("added_by", ""),
            "metadata": {
                "source": "dashboard_manual_add",
                "timestamp": datetime.now().isoformat()
            }
        }
        result = supabase_client.table("added_products").insert(record).execute()
        if result.data:
            print(f"✅ Produit ajouté sauvegardé dans Supabase: {product_data.get('product_name')}")
            return result.data[0]
        return None
    except Exception as e:
        error_str = str(e)
        if "relation" in error_str and "does not exist" in error_str:
            print(f"ℹ️ Table 'added_products' n'existe pas - création recommandée")
        elif "row-level security policy" in error_str or "42501" in error_str:
            SUPABASE_TRACKING_ENABLED = False
        else:
            print(f"⚠️ Erreur save_added_product: {e}")
        return None


def track_product_action(user_id: str, session_id: str, action: str, product_name: str, details: dict = None):
    """
    Track une action sur un produit (ajout, modification, suppression)
    """
    track_activity(user_id, session_id, f"product_{action}", "overview", {
        "product_name": product_name,
        **(details or {})
    })


def get_added_products(limit: int = 100) -> list:
    """Récupère les produits ajoutés manuellement"""
    if not supabase_client:
        return []

    try:
        result = supabase_client.table("added_products") \
            .select("*, users(username, name)") \
            .order("created_at", desc=True) \
            .limit(limit) \
            .execute()
        return result.data or []
    except Exception as e:
        print(f"⚠️ Erreur get_added_products: {e}")
        return []


def get_product_notes(product_name: str) -> list:
    """Récupère les notes d'un produit"""
    if not supabase_client:
        return []

    try:
        result = supabase_client.table("product_notes") \
            .select("*, users(username, name)") \
            .eq("product_name", product_name) \
            .eq("is_deleted", False) \
            .order("created_at", desc=True) \
            .execute()
        return result.data or []
    except Exception as e:
        print(f"⚠️ Erreur get_product_notes: {e}")
        return []


# Variable globale pour stocker la session active
ACTIVE_SESSIONS = {}  # {username: {"user_id": ..., "session_id": ...}}


if RENDER_ENV:
    print("\n" + "="*60)
    print("🚀 DÉMARRAGE SUR RENDER")
    print("="*60)
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Files: {os.listdir('.')[:10]}")
    print("="*60 + "\n")


# ==========================================
# ✅ TIMEOUT COMPATIBLE WINDOWS
# ==========================================
class CustomTimeoutError(Exception):
    """Exception levée en cas de timeout"""
    pass


def run_with_timeout(func, args=(), kwargs=None, timeout_seconds=30):
    """
    Exécute une fonction avec timeout (compatible Windows)

    Args:
        func: Fonction à exécuter
        args: Arguments positionnels
        kwargs: Arguments nommés
        timeout_seconds: Timeout en secondes

    Returns:
        Résultat de la fonction

    Raises:
        TimeoutError: Si timeout dépassé
    """
    if kwargs is None:
        kwargs = {}

    result = [None]
    exception = [None]

    def target():
        try:
            result[0] = func(*args, **kwargs)
        except Exception as e:
            exception[0] = e

    thread = threading.Thread(target=target)
    thread.daemon = True
    thread.start()
    thread.join(timeout_seconds)

    if thread.is_alive():
        # Thread encore actif = timeout
        raise CustomTimeoutError(f"Opération timeout après {timeout_seconds} secondes")

    if exception[0]:
        raise exception[0]

    return result[0]
# ------------- OpenAI client (clé hardcodée à ta demande) ----------------
# OPENAI_API_KEY_HARDCODED = "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"
# Ligne ~45 dans votre code
# api_key = os.getenv("OPENAI_API_KEY_HARDCODED")

# AJOUTER CES LIGNES DE DEBUG
# print("=" * 60)
# print("DEBUG OPENAI CLIENT")
# print("=" * 60)
# if api_key:
#   print(f"✅ API Key trouvée : {api_key[:15]}...{api_key[-4:]}")  # Masquer le milieu
# else:
#   print("❌ API Key NON trouvée dans l'environnement")
# print("=" * 60)


# openai_client = None
# try:
#   if api_key:
#      from openai import OpenAI
#     openai_client = OpenAI(api_key=api_key)
#    print("✅ Client OpenAI initialisé avec succès")
# else:
#   print("⚠️ Pas de clé API, chatbot utilisera fallback")
# except Exception as e:
#    print(f"❌ Erreur initialisation OpenAI : {type(e).__name__}: {e}")
#   openai_client = None
# print("=" * 60)


# ====== GEMINI: config + client ======

from google.api_core.exceptions import ResourceExhausted, DeadlineExceeded, InternalServerError
import time, json
from typing import Optional, Dict, Any, List

# Modèles Gemini
GEMINI_FLASH = "gemini-2.5-flash-lite"
GEMINI_PRO = "gemini-1.5-pro"
# Charger les variables d'environnement depuis le fichier .env
load_dotenv()


def configure_gemini():
    """
    Configure l’API Gemini depuis la variable d'env GEMINI_API_KEY.
    """
    api_key = os.getenv("GEMINI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY n'est pas défini dans l'environnement.")
    genai.configure(api_key=api_key)
    print("✅ Gemini configuré")


# Appelle la configuration
configure_gemini()

# Client Gemini minimal & robuste
DEFAULT_SAFETY = None


class GeminiClient:
    def __init__(self, model: str, system_instruction: Optional[str] = None,
                 max_output_tokens: int = 2048, temperature: float = 0.2):
        self.model = genai.GenerativeModel(
            model_name=model,
            system_instruction=system_instruction or "Tu es un assistant supply chain francophone, précis et concis."
        )
        self.gen_kwargs = dict(
            generation_config={
                "temperature": temperature,
                "max_output_tokens": max_output_tokens,
            },
            safety_settings=DEFAULT_SAFETY
        )

    def generate(self, prompt: str, stream: bool = False) -> str:
        tries, backoff = 0, 1.0
        while True:
            try:
                if stream:
                    out = []
                    for ev in self.model.generate_content([prompt], stream=True, **self.gen_kwargs):
                        if getattr(ev, "text", None):
                            out.append(ev.text)
                    return "".join(out).strip()
                else:
                    resp = self.model.generate_content([prompt], **self.gen_kwargs)
                    return (resp.text or "").strip()
            except (ResourceExhausted, DeadlineExceeded, InternalServerError):
                tries += 1
                if tries >= 3:
                    raise
                time.sleep(backoff);
                backoff *= 2

    def generate_json(self, prompt: str, schema: Dict[str, Any], strict: bool = True) -> Dict[str, Any]:
        gen_kwargs = dict(self.gen_kwargs)
        gen_kwargs["generation_config"] = {
            **gen_kwargs.get("generation_config", {}),
            "response_mime_type": "application/json",
            "response_schema": schema,
        }
        resp = self.model.generate_content([prompt], **gen_kwargs)
        raw = resp.text or "{}"
        try:
            return json.loads(raw)
        except Exception:
            if strict: raise
            return {}

    def embed(self, texts: List[str]) -> List[List[float]]:
        emb = genai.embed_content(model="text-embedding-004", content=texts)
        if isinstance(texts, list) and len(texts) > 1:
            return [r["values"] for r in emb["embedding"]]
        return emb["embedding"]["values"]


# Instancier le client (choisis le modèle)
gemini_client = GeminiClient(
    model=GEMINI_FLASH,
    system_instruction=("Tu es un assistant supply chain francophone. Sois précis, concis et pragmatique."
                        "Tu es maad_assistante, expert Supply Chain avec 30 ans d'expérience. "
                        "Tu es conversationnel et réponds naturellement aux questions.\n\n"
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
                        ),
    max_output_tokens=1024,
    temperature=0.3
)

# ----------------------------- Brand & Meta --------------------------------------
APP_TITLE = "Supply Chain Command Center"
APP_BRAND = "Maad SaSu"
AUTHOR = "Tony SARRE"
THEME = dbc.themes.CYBORG  # sobre & premium

# ------------------------------ Company / Branding -------------------------------
COMPANY_NAME = "Maad SaSu"
COMPANY_CAPITAL = os.getenv("COMPANY_CAPITAL", "")
COMPANY_RCS = os.getenv("COMPANY_RCS", "")
COMPANY_ADDRESS = os.getenv("COMPANY_ADDRESS", "")
COMPANY_PHONE = os.getenv("COMPANY_PHONE", "")
COMPANY_EMAIL = os.getenv("COMPANY_EMAIL", "")
DEFAULT_TVA_RATE = float(os.getenv("COMPANY_TVA_RATE", "0.18"))  # 18% par défaut


def send_notification_email(recipient_email, recipient_name, product_name, author, message):
    """
    Envoie un email de notification à un membre de l'équipe
    """
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from datetime import datetime

    print(f"\n{'=' * 60}")
    print(f"📧 ENVOI EMAIL À {recipient_name}")
    print(f"{'=' * 60}")

    # Configuration SMTP
    smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER", "")
    smtp_password = os.getenv("SMTP_PASSWORD", "")

    print(f"📨 SMTP: {smtp_server}:{smtp_port}")
    print(f"📨 From: {smtp_user}")
    print(f"📨 To: {recipient_email}")
    print(f"📨 Password: {'✅ OK' if smtp_password else '❌ Manquant'}")

    # Vérification
    if not smtp_user or not smtp_password:
        print("❌ SMTP_USER ou SMTP_PASSWORD non défini")
        print("   Créez un fichier .env avec :")
        print("   SMTP_USER=votre.email@gmail.com")
        print("   SMTP_PASSWORD=votre_mot_de_passe_application")
        return False

    try:
        # Créer le message
        msg = MIMEMultipart("alternative")
        msg["From"] = smtp_user
        msg["To"] = recipient_email
        msg["Subject"] = f"📝 {author} vous a mentionné sur {product_name}"

        # Corps HTML
        html_body = f"""
        <html>
        <head>
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    line-height: 1.6;
                    color: #333;
                    background: #f9f9f9;
                    padding: 20px;
                }}
                .container {{
                    max-width: 600px;
                    margin: 0 auto;
                    background: white;
                    border-radius: 12px;
                    overflow: hidden;
                    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
                }}
                .header {{
                    background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%);
                    color: white;
                    padding: 24px;
                    text-align: center;
                }}
                .header h2 {{
                    margin: 0;
                    font-size: 24px;
                }}
                .content {{
                    padding: 24px;
                }}
                .mention-badge {{
                    display: inline-block;
                    background: #0ea5e9;
                    color: #001018;
                    padding: 4px 12px;
                    border-radius: 20px;
                    font-weight: 700;
                    font-size: 14px;
                    margin: 8px 0;
                }}
                .product {{
                    font-size: 18px;
                    font-weight: 700;
                    color: #1e40af;
                    margin: 16px 0;
                    padding: 12px;
                    background: #f0f4f8;
                    border-left: 4px solid #0ea5e9;
                    border-radius: 4px;
                }}
                .note {{
                    background: #f9fafb;
                    padding: 16px;
                    border-radius: 8px;
                    margin: 16px 0;
                    border: 1px solid #e5e7eb;
                }}
                .note p {{
                    margin: 0;
                    color: #374151;
                    line-height: 1.6;
                }}
                .author {{
                    color: #6b7280;
                    font-size: 14px;
                    margin-top: 16px;
                    padding-top: 16px;
                    border-top: 1px solid #e5e7eb;
                }}
                .footer {{
                    text-align: center;
                    padding: 20px;
                    background: #f3f4f6;
                    color: #6b7280;
                    font-size: 12px;
                }}
                .button {{
                    display: inline-block;
                    background: #0ea5e9;
                    color: #001018;
                    padding: 12px 24px;
                    border-radius: 8px;
                    text-decoration: none;
                    font-weight: 700;
                    margin: 16px 0;
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h2>💬 Nouvelle mention</h2>
                </div>

                <div class="content">
                    <p>Bonjour <strong>{recipient_name}</strong>,</p>

                    <p style="margin: 16px 0;">
                        <strong>{author}</strong> vous a mentionné dans une note concernant :
                    </p>

                    <div class="product">
                        🏷️ {product_name}
                    </div>

                    <div class="note">
                        <p>{message.replace('@' + recipient_name.split()[0].lower(), f'<span class="mention-badge">@{recipient_name.split()[0]}</span>')}</p>
                    </div>

                    <div class="author">
                        <strong>✍️ Auteur :</strong> {author}<br>
                        <strong>📅 Date :</strong> {datetime.now().strftime('%d/%m/%Y à %H:%M')}
                    </div>
                </div>

                <div class="footer">
                    <p><strong>MAAD SAS</strong> - Supply Chain Command Center</p>
                    <p>92 Neto Foire Azur • Tél: 221 77 875 20 20</p>
                    <p style="margin-top: 12px; font-size: 11px;">
                        Cet email a été envoyé automatiquement. Ne pas répondre.
                    </p>
                </div>
            </div>
        </body>
        </html>
        """

        # Version texte
        text_body = f"""
Bonjour {recipient_name},

{author} vous a mentionné dans une note concernant :

📦 Produit : {product_name}

Note :
{message}

---
Auteur : {author}
Date : {datetime.now().strftime('%d/%m/%Y à %H:%M')}

---
MAAD SAS - Supply Chain Command Center
92 Neto Foire Azur
Tél: 221 77 875 20 20

Cet email a été envoyé automatiquement.
        """

        # Attacher les deux versions
        msg.attach(MIMEText(text_body, "plain", "utf-8"))
        msg.attach(MIMEText(html_body, "html", "utf-8"))

        print("✅ Message créé")

        # Connexion SMTP
        print(f"🔌 Connexion à {smtp_server}...")
        server = smtplib.SMTP(smtp_server, smtp_port, timeout=10)
        server.ehlo()
        server.starttls()
        server.ehlo()

        print("🔐 Authentification...")
        server.login(smtp_user, smtp_password)

        print("📤 Envoi...")
        server.sendmail(smtp_user, [recipient_email], msg.as_string())
        server.quit()

        print(f"✅ Email envoyé à {recipient_name} !")
        print(f"{'=' * 60}\n")

        return True

    except smtplib.SMTPAuthenticationError as e:
        print(f"❌ ERREUR AUTHENTIFICATION : {e}")
        print("   → Vérifiez SMTP_USER et SMTP_PASSWORD")
        print("   → Créez un mot de passe d'application : https://myaccount.google.com/apppasswords")
        return False

    except smtplib.SMTPException as e:
        print(f"❌ ERREUR SMTP : {e}")
        return False

    except Exception as e:
        print(f"❌ ERREUR : {e}")
        import traceback
        traceback.print_exc()
        return False

# ==================== VÉRIFIER LA CONFIGURATION EMAIL ====================
# Vers ligne 200-250, vérifie que ces variables existent :

SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "")  # Ton email Gmail
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")  # Mot de passe d'application Gmail

TEAM_MEMBERS = {
    "tony": {"name": "Tony SARRE", "email": "tony.sarre@maad.io"},
    "samuel": {"name": "Samuel Essodeke", "email": "essodeke@maad.io"},
    "maimouna": {"name": "Maimouna Dagois", "email": "maimouna@maad.io"},
    "seydouna": {"name": "Seydouna Oumar Niang", "email": "seydouna@maad.io"},
    "arame": {"name": "Arame Toure", "email": "arame.toure@maad.io"},
    "coumba": {"name": "Coumba Cisse", "email": "ndeyecoumba.cisse@maad.io"},
    "fallou": {"name": "Fallou Diop", "email": "serigne.diop@maad.io"},
    "ravane": {"name": "Ravane Diop", "email": "pr.diop@maad.io"},
    "insa": {"name": "Insa Niang", "email": "insa.niang@maad.io"},
}


# ============================================================
# 👤 SYSTÈME DE GESTION DES AGENTS - SIMPLE ET PERFORMANT
# ============================================================
# AUCUN CSV externe - Utilise UNIQUEMENT initial_df

MANAGER_ROLES = ["admin", "manager"]

# ============================================================
# 📊 TRACKING SUPABASE - COMPATIBLE AVEC SCHÉMA EXISTANT
# ============================================================
# Tables utilisées: users, user_sessions, user_activities, qac_edits_history

# Cache des user_id pour éviter les requêtes répétées
_user_id_cache = {}

# Sessions actives en mémoire
ACTIVE_SESSIONS = {}


def get_user_id_from_db(username: str) -> str:
    """Récupère l'UUID de l'utilisateur depuis Supabase"""
    if not supabase_client:
        return None

    username_lower = username.lower().strip()

    # Vérifier le cache
    if username_lower in _user_id_cache:
        return _user_id_cache[username_lower]

    try:
        result = supabase_client.table("users").select("id").eq("username", username_lower).execute()
        if result.data and len(result.data) > 0:
            user_id = result.data[0]["id"]
            _user_id_cache[username_lower] = user_id
            return user_id
    except Exception as e:
        print(f"⚠️ Erreur get_user_id: {e}")

    return None


def create_session(username: str, ip_address: str = None, user_agent: str = None) -> str:
    """Crée une nouvelle session dans Supabase et retourne son ID"""
    if not supabase_client:
        return None

    user_id = get_user_id_from_db(username)
    if not user_id:
        return None

    try:
        data = {
            "user_id": user_id,
            "ip_address": ip_address,
            "user_agent": user_agent,
            "is_active": True
        }
        result = supabase_client.table("user_sessions").insert(data).execute()

        if result.data and len(result.data) > 0:
            session_id = result.data[0]["id"]

            # Stocker en mémoire
            ACTIVE_SESSIONS[username.lower()] = {
                "session_id": session_id,
                "user_id": user_id
            }

            print(f"   📊 Session créée: {session_id[:8]}...")
            return session_id
    except Exception as e:
        print(f"⚠️ Erreur create_session: {e}")

    return None


def end_session(session_id: str):
    """Termine une session"""
    if not supabase_client or not session_id:
        return

    try:
        supabase_client.table("user_sessions").update({
            "logout_at": datetime.now().isoformat(),
            "is_active": False
        }).eq("id", session_id).execute()
        print(f"   📊 Session terminée: {session_id[:8]}...")
    except Exception as e:
        print(f"⚠️ Erreur end_session: {e}")


def track_activity(username: str, action_type: str, page: str = None, details: dict = None):
    """
    Enregistre une activité dans user_activities (schéma existant).

    action_type: 'login', 'logout', 'page_view', 'filter', 'search',
                 'export_pdf', 'export_excel', 'edit_qac', 'save_qac',
                 'add_note', 'generate_bc', 'agent_ia', 'ai_chat'
    """
    if not supabase_client:
        return False

    username_lower = username.lower().strip()
    user_id = get_user_id_from_db(username_lower)

    if not user_id:
        print(f"⚠️ Track: user_id non trouvé pour {username}")
        return False

    # Récupérer session_id si disponible
    session_id = None
    if username_lower in ACTIVE_SESSIONS:
        session_id = ACTIVE_SESSIONS[username_lower].get("session_id")

    try:
        data = {
            "user_id": user_id,
            "session_id": session_id,
            "action_type": action_type,
            "page": page,
            "action_details": details or {}
        }

        supabase_client.table("user_activities").insert(data).execute()
        print(f"   📊 TRACK: {username} → {action_type} ({page or '-'})")
        return True
    except Exception as e:
        print(f"⚠️ Erreur track_activity: {e}")
        return False


def track_login(username: str, success: bool, ip: str = None, user_agent: str = None):
    """Enregistre une connexion et crée une session"""
    if success:
        session_id = create_session(username, ip, user_agent)
        track_activity(username, "login", "login", {"success": True, "session_id": session_id})
    else:
        track_activity(username, "login", "login", {"success": False})


def track_logout(username: str):
    """Enregistre une déconnexion et termine la session"""
    username_lower = username.lower().strip()

    # Terminer la session
    if username_lower in ACTIVE_SESSIONS:
        session_id = ACTIVE_SESSIONS[username_lower].get("session_id")
        if session_id:
            end_session(session_id)
        del ACTIVE_SESSIONS[username_lower]

    track_activity(username, "logout", "logout", {})


def track_page_view(username: str, page: str):
    """Enregistre une visite de page"""
    track_activity(username, "page_view", page, {"timestamp": datetime.now().isoformat()})


def track_filter_action(username: str, filters: dict):
    """Enregistre un changement de filtre"""
    track_activity(username, "filter", "overview", {"filters": filters})


def track_bc_generated(username: str, supplier: str, products_count: int, total_amount: float, po_number: str):
    """Enregistre la génération d'un bon de commande"""
    track_activity(username, "generate_bc", "overview", {
        "supplier": supplier,
        "products_count": products_count,
        "total_amount": total_amount,
        "po_number": po_number
    })


def track_agent_ia(username: str, products_selected: int, total_qac: float):
    """Enregistre l'utilisation de l'Agent IA"""
    track_activity(username, "agent_ia", "overview", {
        "products_selected": products_selected,
        "total_qac": total_qac
    })


def track_qac_edit(username: str, product_name: str, product_id: int, supplier: str, old_value: int, new_value: int):
    """Enregistre une modification QAC dans qac_edits_history"""
    if not supabase_client:
        return

    user_id = get_user_id_from_db(username)
    session_id = ACTIVE_SESSIONS.get(username.lower(), {}).get("session_id")

    try:
        data = {
            "user_id": user_id,
            "session_id": session_id,
            "product_name": product_name,
            "product_id": product_id,
            "supplier": supplier,
            "old_qac_value": old_value,
            "new_qac_value": new_value
        }
        supabase_client.table("qac_edits_history").insert(data).execute()
        print(f"   📊 QAC Edit: {product_name} ({old_value} → {new_value})")
    except Exception as e:
        print(f"⚠️ Erreur track_qac_edit: {e}")


def get_user_stats_summary(username: str) -> dict:
    """Résumé des stats d'un utilisateur depuis Supabase"""
    if not supabase_client:
        return {"logins": 0, "bc_count": 0, "bc_total": 0, "pages_viewed": 0}

    user_id = get_user_id_from_db(username)
    if not user_id:
        return {"logins": 0, "bc_count": 0, "bc_total": 0, "pages_viewed": 0}

    try:
        # Récupérer les activités des 30 derniers jours
        result = supabase_client.table("user_activities") \
            .select("action_type, action_details") \
            .eq("user_id", user_id) \
            .gte("created_at", (datetime.now() - timedelta(days=30)).isoformat()) \
            .execute()

        activities = result.data if result.data else []

        logins = len([a for a in activities if a.get("action_type") == "login"])
        page_views = len([a for a in activities if a.get("action_type") == "page_view"])

        bc_activities = [a for a in activities if a.get("action_type") == "generate_bc"]
        bc_count = len(bc_activities)
        bc_total = sum(
            (a.get("action_details") or {}).get("total_amount", 0)
            for a in bc_activities
        )

        return {
            "logins": logins,
            "bc_count": bc_count,
            "bc_total": bc_total,
            "pages_viewed": page_views
        }
    except Exception as e:
        print(f"⚠️ Erreur get_user_stats: {e}")
        return {"logins": 0, "bc_count": 0, "bc_total": 0, "pages_viewed": 0}


# ============================================================
# 🎯 OBJECTIFS ET FOURNISSEURS AGENTS
# ============================================================

def get_agent_targets(username: str) -> dict:
    """
    Retourne les objectifs mensuels d'un agent.
    Les objectifs sont définis par rôle ou personnalisés par agent.
    """
    # Objectifs par défaut selon le rôle
    user_info = get_user_info(username) or {}
    user_role = user_info.get("role", "user")

    # Objectifs différenciés par rôle
    targets_by_role = {
        "admin": {"target_monthly": 50_000_000, "target_orders": 30},      # 50M / 30 BC
        "manager": {"target_monthly": 30_000_000, "target_orders": 20},    # 30M / 20 BC
        "user": {"target_monthly": 15_000_000, "target_orders": 15},       # 15M / 15 BC
    }

    # Objectifs personnalisés par agent (optionnel)
    custom_targets = {
        "arame": {"target_monthly": 20_000_000, "target_orders": 18},
        "fallou": {"target_monthly": 18_000_000, "target_orders": 16},
        "coumba": {"target_monthly": 15_000_000, "target_orders": 14},
        "ravane": {"target_monthly": 15_000_000, "target_orders": 14},
        "insa": {"target_monthly": 12_000_000, "target_orders": 12},
    }

    username_lower = username.lower().strip()

    # Priorité aux objectifs personnalisés
    if username_lower in custom_targets:
        return custom_targets[username_lower]

    # Sinon objectifs par rôle
    return targets_by_role.get(user_role, targets_by_role["user"])


def get_agent_suppliers(username: str, user_role: str) -> list:
    """
    Retourne la liste des fournisseurs assignés à un agent.
    - Admin/Manager: voit TOUS les fournisseurs (retourne None)
    - User: voit uniquement SES fournisseurs assignés
    """
    # Admin et Manager voient tout
    if user_role in MANAGER_ROLES:
        return None  # None = pas de filtre = tout visible

    username_lower = username.lower().strip()

    # Charger le mapping agents → fournisseurs depuis Google Sheets
    try:
        agents_df = load_agents_suppliers()

        if agents_df.empty:
            print(f"   ⚠️ DataFrame agents vide pour {username}")
            return None

        # Trouver la colonne Owner
        owner_col = None
        supplier_col = None

        for col in agents_df.columns:
            col_lower = col.lower().strip()
            if 'owner' in col_lower or 'agent' in col_lower:
                owner_col = col
            if 'fournisseur' in col_lower or 'supplier' in col_lower:
                supplier_col = col

        if not owner_col or not supplier_col:
            print(f"   ⚠️ Colonnes Owner/Supplier non trouvées: {agents_df.columns.tolist()}")
            return None

        # Filtrer par agent
        agents_df[owner_col] = agents_df[owner_col].astype(str).str.lower().str.strip()
        agent_rows = agents_df[agents_df[owner_col] == username_lower]

        if agent_rows.empty:
            # Essayer avec le prénom seulement
            user_info = get_user_info(username) or {}
            first_name = user_info.get("name", "").split()[0].lower() if user_info.get("name") else ""

            if first_name:
                agent_rows = agents_df[agents_df[owner_col].str.contains(first_name, na=False)]

        if agent_rows.empty:
            print(f"   ⚠️ Aucun fournisseur trouvé pour {username}")
            return None

        # Extraire les fournisseurs uniques
        suppliers = agent_rows[supplier_col].dropna().unique().tolist()
        suppliers = [s.strip() for s in suppliers if s and str(s).strip()]

        print(f"   🔐 Agent {username}: {len(suppliers)} fournisseurs assignés")
        return suppliers if suppliers else None

    except Exception as e:
        print(f"   ⚠️ Erreur get_agent_suppliers: {e}")
        return None

def track_agent_order(username: str, supplier: str, order_total: float, products_count: int, po_number: str):
    """
    Enregistre une commande dans l'historique de l'agent.
    """
    if not supabase_client:
        return False

    try:
        data = {
            "agent_username": username.lower().strip(),
            "supplier": supplier,
            "order_total": order_total,
            "products_count": products_count,
            "po_number": po_number,
            "created_at": datetime.now().isoformat()
        }

        result = supabase_client.table("agent_orders").insert(data).execute()
        print(f"   📊 Commande trackée: {username} → {supplier} ({order_total:,.0f} FCFA)")
        return True
    except Exception as e:
        print(f"   ⚠️ Erreur tracking commande: {e}")
        return False


def get_agent_stats(username: str, period_days: int = 30) -> dict:
    """
    Récupère les statistiques d'un agent sur une période.
    """
    if not supabase_client:
        return _get_default_agent_stats()

    try:
        # Date de début de la période
        start_date = (datetime.now() - timedelta(days=period_days)).isoformat()

        result = supabase_client.table("agent_orders") \
            .select("*") \
            .eq("agent_username", username.lower().strip()) \
            .gte("created_at", start_date) \
            .execute()

        if not result.data:
            return _get_default_agent_stats()

        orders = result.data

        # Calculs
        total_amount = sum(o.get("order_total", 0) for o in orders)
        total_orders = len(orders)
        total_products = sum(o.get("products_count", 0) for o in orders)

        # Par fournisseur
        by_supplier = {}
        for o in orders:
            sup = o.get("supplier", "N/A")
            if sup not in by_supplier:
                by_supplier[sup] = {"orders": 0, "amount": 0, "products": 0}
            by_supplier[sup]["orders"] += 1
            by_supplier[sup]["amount"] += o.get("order_total", 0)
            by_supplier[sup]["products"] += o.get("products_count", 0)

        return {
            "total_amount": total_amount,
            "total_orders": total_orders,
            "total_products": total_products,
            "avg_order_value": total_amount / total_orders if total_orders > 0 else 0,
            "by_supplier": by_supplier,
            "period_days": period_days,
            "orders_list": orders[-10:]  # 10 dernières commandes
        }

    except Exception as e:
        print(f"   ⚠️ Erreur récupération stats agent: {e}")
        return _get_default_agent_stats()


def _get_default_agent_stats() -> dict:
    """Stats par défaut si pas de données"""
    return {
        "total_amount": 0,
        "total_orders": 0,
        "total_products": 0,
        "avg_order_value": 0,
        "by_supplier": {},
        "period_days": 30,
        "orders_list": []
    }


def get_agent_performance_kpis(username: str) -> dict:
    """
    Calcule les indicateurs de performance d'un agent.
    """
    stats = get_agent_stats(username, period_days=30)
    targets = get_agent_targets(username)

    # Calcul des % d'atteinte
    amount_pct = (stats["total_amount"] / targets["target_monthly"] * 100) if targets["target_monthly"] > 0 else 0
    orders_pct = (stats["total_orders"] / targets["target_orders"] * 100) if targets["target_orders"] > 0 else 0

    # Score global (moyenne pondérée)
    score = (amount_pct * 0.6 + orders_pct * 0.4)

    # Niveau de performance
    if score >= 100:
        level = "🏆 Excellent"
        color = "#22c55e"
    elif score >= 80:
        level = "✅ Bon"
        color = "#3b82f6"
    elif score >= 60:
        level = "⚠️ À améliorer"
        color = "#f59e0b"
    else:
        level = "🔴 En difficulté"
        color = "#ef4444"

    return {
        "stats": stats,
        "targets": targets,
        "amount_pct": round(amount_pct, 1),
        "orders_pct": round(orders_pct, 1),
        "score": round(score, 1),
        "level": level,
        "level_color": color
    }


# ============================================================
# 🔐 SYSTÈME D'AUTHENTIFICATION
# ============================================================

# Mot de passe par défaut: "maad2025" (à changer en production via variables d'environnement)
DEFAULT_PASSWORD_HASH = hashlib.sha256("maad2025".encode()).hexdigest()

# Utilisateurs autorisés avec leurs mots de passe hashés
AUTH_USERS = {
    "tony": {
        "name": "Tony SARRE",
        "email": "tony.sarre@maad.io",
        "password_hash": hashlib.sha256(os.getenv("TONY_PASSWORD", "maad2025").encode()).hexdigest(),
        "role": "admin"
    },
    "samuel": {
        "name": "Samuel Essodeke",
        "email": "essodeke@maad.io",
        "password_hash": DEFAULT_PASSWORD_HASH,
        "role": "manager"
    },
    "maimouna": {
        "name": "Maimouna Dagois",
        "email": "maimouna@maad.io",
        "password_hash": DEFAULT_PASSWORD_HASH,
        "role": "user"
    },
    "seydouna": {
        "name": "Seydouna Oumar Niang",
        "email": "seydouna@maad.io",
        "password_hash": DEFAULT_PASSWORD_HASH,
        "role": "manager"
    },
    "arame": {
        "name": "Arame Toure",
        "email": "arame.toure@maad.io",
        "password_hash": DEFAULT_PASSWORD_HASH,
        "role": "user"
    },
    "coumba": {
        "name": "Coumba Cisse",
        "email": "ndeyecoumba.cisse@maad.io",
        "password_hash": DEFAULT_PASSWORD_HASH,
        "role": "user"
    },
    "fallou": {
        "name": "Fallou Diop",
        "email": "serigne.diop@maad.io",
        "password_hash": DEFAULT_PASSWORD_HASH,
        "role": "user"
    },
    "ravane": {
        "name": "Ravane Diop",
        "email": "pr.diop@maad.io",
        "password_hash": DEFAULT_PASSWORD_HASH,
        "role": "user"
    },
    "insa": {
        "name": "Insa Niang",
        "email": "insa.niang@maad.io",
        "password_hash": DEFAULT_PASSWORD_HASH,
        "role": "user"
    },
}


def verify_password(username: str, password: str) -> bool:
    """Vérifie le mot de passe d'un utilisateur et track la connexion"""
    username = username.lower().strip()
    print(f"   🔍 DEBUG verify_password: username='{username}'")

    if username not in AUTH_USERS:
        print(f"   ❌ Utilisateur '{username}' non trouvé dans AUTH_USERS")
        print(f"   📋 Utilisateurs disponibles: {list(AUTH_USERS.keys())}")
        track_login(username, success=False)  # Track tentative échouée
        return False

    password_hash = hashlib.sha256(password.encode()).hexdigest()
    expected_hash = AUTH_USERS[username]["password_hash"]

    print(f"   🔐 Hash fourni: {password_hash[:20]}...")
    print(f"   🔐 Hash attendu: {expected_hash[:20]}...")

    if password_hash == expected_hash:
        print(f"   ✅ Mot de passe correct!")
        track_login(username, success=True)  # Track connexion réussie
        return True
    else:
        print(f"   ❌ Mot de passe incorrect")
        track_login(username, success=False)  # Track tentative échouée
        return False


def get_user_info(username: str) -> dict:
    """Retourne les informations d'un utilisateur"""
    username = username.lower().strip()
    if username in AUTH_USERS:
        return {
            "username": username,
            "name": AUTH_USERS[username]["name"],
            "email": AUTH_USERS[username]["email"],
            "role": AUTH_USERS[username]["role"]
        }
    return None


# ============================================================
# 🎨 CSS STYLES POUR LA PAGE DE CONNEXION
# ============================================================

LOGIN_CSS = """
/* ===== FOND ANIMÉ ===== */
.login-background {
    min-height: 100vh;
    background: linear-gradient(-45deg, #0f172a, #1e293b, #0f172a, #1a1a2e);
    background-size: 400% 400%;
    animation: gradientBG 15s ease infinite;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
}

@keyframes gradientBG {
    0% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}

/* ===== CARTE LOGIN - FOND BLANC ===== */
.login-card {
    background: rgba(255, 255, 255, 0.98) !important;
    backdrop-filter: blur(20px);
    border-radius: 24px;
    padding: 48px;
    width: 100%;
    max-width: 420px;
    box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5), 0 0 0 1px rgba(255, 255, 255, 0.1);
    border: 1px solid rgba(255, 255, 255, 0.1);
}

/* ===== LOGO ===== */
.login-logo {
    text-align: center;
    margin-bottom: 32px;
}

.login-logo-icon {
    font-size: 64px;
    margin-bottom: 16px;
    animation: pulse 2s infinite;
}

@keyframes pulse {
    0%, 100% { transform: scale(1); }
    50% { transform: scale(1.05); }
}

.login-title {
    font-size: 28px;
    font-weight: 800;
    background: linear-gradient(135deg, #0ea5e9 0%, #7c3aed 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    margin-bottom: 8px;
}

.login-subtitle {
    color: #475569;
    font-size: 14px;
}

/* ===== FORMULAIRE ===== */
.login-form {
    margin-top: 24px;
}

.login-input-group {
    position: relative;
    margin-bottom: 20px;
}

.login-input-icon {
    position: absolute;
    left: 16px;
    top: 50%;
    transform: translateY(-50%);
    color: #64748b;
    font-size: 18px;
    z-index: 10;
}

/* ✅ INPUTS FORCÉS LISIBLES - FOND BLANC, TEXTE NOIR */
.login-input,
#login-username,
#login-password,
.login-form input,
.login-form input[type="text"],
.login-form input[type="password"] {
    width: 100%;
    padding: 16px 16px 16px 48px;
    background: #ffffff !important;
    border: 2px solid #e2e8f0 !important;
    border-radius: 12px;
    color: #1e293b !important;
    font-size: 15px;
    transition: all 0.3s ease;
    -webkit-text-fill-color: #1e293b !important;
}

#login-username:focus,
#login-password:focus,
.login-input:focus {
    outline: none;
    border-color: #0ea5e9 !important;
    box-shadow: 0 0 0 4px rgba(14, 165, 233, 0.15);
    background: #ffffff !important;
}

#login-username::placeholder,
#login-password::placeholder,
.login-input::placeholder {
    color: #94a3b8 !important;
    -webkit-text-fill-color: #94a3b8 !important;
}

/* ===== BOUTON ===== */
.login-button {
    width: 100%;
    padding: 16px;
    background: linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%);
    border: none;
    border-radius: 12px;
    color: #ffffff;
    font-size: 16px;
    font-weight: 700;
    cursor: pointer;
    transition: all 0.3s ease;
    margin-top: 8px;
    text-transform: uppercase;
    letter-spacing: 1px;
}

.login-button:hover {
    transform: translateY(-2px);
    box-shadow: 0 10px 25px rgba(34, 211, 238, 0.4);
}

.login-button:active {
    transform: translateY(0);
}

/* ===== MESSAGE ERREUR ===== */
.login-error {
    background: rgba(239, 68, 68, 0.15);
    border: 1px solid #ef4444;
    border-radius: 10px;
    padding: 12px 16px;
    margin-bottom: 20px;
    color: #dc2626;
    font-size: 13px;
    display: flex;
    align-items: center;
    gap: 10px;
}

/* ===== USER NAVBAR ===== */
.user-navbar {
    position: fixed;
    top: 16px;
    right: 16px;
    z-index: 9999;
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 10px 16px;
    background: rgba(30, 41, 59, 0.95);
    backdrop-filter: blur(10px);
    border-radius: 50px;
    border: 1px solid #334155;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
}

.user-avatar {
    width: 36px;
    height: 36px;
    background: linear-gradient(135deg, #0ea5e9, #7c3aed);
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 700;
    color: #0f172a;
    font-size: 14px;
}

.user-info {
    display: flex;
    flex-direction: column;
}

.user-name {
    color: #f0f4f8;
    font-weight: 600;
    font-size: 14px;
}

.user-role {
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.logout-btn {
    background: rgba(239, 68, 68, 0.2);
    border: 1px solid rgba(239, 68, 68, 0.4);
    border-radius: 8px;
    padding: 8px 12px;
    cursor: pointer;
    font-size: 16px;
    transition: all 0.2s;
    color: #f87171;
}

.logout-btn:hover {
    background: rgba(239, 68, 68, 0.4);
    border-color: #ef4444;
}
"""


def create_login_layout():
    """Crée le layout de la page de connexion avec image de fond Supply Chain"""

    # Image de fond encodée en base64
    bg_image_url = "url('data:image/jpeg;base64,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')"

    return html.Div([
        # Conteneur principal avec image de fond
        html.Div(className="login-background", style={
            "minHeight": "100vh",
            "backgroundImage": bg_image_url,
            "backgroundSize": "cover",
            "backgroundPosition": "center",
            "backgroundRepeat": "no-repeat",
            "display": "flex",
            "alignItems": "center",
            "justifyContent": "center",
            "padding": "20px",
            "fontFamily": "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
            "position": "relative"
        }, children=[

            # Overlay sombre semi-transparent pour lisibilité
            html.Div(style={
                "position": "absolute",
                "top": "0",
                "left": "0",
                "right": "0",
                "bottom": "0",
                "background": "rgba(15, 23, 42, 0.7)",
                "zIndex": "1"
            }),

            # Carte de connexion
            html.Div(className="login-card", style={
                "background": "rgba(255, 255, 255, 0.95)",
                "backdropFilter": "blur(20px)",
                "borderRadius": "24px",
                "padding": "48px",
                "width": "100%",
                "maxWidth": "420px",
                "boxShadow": "0 25px 50px -12px rgba(0, 0, 0, 0.4), 0 0 0 1px rgba(255, 255, 255, 0.2)",
                "border": "1px solid rgba(255, 255, 255, 0.3)",
                "position": "relative",
                "zIndex": "10"
            }, children=[

                # Logo MAAD et titre
                html.Div(className="login-logo", style={"textAlign": "center", "marginBottom": "32px"}, children=[
                    # Logo MAAD (image)
                    html.Img(src="/assets/logo_maad.png", style={
                        "width": "160px",
                        "height": "auto",
                        "marginBottom": "20px"
                    }),
                    html.H1("Supply Chain Analytics", style={
                        "fontSize": "24px",
                        "fontWeight": "700",
                        "color": "#0f172a",
                        "marginBottom": "6px"
                    }),
                    html.P("Connectez-vous pour accéder au dashboard", style={
                        "color": "#64748b",
                        "fontSize": "13px"
                    })
                ]),

                # Zone d'erreur
                html.Div(id="login-error", style={"display": "none"}),

                # Formulaire
                html.Div(className="login-form", style={"marginTop": "20px"}, children=[

                    html.Label("Identifiant", style={
                        "display": "block", "marginBottom": "6px", "color": "#374151",
                        "fontWeight": "500", "fontSize": "13px"
                    }),

                    dcc.Input(
                        id="login-username", type="text", placeholder="Votre identifiant",
                        style={
                            "width": "100%", "padding": "12px 14px", "background": "#fff",
                            "border": "1px solid #d1d5db", "borderRadius": "8px",
                            "color": "#1e293b", "fontSize": "14px", "boxSizing": "border-box",
                            "marginBottom": "16px"
                        },
                        autoComplete="username"
                    ),

                    html.Label("Mot de passe", style={
                        "display": "block", "marginBottom": "6px", "color": "#374151",
                        "fontWeight": "500", "fontSize": "13px"
                    }),

                    dcc.Input(
                        id="login-password", type="password", placeholder="Votre mot de passe",
                        style={
                            "width": "100%", "padding": "12px 14px", "background": "#fff",
                            "border": "1px solid #d1d5db", "borderRadius": "8px",
                            "color": "#1e293b", "fontSize": "14px", "boxSizing": "border-box",
                            "marginBottom": "20px"
                        },
                        autoComplete="current-password"
                    ),

                    dbc.Button(
                        "Se connecter",
                        id="login-button", n_clicks=0, color="primary", className="w-100",
                        style={
                            "padding": "12px", "background": "#0ea5e9", "border": "none",
                            "borderRadius": "8px", "color": "#fff", "fontSize": "14px",
                            "fontWeight": "600", "cursor": "pointer"
                        }
                    )
                ]),

                # Footer
                html.Div([
                    html.P([
                        "Mot de passe oublié? Contactez ",
                        html.A("l'administrateur", href="mailto:tony.sarre@maad.io",
                               style={"color": "#0ea5e9", "textDecoration": "none"})
                    ], style={"color": "#64748b", "fontSize": "11px", "textAlign": "center", "marginTop": "24px"}),
                    html.P("© 2024 MAAD",
                           style={"color": "#94a3b8", "fontSize": "10px", "textAlign": "center", "marginTop": "6px"})
                ])
            ])
        ])
    ])


def create_user_navbar(username: str):
    """Navbar horizontale professionnelle."""
    if not username or username.lower() not in AUTH_USERS:
        return html.Div()

    user = AUTH_USERS[username.lower()]
    initials = "".join([n[0].upper() for n in user["name"].split()[:2]])
    first_name = user["name"].split()[0] if user["name"] else username
    user_role = user.get("role", "user")

    role_colors = {"admin": "#ef4444", "manager": "#f59e0b", "user": "#0ea5e9"}
    role_labels = {"admin": "Admin", "manager": "Manager", "user": "User"}

    nav_items = [
        {"label": "Overview", "href": "/"},
        {"label": "Mon Espace", "href": "/myspace"},
        {"label": "Analytics", "href": "/analytics"},
        {"label": "Prédictions", "href": "/predictions"},
        {"label": "Promos", "href": "/promotions"},
        {"label": "About", "href": "/about"},
    ]

    return html.Div([
        html.Nav(style={
            "position": "fixed", "top": "0", "left": "0", "right": "0",
            "height": "50px", "background": "#0f172a",
            "borderBottom": "1px solid #1e293b",
            "display": "flex", "alignItems": "center", "justifyContent": "space-between",
            "padding": "0 16px", "zIndex": "9999"
        }, children=[

            # Logo + Brand
            html.Div(style={"display": "flex", "alignItems": "center", "gap": "10px"}, children=[
                html.Img(src=LOGO_DATA_URI, style={"height": "32px"}) if LOGO_DATA_URI else html.Div(),
                html.Div([
                    html.Div(APP_BRAND, style={"color": "#f8fafc", "fontWeight": "600", "fontSize": "14px"}),
                    html.Div("Supply Chain", style={"color": "#64748b", "fontSize": "9px", "textTransform": "uppercase"})
                ])
            ]),

            # Navigation
            html.Div(style={"display": "flex", "alignItems": "center", "gap": "2px"}, children=[
                dcc.Link(
                    html.Div(item["label"], style={
                        "padding": "6px 12px", "borderRadius": "4px", "color": "#94a3b8",
                        "fontSize": "12px", "fontWeight": "500"
                    }),
                    href=item["href"],
                    style={"textDecoration": "none"},
                    className="nav-link-hover"
                ) for item in nav_items
            ]),

            # User Info
            html.Div(style={"display": "flex", "alignItems": "center", "gap": "8px"}, children=[
                html.Div(style={
                    "display": "flex", "alignItems": "center", "gap": "8px",
                    "padding": "4px 10px", "background": "rgba(255,255,255,0.05)",
                    "borderRadius": "6px", "border": "1px solid #334155"
                }, children=[
                    html.Div(initials, style={
                        "width": "28px", "height": "28px",
                        "background": f"linear-gradient(135deg, {role_colors.get(user_role, '#0ea5e9')}, #7c3aed)",
                        "borderRadius": "50%", "display": "flex", "alignItems": "center",
                        "justifyContent": "center", "fontWeight": "600", "color": "#fff", "fontSize": "11px"
                    }),
                    html.Div([
                        html.Div(first_name, style={"color": "#f1f5f9", "fontWeight": "500", "fontSize": "12px"}),
                        html.Div(role_labels.get(user_role, "User"), style={
                            "color": role_colors.get(user_role, "#64748b"), "fontSize": "9px", "textTransform": "uppercase"
                        })
                    ])
                ]),
                html.Button("Déconnexion", id="logout-button", n_clicks=0, style={
                    "background": "rgba(239, 68, 68, 0.1)", "border": "1px solid rgba(239, 68, 68, 0.3)",
                    "borderRadius": "4px", "padding": "4px 8px", "cursor": "pointer",
                    "fontSize": "11px", "color": "#f87171"
                })
            ])
        ])
    ])


# ------------------------------ Email Configuration -------------------------------
# SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
# SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
# SMTP_USER = os.getenv("SMTP_USER", "")  # your-email@gmail.com
# SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")  # App password

# Mapping des utilisateurs (à personnaliser)
# TEAM_MEMBERS = {
#   "tony": {"name": "Tony SARRE", "email": "tony.sarre@maad.io"},
#  "Samuel": {"name": "Samuel Essodeke", "email": "essodeke@maad.io"},
# "Maimouna": {"name": "Maimouna Dagois", "email": "maimouna@maad.io"},
# "Seydouna": {"name": "Seydouna Oumar Niang", "email": "seydouna@maad.io"},
# }


# def send_notification_email(to_email: str, to_name: str, product_name: str, author: str, message: str):
#   """Envoie un email de notification"""
#  if not SMTP_USER or not SMTP_PASSWORD:
#     print("⚠️ Email non configuré (SMTP_USER/SMTP_PASSWORD manquants)")
#    return False

# try:
#   import smtplib
#  from email.mime.text import MIMEText
# from email.mime.multipart import MIMEMultipart

# msg = MIMEMultipart("alternative")
# msg["Subject"] = f"[Maad SaSu] Nouvelle mention sur {product_name}"
# msg["From"] = SMTP_USER
# msg["To"] = to_email

# html = f"""
# <html>
# <body style="font-family: Arial, sans-serif; color: #333;">
#   <div style="background: #0b1220; padding: 20px; border-radius: 10px;">
#      <h2 style="color: #0ea5e9;">📌 Nouvelle mention</h2>
#     <p style="color: #e5e7eb;">Bonjour {to_name},</p>
#    <p style="color: #e5e7eb;">
#       <strong>{author}</strong> vous a mentionné dans une note sur le produit
#      <strong style="color: #0ea5e9;">{product_name}</strong> :
# </p>
# <blockquote style="background: #1f2937; padding: 15px; border-left: 4px solid #0ea5e9; margin: 20px 0;">
#   <p style="color: #e5e7eb; font-style: italic;">{message}</p>
# </blockquote>
# <p style="color: #9ca3af; font-size: 12px;">
#   Date : {datetime.now().strftime('%d/%m/%Y à %H:%M')}
# </p>
# <a href="https://your-dashboard-url.com"
#  style="display: inline-block; background: #0ea5e9; color: #001018;
#        padding: 10px 20px; text-decoration: none; border-radius: 5px;
#       font-weight: bold; margin-top: 10px;">
# Voir le dashboard
# </a>
#      </div>
# </body>
# </html>
# """

# part = MIMEText(html, "html")
# msg.attach(part)

# with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
#   server.starttls()
#  server.login(SMTP_USER, SMTP_PASSWORD)
# server.send_message(msg)

# print(f" Email envoyé à {to_email}")
# return True

# except Exception as e:
#   print(f" Erreur envoi email: {e}")
#  return False

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
                print(f" Logo loaded from base64 ({len(raw)} bytes)")
                return ImageReader(io.BytesIO(raw))
            except Exception as e:
                print(f" Base64 logo decode failed: {e}")

        print(" No logo available, continuing without")
        return None

    except Exception as e:
        print(f" Logo loading error: {e}")
        return None


# ------------------------------ PO Numbering -------------------------------------
PO_COUNTER_PATH = Path("./po_counter.json")

# ------------------------------ Notes System -------------------------------------
NOTES_DB_PATH = Path("./notes_database.json")
NOTES_LOCK = threading.Lock()


def _load_notes():
    """Charge toutes les notes depuis le fichier JSON"""
    try:
        if NOTES_DB_PATH.exists():
            return json.loads(NOTES_DB_PATH.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"Erreur chargement notes: {e}")
    return []


def _save_notes(notes_list: list):
    """Sauvegarde les notes dans le fichier JSON"""
    try:
        with NOTES_LOCK:
            NOTES_DB_PATH.write_text(json.dumps(notes_list, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        print(f"Erreur sauvegarde notes: {e}")


def get_notes_for_product(product_name: str):
    """
    🚀 VERSION OPTIMISÉE - Récupère les notes d'un produit
    Priorité: Supabase > Fichier local
    """
    # 1. Essayer Supabase d'abord (persistant)
    if supabase_client:
        try:
            result = supabase_client.table("product_notes") \
                .select("*") \
                .ilike("product_name", product_name) \
                .eq("is_deleted", False) \
                .order("created_at", desc=True) \
                .limit(50) \
                .execute()

            if result.data:
                # Convertir au format attendu
                notes = []
                for n in result.data:
                    notes.append({
                        "id": n.get("id"),
                        "product_name": n.get("product_name"),
                        "author": n.get("user_id", "Utilisateur"),
                        "message": n.get("note_text", ""),
                        "mentions": n.get("mentions", []),
                        "timestamp": n.get("created_at", datetime.now().isoformat())
                    })
                return notes
        except Exception as e:
            print(f"⚠️ Supabase notes error: {e}")

    # 2. Fallback: fichier local
    notes = _load_notes()
    return [n for n in notes if n.get("product_name", "").lower() == product_name.lower()]


def add_note(product_name: str, author: str, message: str, mentions: list = None):
    """
    🚀 VERSION OPTIMISÉE - Ajoute une note
    Sauvegarde dans Supabase + fichier local (backup)
    """
    new_note = {
        "id": f"note_{datetime.now().strftime('%Y%m%d_%H%M%S')}",
        "product_name": product_name,
        "author": author,
        "message": message,
        "mentions": mentions or [],
        "timestamp": datetime.now().isoformat()
    }

    # 1. Sauvegarder dans Supabase
    if supabase_client:
        try:
            supabase_data = {
                "product_name": product_name,
                "user_id": author,
                "note_text": message,
                "mentions": mentions or [],
                "supplier": None
            }
            result = supabase_client.table("product_notes").insert(supabase_data).execute()
            if result.data:
                print(f"✅ Note sauvegardée dans Supabase pour: {product_name}")
        except Exception as e:
            print(f"⚠️ Supabase note save error: {e}")

    # 2. Backup local
    notes = _load_notes()
    notes.append(new_note)
    _save_notes(notes)

    return new_note


PO_LOCK = threading.Lock()

# ============================================================
# 🚀 SYSTÈME PO OPTIMISÉ - Cache en mémoire + persistence async
# ============================================================
_PO_STATE_CACHE = {"date": None, "seq": 0, "dirty": False}
_PO_SAVE_SCHEDULED = False


def _load_po_state():
    """Charge l'état depuis le fichier (uniquement au démarrage)"""
    global _PO_STATE_CACHE
    try:
        if PO_COUNTER_PATH.exists():
            import json as _json
            file_state = _json.loads(PO_COUNTER_PATH.read_text(encoding="utf-8"))
            _PO_STATE_CACHE = {**file_state, "dirty": False}
            return _PO_STATE_CACHE
    except Exception:
        pass
    return {"date": None, "seq": 0, "dirty": False}


def _save_po_state_async():
    """Sauvegarde l'état en arrière-plan (non-bloquant)"""
    global _PO_SAVE_SCHEDULED, _PO_STATE_CACHE

    def _do_save():
        global _PO_SAVE_SCHEDULED, _PO_STATE_CACHE
        try:
            with PO_LOCK:
                if _PO_STATE_CACHE.get("dirty"):
                    import json as _json
                    state_to_save = {k: v for k, v in _PO_STATE_CACHE.items() if k != "dirty"}
                    PO_COUNTER_PATH.write_text(_json.dumps(state_to_save, ensure_ascii=False), encoding="utf-8")
                    _PO_STATE_CACHE["dirty"] = False
        except Exception as e:
            print(f"⚠️ PO save error: {e}")
        finally:
            _PO_SAVE_SCHEDULED = False

    if not _PO_SAVE_SCHEDULED:
        _PO_SAVE_SCHEDULED = True
        _async_executor.submit(_do_save)


def _save_po_state(state: dict):
    """Sauvegarde synchrone (fallback)"""
    try:
        import json as _json
        PO_COUNTER_PATH.write_text(_json.dumps(state, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


def get_next_po_number() -> str:
    """
    🚀 VERSION OPTIMISÉE - Génère le prochain numéro PO
    Utilise un cache en mémoire et sauvegarde en arrière-plan
    """
    global _PO_STATE_CACHE
    today = datetime.now().strftime("%Y%m%d")

    with PO_LOCK:
        # Charger depuis fichier seulement si cache vide
        if _PO_STATE_CACHE.get("date") is None:
            _load_po_state()

        # Incrémenter
        if _PO_STATE_CACHE.get("date") != today:
            _PO_STATE_CACHE = {"date": today, "seq": 1, "dirty": True}
        else:
            _PO_STATE_CACHE["seq"] = int(_PO_STATE_CACHE.get("seq", 0)) + 1
            _PO_STATE_CACHE["dirty"] = True

        po_number = f"PO-{today}-{_PO_STATE_CACHE['seq']:03d}"

    # Sauvegarder en arrière-plan (non-bloquant)
    _save_po_state_async()

    return po_number


#import pandas as pd
#import numpy as np
from sklearn.ensemble import RandomForestRegressor, RandomForestClassifier
from sklearn.preprocessing import OneHotEncoder
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.impute import SimpleImputer
from sklearn.model_selection import train_test_split
from sklearn.metrics import precision_score, f1_score, mean_absolute_error, r2_score
#import warnings

warnings.filterwarnings('ignore')


def calculate_formula_based_target(df):
    """
    Formule de fallback pour produits sans historique de ventes.
    Calcule target_quantity basée sur la demande projetée.
    """
    # Gérer les colonnes qui peuvent ne pas exister
    ads = df.get('Average Daily Sales', pd.Series(0, index=df.index))
    leadtime = df.get('ADJUSTED_LEADTIME', pd.Series(7, index=df.index))
    credit = df.get('credit_days', pd.Series(14, index=df.index))
    buffer = df.get('AJUSTER_BUFFER', pd.Series(0, index=df.index))
    stock = df.get('total_stock', pd.Series(0, index=df.index))

    # Formule : demande × période + buffer - stock
    target = (
            ads * (leadtime + credit) +
            buffer * ads -
            stock
    )

    return np.maximum(0, target)


def load_sales_history():
    """Charge et nettoie l'historique des ventes Pikine."""
    SALES_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQAK0IcIDJS8ysyCB0wnLp-rR-t-zu_2_6bYV4-YIhPuL3fZQyo7fgMXZnJ4rcz-5mNur_UHgMenRiU/pub?gid=1493123930&single=true&output=csv"

    try:
        # Lire à partir de la ligne 2 (skiprows=1)
        sales_df = pd.read_csv(SALES_URL, skiprows=1)

        print(f"\nHistorique ventes brut : {len(sales_df)} lignes")
        print(f"Colonnes : {sales_df.columns.tolist()[:5]}")

        # Extraire colonnes pertinentes (indices 0, 1, 2)
        sales_clean = sales_df.iloc[:, [0, 1, 2]].copy()
        sales_clean.columns = ['date', 'product_name', 'quantity_sold']

        # Nettoyer
        sales_clean['date'] = pd.to_datetime(sales_clean['date'], errors='coerce')
        sales_clean['product_name'] = sales_clean['product_name'].astype(str).str.lower().str.strip()
        sales_clean['quantity_sold'] = pd.to_numeric(sales_clean['quantity_sold'], errors='coerce')

        # Filtrer données valides
        sales_clean = sales_clean[
            (sales_clean['date'].notna()) &
            (sales_clean['product_name'].notna()) &
            (sales_clean['quantity_sold'].notna()) &
            (sales_clean['quantity_sold'] > 0)
            ].copy()

        print(f"✅ Ventes valides : {len(sales_clean)} lignes")
        print(f"   Période : {sales_clean['date'].min()} à {sales_clean['date'].max()}")
        print(f"   Produits uniques : {sales_clean['product_name'].nunique()}")

        return sales_clean

    except Exception as e:
        print(f"❌ Erreur chargement ventes : {e}")
        return pd.DataFrame()


def create_supervised_target_from_sales(sales_history_df: pd.DataFrame, current_df: pd.DataFrame) -> tuple:
    """
    Entraîne un modèle supervisé basé sur l'historique réel des ventes.
    """
    from sklearn.ensemble import GradientBoostingRegressor
    from sklearn.model_selection import train_test_split
    from sklearn.metrics import mean_absolute_error, r2_score

    # ✅ PRÉSERVER product_id dès le début
    has_product_id = 'product_id' in current_df.columns
    product_id_backup = None
    if has_product_id:
        product_id_backup = current_df[['product_name', 'product_id']].copy()

    if sales_history_df.empty:
        print("Pas d'historique ventes")
        current_df['target_quantity'] = calculate_formula_based_target(current_df)
        return current_df, None

    # =========================
    # Agrégation ventes par produit
    # =========================
    max_date = sales_history_df['date'].max()

    # Statistiques globales
    sales_stats = sales_history_df.groupby('product_name').agg({
        'quantity_sold': ['sum', 'mean', 'std', 'max', 'count']
    }).reset_index()

    sales_stats.columns = [
        'product_name', 'total_sold_30d', 'avg_daily_sold',
        'std_sold', 'max_daily_sold', 'days_with_sales'
    ]

    # Calculer trend (15 derniers jours vs 15 précédents)
    cutoff = max_date - pd.Timedelta(days=15)
    recent = sales_history_df[sales_history_df['date'] >= cutoff]
    previous = sales_history_df[sales_history_df['date'] < cutoff]

    recent_avg = recent.groupby('product_name')['quantity_sold'].mean()
    previous_avg = previous.groupby('product_name')['quantity_sold'].mean()

    trend = (recent_avg / previous_avg).fillna(1.0).clip(0.5, 2.0)  # Limiter variations extrêmes
    sales_stats['trend_factor'] = sales_stats['product_name'].map(trend).fillna(1.0)

    print(f"\n📊 Statistiques ventes :")
    print(f"   Produits avec ventes : {len(sales_stats)}")
    print(f"   Ventes totales 30j : {sales_stats['total_sold_30d'].sum():.0f} unités")

    # =========================
    # Merge avec données actuelles
    # =========================
    df_ml = current_df.merge(sales_stats, on='product_name', how='left')

    # Features
    feature_cols = [
        'total_stock',
        'credit_days',
        'ADJUSTED_LEADTIME',
        'AJUSTER_BUFFER',
        'Daily OOS Rate (30d)',
        'avg_daily_sold',  # ← Ventes réelles historiques
        'std_sold',  # ← Volatilité
        'max_daily_sold',  # ← Pic
        'trend_factor',  # ← Tendance récente
        'days_with_sales'  # ← Fréquence vente
    ]

    # =========================
    # TARGET : Quantité optimale basée sur ventes réelles
    # =========================
    df_ml['replenishment_period'] = df_ml['ADJUSTED_LEADTIME'] + df_ml['credit_days']

    # Demande projetée = ventes moyennes × trend × période
    df_ml['projected_demand'] = (
            df_ml['avg_daily_sold'] *
            df_ml['trend_factor'] *
            df_ml['replenishment_period']
    )

    # Buffer sécurité = volatilité × période
    df_ml['safety_stock'] = df_ml['std_sold'] * np.sqrt(df_ml['replenishment_period'])

    # Target = demande + sécurité - stock actuel
    df_ml['target_supervised'] = np.maximum(
        0,
        df_ml['projected_demand'] + df_ml['safety_stock'] - df_ml['total_stock']
    )

    # =========================
    # Entraînement ML
    # =========================
    has_history = df_ml['avg_daily_sold'].notna()
    df_train = df_ml[has_history & df_ml[feature_cols].notna().all(axis=1)].copy()

    if len(df_train) < 50:
        print(f"Échantillons insuffisants ({len(df_train)}), utilisation target calculée")
        df_ml['target_quantity'] = df_ml['target_supervised'].fillna(
            calculate_formula_based_target(df_ml)
        )
        # ✅ RESTAURER product_id s'il a été perdu
        if product_id_backup is not None and 'product_id' not in df_ml.columns:
            df_ml = df_ml.merge(product_id_backup, on='product_name', how='left')
            df_ml['product_id'] = df_ml['product_id'].fillna(0).astype(int)
        return df_ml, None

    X = df_train[feature_cols].fillna(0)
    y = df_train['target_supervised']

    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42
    )

    model = GradientBoostingRegressor(
        n_estimators=200,
        max_depth=5,
        learning_rate=0.1,
        subsample=0.8,
        random_state=42
    )

    model.fit(X_train, y_train)

    y_pred = model.predict(X_test)
    mae = mean_absolute_error(y_test, y_pred)
    r2 = r2_score(y_test, y_pred)

    print(f"\n{'=' * 60}")
    print(f"ML SUPERVISÉ - Ventes historiques RÉELLES")
    print(f"{'=' * 60}")
    print(f"  Produits entraînés : {len(df_train)}")
    print(f"  MAE (test) : {mae:.1f} unités")
    print(f"  R² score : {r2:.3f}")

    importance = sorted(
        zip(feature_cols, model.feature_importances_),
        key=lambda x: x[1], reverse=True
    )[:5]
    print(f"  Top features :")
    for feat, imp in importance:
        print(f"    - {feat}: {imp:.3f}")
    print(f"{'=' * 60}\n")

    # Prédictions
    X_all = df_ml[feature_cols].fillna(0)
    predictions = model.predict(X_all)

    # Produits avec historique : prédiction ML
    # Clipper les prédictions
    predictions_clipped = np.maximum(predictions, 0)

    # Appliquer selon historique
    df_ml['target_quantity'] = np.where(
        has_history,
        predictions_clipped,  # ✅
        calculate_formula_based_target(df_ml)
    )

    # ✅ RESTAURER product_id s'il a été perdu
    if product_id_backup is not None and 'product_id' not in df_ml.columns:
        df_ml = df_ml.merge(product_id_backup, on='product_name', how='left')
        df_ml['product_id'] = df_ml['product_id'].fillna(0).astype(int)
        print(f"   ✅ product_id restauré dans ML: {(df_ml['product_id'] > 0).sum()} valides")

    return df_ml, model


def load_and_analyze_promotions():
    """
    Charge l'historique des promotions et analyse leur statut actuel.
    """
    PROMO_URL = "https://data.heroku.com/dataclips/uetqonheumawnovajgbuxrirklmg.csv"

    try:
        promo_df = pd.read_csv(PROMO_URL)
        print(f"\n Promotions chargées : {len(promo_df)} lignes")

        # ✅ Renommer 'name' → 'product_name'
        promo_df.rename(columns={'name': 'product_name'}, inplace=True)

        # Nettoyer
        promo_df['product_name'] = promo_df['product_name'].astype(str).str.lower().str.strip()
        promo_df['start_date'] = pd.to_datetime(promo_df['start_date'], errors='coerce')
        promo_df['end_date'] = pd.to_datetime(promo_df['end_date'], errors='coerce')

        # ✅ Discount par défaut (CSV n'a pas cette colonne)
        promo_df['discount_pct'] = 15  # 15% par défaut pour les promos

        # Filtrer promos valides uniquement
        promo_df = promo_df[
            promo_df['product_name'].notna() &
            promo_df['start_date'].notna() &
            promo_df['end_date'].notna()
            ].copy()

        # Statut actuel
        today = pd.Timestamp.now()
        promo_df['promo_status'] = promo_df.apply(
            lambda row: 'Active' if (row['start_date'] <= today <= row['end_date'])
            else 'Terminée' if row['end_date'] < today
            else 'À venir',
            axis=1
        )

        promo_df['days_remaining'] = (promo_df['end_date'] - today).dt.days
        promo_df['days_remaining'] = promo_df['days_remaining'].where(
            promo_df['promo_status'] == 'Active', 0
        ).fillna(0).astype(int)

        print(f" Promotions valides : {len(promo_df)}")
        print(f"   Actives : {(promo_df['promo_status'] == 'Active').sum()}")
        print(f"   Terminées : {(promo_df['promo_status'] == 'Terminée').sum()}")
        print(f"   À venir : {(promo_df['promo_status'] == 'À venir').sum()}")

        return promo_df

    except Exception as e:
        print(f" Erreur promotions : {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()


def calculate_promo_roi_analysis(sales_df: pd.DataFrame, promo_df: pd.DataFrame,
                                 current_df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcule le ROI réel des promotions en tenant compte :
    - Uplift des ventes
    - Perte de marge due au discount
    - Coûts promotionnels
    - Profit net généré
    """

    if sales_df.empty or promo_df.empty:
        print("Données insuffisantes pour analyse ROI promotions")
        return pd.DataFrame()

    # =========================
    # 1. Calculer l'uplift (comme avant)
    # =========================
    sales_df['date'] = pd.to_datetime(sales_df['date'])

    sales_with_promo = sales_df.merge(
        promo_df[['product_name', 'start_date', 'end_date', 'discount_pct']],
        on='product_name',
        how='left'
    )

    sales_with_promo['was_promo'] = (
            (sales_with_promo['date'] >= sales_with_promo['start_date']) &
            (sales_with_promo['date'] <= sales_with_promo['end_date'])
    ).fillna(False)

    # Agrégation ventes avec/sans promo
    promo_stats = sales_with_promo.groupby(['product_name', 'was_promo']).agg({
        'quantity_sold': ['mean', 'sum', 'count']
    }).reset_index()

    promo_stats.columns = ['product_name', 'was_promo', 'avg_sales', 'total_sales', 'days_count']

    # Pivot pour avoir une ligne par produit
    promo_pivot = promo_stats.pivot(index='product_name', columns='was_promo', values='avg_sales').reset_index()
    promo_pivot.columns = ['product_name', 'sales_without_promo', 'sales_with_promo']
    promo_pivot = promo_pivot.fillna(0)

    # Uplift en %
    promo_pivot['uplift_pct'] = np.where(
        promo_pivot['sales_without_promo'] > 0,
        ((promo_pivot['sales_with_promo'] - promo_pivot['sales_without_promo']) /
         promo_pivot['sales_without_promo'] * 100),
        0
    )

    # =========================
    # 2. Récupérer prix et marges depuis current_df
    # =========================
    # Merger avec données produit actuelles
    # ✅ product_id est optionnel - ne pas échouer s'il est absent
    merge_cols = ['product_name']
    if 'product_id' in current_df.columns:
        merge_cols.append('product_id')

    roi_df = promo_pivot.merge(
        current_df[merge_cols],
        on='product_name',
        how='left'
    )

    # S'assurer que product_id existe
    if 'product_id' not in roi_df.columns:
        roi_df['product_id'] = 0

    # Supposons qu'on a ces colonnes (ajustez selon vos données)
    # Si pas disponibles, utiliser des estimations
    if 'unit_price' in current_df.columns:
        roi_df = roi_df.merge(
            current_df[['product_name', 'unit_price']],
            on='product_name',
            how='left'
        )
    else:
        # Estimation prix moyen par catégorie ou fixe
        roi_df['unit_price'] = 5000  # CFA - À ajuster

    if 'margin_pct' in current_df.columns:
        roi_df = roi_df.merge(
            current_df[['product_name', 'margin_pct']],
            on='product_name',
            how='left'
        )
    else:
        # Estimation marge standard
        roi_df['margin_pct'] = 25  # 25% de marge - À ajuster

    # Récupérer discount moyen pratiqué
    avg_discount = promo_df.groupby('product_name')['discount_pct'].mean().reset_index()
    roi_df = roi_df.merge(avg_discount, on='product_name', how='left')
    roi_df['discount_pct'] = roi_df['discount_pct'].fillna(10)  # Défaut 10%

    # =========================
    # 3. CALCUL DU ROI RÉEL
    # =========================
    def calculate_net_roi(row):
        """
        Calcule le profit net et ROI d'une promotion.
        """
        baseline_sales = row['sales_without_promo']  # Ventes journalières moyennes sans promo
        uplift_pct = row['uplift_pct']
        unit_price = row['unit_price']
        margin_pct = row['margin_pct'] / 100
        discount_pct = row['discount_pct'] / 100

        if baseline_sales == 0 or uplift_pct <= 0:
            return {
                'additional_sales': 0,
                'revenue_loss': 0,
                'additional_profit': 0,
                'net_profit': 0,
                'roi_pct': 0
            }

        # Ventes additionnelles générées
        additional_sales = baseline_sales * (uplift_pct / 100)

        # COÛT 1 : Perte de marge sur ventes de base (qui auraient eu lieu de toute façon)
        revenue_loss_baseline = baseline_sales * unit_price * discount_pct * margin_pct

        # REVENUS : Profit sur ventes additionnelles (après discount)
        revenue_additional = additional_sales * unit_price * (1 - discount_pct)
        profit_additional = revenue_additional * margin_pct

        # PROFIT NET = Gain sur ventes additionnelles - Perte sur ventes de base
        net_profit = profit_additional - revenue_loss_baseline

        # ROI = (Profit net / Coût) × 100
        roi = (net_profit / revenue_loss_baseline * 100) if revenue_loss_baseline > 0 else 0

        return {
            'additional_sales_per_day': additional_sales,
            'revenue_loss_per_day': revenue_loss_baseline,
            'additional_profit_per_day': profit_additional,
            'net_profit_per_day': net_profit,
            'roi_pct': roi
        }

    # Appliquer calcul
    roi_results = roi_df.apply(calculate_net_roi, axis=1, result_type='expand')
    roi_df = pd.concat([roi_df, roi_results], axis=1)

    # =========================
    # 4. RECOMMANDATION BASÉE SUR ROI
    # =========================
    def recommend_frequency(roi_pct, uplift_pct):
        """
        Recommande fréquence promo basée sur ROI ET uplift.
        """
        if roi_pct > 100 and uplift_pct > 30:
            return 'Mensuelle'  # Très rentable
        elif roi_pct > 50 and uplift_pct > 20:
            return 'Bimensuelle'  # Rentable
        elif roi_pct > 20 and uplift_pct > 10:
            return 'Trimestrielle'  # Modérément rentable
        elif roi_pct > 0:
            return 'Semestrielle'  # Faiblement rentable
        else:
            return 'Non rentable'  # Perte d\'argent

    roi_df['promo_recommendation'] = roi_df.apply(
        lambda row: recommend_frequency(row['roi_pct'], row['uplift_pct']),
        axis=1
    )

    # Priorité stratégique
    roi_df['promo_priority'] = roi_df.apply(
        lambda row: 'Haute' if row['roi_pct'] > 100
        else 'Moyenne' if row['roi_pct'] > 50
        else 'Faible' if row['roi_pct'] > 0
        else 'Éviter',
        axis=1
    )

    # =========================
    # 5. STATS FINALES
    # =========================
    print(f"\n{'=' * 60}")
    print(f"ANALYSE ROI PROMOTIONS")
    print(f"{'=' * 60}")
    print(f"  Produits analysés : {len(roi_df)}")
    print(f"  ROI moyen : {roi_df['roi_pct'].mean():.1f}%")
    print(f"  Promos rentables (ROI > 0) : {(roi_df['roi_pct'] > 0).sum()}")
    print(f"  Promos très rentables (ROI > 100) : {(roi_df['roi_pct'] > 100).sum()}")
    print(f"  Promos non rentables : {(roi_df['roi_pct'] <= 0).sum()}")
    print(f"\nRecommandations :")
    print(roi_df['promo_recommendation'].value_counts().to_string())
    print(f"{'=' * 60}\n")

    return roi_df


# ==================== CALCUL ADS DYNAMIQUE DEPUIS HISTORIQUE ====================
def calculate_ads_by_period(period_days: int = 7) -> pd.DataFrame:
    """
    Calcule Average Daily Sales sur N jours depuis l'historique sales_pikine.

    Args:
        period_days: Nombre de jours (3, 7 ou 30)

    Returns:
        DataFrame avec colonnes ['product_name', 'Average Daily Sales ({period}d)']
    """
    SALES_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQAK0IcIDJS8ysyCB0wnLp-rR-t-zu_2_6bYV4-YIhPuL3fZQyo7fgMXZnJ4rcz-5mNur_UHgMenRiU/pub?gid=1493123930&single=true&output=csv"

    print(f"\n📊 Calcul ADS {period_days} jours depuis historique...")

    try:
        # ✅ Charger historique ventes (skiprows=1 comme dans votre code)
        sales_df = pd.read_csv(SALES_URL, skiprows=1, low_memory=False)

        print(f"   📥 {len(sales_df)} lignes chargées")

        # ✅ Extraire colonnes pertinentes (indices 0, 1, 2 selon votre structure)
        # Colonne 0 = date, Colonne 1 = product_name, Colonne 2 = quantity_sold
        if sales_df.shape[1] >= 3:
            sales_clean = sales_df.iloc[:, [0, 1, 2]].copy()
            sales_clean.columns = ['date', 'product_name', 'quantity_sold']
        else:
            print(f"   ❌ Structure inattendue ({sales_df.shape[1]} colonnes)")
            return pd.DataFrame(columns=['product_name', f'Average Daily Sales ({period_days}d)'])

        # ✅ Nettoyer données
        sales_clean['date'] = pd.to_datetime(sales_clean['date'], errors='coerce')
        sales_clean['product_name'] = sales_clean['product_name'].astype(str).str.lower().str.strip()
        sales_clean['quantity_sold'] = pd.to_numeric(sales_clean['quantity_sold'], errors='coerce')

        # Filtrer données valides
        sales_clean = sales_clean[
            (sales_clean['date'].notna()) &
            (sales_clean['product_name'].notna()) &
            (sales_clean['product_name'] != 'nan') &
            (sales_clean['quantity_sold'].notna()) &
            (sales_clean['quantity_sold'] > 0)
            ].copy()

        if sales_clean.empty:
            print(f"   ⚠️ Aucune vente valide")
            return pd.DataFrame(columns=['product_name', f'Average Daily Sales ({period_days}d)'])

        print(f"   ✅ {len(sales_clean)} ventes valides")

        # ✅ Filtrer sur les N derniers jours
        max_date = sales_clean['date'].max()
        cutoff_date = max_date - pd.Timedelta(days=period_days)

        sales_period = sales_clean[sales_clean['date'] >= cutoff_date].copy()

        if sales_period.empty:
            print(f"   ⚠️ Pas de ventes dans les {period_days} derniers jours")
            return pd.DataFrame(columns=['product_name', f'Average Daily Sales ({period_days}d)'])

        print(f"   📅 Période : {cutoff_date.date()} → {max_date.date()}")
        print(f"   📦 {len(sales_period)} ventes sur {period_days} jours")

        # ✅ Calculer moyenne par produit
        # Total ventes / Nombre de jours réels
        actual_days = (max_date - cutoff_date).days + 1

        ads = (
            sales_period.groupby('product_name')['quantity_sold']
            .sum()
            .reset_index()
        )

        ads['avg_daily_sales'] = ads['quantity_sold'] / actual_days
        ads = ads[['product_name', 'avg_daily_sales']].copy()
        ads.columns = ['product_name', f'Average Daily Sales ({period_days}d)']

        # ✅ Valeur minimale = 0.1 (éviter division par zéro)
        ads[f'Average Daily Sales ({period_days}d)'] = (
            ads[f'Average Daily Sales ({period_days}d)']
            .clip(lower=0.1)
            .round(2)
        )

        print(f"   ✅ {len(ads)} produits avec ADS calculée")
        print(f"   📊 ADS min : {ads[f'Average Daily Sales ({period_days}d)'].min():.2f}")
        print(f"   📊 ADS médiane : {ads[f'Average Daily Sales ({period_days}d)'].median():.2f}")
        print(f"   📊 ADS max : {ads[f'Average Daily Sales ({period_days}d)'].max():.2f}")
        print(f"   📊 ADS moyenne : {ads[f'Average Daily Sales ({period_days}d)'].mean():.2f}")

        return ads

    except Exception as e:
        print(f"   ❌ Erreur calcul ADS {period_days}j : {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame(columns=['product_name', f'Average Daily Sales ({period_days}d)'])

def load_supply_data(period_days: str = "7d") -> pd.DataFrame:
    """
    🚀 VERSION OPTIMISÉE - Chargement parallèle des données
    """
    import numpy as np
    import pandas as pd

    load_start = time.time()
    print(f"\n{'='*60}")
    print(f"🚀 LOAD_SUPPLY_DATA OPTIMISÉ (période: {period_days})")
    print(f"{'='*60}")

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
    inventory_pikine_staging = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQAK0IcIDJS8ysyCB0wnLp-rR-t-zu_2_6bYV4-YIhPuL3fZQyo7fgMXZnJ4rcz-5mNur_UHgMenRiU/pub?gid=1876150276&single=true&output=csv"
    sales_pikine = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQAK0IcIDJS8ysyCB0wnLp-rR-t-zu_2_6bYV4-YIhPuL3fZQyo7fgMXZnJ4rcz-5mNur_UHgMenRiU/pub?gid=1493123930&single=true&output=csv"
    Tbh_7dsales = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQAK0IcIDJS8ysyCB0wnLp-rR-t-zu_2_6bYV4-YIhPuL3fZQyo7fgMXZnJ4rcz-5mNur_UHgMenRiU/pub?gid=1080970598&single=true&output=csv"
    Tbh_30dsales_products = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQAK0IcIDJS8ysyCB0wnLp-rR-t-zu_2_6bYV4-YIhPuL3fZQyo7fgMXZnJ4rcz-5mNur_UHgMenRiU/pub?gid=1655420642&single=true&output=csv"
    Product_category = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQbJqHsr6Kifee7I91YD-7-sCZDWgM5GvxCeN0OqUvZhok0j-kDywguqe5I61y97b-uBhHbWraTIrux/pub?gid=803048228&single=true&output=csv"
    DELISTING_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQax3ZQW2QhDLE-waDewtdD8x_Q5tpn2FWzVJftr9egik4_JF3s2ytSYJmXh55aUnp79vmF-XtkaTmN/pub?gid=1681543945&single=true&output=csv"
    PARAMETRES_REPLENISH_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQAK0IcIDJS8ysyCB0wnLp-rR-t-zu_2_6bYV4-YIhPuL3fZQyo7fgMXZnJ4rcz-5mNur_UHgMenRiU/pub?gid=1011110883&single=true&output=csv"
    SUPPLIER_CATEGORIZATION_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQax3ZQW2QhDLE-waDewtdD8x_Q5tpn2FWzVJftr9egik4_JF3s2ytSYJmXh55aUnp79vmF-XtkaTmN/pub?gid=1938047484&single=true&output=csv"
    CATALOG_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRTyAxh6v8o0FXV0r7f6ALPDgmeJNkjTZITjrEoKBHo2gs_f3iyV8sFk8fOzcAsUSkJMXBJCpJnhQKi/pub?gid=751531326&single=true&output=csv"

    # ✅ NOUVEAU: URL des dernières réceptions (Heroku Dataclip)
    LAST_RECEPTIONS_URL = "https://data.heroku.com/dataclips/vftfnhyqonsbclucridburrwdfpe.csv"

    # =========================
    # 🚀 CHARGEMENT PARALLÈLE DES CSV
    # =========================
    print("\n📥 Chargement parallèle des données...")
    csv_start = time.time()

    # Configuration des CSV à charger en parallèle
    csv_configs = [
        {"name": "suppliers", "url": SUPPLIERS_URL},
        {"name": "leadtime", "url": url_leadtime},
        {"name": "inventory", "url": inventory_pikine_staging, "skiprows": 1},
        {"name": "sales", "url": sales_pikine, "header": 1, "low_memory": False},
        {"name": "sales_7d", "url": Tbh_7dsales},
        {"name": "sales_30d", "url": Tbh_30dsales_products},
        {"name": "categories", "url": Product_category},
        {"name": "parametres", "url": PARAMETRES_REPLENISH_URL},
        {"name": "catalog", "url": CATALOG_URL, "skiprows": 1, "keep_default_na": False, "na_values": []},
    ]

    # Charger tous les CSV en parallèle
    csv_data = load_csvs_parallel(csv_configs)

    csv_elapsed = time.time() - csv_start
    print(f"⚡ CSV chargés en parallèle: {csv_elapsed:.2f}s")

    # Extraire les DataFrames
    suppliers_df = csv_data.get("suppliers", pd.DataFrame())
    df_leadtime = csv_data.get("leadtime", pd.DataFrame())
    inventory_pikine_staging_df = csv_data.get("inventory", pd.DataFrame())
    sales_pikine_df = csv_data.get("sales", pd.DataFrame())
    Tbh_7dsales_df = csv_data.get("sales_7d", pd.DataFrame())
    Tbh_30dsales_products_df = csv_data.get("sales_30d", pd.DataFrame())
    Product_category_df = csv_data.get("categories", pd.DataFrame())
    parametres_replenish_df = csv_data.get("parametres", pd.DataFrame())
    catalog_raw = csv_data.get("catalog", pd.DataFrame())

    # ✅ DEBUG: Afficher les colonnes chargées
    print(f"\n📋 Inventaire chargé:")
    print(f"   - Colonnes: {inventory_pikine_staging_df.columns.tolist()[:8]}...")
    print(f"   - Lignes: {len(inventory_pikine_staging_df)}")

    # ✅ Vérifier si product_id existe déjà
    if 'product_id' in inventory_pikine_staging_df.columns:
        print(f"   - product_id existe! Échantillon: {inventory_pikine_staging_df['product_id'].head(5).tolist()}")
    else:
        # La première colonne doit être product_id
        if len(inventory_pikine_staging_df.columns) > 0:
            first_col = inventory_pikine_staging_df.columns[0]
            print(f"   - Première colonne: '{first_col}'")
            if first_col != 'product_id':
                inventory_pikine_staging_df.rename(columns={first_col: 'product_id'}, inplace=True)
                print(f"   → Renommée en 'product_id'")

    # Charger delisting séparément (skiprows spécifique)
    try:
        delisting_df = load_csv_fast(DELISTING_URL, skiprows=3, usecols=[1, 3])
        if not delisting_df.empty:
            delisting_df.columns = ["product_name", "delisting_status"]
    except Exception as e:
        print(f"Warning: Could not load delisting data: {e}")
        delisting_df = pd.DataFrame(columns=["product_name", "delisting_status"])

    if not parametres_replenish_df.empty:
        print("'Parametres Replenish' data loaded successfully.")

    # ✅ NOUVEAU: Charger les dernières réceptions (Heroku Dataclip)
    # Utilise le système de chargement asynchrone avec cache
    print("\n📦 Chargement des dernières réceptions...")

    try:
        # Essayer d'abord le cache
        last_receptions_df = get_cached_receptions()

        if last_receptions_df.empty:
            # Charger avec timeout de 8 secondes
            last_receptions_df = load_receptions_async(LAST_RECEPTIONS_URL, timeout=8)

            if not last_receptions_df.empty:
                print(f"   ✅ {len(last_receptions_df)} réceptions chargées")
            else:
                print("   ℹ️ Pas de données de réception disponibles")
                last_receptions_df = pd.DataFrame(columns=['product_name', 'last_reception_qty', 'last_reception_date'])
        else:
            print(f"   ✅ {len(last_receptions_df)} réceptions (depuis cache)")

    except Exception as e:
        print(f"   ⚠️ Erreur réceptions: {e}")
        last_receptions_df = pd.DataFrame(columns=['product_name', 'last_reception_qty', 'last_reception_date'])

    # Lancer le préchargement en arrière-plan pour la prochaine fois
    preload_receptions_background(LAST_RECEPTIONS_URL)

    # try:
    #     import requests
    #     print("\n📦 Chargement des dernières réceptions...")
    #     # Utiliser requests avec timeout au lieu de pd.read_csv direct
    #     response = requests.get(LAST_RECEPTIONS_URL, timeout=10)
    #     if response.status_code == 200:
    #         from io import StringIO
    #         last_receptions_df = pd.read_csv(StringIO(response.text))
    #         print(f"   ✅ {len(last_receptions_df)} réceptions chargées")
    #         print(f"   📊 Colonnes: {last_receptions_df.columns.tolist()}")
    #
    #         # Harmoniser les noms de colonnes
    #         last_receptions_df.columns = (
    #             last_receptions_df.columns.astype(str)
    #             .str.strip()
    #             .str.lower()
    #             .str.replace(" ", "_")
    #         )
    #
    #         # Identifier les colonnes clés
    #         col_renames = {}
    #         for col in last_receptions_df.columns:
    #             if 'product' in col and 'name' in col:
    #                 col_renames[col] = 'product_name'
    #             elif col in ['product', 'produit', 'nom_produit']:
    #                 col_renames[col] = 'product_name'
    #             elif 'quantity' in col or 'qty' in col or 'qte' in col:
    #                 col_renames[col] = 'last_reception_qty'
    #             elif 'date' in col:
    #                 col_renames[col] = 'last_reception_date'
    #
    #         if col_renames:
    #             last_receptions_df.rename(columns=col_renames, inplace=True)
    #             print(f"   🔄 Colonnes renommées: {col_renames}")
    #     else:
    #         print(f"⚠️ Erreur HTTP {response.status_code}")
    #         last_receptions_df = pd.DataFrame(columns=['product_name', 'last_reception_qty', 'last_reception_date'])
    #
    # except Exception as e:
    #     print(f"⚠️ Warning: Could not load last receptions data: {e}")
    #     last_receptions_df = pd.DataFrame(columns=['product_name', 'last_reception_qty', 'last_reception_date'])

    # =========================
    # Harmonisation inventaire
    # =========================
    inv = inventory_pikine_staging_df.copy()

    print(f"\n📋 Harmonisation inventaire:")
    print(f"   Colonnes originales: {inv.columns.tolist()[:10]}...")

    # ✅ Normaliser les noms de colonnes en minuscules
    inv.columns = (
        inv.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace("-", "_")
    )

    print(f"   Colonnes normalisées: {inv.columns.tolist()[:10]}...")

    # ✅ Mapping des colonnes (tout en minuscules maintenant)
    col_map = {
        "product_name": ["product_name", "produit", "nom_produit", "name"],
        "supplier": ["supplier", "supplier.1", "fournisseur", "vendor"],
        "total_stock": ["total_stock", "stock_total", "qte_stock", "current_stock"],
    }
    for target, aliases in col_map.items():
        if target not in inv.columns:
            for alias in aliases:
                if alias in inv.columns:
                    inv.rename(columns={alias: target}, inplace=True)
                    print(f"   ✅ '{alias}' → '{target}'")
                    break
        if target not in inv.columns and target != "supplier":
            inv[target] = 0
            print(f"   ⚠️ '{target}' créée avec valeur 0")

    # ✅ Renommer supplier en Supplier (majuscule) pour cohérence avec le reste du code
    if "supplier" in inv.columns:
        inv.rename(columns={"supplier": "Supplier"}, inplace=True)

    # ✅ Convertir product_id en int
    if "product_id" in inv.columns:
        inv['product_id'] = pd.to_numeric(inv['product_id'], errors='coerce').fillna(0).astype(int)
        valid = (inv['product_id'] > 0).sum()
        print(f"   ✅ product_id: {valid}/{len(inv)} valides")
        print(f"      Échantillon: {inv['product_id'].head(5).tolist()}")
    else:
        print("   ⚠️ product_id ABSENT - création avec 0")
        inv['product_id'] = 0

    inventory_pikine_staging_df = inv
    print(f"   Colonnes finales: {inventory_pikine_staging_df.columns.tolist()[:10]}...")

    # Charger le catalogue pour is_active ET product_id
    # ✅ Le catalogue utilise "id" et "name" (pas "product_id" et "product_name")
    try:
        product_id_map = pd.DataFrame(columns=['product_name', 'product_id', 'is_active'])

        # Essayer différentes valeurs de skiprows
        for skip in [0, 1, 2, 3]:
            print(f"\n📋 Tentative catalogue avec skiprows={skip}...")
            # ✅ IMPORTANT: keep_default_na=False pour ne pas convertir #N/A en NaN
            catalog_df = pd.read_csv(
                CATALOG_URL,
                skiprows=skip if skip > 0 else None,
                keep_default_na=False,
                na_values=[]
            )

            print(f"   Colonnes: {catalog_df.columns.tolist()[:10]}...")

            # Normaliser les noms de colonnes
            catalog_df.columns = [str(c).lower().strip() for c in catalog_df.columns]

            # Vérifier si on a des colonnes avec de vrais noms (pas "unnamed")
            named_cols = [c for c in catalog_df.columns if 'unnamed' not in c.lower() and c != '' and '#n/a' not in c.lower()]
            print(f"   Colonnes nommées: {len(named_cols)} -> {named_cols[:8]}")

            # Chercher product_id et product_name
            # ✅ IMPORTANT: Le catalogue utilise "id" et "name" !
            col_product_id = None
            col_product_name = None
            col_is_active = None

            for col in catalog_df.columns:
                col_lower = col.lower().strip()

                # Chercher product_id (peut s'appeler "id" ou "product_id")
                if col_product_id is None and col_lower in ['id', 'product_id']:
                    test_vals = pd.to_numeric(catalog_df[col], errors='coerce')
                    if test_vals.notna().sum() > len(catalog_df) * 0.3:
                        col_product_id = col
                        print(f"   ✅ product_id trouvé: '{col}' -> échantillon: {test_vals.head(3).tolist()}")

                # Chercher product_name (peut s'appeler "name" ou "product_name")
                if col_product_name is None and col_lower in ['name', 'product_name']:
                    col_product_name = col
                    print(f"   ✅ product_name trouvé: '{col}' -> échantillon: {catalog_df[col].head(3).tolist()}")

                # Chercher is_active
                if col_is_active is None and 'is_active' in col_lower:
                    col_is_active = col
                    print(f"   ✅ is_active trouvé: '{col}'")

            # Si on a trouvé les colonnes, construire le mapping
            if col_product_id and col_product_name:
                print(f"   🎯 SUCCÈS avec skiprows={skip}")

                catalog_subset = pd.DataFrame()
                catalog_subset['product_id'] = pd.to_numeric(catalog_df[col_product_id], errors='coerce').fillna(0).astype(int)
                catalog_subset['product_name'] = catalog_df[col_product_name].astype(str).str.lower().str.strip()

                if col_is_active:
                    catalog_subset['is_active'] = (
                        catalog_df[col_is_active]
                        .astype(str)
                        .str.upper()
                        .str.strip()
                        .isin(['TRUE', '1', 'YES', 'OUI'])
                    )
                else:
                    catalog_subset['is_active'] = True

                product_map = catalog_subset.drop_duplicates(subset=['product_name'])

                # Compter actifs/inactifs
                active_count = product_map['is_active'].sum()
                inactive_count = len(product_map) - active_count

                print(f"\n✅ Catalogue traité (skiprows={skip}):")
                print(f"   - Total produits : {len(product_map)}")
                print(f"   - Avec product_id > 0 : {(product_map['product_id'] > 0).sum()}")
                print(f"   - ✅ Actifs (is_active=True) : {active_count}")
                print(f"   - ❌ Inactifs (is_active=False) : {inactive_count}")
                print(f"   - Échantillon:")
                print(product_map.head(3).to_string())

                product_id_map = product_map
                break  # Sortir de la boucle, on a trouvé
            else:
                print(f"   ❌ skiprows={skip} - colonnes manquantes (id={col_product_id}, name={col_product_name})")

        if product_id_map.empty:
            print(f"\n⚠️ Catalogue non chargé - product_id viendra de l'inventaire uniquement")

    except Exception as e:
        print(f"❌ Erreur chargement catalogue : {e}")
        import traceback
        print(traceback.format_exc())
        product_id_map = pd.DataFrame(columns=['product_name', 'product_id', 'is_active'])
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
    # Total stock - ✅ GARDER TOUS LES PRODUITS SANS REDONDANCE
    # =========================

    print(f"\n📊 Inventaire brut: {len(inventory_pikine_staging_df)} lignes")

    # Normaliser product_name pour éviter les doublons dûs aux espaces/casse
    inventory_pikine_staging_df['product_name'] = (
        inventory_pikine_staging_df['product_name']
        .astype(str)
        .str.strip()
        .str.lower()
    )

    # ✅ GARDER product_id de l'inventaire dans le groupby
    # L'inventaire a déjà des product_id valides qu'on ne doit pas perdre
    if 'product_id' in inventory_pikine_staging_df.columns:
        total_stock_df = (
            inventory_pikine_staging_df.groupby(["product_name", "Supplier"], as_index=False)
            .agg({
                "total_stock": "sum",
                "product_id": "first"  # Garder le premier product_id
            })
        )
        # Convertir product_id en int
        total_stock_df['product_id'] = pd.to_numeric(total_stock_df['product_id'], errors='coerce').fillna(0).astype(int)
        inv_valid = (total_stock_df['product_id'] > 0).sum()
        print(f"   ✅ Après groupby: {len(total_stock_df)} produits, {inv_valid} avec product_id de l'inventaire")
    else:
        total_stock_df = (
            inventory_pikine_staging_df.groupby(["product_name", "Supplier"], as_index=False)["total_stock"]
            .sum()
        )
        total_stock_df['product_id'] = 0
        print(f"   ✅ Après groupby: {len(total_stock_df)} produits (pas de product_id dans inventaire)")

    print(f"Shape of total_stock_df: {total_stock_df.shape}")

    # ✅ MERGER product_id ET is_active depuis le catalogue (pour compléter les manquants)

    if not product_id_map.empty:
        # S'assurer que total_stock_df a product_name normalisé
        total_stock_df['product_name'] = total_stock_df['product_name'].astype(str).str.lower().str.strip()

        # Préparer le mapping
        merge_map = product_id_map[['product_name', 'product_id', 'is_active']].copy()
        merge_map.columns = ['product_name', 'catalog_product_id', 'is_active']
        merge_map['product_name'] = merge_map['product_name'].astype(str).str.lower().str.strip()

        # ✅ DEBUG DÉTAILLÉ: Voir pourquoi le merge échoue
        print(f"\n🔍 DEBUG MERGE:")
        print(f"   Inventaire: {len(total_stock_df)} produits")
        print(f"   Catalogue: {len(merge_map)} produits")
        print(f"   product_id déjà valides (inventaire): {(total_stock_df['product_id'] > 0).sum()}")

        # Vérifier combien de noms correspondent
        inv_names = set(total_stock_df['product_name'].unique())
        cat_names = set(merge_map['product_name'].unique())
        common = inv_names & cat_names
        only_inv = inv_names - cat_names
        only_cat = cat_names - inv_names

        print(f"   Noms communs: {len(common)}")
        print(f"   Seulement dans inventaire: {len(only_inv)}")
        print(f"   Seulement dans catalogue: {len(only_cat)}")

        if len(common) == 0:
            print(f"\n   ⚠️ AUCUN NOM EN COMMUN!")
            print(f"   Exemples inventaire: {list(inv_names)[:5]}")
            print(f"   Exemples catalogue: {list(cat_names)[:5]}")
        elif len(common) < len(inv_names) * 0.5:
            print(f"\n   ⚠️ MOINS DE 50% DE CORRESPONDANCE!")
            print(f"   Non trouvés: {list(only_inv)[:5]}")

        # Merge pour récupérer catalog_product_id et is_active
        total_stock_df = total_stock_df.merge(
            merge_map,
            on="product_name",
            how="left"
        )

        # ✅ PRIORITÉ: Garder product_id de l'inventaire, compléter avec catalogue si manquant
        if 'catalog_product_id' in total_stock_df.columns:
            # Si product_id inventaire = 0, utiliser catalog_product_id
            mask_missing = total_stock_df['product_id'] == 0
            total_stock_df.loc[mask_missing, 'product_id'] = total_stock_df.loc[mask_missing, 'catalog_product_id']
            total_stock_df.drop(columns=['catalog_product_id'], errors='ignore', inplace=True)

        # Remplir les valeurs manquantes
        total_stock_df['product_id'] = pd.to_numeric(total_stock_df['product_id'], errors='coerce').fillna(0).astype(int)
        total_stock_df['is_active'] = total_stock_df['is_active'].fillna(True)

        print(f"\n✅ Merge catalogue effectué:")
        print(f"   - product_id valides: {(total_stock_df['product_id'] > 0).sum()}/{len(total_stock_df)}")
        print(f"   - Produits actifs: {total_stock_df['is_active'].sum()}")

        # Échantillon des résultats
        sample = total_stock_df[['product_name', 'product_id']].head(5)
        print(f"   Échantillon résultat:")
        print(sample.to_string())
    else:
        total_stock_df['product_id'] = 0
        total_stock_df['is_active'] = True

    # ✅ S'assurer que product_id existe et est propre
    if 'product_id' not in total_stock_df.columns:
        total_stock_df['product_id'] = 0
    total_stock_df['product_id'] = pd.to_numeric(total_stock_df['product_id'], errors='coerce').fillna(0).astype(int)

    print(f"✅ État final total_stock_df:")
    print(f"   - Colonnes: {total_stock_df.columns.tolist()}")
    print(f"   - Produits: {len(total_stock_df)}")
    print(f"   - product_id valides: {(total_stock_df['product_id'] > 0).sum()}")

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

    # ✅ DEBUG: Vérifier product_id après premier merge
    if 'product_id' in final_stock_sales_df.columns:
        print(f"✅ product_id présent dans final_stock_sales_df: {(final_stock_sales_df['product_id'] > 0).sum()} valeurs valides")
    else:
        print(f"⚠️ product_id ABSENT de final_stock_sales_df - colonnes: {final_stock_sales_df.columns.tolist()[:8]}")

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
    # SUPPLIER CATEGORIZATION → UPDATE PRODUCT CATEGORY
    # =========================
    # SUPPLIER_CATEGORIZATION_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQax3ZQW2QhDLE-waDewtdD8x_Q5tpn2FWzVJftr9egik4_JF3s2ytSYJmXh55aUnp79vmF-XtkaTmN/pub?gid=1938047484&single=true&output=csv"

    try:
        supplier_categorization_df = pd.read_csv(SUPPLIER_CATEGORIZATION_URL)
        print("'Supplier Categorization' data loaded successfully.")
    except Exception as e:
        print(f"Error loading 'Supplier Categorization' data: {e}")
        supplier_categorization_df = pd.DataFrame()

    supplier_categorization_lookup_df = pd.DataFrame()
    if not supplier_categorization_df.empty and supplier_categorization_df.shape[1] > 7:
        supplier_categorization_lookup_df = supplier_categorization_df.iloc[:, [0, 7]].copy()
        supplier_categorization_lookup_df.columns = ['Supplier_Name_Lookup', 'Supplier_Categorization_Lookup']
        supplier_categorization_lookup_df['Supplier_Name_Lookup'] = (
            supplier_categorization_lookup_df['Supplier_Name_Lookup'].astype(str).str.lower().str.strip()
        )
    else:
        print("Warning: 'Supplier Categorization' data does not have enough columns.")

    # Merge
    if not supplier_categorization_lookup_df.empty and 'Supplier' in final_stock_sales_df.columns:
        final_stock_sales_df['Supplier'] = final_stock_sales_df['Supplier'].astype(str).str.lower().str.strip()

        if 'Supplier_Categorization_Lookup' in final_stock_sales_df.columns:
            final_stock_sales_df.drop(columns=['Supplier_Categorization_Lookup'], inplace=True)

        final_stock_sales_df = pd.merge(
            final_stock_sales_df,
            supplier_categorization_lookup_df.drop_duplicates(subset=['Supplier_Name_Lookup']),
            left_on='Supplier',
            right_on='Supplier_Name_Lookup',
            how='left'
        )

        if 'Supplier_Name_Lookup' in final_stock_sales_df.columns:
            final_stock_sales_df.drop(columns=['Supplier_Name_Lookup'], inplace=True)

        final_stock_sales_df['Supplier_Categorization_Lookup'] = (
            final_stock_sales_df['Supplier_Categorization_Lookup'].fillna('Unknown')
        )

        # RÈGLE : Si credit_days > 0 → supplier categorization, sinon ABC-XYZ
        if all(col in final_stock_sales_df.columns for col in
               ['credit_days', 'Product Category', 'Supplier_Categorization_Lookup']):

            final_stock_sales_df['credit_days'] = pd.to_numeric(
                final_stock_sales_df['credit_days'], errors='coerce'
            ).fillna(0)

            final_stock_sales_df['Product Category'] = final_stock_sales_df.apply(
                lambda row: row['Supplier_Categorization_Lookup'] if row['credit_days'] > 0 else row[
                    'Product Category'],
                axis=1
            )
            print("'Product Category' column updated based on supplier categorization and credit_days.")
        else:
            print("Warning: Required columns for updating 'Product Category' are missing.")

        # Nettoyer colonne temporaire
        if 'Supplier_Categorization_Lookup' in final_stock_sales_df.columns:
            final_stock_sales_df.drop(columns=['Supplier_Categorization_Lookup'], inplace=True)
    else:
        print("Warning: Cannot update 'Product Category' - missing data.")

    # =========================
    # ADS 7d / ADS 30d + OOS 7d / OOS 30d - MERGE UNIQUE
    # =========================
    # =========================
    # ✅ ADS MULTI-PÉRIODES - LOGIQUE ORIGINALE PRÉSERVÉE
    # =========================
    print(f"\n{'=' * 60}")
    print(f"📊 CALCUL AVERAGE DAILY SALES - Période : {period_days}")
    print(f"{'=' * 60}")

    # ✅ ÉTAPE 1 : Merger ADS 7j et 30j (VOTRE CODE ORIGINAL)
    if {"2", "9"}.issubset(Tbh_7dsales_df.columns):
        ads7 = (
            Tbh_7dsales_df[["2", "9"]]
            .rename(columns={"2": "product_name", "9": "Average Daily Sales (7d)"})
            .drop_duplicates(subset=["product_name"])
        )
        ads7["product_name"] = ads7["product_name"].astype(str).str.lower().str.strip()
        ads7["Average Daily Sales (7d)"] = safe_numeric(ads7["Average Daily Sales (7d)"], 0)
    else:
        ads7 = pd.DataFrame(columns=["product_name", "Average Daily Sales (7d)"])

    if {"2", "9"}.issubset(Tbh_30dsales_products_df.columns):
        ads30 = (
            Tbh_30dsales_products_df[["2", "9"]]
            .rename(columns={"2": "product_name", "9": "Average Daily Sales (30d)"})
            .drop_duplicates(subset=["product_name"])
        )
        ads30["product_name"] = ads30["product_name"].astype(str).str.lower().str.strip()
        ads30["Average Daily Sales (30d)"] = safe_numeric(ads30["Average Daily Sales (30d)"], 0)
    else:
        ads30 = pd.DataFrame(columns=["product_name", "Average Daily Sales (30d)"])

    # ✅ Merger OOS Rates (VOTRE CODE ORIGINAL)
    if {"2", "12"}.issubset(Tbh_7dsales_df.columns):
        oos7 = (
            Tbh_7dsales_df[["2", "12"]]
            .rename(columns={"2": "product_name", "12": "Daily OOS Rate (7d)"})
            .drop_duplicates(subset=["product_name"])
        )
        oos7["product_name"] = oos7["product_name"].astype(str).str.lower().str.strip()
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
        oos30["product_name"] = oos30["product_name"].astype(str).str.lower().str.strip()
        oos30["Daily OOS Rate (30d)"] = safe_numeric(
            oos30["Daily OOS Rate (30d)"].astype(str).str.replace("%", "", regex=False), 0
        ) / 100.0
    else:
        oos30 = pd.DataFrame(columns=["product_name", "Daily OOS Rate (30d)"])

    # ✅ MERGE UNIQUE (VOTRE CODE ORIGINAL)
    final_stock_sales_df = final_stock_sales_df.merge(ads7, on="product_name", how="left", validate="m:1")
    final_stock_sales_df = final_stock_sales_df.merge(ads30, on="product_name", how="left", validate="m:1")
    final_stock_sales_df = final_stock_sales_df.merge(oos7, on="product_name", how="left", validate="m:1")
    final_stock_sales_df = final_stock_sales_df.merge(oos30, on="product_name", how="left", validate="m:1")

    print(f"   ✅ ADS 7j mergée : {(final_stock_sales_df['Average Daily Sales (7d)'].notna()).sum()} produits")
    print(f"   ✅ ADS 30j mergée : {(final_stock_sales_df['Average Daily Sales (30d)'].notna()).sum()} produits")

    # =========================
    # ✅ NOUVEAU : Calcul ADS selon période sélectionnée
    # =========================
    period_map = {"3d": 3, "7d": 7, "30d": 30}
    n_days = period_map.get(period_days, 7)

    print(f"\n🎯 Application période {period_days} ({n_days} jours)...")

    if period_days == "3d":
        # ✅ POUR 3J : Tenter calcul depuis historique, sinon utiliser ADS 7j
        try:
            # ✅ SIMPLIFIÉ: Utiliser ADS 7j directement au lieu de recalculer depuis l'historique
            # Cela évite un rechargement des données qui peut bloquer
            print(f"   → Utilisation de ADS 7j comme base pour période 3j")

            final_stock_sales_df['Average Daily Sales'] = (
                final_stock_sales_df.get('Average Daily Sales (7d)', pd.Series([0.1] * len(final_stock_sales_df)))
                .fillna(0.1)
                .clip(lower=0.1)
            )
            print(f"   ✅ ADS 3j = ADS 7j (optimisé)")

        except Exception as e:
            print(f"   ⚠️ Erreur ADS 3j : {e}")
            print(f"   → Fallback : ADS 3j = ADS 7j")

            # Fallback : utiliser ADS 7j
            final_stock_sales_df['Average Daily Sales'] = (
                final_stock_sales_df.get('Average Daily Sales (7d)', pd.Series([0.1] * len(final_stock_sales_df)))
                .fillna(0.1)
                .clip(lower=0.1)
            )

    elif period_days == "7d":
        # ✅ POUR 7J : VOTRE LOGIQUE ORIGINALE calc_recalculated_ads
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

        # ✅ Renommer immédiatement
        final_stock_sales_df.rename(
            columns={"Recalculated Average Daily Sales": "Average Daily Sales"},
            inplace=True
        )

        print(f"   ✅ ADS 7j calculée (logique originale)")

    elif period_days == "30d":
        # ✅ POUR 30J : Privilégier ADS 30j, sinon moyenne 7j/30j
        def calc_ads_30d(row):
            sales_30d = row.get("Average Daily Sales (30d)", np.nan)
            sales_7d = row.get("Average Daily Sales (7d)", np.nan)
            oos_30d = row.get("Daily OOS Rate (30d)", 1.0)

            if pd.notna(sales_30d):
                # Si OOS élevé, pondérer avec 7j
                if pd.notna(oos_30d) and (oos_30d >= 0.6) and pd.notna(sales_7d):
                    return (sales_30d + sales_7d) / 2.0 + 0.1
                return sales_30d + 0.1
            elif pd.notna(sales_7d):
                return sales_7d + 0.1
            else:
                return 0.1

        final_stock_sales_df['Average Daily Sales'] = (
            final_stock_sales_df.apply(calc_ads_30d, axis=1)
        )

        print(f"   ✅ ADS 30j calculée")

    # =========================
    # ✅ GARANTIE FINALE (au cas où)
    # =========================
    if 'Average Daily Sales' not in final_stock_sales_df.columns:
        print("   ⚠️ ERREUR : Colonne ADS manquante, création par défaut")
        final_stock_sales_df['Average Daily Sales'] = 0.1

    # ✅ Validation et nettoyage
    final_stock_sales_df['Average Daily Sales'] = (
        pd.to_numeric(final_stock_sales_df['Average Daily Sales'], errors='coerce')
        .fillna(0.1)
        .clip(lower=0.1)
    )

    # ✅ Statistiques
    print(f"\n✅ Average Daily Sales {period_days} - Stats finales :")
    print(f"   Min     : {final_stock_sales_df['Average Daily Sales'].min():.2f}")
    print(f"   Max     : {final_stock_sales_df['Average Daily Sales'].max():.2f}")
    print(f"   Médiane : {final_stock_sales_df['Average Daily Sales'].median():.2f}")
    print(f"   Moyenne : {final_stock_sales_df['Average Daily Sales'].mean():.2f}")
    print(f"{'=' * 60}\n")

    # =========================
    # ✅ Max Coverage Day (VOTRE CODE ORIGINAL)
    # =========================
    final_stock_sales_df["Max Coverage Day"] = np.minimum(
        safe_numeric(final_stock_sales_df["total_stock"], 0) /
        np.maximum(safe_numeric(final_stock_sales_df["Average Daily Sales"], 0.01), 0.01),
        365
    )

    print(f"✅ Max Coverage Day recalculé")

    # =========================
    # Lead time (suite de votre code EXACT)
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
    '''if {"2", "9"}.issubset(Tbh_7dsales_df.columns):
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

    # ✅ Renommer immédiatement après le calcul
    final_stock_sales_df.rename(columns={"Recalculated Average Daily Sales": "Average Daily Sales"}, inplace=True)

    print(f"✅ Average Daily Sales - Stats:")
    print(f"   Min: {final_stock_sales_df['Average Daily Sales'].min():.2f}")
    print(f"   Max: {final_stock_sales_df['Average Daily Sales'].max():.2f}")
    print(f"   Médiane: {final_stock_sales_df['Average Daily Sales'].median():.2f}")
    print(f"   Moyenne: {final_stock_sales_df['Average Daily Sales'].mean():.2f}")
    # Max Coverage Day
    final_stock_sales_df["Max Coverage Day"] = np.minimum(
        safe_numeric(final_stock_sales_df["total_stock"], 0) /
        np.maximum(safe_numeric(final_stock_sales_df["Average Daily Sales"], 0.01), 0.01),
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
'''
    # =========================
    # ✅ DÉPLACER ICI : Supplier credit info (AVANT supplier categorization)
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
        # SUPPLIER_CATEGORIZATION_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQax3ZQW2QhDLE-waDewtdD8x_Q5tpn2FWzVJftr9egik4_JF3s2ytSYJmXh55aUnp79vmF-XtkaTmN/pub?gid=1938047484&single=true&output=csv"

        try:
            supplier_categorization_df = pd.read_csv(SUPPLIER_CATEGORIZATION_URL)
            print("'Supplier Categorization' data loaded successfully.")
        except Exception as e:
            print(f"Error loading 'Supplier Categorization' data: {e}")
            supplier_categorization_df = pd.DataFrame()

        supplier_categorization_lookup_df = pd.DataFrame()
        if not supplier_categorization_df.empty and supplier_categorization_df.shape[1] > 7:
            supplier_categorization_lookup_df = supplier_categorization_df.iloc[:, [0, 7]].copy()
            supplier_categorization_lookup_df.columns = ['Supplier_Name_Lookup', 'Supplier_Categorization_Lookup']
            supplier_categorization_lookup_df['Supplier_Name_Lookup'] = (
                supplier_categorization_lookup_df['Supplier_Name_Lookup'].astype(str).str.lower().str.strip()
            )
        else:
            print("Warning: 'Supplier Categorization' data does not have enough columns.")

        # Merge
        if not supplier_categorization_lookup_df.empty and 'Supplier' in final_stock_sales_df.columns:
            final_stock_sales_df['Supplier'] = final_stock_sales_df['Supplier'].astype(str).str.lower().str.strip()

            if 'Supplier_Categorization_Lookup' in final_stock_sales_df.columns:
                final_stock_sales_df.drop(columns=['Supplier_Categorization_Lookup'], inplace=True)

            final_stock_sales_df = pd.merge(
                final_stock_sales_df,
                supplier_categorization_lookup_df.drop_duplicates(subset=['Supplier_Name_Lookup']),
                left_on='Supplier',
                right_on='Supplier_Name_Lookup',
                how='left'
            )

            if 'Supplier_Name_Lookup' in final_stock_sales_df.columns:
                final_stock_sales_df.drop(columns=['Supplier_Name_Lookup'], inplace=True)

            final_stock_sales_df['Supplier_Categorization_Lookup'] = (
                final_stock_sales_df['Supplier_Categorization_Lookup'].fillna('Unknown')
            )

            # RÈGLE : Si credit_days > 0 → supplier categorization, sinon ABC-XYZ
            if all(col in final_stock_sales_df.columns for col in
                   ['credit_days', 'Product Category', 'Supplier_Categorization_Lookup']):

                final_stock_sales_df['credit_days'] = pd.to_numeric(
                    final_stock_sales_df['credit_days'], errors='coerce'
                ).fillna(0)

                final_stock_sales_df['Product Category'] = final_stock_sales_df.apply(
                    lambda row: row['Supplier_Categorization_Lookup'] if row['credit_days'] > 0 else row[
                        'Product Category'],
                    axis=1
                )
                print("'Product Category' column updated based on supplier categorization and credit_days.")
            else:
                print("Warning: Required columns for updating 'Product Category' are missing.")

            # Nettoyer colonne temporaire
            if 'Supplier_Categorization_Lookup' in final_stock_sales_df.columns:
                final_stock_sales_df.drop(columns=['Supplier_Categorization_Lookup'], inplace=True)
        else:
            print("Warning: Cannot update 'Product Category' - missing data.")
    # =========================
    # Parametres Replenish (Buffer lookup) - APRÈS update Product Category
    # =========================
    if not parametres_replenish_df.empty and parametres_replenish_df.shape[1] > 1:
        pr_buf = parametres_replenish_df.iloc[:, [0, 1]].copy()
        pr_buf.columns = ["Param_Product_Category", "Param_Buffer_Value_Lookup"]
        pr_buf["Param_Product_Category"] = pr_buf["Param_Product_Category"].astype(str).str.lower().str.strip()
        pr_buf["Param_Buffer_Value_Lookup"] = safe_numeric(pr_buf["Param_Buffer_Value_Lookup"], 0)

        # ✅ IMPORTANT : Normaliser Product Category AVANT le merge
        final_stock_sales_df["Product Category"] = final_stock_sales_df["Product Category"].astype(
            str).str.lower().str.strip()

        final_stock_sales_df = final_stock_sales_df.merge(
            pr_buf.drop_duplicates(subset=["Param_Product_Category"]),
            left_on="Product Category",  # ✅ Utilise bien la Product Category mise à jour
            right_on="Param_Product_Category",
            how="left",
            validate="m:1",
        )
        if "Param_Product_Category" in final_stock_sales_df.columns:
            final_stock_sales_df.drop(columns=["Param_Product_Category"], inplace=True)
        final_stock_sales_df["Param_Buffer_Value_Lookup"] = final_stock_sales_df["Param_Buffer_Value_Lookup"].fillna(0)
        print("✅ 'Param_Buffer_Value_Lookup' mergé basé sur la Product Category mise à jour")
    else:
        final_stock_sales_df["Param_Buffer_Value_Lookup"] = 0.0

    # 1. Calculer MAX_CREDIT_BUFFER d'abord
    # =========================
    # Calcul des buffers (ORDRE IMPORTANT)
    # =========================

    # 1. Calculer MAX_CREDIT_BUFFER d'abord
    def calc_max_credit_buffer(row):
        cd = row["credit_days"]
        cc = str(row.get("Credit_cumulable", "")).lower()
        return min(cd, 20) if cc == "oui" else cd

    final_stock_sales_df["MAX_CREDIT_BUFFER"] = final_stock_sales_df.apply(calc_max_credit_buffer, axis=1)

    # 2. Calculer AJUSTER_BUFFER (max entre Param_Buffer_Value_Lookup et MAX_CREDIT_BUFFER)
    final_stock_sales_df["AJUSTER_BUFFER"] = np.maximum(
        safe_numeric(final_stock_sales_df["Param_Buffer_Value_Lookup"], 0),
        safe_numeric(final_stock_sales_df["MAX_CREDIT_BUFFER"], 0),
    )

    print("✅ 'MAX_CREDIT_BUFFER' et 'AJUSTER_BUFFER' calculés")
    print(f"   - AJUSTER_BUFFER moyen : {final_stock_sales_df['AJUSTER_BUFFER'].mean():.2f} jours")
    print(f"   - AJUSTER_BUFFER max : {final_stock_sales_df['AJUSTER_BUFFER'].max():.0f} jours")

    # 3. Calculer optimal stock (utilise AJUSTER_BUFFER)
    final_stock_sales_df["Max Daily Sales (Pikine)"] = safe_numeric(
        final_stock_sales_df["Max Daily Sales (Pikine)"], 0
    )
    final_stock_sales_df["Average Daily Sales"] = safe_numeric(
        final_stock_sales_df["Average Daily Sales"], 0.1
    )

    final_stock_sales_df["optimal stock (Reorder Point)"] = final_stock_sales_df.apply(
        lambda r: max(
            r["Max Daily Sales (Pikine)"],
            (r["Max Daily Sales (Pikine)"] / 2.0) + (r["AJUSTER_BUFFER"] * r["Average Daily Sales"])
        ),
        axis=1,
    )

    # Renommer pour cohérence
    final_stock_sales_df.rename(columns={"optimal stock (Reorder Point)": "optimal stock"}, inplace=True)
    print("✅ 'optimal stock' calculé avec AJUSTER_BUFFER")
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
    # ✅ NOUVEAU: Merge des dernières réceptions
    # =========================
    if not last_receptions_df.empty and 'product_name' in last_receptions_df.columns:
        print("\n📦 Fusion des données de réception...")

        # Préparer les données de réception
        receptions_cols = ['product_name']
        if 'last_reception_qty' in last_receptions_df.columns:
            receptions_cols.append('last_reception_qty')
        if 'last_reception_date' in last_receptions_df.columns:
            receptions_cols.append('last_reception_date')

        # Ajouter toutes les colonnes pertinentes du dataclip
        for col in last_receptions_df.columns:
            if col not in receptions_cols and col != 'product_name':
                receptions_cols.append(col)

        receptions_to_merge = last_receptions_df[receptions_cols].drop_duplicates(subset=['product_name'], keep='first')

        # Harmoniser product_name pour le merge
        receptions_to_merge['product_name'] = receptions_to_merge['product_name'].astype(str).str.strip().str.lower()
        final_stock_sales_df['_merge_key'] = final_stock_sales_df['product_name'].astype(str).str.strip().str.lower()

        before_merge = len(final_stock_sales_df)
        final_stock_sales_df = final_stock_sales_df.merge(
            receptions_to_merge.rename(columns={'product_name': '_merge_key'}),
            on='_merge_key',
            how='left',
            validate='m:1'
        )
        final_stock_sales_df.drop(columns=['_merge_key'], inplace=True)

        # Formater la date de réception si présente
        if 'last_reception_date' in final_stock_sales_df.columns:
            try:
                final_stock_sales_df['last_reception_date'] = pd.to_datetime(
                    final_stock_sales_df['last_reception_date'],
                    errors='coerce'
                )
                # Calculer les jours depuis dernière réception
                final_stock_sales_df['days_since_reception'] = (
                        pd.Timestamp.now() - final_stock_sales_df['last_reception_date']
                ).dt.days
                final_stock_sales_df['last_reception_date'] = final_stock_sales_df['last_reception_date'].dt.strftime('%d/%m/%Y')
            except Exception as e:
                print(f"   ⚠️ Erreur formatage date: {e}")

        # Formater la quantité reçue
        if 'last_reception_qty' in final_stock_sales_df.columns:
            final_stock_sales_df['last_reception_qty'] = safe_numeric(
                final_stock_sales_df['last_reception_qty'], 0
            ).astype(int)

        matched = final_stock_sales_df['last_reception_qty'].notna().sum() if 'last_reception_qty' in final_stock_sales_df.columns else 0
        print(f"   ✅ {matched}/{before_merge} produits avec données de réception")
    else:
        # ℹ️ Les données de réception ne sont pas disponibles (timeout ou erreur)
        print("   ℹ️ Données de réception non disponibles - colonnes initialisées à 0")
        final_stock_sales_df['last_reception_qty'] = 0
        final_stock_sales_df['last_reception_date'] = None
        final_stock_sales_df['days_since_reception'] = None

    # =========================
    # Dédupes
    # =========================
    final_stock_sales_df = final_stock_sales_df.drop_duplicates(subset=["product_name"], keep="first")
    print(f"Shape after removing duplicates: {final_stock_sales_df.shape}")

    # ✅ CORRECTION: Renommer product_id_x en product_id si nécessaire
    # Cela arrive quand un merge crée un conflit de noms
    if 'product_id_x' in final_stock_sales_df.columns and 'product_id' not in final_stock_sales_df.columns:
        final_stock_sales_df.rename(columns={'product_id_x': 'product_id'}, inplace=True)
        print(f"   ✅ product_id_x renommé en product_id")

    # Supprimer product_id_y si présent
    if 'product_id_y' in final_stock_sales_df.columns:
        final_stock_sales_df.drop(columns=['product_id_y'], inplace=True)

    # ✅ DEBUG: Vérifier product_id après déduplication
    if 'product_id' in final_stock_sales_df.columns:
        final_stock_sales_df['product_id'] = pd.to_numeric(final_stock_sales_df['product_id'], errors='coerce').fillna(0).astype(int)
        valid = (final_stock_sales_df['product_id'] > 0).sum()
        print(f"   🔍 product_id après dédup: {valid}/{len(final_stock_sales_df)} valides")
    else:
        # Dernier recours : créer product_id à 0
        print(f"   ⚠️ product_id toujours absent après corrections! Colonnes: {final_stock_sales_df.columns.tolist()[:8]}...")
        final_stock_sales_df['product_id'] = 0

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

    # def calc_max_credit_buffer(row):
    #   cd = row["credit_days"]
    #  cc = str(row.get("Credit_cumulable", "")).lower()
    # return min(cd, 20) if cc == "oui" else cd

    # final_stock_sales_df["MAX_CREDIT_BUFFER"] = final_stock_sales_df.apply(calc_max_credit_buffer, axis=1)

    # final_stock_sales_df["AJUSTER_BUFFER"] = np.maximum(
    #   safe_numeric(final_stock_sales_df["Param_Buffer_Value_Lookup"], 0),
    #  safe_numeric(final_stock_sales_df["MAX_CREDIT_BUFFER"], 0),
    # )

    final_stock_sales_df["MOQ MAAD"] = (safe_numeric(final_stock_sales_df["ADJUSTED_LEADTIME"], 0) + 3) * safe_numeric(
        final_stock_sales_df["Average Daily Sales"], 0.1)

    if "Param_Supplier_Factor" not in final_stock_sales_df.columns:
        final_stock_sales_df["Param_Supplier_Factor"] = 0.0

    # Purchase Need
    req_cols = ["total_stock", "Max Daily Sales (Pikine)", "Average Daily Sales",
                "optimal stock", "Param_Buffer_Value_Lookup", "Param_Supplier_Factor"]
    if all(c in final_stock_sales_df.columns for c in req_cols):
        for c in req_cols:
            final_stock_sales_df[c] = safe_numeric(final_stock_sales_df[c], 0)

        def calculate_purchase_need(row):
            total_stock = row["total_stock"]
            max_daily_sales = row["Max Daily Sales (Pikine)"]
            avg_daily_sales = row["Average Daily Sales"]
            optimal_stock = row["optimal stock"]
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

    # =========================
    # Calcul target_quantity (formule simple)
    # =========================
    # final_stock_sales_df['target_quantity'] = np.maximum(
    #   0,
    ##  final_stock_sales_df['Average Daily Sales'] *
    # (final_stock_sales_df['ADJUSTED_LEADTIME'] + final_stock_sales_df['credit_days']) +
    # final_stock_sales_df['AJUSTER_BUFFER'] * final_stock_sales_df['Average Daily Sales'] -
    # final_stock_sales_df['total_stock']
    # )

    # print(f"✅ target_quantity calculée (formule métier)")

    # =========================
    # ML SUPERVISÉ avec ventes historiques
    # =========================
    print("\nChargement ventes historiques Pikine...")
    sales_history = load_sales_history()

    if not sales_history.empty:
        print("Entraînement modèle supervisé...")

        # ✅ DEBUG: Vérifier product_id AVANT ML
        before_ml_valid = (final_stock_sales_df['product_id'] > 0).sum() if 'product_id' in final_stock_sales_df.columns else 'ABSENT'
        print(f"   🔍 product_id AVANT ML: {before_ml_valid}")

        final_stock_sales_df, ml_model = create_supervised_target_from_sales(
            sales_history,
            final_stock_sales_df
        )

        # ✅ DEBUG: Vérifier product_id APRÈS ML
        after_ml_valid = (final_stock_sales_df['product_id'] > 0).sum() if 'product_id' in final_stock_sales_df.columns else 'ABSENT'
        print(f"   🔍 product_id APRÈS ML: {after_ml_valid}")
        if after_ml_valid == 'ABSENT' and before_ml_valid != 'ABSENT':
            print(f"   ⚠️⚠️⚠️ product_id PERDU dans create_supervised_target_from_sales! ⚠️⚠️⚠️")
            print(f"   Colonnes retournées: {final_stock_sales_df.columns.tolist()[:10]}...")
    else:
        print("Pas d'historique, utilisation formule")
        final_stock_sales_df['target_quantity'] = calculate_formula_based_target(
            final_stock_sales_df
        )

    if 'target_quantity' in final_stock_sales_df.columns:
        print(f"✅ target_quantity finalisée :")
        print(f"   Moyenne : {final_stock_sales_df['target_quantity'].mean():.0f}")
        print(f"   Médiane : {final_stock_sales_df['target_quantity'].median():.0f}")

    # ✅ PROMOTIONS avec analyse ROI
    # =========================
    # ✅ PROMOTIONS avec analyse ROI (CORRECTION)
    # =========================
    print("\nChargement et analyse des promotions...")

    # ✅ Initialiser promo_roi AVANT de l'utiliser
    promo_roi = pd.DataFrame()
    promo_history = pd.DataFrame()

    # ✅ DEBUG: Vérifier product_id avant l'appel à calculate_promo_roi_analysis
    print(f"🔍 DEBUG avant promo: colonnes={final_stock_sales_df.columns.tolist()[:10]}...")
    print(f"   product_id présent: {'product_id' in final_stock_sales_df.columns}")
    if 'product_id' in final_stock_sales_df.columns:
        print(f"   product_id valides: {(final_stock_sales_df['product_id'] > 0).sum()}")

    try:
        promo_history = load_and_analyze_promotions()

        if not sales_history.empty and not promo_history.empty:
            print("Calcul ROI des promotions...")

            # ✅ S'assurer que product_id existe avant l'appel
            if 'product_id' not in final_stock_sales_df.columns:
                print("⚠️ product_id absent - création temporaire")
                final_stock_sales_df['product_id'] = 0

            promo_roi = calculate_promo_roi_analysis(
                sales_history,
                promo_history,
                final_stock_sales_df
            )

            if not promo_roi.empty:
                # Merger avec données principales
                final_stock_sales_df = final_stock_sales_df.merge(
                    promo_roi[[
                        'product_name',
                        'uplift_pct',
                        'roi_pct',
                        'net_profit_per_day',
                        'promo_recommendation',
                        'promo_priority'
                    ]],
                    on='product_name',
                    how='left'
                )

                # Statut promo actuel
                promo_active = promo_history[promo_history['promo_status'] == 'Active'][
                    ['product_name', 'promo_status', 'days_remaining', 'discount_pct']
                ].drop_duplicates()

                final_stock_sales_df = final_stock_sales_df.merge(
                    promo_active,
                    on='product_name',
                    how='left'
                )

                final_stock_sales_df['promo_status'] = final_stock_sales_df['promo_status'].fillna('Pas de promo')

                print(f"✅ Analyse ROI promotions intégrée")

                # Debug : vérifier colonnes créées
                print(f"\n✅ Colonnes promo ajoutées :")
                for col in ['uplift_pct', 'roi_pct', 'promo_recommendation', 'promo_priority']:
                    if col in final_stock_sales_df.columns:
                        print(f"   ✓ {col}")
                    else:
                        print(f"   ✗ {col} MANQUANTE")
            else:
                print("⚠️ Analyse ROI vide, pas de données promo")
        else:
            print("⚠️ Pas d'historique ventes ou promo disponible")

    except Exception as e:
        print(f"⚠️ Erreur analyse promotions : {e}")
        import traceback
        traceback.print_exc()
        # Continuer sans les promos

        # Ajusted_total_need

    def calc_adjusted_total_need(row):
        if str(row.get("delisting_status", "")).lower() == "delisted":
            return "NO NEED"
        mcd = float(row.get("Max Coverage Day", 0))
        alt = float(row.get("ADJUSTED_LEADTIME", 0))
        opt = float(row.get("optimal stock", 0))
        ads = float(row.get("Average Daily Sales", 0.1))
        optimal_days = (opt / ads) if ads > 0 else 0
        if mcd <= alt + 3:
            return "ORDER NOW"
        elif mcd < alt + optimal_days:
            return "ORDER NOT URGENT"
        else:
            return "NO NEED"

    final_stock_sales_df["Ajusted_total_need"] = final_stock_sales_df.apply(calc_adjusted_total_need, axis=1)

    # Predicted Stockout
    final_stock_sales_df["Predicted Stockout"] = safe_numeric(final_stock_sales_df["total_stock"], 0) <= safe_numeric(
        final_stock_sales_df["optimal stock"], 0)

    # Stock Status
    def get_stock_status(row):
        stock = float(row.get("total_stock", 0))
        reorder_point = float(row.get("optimal stock", 0))
        alt = float(row.get("ADJUSTED_LEADTIME", 0))
        ads = float(row.get("Average Daily Sales", 0))
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
    for c in ["total_stock", "Average Daily Sales", "ADJUSTED_LEADTIME", "credit_days", "Predicted Order Quantity"]:
        if c not in final_stock_sales_df.columns:
            final_stock_sales_df[c] = 0.0
        final_stock_sales_df[c] = safe_numeric(final_stock_sales_df[c], 0)

    cds = final_stock_sales_df["credit_days"].clip(lower=0)
    ads = final_stock_sales_df["Average Daily Sales"].clip(lower=0)
    alt = final_stock_sales_df["ADJUSTED_LEADTIME"].clip(lower=0)
    s0 = final_stock_sales_df["total_stock"].clip(lower=0)
    q = final_stock_sales_df["Predicted Order Quantity"].clip(lower=0)

    target_low = ads * cds
    target_high = ads * (cds + alt)
    s_post = s0 + q
    gap_low = (target_low - s_post).clip(lower=0)
    excess = (s_post - target_high).clip(lower=0)
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
    final_stock_sales_df.loc[
        dmask, ["Predicted Order Quantity", "Predicted Stockout", "purchase_need", "QAC", "MOQ MAAD"]] = [0, False, 0,
                                                                                                          0, 0]
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
    # 🤖 ML AMÉLIORÉ: Stockout Probability
    # =========================
    try:
        # Features étendues pour meilleure prédiction
        feature_cols = [
            "total_stock", "Average Daily Sales",
            "Max Daily Sales (Pikine)", "optimal stock",
            "ADJUSTED_LEADTIME", "MAX_CREDIT_BUFFER", "AJUSTER_BUFFER"
        ]
        feature_cols = [c for c in feature_cols if c in final_stock_sales_df.columns]

        # Ajouter features calculées
        if "total_stock" in final_stock_sales_df.columns and "Average Daily Sales" in final_stock_sales_df.columns:
            stock = pd.to_numeric(final_stock_sales_df["total_stock"], errors="coerce").fillna(0)
            ads = pd.to_numeric(final_stock_sales_df["Average Daily Sales"], errors="coerce").fillna(0.01).replace(0, 0.01)

            # Coverage ratio (jours de stock)
            final_stock_sales_df["_coverage_ratio"] = stock / ads

            # Stock health score (0-1)
            optimal = pd.to_numeric(final_stock_sales_df.get("optimal stock", stock), errors="coerce").fillna(stock)
            final_stock_sales_df["_stock_health"] = (stock / optimal.replace(0, 1)).clip(0, 2)

            # Velocity score (ventes rapides = plus à risque)
            max_sales = pd.to_numeric(final_stock_sales_df.get("Max Daily Sales (Pikine)", ads), errors="coerce").fillna(ads)
            final_stock_sales_df["_velocity"] = max_sales / ads.replace(0, 0.01)

            feature_cols.extend(["_coverage_ratio", "_stock_health", "_velocity"])

        feature_cols = [c for c in feature_cols if c in final_stock_sales_df.columns]

        if feature_cols and "Predicted Stockout" in final_stock_sales_df.columns:
            df_train = final_stock_sales_df.dropna(subset=feature_cols + ["Predicted Stockout"]).copy()
            X = df_train[feature_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
            y = df_train["Predicted Stockout"].astype(int)

            if len(df_train) > 50 and y.nunique() > 1:
                X_train, X_test, y_train, y_test = train_test_split(
                    X, y, test_size=0.2, random_state=42, stratify=y
                )

                # Random Forest optimisé
                model = RandomForestClassifier(
                    n_estimators=150,  # Réduit pour rapidité
                    max_depth=10,
                    min_samples_split=5,
                    min_samples_leaf=2,
                    class_weight="balanced",
                    random_state=42,
                    n_jobs=-1
                )
                model.fit(X_train, y_train)

                y_pred = model.predict(X_test)
                prec = precision_score(y_test, y_pred)
                f1 = f1_score(y_test, y_pred)

                print(f"🤖 [ML Stockout] Modèle entraîné:")
                print(f"   - Précision: {prec:.2%}")
                print(f"   - F1-Score: {f1:.2%}")
                print(f"   - Features: {len(feature_cols)}")

                # Prédire probabilités
                X_all = final_stock_sales_df[feature_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
                probs = model.predict_proba(X_all)[:, 1]
                final_stock_sales_df["Stockout Probability"] = probs

                # Feature importance
                importances = dict(zip(feature_cols, model.feature_importances_))
                top_features = sorted(importances.items(), key=lambda x: x[1], reverse=True)[:3]
                print(f"   - Top features: {', '.join([f'{k}({v:.2f})' for k,v in top_features])}")

            else:
                # Fallback: règles métier simples
                print("🤖 [ML Stockout] Pas assez de données → utilisation règles métier")
                if "_coverage_ratio" in final_stock_sales_df.columns:
                    cov = final_stock_sales_df["_coverage_ratio"]
                    # Probabilité basée sur couverture
                    final_stock_sales_df["Stockout Probability"] = np.where(
                        cov <= 0, 1.0,
                        np.where(cov <= 3, 0.9,
                                 np.where(cov <= 7, 0.7,
                                          np.where(cov <= 14, 0.4,
                                                   np.where(cov <= 21, 0.2, 0.1))))
                    )
                else:
                    final_stock_sales_df["Stockout Probability"] = 0.0
        else:
            # Règles métier si pas de target
            print("🤖 [ML Stockout] Colonnes ML absentes → règles métier")
            if "total_stock" in final_stock_sales_df.columns and "Average Daily Sales" in final_stock_sales_df.columns:
                stock = pd.to_numeric(final_stock_sales_df["total_stock"], errors="coerce").fillna(0)
                ads = pd.to_numeric(final_stock_sales_df["Average Daily Sales"], errors="coerce").fillna(0.01).replace(0, 0.01)
                coverage = stock / ads

                final_stock_sales_df["Stockout Probability"] = np.where(
                    stock <= 0, 1.0,
                    np.where(coverage <= 3, 0.85,
                             np.where(coverage <= 7, 0.6,
                                      np.where(coverage <= 14, 0.3, 0.1)))
                )
            else:
                final_stock_sales_df["Stockout Probability"] = 0.0

    except Exception as e:
        print(f"🤖 [ML Stockout] Erreur: {e}")
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

    if "Average Daily Sales" not in final_stock_sales_df.columns:
        final_stock_sales_df["Average Daily Sales"] = 0.1

    final_stock_sales_df["Average Daily Sales"] = (
        pd.to_numeric(final_stock_sales_df["Average Daily Sales"], errors="coerce")
        .fillna(0.1)
        .clip(lower=0.1)
    )

    # ✅ DEBUG: Vérifier product_id avant réordonnancement
    print(f"🔍 DEBUG avant réord: product_id présent = {'product_id' in final_stock_sales_df.columns}")
    if 'product_id' in final_stock_sales_df.columns:
        print(f"   product_id valides: {(final_stock_sales_df['product_id'] > 0).sum()}")

    # Colonnes en tête pour visibilité
    _front = ["product_name", "Supplier", "Average Daily Sales"]
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
                                   "Average Daily Sales"]].head(10)
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

    # =========================
    # ✅ FILTRE 1 : Exclure les produits INACTIFS (is_active = False)
    # =========================
    if 'is_active' in final_stock_sales_df.columns:
        before_filter = len(final_stock_sales_df)

        # Convertir is_active en booléen
        def parse_active(x):
            if pd.isna(x):
                return True
            if isinstance(x, bool):
                return x
            return str(x).lower() in ('true', '1', 'yes', 'oui', 'actif')

        final_stock_sales_df['is_active'] = final_stock_sales_df['is_active'].apply(parse_active)

        # FILTRER : garder uniquement les actifs
        final_stock_sales_df = final_stock_sales_df[final_stock_sales_df['is_active'] == True].copy()

        after_filter = len(final_stock_sales_df)
        excluded_inactive = before_filter - after_filter

        print(f"\n{'=' * 60}")
        print(f"✅ FILTRE PRODUITS INACTIFS")
        print(f"{'=' * 60}")
        print(f"   Avant : {before_filter} produits")
        print(f"   Exclus (inactifs) : {excluded_inactive}")
        print(f"   Après : {after_filter} produits actifs")
        print(f"{'=' * 60}\n")
    else:
        print("\n⚠️ Colonne 'is_active' absente - TOUS les produits sont affichés")
        final_stock_sales_df['is_active'] = True

    # =========================
    # ✅ FILTRE 2 : Exclure les produits DELISTED
    # =========================
    if 'delisting_status' in final_stock_sales_df.columns:
        before_delist = len(final_stock_sales_df)

        # FILTRER : exclure les delisted
        delist_mask = final_stock_sales_df['delisting_status'].astype(str).str.lower() == 'delisted'
        final_stock_sales_df = final_stock_sales_df[~delist_mask].copy()

        after_delist = len(final_stock_sales_df)
        excluded_delisted = before_delist - after_delist

        if excluded_delisted > 0:
            print(f"✅ FILTRE DELISTED : {excluded_delisted} produits exclus")
        print(f"   → Reste : {after_delist} produits")

    # =========================
    # ✅ FILTRE 3 : Exclure les produits "cadeau"
    # =========================
    if 'product_name' in final_stock_sales_df.columns:
        before_cadeau = len(final_stock_sales_df)

        # Exclure tout produit contenant "cadeau" (insensible à la casse)
        cadeau_mask = final_stock_sales_df['product_name'].astype(str).str.lower().str.contains(
            'cadeau',
            na=False,
            regex=False
        )

        final_stock_sales_df = final_stock_sales_df[~cadeau_mask].copy()

        excluded_count = before_cadeau - len(final_stock_sales_df)

        if excluded_count > 0:
            print(f"✅ FILTRE CADEAU : {excluded_count} produits exclus")
        else:
            print("ℹ️ Aucun produit 'cadeau' détecté")

        # =========================
        # ✅ FILTRE : Exclure les frais (livraison, transport, etc.)
        # =========================
        if 'product_name' in final_stock_sales_df.columns:
            before_frais = len(final_stock_sales_df)

            # Liste des termes à exclure
            frais_keywords = [
                'frais de livraison',
                'frais de transport',
                'frais de majoration',
                'remboursement prêt',
                'rubyx',
                'cfa - cash',
                'affiches',
                'remise',
                'livraison'
            ]

            # Créer un pattern regex pour matcher n'importe lequel de ces termes
            pattern = '|'.join(frais_keywords)

            frais_mask = final_stock_sales_df['product_name'].astype(str).str.lower().str.contains(
                pattern,
                na=False,
                regex=True
            )

            final_stock_sales_df = final_stock_sales_df[~frais_mask].copy()

            excluded_count = before_frais - len(final_stock_sales_df)

            if excluded_count > 0:
                print(f"✅ FILTRE FRAIS : {excluded_count} produits exclus")
                print(f"   (livraison, transport, majoration, remboursement)")

            if not promo_roi.empty:
                # Debug : vérifier colonnes créées
                print(f"\n✅ Colonnes promo ajoutées :")
                for col in ['uplift_pct', 'roi_pct', 'promo_recommendation', 'promo_priority']:
                    if col in final_stock_sales_df.columns:
                        print(f"   ✓ {col}")
                    else:
                        print(f"   ✗ {col} MANQUANTE")

        # ✅ GARANTIR product_id avant le return
        if 'product_id' not in final_stock_sales_df.columns:
            final_stock_sales_df['product_id'] = 0
            print("⚠️ product_id créé avec 0 (absent)")
        else:
            final_stock_sales_df['product_id'] = pd.to_numeric(
                final_stock_sales_df['product_id'], errors='coerce'
            ).fillna(0).astype(int)
            valid_count = (final_stock_sales_df['product_id'] > 0).sum()
            print(f"✅ product_id final: {valid_count}/{len(final_stock_sales_df)} valides")

        # ✅ AJOUTER JUSTE AVANT LE RETURN FINAL
        print(f"\n{'=' * 70}")
        print(f"🎉 LOAD_SUPPLY_DATA TERMINÉ")
        print(f"{'=' * 70}")
        print(f"   Période : {period_days}")
        print(f"   Produits : {len(final_stock_sales_df)}")
        print(f"   Colonnes : {len(final_stock_sales_df.columns)}")
        print(f"   product_id valides : {(final_stock_sales_df['product_id'] > 0).sum()}")

        if 'Average Daily Sales' in final_stock_sales_df.columns:
            print(f"   ADS min : {final_stock_sales_df['Average Daily Sales'].min():.2f}")
            print(f"   ADS max : {final_stock_sales_df['Average Daily Sales'].max():.2f}")
            print(f"   ADS moy : {final_stock_sales_df['Average Daily Sales'].mean():.2f}")

        print(f"{'=' * 70}\n")

        return final_stock_sales_df


# ==========================================
# ⚡ PRÉ-CHARGEMENT DES 3 PÉRIODES
# ==========================================
print("\n" + "=" * 80)
print("🚀 DÉMARRAGE - PRÉ-CHARGEMENT DES DONNÉES (peut prendre 2-3 minutes)")
print("=" * 80)
print("⏳ Patience, ce chargement n'aura lieu qu'UNE SEULE FOIS...")
print()

import time

start_total = time.time()

# ✅ Chargement 7 jours (défaut)
print("📥 [1/3] Chargement période 7 jours...")
start = time.time()
try:
    DATA_7D = load_supply_data(period_days="7d")
    elapsed = time.time() - start
    print(f"    ✅ 7j : {len(DATA_7D)} produits en {elapsed:.1f}s")
    print(
        f"       ADS min/max/moy : {DATA_7D['Average Daily Sales'].min():.2f} / {DATA_7D['Average Daily Sales'].max():.2f} / {DATA_7D['Average Daily Sales'].mean():.2f}")
except Exception as e:
    print(f"    ❌ Erreur 7j : {e}")
    DATA_7D = pd.DataFrame()

print()

# ✅ Chargement 3 jours
print("📥 [2/3] Chargement période 3 jours...")
start = time.time()
try:
    DATA_3D = load_supply_data(period_days="3d")
    elapsed = time.time() - start
    print(f"    ✅ 3j : {len(DATA_3D)} produits en {elapsed:.1f}s")
    print(
        f"       ADS min/max/moy : {DATA_3D['Average Daily Sales'].min():.2f} / {DATA_3D['Average Daily Sales'].max():.2f} / {DATA_3D['Average Daily Sales'].mean():.2f}")
except Exception as e:
    print(f"    ❌ Erreur 3j : {e}")
    DATA_3D = pd.DataFrame()

print()

# ✅ Chargement 30 jours
print("📥 [3/3] Chargement période 30 jours...")
start = time.time()
try:
    DATA_30D = load_supply_data(period_days="30d")
    elapsed = time.time() - start
    print(f"    ✅ 30j : {len(DATA_30D)} produits en {elapsed:.1f}s")
    print(
        f"       ADS min/max/moy : {DATA_30D['Average Daily Sales'].min():.2f} / {DATA_30D['Average Daily Sales'].max():.2f} / {DATA_30D['Average Daily Sales'].mean():.2f}")
except Exception as e:
    print(f"    ❌ Erreur 30j : {e}")
    DATA_30D = pd.DataFrame()

elapsed_total = time.time() - start_total
print()
print("=" * 80)
print(f"🎉 PRÉ-CHARGEMENT TERMINÉ en {elapsed_total:.1f}s")
print("=" * 80)
print("✨ Le changement de période sera maintenant INSTANTANÉ (< 1 seconde) !")
print("=" * 80)
print()


# ==========================================
# ⚡ FONCTION D'ACCÈS ULTRA-RAPIDE
# ==========================================
def get_preloaded_data(period_days="7d"):
    """
    Retourne données pré-chargées (INSTANTANÉ - < 0.1s)

    Args:
        period_days: "3d", "7d" ou "30d"

    Returns:
        DataFrame avec toutes les colonnes déjà calculées
    """
    if period_days == "3d":
        return DATA_3D.copy()
    elif period_days == "30d":
        return DATA_30D.copy()
    else:
        return DATA_7D.copy()
# ==================== CALLBACK ROTATION ADS ====================
'''
@app.callback(
    [Output("master-data", "data", allow_duplicate=True),
     Output("filtered-data", "data", allow_duplicate=True),
     Output("main-table", "data", allow_duplicate=True),
     Output("rotation-indicator", "children")],
    Input("rotation-period", "value"),
    prevent_initial_call=True
)
def update_rotation_period(period_value):
    """
    Recalcule toutes les données avec nouvelle période de rotation
    ✅ GARANTI de toujours retourner 4 valeurs
    """
    print(f"\n{'=' * 70}")
    print(f"🔥 CALLBACK ROTATION DÉCLENCHÉ : {period_value}")
    print(f"{'=' * 70}")

    # ✅ PROTECTION #1 : Valeur vide
    if not period_value:
        print("❌ Valeur vide reçue")
        error_msg = html.Div("⚠️ Erreur : valeur vide", style={"color": "#ef4444"})
        return no_update, no_update, no_update, error_msg

    try:
        # ✅ Invalider cache
        print("🗑️ Invalidation cache...")
        get_df_cached.cache_clear()

        # ✅ Recharger données
        print(f"📥 Chargement données pour période {period_value}...")
        df = load_supply_data(period_days=period_value)
        print(f"✅ {len(df)} produits chargés")

        # ✅ PROTECTION #2 : DataFrame vide
        if df.empty:
            print("⚠️ DataFrame vide reçu")
            error_msg = html.Div("⚠️ Aucune donnée chargée", style={"color": "#ef4444"})
            return no_update, no_update, no_update, error_msg

        # ✅ Validation colonnes critiques
        print("🔍 Validation colonnes...")

        # Supplier
        if "Supplier" not in df.columns:
            df["Supplier"] = "unknown"
        else:
            df["Supplier"] = (
                df["Supplier"]
                .fillna("unknown")
                .astype(str)
                .str.strip()
                .replace("", "unknown")
            )

        # Average Daily Sales
        if "Average Daily Sales" not in df.columns:
            print("⚠️ Colonne ADS manquante, création par défaut")
            df["Average Daily Sales"] = 0.1
        else:
            df["Average Daily Sales"] = (
                pd.to_numeric(df["Average Daily Sales"], errors="coerce")
                .fillna(0.1)
                .clip(lower=0.1)
            )

        # ✅ NETTOYAGE COLONNES DUPLIQUÉES product_id
        # Chercher toutes les colonnes qui contiennent "product_id"
        product_id_cols = [c for c in df.columns if 'product_id' in c.lower()]
        print(f"📋 Colonnes product_id trouvées: {product_id_cols}")

        if len(product_id_cols) > 1:
            # Trouver la colonne avec les vraies valeurs (pas que des 0)
            best_col = None
            best_count = 0
            for col in product_id_cols:
                try:
                    col_numeric = pd.to_numeric(df[col], errors='coerce').fillna(0)
                    valid_count = (col_numeric > 0).sum()
                    print(f"   - {col}: {valid_count} valeurs > 0")
                    if valid_count > best_count:
                        best_count = valid_count
                        best_col = col
                except:
                    pass

            if best_col:
                print(f"   ✅ Colonne retenue: {best_col} ({best_count} valeurs)")
                # Garder la meilleure colonne et supprimer les autres
                df['product_id'] = pd.to_numeric(df[best_col], errors='coerce').fillna(0).astype(int)
                # Supprimer les doublons
                cols_to_drop = [c for c in product_id_cols if c != 'product_id']
                df.drop(columns=cols_to_drop, errors='ignore', inplace=True)
                print(f"   ✅ Colonnes supprimées: {cols_to_drop}")

        # S'assurer que product_id existe et est propre
        if "product_id" not in df.columns:
            df["product_id"] = 0
        df["product_id"] = pd.to_numeric(df["product_id"], errors="coerce").fillna(0).astype(int)

        print(f"✅ product_id final: {(df['product_id'] > 0).sum()} valeurs valides sur {len(df)}")

        # ✅ Préparer colonnes table
        print("📋 Préparation colonnes...")

        cols_priority = [
            "product_id", "product_name", "Supplier", "total_stock",
            "Average Daily Sales", "Max Daily Sales (Pikine)", "optimal stock",
            "Ajusted_total_need", "QAC", "target_quantity",
            "Max Coverage Day", "Product Category", "credit_days"
        ]

        cols_banned = [
            "Average Daily Sales (7d)", "Average Daily Sales (30d)",
            "Average Daily Sales (3d)", "Daily OOS Rate (7d)",
            "Daily OOS Rate (30d)", "Stockout Probability",
            "is_active", "promo_status", "uplift_pct"
        ]

        # Filtrer colonnes disponibles
        available_priority = [c for c in cols_priority if c in df.columns]
        extra_cols = [
            c for c in df.columns
            if c not in cols_priority and c not in cols_banned
        ]

        # Supprimer colonnes bannies
        df_overview = df.drop(
            columns=[c for c in cols_banned if c in df.columns],
            errors='ignore'
        )

        available_cols = available_priority + extra_cols

        # ✅ PROTECTION #3 : Garantir product_name
        if "product_name" not in available_cols and "product_name" in df_overview.columns:
            available_cols.insert(0, "product_name")

        print(f"✅ Colonnes disponibles : {len(available_cols)}")

        # ✅ Créer indicateur visuel
        period_labels = {
            "3d": "⚡ 3 jours - Ultra court terme",
            "7d": "📅 7 jours - Court terme (défaut)",
            "30d": "📊 30 jours - Moyen terme"
        }

        stats_text = (
            f"Min: {df['Average Daily Sales'].min():.2f} • "
            f"Med: {df['Average Daily Sales'].median():.2f} • "
            f"Max: {df['Average Daily Sales'].max():.2f} • "
            f"Moy: {df['Average Daily Sales'].mean():.2f}"
        )

        indicator = html.Div([
            html.Div([
                html.Span("✓ ", style={
                    "color": "#10b981",
                    "fontSize": "18px",
                    "marginRight": "6px"
                }),
                html.Strong(
                    period_labels.get(period_value, f"Période {period_value}"),
                    style={
                        "color": "#3b82f6",
                        "fontSize": "14px",
                        "fontWeight": "700"
                    }
                )
            ], style={"marginBottom": "6px"}),
            html.Div([
                html.Span(
                    f"📦 {len(df)} produits • ",
                    style={
                        "color": "#94a3b8",
                        "marginRight": "10px",
                        "fontWeight": "600"
                    }
                ),
                html.Span(
                    f"ADS : {stats_text}",
                    style={
                        "color": "#64748b",
                        "fontSize": "11px"
                    }
                )
            ]),
            html.Small(
                f"Mis à jour : {datetime.now().strftime('%H:%M:%S')}",
                style={
                    "color": "#6b7280",
                    "fontSize": "10px",
                    "display": "block",
                    "marginTop": "4px"
                }
            )
        ])

        # ✅ PROTECTION #4 : Vérifier colonnes avant to_dict
        missing_cols = [c for c in available_cols if c not in df_overview.columns]
        if missing_cols:
            print(f"⚠️ Colonnes manquantes : {missing_cols}")
            # Retirer colonnes manquantes
            available_cols = [c for c in available_cols if c in df_overview.columns]

        # ✅ NETTOYAGE FINAL: Garantir UNE SEULE colonne product_id
        product_id_variants = [c for c in df.columns if 'product_id' in c.lower()]
        if len(product_id_variants) > 1:
            print(f"⚠️ NETTOYAGE: {len(product_id_variants)} colonnes product_id détectées: {product_id_variants}")
            # Garder celle avec le plus de valeurs valides
            best_col = max(product_id_variants, key=lambda c: (pd.to_numeric(df[c], errors='coerce').fillna(0) > 0).sum())
            df['product_id'] = pd.to_numeric(df[best_col], errors='coerce').fillna(0).astype(int)
            cols_to_drop = [c for c in product_id_variants if c != 'product_id']
            df.drop(columns=cols_to_drop, errors='ignore', inplace=True)
            df_overview.drop(columns=[c for c in cols_to_drop if c in df_overview.columns], errors='ignore', inplace=True)
            print(f"   ✅ Colonnes supprimées: {cols_to_drop}")

        # S'assurer que product_id est en première position dans available_cols
        if 'product_id' in df_overview.columns and 'product_id' not in available_cols:
            available_cols.insert(0, 'product_id')
        elif 'product_id' in available_cols and available_cols[0] != 'product_id':
            available_cols.remove('product_id')
            available_cols.insert(0, 'product_id')

        print(f"✅ product_id final: {(df['product_id'] > 0).sum()} valeurs valides")

        # ✅ Préparer données retour
        print("📤 Préparation données retour...")

        master_json = df.to_json(orient="records")
        filtered_json = df_overview.to_json(orient="records")
        table_data = df_overview[available_cols].to_dict("records")

        print(f"✅ CALLBACK TERMINÉ AVEC SUCCÈS")
        print(f"   - Master data : {len(master_json)} chars")
        print(f"   - Filtered data : {len(filtered_json)} chars")
        print(f"   - Table rows : {len(table_data)}")
        print(f"{'=' * 70}\n")

        # ✅ RETOUR GARANTI (4 valeurs)
        return master_json, filtered_json, table_data, indicator

    except Exception as e:
        print(f"❌ ERREUR DANS CALLBACK : {e}")
        import traceback
        traceback.print_exc()

        # ✅ En cas d'erreur, retourner 4 valeurs quand même
        error_indicator = html.Div([
            html.Span("⚠️ ", style={
                "fontSize": "18px",
                "marginRight": "6px",
                "color": "#ef4444"
            }),
            html.Span(
                f"Erreur : {str(e)[:100]}",
                style={
                    "color": "#ef4444",
                    "fontWeight": "600",
                    "fontSize": "12px"
                }
            )
        ])

        # ✅ RETOUR EN CAS D'ERREUR (4 valeurs)
        return no_update, no_update, no_update, error_indicator
'''
# Utility: add Actions columns
def add_action_cols(df: pd.DataFrame) -> pd.DataFrame:
    df2 = df.copy()
    df2["edit Edit"] = "edit"
    df2["delete Delete"] = "delete"
    return df2


# ✅ ✅ ✅ PLACER ICI LA FONCTION recalculate_product_metrics ✅ ✅ ✅

def recalculate_product_metrics(row: pd.Series) -> pd.Series:
    """
    Recalcule TOUS les paramètres dépendants d'un produit après modification.
    Garantit la cohérence métier Supply Chain.

    Args:
        row: Ligne produit (pd.Series) avec valeurs potentiellement modifiées

    Returns:
        pd.Series avec toutes les métriques recalculées
    """


    row = row.copy()

    # === VALIDATION & NETTOYAGE ===
    numeric_cols = {
        #'total_stock': 0,
        'Average Daily Sales': 0.1,
        'Max Daily Sales (Pikine)': 0,
        'QAC': 0,
        'MOQ MAAD': 0,
        'purchase_need': 0,
        'credit_days': 0,
        'ADJUSTED_LEADTIME': 7,
        'AJUSTER_BUFFER': 0,
        #'target_quantity': 0,
        #'optimal stock': 0,
        'Max Coverage Day': 0
    }

    for col, default in numeric_cols.items():
        if col in row.index:
            val = pd.to_numeric(row[col], errors='coerce')
            row[col] = max(val if pd.notna(val) else default, default)

    # === PARAMÈTRES DE BASE ===
    ads = max(row.get('Average Daily Sales', 0.1), 0.1)
    leadtime = row.get('ADJUSTED_LEADTIME', 7)
    credit = row.get('credit_days', 0)
    buffer = row.get('AJUSTER_BUFFER', 0)
    stock = row.get('total_stock', 0)
    max_daily = row.get('Max Daily Sales (Pikine)', 0)

    # Si Max Daily Sales non défini, estimer à 1.5x ADS
    if max_daily == 0:
        max_daily = ads * 1.5
        row['Max Daily Sales (Pikine)'] = max_daily

    # === RECALCULS MÉTIER (ordre critique) ===

    # 1. Couverture en jours
    row['Max Coverage Day'] = min(stock / ads, 365) if ads > 0 else 0

    # 2. Stock optimal
    row['optimal stock'] = max(
        max_daily,
        (max_daily / 2.0) + (buffer * ads)
    )

    # 3. MOQ MAAD
    row['MOQ MAAD'] = (leadtime + 3) * ads

    # 4. Purchase Need
    optimal = row['optimal stock']
    param_supplier_factor = row.get('Param_Supplier_Factor', 0)
    param_buffer_value = row.get('Param_Buffer_Value_Lookup', 0)

    if stock <= 0:
        row['purchase_need'] = max_daily + (param_buffer_value * ads)
    elif stock < optimal + (param_supplier_factor * ads):
        row['purchase_need'] = max(0, (optimal - stock) + (param_supplier_factor * ads))
    else:
        row['purchase_need'] = 0

    # 5. QAC (ne recalculer que si vide)
    current_qac = row.get('QAC', 0)
    if current_qac == 0 or pd.isna(current_qac):
        row['QAC'] = max(row['MOQ MAAD'], row['purchase_need'])

    # 6. Target Quantity
    row['target_quantity'] = max(
        0,
        ads * (leadtime + credit) + buffer * ads - stock
    )

    # 7. Adjusted Total Need
    delisting = str(row.get('delisting_status', '')).lower()
    if delisting == 'delisted':
        row['Ajusted_total_need'] = 'NO NEED'
    else:
        coverage = row['Max Coverage Day']
        replenishment_days = leadtime + credit
        optimal_coverage = optimal / ads if ads > 0 else 0

        if coverage <= leadtime + 3:
            row['Ajusted_total_need'] = 'ORDER NOW'
        elif coverage < replenishment_days + optimal_coverage:
            row['Ajusted_total_need'] = 'ORDER NOT URGENT'
        else:
            row['Ajusted_total_need'] = 'NO NEED'

    # 8. Stock Status
    if stock <= 0:
        row['Stock Status'] = 'Out of Stock'
    elif stock <= optimal:
        row['Stock Status'] = 'Predicted Stockout Soon'
    elif stock <= optimal + (leadtime * ads):
        row['Stock Status'] = 'Order Soon'
    else:
        row['Stock Status'] = 'Stock OK'

    # 9. Predicted Stockout
    row['Predicted Stockout'] = (stock <= optimal)

    # 10. Credit Adequacy
    if credit > 0:
        target_low = ads * credit
        target_high = ads * (credit + leadtime)
        stock_post = stock + row.get('Predicted Order Quantity', row['QAC'])

        gap_low = max(0, target_low - stock_post)
        excess = max(0, stock_post - target_high)

        denom = target_low if target_low > 0 else 1.0
        err = (gap_low + excess) / denom

        row['Credit Adequacy Score'] = np.exp(-1.5 * err)
        row['Credit Adequacy Risk'] = (err > 0)
    else:
        row['Credit Adequacy Score'] = 1.0 if ads > 0 else 0.0
        row['Credit Adequacy Risk'] = False

    # 11. Timestamp de modification
    row['last_modified'] = datetime.now().isoformat()

    return row


def train_optimal_order_quantity_model(df: pd.DataFrame) -> tuple:
    """
    Entraîne un modèle ML pour prédire la quantité optimale de commande.

    Objectif : Commander juste assez pour couvrir la demande pendant credit_days + lead_time
    sans rupture ni surstock excessif.

    Returns:
        (df_with_predictions, model) ou (df_with_predictions, None) si échec
    """
    from sklearn.ensemble import GradientBoostingRegressor
    from sklearn.model_selection import train_test_split
    from sklearn.metrics import mean_absolute_error, r2_score

    df_work = df.copy()

    # =========================
    # Feature Engineering
    # =========================
    required_cols = [
        'credit_days', 'total_stock', 'Average Daily Sales',
        'Max Daily Sales (Pikine)', 'ADJUSTED_LEADTIME'
    ]

    # Vérifier colonnes requises
    missing = [c for c in required_cols if c not in df_work.columns]
    if missing:
        print(f"❌ Colonnes manquantes pour ML : {missing}")
        df_work['Predicted Order Quantity'] = 0
        return df_work, None

    # Features
    df_work['demand_during_credit'] = (
            df_work['Average Daily Sales'] * df_work['credit_days']
    ).clip(lower=0)

    df_work['demand_during_leadtime'] = (
            df_work['Average Daily Sales'] * df_work['ADJUSTED_LEADTIME']
    ).clip(lower=0)

    df_work['total_demand_period'] = (
            df_work['demand_during_credit'] + df_work['demand_during_leadtime']
    )

    df_work['stock_coverage_ratio'] = pd.Series(
        np.where(
            df_work['Average Daily Sales'] > 0,
            df_work['total_stock'] / df_work['Average Daily Sales'],
            0
        ),
        index=df_work.index
    ).clip(upper=365)

    df_work['demand_volatility'] = pd.Series(
        df_work['Max Daily Sales (Pikine)'] /
        np.maximum(df_work['Average Daily Sales'], 0.1),
        index=df_work.index
    ).clip(upper=10)

    # =========================
    # Target : Quantité optimale théorique
    # =========================
    def calculate_optimal_quantity(row):
        """
        Quantité optimale = demande totale période - stock actuel + buffer sécurité
        """
        total_demand = row['total_demand_period']
        current_stock = row['total_stock']
        safety_buffer = row['Max Daily Sales (Pikine)'] * 0.5  # 50% du pic comme sécurité

        optimal = total_demand - current_stock + safety_buffer
        return max(0, optimal)  # Jamais négatif

    df_work['target_quantity'] = df_work.apply(calculate_optimal_quantity, axis=1)

    # =========================
    # Préparer données d'entraînement
    # =========================
    feature_cols = [
        'credit_days', 'total_stock', 'Average Daily Sales',
        'Max Daily Sales (Pikine)', 'ADJUSTED_LEADTIME',
        'demand_during_credit', 'demand_during_leadtime',
        'stock_coverage_ratio', 'demand_volatility'
    ]

    # Filtrer données valides
    valid_mask = (
            (df_work['target_quantity'] > 0) &
            (df_work['Average Daily Sales'] > 0) &
            (df_work[feature_cols].notna().all(axis=1))
    )

    df_train = df_work[valid_mask].copy()

    if len(df_train) < 50:
        print(f"⚠️ Données insuffisantes pour ML ({len(df_train)} échantillons)")
        df_work['Predicted Order Quantity'] = df_work['target_quantity']
        return df_work, None

    X = df_train[feature_cols].fillna(0)
    y = df_train['target_quantity']

    # =========================
    # Entraînement
    # =========================
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42
    )

    try:
        model = GradientBoostingRegressor(
            n_estimators=300,
            max_depth=6,
            learning_rate=0.05,
            subsample=0.8,
            min_samples_leaf=5,
            random_state=42
        )

        model.fit(X_train, y_train)

        # Évaluation
        y_pred_test = model.predict(X_test)
        mae = mean_absolute_error(y_test, y_pred_test)
        r2 = r2_score(y_test, y_pred_test)

        print(f"\n{'=' * 60}")
        print(f"✅ MODÈLE ML - PREDICTED ORDER QUANTITY")
        print(f"{'=' * 60}")
        print(f"   Échantillons entraînement : {len(X_train)}")
        print(f"   MAE (test) : {mae:.2f} unités")
        print(f"   R² score : {r2:.3f}")
        print(f"   Features importance :")

        importance = sorted(
            zip(feature_cols, model.feature_importances_),
            key=lambda x: x[1],
            reverse=True
        )[:5]
        for feat, imp in importance:
            print(f"      - {feat}: {imp:.3f}")
        print(f"{'=' * 60}\n")

        # Prédire sur tout le dataset
        X_all = df_work[feature_cols].fillna(0)
        df_work['Predicted Order Quantity'] = model.predict(X_all).clip(lower=0)

        return df_work, model

    except Exception as e:
        print(f" Erreur entraînement ML : {e}")
        df_work['Predicted Order Quantity'] = df_work['target_quantity']
        return df_work, None


# ----------------------------- App & Cache ---------------------------------------
#app = Dash(__name__, title=APP_TITLE, external_stylesheets=[THEME], suppress_callback_exceptions=True, prevent_initial_callbacks='initial_duplicate')
#server = app.server
#cache = Cache(app.server, config={"CACHE_TYPE": "SimpleCache", "CACHE_DEFAULT_TIMEOUT": 3600})

# ------------------------------ Custom CSS & JS ----------------------------------
app.index_string = """
<!DOCTYPE html>
<html>
    <script>
    // ========================================
    // 🎨 COLORIER LES LIGNES SELON BESOIN D'ACHAT
    // ========================================
    function colorizeTableRows() {
        const table = document.querySelector('.dash-table-container table');
        if (!table) return;

        const rows = table.querySelectorAll('tbody tr');

        rows.forEach(row => {
            const cells = row.querySelectorAll('td');

            // Trouver la cellule Ajusted_total_need
            cells.forEach((cell, index) => {
                const text = cell.textContent.trim();

                if (text === 'ORDER NOW') {
                    // 🔴 Ligne rouge
                    row.style.backgroundColor = 'rgba(239, 68, 68, 0.15)';
                    cells.forEach(c => {
                        c.style.color = '#fecaca';
                        c.style.fontWeight = '600';
                    });

                    // Badge rouge pour la cellule
                    cell.style.backgroundColor = 'rgba(239, 68, 68, 0.4)';
                    cell.style.color = '#ffffff';
                    cell.style.fontWeight = '800';
                    cell.style.border = '2px solid #ef4444';
                    cell.style.borderRadius = '6px';
                    cell.style.textTransform = 'uppercase';

                    // Bordure gauche sur product_name
                    if (cells[1]) {
                        cells[1].style.borderLeft = '4px solid #ef4444';
                        cells[1].style.backgroundColor = 'rgba(239, 68, 68, 0.25)';
                        cells[1].style.color = '#fee2e2';
                        cells[1].style.fontWeight = '700';
                    }

                    // Highlight QAC
                    cells.forEach((c, i) => {
                        const header = table.querySelectorAll('thead th')[i];
                        if (header && header.textContent.includes('QAC')) {
                            c.style.backgroundColor = 'rgba(239, 68, 68, 0.3)';
                            c.style.color = '#ffffff';
                            c.style.fontWeight = '800';
                            c.style.fontSize = '15px';
                        }
                    });
                }

                else if (text === 'ORDER NOT URGENT') {
                    // 🟠 Ligne orange
                    row.style.backgroundColor = 'rgba(245, 158, 11, 0.12)';
                    cells.forEach(c => {
                        c.style.color = '#fde68a';
                        c.style.fontWeight = '500';
                    });

                    // Badge orange
                    cell.style.backgroundColor = 'rgba(245, 158, 11, 0.35)';
                    cell.style.color = '#ffffff';
                    cell.style.fontWeight = '700';
                    cell.style.border = '2px solid #f59e0b';
                    cell.style.borderRadius = '6px';
                    cell.style.textTransform = 'uppercase';

                    // Bordure gauche
                    if (cells[1]) {
                        cells[1].style.borderLeft = '4px solid #f59e0b';
                        cells[1].style.backgroundColor = 'rgba(245, 158, 11, 0.2)';
                        cells[1].style.color = '#fef3c7';
                        cells[1].style.fontWeight = '600';
                    }
                }

                else if (text === 'NO NEED') {
                    // 🟢 Ligne verte
                    row.style.backgroundColor = 'rgba(16, 185, 129, 0.08)';
                    cells.forEach(c => {
                        c.style.color = '#d1fae5';
                    });

                    // Badge vert
                    cell.style.backgroundColor = 'rgba(16, 185, 129, 0.3)';
                    cell.style.color = '#ffffff';
                    cell.style.fontWeight = '700';
                    cell.style.border = '2px solid #10b981';
                    cell.style.borderRadius = '6px';
                    cell.style.textTransform = 'uppercase';

                    // Bordure gauche
                    if (cells[1]) {
                        cells[1].style.borderLeft = '4px solid #10b981';
                    }
                }
            });
        });
    }

    // Exécuter au chargement et après chaque mise à jour
    window.addEventListener('load', () => {
        colorizeTableRows();

        // Observer les changements du DOM
        const observer = new MutationObserver(() => {
            setTimeout(colorizeTableRows, 100);
        });

        const tableContainer = document.querySelector('.dash-table-container');
        if (tableContainer) {
            observer.observe(tableContainer, {
                childList: true,
                subtree: true
            });
        }
    });
    </script>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <style>
            /* ========================================
               VARIABLES - THÈME CLAIR MODERNE
               ======================================== */
            :root {
                --brand-accent: #0ea5e9;
                --brand-hover: #0284c7;
                --badge-danger: #dc2626;
                --badge-warning: #d97706;
                --badge-ok: #059669;
                --bg-primary: #f8fafc;
                --bg-secondary: #ffffff;
                --bg-tertiary: #f1f5f9;
                --bg-input: #ffffff;
                --text-primary: #1e293b;
                --text-secondary: #475569;
                --text-muted: #64748b;
                --border-color: #e2e8f0;
                --border-hover: #cbd5e1;
            }

            /* ========================================
               LAYOUT
               ======================================== */
            body {
                background: var(--bg-primary) !important;
                color: var(--text-primary) !important;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
            }

            .sidebar { 
                position: fixed; 
                top: 0; 
                bottom: 0; 
                left: 0; 
                width: 270px;
                padding: 14px 12px; 
                background: linear-gradient(180deg, #1e3a5f 0%, #0f2744 100%); 
                border-right: none; 
                overflow-y: auto;
                box-shadow: 4px 0 20px rgba(0, 0, 0, 0.08);
            }

            .content { 
                margin-left: 150px; 
                padding: 16px 20px 100px 20px; 
                background: var(--bg-primary) !important; 
                color: var(--text-primary) !important; 
                min-height: 100vh;
            }

            .brand { 
                font-weight: 800; 
                font-size: 20px; 
                letter-spacing: 0.6px; 
                color: var(--text-primary); 
            }

            .muted { 
                color: var(--text-muted); 
                font-size: 13px; 
            }

            /* ========================================
               CARTES & CONTAINERS
               ======================================== */
            .pill { 
                border: 1px solid var(--border-color); 
                padding: 10px 12px; 
                border-radius: 12px; 
                background: var(--bg-secondary); 
                color: var(--text-primary) !important;
                box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
            }

            .kpi { 
                border-radius: 16px; 
                padding: 20px; 
                border: 1px solid var(--border-color); 
                background: var(--bg-secondary);
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
                transition: all 0.3s ease;
            }

            .kpi:hover {
                transform: translateY(-2px);
                box-shadow: 0 8px 24px rgba(14, 165, 233, 0.15);
                border-color: rgba(14, 165, 233, 0.3);
            }

            /* ========================================
               CARTES PRODUITS À RISQUE (CLIQUABLES)
               ======================================== */
            .risk-product-card {
                transition: all 0.2s ease !important;
            }

            .risk-product-card:hover {
                background: var(--bg-tertiary) !important;
                transform: translateX(4px);
                box-shadow: 0 4px 12px rgba(34, 211, 238, 0.2);
                border-color: #0ea5e9 !important;
            }

            .risk-product-card:active {
                transform: translateX(2px);
                background: #e2e8f0 !important;
            }

            .kpi h3 {
                color: var(--brand-accent) !important;
                font-weight: 700 !important;
                font-size: 32px !important;
                margin-top: 8px !important;
            }

            .kpi small {
                color: var(--text-secondary) !important;
                font-weight: 600 !important;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                font-size: 11px !important;
            }

            .soft-card { 
                border: 1px solid var(--border-color); 
                border-radius: 16px; 
                padding: 20px; 
                background: var(--bg-secondary);
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
            }

            /* ========================================
               BADGES
               ======================================== */
            .badge { 
                padding: 4px 10px; 
                border-radius: 12px; 
                font-size: 11px; 
                font-weight: 700;
                letter-spacing: 0.3px;
            }

            .badge-danger { 
                background: linear-gradient(135deg, var(--badge-danger) 0%, #dc2626 100%); 
                color: #fff; 
                border: none;
            }

            .badge-warn { 
                background: linear-gradient(135deg, var(--badge-warning) 0%, #d97706 100%); 
                color: #0a0e1a; 
                border: none;
            }

            .badge-ok { 
                background: linear-gradient(135deg, var(--badge-ok) 0%, #059669 100%); 
                color: #0a0e1a; 
                border: none;
            }

            /* ========================================
               BOUTONS
               ======================================== */
            .btn-primary { 
                background: linear-gradient(135deg, var(--brand-accent) 0%, var(--brand-hover) 100%) !important;
                color: #0a0e1a !important;
                font-weight: 700 !important;
                border: none !important;
                border-radius: 10px !important;
                padding: 10px 20px !important;
                transition: all 0.3s ease !important;
            }

            .btn-primary:hover {
                background: linear-gradient(135deg, var(--brand-hover) 0%, #0891b2 100%) !important;
                transform: translateY(-2px);
                box-shadow: 0 6px 20px rgba(34, 211, 238, 0.4) !important;
            }

            .btn-outline-secondary {
                border: 2px solid var(--border-color) !important;
                color: var(--text-primary) !important;
                background: transparent !important;
                border-radius: 10px !important;
                font-weight: 600 !important;
                transition: all 0.3s ease !important;
            }

            .btn-outline-secondary:hover {
                background: rgba(34, 211, 238, 0.1) !important;
                border-color: var(--brand-accent) !important;
                color: var(--brand-accent) !important;
            }

            /* ========================================
               🔍 BARRE DE RECHERCHE (CRITIQUE)
               ======================================== */
            #search-input {
                background: var(--bg-input) !important;
                color: var(--text-primary) !important;
                border: 2px solid var(--border-color) !important;
                border-radius: 12px !important;
                padding: 12px 16px !important;
                font-size: 14px !important;
                font-weight: 500 !important;
                transition: all 0.3s ease !important;
            }

            #search-input::placeholder {
                color: var(--text-muted) !important;
                opacity: 0.8 !important;
            }

            #search-input:focus {
                background: var(--bg-tertiary) !important;
                border-color: var(--brand-accent) !important;
                color: var(--text-primary) !important;
                box-shadow: 0 0 0 3px rgba(34, 211, 238, 0.15) !important;
                outline: none !important;
            }

            .search-input input { 
                background: var(--bg-input) !important;
                color: var(--text-primary) !important;
                border: 2px solid var(--border-color) !important;
                border-radius: 12px !important;
                padding: 12px 16px !important;
            }

            /* ========================================
               📊 DROPDOWN ROTATION PÉRIODE (CRITIQUE)
               ======================================== */
            .dark-dropdown {
                background: var(--bg-input) !important;
            }

            /* ========================================
               DROPDOWN ROTATION COMPACT
               ======================================== */
            .rotation-dropdown-compact .Select-control {
                background: #ffffff !important;
                border: 1px solid #e2e8f0 !important;
                border-radius: 6px !important;
                min-height: 32px !important;
                height: 32px !important;
            }

            .rotation-dropdown-compact .Select-value-label {
                color: #1e293b !important;
                font-weight: 600 !important;
                font-size: 13px !important;
                line-height: 30px !important;
            }

            .rotation-dropdown-compact .Select-arrow-zone {
                padding: 4px 8px !important;
            }

            .rotation-dropdown-compact .Select-menu-outer {
                background: #ffffff !important;
                border: 1px solid #e2e8f0 !important;
                border-radius: 6px !important;
                box-shadow: 0 4px 12px rgba(0,0,0,0.1) !important;
            }

            .rotation-dropdown-compact .Select-option {
                padding: 8px 12px !important;
                font-size: 13px !important;
                color: #1e293b !important;
            }

            .rotation-dropdown-compact .Select-option:hover,
            .rotation-dropdown-compact .Select-option.is-focused {
                background: #f0f9ff !important;
                color: #0ea5e9 !important;
            }

            /* ========================================
               TABLE - SCROLL HORIZONTAL UNIQUEMENT
               ======================================== */
            .dash-table-container {
                overflow-x: auto !important;
                overflow-y: visible !important;
            }

            .dash-table-container .dash-spreadsheet-container {
                overflow-x: auto !important;
                overflow-y: visible !important;
            }

            .dash-table-container .dash-spreadsheet-inner {
                table-layout: fixed !important;
            }

            /* Empêcher le redimensionnement lors de la sélection */
            .dash-table-container td {
                box-sizing: border-box !important;
            }

            /* Checkbox de sélection - taille fixe */
            .dash-table-container .dash-select-cell {
                width: 40px !important;
                min-width: 40px !important;
                max-width: 40px !important;
                text-align: center !important;
            }

            /* Ligne sélectionnée - pas de changement de largeur */
            .dash-table-container tr.row-selected {
                outline: none !important;
            }

            /* ✅ SCROLLBAR HORIZONTAL */
            .dash-table-container::-webkit-scrollbar {
                height: 10px;
            }

            .dash-table-container::-webkit-scrollbar-track {
                background: #f1f5f9;
                border-radius: 5px;
            }

            .dash-table-container::-webkit-scrollbar-thumb {
                background: #94a3b8;
                border-radius: 5px;
            }

            .dash-table-container::-webkit-scrollbar-thumb:hover {
                background: #64748b;
            }

            /* ✅ CURSEUR NAVIGATION */
            .dash-table-container .dash-cell {
                cursor: pointer !important;
            }

            .dash-table-container .dash-cell:hover {
                background-color: #f0f9ff !important;
            }

            /* ========================================
               AUTRES DROPDOWNS (FILTRES)
               ======================================== */
            #filter-supplier .Select-control,
            #filter-category .Select-control,
            #filter-need .Select-control,
            #filter-agent .Select-conytrol {
                background: var(--bg-input) !important;
                color: var(--text-primary) !important;
                border: 2px solid var(--border-color) !important;
                border-radius: 10px !important;
            }

            #filter-supplier .Select-value-label,
            #filter-category .Select-value-label,
            #filter-need .Select-value-label,
            #filter-agent .Select-value-label,
            #filter-supplier .Select-placeholder,
            #filter-category .Select-placeholder,
            #filter-need .Select-placeholder,
            #filter-agent .Select-placeholder {
                color: var(--text-primary) !important;
            }

            #filter-supplier .Select-menu-outer,
            #filter-category .Select-menu-outer,
            #filter-need .Select-menu-outer,
            #filter-agent .Select-menu-outer {
                background: var(--bg-tertiary) !important;
                border: 2px solid var(--border-color) !important;
            }

            #filter-supplier .Select-option,
            #filter-category .Select-option,
            #filter-need .Select-option,
            #filter-agent .Select-option {
                background: var(--bg-tertiary) !important;
                color: var(--text-primary) !important;
            }

            #filter-supplier .Select-option:hover,
            #filter-category .Select-option:hover,
            #filter-need .Select-option:hover,
            #filter-agent .Select-option:hover  {
                background: rgba(34, 211, 238, 0.15) !important;
                color: var(--brand-accent) !important;
            }

            /* ========================================
               LABELS & TITRES
               ======================================== */
            label {
                color: var(--text-primary) !important;
                font-weight: 600 !important;
                font-size: 14px !important;
                margin-bottom: 8px !important;
                display: block;
            }

            .section-title { 
                font-weight: 700; 
                font-size: 18px; 
                margin-bottom: 10px; 
                color: var(--brand-accent) !important;
                letter-spacing: 0.5px;
            }

            h2, h3, h4 {
                color: var(--text-primary) !important;
                font-weight: 700 !important;
            }

            /* ========================================
               TABLEAU
               ======================================== */
            .dash-table-container {
                background: var(--bg-secondary) !important;
                border-radius: 16px !important;
                padding: 16px !important;
                box-shadow: 0 4px 16px rgba(0, 0, 0, 0.06) !important;
                border: 1px solid var(--border-color) !important;
            }

            .dash-header {
                background: linear-gradient(135deg, #1e3a5f 0%, #0f2744 100%) !important;
                color: #ffffff !important;
                font-weight: 700 !important;
                text-transform: uppercase !important;
                letter-spacing: 0.5px !important;
                font-size: 11px !important;
                border-bottom: 2px solid var(--brand-accent) !important;
            }

            .dash-cell {
                background: var(--bg-secondary) !important;
                color: var(--text-primary) !important;
                border-color: var(--border-color) !important;
                font-size: 13px !important;
            }

            .dash-table-container tr:nth-child(even) .dash-cell {
                background: var(--bg-tertiary) !important;
            }

            /* HOVER CLAIR ET VISIBLE */
            .dash-table-container tr:hover .dash-cell {
                background: #e0f2fe !important;
                border-color: #0ea5e9 !important;
                cursor: pointer;
            }

            /* Ligne sélectionnée - TRÈS VISIBLE */
            .dash-table-container tr.row-selected .dash-cell,
            .dash-table-container .dash-cell.cell-selected {
                background: #0ea5e9 !important;
                color: #ffffff !important;
                font-weight: 600 !important;
            }

            /* ========================================
               AMÉLIORATION GRAPHIQUES
               ======================================== */
            .js-plotly-plot .plotly .main-svg {
                background: transparent !important;
            }

            .js-plotly-plot .plotly text {
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
            }

            .js-plotly-plot .plotly .gtitle {
                font-weight: 700 !important;
                font-size: 16px !important;
                fill: #1e293b !important;
            }

            .js-plotly-plot .plotly .xtick text,
            .js-plotly-plot .plotly .ytick text {
                font-size: 11px !important;
                fill: #64748b !important;
            }

            .js-plotly-plot .plotly .legendtext {
                font-size: 12px !important;
                fill: #475569 !important;
            }

            /* ========================================
               CHAT (LIGHT THEME PRÉSERVÉ)
               ======================================== */
            .chat-fab {
                position: fixed; 
                right: 24px; 
                bottom: 24px; 
                z-index: 10000;
                border-radius: 9999px; 
                padding: 12px 16px; 
                border: none;
                background: linear-gradient(135deg, #0ea5e9 0%, #6366f1 100%); 
                color: #ffffff; 
                font-weight: 800;
                box-shadow: 0 12px 24px rgba(14, 165, 233, 0.35);
                display: flex; 
                align-items: center; 
                gap: 8px; 
                cursor: pointer;
                transition: all 0.3s ease;
            }

            .chat-fab:hover {
                transform: translateY(-2px);
                box-shadow: 0 16px 32px rgba(14, 165, 233, 0.4);
            }

            .chat-window {
                position: fixed; 
                right: 24px; 
                bottom: 92px; 
                width: 520px; 
                max-width: 96vw;
                background: #ffffff; 
                border: 1px solid #e5e7eb; 
                border-radius: 16px;
                z-index: 10000; 
                box-shadow: 0 24px 48px rgba(0, 0, 0, 0.15);
                display: flex; 
                flex-direction: column; 
                overflow: hidden; 
                color: #111827;
            }

            .chat-header {
                padding: 10px 12px; 
                display: flex; 
                align-items: center; 
                justify-content: space-between;
                background: #f3f4f6; 
                border-bottom: 1px solid #e5e7eb; 
                color: #111827;
            }

            .chat-body {
                padding: 12px; 
                max-height: 56vh; 
                overflow: auto; 
                display: flex; 
                flex-direction: column; 
                gap: 10px;
                background: #fafafa;
            }

            .chat-input-wrap {
                padding: 10px; 
                background: #f9fafb; 
                border-top: 1px solid #e5e7eb; 
                display: grid;
                grid-template-columns: 1fr auto; 
                gap: 10px;
            }

            .chat-textarea { 
                background: #ffffff; 
                color: #111827; 
                border: 1px solid #d1d5db; 
                border-radius: 10px; 
            }

            .chat-avatar {
                min-width: 32px; 
                height: 32px; 
                border-radius: 50%; 
                display: flex; 
                align-items: center; 
                justify-content: center;
                font-size: 16px; 
                color: #fff; 
                background: #6366f1; 
                box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1);
            }

            .chat-avatar.user { 
                background: #3b82f6; 
            }

            .chat-bubble { 
                display: flex; 
                gap: 10px; 
            }

            .chat-bubble.user { 
                flex-direction: row; 
            }

            .chat-bubble.bot { 
                flex-direction: row-reverse; 
            }

            .chat-msg { 
                background: #ffffff; 
                border: 1px solid #e5e7eb; 
                border-radius: 12px; 
                padding: 10px 12px; 
                max-width: 90%; 
            }

            .upload-box {
                border: 2px dashed #9ca3af; 
                border-radius: 10px; 
                padding: 10px; 
                text-align: center; 
                color: #6b7280; 
                background: #ffffff;
            }

            /* ========================================
               SCROLLBARS
               ======================================== */
            ::-webkit-scrollbar {
                width: 10px;
                height: 10px;
            }

            ::-webkit-scrollbar-track {
                background: var(--bg-secondary);
                border-radius: 10px;
            }

            ::-webkit-scrollbar-thumb {
                background: linear-gradient(135deg, var(--brand-accent) 0%, var(--brand-hover) 100%);
                border-radius: 10px;
            }

            ::-webkit-scrollbar-thumb:hover {
                background: linear-gradient(135deg, var(--brand-hover) 0%, #0891b2 100%);
            }

            /* ========================================
               NAVBAR NAVIGATION LINKS
               ======================================== */
            .nav-link-hover:hover > div {
                background: rgba(14, 165, 233, 0.15) !important;
                color: #0ea5e9 !important;
            }

            .nav-link-hover:active > div {
                background: rgba(14, 165, 233, 0.25) !important;
            }

            /* ========================================
               SIDEBAR - DROPDOWNS VISIBLES
               ======================================== */
            .sidebar {
                background: linear-gradient(180deg, #1e293b 0%, #0f172a 100%) !important;
                width: 240px !important;
            }

            .sidebar label {
                color: #e2e8f0 !important;
                font-weight: 500 !important;
            }

            /* Tous les dropdowns dans le sidebar */
            .sidebar .Select-control,
            .sidebar .Select--single > .Select-control,
            .sidebar .Select--multi > .Select-control,
            .sidebar div[class*="control"] {
                background-color: #1e293b !important;
                border: 1px solid #475569 !important;
                border-radius: 6px !important;
                min-height: 32px !important;
            }

            /* Texte dans les dropdowns */
            .sidebar .Select-value-label,
            .sidebar .Select-placeholder,
            .sidebar div[class*="singleValue"],
            .sidebar div[class*="placeholder"],
            .sidebar input[class*="Input"],
            .sidebar .Select-input > input {
                color: #1e293b !important;
            }

            /* Menu déroulant */
            .sidebar .Select-menu-outer,
            .sidebar div[class*="menu"] {
                background-color: #fff !important;
                border: 1px solid #e2e8f0 !important;
                border-radius: 6px !important;
                z-index: 9999 !important;
                box-shadow: 0 4px 12px rgba(0,0,0,0.15) !important;
            }

            /* Options dans le menu */
            .sidebar .Select-option,
            .sidebar div[class*="option"] {
                background-color: #fff !important;
                color: #1e293b !important;
                padding: 8px 12px !important;
            }

            .sidebar .Select-option:hover,
            .sidebar .Select-option.is-focused,
            .sidebar div[class*="option"]:hover {
                background-color: #f1f5f9 !important;
                color: #0f172a !important;
            }

            /* Tags sélectionnés (multi-select) */
            .sidebar .Select-value,
            .sidebar div[class*="multiValue"] {
                background-color: #0ea5e9 !important;
                border-color: #0ea5e9 !important;
                color: #fff !important;
                border-radius: 4px !important;
            }

            .sidebar .Select-value-label,
            .sidebar div[class*="multiValue"] span {
                color: #fff !important;
            }

            /* Flèche et clear */
            .sidebar .Select-arrow-zone,
            .sidebar .Select-clear-zone,
            .sidebar div[class*="indicator"] {
                color: #64748b !important;
            }

            /* Input recherche sidebar */
            .sidebar input[type="text"],
            .sidebar .form-control {
                background: #334155 !important;
                border: 1px solid #475569 !important;
                color: #f8fafc !important;
            }

            .sidebar input[type="text"]::placeholder,
            .sidebar .form-control::placeholder {
                color: #94a3b8 !important;
            }

            /* Switch sidebar */
            .sidebar .form-check-label {
                color: #e2e8f0 !important;
            }

            /* ========================================
               RESPONSIVE
               ======================================== */
            @media (max-width: 992px) {
                .sidebar {
                    width: 200px !important;
                }
                #page-container {
                    margin-left: 200px !important;
                }
            }

            @media (max-width: 768px) {
                .sidebar {
                    display: none !important;
                }
                #page-container {
                    margin-left: 0 !important;
                    padding: 12px !important;
                }
                nav > div:nth-child(2) {
                    display: none !important; /* Cacher nav links sur mobile */
                }
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

# NOTE: get_df_cached est défini au début du fichier (ligne ~294) avec @cache.memoize
# NE PAS redéfinir ici pour éviter les conflits

# ------------------------------ Sidebar ------------------------------------------
def make_sidebar():
    """
    Sidebar simplifié - Filtres uniquement.
    La navigation est dans la navbar horizontale.
    """
    df = get_df_cached()

    # Extraction des valeurs uniques pour les dropdowns
    suppliers = sorted(
        [s for s in df['Supplier'].dropna().unique().tolist() if s != '']) if 'Supplier' in df.columns else []
    cats = sorted([c for c in df['Product Category'].dropna().unique().tolist() if
                   c != '']) if 'Product Category' in df.columns else []

    # Charger la liste des agents
    agents_list = []
    try:
        agents_df = load_agents_suppliers()
        if not agents_df.empty:
            owner_col = 'Owner' if 'Owner' in agents_df.columns else None
            if not owner_col:
                for col in agents_df.columns:
                    if 'owner' in col.lower() or 'agent' in col.lower():
                        owner_col = col
                        break
            if owner_col:
                agents_list = sorted([a for a in agents_df[owner_col].dropna().unique().tolist()
                                      if a and str(a).strip() != '' and str(a) != 'nan'])
    except Exception as e:
        print(f"⚠️ Erreur chargement agents pour sidebar: {e}")

    return html.Div(className="sidebar", style={
        "position": "fixed",
        "top": "50px",
        "left": "0",
        "width": "240px",
        "height": "calc(100vh - 56px)",
        "background": "linear-gradient(180deg, #1e293b 0%, #0f172a 100%)",
        "borderRight": "1px solid #334155",
        "padding": "16px",
        "overflowY": "auto",
        "zIndex": "100"
    }, children=[

        # Recherche
        html.Div([
            html.Div("RECHERCHE", style={
                "color": "#94a3b8", "fontSize": "10px", "fontWeight": "600",
                "textTransform": "uppercase", "letterSpacing": "1px", "marginBottom": "6px"
            }),
            dbc.Input(
                id="search-input", placeholder="Rechercher...", type="text", debounce=True,
                style={"background": "#334155", "border": "1px solid #475569", "color": "#f8fafc",
                       "fontSize": "12px", "borderRadius": "6px", "padding": "8px 10px"}
            )
        ], style={"marginBottom": "16px"}),

        # Filtres
        html.Div([
            html.Div("FILTRES", style={
                "color": "#94a3b8", "fontSize": "10px", "fontWeight": "600",
                "textTransform": "uppercase", "letterSpacing": "1px", "marginBottom": "10px"
            }),

            # Fournisseur
            html.Div([
                html.Label("Fournisseur", style={"color": "#e2e8f0", "fontSize": "11px", "fontWeight": "500", "marginBottom": "4px", "display": "block"}),
                dcc.Dropdown(
                    id="filter-supplier",
                    options=[{"label": s, "value": s} for s in suppliers],
                    multi=True, placeholder="Tous", persistence=True,
                    style={"fontSize": "11px"}
                )
            ], style={"marginBottom": "10px"}),

            # Catégorie
            html.Div([
                html.Label("Catégorie", style={"color": "#e2e8f0", "fontSize": "11px", "fontWeight": "500", "marginBottom": "4px", "display": "block"}),
                dcc.Dropdown(
                    id="filter-category",
                    options=[{"label": c, "value": c} for c in cats],
                    multi=True, placeholder="Toutes", persistence=True,
                    style={"fontSize": "11px"}
                )
            ], style={"marginBottom": "10px"}),

            # Besoin
            html.Div([
                html.Label("Besoin", style={"color": "#e2e8f0", "fontSize": "11px", "fontWeight": "500", "marginBottom": "4px", "display": "block"}),
                dcc.Dropdown(
                    id="filter-need",
                    options=[
                        {"label": "ORDER NOW", "value": "ORDER NOW"},
                        {"label": "ORDER NOT URGENT", "value": "ORDER NOT URGENT"},
                        {"label": "NO NEED", "value": "NO NEED"}
                    ],
                    multi=True, placeholder="Tous", persistence=True,
                    style={"fontSize": "11px"}
                )
            ], style={"marginBottom": "10px"}),

            # Agent
            html.Div([
                html.Label("Agent", style={"color": "#e2e8f0", "fontSize": "11px", "fontWeight": "500", "marginBottom": "4px", "display": "block"}),
                dcc.Dropdown(
                    id="filter-agent",
                    options=[{"label": "Tous", "value": "all"}] + [{"label": a, "value": a} for a in agents_list],
                    value="all", placeholder="Tous", persistence=True,
                    style={"fontSize": "11px"}
                )
            ], style={"marginBottom": "10px"}),

            # Regroupement
            dbc.Checklist(
                options=[{"label": " Regrouper par produit", "value": "by_product"}],
                value=[], id="toggle-options", switch=True,
                style={"fontSize": "11px", "color": "#cbd5e1"}
            )
        ], style={"marginBottom": "14px"}),

        html.Hr(style={"borderColor": "#334155", "margin": "12px 0"}),

        # Footer
        html.Div([
            html.Small(f"© {datetime.now().year} • {AUTHOR}", style={"color": "#64748b", "fontSize": "9px"})
        ], style={"position": "absolute", "bottom": "12px", "left": "16px", "right": "16px"}),

        html.Div(id="debug-info", style={"display": "none"}),
        dcc.Download(id="download-data"),
        dcc.Download(id="download-po"),

        dbc.Modal([
            dbc.ModalBody([
                html.Div([
                    dbc.Spinner(color="primary", size="lg"),
                    html.H5("Génération du BC...", className="mt-3", style={"color": "#0ea5e9"}),
                    html.P("Patientez...", style={"color": "#9ca3af", "fontSize": "11px"})
                ], style={"textAlign": "center", "padding": "16px"})
            ], style={"background": "#f8fafc"})
        ], id="po-loading-overlay", is_open=False, centered=True, backdrop="static", keyboard=False),
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
        'optimal stock', 'Max Lead Time', 'Avg Daily Sales', 'Buffer Value',
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
    if {'product_name', 'Supplier', 'total_stock'}.issubset(tmp.columns):
        idx_max = tmp.groupby('product_name')['total_stock'].idxmax()
        main_sup = tmp.loc[idx_max, ['product_name', 'Supplier']].rename(columns={'Supplier': 'Main Supplier'})
    else:
        main_sup = pd.DataFrame(columns=['product_name', 'Main Supplier'])

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
    for c in ['total_stock', 'Predicted Order Quantity', 'purchase_need', 'MOQ MAAD', 'QAC']:
        if c in tmp.columns: agg_map[c] = 'sum'

    # Max (couverture, LT, OOS, buffers…)
    for c in [
        'optimal stock', 'Max Lead Time', 'Avg Daily Sales', 'Buffer Value',
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



# ============================================================
# FONCTION CENTRALE DE CALCUL DES KPIS - COHÉRENCE GARANTIE
# ============================================================
def calculate_stock_kpis(df: pd.DataFrame) -> dict:
    """
    Fonction UNIQUE et RIGOUREUSE pour calculer les KPIs de stock.
    Utilisée par Overview ET Agents pour garantir la cohérence ABSOLUE.

    DÉFINITIONS STRICTES:
    - SKUs: Produits ACTIFS uniquement (is_active = True)
    - Rupture (Out of Stock): stock actuel <= 0
    - À risque de rupture: produits avec proba ML >= 0.5 OU couverture < 7 jours

    Returns:
        dict avec: total_skus, out_of_stock, at_risk, suppliers
    """
    if df is None or df.empty:
        return {
            'total_skus': 0,
            'out_of_stock': 0,
            'at_risk': 0,
            'suppliers': 0,
            'method_rupture': 'N/A',
            'method_prediction': 'N/A'
        }

    df_clean = df.copy()

    # ========================================
    # 🔍 FILTRAGE 1: Exclure les produits delisted
    # ========================================
    if 'delisting_status' in df_clean.columns:
        delist = df_clean['delisting_status'].astype(str).str.lower()
        mask_valid = delist != 'delisted'
        df_clean = df_clean.loc[mask_valid].copy()

    # ========================================
    # 🔍 FILTRAGE 2: Produits ACTIFS uniquement (is_active = True)
    # ========================================
    if 'is_active' in df_clean.columns:
        # Convertir en booléen de façon robuste
        def parse_is_active(x):
            if pd.isna(x):
                return True  # Par défaut actif si non spécifié
            if isinstance(x, bool):
                return x
            return str(x).lower() in ('true', '1', 'yes', 'oui', 'actif')

        df_clean['_is_active'] = df_clean['is_active'].apply(parse_is_active)
        total_before = len(df_clean)
        df_clean = df_clean[df_clean['_is_active'] == True].copy()
        total_after = len(df_clean)
        if total_before != total_after:
            print(f"   🔍 Filtre is_active: {total_before} → {total_after} produits actifs")

    # KPI 1 : SKUs uniques (ACTIFS uniquement)
    total_skus = df_clean['product_name'].nunique() if 'product_name' in df_clean.columns else len(df_clean)

    # KPI 2 : RUPTURES RÉELLES (Out of Stock) - stock <= 0
    out_of_stock = 0
    method_rupture = 'N/A'

    if 'total_stock' in df_clean.columns:
        stock_values = pd.to_numeric(df_clean['total_stock'], errors='coerce').fillna(0)
        out_of_stock = int((stock_values <= 0).sum())
        method_rupture = 'total_stock <= 0'
    elif 'Stock Status' in df_clean.columns:
        out_of_stock = int((df_clean['Stock Status'] == 'Out of Stock').sum())
        method_rupture = 'Stock Status == Out of Stock'

    # KPI 3 : PRODUITS À RISQUE (Prédiction)
    at_risk = 0
    method_prediction = 'N/A'

    # Préparer colonnes numériques
    if 'total_stock' in df_clean.columns:
        df_clean['_stock'] = pd.to_numeric(df_clean['total_stock'], errors='coerce').fillna(0)
    else:
        df_clean['_stock'] = 0

    if 'Average Daily Sales' in df_clean.columns:
        df_clean['_ads'] = pd.to_numeric(df_clean['Average Daily Sales'], errors='coerce').fillna(0.01).replace(0, 0.01)
    else:
        df_clean['_ads'] = 0.01

    if 'optimal stock' in df_clean.columns:
        df_clean['_optimal'] = pd.to_numeric(df_clean['optimal stock'], errors='coerce').fillna(0)
    else:
        df_clean['_optimal'] = 0

    df_clean['_coverage'] = df_clean['_stock'] / df_clean['_ads']

    # MÉTHODE 1: ML Stockout Probability >= 0.5
    if 'Stockout Probability' in df_clean.columns:
        probs = pd.to_numeric(df_clean['Stockout Probability'], errors='coerce').fillna(0)
        non_zero = (probs > 0).sum()

        if non_zero > 10:
            # Exclure les produits déjà en rupture (stock <= 0)
            at_risk = int(((probs >= 0.5) & (df_clean['_stock'] > 0)).sum())
            method_prediction = f'ML Probability >= 0.5'

    # MÉTHODE 2: Règle métier si ML non disponible
    if at_risk == 0:
        # Couverture < 7 jours ET stock > 0 (pas déjà en rupture)
        risk_mask = (df_clean['_stock'] > 0) & (df_clean['_coverage'] < 7) & (df_clean['_coverage'] > 0)
        at_risk = int(risk_mask.sum())
        method_prediction = 'Couverture < 7 jours'

    # KPI 4 : Fournisseurs
    suppliers = df_clean['Supplier'].nunique() if 'Supplier' in df_clean.columns else 0

    print(f"   📊 KPIs: {total_skus} SKUs actifs | {out_of_stock} ruptures | {at_risk} à risque | {suppliers} fournisseurs")

    # 🚀 Tracking Supabase (asynchrone, non-bloquant)
    try:
        total_in_df = len(df) if df is not None else 0
        track_stock_stats(total_in_df, total_skus, out_of_stock, at_risk, suppliers)
    except:
        pass  # Ne jamais bloquer sur le tracking

    return {
        'total_skus': total_skus,
        'out_of_stock': out_of_stock,
        'at_risk': at_risk,
        'suppliers': suppliers,
        'method_rupture': method_rupture,
        'method_prediction': method_prediction
    }


# ------------------------------ Pages --------------------------------------------
def make_kpis(df: pd.DataFrame):
    """
    Calcule les KPIs + liste des produits à risque.
    UTILISE calculate_stock_kpis() pour COHÉRENCE avec page Agents.
    """

    # ---------- helpers ----------
    def fmt(n):
        if pd.isna(n):
            return "-"
        if isinstance(n, (int, float)):
            try:
                return f"{n:,.0f}".replace(",", " ")
            except Exception:
                return str(n)
        return str(n)

    def as_float(s, default=0.0):
        return pd.to_numeric(s, errors="coerce").fillna(default)

    # ==========================================
    # UTILISER LA FONCTION COMMUNE - COHÉRENCE GARANTIE
    # ==========================================
    stock_kpis = calculate_stock_kpis(df)

    total_skus = stock_kpis['total_skus']
    out_of_stock = stock_kpis['out_of_stock']
    suppliers = stock_kpis['suppliers']
    risk_count = stock_kpis['at_risk']  # Utiliser la même valeur que page Agents

    # Préparer df_valid pour construire la liste des produits à afficher
    if 'delisting_status' in df.columns:
        delist = df['delisting_status'].astype(str).str.lower()
        mask_valid = delist != 'delisted'
        df_valid = df.loc[mask_valid].copy()
    else:
        df_valid = df.copy()

    # ========================================
    # CONSTRUIRE LA LISTE DES PRODUITS À RISQUE (pour dropdown)
    # ========================================
    print("\n" + "=" * 60)
    print("🤖 CONSTRUCTION LISTE PRODUITS À RISQUE")
    print("=" * 60)

    risk_products = []

    # Préparer colonnes
    if 'total_stock' in df_valid.columns:
        df_valid['_stock'] = pd.to_numeric(df_valid['total_stock'], errors='coerce').fillna(0)
    else:
        df_valid['_stock'] = 0

    if 'Average Daily Sales' in df_valid.columns:
        df_valid['_ads'] = pd.to_numeric(df_valid['Average Daily Sales'], errors='coerce').fillna(0.01).replace(0, 0.01)
    else:
        df_valid['_ads'] = 0.01

    df_valid['_coverage'] = df_valid['_stock'] / df_valid['_ads']

    # MÉTHODE 1 : ML Stockout Probability >= 0.5
    if 'Stockout Probability' in df_valid.columns:
        probs = as_float(df_valid['Stockout Probability'])
        non_zero = (probs > 0).sum()

        if non_zero > 10:
            print("✅ Utilisation ML (Stockout Probability >= 0.5)")
            # Même critère que calculate_stock_kpis
            risk_mask = (probs >= 0.5) & (df_valid['_stock'] > 0)
            risk_df = df_valid.loc[risk_mask].copy()
            risk_df = risk_df.sort_values('Stockout Probability', ascending=False)

            for _, row in risk_df.head(50).iterrows():
                prob = float(row.get('Stockout Probability', 0))
                risk_products.append({
                    'product_name': str(row.get('product_name', 'N/A'))[:60],
                    'supplier': str(row.get('Supplier', 'N/A')),
                    'coverage_days': float(row.get('_coverage', 0) or 0),
                    'stock': float(row.get('_stock', 0) or 0),
                    'ads': float(row.get('_ads', 0) or 0),
                    'category': str(row.get('Product Category', 'N/A')),
                    'ml_probability': prob,
                    'risk_level': 'HIGH' if prob >= 0.7 else 'MEDIUM'
                })

    # MÉTHODE 2 : Règle métier si ML non disponible
    if len(risk_products) == 0:
        print("⚠️ Fallback: Couverture < 7 jours")
        # Même critère que calculate_stock_kpis
        risk_mask = (df_valid['_stock'] > 0) & (df_valid['_coverage'] < 7) & (df_valid['_coverage'] > 0)
        risk_df = df_valid.loc[risk_mask].copy()
        risk_df = risk_df.sort_values('_coverage', ascending=True)

        for _, row in risk_df.head(50).iterrows():
            coverage = float(row.get('_coverage', 0) or 0)
            pseudo_prob = 0.7 if coverage < 3 else 0.5
            risk_products.append({
                'product_name': str(row.get('product_name', 'N/A'))[:60],
                'supplier': str(row.get('Supplier', 'N/A')),
                'coverage_days': coverage,
                'stock': float(row.get('_stock', 0) or 0),
                'ads': float(row.get('_ads', 0) or 0),
                'category': str(row.get('Product Category', 'N/A')),
                'ml_probability': pseudo_prob,
                'risk_level': 'HIGH' if coverage < 3 else 'MEDIUM'
            })

    print(f"   📋 Produits dans dropdown: {len(risk_products)} (total à risque: {risk_count})")
    print("=" * 60 + "\n")

    # ---------- Cartes KPI ----------
    cards = dbc.Row([
        dbc.Col(html.Div(className="kpi", children=[
            html.Small("SKUs"),
            html.H3(fmt(total_skus))
        ]), md=4),
        dbc.Col(html.Div(className="kpi", children=[
            html.Small("Produits en rupture"),
            html.H3(fmt(out_of_stock), style={"color": "#ef4444" if out_of_stock > 0 else "#10b981"})
        ]), md=4),
        dbc.Col(html.Div(className="kpi", children=[
            html.Small("Fournisseurs actifs"),
            html.H3(fmt(suppliers))
        ]), md=4),
    ], className="gy-3")

    # ---------- Dropdown avec vraie proba ML + NAVIGATION CLIQUABLE ----------
    dropdown_children = []

    if risk_products:
        for idx, prod in enumerate(risk_products):
            # ✅ Couleur basée sur ML probability
            ml_prob = prod.get('ml_probability', 0)

            if ml_prob > 0.7:
                color_border = '#ef4444'  # Rouge
                badge_text = f"URGENT ({ml_prob:.0%})"
                badge_color = "danger"
            elif ml_prob > 0.3:
                color_border = '#f59e0b'  # Orange
                badge_text = f"À surveiller ({ml_prob:.0%})"
                badge_color = "warning"
            else:
                color_border = '#3b82f6'  # Bleu
                badge_text = f"Risque faible ({ml_prob:.0%})"
                badge_color = "info"

            # ✅ NOUVEAU: Carte cliquable avec ID pour navigation
            dropdown_children.append(
                html.Div(
                    id={'type': 'risk-product-item', 'index': idx},
                    n_clicks=0,
                    style={
                        "padding": "12px",
                        "marginBottom": "8px",
                        "background": "#f8fafc",
                        "border": "1px solid #1f2937",
                        "borderLeft": f"4px solid {color_border}",
                        "borderRadius": "8px",
                        "cursor": "pointer",
                        "transition": "all 0.2s ease",
                    },
                    className="risk-product-card",
                    children=[
                        # ✅ Stocker le nom du produit pour le callback
                        dcc.Store(
                            id={'type': 'risk-product-name', 'index': idx},
                            data=prod['product_name']
                        ),

                        # Header avec nom et badge
                        html.Div([
                            html.Strong(prod['product_name'], style={
                                "color": "#0ea5e9",
                                "fontSize": "13px",
                                "marginRight": "8px"
                            }),
                            dbc.Badge(badge_text, color=badge_color, pill=True, style={"fontSize": "9px"}),
                            # ✅ NOUVEAU: Icône de navigation
                            html.Span("→", style={
                                "marginLeft": "auto",
                                "color": "#6b7280",
                                "fontSize": "14px",
                                "fontWeight": "bold"
                            })
                        ], style={"marginBottom": "6px", "display": "flex", "alignItems": "center"}),

                        html.Div([
                            html.Small(f"📦 {prod['supplier']}", style={"color": "#9ca3af", "fontSize": "11px"}),
                            html.Span(" • ", style={"color": "#4b5563"}),
                            html.Small(f"Stock: {prod['stock']:.0f}", style={"color": "#9ca3af", "fontSize": "11px"}),
                        ], style={"marginBottom": "4px"}),

                        html.Div([
                            html.Small(
                                f"⏱️ Couverture: {prod['coverage_days']:.1f} jours",
                                style={
                                    "color": "#ef4444" if prod['coverage_days'] < 7 else "#f59e0b",
                                    "fontSize": "11px",
                                    "fontWeight": "700",
                                    "marginRight": "12px"
                                }
                            ),
                            html.Small(f"📈 ADS: {prod['ads']:.1f}/j", style={"color": "#6b7280", "fontSize": "11px"}),
                        ], style={"marginBottom": "6px"}),

                        html.Div([
                            dbc.Badge(prod['category'], color="secondary", pill=True, className="me-1",
                                      style={"fontSize": "9px"}),
                            dbc.Badge(
                                f"🤖 ML: {ml_prob:.0%}",
                                color="danger" if ml_prob > 0.7 else "warning" if ml_prob > 0.3 else "info",
                                pill=True,
                                style={"fontSize": "9px"}
                            ),
                        ]),

                        # ✅ Hint de clic
                        html.Small("Cliquer pour voir les détails", style={
                            "color": "#4b5563",
                            "fontSize": "10px",
                            "marginTop": "6px",
                            "display": "block",
                            "fontStyle": "italic"
                        })
                    ]
                )
            )
    else:
        dropdown_children = [
            html.Div(
                children=[
                    html.Span("✅", style={"fontSize": "24px", "marginBottom": "8px"}),
                    html.Div("Aucun produit à risque détecté", style={
                        "color": "#10b981",
                        "fontSize": "13px",
                        "fontWeight": "600"
                    }),
                    html.Small("Tous les produits ont une couverture adéquate", style={
                        "color": "#6b7280",
                        "fontSize": "11px",
                        "marginTop": "4px"
                    })
                ],
                style={
                    "textAlign": "center",
                    "padding": "30px 20px",
                    "display": "flex",
                    "flexDirection": "column",
                    "alignItems": "center"
                }
            )
        ]

    # ---------- Badge avec ML ----------
    bell = html.Div(
        className="pill",
        style={"marginTop": "10px"},
        children=[
            dbc.Button([
                html.Span("🤖", style={"fontSize": "16px", "marginRight": "8px"}),
                html.B("À risque de rupture : "),
                dbc.Badge(
                    f"{risk_count}",
                    color="danger" if any(p.get('ml_probability', 0) > 0.7 for p in
                                          risk_products) else "warning" if risk_count > 0 else "success",
                    pill=True,
                    className="ms-2"
                )
            ],
                id="risk-alert-toggle",
                color="link",
                className="p-2 text-start w-100",
                style={
                    "textDecoration": "none",
                    "color": "#1e293b",
                    "border": "1px solid #374151",
                    "borderRadius": "10px",
                    "background": "rgba(239, 68, 68, 0.1)" if any(p.get('ml_probability', 0) > 0.7 for p in
                                                                  risk_products) else "rgba(245, 158, 11, 0.1)" if risk_count > 0 else "rgba(16, 185, 129, 0.1)",
                    "cursor": "pointer" if risk_count > 0 else "not-allowed"
                },
                disabled=(risk_count == 0)
            ),

            dbc.Collapse(
                id="risk-alert-collapse",
                is_open=False,
                children=html.Div(
                    dropdown_children + [
                        html.Hr(style={"borderColor": "#374151", "margin": "12px 0"}),
                        html.Div([
                            html.Small(
                                f" Affichant {min(len(risk_products), 50)} produit(s)",
                                style={"color": "#6b7280", "fontSize": "10px", "marginRight": "10px"}
                            ),
                            html.Small(
                                "• 🤖 Calculé avec ML RandomForest",
                                style={"color": "#4b5563", "fontSize": "10px"}
                            )
                        ], style={"display": "flex", "alignItems": "center", "justifyContent": "center"})
                    ],
                    style={
                        "maxHeight": "450px",
                        "overflowY": "auto",
                        "marginTop": "10px",
                        "padding": "12px",
                        "background": "#0a1320",
                        "border": "1px solid #1f2937",
                        "borderRadius": "10px",
                        "boxShadow": "0 4px 12px rgba(0,0,0,0.3)"
                    }
                )
            )
        ]
    )

    return cards, bell
'''
def make_kpis(df: pd.DataFrame):
    """Calcule les KPIs (SKUs, ruptures, fournisseurs) + alerte dropdown produits à risque (ML)."""

    # ---------- helpers ----------
    def fmt(n):
        if pd.isna(n):
            return "-"
        if isinstance(n, (int, float)):
            try:
                return f"{n:,.0f}".replace(",", " ")
            except Exception:
                return str(n)
        return str(n)

    def as_float(s, default=0.0):
        return pd.to_numeric(s, errors="coerce").fillna(default)

    # ---------- nettoyage : exclure les delisted sans planter sur dtypes ----------
    if 'delisting_status' in df.columns:
        delist = df['delisting_status'].astype(str).str.lower()
        mask_valid = delist != 'delisted'
        df_valid = df.loc[mask_valid].copy()
    else:
        df_valid = df.copy()

    # ---------- KPI 1 : SKUs ----------
    total_skus = df_valid['product_name'].nunique() if 'product_name' in df_valid.columns else len(df_valid)

    # ---------- KPI 2 : Ruptures réelles ----------
    if 'Stock Status' in df_valid.columns:
        out_of_stock = int((df_valid['Stock Status'] == 'Out of Stock').sum())
    elif 'total_stock' in df_valid.columns:
        out_of_stock = int(as_float(df_valid['total_stock']) <= 0.0)
    elif 'Ajusted_total_need' in df_valid.columns:
        out_of_stock = int((df_valid['Ajusted_total_need'] == 'ORDER NOW').sum())
    else:
        out_of_stock = 0

    # ---------- KPI 3 : Fournisseurs ----------
    if 'Suppliers (all)' in df_valid.columns:
        suppliers = df_valid['Suppliers (all)'].nunique()
    elif 'Supplier' in df_valid.columns:
        suppliers = df_valid['Supplier'].nunique()
    else:
        suppliers = 0

    # ---------- Produits à risque (pour dropdown) ----------
    risk_products = []

    if {'Predicted Stockout', 'Credit Adequacy Score'}.issubset(df_valid.columns):
        risk_mask = (df_valid['Predicted Stockout'] == True) & (as_float(df_valid['Credit Adequacy Score']) < 0.5)
        risk_df = df_valid.loc[risk_mask].copy()

        if 'Max Coverage Day' in risk_df.columns:
            risk_df = risk_df.sort_values('Max Coverage Day', ascending=True)

        for _, row in risk_df.head(50).iterrows():
            risk_products.append({
                'product_name': str(row.get('product_name', 'N/A'))[:60],
                'supplier'    : str(row.get('Supplier', 'N/A')),
                'coverage_days': float(row.get('Max Coverage Day', 0) or 0),
                'stock'       : float(row.get('total_stock', 0) or 0),
                'ads'         : float(row.get('Average Daily Sales', 0) or 0),
                'category'    : str(row.get('Product Category', 'N/A')),
                'credit_score': float(row.get('Credit Adequacy Score', 0) or 0),
            })

    elif {'Max Coverage Day', 'ADJUSTED_LEADTIME', 'credit_days'}.issubset(df_valid.columns):
        tmp = df_valid.assign(
            replenishment_days = as_float(df_valid['ADJUSTED_LEADTIME'], 7) + as_float(df_valid['credit_days'], 14),
            mcd                = as_float(df_valid['Max Coverage Day'])
        )
        risk_mask = (tmp['mcd'] > 0) & (tmp['mcd'] < tmp['replenishment_days'])
        risk_df = tmp.loc[risk_mask].sort_values('mcd', ascending=True)

        for _, row in risk_df.head(50).iterrows():
            risk_products.append({
                'product_name': str(row.get('product_name', 'N/A'))[:60],
                'supplier'    : str(row.get('Supplier', 'N/A')),
                'coverage_days': float(row.get('mcd', 0) or 0),
                'stock'       : float(row.get('total_stock', 0) or 0),
                'ads'         : float(row.get('Average Daily Sales', 0) or 0),
                'category'    : str(row.get('Product Category', 'N/A')),
                'credit_score': 0.3,
            })

    elif 'Ajusted_total_need' in df_valid.columns:
        tmp = df_valid.copy()
        tmp['priority'] = tmp['Ajusted_total_need'].map({'ORDER NOW': 1, 'ORDER NOT URGENT': 2}).fillna(9)
        risk_df = tmp.loc[tmp['Ajusted_total_need'].isin(['ORDER NOW', 'ORDER NOT URGENT'])] \
                     .sort_values('priority', ascending=True)

        for _, row in risk_df.head(50).iterrows():
            risk_products.append({
                'product_name': str(row.get('product_name', 'N/A'))[:60],
                'supplier'    : str(row.get('Supplier', 'N/A')),
                'coverage_days': float(row.get('Max Coverage Day', 0) or 0),
                'stock'       : float(row.get('total_stock', 0) or 0),
                'ads'         : float(row.get('Average Daily Sales', 0) or 0),
                'category'    : str(row.get('Product Category', 'N/A')),
                'credit_score': 0.3,
            })

    risk_count = len(risk_products)

    # ---------- Cartes KPI ----------
    cards = dbc.Row(
        [
            dbc.Col(html.Div(className="kpi", children=[html.Small("SKUs"), html.H3(fmt(total_skus))]), md=4),
            dbc.Col(html.Div(className="kpi", children=[
                html.Small("Produits en rupture"),
                html.H3(fmt(out_of_stock), style={"color": "#ef4444" if out_of_stock > 0 else "#10b981"})
            ]), md=4),
            dbc.Col(html.Div(className="kpi", children=[html.Small("Fournisseurs actifs"), html.H3(fmt(suppliers))]), md=4),
        ],
        className="gy-3"
    )

    # ---------- Alerte dropdown ----------
    # (pas de pseudo-sélecteurs CSS comme ':hover' dans style inline → ignorés par Dash)
    dropdown_children = []

    if risk_products:
        for prod in risk_products:
            color_border = '#ef4444' if prod['coverage_days'] < 7 else '#f59e0b'
            dropdown_children.append(
                html.Div(
                    style={
                        "padding": "10px",
                        "marginBottom": "8px",
                        "background": "#f8fafc",
                        "border": "1px solid #1f2937",
                        "borderLeft": f"4px solid {color_border}",
                        "borderRadius": "8px",
                    },
                    children=[
                        html.Div([
                            html.Strong(prod['product_name'], style={"color": "#0ea5e9", "fontSize": "13px", "marginRight": "8px"}),
                            dbc.Badge("URGENT" if prod['coverage_days'] < 7 else "À surveiller",
                                      color="danger" if prod['coverage_days'] < 7 else "warning",
                                      pill=True, style={"fontSize": "9px"})
                        ], style={"marginBottom": "6px"}),
                        html.Div([
                            html.Span(" "),
                            html.Small(f"{prod['supplier']}", style={"color": "#9ca3af", "fontSize": "11px"}),
                            html.Span(" • ", style={"color": "#4b5563"}),
                            html.Span(" "),
                            html.Small(f"Stock: {prod['stock']:.0f}", style={"color": "#9ca3af", "fontSize": "11px"}),
                        ], style={"marginBottom": "4px"}),
                        html.Div([
                            html.Span(" "),
                            html.Small(
                                f"Couverture: {prod['coverage_days']:.1f} jours",
                                style={
                                    "color": "#ef4444" if prod['coverage_days'] < 7 else "#f59e0b",
                                    "fontSize": "11px", "fontWeight": "700", "marginRight": "12px"
                                }
                            ),
                            html.Span(" "),
                            html.Small(f"ADS: {prod['ads']:.1f}/j", style={"color": "#6b7280", "fontSize": "11px"}),
                        ], style={"marginBottom": "6px"}),
                        html.Div([
                            dbc.Badge(prod['category'], color="secondary", pill=True, className="me-1", style={"fontSize": "9px"}),
                            dbc.Badge(f"Score: {prod['credit_score']:.2f}",
                                      color="danger" if prod['credit_score'] < 0.3 else "warning",
                                      pill=True, style={"fontSize": "9px"}),
                        ])
                    ]
                )
            )
    else:
        dropdown_children = [
            html.Div(
                children=[
                    html.Span("✅", style={"fontSize": "24px", "marginBottom": "8px"}),
                    html.Div("Aucun produit à risque détecté", style={"color": "#10b981", "fontSize": "13px", "fontWeight": "600"}),
                    html.Small("Tous les produits ont une couverture adéquate", style={"color": "#6b7280", "fontSize": "11px", "marginTop": "4px"})
                ],
                style={"textAlign": "center", "padding": "30px 20px", "display": "flex", "flexDirection": "column", "alignItems": "center"}
            )
        ]

    bell = html.Div(
        className="pill",
        style={"marginTop": "10px"},
        children=[
            dbc.Button(
                [
                    html.Span("🔔", style={"fontSize": "16px", "marginRight": "8px"}),
                    html.B("À risque de rupture (ML) : "),
                    dbc.Badge(f"{risk_count}", color="warning" if risk_count > 0 else "success", pill=True, className="ms-2")
                ],
                id="risk-alert-toggle",
                color="link",
                className="p-2 text-start w-100",
                style={
                    "textDecoration": "none",
                    "color": "#1e293b",
                    "border": "1px solid #374151",
                    "borderRadius": "10px",
                    "background": "rgba(245, 158, 11, 0.1)" if risk_count > 0 else "rgba(16, 185, 129, 0.1)",
                    "cursor": "pointer" if risk_count > 0 else "not-allowed"
                },
                disabled=(risk_count == 0)
            ),
            dbc.Collapse(
                id="risk-alert-collapse",
                is_open=False,
                children=html.Div(
                    dropdown_children + [
                        html.Hr(style={"borderColor": "#374151", "margin": "12px 0"}),
                        html.Div([
                            html.Small(
                                f" Affichant {min(len(risk_products), 50)} produit(s)",
                                style={"color": "#6b7280", "fontSize": "10px", "marginRight": "10px"}
                            ),
                            html.Small("• Calculé avec ML (Predicted Stockout + Credit Adequacy)",
                                       style={"color": "#4b5563", "fontSize": "10px"})
                        ], style={"display": "flex", "alignItems": "center", "justifyContent": "center"})
                    ],
                    style={
                        "maxHeight": "450px", "overflowY": "auto", "marginTop": "10px",
                        "padding": "12px", "background": "#0a1320", "border": "1px solid #1f2937",
                        "borderRadius": "10px", "boxShadow": "0 4px 12px rgba(0,0,0,0.3)"
                    }
                )
            )
        ]
    )

    return cards, bell
'''
@app.callback(
    Output("risk-alert-collapse", "is_open"),
    Input("risk-alert-toggle", "n_clicks"),
    State("risk-alert-collapse", "is_open"),
    prevent_initial_call=True
)
def _toggle_risk_dropdown(n, is_open):
    if not n:
        raise dash.exceptions.PreventUpdate
    return not is_open


# ✅ NOUVEAU CALLBACK: Navigation vers le produit cliqué dans le dropdown
@app.callback(
    [Output("search-input", "value", allow_duplicate=True),
     Output("risk-alert-collapse", "is_open", allow_duplicate=True)],
    Input({'type': 'risk-product-item', 'index': ALL}, 'n_clicks'),
    State({'type': 'risk-product-name', 'index': ALL}, 'data'),
    prevent_initial_call=True
)
def navigate_to_risk_product(n_clicks_list, product_names):
    """
    Quand on clique sur un produit à risque dans le dropdown:
    1. Ferme le dropdown
    2. Met le nom du produit dans la barre de recherche
    3. La table se filtre automatiquement sur ce produit
    """
    if not n_clicks_list or not any(n_clicks_list):
        raise dash.exceptions.PreventUpdate

    # Trouver quel produit a été cliqué
    ctx = dash.callback_context
    if not ctx.triggered:
        raise dash.exceptions.PreventUpdate

    # Extraire l'index du produit cliqué
    triggered_id = ctx.triggered[0]['prop_id']

    # Parser l'ID pour obtenir l'index
    try:
        import json as json_parser
        # Format: {"type":"risk-product-item","index":0}.n_clicks
        id_str = triggered_id.split('.')[0]
        id_dict = json_parser.loads(id_str)
        clicked_index = id_dict.get('index', 0)
    except Exception:
        clicked_index = 0

    # Récupérer le nom du produit
    if clicked_index < len(product_names):
        product_name = product_names[clicked_index]
        print(f"🎯 Navigation vers produit: {product_name}")

        # Retourner le nom du produit pour la recherche et fermer le dropdown
        return product_name, False

    raise dash.exceptions.PreventUpdate


def page_overview(master_df: pd.DataFrame = None, username: str = None, user_role: str = None):
    # Charger les données
    df = master_df if master_df is not None else get_df_cached()
    df = df.copy()

    # ============================================
    # 🔐 FILTRAGE PAR AGENT (si pas admin/manager)
    # ============================================
    agent_suppliers = None
    if username and user_role:
        agent_suppliers = get_agent_suppliers(username, user_role)

        if agent_suppliers and 'Supplier' in df.columns:
            # Normaliser pour comparaison
            df_suppliers_upper = df['Supplier'].str.upper().str.strip()
            agent_suppliers_upper = [s.upper().strip() for s in agent_suppliers]

            before_filter = len(df)
            df = df[df_suppliers_upper.isin(agent_suppliers_upper)].copy()
            after_filter = len(df)

            print(f"   🔐 Filtre agent {username}: {before_filter} → {after_filter} produits")
            print(f"      Fournisseurs: {agent_suppliers}")

    print(df.head())



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

        # Average Daily Sales non vide et numérique
        if "Average Daily Sales" not in df.columns:
            df["Average Daily Sales"] = 0.1

        df["Average Daily Sales"] = (
            pd.to_numeric(df["Average Daily Sales"], errors="coerce")
            .fillna(0.1)
            .clip(lower=0.1)
        )

        # ✅ NETTOYAGE COLONNES DUPLIQUÉES product_id
        product_id_cols = [c for c in df.columns if 'product_id' in c.lower()]
        if len(product_id_cols) > 1:
            # Trouver la colonne avec les vraies valeurs
            best_col = None
            best_count = 0
            for col in product_id_cols:
                try:
                    col_numeric = pd.to_numeric(df[col], errors='coerce').fillna(0)
                    valid_count = (col_numeric > 0).sum()
                    if valid_count > best_count:
                        best_count = valid_count
                        best_col = col
                except:
                    pass

            if best_col:
                df['product_id'] = pd.to_numeric(df[best_col], errors='coerce').fillna(0).astype(int)
                cols_to_drop = [c for c in product_id_cols if c != 'product_id']
                df.drop(columns=cols_to_drop, errors='ignore', inplace=True)

        if "product_id" not in df.columns:
            df["product_id"] = 0
        df["product_id"] = pd.to_numeric(df["product_id"], errors="coerce").fillna(0).astype(int)

        return df

    df = _ui_harden(df)

    # ✅ Harmoniser quelques alias pour l'UI
    if "supplier_categorization" in df.columns and "Product Category" not in df.columns:
        df["Product Category"] = df["supplier_categorization"]
    elif "product_category" in df.columns and "Product Category" not in df.columns:
        df["Product Category"] = df["product_category"]

    if "credit_cumulable" in df.columns and "Credit_cumulable" not in df.columns:
        df["Credit_cumulable"] = df["credit_cumulable"]

    if "supplier" in df.columns and "Supplier" not in df.columns:
        df["Supplier"] = df["supplier"]

    # ✅ 1. DÉFINIR colonnes prioritaires overview
    cols_priority_overview = [

        #"delete",
        "product_id",
        "product_name",
        "Supplier",
        "total_stock",
        "Average Daily Sales",
        "Max Daily Sales (Pikine)",
        "optimal stock",
        "Ajusted_total_need",
        "QAC edited",
        "QAC",
        "target_quantity",
        "Max Coverage Day",
        # ✅ NOUVEAU: Colonnes de dernière réception
        "last_reception_qty",
        "last_reception_date",
        "days_since_reception",
        "Product Category",
        "credit_days",
        "Credit_cumulable",
        "AJUSTER_BUFFER",
        "MAX_CREDIT_BUFFER",
        "ADJUSTED_LEADTIME",
        "MOQ MAAD",
        "delisting_status",
        "Daily OOS Rate (30d)",
        # "📝 Notes"

    ]

    # ✅ 2. Colonnes INTERDITES dans overview (techniques + promo)
    cols_banned_in_overview = [
        "_Product_Category_ABC_XYZ",
        "_Supplier_Categorization",
        "Average Daily Sales (7d)",
        "Average Daily Sales (30d)",
        "Daily OOS Rate (7d)",
        "Stockout Probability",
        "Credit Adequacy Score",
        # "Stock Status",
        "is_active",
        "demand_stability",
        "oos_risk",
        "target_quantity_calc",
        "Predicted Stockout",
        "replenishment_period",
        "Predicted Order Quantity",
        "purchase_need",
        "total_sold_30d",
        "avg_daily_sold",
        "std_sold",
        "max_daily_sold",
        "days_with_sales",
        "trend_factor",
        "projected_demand",
        "safety_stock",
        "target_supervised",
        # ✅ COLONNES PROMO INTERDITES
        'promo_status',
        'days_remaining',
        'uplift_pct',
        'roi_pct',
        'promo_recommendation',
        'promo_priority',
        'net_profit_per_day',
        'discount_pct',
        'sales_with_promo',
        'sales_without_promo',
        'additional_sales_per_day',
        'revenue_loss_per_day',
        'additional_profit_per_day'
    ]

    # ✅ 3. Filtrer colonnes disponibles
    available_priority = [c for c in cols_priority_overview if c in df.columns]
    extra_cols = [c for c in df.columns if c not in cols_priority_overview and c not in cols_banned_in_overview]

    # ✅ 4. Supprimer physiquement les colonnes bannies
    df_overview = df.drop(columns=[c for c in cols_banned_in_overview if c in df.columns], errors='ignore')

    available_cols = available_priority + extra_cols

    # ✅ 5. Garantir colonnes critiques
    for must in ["Supplier", "Average Daily Sales"]:
        if must not in available_cols:
            available_cols.insert(1, must)

    # ✅ NETTOYAGE: Supprimer colonnes product_id dupliquées
    product_id_variants = [c for c in df_overview.columns if 'product_id' in c.lower()]
    if len(product_id_variants) > 1:
        print(f"⚠️ page_overview: {len(product_id_variants)} colonnes product_id: {product_id_variants}")
        # Garder celle avec le plus de valeurs valides
        best_col = max(product_id_variants, key=lambda c: (pd.to_numeric(df_overview[c], errors='coerce').fillna(0) > 0).sum())
        df_overview['product_id'] = pd.to_numeric(df_overview[best_col], errors='coerce').fillna(0).astype(int)
        cols_to_drop = [c for c in product_id_variants if c != 'product_id']
        df_overview.drop(columns=cols_to_drop, errors='ignore', inplace=True)
        # Aussi nettoyer available_cols
        available_cols = [c for c in available_cols if c not in cols_to_drop]
        print(f"   ✅ Supprimé: {cols_to_drop}")

    # ✅ Vérifier que product_id est en PREMIÈRE position
    if "product_id" in df_overview.columns:
        if "product_id" not in available_cols:
            available_cols.insert(0, "product_id")
        elif available_cols[0] != "product_id":
            available_cols.remove("product_id")
            available_cols.insert(0, "product_id")
        valid_ids = (df_overview['product_id'] > 0).sum()
        total = len(df_overview)
        print(f"✅ product_id en première position - valeurs valides: {valid_ids}/{total}")
        if valid_ids == 0:
            print(f"   ⚠️⚠️⚠️ TOUS LES PRODUCT_ID SONT 0 ⚠️⚠️⚠️")
            print(f"   Échantillon df_overview:")
            print(df_overview[['product_id', 'product_name']].head(5).to_string())
    else:
        print(f"⚠️ product_id ABSENT de df_overview - colonnes: {df_overview.columns.tolist()[:10]}")

    # Vérifier que product_name est dans available_cols
    if "product_name" not in available_cols and "product_name" in df_overview.columns:
        available_cols.insert(1, "product_name")

    # Ligne ~145, juste avant make_kpis()

    # ✅ Recalculer Stock Status pour les KPIs (même si masquée du tableau)
    if 'total_stock' in df_overview.columns and 'optimal stock' in df_overview.columns:
        df_overview['Stock Status'] = df_overview.apply(
            lambda row: "Out of Stock" if row['total_stock'] <= 0
            else "Predicted Stockout Soon" if row['total_stock'] <= row.get('optimal stock', 0)
            else "Stock OK",
            axis=1
        )
    # Ajouter les colonnes d'actions "edit" et "delete"
    df_with_actions = df.copy()

    # Insérer les colonnes "edit" et "delete" dans le DataFrame avec un message d'action ou une valeur par défaut
    # df_with_actions.insert(0, "delete", "delete")  # Colonne Delete
    #df_with_actions.insert(0, "edit", "edit")  # Colonne Edit
    # Appliquer la fonction pour ajouter les colonnes "edit" et "delete"
    #df_with_actions = add_action_cols(df)

    # Mettre à jour la liste des colonnes disponibles
    available_cols_with_actions = ["edit", "delete"] + available_cols

    # KPIs
    kpi_cards, risk_bell = make_kpis(df_overview)

    # En-tête
    header_row = dbc.Row([
        dbc.Col(html.H2("Overview"), md=8),
        dbc.Col(html.Div(risk_bell, style={"textAlign": "right"}), md=4),
    ])

    # ✅ DROPDOWN COMPACT ET CLAIR
    rotation_dropdown = html.Div([
        html.Span("📊 Période ADS :", style={
            "fontWeight": "600",
            "marginRight": "10px",
            "color": "#475569",
            "fontSize": "13px"
        }),
        dcc.Dropdown(
            id="rotation-period",
            options=[
                {"label": "3 jours", "value": "3d"},
                {"label": "7 jours", "value": "7d"},
                {"label": "30 jours", "value": "30d"},
            ],
            value="7d",
            clearable=False,
            searchable=False,
            style={"width": "120px", "display": "inline-block", "verticalAlign": "middle"},
            className="rotation-dropdown-compact"
        ),
    ], style={
        "display": "inline-flex",
        "alignItems": "center",
        "padding": "8px 14px",
        "background": "#f1f5f9",
        "borderRadius": "8px",
        "border": "1px solid #e2e8f0",
        "marginBottom": "12px"
    })

    # Ajouter la nouvelle colonne 'QAC edited' dans available_cols
    available_cols = available_cols #+ ['QAC edited']  # Ajoute 'QAC edited' à la liste des colonnes

    # Générer dynamiquement les colonnes et rendre 'QAC edited' editable
    columns = [
        {"name": c, "id": c, "deletable": False, "hideable": False, "editable": True if c == 'QAC edited' else False}
        for c in available_cols
    ]

    # DataTable avec couleurs améliorées par besoin d'achat
    table = dash_table.DataTable(
        id="main-table",
        columns=columns,
        data=df_overview[available_cols].to_dict("records"),
        page_size=20,  # ✅ Augmenté légèrement
        filter_action="native",
        sort_action="native",
        sort_mode="single",  # ✅ Single pour performance
        column_selectable="single",
        editable=True,
        active_cell=None,
        dropdown_conditional=[],
        row_selectable="multi",
        selected_rows=[],

        # ✅ OPTIMISATION PERFORMANCE
        virtualization=False,  # Désactivé pour scroll fluide
        page_action='native',

        # ========================================
        # TABLE STYLES - SCROLL HORIZONTAL UNIQUEMENT
        # ========================================
        style_table={
            "overflowX": "auto",
            "overflowY": "visible",
            "maxWidth": "100%",
            "border": "1px solid #e2e8f0",
            "borderRadius": "8px"
        },

        style_header={
            "backgroundColor": "#1e3a5f",
            "border": "1px solid #e2e8f0",
            "fontWeight": "700",
            "textAlign": "center",
            "color": "#ffffff",
            "fontSize": "12px",
            "textTransform": "uppercase",
            "letterSpacing": "0.5px",
            "padding": "12px 8px"
        },

        style_cell={
            "backgroundColor": "#ffffff",
            "color": "#1e293b",
            "border": "1px solid #e2e8f0",
            "fontSize": 12,
            "textAlign": "center",
            "padding": "8px",
            "minWidth": "80px",
            "maxWidth": "300px",
            "overflow": "hidden",
            "textOverflow": "ellipsis"
        },

        # ========================================
        # STYLES PAR COLONNE
        # ========================================
        style_cell_conditional=[
            # QAC edited - éditable
            {
                "if": {"column_id": "QAC edited"},
                "width": "120px",
                "minWidth": "120px",
                "maxWidth": "120px",
                "textAlign": "center",
                "cursor": "text",
                "fontWeight": "700",
                "fontSize": "14px",
                "backgroundColor": "#e0f2fe",
                "border": "2px dashed #0ea5e9",
                "color": "#0369a1"
            },

            # QAC - valeur calculée
            {
                "if": {"column_id": "QAC"},
                "backgroundColor": "#f0f9ff",
                "cursor": "default",
                "fontWeight": "700",
                "fontSize": "13px",
                "color": "#0369a1"
            },

            # product_name - aligné à gauche
            {
                "if": {"column_id": "product_name"},
                "textAlign": "left",
                "fontWeight": "600",
                "minWidth": "250px",
                "maxWidth": "350px",
                "paddingLeft": "12px"
            },

            # Supplier
            {
                "if": {"column_id": "Supplier"},
                "textAlign": "left",
                "color": "#64748b",
                "fontSize": "11px",
                "fontStyle": "italic"
            },

            # Colonnes numériques importantes
            {
                "if": {"column_id": ["total_stock", "target_quantity"]},
                "fontWeight": "700",
                "fontSize": "13px"
            },

            # Average Daily Sales
            {
                "if": {"column_id": "Average Daily Sales"},
                "fontWeight": "700",
                "color": "#7c3aed",
                "fontSize": "13px"
            },

            # Max Daily Sales
            {
                "if": {"column_id": "Max Daily Sales (Pikine)"},
                "fontWeight": "600",
                "color": "#d97706"
            },

            # Product Category - badge style
            {
                "if": {"column_id": "Product Category"},
                "fontWeight": "700",
                "fontSize": "11px",
                "textTransform": "uppercase",
                "letterSpacing": "0.5px"
            },

            # ✅ NOUVEAU: Colonnes de dernière réception
            {
                "if": {"column_id": "last_reception_qty"},
                "fontWeight": "700",
                "fontSize": "13px",
                "color": "#34d399",  # Vert
                "backgroundColor": "rgba(16, 185, 129, 0.1)",
                "textAlign": "center"
            },
            {
                "if": {"column_id": "last_reception_date"},
                "fontSize": "11px",
                "color": "#94a3b8",
                "textAlign": "center",
                "fontStyle": "italic"
            },
            {
                "if": {"column_id": "days_since_reception"},
                "fontWeight": "600",
                "fontSize": "12px",
                "textAlign": "center"
            }
        ],

        # ========================================
        # 🎨 STYLES CONDITIONNELS - COULEURS PAR BESOIN
        # ========================================
        style_data_conditional=[
            # ========================================
            # 🔴 ORDER NOW - LIGNE ENTIÈRE ROUGE
            # ========================================
            {
                "if": {"filter_query": "{Ajusted_total_need} = 'ORDER NOW'"},
                "backgroundColor": "rgba(239, 68, 68, 0.12)",
                "color": "#991b1b",
                "fontWeight": "600"
            },

            # Colonne Ajusted_total_need - ORDER NOW (badge fort)
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'ORDER NOW'",
                    "column_id": "Ajusted_total_need"
                },
                "backgroundColor": "#dc2626",
                "color": "#ffffff",
                "fontWeight": "800",
                "fontSize": "12px",
                "border": "none",
                "borderRadius": "6px",
                "textTransform": "uppercase"
            },

            # product_name - ORDER NOW (bordure rouge)
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'ORDER NOW'",
                    "column_id": "product_name"
                },
                "backgroundColor": "rgba(239, 68, 68, 0.15)",
                "color": "#7f1d1d",
                "fontWeight": "700",
                "borderLeft": "4px solid #ef4444"
            },

            # Stock - ORDER NOW (highlight)
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'ORDER NOW'",
                    "column_id": "total_stock"
                },
                "backgroundColor": "rgba(239, 68, 68, 0.18)",
                "color": "#991b1b",
                "fontWeight": "700",
                "fontSize": "14px"
            },

            # QAC - ORDER NOW (highlight)
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'ORDER NOW'",
                    "column_id": "QAC"
                },
                "backgroundColor": "#fecaca",
                "color": "#7f1d1d",
                "fontWeight": "800",
                "fontSize": "15px",
                "border": "2px solid #ef4444"
            },

            # ========================================
            # 🟠 ORDER NOT URGENT - LIGNE ORANGE
            # ========================================
            {
                "if": {"filter_query": "{Ajusted_total_need} = 'ORDER NOT URGENT'"},
                "backgroundColor": "rgba(245, 158, 11, 0.1)",
                "color": "#92400e",
                "fontWeight": "500"
            },

            # Colonne Ajusted_total_need - ORDER NOT URGENT
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'ORDER NOT URGENT'",
                    "column_id": "Ajusted_total_need"
                },
                "backgroundColor": "#f59e0b",
                "color": "#ffffff",
                "fontWeight": "700",
                "fontSize": "12px",
                "border": "none",
                "borderRadius": "6px",
                "textTransform": "uppercase"
            },

            # product_name - ORDER NOT URGENT
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'ORDER NOT URGENT'",
                    "column_id": "product_name"
                },
                "backgroundColor": "rgba(245, 158, 11, 0.12)",
                "color": "#78350f",
                "fontWeight": "600",
                "borderLeft": "4px solid #f59e0b"
            },

            # QAC - ORDER NOT URGENT
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'ORDER NOT URGENT'",
                    "column_id": "QAC"
                },
                "backgroundColor": "#fef3c7",
                "color": "#92400e",
                "fontWeight": "700",
                "fontSize": "14px"
            },

            # ========================================
            # 🟢 NO NEED - LIGNE VERTE
            # ========================================
            {
                "if": {"filter_query": "{Ajusted_total_need} = 'NO NEED'"},
                "backgroundColor": "rgba(16, 185, 129, 0.08)",
                "color": "#065f46"
            },

            # Colonne Ajusted_total_need - NO NEED
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'NO NEED'",
                    "column_id": "Ajusted_total_need"
                },
                "backgroundColor": "#10b981",
                "color": "#ffffff",
                "fontWeight": "700",
                "fontSize": "12px",
                "border": "none",
                "borderRadius": "6px",
                "textTransform": "uppercase"
            },

            # product_name - NO NEED (bordure verte subtile)
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'NO NEED'",
                    "column_id": "product_name"
                },
                "borderLeft": "4px solid #10b981",
                "color": "#064e3b"
            },

            # ========================================
            # 📊 MAX COVERAGE DAY - DÉGRADÉ DE COULEURS
            # ========================================

            # Moins de 7 jours - CRITIQUE
            {
                "if": {
                    "filter_query": "{Max Coverage Day} < 7",
                    "column_id": "Max Coverage Day"
                },
                "backgroundColor": "#fecaca",
                "color": "#7f1d1d",
                "fontWeight": "800",
                "fontSize": "13px",
                "border": "1px solid #ef4444"
            },

            # 7-14 jours - ATTENTION
            {
                "if": {
                    "filter_query": "{Max Coverage Day} >= 7 && {Max Coverage Day} < 14",
                    "column_id": "Max Coverage Day"
                },
                "backgroundColor": "#fef3c7",
                "color": "#92400e",
                "fontWeight": "700",
                "fontSize": "13px"
            },

            # 14-30 jours - BON
            {
                "if": {
                    "filter_query": "{Max Coverage Day} >= 14 && {Max Coverage Day} < 30",
                    "column_id": "Max Coverage Day"
                },
                "backgroundColor": "#d1fae5",
                "color": "#065f46",
                "fontWeight": "600"
            },

            # Plus de 30 jours - EXCELLENT
            {
                "if": {
                    "filter_query": "{Max Coverage Day} >= 30",
                    "column_id": "Max Coverage Day"
                },
                "backgroundColor": "#a7f3d0",
                "color": "#064e3b",
                "fontWeight": "600"
            },

            # ========================================
            # 🏷️ PRODUCT CATEGORY - COULEURS PAR TYPE
            # ========================================
            {
                "if": {
                    "filter_query": "{Product Category} contains 'a'",
                    "column_id": "Product Category"
                },
                "backgroundColor": "#fee2e2",
                "color": "#991b1b"
            },
            {
                "if": {
                    "filter_query": "{Product Category} contains 'b'",
                    "column_id": "Product Category"
                },
                "backgroundColor": "#fef3c7",
                "color": "#92400e"
            },
            {
                "if": {
                    "filter_query": "{Product Category} contains 'c'",
                    "column_id": "Product Category"
                },
                "backgroundColor": "#d1fae5",
                "color": "#065f46"
            },

            # ========================================
            # 📦 JOURS DEPUIS RÉCEPTION - COULEURS PAR ANCIENNETÉ
            # ========================================
            # Réception récente (< 7 jours) - Vert
            {
                "if": {
                    "filter_query": "{days_since_reception} < 7",
                    "column_id": "days_since_reception"
                },
                "backgroundColor": "#d1fae5",
                "color": "#065f46",
                "fontWeight": "700"
            },
            # Réception moyenne (7-14 jours) - Jaune
            {
                "if": {
                    "filter_query": "{days_since_reception} >= 7 && {days_since_reception} < 14",
                    "column_id": "days_since_reception"
                },
                "backgroundColor": "#fef3c7",
                "color": "#92400e",
                "fontWeight": "600"
            },
            # Réception ancienne (14-30 jours) - Orange
            {
                "if": {
                    "filter_query": "{days_since_reception} >= 14 && {days_since_reception} < 30",
                    "column_id": "days_since_reception"
                },
                "backgroundColor": "#fed7aa",
                "color": "#9a3412",
                "fontWeight": "600"
            },
            # Réception très ancienne (> 30 jours) - Rouge
            {
                "if": {
                    "filter_query": "{days_since_reception} >= 30",
                    "column_id": "days_since_reception"
                },
                "backgroundColor": "#fecaca",
                "color": "#991b1b",
                "fontWeight": "700"
            },

            # ========================================
            # 🎯 ÉTATS INTERACTIFS
            # ========================================

            # Lignes sélectionnées - TRÈS VISIBLE
            {
                "if": {"state": "selected"},
                "backgroundColor": "#0ea5e9 !important",
                "border": "2px solid #0369a1",
                "fontWeight": "700",
                "color": "#ffffff !important"
            },

            # Cellule active (en cours d'édition)
            {
                "if": {"state": "active"},
                "backgroundColor": "#e0f2fe",
                "border": "2px solid #0ea5e9",
                "outline": "none",
                "color": "#0c4a6e",
                "fontWeight": "700"
            },

            # QAC edited - toujours visible
            {
                "if": {"column_id": "QAC edited"},
                "backgroundColor": "#f0f9ff"
            }
        ],

        style_data={
            "whiteSpace": "normal",
            "height": "auto"
        },

        export_format="csv",
        export_headers="display",
        persistence=False,
        persisted_props=["filter_query", "sort_by", "page_current",
                         "selected_rows", "selected_columns", "hidden_columns"],
    )

    # Barre d'actions compacte
    action_buttons = html.Div([
        html.Div([
            dbc.Button("Actualiser", id="btn-refresh", color="light", size="sm",
                       style={"fontSize": "12px", "padding": "4px 10px", "marginRight": "6px"}),
            dbc.Button("Recharger", id="btn-reload-data", color="warning", size="sm", outline=True,
                       style={"fontSize": "12px", "padding": "4px 10px", "marginRight": "6px"}),
            dbc.Button("Tout sélect.", id="btn-select-all", color="secondary", size="sm", outline=True,
                       style={"fontSize": "12px", "padding": "4px 10px", "marginRight": "6px"}),
            dbc.Button("+ Ajouter", id="btn-add-row", color="info", size="sm",
                       style={"fontSize": "12px", "padding": "4px 10px"}),
        ], style={"display": "flex", "alignItems": "center"}),

        html.Span(id="selection-count-inline", style={
            "marginLeft": "12px", "marginRight": "auto", "fontSize": "12px", "color": "#64748b"
        }),

        html.Div([
            dbc.Button("Agent IA", id="btn-run-agent-ia", color="success", size="sm",
                       style={"fontSize": "12px", "padding": "4px 12px", "marginRight": "6px"}),
            dbc.Button("Bon de Commande", id="btn-po-pdf", color="primary", size="sm", disabled=True,
                       style={"fontSize": "12px", "padding": "4px 12px"}),
        ], style={"display": "flex", "alignItems": "center"})

    ], style={
        "display": "flex", "alignItems": "center", "justifyContent": "space-between",
        "padding": "8px 12px", "background": "#fff", "borderRadius": "8px",
        "border": "1px solid #e2e8f0", "marginBottom": "12px"
    })


    # Dropdown filter-status
    #dcc.Dropdown(
    #   id="filter-status",
    #  options=[
    #     {"label": "Tous", "value": "all"},
    #    {"label": "Stock OK", "value": "ok"},
    #   {"label": "Rupture", "value": "oos"}
    #],
    #value="all",
    #className="filter-dropdown"
    #)

    # Modal édition
    #edit_modal = dbc.Modal(
    #   [
    #      dbc.ModalHeader(dbc.ModalTitle("Éditer produit")),
    #     dbc.ModalBody([
    #        html.Div("Formulaire d'édition à implémenter ici…"),
    #       dcc.Input(id="edit-input", type="text", placeholder="Modifier la valeur")
    #  ]),
    # dbc.ModalFooter(
    #    dbc.Button("Fermer", id="close-edit", className="ms-auto", n_clicks=0)
    #),
    #],
    #id="edit-modal",
    #is_open=False,
    #)

    # ✅ NOUVEAU : Bouton flottant + Modal amélioré
    notes_fab_button = html.Button(
        "💬 Notes",
        id="notes-fab",
        className="btn btn-primary",
        style={
            "position": "fixed",
            "right": "24px",
            "bottom": "100px",  # Au-dessus du chat
            "zIndex": "9999",
            "borderRadius": "50px",
            "padding": "12px 20px",
            "boxShadow": "0 4px 12px rgba(34, 211, 238, 0.4)",
            "background": "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
            "border": "none",
            "color": "white",
            "fontWeight": "700",
            "fontSize": "14px",
            "cursor": "pointer"
        }
    )

    # ✅ Modal avec sélecteur de produit - THÈME CLAIR LISIBLE
    notes_modal = dbc.Modal([
        dbc.ModalHeader([
            html.Div([
                html.Span("💬", style={"fontSize": "20px", "marginRight": "8px"}),
                html.Span("Notes Produit", style={"fontWeight": "700", "color": "#1e293b"})
            ], style={"display": "flex", "alignItems": "center"})
        ], style={
            "background": "linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%)",
            "borderBottom": "2px solid #0ea5e9",
            "padding": "16px 20px"
        }),

        dbc.ModalBody([
            # ========================================
            # 🎯 SÉLECTEUR DE PRODUIT
            # ========================================
            html.Div([
                html.Label("Produit concerné", style={
                    "color": "#1e293b",
                    "fontWeight": "700",
                    "marginBottom": "8px",
                    "fontSize": "14px"
                }),
                dcc.Dropdown(
                    id="note-product-selector",
                    options=[],  # Sera rempli dynamiquement
                    placeholder="Sélectionner un produit...",
                    style={
                        "marginBottom": "16px"
                    },
                    className="light-dropdown"
                )
            ], style={
                "padding": "16px",
                "background": "#ffffff",
                "borderRadius": "12px",
                "border": "1px solid #e2e8f0",
                "marginBottom": "16px",
                "boxShadow": "0 1px 3px rgba(0,0,0,0.05)"
            }),

            html.Hr(style={"borderColor": "#cbd5e1", "margin": "16px 0"}),

            # ========================================
            # 📋 LISTE DES NOTES EXISTANTES
            # ========================================
            html.Div([
                html.Div("📋 Historique des notes", style={
                    "fontWeight": "700",
                    "color": "#1e293b",
                    "marginBottom": "12px",
                    "fontSize": "14px"
                }),
                html.Div(
                    id="notes-display-list",
                    style={
                        "maxHeight": "300px",
                        "overflowY": "auto",
                        "padding": "12px",
                        "background": "#f8fafc",
                        "borderRadius": "10px",
                        "border": "1px solid #e2e8f0"
                    }
                )
            ], style={
                "marginBottom": "16px"
            }),

            html.Hr(style={"borderColor": "#cbd5e1", "margin": "16px 0"}),

            # ========================================
            # ✍️ FORMULAIRE NOUVELLE NOTE
            # ========================================
            html.Div([
                html.Label("✍️ Nouvelle note", style={
                    "color": "#1e293b",
                    "fontWeight": "700",
                    "marginBottom": "8px",
                    "fontSize": "14px"
                }),

                dbc.Textarea(
                    id="note-text-input",
                    placeholder="Écrivez votre message... Utilisez @tony, @samuel, etc. pour notifier",
                    rows=4,
                    style={
                        "background": "#ffffff",
                        "color": "#1e293b",
                        "border": "2px solid #cbd5e1",
                        "borderRadius": "10px",
                        "marginBottom": "12px",
                        "fontSize": "14px",
                        "padding": "12px",
                        "fontFamily": "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif"
                    }
                ),

                dbc.Input(
                    id="note-author-input",
                    placeholder="Votre nom",
                    type="text",
                    style={
                        "background": "#ffffff",
                        "color": "#1e293b",
                        "border": "2px solid #cbd5e1",
                        "borderRadius": "10px",
                        "marginBottom": "12px",
                        "fontSize": "14px",
                        "padding": "12px",
                        "fontWeight": "600"
                    }
                ),

                # Mentions disponibles (badges)
                html.Div([
                    html.Strong("💡 Mentions disponibles :", style={
                        "color": "#475569",
                        "fontSize": "12px",
                        "display": "block",
                        "marginBottom": "8px"
                    }),
                    html.Div([
                        html.Span("@tony", className="mention-badge"),
                        html.Span("@samuel", className="mention-badge"),
                        html.Span("@maimouna", className="mention-badge"),
                        html.Span("@seydouna", className="mention-badge"),
                        html.Span("@Arame", className="mention-badge"),
                        html.Span("@Coumba", className="mention-badge"),
                        html.Span("@Fallou", className="mention-badge"),
                        html.Span("@Ravane", className="mention-badge"),
                        html.Span("@Insa", className="mention-badge"),
                    ], style={"display": "flex", "flexWrap": "wrap", "gap": "6px"})
                ], style={
                    "padding": "12px",
                    "background": "#f1f5f9",
                    "borderRadius": "8px",
                    "marginBottom": "10px"
                }),

                # Feedback
                html.Div(id="note-feedback-new", style={
                    "color": "#059669",
                    "fontSize": "13px",
                    "marginTop": "8px",
                    "fontWeight": "600"
                })
            ], style={
                "padding": "16px",
                "background": "#ffffff",
                "borderRadius": "12px",
                "border": "1px solid #e2e8f0",
                "boxShadow": "0 1px 3px rgba(0,0,0,0.05)"
            })
        ], style={
            "background": "linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%)",
            "padding": "20px"
        }),

        dbc.ModalFooter([
            dbc.Button(
                "Fermer",
                id="note-modal-close",
                outline=True,
                color="secondary",
                size="sm",
                style={
                    "marginRight": "8px",
                    "fontWeight": "600",
                    "borderWidth": "2px",
                    "color": "#475569",
                    "borderColor": "#cbd5e1"
                }
            ),
            dbc.Button(
                "✉️ Envoyer",
                id="note-modal-send",
                color="primary",
                size="sm",
                style={
                    "fontWeight": "700",
                    "background": "linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%)",
                    "border": "none",
                    "padding": "8px 20px"
                }
            )
        ], style={
            "background": "#f8fafc",
            "borderTop": "2px solid #cbd5e1",
            "padding": "16px 20px"
        })
    ],
        id="notes-modal-new",
        size="lg",
        is_open=False,
        backdrop="static",
        scrollable=True
    )

    # ========== RETURN DU LAYOUT PAGE OVERVIEW ==========
    return html.Div(className="content", children=[
        # En-tête avec titre + KPI risque de rupture
        dbc.Row([
            dbc.Col(html.H2(" Overview", style={"color": "#0ea5e9", "fontWeight": "800"}), md=8),
            dbc.Col(html.Div(risk_bell, style={"textAlign": "right"}), md=4),
        ], className="mb-3"),

        # KPIs (SKUs, Ruptures, Fournisseurs)
        html.Div(kpi_cards),
        html.Br(),

        # Dropdown période rotation
        rotation_dropdown,
        html.Br(),

        # Boutons d'action
        action_buttons,

        # Table principale
        html.Div(className="soft-card", children=[table]),

        # Modal notes
        notes_fab_button,
        notes_modal
    ])


@app.callback(
    [Output("filtered-data", "data", allow_duplicate=True),
     Output("main-table", "data", allow_duplicate=True),
     Output("main-table", "selected_rows", allow_duplicate=True)],
    [Input("search-input", "value"),
     Input("filter-supplier", "value"),
     Input("filter-category", "value"),
     Input("filter-need", "value"),
     Input("toggle-options", "value"),
     Input("master-data", "data")],
    prevent_initial_call=True  # ✅ CORRIGÉ: Évite erreur main-table au démarrage
)
def apply_filters(search, sup, cat, need, options, master_json):
    """
    ⚡ CALLBACK UNIFIÉ OPTIMISÉ - Initialisation + Filtrage
    - Utilise orjson pour parsing rapide
    - Minimise les copies DataFrame
    - Filtrage vectorisé
    """
    start_time = time.time()

    # ✅ Charger données avec orjson (plus rapide)
    if not master_json:
        base = get_df_cached()
    else:
        try:
            # orjson est 5-10x plus rapide que json.loads
            base = pd.DataFrame(orjson.loads(master_json))
        except:
            base = pd.DataFrame(json.loads(master_json))

    if base.empty:
        return no_update, no_update, []

    # ✅ PRÉSERVER product_id IMMÉDIATEMENT après chargement
    if 'product_id' in base.columns:
        base['product_id'] = pd.to_numeric(base['product_id'], errors='coerce').fillna(0).astype(int)

    # ✅ Validation (inline pour éviter appel fonction)
    for col in ['product_name', 'Supplier', 'Product Category']:
        if col not in base.columns:
            base[col] = ""

    # ✅ Supprimer colonnes bannies (in-place)
    promo_cols = [
        'promo_status', 'uplift_pct', 'roi_pct',
        'Average Daily Sales (7d)', 'Average Daily Sales (30d)',
        'Average Daily Sales (3d)', 'Daily OOS Rate (7d)',
        'Daily OOS Rate (30d)', 'Stockout Probability',
        'Credit Adequacy Score'
    ]
    cols_to_drop = [c for c in promo_cols if c in base.columns]
    if cols_to_drop:
        base.drop(columns=cols_to_drop, inplace=True, errors='ignore')

    # ✅ APPLIQUER FILTRES VECTORISÉS (sans copie intermédiaire)
    mask = pd.Series(True, index=base.index)

    if search and search.strip():
        search_lower = search.strip().lower()
        search_mask = pd.Series(False, index=base.index)
        for col in ['product_name', 'Supplier', 'Product Category']:
            if col in base.columns:
                search_mask |= base[col].astype(str).str.lower().str.contains(search_lower, na=False, regex=False)
        mask &= search_mask

    if sup and len(sup) > 0 and 'Supplier' in base.columns:
        mask &= base['Supplier'].isin(sup)

    if cat and len(cat) > 0 and 'Product Category' in base.columns:
        mask &= base['Product Category'].isin(cat)

    if need and len(need) > 0 and 'Ajusted_total_need' in base.columns:
        mask &= base['Ajusted_total_need'].isin(need)

    if options and len(options) > 0 and 'Stock Status' in base.columns:
        status_mask = pd.Series(False, index=base.index)
        for opt in options:
            if opt == 'show_out_of_stock':
                status_mask |= (base['Stock Status'] == 'Out of Stock')
            elif opt == 'show_predicted_stockout':
                status_mask |= (base['Stock Status'] == 'Predicted Stockout Soon')
            elif opt == 'show_order_soon':
                status_mask |= (base['Stock Status'] == 'Order Soon')
        mask &= status_mask

    # Appliquer le masque une seule fois
    fdf = base.loc[mask].reset_index(drop=True)

    # ✅ FORMATAGE RAPIDE (vectorisé)
    int_cols = ['total_stock', 'QAC', 'target_quantity', 'product_id']
    for col in int_cols:
        if col in fdf.columns:
            fdf[col] = pd.to_numeric(fdf[col], errors='coerce').fillna(0).astype(int)

    if 'Max Coverage Day' in fdf.columns:
        fdf['Max Coverage Day'] = pd.to_numeric(fdf['Max Coverage Day'], errors='coerce').fillna(0).round(1)

    if 'Average Daily Sales' in fdf.columns:
        fdf['Average Daily Sales'] = pd.to_numeric(fdf['Average Daily Sales'], errors='coerce').fillna(0.1).round(2)

    # ✅ Actions
    fdf = add_action_cols(fdf)

    elapsed_total = time.time() - start_time
    print(f"⚡ FILTRE: {len(fdf)} produits en {elapsed_total:.3f}s")

    # ✅ Sérialisation rapide avec orjson
    try:
        json_data = orjson.dumps(fdf.to_dict("records"), option=ORJSON_OPTS).decode('utf-8')
    except:
        json_data = fdf.to_json(orient="records")

    return (
        json_data,
        fdf.to_dict("records"),
        []
    )

@app.callback(
    Output("master-data", "data", allow_duplicate=True),
    Input("rotation-period", "value"),
    prevent_initial_call=True
)
def update_rotation_simple(period_value):
    """
    ⚡ CALLBACK ROTATION INSTANTANÉ
    Met à jour master-data uniquement
    → apply_filters se déclenchera automatiquement
    """
    start_time = time.time()

    print(f"\n{'=' * 70}")
    print(f"⚡ ROTATION : {period_value or '7d'}")
    print(f"{'=' * 70}")

    if not period_value:
        period_value = "7d"

    try:
        # ⚡ RÉCUPÉRATION INSTANTANÉE (< 0.1s)
        df = get_preloaded_data(period_value)

        elapsed_fetch = time.time() - start_time
        print(f"⚡ {len(df)} produits en {elapsed_fetch:.3f}s")

        # ✅ Validation rapide
        if "Supplier" not in df.columns:
            df["Supplier"] = "unknown"
        else:
            df["Supplier"] = df["Supplier"].fillna("unknown")

        if "Average Daily Sales" not in df.columns:
            df["Average Daily Sales"] = 0.1
        else:
            df["Average Daily Sales"] = df["Average Daily Sales"].fillna(0.1).clip(lower=0.1)

        if "product_id" in df.columns:
            df["product_id"] = df["product_id"].fillna(0).astype(int)

        # ✅ Sérialisation
        master_json = df.to_json(orient="records", date_format='iso')

        elapsed_total = time.time() - start_time

        print(f"📊 ADS {period_value} :")
        print(
            f"   Min/Max/Moy : {df['Average Daily Sales'].min():.2f} / {df['Average Daily Sales'].max():.2f} / {df['Average Daily Sales'].mean():.2f}")
        print(f"⚡ CALLBACK TERMINÉ en {elapsed_total:.3f}s")
        print(f"{'=' * 70}\n")

        return master_json

    except Exception as e:
        print(f"❌ ERREUR : {e}\n")
        import traceback
        traceback.print_exc()
        return no_update

'''
# Callback 2 : Filtrage (avec allow_duplicate)
@app.callback(
    [Output("filtered-data", "data", allow_duplicate=True),
     Output("main-table", "data", allow_duplicate=True),
     #Output("risk-banner", "children", allow_duplicate=True),
     Output("main-table", "selected_rows", allow_duplicate=True)],
    [Input("search-input", "value"),
     Input("filter-supplier", "value"),
     Input("filter-category", "value"),
     Input("filter-need", "value"),
     Input("toggle-options", "value"),
     Input("master-data", "data")],  # ✅ 6ème Input
    # ❌ SUPPRIMER CE State (doublon avec Input ci-dessus)
    # State("master-data", "data"),
    prevent_initial_call=True
)
def apply_filters(search, sup, cat, need, options, master_json):
    """
    Applique tous les filtres
    ✅ CORRIGÉ : 6 Inputs = 6 arguments
    """

    # ✅ Utiliser master_json directement (c'est déjà un Input)
    if not master_json:
        print("⚠️ master_json vide, rechargement...")
        base = get_df_cached()
    else:
        print("📥 Utilisation master_json...")
        base = pd.DataFrame(json.loads(master_json))

    # Validation colonnes critiques
    base = validate_core_columns(base)

    # Supprimer colonnes promo
    promo_cols_to_remove = [
        'promo_status', 'days_remaining', 'uplift_pct', 'roi_pct',
        'promo_recommendation', 'promo_priority', 'net_profit_per_day',
        'discount_pct', 'sales_with_promo', 'sales_without_promo',
        'additional_sales_per_day', 'revenue_loss_per_day',
        'additional_profit_per_day',
        'Average Daily Sales (7d)', 'Average Daily Sales (30d)',
        'Average Daily Sales (3d)', 'Daily OOS Rate (7d)',
        'Daily OOS Rate (30d)', 'Stockout Probability'
    ]
    base = base.drop(columns=[c for c in promo_cols_to_remove if c in base.columns], errors='ignore')

    print(f"\n{'=' * 60}")
    print(f"🔍 FILTRAGE EN COURS")
    print(f"{'=' * 60}")
    print(f"📊 Base de départ: {len(base)} produits")

    # Copie de travail
    fdf = base.copy()

    # ========== FILTRE 1: RECHERCHE ==========
    if search and search.strip():
        search_lower = search.strip().lower()
        print(f"🔎 Recherche: '{search_lower}'")

        mask = pd.Series([False] * len(fdf), index=fdf.index)
        search_columns = ['product_name', 'Supplier', 'Product Category']

        for col in search_columns:
            if col in fdf.columns:
                mask |= fdf[col].astype(str).str.lower().str.contains(search_lower, na=False, regex=False)

        fdf = fdf[mask].reset_index(drop=True)
        print(f"   → {len(fdf)} produits")

    # ========== FILTRE 2: FOURNISSEUR ==========
    if sup and len(sup) > 0 and 'Supplier' in fdf.columns:
        print(f"🏭 Fournisseurs: {len(sup)} sélectionnés")
        fdf = fdf[fdf['Supplier'].isin(sup)]
        print(f"   → {len(fdf)} produits")

    # ========== FILTRE 3: CATÉGORIE ==========
    if cat and len(cat) > 0 and 'Product Category' in fdf.columns:
        print(f"🏷️ Catégories: {len(cat)} sélectionnées")
        fdf = fdf[fdf['Product Category'].isin(cat)]
        print(f"   → {len(fdf)} produits")

    # ========== FILTRE 4: BESOIN ==========
    if need and len(need) > 0 and 'Ajusted_total_need' in fdf.columns:
        print(f"📦 Besoins: {need}")
        fdf = fdf[fdf['Ajusted_total_need'].isin(need)]
        print(f"   → {len(fdf)} produits")

    # ========== FILTRE 5: OPTIONS ==========
    if options and len(options) > 0:
        print(f"⚙️ Options: {options}")

        if 'show_out_of_stock' in options and 'Stock Status' in fdf.columns:
            fdf = fdf[fdf['Stock Status'] == 'Out of Stock']
            print(f"   → OOS: {len(fdf)} produits")

        if 'show_predicted_stockout' in options and 'Stock Status' in fdf.columns:
            fdf = fdf[fdf['Stock Status'] == 'Predicted Stockout Soon']
            print(f"   → Predicted: {len(fdf)} produits")

        if 'show_order_soon' in options and 'Stock Status' in fdf.columns:
            fdf = fdf[fdf['Stock Status'] == 'Order Soon']
            print(f"   → Order Soon: {len(fdf)} produits")

    # ========== FORMATAGE VALEURS ==========
    for col in ['total_stock', 'QAC', 'target_quantity']:
        if col in fdf.columns:
            fdf[col] = pd.to_numeric(fdf[col], errors='coerce').fillna(0).round(0).astype(int)

    if 'Max Coverage Day' in fdf.columns:
        fdf['Max Coverage Day'] = pd.to_numeric(fdf['Max Coverage Day'], errors='coerce').fillna(0).round(1)

    if 'Average Daily Sales' in fdf.columns:
        fdf['Average Daily Sales'] = pd.to_numeric(fdf['Average Daily Sales'], errors='coerce').fillna(0.1).round(2)

    # ========== NETTOYAGE ==========
    # Supprimer colonnes actions si présentes
    for action_col in ["QAC edited", "Delete"]:
        if action_col in fdf.columns:
            fdf = fdf.drop(columns=[action_col])

    # Ajouter colonnes actions
    fdf_actions = add_action_cols(fdf)

    # Sécurité finale
    final_cols = [c for c in fdf_actions.columns if c not in promo_cols_to_remove]
    fdf_actions = fdf_actions[final_cols]

    # ========== BANNER ==========
    risk_count = 0
    if 'Stock Status' in fdf_actions.columns:
        risk_count = int(fdf_actions['Stock Status'].isin(['Out of Stock', 'Predicted Stockout Soon']).sum())

    banner = html.Div([
        html.Span("🔔 ", style={"fontSize": "16px"}),
        html.B(f"{risk_count} produits à risque", style={"color": "#f59e0b" if risk_count > 0 else "#10b981"})
    ])

    # ========== RÉSULTAT ==========
    print(f"{'=' * 60}")
    print(f"✅ FILTRAGE TERMINÉ: {len(fdf_actions)} produits")
    print(f"{'=' * 60}\n")

    return (
        fdf_actions.to_json(orient="records"),  # filtered-data
        fdf_actions.to_dict("records"),  # main-table data
       # banner,  # risk-banner
        []  # selected_rows reset
    )


# ==========================================
# ✅ CALLBACK ROTATION SIMPLE (ajouter celui-ci)
# ==========================================
@app.callback(
    Output("master-data", "data", allow_duplicate=True),
    Input("rotation-period", "value"),
    prevent_initial_call=True
)
def update_rotation(period_value):
    """
    Met à jour master-data quand période change
    → Déclenche automatiquement apply_filters()
    """
    print(f"\n🔄 ROTATION : {period_value}")

    if not period_value:
        period_value = "7d"

    try:
        # Invalider cache
        get_df_cached.cache_clear()

        # Recharger avec nouvelle période
        df = run_with_timeout(
            func=load_supply_data,
            kwargs={'period_days': period_value},
            timeout_seconds=60
        )

        print(f"✅ {len(df)} produits chargés")
        print(f"   ADS moy: {df['Average Daily Sales'].mean():.2f}\n")

        return df.to_json(orient="records")

    except TimeoutError:
        print("❌ Timeout\n")
        return no_update

    except Exception as e:
        print(f"❌ Erreur : {e}\n")
        return no_update
'''

# NOTE: Ce callback est désactivé car dupliqué avec refresh_data (ligne ~7740)
'''
@app.callback(
    Output("master-data", "data", allow_duplicate=True),
    Input("btn-refresh", "n_clicks"),
    prevent_initial_call=True
)
def force_refresh_master(n_clicks):
    """Force le rechargement complet des données"""
    if n_clicks:
        df = load_supply_data()  # Recharge depuis sources
        print(f"🔄 Données rechargées : {len(df)} produits")
        return df.to_json(orient="records")
    return no_update
'''

'''@app.callback(
    Output("main-table", "selected_rows", allow_duplicate=True),
    [Input("search-input", "value"),
     Input("filter-supplier", "value"),
     Input("filter-category", "value"),
     Input("filter-need", "value"),
     Input("toggle-options", "value")],
    prevent_initial_call=True
)
def reset_selection_on_filter(search, supplier, category, need, options):
    """
    Réinitialise la sélection à chaque changement de filtre.
    Évite les sélections fantômes après filtrage.
    """
    print("🔄 Réinitialisation des sélections suite à filtrage")
    return []  # ✅ Aucune ligne sélectionnée
'''


# ==================== CALLBACKS NOTES ====================

# Callback 1 : Ouvrir modal + charger liste produits
@app.callback(
    [Output("notes-modal-new", "is_open"),
     Output("note-product-selector", "options")],
    Input("notes-fab", "n_clicks"),
    State("main-table", "data"),
    prevent_initial_call=True
)
def open_notes_modal_fab(n_clicks, table_data):
    """Ouvre le modal et charge la liste des produits"""
    if not n_clicks:
        return False, []

    # Charger liste produits depuis le tableau
    if table_data:
        products = [
            {"label": f"{row.get('product_name', 'N/A')} ({row.get('Supplier', 'N/A')})",
             "value": row.get('product_name', '')}
            for row in table_data
            if row.get('product_name')
        ]
        # Trier alphabétiquement
        products = sorted(products, key=lambda x: x['label'])
    else:
        products = []

    return True, products


# Callback 2 : Fermer modal
@app.callback(
    Output("notes-modal-new", "is_open", allow_duplicate=True),
    Input("note-modal-close", "n_clicks"),
    prevent_initial_call=True
)
def close_notes_modal_fab(n_clicks):
    if n_clicks:
        return False
    return no_update


# Callback 3 : Charger notes du produit sélectionné
@app.callback(
    [Output("notes-display-list", "children"),
     Output("selected-product-for-notes", "data", allow_duplicate=True)],
    Input("note-product-selector", "value"),
    prevent_initial_call=True
)
def load_product_notes(product_name):
    """Charge les notes quand un produit est sélectionné"""
    if not product_name:
        return [
            html.Div([
                html.Div("👆", style={"fontSize": "32px", "marginBottom": "8px"}),
                html.Div(
                    "Sélectionnez un produit pour voir ses notes",
                    style={"color": "#64748b", "fontSize": "14px", "fontWeight": "600"}
                )
            ], style={
                "textAlign": "center",
                "padding": "40px 20px"
            })
        ], None

    notes = get_notes_for_product(product_name)

    if notes:
        notes_display = []
        for note in reversed(notes):
            timestamp = datetime.fromisoformat(note["timestamp"]).strftime("%d/%m/%Y %H:%M")
            message_html = note["message"]

            # Colorier les mentions
            for mention in note.get("mentions", []):
                message_html = message_html.replace(
                    f"@{mention}",
                    f'<span style="background:#dbeafe;color:#1e40af;padding:2px 8px;border-radius:4px;font-weight:700;font-size:12px">@{mention}</span>'
                )

            notes_display.append(
                html.Div([
                    # En-tête
                    html.Div([
                        html.Div([
                            html.Span("👤", style={"marginRight": "8px", "fontSize": "16px"}),
                            html.Strong(note["author"], style={
                                "color": "#1e40af",
                                "fontSize": "14px",
                                "fontWeight": "700"
                            }),
                        ], style={"display": "flex", "alignItems": "center"}),
                        html.Span(
                            timestamp,
                            style={
                                "color": "#64748b",
                                "fontSize": "11px",
                                "fontWeight": "600"
                            }
                        )
                    ], style={
                        "display": "flex",
                        "justifyContent": "space-between",
                        "alignItems": "center",
                        "marginBottom": "10px",
                        "paddingBottom": "8px",
                        "borderBottom": "1px solid #e2e8f0"
                    }),

                    # Message
                    dcc.Markdown(
                        message_html,
                        dangerously_allow_html=True,
                        style={
                            "color": "#1e293b",
                            "fontSize": "13px",
                            "lineHeight": "1.6",
                            "whiteSpace": "pre-wrap",
                            "fontFamily": "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif"
                        }
                    )
                ], style={
                    "marginBottom": "10px",
                    "padding": "14px",
                    "background": "#ffffff",
                    "borderRadius": "10px",
                    "border": "1px solid #e2e8f0",
                    "boxShadow": "0 1px 3px rgba(0,0,0,0.08)",
                    "transition": "all 0.2s ease"
                })
            )
    else:
        notes_display = [
            html.Div([
                html.Div("📝", style={"fontSize": "36px", "marginBottom": "10px"}),
                html.Div(
                    "Aucune note pour ce produit",
                    style={
                        "fontWeight": "700",
                        "color": "#475569",
                        "marginBottom": "6px",
                        "fontSize": "15px"
                    }
                ),
                html.Div(
                    "Soyez le premier à ajouter une note !",
                    style={"fontSize": "12px", "color": "#94a3b8"}
                )
            ], style={
                "textAlign": "center",
                "padding": "50px 20px",
                "background": "#ffffff",
                "borderRadius": "10px",
                "border": "2px dashed #cbd5e1"
            })
        ]

    return notes_display, product_name


# Callback 4 : Envoyer note avec notifications
# Callback 4 : Envoyer note avec notifications
@app.callback(
    [Output("notes-display-list", "children", allow_duplicate=True),
     Output("note-text-input", "value"),
     Output("note-author-input", "value", allow_duplicate=True),
     Output("note-feedback-new", "children")],
    Input("note-modal-send", "n_clicks"),
    [State("note-text-input", "value"),
     State("note-author-input", "value"),
     State("selected-product-for-notes", "data")],  # ✅ ID simple
    prevent_initial_call=True
)
def send_note_with_notifications(n_clicks, message, author, product_name):
    """Envoie note + emails aux mentions"""
    if not n_clicks:
        raise PreventUpdate

    print(f"\n{'=' * 60}")
    print(f"📝 NOUVELLE NOTE")
    print(f"{'=' * 60}")
    print(f"Produit: {product_name}")
    print(f"Auteur: {author}")
    print(f"Message: {message[:50]}...")

    # Vérifications
    if not product_name:
        print("❌ Aucun produit sélectionné")
        return no_update, no_update, no_update, "❌ Sélectionnez un produit d'abord"

    if not message or not message.strip():
        print("❌ Message vide")
        return no_update, no_update, no_update, "❌ Message vide"

    if not author or not author.strip():
        print("❌ Auteur manquant")
        return no_update, no_update, no_update, "❌ Nom d'auteur requis"

    # Extraire mentions
    mentions = re.findall(r'@(\w+)', message)
    print(f"Mentions détectées : {mentions}")

    emails_sent = []
    emails_failed = []

    # Envoyer emails
    for username in mentions:
        username_lower = username.lower()
        if username_lower in TEAM_MEMBERS:
            user_info = TEAM_MEMBERS[username_lower]
            print(f"\n📧 Envoi à {user_info['name']} ({user_info['email']})...")

            try:
                success = send_notification_email(
                    recipient_email=user_info["email"],
                    recipient_name=user_info["name"],
                    product_name=product_name,
                    author=author,
                    message=message
                )

                if success:
                    emails_sent.append(user_info["name"])
                    print(f"  ✅ Envoyé à {user_info['name']}")
                else:
                    emails_failed.append(user_info["name"])
                    print(f"  ❌ Échec pour {user_info['name']}")

            except Exception as e:
                emails_failed.append(user_info["name"])
                print(f"  ❌ Erreur pour {user_info['name']}: {e}")
                import traceback
                traceback.print_exc()
        else:
            print(f"  ⚠️ @{username} non trouvé dans TEAM_MEMBERS")

    # Sauvegarder note
    try:
        add_note(product_name, author, message, mentions)
        print(f"\n✅ Note sauvegardée dans {NOTES_DB_PATH}")
    except Exception as e:
        print(f"❌ Erreur sauvegarde : {e}")
        return no_update, no_update, no_update, f"❌ Erreur sauvegarde : {str(e)}"

    # Recharger notes
    notes = get_notes_for_product(product_name)
    notes_display = []

    for note in reversed(notes):
        timestamp = datetime.fromisoformat(note["timestamp"]).strftime("%d/%m/%Y %H:%M")
        message_html = note["message"]

        # Colorier mentions
        for mention in note.get("mentions", []):
            message_html = message_html.replace(
                f"@{mention}",
                f'<span style="color:#0ea5e9;font-weight:600">@{mention}</span>'
            )

        notes_display.append(
            html.Div([
                html.Div([
                    html.Span("👤", style={"marginRight": "6px"}),
                    html.Strong(note["author"], style={"color": "#0ea5e9"}),
                    html.Span(f" • {timestamp}",
                              style={"color": "#6b7280", "fontSize": "11px", "marginLeft": "6px"})
                ], style={"marginBottom": "6px"}),

                dcc.Markdown(
                    message_html,
                    dangerously_allow_html=True,
                    style={"color": "#1e293b", "fontSize": "13px", "lineHeight": "1.5"}
                ),

                html.Hr(style={"borderColor": "#1f2937", "margin": "10px 0"})
            ], style={"marginBottom": "12px"})
        )

    # Feedback
    if emails_sent:
        feedback = f"✅ Note envoyée • {len(emails_sent)} email(s) : {', '.join(emails_sent)}"
        if emails_failed:
            feedback += f" • ⚠️ Échec : {', '.join(emails_failed)}"
    elif emails_failed:
        feedback = f"⚠️ Note enregistrée mais emails non envoyés : {', '.join(emails_failed)}"
    else:
        feedback = "✅ Note enregistrée (aucune mention)"

    print(f"\n{feedback}")
    print(f"{'=' * 60}\n")

    return notes_display, "", author, feedback

@app.callback(
    [Output('main-table', 'data', allow_duplicate=True),
     Output('action-feedback', 'children', allow_duplicate=True),
     Output('search-input', 'value', allow_duplicate=True),
     Output('filter-supplier', 'value', allow_duplicate=True),
     Output('filter-category', 'value', allow_duplicate=True),
     Output('filter-need', 'value', allow_duplicate=True),
     Output('main-table', 'selected_rows', allow_duplicate=True)],
    Input('btn-refresh', 'n_clicks'),
    State('master-data', 'data'),
    prevent_initial_call=True
)
def refresh_data(n_clicks, current_master_data):
    """
    ⚡ ACTUALISATION RAPIDE - Réinitialise les filtres sans recharger les données
    NE régénère PAS la page entière, juste la table
    """
    if not n_clicks:
        return [no_update] * 7

    try:
        start = time.time()
        print(f"\n⚡ ACTUALISATION RAPIDE")

        # Utiliser les données existantes
        if current_master_data:
            updated_df = pd.DataFrame(json.loads(current_master_data))
        else:
            updated_df = get_df_cached()

        if updated_df.empty:
            return [no_update] * 7

        # S'assurer que QAC edited existe
        if 'QAC edited' not in updated_df.columns:
            updated_df['QAC edited'] = ' '

        # Convertir en records (rapide)
        records = updated_df.to_dict('records')

        elapsed = time.time() - start
        print(f"   ✅ Fait en {elapsed:.2f}s - {len(records)} produits")

        feedback = dbc.Alert([
            html.I(className="fas fa-sync-alt me-2"),
            f"✅ Filtres réinitialisés ({len(records)} produits)"
        ], color="info", duration=1500, dismissable=True)

        # Retourner: data, feedback, search vide, filtres vides, sélection vide
        return records, feedback, "", [], [], [], []

    except Exception as e:
        print(f"❌ Erreur refresh: {e}")
        return [no_update] * 7


# ==========================================
# 📥 CALLBACK RECHARGEMENT COMPLET DONNÉES
# ==========================================
@app.callback(
    [Output('main-table', 'data', allow_duplicate=True),
     Output('master-data', 'data', allow_duplicate=True),
     Output('filtered-data', 'data', allow_duplicate=True),
     Output('action-feedback', 'children', allow_duplicate=True),
     Output('page-container', 'children', allow_duplicate=True)],
    Input('btn-reload-data', 'n_clicks'),
    State('auth-state', 'data'),
    prevent_initial_call=True
)
def reload_data_from_source(n_clicks, auth_state):
    """Recharge TOUTES les données depuis Google Sheets (lent mais complet)"""
    global initial_df

    if not n_clicks:
        return [no_update] * 5

    try:
        print(f"\n{'='*60}")
        print(f"📥 RECHARGEMENT COMPLET DEPUIS GOOGLE SHEETS")
        print(f"{'='*60}")

        username = auth_state.get("username", "") if auth_state else ""

        # Invalider le cache
        invalidate_user_cache(username)
        clear_all_caches()

        # Recharger depuis la source
        updated_df = load_supply_data()

        # Mettre à jour le DataFrame global
        with _data_lock:
            initial_df = updated_df

        if 'QAC edited' not in updated_df.columns:
            updated_df['QAC edited'] = ' '

        records = updated_df.to_dict('records')
        json_data = updated_df.to_json(orient="records")
        new_page_content = page_overview(updated_df)

        print(f"   ✅ {len(updated_df)} produits rechargés")
        print(f"{'='*60}\n")

        feedback = dbc.Alert([
            html.I(className="fas fa-check-circle me-2"),
            f"Données rechargées : {len(updated_df)} produits depuis Google Sheets"
        ], color="warning", duration=4000, dismissable=True)

        return records, json_data, json_data, feedback, new_page_content

    except Exception as e:
        print(f"❌ Erreur reload: {e}")
        import traceback
        traceback.print_exc()

        feedback = dbc.Alert([
            html.I(className="fas fa-exclamation-triangle me-2"),
            f"Erreur rechargement : {str(e)}"
        ], color="danger", duration=5000, dismissable=True)

        return [no_update] * 4 + [feedback]


# ==========================================
# ☑️ CALLBACK SELECT ALL / DESELECT ALL
# ==========================================
@app.callback(
    [Output('main-table', 'selected_rows', allow_duplicate=True),
     Output('btn-select-all', 'children'),
     Output('selection-count-inline', 'children')],
    [Input('btn-select-all', 'n_clicks')],
    [State('main-table', 'data'),
     State('main-table', 'selected_rows')],
    prevent_initial_call=True
)
def toggle_select_all(n_clicks, table_data, current_selection):
    """Sélectionne ou désélectionne toutes les lignes"""
    if not n_clicks or not table_data:
        return no_update, no_update, no_update

    total_rows = len(table_data)
    currently_selected = len(current_selection) if current_selection else 0

    # Si tout est sélectionné, on désélectionne
    if currently_selected == total_rows:
        return [], "☑️ Tout sélect.", ""
    else:
        # Sélectionner toutes les lignes
        all_rows = list(range(total_rows))
        return all_rows, "☐ Tout désélect.", f"✓ {total_rows} produits sélectionnés"


# ==========================================
# 📊 CALLBACK COMPTEUR SÉLECTION
# ==========================================
@app.callback(
    Output('selection-count-inline', 'children', allow_duplicate=True),
    Input('main-table', 'selected_rows'),
    State('main-table', 'data'),
    prevent_initial_call=True
)
def update_selection_count(selected_rows, table_data):
    """Met à jour le compteur de sélection"""
    if not selected_rows:
        return ""
    count = len(selected_rows)
    return f"✓ {count} produit{'s' if count > 1 else ''} sélectionné{'s' if count > 1 else ''}"

@app.callback(
    [Output('edit-modal', 'is_open', allow_duplicate=True),
     Output('edit-modal-title', 'children', allow_duplicate=True),
     Output('edit-product-name', 'value', allow_duplicate=True),
     Output('edit-supplier', 'value', allow_duplicate=True),
     Output('edit-category', 'value', allow_duplicate=True),
     Output('edit-stock', 'value', allow_duplicate=True),
     Output('adding-new-product-flag', 'data')],
    Input('btn-add-row', 'n_clicks'),
    prevent_initial_call=True
)
def open_add_product_modal(n_clicks):
    """Ouvre le modal pour ajouter un nouveau produit"""
    if n_clicks:
        return True, "➕ Ajouter un nouveau produit", "", "", "", 0, True
    return no_update, no_update, no_update, no_update, no_update, no_update, no_update


def page_analytics(master_df: pd.DataFrame = None):
    """Page Analytics avec filtres dynamiques"""
    df = master_df if master_df is not None else get_df_cached()

    # Colonnes nécessaires
    cols_to_keep = [
        "product_name", "Supplier", "Product Category",
        "total_stock", "Average Daily Sales",
        "Max Coverage Day", "ADJUSTED_LEADTIME",
        "Ajusted_total_need", "purchase_need",
        "QAC", "optimal stock", "credit_days"
    ]
    cols_to_keep = [c for c in cols_to_keep if c in df.columns]
    analytics_df = df[cols_to_keep].dropna()

    # KPI cards
    kpi_cards, _ = make_kpis(df)

    # ========================================
    # 🎛️ FILTRES DYNAMIQUES
    # ========================================

    # Options pour les dropdowns
    suppliers_list = ["Tous"] + sorted(analytics_df["Supplier"].unique().tolist())
    categories_list = ["Toutes"] + sorted(analytics_df["Product Category"].unique().tolist())
    needs_list = ["Tous"] + sorted(analytics_df["Ajusted_total_need"].unique().tolist())

    filters_section = html.Div([
        dbc.Row([
            dbc.Col([
                html.Label(" Fournisseur", style={
                    "fontWeight": "700",
                    "color": "#1e293b",
                    "marginBottom": "8px",
                    "fontSize": "14px"
                }),
                dcc.Dropdown(
                    id="analytics-filter-supplier",
                    options=[{"label": s, "value": s} for s in suppliers_list],
                    value="Tous",
                    clearable=False,
                    className="dark-dropdown"
                )
            ], md=4),

            dbc.Col([
                html.Label(" Catégorie", style={
                    "fontWeight": "700",
                    "color": "#1e293b",
                    "marginBottom": "8px",
                    "fontSize": "14px"
                }),
                dcc.Dropdown(
                    id="analytics-filter-category",
                    options=[{"label": c, "value": c} for c in categories_list],
                    value="Toutes",
                    clearable=False,
                    className="dark-dropdown"
                )
            ], md=4),

            dbc.Col([
                html.Label(" Besoin d'achat", style={
                    "fontWeight": "700",
                    "color": "#1e293b",
                    "marginBottom": "8px",
                    "fontSize": "14px"
                }),
                dcc.Dropdown(
                    id="analytics-filter-need",
                    options=[{"label": n, "value": n} for n in needs_list],
                    value="Tous",
                    clearable=False,
                    className="dark-dropdown"
                )
            ], md=4)
        ], className="mb-3"),

        # Indicateur de filtrage
        html.Div(id="analytics-filter-indicator", style={
            "marginTop": "12px",
            "padding": "10px 14px",
            "background": "rgba(34, 211, 238, 0.1)",
            "border": "1px solid rgba(34, 211, 238, 0.3)",
            "borderRadius": "8px",
            "fontSize": "12px",
            "color": "#0ea5e9",
            "fontWeight": "600"
        })
    ], style={
        "padding": "20px",
        "background": "linear-gradient(135deg, #1a2332 0%, #141b2d 100%)",
        "borderRadius": "16px",
        "border": "1px solid #2d3748",
        "marginBottom": "24px",
        "boxShadow": "0 4px 12px rgba(0, 0, 0, 0.3)"
    })

    return html.Div(className="content", children=[
        dbc.Row([
            dbc.Col(html.H2(" Analyses Avancées", style={"color": "#0ea5e9"}), md=12)
        ]),
        html.Div(kpi_cards),
        html.Br(),

        # Filtres
        filters_section,

        # Graphiques dynamiques
        dbc.Row([
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5(" Lead Time vs Couverture Stock", className="section-title"),
                    dcc.Graph(id="analytics-scatter")
                ])
            ], md=6),
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5(" Distribution du Stock", className="section-title"),
                    dcc.Graph(id="analytics-hist-stock")
                ])
            ], md=6)
        ]),
        html.Br(),

        dbc.Row([
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5(" Ventes Moyennes par Catégorie", className="section-title"),
                    dcc.Graph(id="analytics-box-category")
                ])
            ], md=6),
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5(" Top 10 Besoins par Fournisseur", className="section-title"),
                    dcc.Graph(id="analytics-purchase-supplier")
                ])
            ], md=6)
        ]),
        html.Br(),

        dbc.Row([
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5(" QAC vs Stock Optimal", className="section-title"),
                    dcc.Graph(id="analytics-qac-optimal")
                ])
            ], md=6),
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5(" Répartition par Besoin d'Achat", className="section-title"),
                    dcc.Graph(id="analytics-pie-need")
                ])
            ], md=6)
        ])
    ])


# ==========================================
# 📊 CALLBACKS ANALYTICS DYNAMIQUES
# ==========================================

# ==========================================
# 📊 CALLBACK ANALYTICS AVEC GESTION D'ERREURS
# ==========================================

# NOTE: Ce callback est désactivé car dupliqué avec update_analytics_all_charts (ligne ~9232)
'''
@app.callback(
    [Output("analytics-scatter", "figure"),
     Output("analytics-hist-stock", "figure"),
     Output("analytics-box-category", "figure"),
     Output("analytics-purchase-supplier", "figure"),
     Output("analytics-qac-optimal", "figure"),
     Output("analytics-pie-need", "figure"),
     Output("analytics-filter-indicator", "children")],
    [Input("analytics-filter-supplier", "value"),
     Input("analytics-filter-category", "value"),
     Input("analytics-filter-need", "value"),
     Input("master-data", "data")],
    prevent_initial_call=False
)
def update_analytics_charts(supplier_value, category_value, need_value, master_json):
    """
    Met à jour tous les graphiques analytics selon les filtres
    """
    print(f"\n{'=' * 60}")
    print(f"🔄 UPDATE ANALYTICS")
    print(f"{'=' * 60}")
    print(f"Filtres: Supplier={supplier_value}, Category={category_value}, Need={need_value}")

    try:
        # ✅ Charger données
        if master_json:
            df = pd.DataFrame(json.loads(master_json))
            print(f"📥 Données chargées depuis master-data : {len(df)} lignes")
        else:
            df = get_df_cached()
            print(f"📥 Données chargées depuis cache : {len(df)} lignes")

        # ✅ Vérifier colonnes disponibles
        required_cols = [
            "product_name", "Supplier", "Product Category",
            "total_stock", "Average Daily Sales",
            "Max Coverage Day", "ADJUSTED_LEADTIME",
            "Ajusted_total_need", "purchase_need",
            "QAC", "optimal stock"
        ]

        available_cols = [c for c in required_cols if c in df.columns]
        missing_cols = [c for c in required_cols if c not in df.columns]

        print(f"✅ Colonnes disponibles : {len(available_cols)}/{len(required_cols)}")
        if missing_cols:
            print(f"⚠️  Colonnes manquantes : {missing_cols}")

        # ✅ Filtrer colonnes et nettoyer
        df = df[available_cols].copy()

        # Nettoyer les NaN
        initial_len = len(df)
        df = df.dropna(subset=['product_name', 'Supplier'])
        print(f"🧹 Nettoyage NaN : {initial_len} → {len(df)} lignes")

        # Remplir les NaN numériques avec 0
        numeric_cols = ['total_stock', 'Average Daily Sales', 'Max Coverage Day',
                        'ADJUSTED_LEADTIME', 'purchase_need', 'QAC', 'optimal stock']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        initial_count = len(df)
        print(f"📊 Dataset initial : {initial_count} produits")

        # ✅ APPLIQUER FILTRES
        if supplier_value and supplier_value != "Tous":
            df = df[df["Supplier"] == supplier_value]
            print(f"   🏭 Filtre fournisseur '{supplier_value}' : {len(df)} produits")

        if category_value and category_value != "Toutes":
            df = df[df["Product Category"] == category_value]
            print(f"   🏷️ Filtre catégorie '{category_value}' : {len(df)} produits")

        if need_value and need_value != "Tous":
            df = df[df["Ajusted_total_need"] == need_value]
            print(f"   📦 Filtre besoin '{need_value}' : {len(df)} produits")

        filtered_count = len(df)
        print(f"✅ Dataset filtré : {filtered_count} produits")

        # ✅ VÉRIFICATION : Dataset vide ?
        if filtered_count == 0:
            print("⚠️  AUCUNE DONNÉE après filtrage !")

            # Créer des graphiques vides avec message
            empty_fig = go.Figure()
            empty_fig.add_annotation(
                text="Aucune donnée disponible avec ces filtres",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="#94a3b8")
            )
            empty_fig.update_layout(
                plot_bgcolor="#f8fafc",
                paper_bgcolor="#ffffff",
                xaxis=dict(visible=False),
                yaxis=dict(visible=False)
            )

            indicator = html.Div([
                html.Span("⚠️ Aucune donnée avec ces filtres", style={
                    "fontWeight": "700",
                    "color": "#ef4444"
                })
            ])

            return (empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, indicator)

        # ✅ INDICATEUR DE FILTRE
        filter_text = []
        if supplier_value != "Tous":
            filter_text.append(f"🏭 {supplier_value}")
        if category_value != "Toutes":
            filter_text.append(f"🏷️ {category_value}")
        if need_value != "Tous":
            filter_text.append(f"📦 {need_value}")

        if filter_text:
            indicator = html.Div([
                html.Span("🔍 Filtres actifs : ", style={"fontWeight": "700"}),
                html.Span(" • ".join(filter_text)),
                html.Span(f" ({filtered_count} produits)", style={
                    "marginLeft": "10px",
                    "opacity": "0.8"
                })
            ])
        else:
            indicator = html.Div([
                html.Span("✨ Tous les produits", style={"fontWeight": "700"}),
                html.Span(f" ({filtered_count} produits)", style={
                    "marginLeft": "10px",
                    "opacity": "0.8"
                })
            ])

        # ========================================
        # 📊 GRAPHIQUE 1 : SCATTER LEAD TIME VS COVERAGE
        # ========================================
        print("📊 Création scatter plot...")

        if all(col in df.columns for col in ["ADJUSTED_LEADTIME", "Max Coverage Day"]):
            fig_scatter = px.scatter(
                df,
                x="ADJUSTED_LEADTIME",
                y="Max Coverage Day",
                color="Ajusted_total_need" if "Ajusted_total_need" in df.columns else None,
                hover_data=["product_name", "Supplier", "purchase_need", "QAC"],
                labels={
                    "ADJUSTED_LEADTIME": "Lead Time Ajusté (jours)",
                    "Max Coverage Day": "Couverture Stock (jours)"
                },
                color_discrete_map={
                    "ORDER NOW": "#ef4444",
                    "ORDER NOT URGENT": "#f59e0b",
                    "NO NEED": "#10b981"
                }
            )

            fig_scatter.update_layout(
                plot_bgcolor="#f8fafc",
                paper_bgcolor="#ffffff",
                font=dict(color="#374151", size=12),
                showlegend=True,
                legend=dict(
                    bgcolor="rgba(15, 22, 37, 0.8)",
                    bordercolor="#2d3748",
                    borderwidth=1
                ),
                height=400
            )
            print("   ✅ Scatter plot créé")
        else:
            fig_scatter = go.Figure()
            fig_scatter.add_annotation(
                text="Colonnes manquantes pour ce graphique",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False
            )
            print("   ⚠️ Scatter plot - colonnes manquantes")

        # ========================================
        # 📊 GRAPHIQUE 2 : HISTOGRAMME STOCK
        # ========================================
        print("📊 Création histogramme...")

        if "total_stock" in df.columns:
            fig_hist_stock = px.histogram(
                df,
                x="total_stock",
                nbins=30,
                labels={"total_stock": "Stock Total"},
                color_discrete_sequence=["#0ea5e9"]
            )

            fig_hist_stock.update_layout(
                plot_bgcolor="#f8fafc",
                paper_bgcolor="#ffffff",
                font=dict(color="#374151", size=12),
                showlegend=False,
                height=400
            )
            print("   ✅ Histogramme créé")
        else:
            fig_hist_stock = go.Figure()
            print("   ⚠️ Histogramme - colonne manquante")

        # ========================================
        # 📊 GRAPHIQUE 3 : BOXPLOT PAR CATÉGORIE
        # ========================================
        print("📊 Création boxplot...")

        if all(col in df.columns for col in ["Product Category", "Average Daily Sales"]):
            fig_box_category = px.box(
                df,
                x="Product Category",
                y="Average Daily Sales",
                color="Product Category",
                labels={"Average Daily Sales": "Ventes Moyennes (ADS)"}
            )

            fig_box_category.update_layout(
                plot_bgcolor="#f8fafc",
                paper_bgcolor="#ffffff",
                font=dict(color="#374151", size=12),
                showlegend=False,
                height=400
            )
            print("   ✅ Boxplot créé")
        else:
            fig_box_category = go.Figure()
            print("   ⚠️ Boxplot - colonnes manquantes")

        # ========================================
        # 📊 GRAPHIQUE 4 : TOP 10 FOURNISSEURS
        # ========================================
        print("📊 Création bar chart fournisseurs...")

        if all(col in df.columns for col in ["Supplier", "purchase_need"]):
            purchase_by_supplier = (
                df.groupby("Supplier")["purchase_need"]
                .sum()
                .reset_index()
                .sort_values("purchase_need", ascending=False)
                .head(10)
            )

            print(f"   📊 Top 10 fournisseurs : {len(purchase_by_supplier)} lignes")

            if len(purchase_by_supplier) > 0:
                fig_purchase_supplier = px.bar(
                    purchase_by_supplier,
                    x="Supplier",
                    y="purchase_need",
                    labels={"purchase_need": "Besoin d'Achat Total"},
                    color="purchase_need",
                    color_continuous_scale="Reds"
                )

                fig_purchase_supplier.update_layout(
                    plot_bgcolor="#f8fafc",
                    paper_bgcolor="#ffffff",
                    font=dict(color="#374151", size=12),
                    xaxis=dict(tickangle=-45),
                    showlegend=False,
                    height=400
                )
                print("   ✅ Bar chart fournisseurs créé")
            else:
                fig_purchase_supplier = go.Figure()
                print("   ⚠️ Bar chart fournisseurs - pas de données")
        else:
            fig_purchase_supplier = go.Figure()
            print("   ⚠️ Bar chart fournisseurs - colonnes manquantes")

        # ========================================
        # 📊 GRAPHIQUE 5 : QAC VS OPTIMAL
        # ========================================
        print("📊 Création scatter QAC vs Optimal...")

        if all(col in df.columns for col in ["QAC", "optimal stock"]):
            fig_qac_optimal = px.scatter(
                df,
                x="QAC",
                y="optimal stock",
                color="Ajusted_total_need" if "Ajusted_total_need" in df.columns else None,
                hover_data=["product_name", "Supplier"],
                labels={"QAC": "Quantité Ajustée Commandée"},
                color_discrete_map={
                    "ORDER NOW": "#ef4444",
                    "ORDER NOT URGENT": "#f59e0b",
                    "NO NEED": "#10b981"
                }
            )

            fig_qac_optimal.update_layout(
                plot_bgcolor="#f8fafc",
                paper_bgcolor="#ffffff",
                font=dict(color="#374151", size=12),
                showlegend=True,
                height=400
            )
            print("   ✅ Scatter QAC créé")
        else:
            fig_qac_optimal = go.Figure()
            print("   ⚠️ Scatter QAC - colonnes manquantes")

        # ========================================
        # 📊 GRAPHIQUE 6 : PIE CHART BESOIN D'ACHAT
        # ========================================
        print("📊 Création pie chart...")

        if "Ajusted_total_need" in df.columns:
            need_distribution = df["Ajusted_total_need"].value_counts().reset_index()
            need_distribution.columns = ["Besoin", "Nombre"]

            print(f"   📊 Distribution besoins : {len(need_distribution)} catégories")

            fig_pie_need = px.pie(
                need_distribution,
                values="Nombre",
                names="Besoin",
                color="Besoin",
                color_discrete_map={
                    "ORDER NOW": "#ef4444",
                    "ORDER NOT URGENT": "#f59e0b",
                    "NO NEED": "#10b981"
                }
            )

            fig_pie_need.update_layout(
                plot_bgcolor="#f8fafc",
                paper_bgcolor="#ffffff",
                font=dict(color="#374151", size=12),
                height=400
            )
            print("   ✅ Pie chart créé")
        else:
            fig_pie_need = go.Figure()
            print("   ⚠️ Pie chart - colonne manquante")

        print(f"{'=' * 60}")
        print(f"✅ ANALYTICS MIS À JOUR")
        print(f"{'=' * 60}\n")

        return (
            fig_scatter,
            fig_hist_stock,
            fig_box_category,
            fig_purchase_supplier,
            fig_qac_optimal,
            fig_pie_need,
            indicator
        )

    except Exception as e:
        print(f"\n❌ ERREUR dans update_analytics_charts : {e}")
        import traceback
        traceback.print_exc()

        # Retourner graphiques vides en cas d'erreur
        empty_fig = go.Figure()
        empty_fig.add_annotation(
            text=f"Erreur : {str(e)[:50]}",
            xref="paper", yref="paper",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=14, color="#ef4444")
        )
        empty_fig.update_layout(
            plot_bgcolor="#f8fafc",
            paper_bgcolor="#ffffff"
        )

        error_indicator = html.Div([
            html.Span(f"❌ Erreur : {str(e)[:100]}", style={
                "fontWeight": "700",
                "color": "#ef4444"
            })
        ])

        return (empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, error_indicator)
'''

def page_predictive(master_df: pd.DataFrame = None):
    """Page Prédictions avec filtres dynamiques"""
    df = master_df if master_df is not None else get_df_cached()

    # Colonnes nécessaires
    cols_to_keep = [
        "product_id", "product_name", "Supplier",
        "total_stock", "Average Daily Sales",
        "Max Daily Sales (Pikine)", "target_quantity",
        "Product Category", "Ajusted_total_need",
        "total_sold_30d", "avg_daily_sold",
        "std_sold", "max_daily_sold",
        "days_with_sales", "trend_factor",
        "projected_demand", "safety_stock",
        "target_supervised"
    ]
    cols_to_keep = [c for c in cols_to_keep if c in df.columns]
    pred_df = df[cols_to_keep].copy()

    # Calcul stockout rate
    if 'avg_daily_sold' in pred_df.columns and 'total_stock' in pred_df.columns:
        pred_df_ml = pred_df[pred_df['avg_daily_sold'].notna()].copy()

        if len(pred_df_ml) > 0:
            pred_df_ml['coverage_days'] = np.where(
                pred_df_ml['avg_daily_sold'] > 0,
                pred_df_ml['total_stock'] / pred_df_ml['avg_daily_sold'],
                999
            )
            replenishment = 21  # Valeur par défaut
            at_risk = (pred_df_ml['coverage_days'] < replenishment).sum()
            stockout_rate = (at_risk / len(pred_df_ml)) * 100
        else:
            stockout_rate = 0.0
    else:
        stockout_rate = 0.0

    # Options pour dropdowns
    suppliers_list = ["Tous"] + sorted(pred_df["Supplier"].unique().tolist())
    categories_list = ["Toutes"] + sorted(pred_df["Product Category"].unique().tolist())
    needs_list = ["Tous"] + sorted(pred_df["Ajusted_total_need"].unique().tolist())

    # ========================================
    # 🎛️ FILTRES
    # ========================================
    filters_section = html.Div([
        dbc.Row([
            dbc.Col([
                html.Label("🏭 Fournisseur", style={
                    "fontWeight": "700",
                    "color": "#1e293b",
                    "marginBottom": "8px"
                }),
                dcc.Dropdown(
                    id="predictive-filter-supplier",
                    options=[{"label": s, "value": s} for s in suppliers_list],
                    value="Tous",
                    clearable=False,
                    className="dark-dropdown"
                )
            ], md=4),

            dbc.Col([
                html.Label("🏷️ Catégorie", style={
                    "fontWeight": "700",
                    "color": "#1e293b",
                    "marginBottom": "8px"
                }),
                dcc.Dropdown(
                    id="predictive-filter-category",
                    options=[{"label": c, "value": c} for c in categories_list],
                    value="Toutes",
                    clearable=False,
                    className="dark-dropdown"
                )
            ], md=4),

            dbc.Col([
                html.Label("📦 Besoin d'achat", style={
                    "fontWeight": "700",
                    "color": "#1e293b",
                    "marginBottom": "8px"
                }),
                dcc.Dropdown(
                    id="predictive-filter-need",
                    options=[{"label": n, "value": n} for n in needs_list],
                    value="Tous",
                    clearable=False,
                    className="dark-dropdown"
                )
            ], md=4)
        ], className="mb-3"),

        html.Div(id="predictive-filter-indicator", style={
            "marginTop": "12px",
            "padding": "10px 14px",
            "background": "rgba(34, 211, 238, 0.1)",
            "border": "1px solid rgba(34, 211, 238, 0.3)",
            "borderRadius": "8px",
            "fontSize": "12px",
            "color": "#0ea5e9",
            "fontWeight": "600"
        })
    ], style={
        "padding": "20px",
        "background": "linear-gradient(135deg, #1a2332 0%, #141b2d 100%)",
        "borderRadius": "16px",
        "border": "1px solid #2d3748",
        "marginBottom": "24px"
    })

    return html.Div(className="content", children=[
        dbc.Row([
            dbc.Col(html.H3("🔮 Prédictions ML", style={"color": "#0ea5e9"}), md=8),
            dbc.Col(
                dbc.Badge(
                    f"Stockout prédit : {stockout_rate:.1f}%",
                    color="danger" if stockout_rate > 15 else "warning" if stockout_rate > 5 else "success",
                    style={"fontSize": "14px", "padding": "8px 15px"}
                ),
                md=4,
                style={"display": "flex", "justifyContent": "flex-end", "alignItems": "center"}
            )
        ], className="mb-4"),

        # Filtres
        filters_section,

        # Graphiques
        dbc.Row([
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("📊 Top 20 Produits - Target Quantity", className="section-title"),
                    dcc.Graph(id="predictive-bar")
                ])
            ], md=12)
        ]),
        html.Br(),

        dbc.Row([
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("🎯 Demande Projetée vs Stock Actuel", className="section-title"),
                    dcc.Graph(id="predictive-scatter")
                ])
            ], md=6),
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("📈 Safety Stock par Catégorie", className="section-title"),
                    dcc.Graph(id="predictive-safety-stock")
                ])
            ], md=6)
        ]),
        html.Br(),

        # Tableau
        html.Div(className="soft-card", children=[
            html.H5("📋 Données Prédictives Détaillées", className="section-title"),
            dash_table.DataTable(
                id="predictive-table",
                columns=[{"name": c, "id": c} for c in pred_df.columns],
                data=pred_df.to_dict("records"),
                page_size=15,
                filter_action="native",
                sort_action="native",
                sort_mode="multi",
                style_table={"overflowX": "auto"},
                style_header={
                    "backgroundColor": "#1e3a5f",
                    "border": "1px solid #2d3748",
                    "fontWeight": "700",
                    "textAlign": "center",
                    "color": "#1e293b"
                },
                style_cell={
                    "backgroundColor": "#f8fafc",
                    "color": "#1e293b",
                    "border": "1px solid #1f2937",
                    "fontSize": 12,
                    "textAlign": "center"
                },
                style_data_conditional=[
                    {
                        "if": {"column_id": "target_quantity"},
                        "fontWeight": "700",
                        "color": "#0ea5e9",
                        "fontSize": "14px"
                    }
                ]
            )
        ])
    ])


# ==========================================
# 🔮 CALLBACKS PRÉDICTIONS DYNAMIQUES
# ==========================================

# NOTE: Ce callback est désactivé car dupliqué avec update_predictive_all_charts (ligne ~9626)
'''
@app.callback(
    [Output("predictive-bar", "figure"),
     Output("predictive-scatter", "figure"),
     Output("predictive-safety-stock", "figure"),
     Output("predictive-table", "data"),
     Output("predictive-filter-indicator", "children")],
    [Input("predictive-filter-supplier", "value"),
     Input("predictive-filter-category", "value"),
     Input("predictive-filter-need", "value"),
     Input("master-data", "data")],
    prevent_initial_call=False
)
def update_predictive_charts(supplier_value, category_value, need_value, master_json):
    """
    Met à jour tous les graphiques prédictifs selon les filtres
    """
    print(f"\n🔮 Update Prédictions - Filtres: {supplier_value}, {category_value}, {need_value}")

    # Charger données
    if master_json:
        df = pd.DataFrame(json.loads(master_json))
    else:
        df = get_df_cached()

    # Colonnes nécessaires
    cols = [
        "product_name", "Supplier", "Product Category",
        "total_stock", "Average Daily Sales",
        "target_quantity", "Ajusted_total_need",
        "avg_daily_sold", "projected_demand",
        "safety_stock", "total_sold_30d"
    ]
    df = df[[c for c in cols if c in df.columns]].copy()

    initial_count = len(df)

    # Appliquer filtres
    if supplier_value and supplier_value != "Tous":
        df = df[df["Supplier"] == supplier_value]

    if category_value and category_value != "Toutes":
        df = df[df["Product Category"] == category_value]

    if need_value and need_value != "Tous":
        df = df[df["Ajusted_total_need"] == need_value]

    filtered_count = len(df)

    print(f"   📊 {initial_count} → {filtered_count} produits après filtres")

    # Indicateur
    filter_text = []
    if supplier_value != "Tous":
        filter_text.append(f"🏭 {supplier_value}")
    if category_value != "Toutes":
        filter_text.append(f"🏷️ {category_value}")
    if need_value != "Tous":
        filter_text.append(f"📦 {need_value}")

    if filter_text:
        indicator = html.Div([
            html.Span("🔍 Filtres actifs : ", style={"fontWeight": "700"}),
            html.Span(" • ".join(filter_text)),
            html.Span(f" ({filtered_count} produits)", style={"marginLeft": "10px"})
        ])
    else:
        indicator = html.Div([
            html.Span("✨ Tous les produits affichés", style={"fontWeight": "700"}),
            html.Span(f" ({filtered_count} produits)", style={"marginLeft": "10px"})
        ])

    # ========================================
    # 📊 GRAPHIQUE 1 : BAR CHART TOP 20
    # ========================================
    df_chart = df[df['target_quantity'] > 0].copy() if 'target_quantity' in df.columns else df.copy()

    fig_bar = px.bar(
        df_chart.sort_values("target_quantity", ascending=False).head(20),
        x="product_name",
        y="target_quantity",
        color="Ajusted_total_need",
        hover_data=["Supplier", "total_stock", "avg_daily_sold"],
        labels={"target_quantity": "Quantité à Commander"},
        color_discrete_map={
            "ORDER NOW": "#ef4444",
            "ORDER NOT URGENT": "#f59e0b",
            "NO NEED": "#10b981"
        }
    )

    fig_bar.update_layout(
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(color="#374151"),
        xaxis=dict(tickangle=-45)
    )

    # ========================================
    # 📊 GRAPHIQUE 2 : SCATTER DEMANDE VS STOCK
    # ========================================
    if 'projected_demand' in df.columns:
        fig_scatter = px.scatter(
            df,
            x="total_stock",
            y="projected_demand",
            color="Ajusted_total_need",
            size="Average Daily Sales",
            hover_data=["product_name", "Supplier"],
            labels={
                "total_stock": "Stock Actuel",
                "projected_demand": "Demande Projetée"
            },
            color_discrete_map={
                "ORDER NOW": "#ef4444",
                "ORDER NOT URGENT": "#f59e0b",
                "NO NEED": "#10b981"
            }
        )
    else:
        fig_scatter = px.scatter(
            df,
            x="total_stock",
            y="Average Daily Sales",
            color="Ajusted_total_need",
            hover_data=["product_name", "Supplier"]
        )

    fig_scatter.update_layout(
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(color="#374151")
    )

    # ========================================
    # 📊 GRAPHIQUE 3 : SAFETY STOCK PAR CATÉGORIE
    # ========================================
    if 'safety_stock' in df.columns:
        safety_by_cat = (
            df.groupby("Product Category")["safety_stock"]
            .sum()
            .reset_index()
            .sort_values("safety_stock", ascending=False)
        )

        fig_safety = px.bar(
            safety_by_cat,
            x="Product Category",
            y="safety_stock",
            labels={"safety_stock": "Safety Stock Total"},
            color="safety_stock",
            color_continuous_scale="Blues"
        )
    else:
        fig_safety = px.bar(
            df.groupby("Product Category").size().reset_index(name="count"),
            x="Product Category",
            y="count"
        )

    fig_safety.update_layout(
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(color="#374151"),
        showlegend=False
    )

    print(f"   ✅ Prédictions mises à jour\n")

    return (
        fig_bar,
        fig_scatter,
        fig_safety,
        df.to_dict("records"),
        indicator
    )
'''

def page_analytics(master_df: pd.DataFrame = None):
    df = master_df if master_df is not None else get_df_cached()

    # Colonnes nécessaires
    cols_to_keep = [
        "product_name", "Supplier", "Product Category",
        "total_stock", "Average Daily Sales",
        "Max Coverage Day", "ADJUSTED_LEADTIME",
        "Ajusted_total_need", "purchase_need",
        "QAC", "optimal stock"
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
        y="Average Daily Sales",
        title="Distribution des ventes moyennes par catégorie",
        labels={"Average Daily Sales": "Ventes moyennes recalculées"}
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

    # Scatter QAC vs optimal stock
    fig_qac_vs_optimal = px.scatter(
        analytics_df,
        x="QAC",
        y="optimal stock",
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
# @app.callback(
#   Output("analytics-scatter", "figure"),
#  Input("analytics-filter-supplier", "value"),
# Input("analytics-filter-category", "value"),
# prevent_initial_call=False
# )
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

    # ✅ Virgule manquante corrigée
    cols_to_keep = [
        "product_id",
        "product_name",
        "Supplier",
        "total_stock",
        "Average Daily Sales",
        "Max Daily Sales (Pikine)",
        "target_quantity",
        "Product Category",
        "total_sold_30d",
        "avg_daily_sold",
        "std_sold",
        "max_daily_sold",
        "days_with_sales",
        "trend_factor",
        "projected_demand",
        "safety_stock",
        "target_supervised"
    ]
    cols_to_keep = [c for c in cols_to_keep if c in df.columns]
    pred_df = df[cols_to_keep].copy()

    # ✅ Calcul stockout_rate corrigé (basé sur données ML)
    if 'avg_daily_sold' in pred_df.columns and 'total_stock' in pred_df.columns:
        # Produits avec historique ML
        pred_df_ml = pred_df[pred_df['avg_daily_sold'].notna()].copy()

        if len(pred_df_ml) > 0:
            # Calculer couverture en jours
            pred_df_ml['coverage_days'] = np.where(
                pred_df_ml['avg_daily_sold'] > 0,
                pred_df_ml['total_stock'] / pred_df_ml['avg_daily_sold'],
                999
            )

            # Période de réapprovisionnement (leadtime + credit)
            # replenishment = pred_df_ml.get('ADJUSTED_LEADTIME', 7).fillna(7) + pred_df_ml.get('credit_days', 14).fillna(
            #   14)
            replenishment = pred_df_ml.get('ADJUSTED_LEADTIME', 7) + pred_df_ml.get('credit_days', 14)

            # Risque si couverture < période réapprovisionnement
            at_risk = (pred_df_ml['coverage_days'] < replenishment).sum()
            stockout_rate = (at_risk / len(pred_df_ml)) * 100

            print(f"📊 Stockout analysis : {at_risk}/{len(pred_df_ml)} produits à risque")
        else:
            stockout_rate = 0.0
    else:
        stockout_rate = 0.0

    # Avant de créer le graphique
    available_hover = []
    for col in ["Supplier", "total_stock", "avg_daily_sold", "projected_demand"]:
        if col in pred_df.columns:
            available_hover.append(col)

    # ✅ Filtrer pour graphique (uniquement produits avec target > 0)
    pred_df_chart = pred_df[
        pred_df['target_quantity'] > 0].copy() if 'target_quantity' in pred_df.columns else pred_df.copy()

    fig_bar = px.bar(
        pred_df_chart.sort_values("target_quantity", ascending=False).head(20),
        x="product_name",
        y="target_quantity",
        color="Product Category",
        hover_data=available_hover,
        title="Top 20 produits par besoin de commande",
        labels={"target_quantity": "Quantité à commander", "product_name": "Produit"}
    )

    # Style du graphique
    fig_bar.update_layout(
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(color="#374151"),
        xaxis=dict(tickangle=-45)
    )

    return html.Div(className="content", children=[
        dbc.Row([
            dbc.Col(html.H3(" Prédictions ", className="page-title"), md=8),
            dbc.Col(
                dbc.Badge(
                    f"Stockout prédit : {stockout_rate:.1f}%",
                    color="danger" if stockout_rate > 15 else "warning" if stockout_rate > 5 else "success",
                    className="ms-auto",
                    style={"fontSize": "14px", "padding": "8px 15px"}
                ),
                md=4,
                style={"display": "flex", "justifyContent": "flex-end", "alignItems": "center"}
            )
        ], className="mb-4"),
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
            sort_action="native",
            sort_mode="multi",
            style_table={"overflowX": "auto"},
            style_header={
                "backgroundColor": "#1e3a5f",
                "border": "1px solid #1f2937",
                "fontWeight": "700",
                "textAlign": "center"
            },
            style_cell={
                "backgroundColor": "#f8fafc",
                "color": "#1e293b",
                "border": "1px solid #1f2937",
                "fontSize": 12,
                "textAlign": "center"
            },
            style_data_conditional=[
                {
                    "if": {"column_id": "target_quantity"},
                    "fontWeight": "bold",
                    "color": "#10b981"
                }
            ]
        )
    ])


# NOTE: Ce callback est désactivé car dupliqué avec update_predictive_all_charts
'''
# Callback : mettre à jour le bar chart prédictif
@app.callback(
    Output("predictive-bar", "figure"),
    Input("predictive-filter-supplier", "value"),
    Input("predictive-filter-category", "value"),
    prevent_initial_call=False
)
def update_predictive_bar(supplier_value, category_value):
    df = get_df_cached()

    # ✅ Colonnes mises à jour (retirer celles qui n'existent plus)
    cols = [
        "product_name",
        "Supplier",
        "Product Category",
        "target_quantity",
        "Ajusted_total_need",
        "total_stock",
        "avg_daily_sold",
        "projected_demand"
    ]
    df = df[[c for c in cols if c in df.columns]].copy()

    # Filtres
    if supplier_value and supplier_value != "Tous":
        df = df[df["Supplier"] == supplier_value]
    if category_value and category_value != "Toutes":
        df = df[df["Product Category"] == category_value]

    # ✅ Hover_data avec colonnes existantes
    available_hover = []
    for col in ["Supplier", "total_stock", "avg_daily_sold", "projected_demand"]:
        if col in df.columns:
            available_hover.append(col)

    fig = px.bar(
        df.sort_values("target_quantity", ascending=False).head(20),
        x="product_name",
        y="target_quantity",
        color="Ajusted_total_need",
        hover_data=available_hover,  # ✅ Colonnes valides
        labels={"target_quantity": "Qté de commande prédite"},
        title="Top 20 produits par besoin de commande"
    )

    # Style cohérent avec le reste de l'app
    fig.update_layout(
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(color="#374151"),
        xaxis=dict(tickangle=-45)
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

    # ✅ Colonnes ML prédictives
    cols = [
        "product_id",
        "product_name",
        "Supplier",
        "total_stock",
        "Average Daily Sales",
        "Max Daily Sales (Pikine)",
        "target_quantity",
        "Product Category",
        "total_sold_30d",
        "avg_daily_sold",
        "std_sold",
        "max_daily_sold",
        "days_with_sales",
        "trend_factor",
        "projected_demand",
        "safety_stock",
        "target_supervised"
    ]
    df = df[[c for c in cols if c in df.columns]].copy()

    # Filtres
    if supplier_value and supplier_value != "Tous":
        df = df[df["Supplier"] == supplier_value]
    if category_value and category_value != "Toutes":
        df = df[df["Product Category"] == category_value]

    return df.to_dict("records")
'''

# ==========================================
# 🔧 HELPER : PLACEHOLDER POUR ÉVITER ERREURS
# ==========================================

# NOTE: Cette fonction est désactivée - main-table est défini dans le layout principal
'''
def create_hidden_table_placeholder():
    """
    Crée un placeholder caché pour main-table
    Nécessaire pour éviter les erreurs de callbacks sur les autres pages
    """
    return dash_table.DataTable(
        id="main-table",
        data=[],
        columns=[],
        style_table={"display": "none"}
    )
'''

def page_analytics(master_df: pd.DataFrame = None):
    """Page Analytics avec filtres sidebar et graphiques dynamiques"""
    df = master_df if master_df is not None else get_df_cached()

    # Colonnes nécessaires
    cols_to_keep = [
        "product_name", "Supplier", "Product Category",
        "total_stock", "Average Daily Sales",
        "Max Coverage Day", "ADJUSTED_LEADTIME",
        "Ajusted_total_need", "purchase_need",
        "QAC", "optimal stock", "credit_days"
    ]
    cols_to_keep = [c for c in cols_to_keep if c in df.columns]
    analytics_df = df[cols_to_keep].dropna()

    # KPI cards
    kpi_cards, _ = make_kpis(df)

    return html.Div(className="content", children=[
        dbc.Row([
            dbc.Col(html.H2("📊 Analyses Avancées", style={
                "color": "#0ea5e9",
                "fontWeight": "800",
                "marginBottom": "20px"
            }), md=12)
        ]),

        # KPIs
        html.Div(kpi_cards),
        html.Br(),

        # ========================================
        # 📈 GRAPHIQUES LIGNE 1
        # ========================================
        dbc.Row([
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("🎯 Lead Time vs Couverture Stock", className="section-title"),
                    dcc.Graph(id="analytics-scatter")
                ])
            ], md=6),

            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("📊 Distribution du Stock", className="section-title"),
                    dcc.Graph(id="analytics-hist-stock")
                ])
            ], md=6)
        ]),
        html.Br(),

        # ========================================
        # 📈 GRAPHIQUES LIGNE 2
        # ========================================
        dbc.Row([
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("📦 Ventes par Catégorie", className="section-title"),
                    dcc.Graph(id="analytics-box-category")
                ])
            ], md=6),

            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("🏭 Top 10 Fournisseurs", className="section-title"),
                    dcc.Graph(id="analytics-purchase-supplier")
                ])
            ], md=6)
        ]),
        html.Br(),

        # ========================================
        # 📈 GRAPHIQUES LIGNE 3
        # ========================================
        dbc.Row([
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("🎯 QAC vs Stock Optimal", className="section-title"),
                    dcc.Graph(id="analytics-qac-optimal")
                ])
            ], md=6),

            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("📈 Répartition Besoins d'Achat", className="section-title"),
                    dcc.Graph(id="analytics-pie-need")
                ])
            ], md=6)
        ]),
        html.Br(),

        # ========================================
        # 📈 GRAPHIQUES LIGNE 4 - NOUVEAUX
        # ========================================
        dbc.Row([
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("⚡ Taux de Rotation par Catégorie", className="section-title"),
                    dcc.Graph(id="analytics-rotation-category")
                ])
            ], md=6),

            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("💰 Valeur Stock par Fournisseur", className="section-title"),
                    dcc.Graph(id="analytics-stock-value")
                ])
            ], md=6)
        ]),

        # ========================================
        # 🔒 MAIN-TABLE CACHÉ (pour callbacks cross-page)
        # ========================================
        html.Div([
            dash_table.DataTable(
                id="main-table",
                data=[],
                columns=[],
                style_table={"display": "none"}
            )
        ], style={"display": "none"})
    ])


# ==========================================
# 📊 CALLBACK ANALYTICS - UTILISE FILTRES SIDEBAR
# ==========================================

@app.callback(
    [Output("analytics-scatter", "figure"),
     Output("analytics-hist-stock", "figure"),
     Output("analytics-box-category", "figure"),
     Output("analytics-purchase-supplier", "figure"),
     Output("analytics-qac-optimal", "figure"),
     Output("analytics-pie-need", "figure"),
     Output("analytics-rotation-category", "figure"),
     Output("analytics-stock-value", "figure")],
    [Input("filter-supplier", "value"),
     Input("filter-category", "value"),
     Input("filter-need", "value"),
     Input("filtered-data", "data")],
    prevent_initial_call=False
)
def update_analytics_all_charts(supplier_filter, category_filter, need_filter, filtered_json):
    """
    Met à jour tous les graphiques analytics selon les filtres sidebar
    """
    print(f"\n📊 Analytics - Filtres: {supplier_filter}, {category_filter}, {need_filter}")

    # Charger données filtrées
    if filtered_json:
        df = pd.DataFrame(json.loads(filtered_json))
    else:
        df = get_df_cached()

    # Colonnes nécessaires
    cols = [
        "product_name", "Supplier", "Product Category",
        "total_stock", "Average Daily Sales",
        "Max Coverage Day", "ADJUSTED_LEADTIME",
        "Ajusted_total_need", "purchase_need",
        "QAC", "optimal stock", "credit_days"
    ]
    df = df[[c for c in cols if c in df.columns]].dropna()

    print(f"   📊 {len(df)} produits pour analytics")

    # Couleurs par besoin
    color_map = {
        "ORDER NOW": "#ef4444",
        "ORDER NOT URGENT": "#f59e0b",
        "NO NEED": "#10b981"
    }

    # ========================================
    # 📊 GRAPHIQUE 1 : SCATTER LEAD TIME VS COVERAGE
    # ========================================
    fig_scatter = px.scatter(
        df,
        x="ADJUSTED_LEADTIME",
        y="Max Coverage Day",
        color="Ajusted_total_need",
        size="purchase_need",
        hover_data=["product_name", "Supplier", "QAC"],
        labels={
            "ADJUSTED_LEADTIME": "Lead Time Ajusté (jours)",
            "Max Coverage Day": "Couverture Stock (jours)"
        },
        color_discrete_map=color_map
    )

    fig_scatter.update_layout(
        title=dict(text="Lead Time vs Couverture Stock", font=dict(size=16, color="#1e293b")),
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(family="-apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif", color="#374151", size=12),
        legend=dict(bgcolor="rgba(255,255,255,0.9)", bordercolor="#e2e8f0", borderwidth=1),
        margin=dict(l=60, r=30, t=60, b=50),
        height=380
    )

    # ========================================
    # 📊 GRAPHIQUE 2 : HISTOGRAMME STOCK
    # ========================================
    fig_hist = px.histogram(
        df,
        x="total_stock",
        nbins=30,
        labels={"total_stock": "Stock Total"},
        color_discrete_sequence=["#0ea5e9"]
    )

    fig_hist.update_layout(
        title=dict(text="Distribution des Niveaux de Stock", font=dict(size=16, color="#1e293b")),
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(family="-apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif", color="#374151", size=12),
        showlegend=False,
        margin=dict(l=60, r=30, t=60, b=50),
        height=380
    )

    # ========================================
    # 📊 GRAPHIQUE 3 : BOXPLOT PAR CATÉGORIE
    # ========================================
    fig_box = px.box(
        df,
        x="Product Category",
        y="Average Daily Sales",
        color="Product Category",
        labels={"Average Daily Sales": "Ventes Moyennes"}
    )

    fig_box.update_layout(
        title=dict(text="Ventes par Catégorie", font=dict(size=16, color="#1e293b")),
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(family="-apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif", color="#374151", size=12),
        showlegend=False,
        margin=dict(l=60, r=30, t=60, b=80),
        height=380
    )

    # ========================================
    # 📊 GRAPHIQUE 4 : TOP 10 FOURNISSEURS
    # ========================================
    purchase_by_supplier = (
        df.groupby("Supplier")["purchase_need"]
        .sum()
        .reset_index()
        .sort_values("purchase_need", ascending=False)
        .head(10)
    )

    fig_supplier = px.bar(
        purchase_by_supplier,
        x="Supplier",
        y="purchase_need",
        labels={"purchase_need": "Besoin d'Achat Total"},
        color="purchase_need",
        color_continuous_scale="Reds"
    )

    fig_supplier.update_layout(
        title=dict(text="Top 10 Fournisseurs - Besoin d'Achat", font=dict(size=16, color="#1e293b")),
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(family="-apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif", color="#374151", size=12),
        xaxis=dict(tickangle=-45, tickfont=dict(size=10)),
        showlegend=False,
        coloraxis_showscale=False,
        margin=dict(l=60, r=30, t=60, b=100),
        height=380
    )

    # ========================================
    # 📊 GRAPHIQUE 5 : QAC VS OPTIMAL
    # ========================================
    fig_qac = px.scatter(
        df,
        x="QAC",
        y="optimal stock",
        color="Ajusted_total_need",
        hover_data=["product_name", "Supplier"],
        labels={"QAC": "Quantité Ajustée Commandée"},
        color_discrete_map=color_map
    )

    fig_qac.update_layout(
        title=dict(text="QAC vs Stock Optimal", font=dict(size=16, color="#1e293b")),
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(family="-apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif", color="#374151", size=12),
        legend=dict(bgcolor="rgba(255,255,255,0.9)", bordercolor="#e2e8f0", borderwidth=1),
        margin=dict(l=60, r=30, t=60, b=50),
        height=380
    )

    # ========================================
    # 📊 GRAPHIQUE 6 : PIE BESOIN D'ACHAT
    # ========================================
    need_dist = df["Ajusted_total_need"].value_counts().reset_index()
    need_dist.columns = ["Besoin", "Nombre"]

    fig_pie = px.pie(
        need_dist,
        values="Nombre",
        names="Besoin",
        color="Besoin",
        color_discrete_map=color_map,
        hole=0.4  # Donut chart
    )

    fig_pie.update_layout(
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(color="#374151", size=11),
        height=400
    )

    # ========================================
    # 📊 GRAPHIQUE 7 : ROTATION PAR CATÉGORIE
    # ========================================
    rotation_by_cat = df.groupby("Product Category").agg({
        "Max Coverage Day": "mean",
        "product_name": "count"
    }).reset_index()
    rotation_by_cat.columns = ["Category", "Rotation Moyenne", "Nombre"]

    fig_rotation = px.bar(
        rotation_by_cat.sort_values("Rotation Moyenne"),
        x="Category",
        y="Rotation Moyenne",
        labels={"Rotation Moyenne": "Jours de Couverture Moyenne"},
        color="Rotation Moyenne",
        color_continuous_scale="RdYlGn_r"
    )

    fig_rotation.update_layout(
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(color="#374151", size=11),
        xaxis=dict(tickangle=-45),
        height=400
    )

    # ========================================
    # 📊 GRAPHIQUE 8 : VALEUR STOCK PAR FOURNISSEUR
    # ========================================
    stock_by_supplier = (
        df.groupby("Supplier")["total_stock"]
        .sum()
        .reset_index()
        .sort_values("total_stock", ascending=False)
        .head(10)
    )

    fig_stock_value = px.bar(
        stock_by_supplier,
        x="Supplier",
        y="total_stock",
        labels={"total_stock": "Stock Total (unités)"},
        color="total_stock",
        color_continuous_scale="Blues"
    )

    fig_stock_value.update_layout(
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(color="#374151", size=11),
        xaxis=dict(tickangle=-45),
        height=400
    )

    print(f"   ✅ Analytics graphiques générés\n")

    return (
        fig_scatter,
        fig_hist,
        fig_box,
        fig_supplier,
        fig_qac,
        fig_pie,
        fig_rotation,
        fig_stock_value
    )


def page_predictive(master_df: pd.DataFrame = None):
    """Page Prédictions avec filtres sidebar et graphiques ML"""
    df = master_df if master_df is not None else get_df_cached()

    # Colonnes ML
    cols_to_keep = [
        "product_id", "product_name", "Supplier",
        "total_stock", "Average Daily Sales",
        "Max Daily Sales (Pikine)", "target_quantity",
        "Product Category", "Ajusted_total_need",
        "total_sold_30d", "avg_daily_sold",
        "std_sold", "max_daily_sold",
        "days_with_sales", "trend_factor",
        "projected_demand", "safety_stock",
        "target_supervised"
    ]
    cols_to_keep = [c for c in cols_to_keep if c in df.columns]
    pred_df = df[cols_to_keep].copy()

    # Calcul stockout rate
    stockout_rate = 0.0
    if 'avg_daily_sold' in pred_df.columns and 'total_stock' in pred_df.columns:
        pred_df_ml = pred_df[pred_df['avg_daily_sold'].notna()].copy()

        if len(pred_df_ml) > 0:
            pred_df_ml['coverage_days'] = np.where(
                pred_df_ml['avg_daily_sold'] > 0,
                pred_df_ml['total_stock'] / pred_df_ml['avg_daily_sold'],
                999
            )
            at_risk = (pred_df_ml['coverage_days'] < 21).sum()
            stockout_rate = (at_risk / len(pred_df_ml)) * 100

    return html.Div(className="content", children=[
        dbc.Row([
            dbc.Col(html.H3("🔮 Prédictions ML", style={
                "color": "#0ea5e9",
                "fontWeight": "800"
            }), md=8),
            dbc.Col(
                dbc.Badge(
                    f"Stockout prédit : {stockout_rate:.1f}%",
                    color="danger" if stockout_rate > 15 else "warning" if stockout_rate > 5 else "success",
                    style={"fontSize": "14px", "padding": "8px 15px"}
                ),
                md=4,
                style={"display": "flex", "justifyContent": "flex-end", "alignItems": "center"}
            )
        ], className="mb-4"),
        html.Br(),

        # ========================================
        # 📈 GRAPHIQUES LIGNE 1
        # ========================================
        dbc.Row([
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("📊 Top 20 - Target Quantity", className="section-title"),
                    dcc.Graph(id="predictive-bar")
                ])
            ], md=12)
        ]),
        html.Br(),

        # ========================================
        # 📈 GRAPHIQUES LIGNE 2
        # ========================================
        dbc.Row([
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("🎯 Demande Projetée vs Stock", className="section-title"),
                    dcc.Graph(id="predictive-scatter")
                ])
            ], md=6),

            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("🛡️ Safety Stock par Catégorie", className="section-title"),
                    dcc.Graph(id="predictive-safety-stock")
                ])
            ], md=6)
        ]),
        html.Br(),

        # ========================================
        # 📈 GRAPHIQUES LIGNE 3 - NOUVEAUX
        # ========================================
        dbc.Row([
            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("📈 Tendances de Ventes", className="section-title"),
                    dcc.Graph(id="predictive-trend")
                ])
            ], md=6),

            dbc.Col([
                html.Div(className="soft-card", children=[
                    html.H5("⚠️ Produits à Risque", className="section-title"),
                    dcc.Graph(id="predictive-risk")
                ])
            ], md=6)
        ]),
        html.Br(),

        # ========================================
        # 📋 TABLEAU DÉTAILLÉ
        # ========================================
        html.Div(className="soft-card", children=[
            html.H5("📋 Données Prédictives Détaillées", className="section-title"),
            dash_table.DataTable(
                id="predictive-table",
                columns=[{"name": c, "id": c} for c in pred_df.columns],
                data=pred_df.to_dict("records"),
                page_size=15,
                filter_action="native",
                sort_action="native",
                sort_mode="multi",
                style_table={"overflowX": "auto"},
                style_header={
                    "backgroundColor": "#1e3a5f",
                    "border": "1px solid #2d3748",
                    "fontWeight": "700",
                    "textAlign": "center",
                    "color": "#1e293b"
                },
                style_cell={
                    "backgroundColor": "#f8fafc",
                    "color": "#1e293b",
                    "border": "1px solid #1f2937",
                    "fontSize": 12,
                    "textAlign": "center"
                },
                style_data_conditional=[
                    {
                        "if": {"column_id": "target_quantity"},
                        "fontWeight": "700",
                        "color": "#0ea5e9",
                        "fontSize": "14px"
                    }
                ]
            )
        ]),

        # ========================================
        # 🔒 MAIN-TABLE CACHÉ (pour callbacks cross-page)
        # ========================================
        html.Div([
            dash_table.DataTable(
                id="main-table",
                data=[],
                columns=[],
                style_table={"display": "none"}
            )
        ], style={"display": "none"})
    ])


# ==========================================
# 🔮 CALLBACK PRÉDICTIONS - UTILISE FILTRES SIDEBAR
# ==========================================

@app.callback(
    [Output("predictive-bar", "figure"),
     Output("predictive-scatter", "figure"),
     Output("predictive-safety-stock", "figure"),
     Output("predictive-trend", "figure"),
     Output("predictive-risk", "figure"),
     Output("predictive-table", "data")],
    [Input("filter-supplier", "value"),
     Input("filter-category", "value"),
     Input("filter-need", "value"),
     Input("filtered-data", "data")],
    prevent_initial_call=False
)
def update_predictive_all_charts(supplier_filter, category_filter, need_filter, filtered_json):
    """
    Met à jour tous les graphiques prédictifs selon les filtres sidebar
    """
    print(f"\n🔮 Prédictions - Filtres: {supplier_filter}, {category_filter}, {need_filter}")

    # Charger données
    if filtered_json:
        df = pd.DataFrame(json.loads(filtered_json))
    else:
        df = get_df_cached()

    # Colonnes ML
    cols = [
        "product_name", "Supplier", "Product Category",
        "total_stock", "Average Daily Sales",
        "target_quantity", "Ajusted_total_need",
        "avg_daily_sold", "projected_demand",
        "safety_stock", "trend_factor",
        "max_daily_sold", "days_with_sales"
    ]
    df = df[[c for c in cols if c in df.columns]].copy()

    print(f"   📊 {len(df)} produits pour prédictions")

    color_map = {
        "ORDER NOW": "#ef4444",
        "ORDER NOT URGENT": "#f59e0b",
        "NO NEED": "#10b981"
    }

    # ========================================
    # 📊 GRAPHIQUE 1 : BAR CHART TOP 20
    # ========================================
    df_chart = df[df['target_quantity'] > 0].copy() if 'target_quantity' in df.columns else df.copy()

    fig_bar = px.bar(
        df_chart.sort_values("target_quantity", ascending=False).head(20),
        x="product_name",
        y="target_quantity",
        color="Ajusted_total_need",
        hover_data=["Supplier", "total_stock"],
        labels={"target_quantity": "Quantité à Commander"},
        color_discrete_map=color_map
    )

    fig_bar.update_layout(
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(color="#374151", size=11),
        xaxis=dict(tickangle=-45),
        height=400
    )

    # ========================================
    # 📊 GRAPHIQUE 2 : SCATTER DEMANDE VS STOCK
    # ========================================
    if 'projected_demand' in df.columns:
        fig_scatter = px.scatter(
            df,
            x="total_stock",
            y="projected_demand",
            color="Ajusted_total_need",
            size="Average Daily Sales",
            hover_data=["product_name", "Supplier"],
            labels={
                "total_stock": "Stock Actuel",
                "projected_demand": "Demande Projetée"
            },
            color_discrete_map=color_map
        )
    else:
        fig_scatter = px.scatter(
            df,
            x="total_stock",
            y="Average Daily Sales",
            color="Ajusted_total_need",
            hover_data=["product_name"],
            color_discrete_map=color_map
        )

    fig_scatter.update_layout(
        title=dict(text="Lead Time vs Couverture Stock", font=dict(size=16, color="#1e293b")),
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(family="-apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif", color="#374151", size=12),
        legend=dict(bgcolor="rgba(255,255,255,0.9)", bordercolor="#e2e8f0", borderwidth=1),
        margin=dict(l=60, r=30, t=60, b=50),
        height=380
    )

    # ========================================
    # 📊 GRAPHIQUE 3 : SAFETY STOCK PAR CATÉGORIE
    # ========================================
    if 'safety_stock' in df.columns:
        safety_by_cat = (
            df.groupby("Product Category")["safety_stock"]
            .sum()
            .reset_index()
            .sort_values("safety_stock", ascending=False)
        )

        fig_safety = px.bar(
            safety_by_cat,
            x="Product Category",
            y="safety_stock",
            labels={"safety_stock": "Safety Stock Total"},
            color="safety_stock",
            color_continuous_scale="Blues"
        )
    else:
        fig_safety = px.bar(
            df.groupby("Product Category").size().reset_index(name="count"),
            x="Product Category",
            y="count"
        )

    fig_safety.update_layout(
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(color="#374151", size=11),
        height=400
    )

    # ========================================
    # 📊 GRAPHIQUE 4 : TENDANCES (TREND FACTOR)
    # ========================================
    if 'trend_factor' in df.columns:
        trend_by_cat = (
            df.groupby("Product Category")["trend_factor"]
            .mean()
            .reset_index()
            .sort_values("trend_factor", ascending=False)
        )

        fig_trend = px.bar(
            trend_by_cat,
            x="Product Category",
            y="trend_factor",
            labels={"trend_factor": "Facteur de Tendance Moyen"},
            color="trend_factor",
            color_continuous_scale="RdYlGn"
        )
    else:
        fig_trend = go.Figure()
        fig_trend.add_annotation(
            text="Données de tendance non disponibles",
            xref="paper", yref="paper",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=14, color="#94a3b8")
        )

    fig_trend.update_layout(
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(color="#374151", size=11),
        height=400
    )

    # ========================================
    # 📊 GRAPHIQUE 5 : PRODUITS À RISQUE
    # ========================================
    if 'avg_daily_sold' in df.columns and 'total_stock' in df.columns:
        df_risk = df[df['avg_daily_sold'].notna()].copy()
        df_risk['coverage'] = np.where(
            df_risk['avg_daily_sold'] > 0,
            df_risk['total_stock'] / df_risk['avg_daily_sold'],
            999
        )
        df_risk['risk_level'] = pd.cut(
            df_risk['coverage'],
            bins=[0, 7, 14, 30, 999],
            labels=["Critique (<7j)", "Attention (7-14j)", "Moyen (14-30j)", "OK (>30j)"]
        )

        risk_dist = df_risk['risk_level'].value_counts().reset_index()
        risk_dist.columns = ["Niveau", "Nombre"]

        fig_risk = px.pie(
            risk_dist,
            values="Nombre",
            names="Niveau",
            color="Niveau",
            color_discrete_map={
                "Critique (<7j)": "#ef4444",
                "Attention (7-14j)": "#f59e0b",
                "Moyen (14-30j)": "#eab308",
                "OK (>30j)": "#10b981"
            },
            hole=0.4
        )
    else:
        fig_risk = go.Figure()

    fig_risk.update_layout(
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#ffffff",
        font=dict(color="#374151", size=11),
        height=400
    )

    print(f"   ✅ Prédictions graphiques générés\n")

    return (
        fig_bar,
        fig_scatter,
        fig_safety,
        fig_trend,
        fig_risk,
        df.to_dict("records")
    )
def page_about():
    return html.Div(className="content", children=[
        html.H2("À propos", className="page-title"),
        html.Div(className="soft-card", children=[
            html.H4(APP_BRAND, style={"color": "#10b981", "marginBottom": "15px"}),
            html.P(
                "Plateforme premium de pilotage Supply Chain : visibilité temps quasi-réel des stocks, "
                "risques de rupture, analyses graphiques, et recommandations d'approvisionnement optimisées par ML.",
                style={"lineHeight": "1.6"}
            ),
            html.Hr(style={"borderColor": "#1f2937"}),
            html.H5("Auteur", style={"color": "#1e293b", "marginTop": "20px"}),
            html.P(
                f"{AUTHOR} — Data Scientist / Ph.D / Supply Chain Analytics",
                style={"marginBottom": "10px"}
            ),
            html.P(
                f"Dashboard construit avec Dash/Plotly, stylé via Bootstrap (thème {THEME.rsplit('.', 1)[-1]}) "
                "et CSS custom. Modèle ML supervisé basé sur Gradient Boosting (R²=0.73).",
                style={"fontSize": "14px", "color": "#9ca3af"}
            ),
        ])
    ])


# ============================================================
# 👤 PAGE MON ESPACE - DASHBOARD AGENT COMPACT
# ============================================================
# Layout compact - Données EXACTES - Pas de débordement

def page_my_space(username: str, master_df: pd.DataFrame = None):
    """Page Mon Espace - Dashboard de l'agent."""
    track_page_view(username, "myspace")

    if master_df is None:
        master_df = initial_df if 'initial_df' in globals() else pd.DataFrame()

    user_info = get_user_info(username) or {}
    user_name = user_info.get("name", username.title())
    user_role = user_info.get("role", "user")

    # Performances depuis Supabase
    perf = get_user_stats_summary(username)
    bc_count = perf.get("bc_count", 0)
    bc_total = perf.get("bc_total", 0)
    logins = perf.get("logins", 0)
    pages = perf.get("pages_viewed", 0)

    # Fournisseurs de l'agent
    agent_suppliers_list = get_agent_suppliers(username, user_role)
    if agent_suppliers_list is None and 'Supplier' in master_df.columns:
        agent_suppliers_list = sorted(master_df['Supplier'].dropna().unique().tolist())

    # Stats par fournisseur
    suppliers_stats = []
    total_skus = 0
    total_ruptures = 0
    total_at_risk = 0

    if not master_df.empty and 'Supplier' in master_df.columns and agent_suppliers_list:
        for supplier in agent_suppliers_list:
            if not supplier or str(supplier).strip() == '':
                continue
            df_sup = master_df[master_df['Supplier'].str.upper().str.strip() == str(supplier).upper().strip()]
            if df_sup.empty:
                continue

            n_products = len(df_sup)
            total_skus += n_products
            oos_count = 0
            at_risk_count = 0

            if 'total_stock' in df_sup.columns:
                stocks = pd.to_numeric(df_sup['total_stock'], errors='coerce').fillna(0)
                oos_count = int((stocks <= 0).sum())
                total_ruptures += oos_count
                if 'Max Coverage Day' in df_sup.columns:
                    cov = pd.to_numeric(df_sup['Max Coverage Day'], errors='coerce').fillna(999)
                    at_risk_count = int(((cov < 7) & (stocks > 0)).sum())
                    total_at_risk += at_risk_count

            oos_pct = (oos_count / n_products * 100) if n_products > 0 else 0
            stock_value = 0
            if 'total_stock' in df_sup.columns and 'Prix Achat TTC' in df_sup.columns:
                stocks = pd.to_numeric(df_sup['total_stock'], errors='coerce').fillna(0)
                prices = pd.to_numeric(df_sup['Prix Achat TTC'], errors='coerce').fillna(0)
                stock_value = (stocks * prices).sum()

            suppliers_stats.append({
                "supplier": supplier, "products": n_products, "oos": oos_count,
                "oos_pct": oos_pct, "at_risk": at_risk_count, "stock_value": stock_value
            })

    suppliers_stats.sort(key=lambda x: x["oos_pct"], reverse=True)
    supplier_options = [{"label": "Tous mes fournisseurs", "value": "all"}]
    supplier_options += [{"label": f"{s['supplier']} ({s['oos_pct']:.0f}% OOS)", "value": s['supplier']} for s in suppliers_stats]

    return html.Div([
        # En-tête
        html.Div([
            html.Div([
                html.H4(user_name, style={"margin": "0", "fontSize": "18px", "color": "#1e293b", "fontWeight": "600"}),
                html.Span(f"{user_role.upper()} | {datetime.now().strftime('%d/%m/%Y %H:%M')}", style={"fontSize": "11px", "color": "#64748b"})
            ]),
            html.Span(f"{len(suppliers_stats)} fournisseurs", style={"fontSize": "11px", "color": "#0ea5e9", "fontWeight": "500"})
        ], style={"display": "flex", "justifyContent": "space-between", "alignItems": "center", "marginBottom": "12px", "paddingBottom": "10px", "borderBottom": "1px solid #e2e8f0"}),

        # Performances (30j)
        html.Div([
            html.Div("Mes Performances (30 jours)", style={"fontSize": "12px", "fontWeight": "600", "marginBottom": "8px", "color": "#374151"}),
            html.Div([
                html.Div([
                    html.Div(f"{bc_count}", style={"fontSize": "20px", "fontWeight": "700", "color": "#3b82f6"}),
                    html.Div("BC Générés", style={"fontSize": "9px", "color": "#6b7280"})
                ], style={"textAlign": "center", "flex": "1", "padding": "10px", "background": "#eff6ff", "borderRadius": "6px"}),
                html.Div([
                    html.Div(f"{bc_total/1e6:.1f}M", style={"fontSize": "20px", "fontWeight": "700", "color": "#22c55e"}),
                    html.Div("FCFA", style={"fontSize": "9px", "color": "#6b7280"})
                ], style={"textAlign": "center", "flex": "1", "padding": "10px", "background": "#f0fdf4", "borderRadius": "6px"}),
                html.Div([
                    html.Div(f"{logins}", style={"fontSize": "20px", "fontWeight": "700", "color": "#8b5cf6"}),
                    html.Div("Connexions", style={"fontSize": "9px", "color": "#6b7280"})
                ], style={"textAlign": "center", "flex": "1", "padding": "10px", "background": "#faf5ff", "borderRadius": "6px"}),
                html.Div([
                    html.Div(f"{pages}", style={"fontSize": "20px", "fontWeight": "700", "color": "#06b6d4"}),
                    html.Div("Pages", style={"fontSize": "9px", "color": "#6b7280"})
                ], style={"textAlign": "center", "flex": "1", "padding": "10px", "background": "#ecfeff", "borderRadius": "6px"}),
            ], style={"display": "flex", "gap": "6px"})
        ], style={"background": "#fff", "padding": "12px", "borderRadius": "8px", "border": "1px solid #e5e7eb", "marginBottom": "12px"}),

        # Résumé Stock
        html.Div([
            html.Div("Résumé Stock", style={"fontSize": "12px", "fontWeight": "600", "marginBottom": "8px", "color": "#374151"}),
            html.Div([
                html.Div([
                    html.Div(f"{total_skus:,}", style={"fontSize": "18px", "fontWeight": "700", "color": "#0ea5e9"}),
                    html.Div("Produits", style={"fontSize": "9px", "color": "#6b7280"})
                ], style={"textAlign": "center", "flex": "1"}),
                html.Div([
                    html.Div(f"{total_ruptures}", style={"fontSize": "18px", "fontWeight": "700", "color": "#ef4444"}),
                    html.Div("Ruptures", style={"fontSize": "9px", "color": "#6b7280"})
                ], style={"textAlign": "center", "flex": "1", "padding": "6px", "background": "#fef2f2", "borderRadius": "6px"}),
                html.Div([
                    html.Div(f"{total_at_risk}", style={"fontSize": "18px", "fontWeight": "700", "color": "#f59e0b"}),
                    html.Div("<7j", style={"fontSize": "9px", "color": "#6b7280"})
                ], style={"textAlign": "center", "flex": "1", "padding": "6px", "background": "#fffbeb", "borderRadius": "6px"}),
                html.Div([
                    html.Div(f"{(total_ruptures/total_skus*100) if total_skus > 0 else 0:.1f}%", style={"fontSize": "18px", "fontWeight": "700", "color": "#dc2626"}),
                    html.Div("OOS", style={"fontSize": "9px", "color": "#6b7280"})
                ], style={"textAlign": "center", "flex": "1", "padding": "6px", "background": "#fef2f2", "borderRadius": "6px"}),
            ], style={"display": "flex", "gap": "6px"})
        ], style={"background": "#fff", "padding": "12px", "borderRadius": "8px", "border": "1px solid #e5e7eb", "marginBottom": "12px"}),

        # Filtre Fournisseur
        html.Div([
            html.Div("Filtrer par Fournisseur", style={"fontSize": "12px", "fontWeight": "600", "marginBottom": "6px", "color": "#374151"}),
            dcc.Dropdown(id="myspace-supplier-filter", options=supplier_options, value="all", clearable=False, style={"fontSize": "12px"})
        ], style={"background": "#fff", "padding": "10px", "borderRadius": "8px", "border": "1px solid #e5e7eb", "marginBottom": "12px"}),

        # Liste Fournisseurs
        html.Div([
            html.Div([
                html.Span("Mes Fournisseurs", style={"fontSize": "12px", "fontWeight": "600", "color": "#374151"}),
                html.Span(f" ({len(suppliers_stats)})", style={"fontSize": "11px", "color": "#64748b"})
            ], style={"marginBottom": "8px"}),

            html.Div([
                # Header
                html.Div([
                    html.Div("Fournisseur", style={"flex": "2", "fontWeight": "500", "fontSize": "10px", "color": "#64748b"}),
                    html.Div("Produits", style={"flex": "1", "fontWeight": "500", "fontSize": "10px", "color": "#64748b", "textAlign": "center"}),
                    html.Div("OOS", style={"flex": "1", "fontWeight": "500", "fontSize": "10px", "color": "#64748b", "textAlign": "center"}),
                    html.Div("% OOS", style={"flex": "1", "fontWeight": "500", "fontSize": "10px", "color": "#64748b", "textAlign": "center"}),
                    html.Div("Risque", style={"flex": "1", "fontWeight": "500", "fontSize": "10px", "color": "#64748b", "textAlign": "center"}),
                    html.Div("Valeur", style={"flex": "1", "fontWeight": "500", "fontSize": "10px", "color": "#64748b", "textAlign": "right"}),
                ], style={"display": "flex", "padding": "6px 10px", "background": "#f8fafc", "borderRadius": "4px", "marginBottom": "4px"}),

                html.Div([_create_supplier_row(s) for s in suppliers_stats[:15]], id="myspace-suppliers-list", style={"maxHeight": "280px", "overflowY": "auto"})
            ]) if suppliers_stats else html.Div("Aucun fournisseur", style={"color": "#94a3b8", "textAlign": "center", "padding": "16px"})
        ], style={"background": "#fff", "padding": "12px", "borderRadius": "8px", "border": "1px solid #e5e7eb", "marginBottom": "12px"}),

        # Actions
        html.Div([
            html.Div("Actions Rapides", style={"fontSize": "12px", "fontWeight": "600", "marginBottom": "8px", "color": "#374151"}),
            html.Div([
                dcc.Link(dbc.Button("Overview", color="primary", size="sm", outline=True, style={"fontSize": "11px"}), href="/"),
                dcc.Link(dbc.Button("Analytics", color="info", size="sm", outline=True, style={"fontSize": "11px"}), href="/analytics"),
                dcc.Link(dbc.Button("Prédictions", color="warning", size="sm", outline=True, style={"fontSize": "11px"}), href="/predictions"),
            ], style={"display": "flex", "gap": "6px"})
        ], style={"background": "#f8fafc", "padding": "12px", "borderRadius": "8px", "border": "1px solid #e5e7eb"}),

    ], style={"padding": "12px"})


def _create_supplier_row(s: dict) -> html.Div:
    """Crée une ligne pour un fournisseur"""
    oos_pct = s["oos_pct"]
    if oos_pct >= 30:
        oos_color, oos_bg = "#dc2626", "#fef2f2"
    elif oos_pct >= 15:
        oos_color, oos_bg = "#f59e0b", "#fffbeb"
    elif oos_pct > 0:
        oos_color, oos_bg = "#eab308", "#fefce8"
    else:
        oos_color, oos_bg = "#22c55e", "#f0fdf4"

    return html.Div([
        html.Div(s["supplier"][:22], style={"flex": "2", "fontSize": "11px", "color": "#1e293b", "overflow": "hidden", "textOverflow": "ellipsis", "whiteSpace": "nowrap"}),
        html.Div(str(s["products"]), style={"flex": "1", "fontSize": "11px", "color": "#64748b", "textAlign": "center"}),
        html.Div(str(s["oos"]), style={"flex": "1", "fontSize": "11px", "color": "#ef4444", "textAlign": "center", "fontWeight": "600"}),
        html.Div(f"{oos_pct:.0f}%", style={"flex": "1", "fontSize": "11px", "color": oos_color, "textAlign": "center", "fontWeight": "600", "background": oos_bg, "borderRadius": "3px", "padding": "1px 0"}),
        html.Div(str(s["at_risk"]), style={"flex": "1", "fontSize": "11px", "color": "#f59e0b", "textAlign": "center"}),
        html.Div(f"{s['stock_value']/1e6:.1f}M", style={"flex": "1", "fontSize": "11px", "color": "#22c55e", "textAlign": "right"}),
    ], style={"display": "flex", "alignItems": "center", "padding": "6px 10px", "borderBottom": "1px solid #f1f5f9"})


def _create_orders_history(orders: list) -> html.Div:
    """Crée l'historique des commandes"""
    if not orders:
        return html.P("Aucune commande enregistrée", style={"color": "#64748b", "textAlign": "center"})

    items = []
    for order in orders[-10:]:  # 10 dernières
        date_str = order.get("created_at", "")[:10]
        supplier = order.get("supplier", "N/A")
        amount = order.get("order_total", 0)
        products = order.get("products_count", 0)
        po = order.get("po_number", "N/A")

        items.append(html.Div([
            html.Div([
                html.Span(f"📄 {po}", style={"fontWeight": "600", "color": "#1e293b"}),
                html.Span(f" • {date_str}", style={"color": "#64748b", "fontSize": "13px"})
            ]),
            html.Div([
                html.Span(supplier, style={"color": "#0ea5e9", "fontWeight": "500"}),
                html.Span(f" • {products} produits • ", style={"color": "#64748b"}),
                html.Span(f"{amount:,.0f} FCFA", style={"color": "#22c55e", "fontWeight": "600"})
            ], style={"marginTop": "4px", "fontSize": "13px"})
        ], style={
            "padding": "12px",
            "background": "#f8fafc",
            "borderRadius": "8px",
            "marginBottom": "8px",
            "borderLeft": "4px solid #0ea5e9"
        }))

    return html.Div(items)


def page_promotions():
    """Page dédiée à l'analyse des promotions."""
    df = get_df_cached()

    if df is None or df.empty:
        return html.Div("Aucune donnée disponible", className="error-message")

    # ✅ COLONNES AUTORISÉES
    allowed_columns_promo = [
        'product_name',
        'promo_status',
        'days_remaining',
        'uplift_pct',
        'roi_pct',
        'promo_recommendation',
        'promo_priority',
        'avg_daily_sold',
        'net_profit_per_day',
        'discount_pct'
    ]

    # ✅ COLONNES INTERDITES
    cols_banned_in_promo = [
        'Average Daily Sales', 'Max Daily Sales (Pikine)', 'optimal stock',
        'Ajusted_total_need', 'QAC', 'target_quantity', 'Max Coverage Day',
        'credit_days', 'Credit_cumulable', 'AJUSTER_BUFFER', 'MAX_CREDIT_BUFFER',
        'ADJUSTED_LEADTIME', 'MOQ MAAD', 'delisting_status', 'Daily OOS Rate (30d)',
        'total_stock', 'product_id', 'Supplier', 'Product Category',
        'sales_with_promo', 'sales_without_promo'
    ]

    # Filtrage
    available_cols = [c for c in allowed_columns_promo if c in df.columns]
    available_cols = [c for c in available_cols if c not in cols_banned_in_promo]
    df_promo = df[available_cols].copy()

    if 'uplift_pct' in df_promo.columns:
        df_promo = df_promo[df_promo['uplift_pct'].notna()].copy()

    print(f"\n Page Promotions : {len(df_promo)} produits, {len(available_cols)} colonnes")

    # Vérification
    if df_promo.empty or 'uplift_pct' not in df_promo.columns:
        return html.Div(className="content", children=[
            html.H2(" Gestion des Promotions", className="page-title"),
            dbc.Alert("Les données de promotions ne sont pas encore chargées.", color="warning")
        ])

    # KPIs
    active_promos = (df_promo['promo_status'] == 'Active').sum() if 'promo_status' in df_promo.columns else 0
    avg_roi = df_promo['roi_pct'].mean() if 'roi_pct' in df_promo.columns else 0
    profitable_promos = (df_promo['roi_pct'] > 0).sum() if 'roi_pct' in df_promo.columns else 0
    high_priority = (df_promo['promo_priority'] == 'Haute').sum() if 'promo_priority' in df_promo.columns else 0

    kpi_cards = dbc.Row([
        dbc.Col(html.Div(className="kpi", children=[
            html.Small(" Promos actives"),
            html.H3(f"{active_promos}")
        ]), md=3),
        dbc.Col(html.Div(className="kpi", children=[
            html.Small(" ROI moyen"),
            html.H3(f"{avg_roi:.1f}%", style={
                "color": "#10b981" if avg_roi > 50 else "#f59e0b" if avg_roi > 0 else "#ef4444"
            })
        ]), md=3),
        dbc.Col(html.Div(className="kpi", children=[
            html.Small(" Promos rentables"),
            html.H3(f"{profitable_promos}")
        ]), md=3),
        dbc.Col(html.Div(className="kpi", children=[
            html.Small(" Haute priorité"),
            html.H3(f"{high_priority}")
        ]), md=3),
    ], className="mb-4")

    # Graphique ROI
    if 'roi_pct' in df_promo.columns and len(df_promo) > 0:
        fig_roi = px.bar(
            df_promo.nlargest(20, 'roi_pct'),
            x='product_name',
            y='roi_pct',
            color='promo_priority' if 'promo_priority' in df_promo.columns else None,
            title="Top 20 produits par ROI",
            labels={'roi_pct': 'ROI (%)', 'product_name': 'Produit'},
            color_discrete_map={'Haute': '#10b981', 'Moyenne': '#f59e0b', 'Faible': '#6b7280', 'Éviter': '#ef4444'}
        )
        fig_roi.update_layout(
            plot_bgcolor="#f8fafc",
            paper_bgcolor="#ffffff",
            font=dict(color="#374151"),
            xaxis=dict(tickangle=-45),
            height=400
        )
    else:
        fig_roi = {'data': [],
                   'layout': {'title': 'Pas de données ROI', 'plot_bgcolor': '#f8fafc', 'paper_bgcolor': '#ffffff',
                              'font': {'color': '#1e293b'}}}

    # Graphique Uplift
    if 'uplift_pct' in df_promo.columns and len(df_promo) > 0:
        fig_uplift = px.bar(
            df_promo.nlargest(20, 'uplift_pct'),
            x='product_name',
            y='uplift_pct',
            title="Top 20 produits par Uplift",
            labels={'uplift_pct': 'Uplift (%)', 'product_name': 'Produit'}
        )
        fig_uplift.update_layout(
            plot_bgcolor="#f8fafc",
            paper_bgcolor="#ffffff",
            font=dict(color="#374151"),
            xaxis=dict(tickangle=-45),
            height=400
        )
    else:
        fig_uplift = {'data': [], 'layout': {'title': 'Pas de données Uplift', 'plot_bgcolor': '#f8fafc',
                                             'paper_bgcolor': '#ffffff', 'font': {'color': '#1e293b'}}}

    # Layout
    return html.Div(className="content", children=[
        html.H2(" Gestion des Promotions", className="page-title"),
        kpi_cards,

        dbc.Row([
            dbc.Col(html.Div(dcc.Graph(figure=fig_roi), className="soft-card"), md=6),
            dbc.Col(html.Div(dcc.Graph(figure=fig_uplift), className="soft-card"), md=6)
        ], className="mb-4"),

        html.Div(className="soft-card", children=[
            html.H5("Détails par produit", className="section-title"),
            dash_table.DataTable(
                id="promo-table",
                columns=[{"name": c, "id": c} for c in available_cols],
                data=df_promo.to_dict("records"),
                page_size=20,
                filter_action="native",
                sort_action="native",
                style_table={"overflowX": "auto"},
                style_header={
                    "backgroundColor": "#1e3a5f",
                    "color": "white",
                    "fontWeight": "bold",
                    "textAlign": "center"
                },
                style_cell={
                    "backgroundColor": "#f8fafc",
                    "color": "#1e293b",
                    "textAlign": "center",
                    "padding": "8px",
                    "fontSize": "12px"
                },
                style_data_conditional=[
                    {"if": {"filter_query": "{promo_status} = 'Active'"}, "backgroundColor": "rgba(16,185,129,.2)",
                     "fontWeight": "bold"},
                    {"if": {"column_id": "roi_pct"}, "fontWeight": "bold", "color": "#10b981"},
                    {"if": {"filter_query": "{promo_priority} = 'Éviter'"}, "backgroundColor": "rgba(239,68,68,.15)",
                     "color": "#fca5a5"},
                    {"if": {"filter_query": "{promo_priority} = 'Haute'"}, "backgroundColor": "rgba(16,185,129,.15)",
                     "color": "#86efac"}
                ]
            )
        ]),

        # ========================================
        # 🔒 MAIN-TABLE CACHÉ (pour callbacks cross-page)
        # ========================================
        html.Div([
            dash_table.DataTable(
                id="main-table",
                data=[],
                columns=[],
                style_table={"display": "none"}
            )
        ], style={"display": "none"})
    ])


# ============================================================
# 👤 PAGE AGENTS - SUIVI PERFORMANCE
# ============================================================
# URL du CSV des agents-fournisseurs
AGENTS_SUPPLIERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQAK0IcIDJS8ysyCB0wnLp-rR-t-zu_2_6bYV4-YIhPuL3fZQyo7fgMXZnJ4rcz-5mNur_UHgMenRiU/pub?gid=2137463594&single=true&output=csv"

# Cache global pour les données agents
_agents_cache = {"data": None, "timestamp": None}


def load_agents_suppliers():
    """Charge le mapping agents → fournisseurs depuis Google Sheets"""
    import time

    # Vérifier le cache (5 minutes)
    if _agents_cache["data"] is not None and _agents_cache["timestamp"]:
        if time.time() - _agents_cache["timestamp"] < 300:
            return _agents_cache["data"]

    try:
        df = pd.read_csv(AGENTS_SUPPLIERS_URL)
        print(f"✅ Agents-Fournisseurs chargés: {len(df)} lignes")
        print(f"   Colonnes: {df.columns.tolist()}")

        # Normaliser les noms de colonnes
        df.columns = df.columns.str.strip()

        # Mettre en cache
        _agents_cache["data"] = df
        _agents_cache["timestamp"] = time.time()

        return df
    except Exception as e:
        print(f"⚠️ Erreur chargement agents: {e}")
        return pd.DataFrame()


def calculate_agent_performance(master_df: pd.DataFrame, agents_df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcule les KPIs de performance pour chaque agent

    Colonnes du CSV agents:
    - Owner → L'agent
    - Liste des Fournisseurs → Le fournisseur géré
    - Total Order Value → Valeur des commandes
    - # Refs OOS → Nombre de références en rupture
    - Traitement → Statut (Pending, Traitee, No Need)
    """

    if agents_df.empty:
        print("⚠️ DataFrame agents vide")
        return pd.DataFrame()

    # Identifier les colonnes clés basées sur la structure réelle
    owner_col = 'Owner' if 'Owner' in agents_df.columns else None
    supplier_col = 'Liste des Fournisseurs' if 'Liste des Fournisseurs' in agents_df.columns else None
    order_value_col = 'Total Order Value' if 'Total Order Value' in agents_df.columns else None
    refs_oos_col = '# Refs OOS' if '# Refs OOS' in agents_df.columns else None
    traitement_col = 'Traitement' if 'Traitement' in agents_df.columns else None

    if not owner_col:
        # Chercher des alternatives
        for col in agents_df.columns:
            col_lower = col.lower()
            if 'owner' in col_lower or 'agent' in col_lower:
                owner_col = col
                break

    if not supplier_col:
        for col in agents_df.columns:
            col_lower = col.lower()
            if 'fournisseur' in col_lower or 'supplier' in col_lower:
                supplier_col = col
                break

    if not owner_col or not supplier_col:
        print(f"⚠️ Colonnes owner/supplier non trouvées dans: {agents_df.columns.tolist()}")
        return pd.DataFrame()

    print(f"   📊 Colonnes identifiées: Owner={owner_col}, Supplier={supplier_col}")

    # Nettoyer les données
    agents_df = agents_df.copy()
    agents_df[owner_col] = agents_df[owner_col].astype(str).str.strip()
    agents_df[supplier_col] = agents_df[supplier_col].astype(str).str.strip()

    # Convertir les valeurs numériques
    if order_value_col and order_value_col in agents_df.columns:
        agents_df[order_value_col] = pd.to_numeric(
            agents_df[order_value_col].astype(str).str.replace(',', '').str.replace(' ', ''),
            errors='coerce'
        ).fillna(0)

    if refs_oos_col and refs_oos_col in agents_df.columns:
        agents_df[refs_oos_col] = pd.to_numeric(agents_df[refs_oos_col], errors='coerce').fillna(0)

    # Liste des agents uniques
    agents_list = agents_df[owner_col].dropna().unique()
    agents_list = [a for a in agents_list if a and a != 'nan']

    print(f"   👥 Agents trouvés: {agents_list}")

    performance_data = []

    for agent_name in agents_list:
        # Filtrer les lignes de cet agent
        agent_rows = agents_df[agents_df[owner_col] == agent_name]

        # Fournisseurs gérés par cet agent
        suppliers = agent_rows[supplier_col].dropna().unique().tolist()
        suppliers = [s for s in suppliers if s and s != 'nan']

        # Calculer les KPIs depuis le CSV agents
        total_order_value = agent_rows[order_value_col].sum() if order_value_col else 0
        total_refs_oos = int(agent_rows[refs_oos_col].sum()) if refs_oos_col else 0

        # Comptage des statuts de traitement
        commandes_traitees = 0
        commandes_pending = 0
        if traitement_col:
            commandes_traitees = (agent_rows[traitement_col].astype(str).str.lower() == 'traitee').sum()
            commandes_pending = (agent_rows[traitement_col].astype(str).str.lower() == 'pending').sum()

        # Calculer les KPIs depuis les données produits (master_df)
        total_skus = 0
        total_stock = 0
        rupture_count = 0
        risk_count = 0
        avg_rotation = 0

        # Chercher la bonne colonne supplier dans master_df
        supplier_col_master = None
        for col in ['Supplier', 'supplier_name', 'supplier']:
            if col in master_df.columns:
                supplier_col_master = col
                break

        if not master_df.empty and supplier_col_master:
            # Normaliser pour la comparaison
            suppliers_normalized = [s.lower().strip() for s in suppliers]

            # Filtrer les produits de cet agent
            agent_products = master_df[
                master_df[supplier_col_master].astype(str).str.lower().str.strip().isin(suppliers_normalized)
            ].copy()

            total_skus = len(agent_products)

            if 'total_stock' in agent_products.columns:
                agent_products['total_stock'] = pd.to_numeric(agent_products['total_stock'], errors='coerce').fillna(0)
                total_stock = int(agent_products['total_stock'].sum())
                rupture_count = int((agent_products['total_stock'] <= 0).sum())

            # Produits à risque (couverture < 7 jours)
            if 'rotation_days' in agent_products.columns:
                agent_products['rotation_days'] = pd.to_numeric(agent_products['rotation_days'], errors='coerce').fillna(999)
                risk_count = int(((agent_products['rotation_days'] < 7) & (agent_products['rotation_days'] > 0)).sum())
                valid_rotations = agent_products[agent_products['rotation_days'] < 999]['rotation_days']
                avg_rotation = valid_rotations.mean() if len(valid_rotations) > 0 else 0
            elif 'Average Daily Sales' in agent_products.columns and 'total_stock' in agent_products.columns:
                ads = pd.to_numeric(agent_products['Average Daily Sales'], errors='coerce').fillna(0.1).replace(0, 0.1)
                stock = pd.to_numeric(agent_products['total_stock'], errors='coerce').fillna(0)
                rotations = stock / ads
                risk_count = int(((rotations < 7) & (rotations > 0)).sum())
                avg_rotation = rotations.mean() if len(rotations) > 0 else 0

        # Taux de rupture
        rupture_rate = (rupture_count / total_skus * 100) if total_skus > 0 else 0

        # Taux de traitement
        total_commandes = len(agent_rows)
        taux_traitement = (commandes_traitees / total_commandes * 100) if total_commandes > 0 else 0

        # Score global (0-100)
        # Basé sur: taux de rupture (30%), taux traitement (30%), refs OOS (20%), rotation (20%)
        score_rupture = max(0, 100 - rupture_rate * 2)
        score_traitement = taux_traitement
        score_oos = max(0, 100 - total_refs_oos * 2) if total_refs_oos < 50 else 0
        score_rotation = min(100, (avg_rotation / 30 * 100)) if pd.notna(avg_rotation) and avg_rotation > 0 else 50

        score_global = (
                score_rupture * 0.30 +
                score_traitement * 0.30 +
                score_oos * 0.20 +
                score_rotation * 0.20
        )

        performance_data.append({
            'agent_name': agent_name,
            'total_suppliers': len(suppliers),
            'suppliers_list': ', '.join(suppliers[:3]) + ('...' if len(suppliers) > 3 else ''),
            'total_order_value': int(total_order_value),
            'total_commandes': total_commandes,
            'commandes_traitees': commandes_traitees,
            'commandes_pending': commandes_pending,
            'taux_traitement': round(taux_traitement, 1),
            'total_skus': total_skus,
            'total_stock': total_stock,
            'rupture_count': rupture_count,
            'rupture_rate': round(rupture_rate, 1),
            'refs_oos': total_refs_oos,
            'risk_count': risk_count,
            'avg_rotation_days': round(avg_rotation, 1) if pd.notna(avg_rotation) else 0,
            'score_global': round(score_global, 1)
        })

    result_df = pd.DataFrame(performance_data)
    if not result_df.empty:
        result_df = result_df.sort_values('score_global', ascending=False).reset_index(drop=True)

    return result_df


def get_agent_suppliers_detail(agent_name: str, agents_df: pd.DataFrame) -> pd.DataFrame:
    """Retourne le détail des fournisseurs d'un agent"""
    if agents_df.empty:
        return pd.DataFrame()

    owner_col = 'Owner' if 'Owner' in agents_df.columns else None
    if not owner_col:
        for col in agents_df.columns:
            if 'owner' in col.lower() or 'agent' in col.lower():
                owner_col = col
                break

    if not owner_col:
        return pd.DataFrame()

    return agents_df[agents_df[owner_col].astype(str).str.strip() == agent_name].copy()


def get_supplier_agent_mapping():
    """Crée un mapping fournisseur → agent depuis les données"""
    agents_df = load_agents_suppliers()
    mapping = {}

    if agents_df.empty:
        return mapping

    owner_col = 'Owner' if 'Owner' in agents_df.columns else None
    supplier_col = 'Liste des Fournisseurs' if 'Liste des Fournisseurs' in agents_df.columns else None

    if not owner_col:
        for col in agents_df.columns:
            if 'owner' in col.lower() or 'agent' in col.lower():
                owner_col = col
                break

    if not supplier_col:
        for col in agents_df.columns:
            if 'fournisseur' in col.lower() or 'supplier' in col.lower():
                supplier_col = col
                break

    if owner_col and supplier_col:
        for _, row in agents_df.iterrows():
            supplier = str(row[supplier_col]).strip().lower()
            agent = str(row[owner_col]).strip()
            if supplier and supplier != 'nan' and agent and agent != 'nan':
                mapping[supplier] = agent

    return mapping


def get_agent_for_product(product_supplier, mapping=None):
    """Retourne l'agent responsable d'un produit basé sur son fournisseur"""
    if mapping is None:
        mapping = get_supplier_agent_mapping()

    if not product_supplier:
        return "Non assigné"

    supplier_lower = str(product_supplier).strip().lower()
    return mapping.get(supplier_lower, "Non assigné")


def page_agents(master_df: pd.DataFrame = None):
    """Page de suivi de performance des agents - Compacte et adaptée au layout"""

    print("\n" + "="*60)
    print("👤 CHARGEMENT PAGE AGENTS")
    print("="*60)

    # Charger les données
    agents_df = load_agents_suppliers()

    if master_df is None:
        master_df = initial_df if 'initial_df' in globals() else pd.DataFrame()

    performance_df = calculate_agent_performance(master_df, agents_df)
    supplier_agent_map = get_supplier_agent_mapping()

    print(f"   📊 {len(performance_df)} agents avec performances calculées")

    # ==========================================
    # CALCULS KPIs - UTILISER FONCTION COMMUNE
    # ==========================================
    total_agents = len(performance_df) if not performance_df.empty else 0

    # Score moyen pondéré
    avg_score = 0
    if not performance_df.empty and 'score_global' in performance_df.columns:
        scores = performance_df['score_global'].dropna()
        avg_score = scores.mean() if len(scores) > 0 else 0

    # Valeur totale commandes (depuis agents_df, pas performance_df)
    total_order_value = 0
    if not performance_df.empty and 'total_order_value' in performance_df.columns:
        total_order_value = performance_df['total_order_value'].sum()

    # Taux traitement moyen
    avg_taux_traitement = 0
    if not performance_df.empty and 'taux_traitement' in performance_df.columns:
        taux = performance_df['taux_traitement'].dropna()
        avg_taux_traitement = taux.mean() if len(taux) > 0 else 0

    # ==========================================
    # ✅ UTILISER LA FONCTION COMMUNE - COHÉRENCE GARANTIE
    # ==========================================
    print("   🔄 Page Agents: Utilisation de calculate_stock_kpis()")
    stock_kpis = calculate_stock_kpis(master_df)

    total_ruptures = stock_kpis['out_of_stock']
    total_at_risk = stock_kpis['at_risk']

    print(f"   📈 KPIs Agents: agents={total_agents}, score={avg_score:.1f}, rupt={total_ruptures}, risk={total_at_risk}")

    # ==========================================
    # KPIs AGRANDIS (style cards)
    # ==========================================
    kpi_card_style = {
        "display": "flex", "flexDirection": "column", "alignItems": "center", "justifyContent": "center",
        "background": "#ffffff", "borderRadius": "12px",
        "padding": "16px 20px", "minWidth": "130px", "border": "1px solid #e2e8f0",
        "boxShadow": "0 2px 8px rgba(0,0,0,0.06)"
    }
    kpi_value_style = {"fontSize": "28px", "fontWeight": "700", "lineHeight": "1.2"}
    kpi_label_style = {"color": "#64748b", "fontSize": "11px", "marginTop": "6px", "textTransform": "uppercase", "letterSpacing": "0.5px"}

    kpis_row = html.Div([
        html.Div([
            html.Div(f"{total_agents}", style={**kpi_value_style, "color": "#0ea5e9"}),
            html.Div(" Agents", style=kpi_label_style)
        ], style=kpi_card_style),
        html.Div([
            html.Div(f"{avg_score:.0f}", style={**kpi_value_style, "color": "#34d399" if avg_score >= 60 else "#f87171"}),
            html.Div(" Score Moyen", style=kpi_label_style)
        ], style=kpi_card_style),
        html.Div([
            html.Div(f"{total_order_value/1e6:.1f}M", style={**kpi_value_style, "color": "#7c3aed"}),
            html.Div(" Valeur Cmd", style=kpi_label_style)
        ], style=kpi_card_style),
        html.Div([
            html.Div(f"{avg_taux_traitement:.0f}%", style={**kpi_value_style, "color": "#34d399"}),
            html.Div(" Traitement", style=kpi_label_style)
        ], style=kpi_card_style),
        html.Div([
            html.Div(f"{total_ruptures}", style={**kpi_value_style, "color": "#f87171"}),
            html.Div(" Ruptures", style=kpi_label_style)
        ], style=kpi_card_style),
        html.Div([
            html.Div(f"{total_at_risk}", style={**kpi_value_style, "color": "#fbbf24"}),
            html.Div(" À Risque", style=kpi_label_style)
        ], style=kpi_card_style),
    ], style={"display": "flex", "flexWrap": "wrap", "gap": "12px", "marginBottom": "20px"})

    # ==========================================
    # PODIUM TOP 3 (agrandi)
    # ==========================================
    podium = html.Div()
    if not performance_df.empty and len(performance_df) >= 1:
        medals = ["🥇", "🥈", "🥉"]
        colors = ["#fbbf24", "#94a3b8", "#b45309"]
        items = []
        for i in range(min(3, len(performance_df))):
            a = performance_df.iloc[i]
            items.append(html.Span([
                html.Span(medals[i], style={"fontSize": "20px"}), " ",
                html.B(a['agent_name'], style={"color": colors[i], "fontSize": "15px"}),
                html.Span(f" ({a['score_global']:.0f})", style={"color": "#64748b", "fontSize": "14px"})
            ], style={"marginRight": "24px"}))
        podium = html.Div(["🏆 TOP 3 : ", *items], style={
            "background": "linear-gradient(135deg, rgba(30, 41, 59, 0.6) 0%, rgba(15, 23, 42, 0.8) 100%)",
            "borderRadius": "10px", "border": "1px solid #334155",
            "padding": "14px 18px", "marginBottom": "20px", "fontSize": "15px"
        })

    # ==========================================
    # TABLE PERFORMANCE AGENTS (ROW 1 - PLEINE LARGEUR)
    # ==========================================
    table_data = []
    if not performance_df.empty:
        for idx, row in performance_df.iterrows():
            table_data.append({
                'rank': idx + 1,
                'agent': row['agent_name'],
                'frs': int(row['total_suppliers']),
                'val': f"{row['total_order_value']/1e6:.1f}M",
                'ok': int(row['commandes_traitees']),
                'pend': int(row['commandes_pending']),
                'pct': f"{row['taux_traitement']:.0f}%",
                'sku': int(row['total_skus']),
                'rupt': int(row['rupture_count']),
                'sc': round(row['score_global'], 1)
            })

    agents_table = dash_table.DataTable(
        id='agents-performance-table',
        columns=[
            {"name": "#", "id": "rank"},
            {"name": "Agent", "id": "agent"},
            {"name": "Fournisseurs", "id": "frs"},
            {"name": "Valeur Cmd", "id": "val"},
            {"name": "Traitées", "id": "ok"},
            {"name": "En attente", "id": "pend"},
            {"name": "Taux %", "id": "pct"},
            {"name": "SKUs", "id": "sku"},
            {"name": "Ruptures", "id": "rupt"},
            {"name": "Score", "id": "sc"},
        ],
        data=table_data,
        style_table={'overflowX': 'auto'},
        style_header={
            'backgroundColor': '#f1f5f9', 'color': '#64748b', 'fontWeight': '600',
            'fontSize': '12px', 'border': 'none', 'padding': '12px 8px', 'textAlign': 'center'
        },
        style_cell={
            'backgroundColor': '#ffffff', 'color': '#1e293b', 'border': 'none',
            'padding': '10px 8px', 'fontSize': '13px', 'textAlign': 'center', 'minWidth': '70px'
        },
        style_data_conditional=[
            {'if': {'filter_query': '{rank} = 1'}, 'borderLeft': '3px solid #fbbf24', 'backgroundColor': 'rgba(251, 191, 36, 0.05)'},
            {'if': {'filter_query': '{rank} = 2'}, 'borderLeft': '3px solid #94a3b8'},
            {'if': {'filter_query': '{rank} = 3'}, 'borderLeft': '3px solid #b45309'},
            {'if': {'column_id': 'sc'}, 'fontWeight': '700', 'color': '#0ea5e9', 'fontSize': '14px'},
            {'if': {'column_id': 'agent'}, 'fontWeight': '600', 'textAlign': 'left'},
            {'if': {'filter_query': '{pend} > 0', 'column_id': 'pend'}, 'color': '#fbbf24', 'fontWeight': '600'},
            {'if': {'filter_query': '{rupt} > 0', 'column_id': 'rupt'}, 'color': '#f87171', 'fontWeight': '600'},
        ],
        row_selectable='single', selected_rows=[], page_size=8, sort_action='native'
    )

    # ROW 1 : Table Performance
    performance_row = dbc.Row([
        dbc.Col([
            html.Div([
                html.Span("📋", style={"fontSize": "18px", "marginRight": "8px"}),
                html.Span("Performance des Agents", style={"fontSize": "16px", "fontWeight": "600", "color": "#1e293b"})
            ], style={"marginBottom": "12px"}),
            agents_table
        ], width=12)
    ], style={"marginBottom": "24px"})

    # ==========================================
    # TABLE PRODUITS À RISQUE (ROW 2 - PLEINE LARGEUR)
    # ==========================================
    risk_products = []
    if not master_df.empty and supplier_agent_map:
        risk_df = master_df.copy()

        # Colonnes de calcul
        if 'total_stock' in risk_df.columns:
            risk_df['_stk'] = pd.to_numeric(risk_df['total_stock'], errors='coerce').fillna(0)
        else:
            risk_df['_stk'] = 0

        if 'Average Daily Sales' in risk_df.columns:
            risk_df['_ads'] = pd.to_numeric(risk_df['Average Daily Sales'], errors='coerce').fillna(0.01).replace(0, 0.01)
        else:
            risk_df['_ads'] = 0.01

        risk_df['_coverage'] = risk_df['_stk'] / risk_df['_ads']

        # Filtrer : stock <= 0 OU couverture < 7 jours
        risk_mask = (risk_df['_stk'] <= 0) | ((risk_df['_coverage'] < 7) & (risk_df['_coverage'] > 0))
        risk_filtered = risk_df[risk_mask].sort_values('_coverage', ascending=True)

        for _, row in risk_filtered.head(30).iterrows():
            sup = str(row.get('Supplier', '')).strip()
            stk = row['_stk']
            cov = row['_coverage']

            if stk <= 0:
                status = '🔴 Rupture'
                status_color = '#f87171'
            elif cov < 3:
                status = '🟠 Critique'
                status_color = '#fb923c'
            else:
                status = '🟡 Risque'
                status_color = '#fbbf24'

            risk_products.append({
                'status': status,
                'prod': str(row.get('product_name', ''))[:40],
                'sup': sup[:20],
                'agt': get_agent_for_product(sup, supplier_agent_map)[:15],
                'stk': int(stk),
                'cov': f"{cov:.0f}j" if cov < 999 else "-"
            })

    risk_table = dash_table.DataTable(
        id='agents-risk-products-table',
        columns=[
            {"name": "Status", "id": "status"},
            {"name": "Produit", "id": "prod"},
            {"name": "Fournisseur", "id": "sup"},
            {"name": "Agent", "id": "agt"},
            {"name": "Stock", "id": "stk"},
            {"name": "Couv.", "id": "cov"},
        ],
        data=risk_products,
        style_table={'overflowX': 'auto'},
        style_header={
            'backgroundColor': '#f1f5f9', 'color': '#64748b', 'fontWeight': '600',
            'fontSize': '12px', 'border': 'none', 'padding': '12px 8px', 'textAlign': 'left'
        },
        style_cell={
            'backgroundColor': '#ffffff', 'color': '#1e293b', 'border': 'none',
            'padding': '10px 8px', 'fontSize': '13px', 'textAlign': 'left',
            'maxWidth': '200px', 'overflow': 'hidden', 'textOverflow': 'ellipsis'
        },
        style_data_conditional=[
            {'if': {'filter_query': '{status} contains "Rupture"'}, 'backgroundColor': 'rgba(239, 68, 68, 0.1)'},
            {'if': {'filter_query': '{status} contains "Critique"'}, 'backgroundColor': 'rgba(251, 146, 60, 0.08)'},
            {'if': {'column_id': 'agt'}, 'fontWeight': '600', 'color': '#0ea5e9'},
            {'if': {'column_id': 'status'}, 'fontWeight': '600'},
            {'if': {'filter_query': '{stk} = 0', 'column_id': 'stk'}, 'color': '#f87171', 'fontWeight': '700'},
        ],
        page_size=10, sort_action='native'
    )

    # ROW 2 : Table Risques
    risk_row = dbc.Row([
        dbc.Col([
            html.Div([
                html.Span("🚨", style={"fontSize": "18px", "marginRight": "8px"}),
                html.Span(f"Produits à Risque ({len(risk_products)})", style={"fontSize": "16px", "fontWeight": "600", "color": "#1e293b"})
            ], style={"marginBottom": "12px"}),
            risk_table
        ], width=12)
    ], style={"marginBottom": "20px"})

    # ==========================================
    # DÉTAIL FOURNISSEURS (agrandi)
    # ==========================================
    detail = html.Div([
        html.Span("🔍 ", style={"marginRight": "8px", "fontSize": "16px"}),
        html.Div(id="agent-suppliers-detail", children=[
            html.Span("Cliquez sur un agent dans la table pour voir ses fournisseurs", style={"color": "#64748b", "fontSize": "13px"})
        ], style={"display": "inline"})
    ], style={
        "background": "linear-gradient(135deg, rgba(30, 41, 59, 0.4) 0%, rgba(15, 23, 42, 0.6) 100%)",
        "borderRadius": "10px", "border": "1px solid #334155",
        "padding": "12px 16px", "fontSize": "13px"
    })

    # ==========================================
    # RETURN LAYOUT FINAL
    # ==========================================
    return html.Div(className="content", children=[
        # Header
        html.Div([
            html.H4(" Performance Agents", style={"margin": "0", "fontSize": "20px", "fontWeight": "700", "color": "#0ea5e9"}),
            html.P("Suivi des performances et gestion des risques par agent", style={"margin": "4px 0 0 0", "color": "#64748b", "fontSize": "13px"})
        ], style={"marginBottom": "20px"}),

        # KPIs
        kpis_row,

        # Podium
        podium,

        # ROW 1 : Table Performance (pleine largeur)
        performance_row,

        # ROW 2 : Table Risques (pleine largeur)
        risk_row,

        # Détail fournisseurs
        detail,

        # ========================================
        # 🔒 MAIN-TABLE CACHÉ (pour callbacks cross-page)
        # ========================================
        html.Div([
            dash_table.DataTable(
                id="main-table",
                data=[],
                columns=[],
                style_table={"display": "none"}
            )
        ], style={"display": "none"})
    ])


# ==========================================
# CALLBACKS PAGE AGENTS
# ==========================================

@app.callback(
    Output("agent-suppliers-detail", "children"),
    [Input("agents-performance-table", "selected_rows")],
    [State("agents-performance-table", "data")],
    prevent_initial_call=True
)
def show_agent_suppliers_detail(selected_rows, table_data):
    """Affiche les fournisseurs détaillés quand un agent est sélectionné"""
    if not selected_rows or not table_data:
        return html.Span("Cliquez sur un agent pour voir ses fournisseurs", style={"color": "#64748b", "fontSize": "10px"})

    # Récupérer l'agent sélectionné (colonne renommée en 'agent')
    agent_name = table_data[selected_rows[0]].get('agent', table_data[selected_rows[0]].get('agent_name', ''))

    if not agent_name:
        return html.Span("Agent non trouvé", style={"color": "#f87171", "fontSize": "10px"})

    # Charger les données agents
    agents_df = load_agents_suppliers()

    if agents_df.empty:
        return html.Span("Données non disponibles", style={"color": "#f87171", "fontSize": "10px"})

    # Identifier les colonnes
    owner_col = 'Owner' if 'Owner' in agents_df.columns else None
    supplier_col = 'Liste des Fournisseurs' if 'Liste des Fournisseurs' in agents_df.columns else None

    if not owner_col:
        for col in agents_df.columns:
            if 'owner' in col.lower() or 'agent' in col.lower():
                owner_col = col
                break

    if not supplier_col:
        for col in agents_df.columns:
            if 'fournisseur' in col.lower() or 'supplier' in col.lower():
                supplier_col = col
                break

    if not owner_col or not supplier_col:
        return html.Span("Colonnes non trouvées", style={"color": "#f87171", "fontSize": "10px"})

    # Filtrer par agent
    agent_data = agents_df[agents_df[owner_col].astype(str).str.strip() == agent_name]

    if agent_data.empty:
        return html.Span(f"Aucun fournisseur pour {agent_name}", style={"color": "#fbbf24", "fontSize": "10px"})

    # Liste des fournisseurs
    suppliers_list = agent_data[supplier_col].dropna().unique().tolist()
    suppliers_str = ", ".join([s[:15] for s in suppliers_list[:8]])
    if len(suppliers_list) > 8:
        suppliers_str += f"... (+{len(suppliers_list)-8})"

    return html.Span([
        html.B(agent_name, style={"color": "#0ea5e9"}),
        f" : {suppliers_str}"
    ], style={"fontSize": "10px"})


@app.callback(
    Output("agents-risk-products-table", "data"),
    [Input("filter-agent", "value"),
     Input("filter-supplier", "value"),
     Input("filter-category", "value")],
    prevent_initial_call=True
)
def filter_risk_products_by_agent(agent_filter, supplier_filter, category_filter):
    """Filtre les produits à risque par agent et autres filtres sidebar"""

    # Charger les données
    master_df = initial_df if 'initial_df' in globals() else pd.DataFrame()

    if master_df.empty:
        return []

    # Créer le mapping
    supplier_agent_map = get_supplier_agent_mapping()

    # Filtrer les données
    filtered_df = master_df.copy()

    # Appliquer filtre supplier
    if supplier_filter and supplier_filter not in [[], None]:
        if isinstance(supplier_filter, list):
            filtered_df = filtered_df[filtered_df['Supplier'].isin(supplier_filter)]
        else:
            filtered_df = filtered_df[filtered_df['Supplier'] == supplier_filter]

    # Appliquer filtre catégorie
    if category_filter and category_filter not in [[], None, "Toutes"]:
        if isinstance(category_filter, list):
            filtered_df = filtered_df[filtered_df['Product Category'].isin(category_filter)]
        else:
            filtered_df = filtered_df[filtered_df['Product Category'] == category_filter]

    # Identifier produits à risque - Correction: utiliser la colonne si elle existe
    if 'rotation_days' in filtered_df.columns:
        filtered_df['rotation_days'] = pd.to_numeric(filtered_df['rotation_days'], errors='coerce').fillna(999)
    else:
        filtered_df['rotation_days'] = 999

    if 'total_stock' in filtered_df.columns:
        filtered_df['total_stock'] = pd.to_numeric(filtered_df['total_stock'], errors='coerce').fillna(0)
    else:
        filtered_df['total_stock'] = 0

    risk_mask = (filtered_df['total_stock'] <= 0) | (filtered_df['rotation_days'] < 7)
    risk_df = filtered_df[risk_mask]

    # Ajouter l'agent - colonnes: st, prod, sup, agt, stk
    risk_products = []
    for _, row in risk_df.iterrows():
        supplier = str(row.get('Supplier', '')).strip()
        agent = get_agent_for_product(supplier, supplier_agent_map)

        # Filtre par agent si sélectionné
        if agent_filter and agent_filter != "all" and agent != agent_filter:
            continue

        stock_val = row.get('total_stock', 0)
        risk_products.append({
            'st': '🔴' if stock_val <= 0 else '🟡',
            'prod': str(row.get('product_name', 'N/A'))[:25],
            'sup': supplier[:10] if supplier else 'N/A',
            'agt': agent[:10],
            'stk': int(stock_val),
        })

    return risk_products[:100]


# =========================
# ROUTING CALLBACK
# =========================
'''@app.callback(
    Output("page-content", "children"),
    Input("url", "pathname")
)
def display_page(pathname):
    """Route vers les différentes pages selon l'URL."""
    print(f"🔀 Navigation vers : {pathname}")  # Debug

    if pathname == "/":
        return page_overview()
    elif pathname == "/predictive":
        return page_predictive()
    elif pathname == "/promotions":  # ✅ AJOUTER
        return page_promotions()
    elif pathname == "/about":
        return page_about()
    else:
        return html.Div([
            html.H2("404"),
            html.P("Page introuvable"),
            html.A("Retour à l'accueil", href="/")
        ], className="error-message")

'''
# ------------------------------ Chatbot helpers ----------------------------------

# =============================== Chatbot helpers ================================
def _detect_lang(text: str) -> str:
    if not text:
        return "fr"
    t = text.lower()
    fr_markers = ["bonjour", "salut", "stock", "commande", "fournisseur", "rupture", "couverture", "jours", "crédit",
                  "delisting"]
    if any(w in t for w in fr_markers) or re.search(r"[àâçéèêëîïôùûüœ]", t):
        return "fr"
    return "en"


def _clean_df_for_advice(df: pd.DataFrame) -> pd.DataFrame:
    """Prépare une vue compacte (ne modifie pas la logique métier)."""
    if not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()

    d = pd.DataFrame()
    d["product_name_display"] = df.get("product_name", "")
    d["supplier_name"] = df.get("Supplier", "")
    d["abc_class"] = df.get("Product Category", "")
    d["xyz_class"] = ""
    d["current_stock"] = pd.to_numeric(df.get("total_stock", 0), errors="coerce").fillna(0).clip(lower=0)
    d["avg_daily_sales"] = pd.to_numeric(df.get("Average Daily Sales", 0), errors="coerce").fillna(0).clip(lower=0)

    # Coverage days
    with np.errstate(divide='ignore', invalid='ignore'):
        cov_calc = np.where(d["avg_daily_sales"] > 0, d["current_stock"] / d["avg_daily_sales"], 0)

    if "Max Coverage Day" in df.columns:
        d["coverage_days"] = pd.to_numeric(df["Max Coverage Day"], errors="coerce").fillna(0)
    else:
        d["coverage_days"] = cov_calc
    d["coverage_days"] = pd.Series(d["coverage_days"], dtype=float).clip(lower=0, upper=365)

    # Lead time
    if "ADJUSTED_LEADTIME" in df.columns:
        d["leadtime_days"] = pd.to_numeric(df["ADJUSTED_LEADTIME"], errors="coerce").fillna(7)
    elif "Max Lead Time" in df.columns:
        d["leadtime_days"] = pd.to_numeric(df["Max Lead Time"], errors="coerce").fillna(7)
    else:
        d["leadtime_days"] = 7
    d["leadtime_days"] = pd.Series(d["leadtime_days"], dtype=float).clip(lower=0, upper=180)

    # Crédit
    if "credit_days" in df.columns:
        d["credit_days"] = pd.to_numeric(df["credit_days"], errors="coerce").fillna(14)
    else:
        d["credit_days"] = 14

    # Rupture ML (utilise Predicted Stockout si présent, sinon statut stock)
    if "Predicted Stockout" in df.columns:
        d["rupture_ml"] = np.where(df["Predicted Stockout"], "OUI", "NON")
    else:
        st = df.get("Stock Status", "").astype(str)
        d["rupture_ml"] = np.where(st.isin(["Out of Stock", "Predicted Stockout Soon"]), "OUI", "NON")

    d["delisting_product"] = "NON"

    # Exclure lignes non pertinentes
    bad = d["product_name_display"].astype(str).str.contains(
        r'\b(CFA|CASH|CFA\s*-\s*CASH|ESPECES|CAISSE)\b', case=False, na=False
    )
    d = d[~bad].copy()

    d["risk"] = (d["rupture_ml"].astype(str).str.upper() == "OUI").astype(int)
    return d


def _bubble(role: str, text: str):
    is_user = (role == "user")
    return html.Div(className=f"chat-bubble {'user' if is_user else 'bot'}", children=[
        html.Div("🧑" if is_user else "🤖", className=f"chat-avatar {'user' if is_user else ''}"),
        html.Div(dcc.Markdown(text or "", link_target="_blank",
                              style={"whiteSpace": "pre-wrap", "wordBreak": "break-word"}), className="chat-msg")
    ])


def _render_messages(msgs: list):
    return [_bubble(m.get("role", "assistant"), m.get("text", "")) for m in (msgs or [])]


# --- Wrapper simple vers Gemini (utilise l'instance gemini_client déjà créée) ---
def ask_ai(prompt: str, stream: bool = False) -> str:
    try:
        return gemini_client.generate(prompt, stream=stream)
    except Exception as e:
        print(f"[Gemini] Erreur: {type(e).__name__}: {e}")
        return ""


def chatbot_reply(user_text: str, df: pd.DataFrame, history_messages: list) -> str:
    """
    🚀 VERSION OPTIMISÉE - Assistant Supply Chain Expert
    - Réponses rapides et pertinentes
    - Personnalité chaleureuse et experte
    - Analyse contextuelle intelligente
    """
    try:
        if not user_text or not str(user_text).strip():
            return "👋 Salut ! Je suis ton assistant Supply Chain. Pose-moi des questions sur tes stocks, ruptures, commandes ou n'importe quel sujet business !"

        # Détection langue rapide
        lang = _detect_lang(user_text)

        # Préparation données (optimisé - seulement si nécessaire)
        needs_data = any(kw in user_text.lower() for kw in [
            'analyse', 'recommande', 'conseil', 'rupture', 'commande', 'stock',
            'produit', 'quels', 'combien', 'urgent', 'priorité', 'fournisseur',
            'risque', 'order', 'achat', 'besoin', 'coverage', 'lead', 'qac',
            'commander', 'approvisionnement', 'inventaire', 'ruptures', 'données'
        ])

        ctx = ""
        if needs_data and isinstance(df, pd.DataFrame) and not df.empty:
            # Extraction rapide des KPIs clés
            try:
                total_products = len(df)

                # Ruptures
                ruptures = 0
                if 'total_stock' in df.columns:
                    stocks = pd.to_numeric(df['total_stock'], errors='coerce').fillna(0)
                    ruptures = int((stocks <= 0).sum())

                # Couverture médiane
                cov_med = None
                if 'coverage_days' in df.columns:
                    cov_med = df['coverage_days'].median()
                elif 'Average Daily Sales' in df.columns and 'total_stock' in df.columns:
                    ads = pd.to_numeric(df['Average Daily Sales'], errors='coerce').fillna(0.01).replace(0, 0.01)
                    stock = pd.to_numeric(df['total_stock'], errors='coerce').fillna(0)
                    cov_med = (stock / ads).median()

                # Fournisseurs
                suppliers = df['Supplier'].nunique() if 'Supplier' in df.columns else 0

                # Produits à risque (couverture < 7j)
                at_risk = 0
                if cov_med is not None:
                    at_risk = int(((stock / ads) < 7).sum())

                # Top 5 produits critiques
                top_critical = []
                if 'product_name' in df.columns and 'total_stock' in df.columns:
                    df_sorted = df.copy()
                    df_sorted['_stock'] = pd.to_numeric(df_sorted['total_stock'], errors='coerce').fillna(0)
                    critical = df_sorted[df_sorted['_stock'] <= 0].head(5)
                    for _, row in critical.iterrows():
                        top_critical.append(f"• {row.get('product_name', 'N/A')[:40]} ({row.get('Supplier', 'N/A')})")

                ctx = f"""
📊 **DONNÉES ACTUELLES** (temps réel):
- Total SKUs: {total_products}
- Ruptures: {ruptures} produits (stock = 0)
- À risque: {at_risk} produits (couverture < 7j)
- Couverture médiane: {cov_med:.1f}j
- Fournisseurs actifs: {suppliers}

🔴 **TOP PRODUITS EN RUPTURE**:
{chr(10).join(top_critical) if top_critical else '• Aucune rupture détectée ✅'}
"""
            except Exception as e:
                ctx = f"⚠️ Données partiellement disponibles. Erreur: {str(e)[:50]}"

        # Système expert avec personnalité
        if lang == "fr":
            system_msg = """Tu es **MAAD Assistant**, expert Supply Chain avec 20+ ans d'expérience chez les plus grands distributeurs africains.

🎯 **TA PERSONNALITÉ**:
- Chaleureux, dynamique et passionné par la supply chain
- Tu tutoies l'utilisateur et utilises des emojis avec modération
- Tu es direct, pragmatique et orienté solutions
- Tu aimes partager ton expertise avec enthousiasme

💼 **TES COMPÉTENCES**:
- Gestion des stocks et approvisionnements
- Analyse ABC/XYZ, prévision de la demande
- Négociation fournisseurs, lead times, crédits
- Optimisation des coûts et du BFR
- Prévention des ruptures et surstock

📋 **COMMENT TU RÉPONDS**:
- Questions simples → Réponses courtes et percutantes
- Demandes d'analyse → Données chiffrées + recommandations concrètes
- Conversations → Échanges naturels et engageants
- Toujours proposer la prochaine étape ou action"""

        else:
            system_msg = """You are **MAAD Assistant**, a Supply Chain expert with 20+ years of experience with top African distributors.

🎯 **YOUR PERSONALITY**:
- Warm, dynamic and passionate about supply chain
- You use a friendly tone and emojis sparingly
- Direct, pragmatic and solution-oriented
- Love sharing expertise with enthusiasm

💼 **YOUR SKILLS**:
- Inventory management and procurement
- ABC/XYZ analysis, demand forecasting
- Supplier negotiation, lead times, credits
- Cost and working capital optimization
- Stockout and overstock prevention

📋 **HOW YOU RESPOND**:
- Simple questions → Short and impactful answers
- Analysis requests → Hard data + concrete recommendations
- Conversations → Natural and engaging exchanges
- Always suggest the next step or action"""

        # Historique compact (5 derniers pour rapidité)
        hist_lines = []
        for m in (history_messages or [])[-5:]:
            role = "👤" if m.get("role") == "user" else "🤖"
            hist_lines.append(f"{role} {m.get('text', '').strip()[:100]}")
        history_txt = "\n".join(hist_lines) if hist_lines else "Début de conversation"

        # Prompt optimisé (plus court = plus rapide)
        prompt_text = f"""{system_msg}

{ctx if ctx else '💬 Conversation générale - pas de données demandées'}

📜 **Historique récent**:
{history_txt}

❓ **Question**: {user_text.strip()}

👉 Réponds en {'français' if lang == 'fr' else 'anglais'}, de façon concise et actionnable."""

        # Appel Gemini
        out = ask_ai(prompt_text).strip()
        return out if out else _chatbot_fallback(user_text, df)

    except Exception as e:
        print(f"[Chatbot] Erreur: {e}")
        return _chatbot_fallback(user_text, df)
        import traceback
        print(f"[Chatbot] Erreur _chatbot_reply: {traceback.format_exc()}")
        return f"⚠️ Erreur technique : {type(e).__name__}"


def _chatbot_fallback(user_text: str, df: pd.DataFrame, prefix: str = "") -> str:
    """Réponse de secours (aucun appel LLM)."""
    lang = _detect_lang(user_text)

    if df.empty:
        return prefix + ("Données insuffisantes. Recharge les sources." if lang == "fr"
                         else "Insufficient data. Please reload sources.")

    try:
        top_risk = df.sort_values(['risk', 'coverage_days'], ascending=[False, True]).head(5) \
            if 'risk' in df.columns else df.head(5)
        cnt_risk = int(df['risk'].sum()) if 'risk' in df.columns else 0
        cov_med = float(df['coverage_days'].median()) if 'coverage_days' in df.columns else 0

        lines = [f"{prefix}" + (f"Risque: {cnt_risk} prod. • Couverture médiane ~ {cov_med:.1f} j." if lang == "fr"
                                else f"Risk: {cnt_risk} SKUs • Median coverage ~ {cov_med:.1f}d.")]

        for _, r in top_risk.iterrows():
            if lang == "fr":
                lines.append(
                    f"- {r.get('product_name_display', '?')} · cov {float(r.get('coverage_days', 0)):.1f}j · "
                    f"LT {int(r.get('leadtime_days', 0))}j · crédit {int(r.get('credit_days', 0))}j"
                )
            else:
                lines.append(
                    f"- {r.get('product_name_display', '?')} · cov {float(r.get('coverage_days', 0)):.1f}d · "
                    f"LT {int(r.get('leadtime_days', 0))}d · credit {int(r.get('credit_days', 0))}d"
                )

        lines += [("Cibles: A/AX en Y/Z <7j; crédit>14j si LT>20j; promos classe C." if lang == "fr"
                   else "Focus A/A+ in Y/Z <7d; credit>14d if LT>20d; promo bundles for class C.")]
        return "\n".join(lines)

    except Exception as e:
        print(f"[Chatbot] Erreur fallback: {e}")
        return prefix + ("Recommandation impossible avec les données disponibles." if lang == "fr"
                         else "Unable to compute recommendation with available data.")


# ===============================================================================

# def _chatbot_fallback(user_text: str, df: pd.DataFrame, prefix: str = "") -> str:
#   """
#  df ici est déjà le DataFrame nettoyé (view), pas besoin de re-nettoyer
# """
# lang = _detect_lang(user_text)

# ✅ NE PAS appeler _clean_df_for_advice ici, df est déjà nettoyé
# if df.empty:
#   return prefix + (
#      "Données insuffisantes. Recharge les sources." if lang == "fr" else "Insufficient data. Please reload sources.")

# try:
# Utiliser directement df (qui est view)
#   top_risk = df.sort_values(['risk', 'coverage_days'], ascending=[False, True]).head(
#      5) if 'risk' in df.columns else df.head(5)
# cnt_risk = int(df['risk'].sum()) if 'risk' in df.columns else 0
# cov_med = float(df['coverage_days'].median()) if 'coverage_days' in df.columns else 0

# lines = [f"{prefix}" + (f"Risque: {cnt_risk} prod. • Couverture médiane ~ {cov_med:.1f} j." if lang == "fr"
#                       else f"Risk: {cnt_risk} SKUs • Median coverage ~ {cov_med:.1f}d.")]

# for _, r in top_risk.iterrows():
#   if lang == "fr":
#      lines.append(
#         f"- {r.get('product_name_display', '?')} · cov {float(r.get('coverage_days', 0)):.1f}j · LT {int(r.get('leadtime_days', 0))}j · crédit {int(r.get('credit_days', 0))}j")
# else:
#    lines.append(
#       f"- {r.get('product_name_display', '?')} · cov {float(r.get('coverage_days', 0)):.1f}d · LT {int(r.get('leadtime_days', 0))}d · credit {int(r.get('credit_days', 0))}d")

# lines += [("Cibles: A/AX en Y/Z <7j; crédit>14j si LT>20j; promos classe C." if lang == "fr"
#          else "Focus A/A+ in Y/Z <7d; credit>14d if LT>20d; promo bundles for class C.")]
# return "\n".join(lines)

# except Exception as e:
#   print(f"[Chatbot] Erreur fallback: {e}")
#  return prefix + (
#     "Recommandation impossible avec les données disponibles." if lang == "fr" else "Unable to compute recommendation with available data.")
# ------------------------------ Layout root (with floating chat) ------------------
try:
    initial_df = get_df_cached()
    if initial_df is None or initial_df.empty:
        raise ValueError("DataFrame vide ou None retourné par get_df_cached()")
    print(f"✅ Application initialisée avec {len(initial_df)} produits")
except Exception as e:
    print(f"❌ ERREUR CRITIQUE lors du chargement initial : {e}")
    import traceback

    traceback.print_exc()
    # Créer un DataFrame vide par sécurité
    initial_df = pd.DataFrame(columns=['product_name', 'Supplier', 'total_stock'])

initial_df = get_df_cached()
initial_df['QAC edited']=' '

# ✅ GARANTIR product_id dans le DataFrame initial
if 'product_id' in initial_df.columns:
    initial_df['product_id'] = pd.to_numeric(initial_df['product_id'], errors='coerce').fillna(0).astype(int)
    print(f"✅ product_id initial: {(initial_df['product_id'] > 0).sum()} valides sur {len(initial_df)}")
else:
    initial_df['product_id'] = 0
    print("⚠️ product_id absent du DataFrame initial - créé avec 0")

# ✅ Mettre à jour les stats journalières dans Supabase
try:
    update_daily_stats_on_load(initial_df)
except Exception as e:
    print(f"⚠️ Erreur update_daily_stats_on_load: {e}")

#initial_df['delete']='delete'
# ==================== LAYOUT CORRIGÉ AVEC AUTHENTIFICATION ====================

app.layout = html.Div([
    # ========== STORES AUTHENTIFICATION ==========
    dcc.Store(id="auth-state", storage_type="session", data={"authenticated": False, "username": None}),
    dcc.Store(id="user-info-store", storage_type="session"),

    # ========== NAVIGATION & STORES ==========
    dcc.Location(id="url", refresh=False),
    dcc.Store(id="master-data", data=initial_df.to_json(orient="records")),
    dcc.Store(id="filtered-data"),
    dcc.Store(id="uploaded-csv"),
    dcc.Store(id="chat-store", data=[]),
    dcc.Store(id="chat-open", data=False),
    dcc.Store(id='selected-product-for-notes', data=None),  # ✅ Simple ID (pas de pattern-matching ici)
    dcc.Store(id="qac-edits-store", storage_type='local', data={}),  # ✅ Un seul Store pour QAC
    dcc.Store(id="edit-mode"),
    dcc.Store(id="edit-original-product"),
    dcc.Store(id="adding-new-product-flag", data=False),  # ✅ Flag pour distinguer ajout vs édition
    # ========== STORES (données partagées entre callbacks) ==========
    dcc.Store(id='agent-ia-recommendations', data=None),
    dcc.Store(id='bc-data-store', data=None),

    # ========== CONTENEUR PRINCIPAL (login ou dashboard) ==========
    html.Div(id="app-container"),

    # ========== FEEDBACKS GLOBAUX (niveau racine pour callbacks) ==========
    html.Div(id="action-feedback", style={"position": "fixed", "top": "80px", "right": "20px", "zIndex": 10000}),
    html.Div(id="selection-counter", style={"position": "fixed", "bottom": "20px", "right": "20px", "zIndex": 9999}),

    # ========== COMPOSANTS CACHÉS (pour callbacks qui ciblent des éléments dynamiques) ==========
    html.Div(id='edit-product-output', style={"display": "none"}),

    # Table principale cachée (sera remplacée par page_overview)
    html.Div(style={"display": "none"}, children=[
        dash_table.DataTable(
            id="main-table",
            columns=[{"name": "ID", "id": "product_id"}],
            data=[],
            row_selectable="multi",
            selected_rows=[]
        ),
        html.Span(id="selection-count-inline"),
        dbc.Button(id="btn-refresh"),
        dbc.Button(id="btn-reload-data"),
        dbc.Button(id="btn-add-row"),
        dbc.Button(id="btn-po-pdf"),
        dbc.Button(id="btn-run-agent-ia"),
        dbc.Button(id="btn-select-all"),
        dbc.Button(id="btn-clear-selection"),
        dbc.Button(id="btn-fill-qac-target"),
        dbc.Button(id="btn-save-qac"),
        dcc.Dropdown(id="rotation-period"),
        dbc.Alert(id="alert-ia-analysis", is_open=False),
        dbc.Alert(id="alert-bc-error", is_open=False),
        html.Div(id="qac-save-feedback"),
    ]),


    # ========== EDIT MODAL (VISIBLE AU NIVEAU RACINE) ==========
    dbc.Modal(
        id="edit-modal",
        is_open=False,
        backdrop="static",
        children=[
            dbc.ModalHeader(dbc.ModalTitle("Éditer produit", id="edit-modal-title")),
            dbc.ModalBody([
                # Nom du produit
                html.Div([
                    html.Label("Nom du produit", style={"fontWeight": "600", "marginBottom": "5px"}),
                    dbc.Input(id="edit-product-name", placeholder="Nom du produit", type="text"),
                ], style={"marginBottom": "15px"}),

                # Fournisseur
                html.Div([
                    html.Label("Fournisseur", style={"fontWeight": "600", "marginBottom": "5px"}),
                    dbc.Input(id="edit-supplier", placeholder="Fournisseur", type="text"),
                ], style={"marginBottom": "15px"}),

                # Catégorie
                html.Div([
                    html.Label("Catégorie", style={"fontWeight": "600", "marginBottom": "5px"}),
                    dbc.Input(id="edit-category", placeholder="Catégorie ABC/XYZ", type="text"),
                ], style={"marginBottom": "15px"}),

                # Stock
                html.Div([
                    html.Label("Stock actuel", style={"fontWeight": "600", "marginBottom": "5px"}),
                    dbc.Input(id="edit-stock", placeholder="Stock", type="number", min=0),
                ], style={"marginBottom": "15px"}),

                # QAC
                html.Div([
                    html.Label("QAC (Quantité à commander)", style={"fontWeight": "600", "marginBottom": "5px"}),
                    dbc.Input(id="edit-qac", placeholder="QAC", type="number", min=0),
                ], style={"marginBottom": "15px"}),

                # Average Daily Sales
                html.Div([
                    html.Label("Average Daily Sales", style={"fontWeight": "600", "marginBottom": "5px"}),
                    dbc.Input(id="edit-ads", placeholder="Ventes journalières moyennes", type="number", min=0,
                              step=0.1),
                ], style={"marginBottom": "15px"}),

                # Crédit
                html.Div([
                    html.Label("Jours de crédit", style={"fontWeight": "600", "marginBottom": "5px"}),
                    dbc.Input(id="edit-credit", placeholder="Crédit (jours)", type="number", min=0),
                ], style={"marginBottom": "15px"}),

                # Lead Time
                html.Div([
                    html.Label("Lead Time ajusté", style={"fontWeight": "600", "marginBottom": "5px"}),
                    dbc.Input(id="edit-leadtime", placeholder="Lead time (jours)", type="number", min=0),
                ], style={"marginBottom": "15px"}),

                # Buffer
                html.Div([
                    html.Label("Buffer ajusté", style={"fontWeight": "600", "marginBottom": "5px"}),
                    dbc.Input(id="edit-buffer", placeholder="Buffer (jours)", type="number", min=0),
                ], style={"marginBottom": "15px"}),

                # Max Daily Sales
                html.Div([
                    html.Label("Max Daily Sales", style={"fontWeight": "600", "marginBottom": "5px"}),
                    dbc.Input(id="edit-max-sales", placeholder="Ventes max journalières", type="number", min=0),
                ], style={"marginBottom": "15px"}),

                # Feedback
                html.Div(id="edit-modal-feedback", style={"color": "#10b981", "marginTop": "10px", "fontSize": "13px"})
            ]),
            dbc.ModalFooter([
                dbc.Button("Annuler", id="edit-modal-close", outline=True, size="sm", style={"marginRight": "8px"}),
                dbc.Button("💾 Enregistrer", id="edit-modal-save", color="primary", size="sm", n_clicks=0),
            ])
        ]
    ),

    # ========== PATTERN-MATCHING STORES (pour localStorage) ==========
    dcc.Store(id={'type': 'storage', 'index': 'memory'}),
    dcc.Store(id={'type': 'storage', 'index': 'local'}, storage_type='local'),

    # ========== BOUTONS CACHÉS (pour callbacks localStorage) ==========
    html.Div(style={"display": "none"}, children=[
        html.Button('Stockage Local', id={'type': 'button-storage', 'index': 'local'}),
        html.Button('Stockage Session', id={'type': 'button-storage', 'index': 'session'}),
        html.Div([html.Span(0, id={'type': 'output-storage', 'index': 'local'}), " Clics"]),
    ]),

    # ========== SIDEBAR ==========
    # Note: La sidebar est maintenant gérée par le callback render_page_content
    # make_sidebar(),

    # ========== CONTENU PRINCIPAL ==========
    # Note: page-container est maintenant géré par le callback d'authentification (render_page_content)
    # html.Div(id="page-container", children=page_overview(initial_df)),

    # Output pour Agent IA (invisible)
    html.Div(id="agent-ia-output", style={"display": "none"}),

    # NOTE: action-feedback et selection-counter sont maintenant au niveau racine (après app-container)

    # ========== CHATBOT FLOTTANT ==========
    html.Button(id="chat-fab", className="chat-fab", children=[html.Span("Assistant"), html.Span("💬")]),
    html.Div(id="chat-window", className="chat-window", style={"display": "none"}, children=[
        html.Div(className="chat-header", children=[
            html.Strong("Assistant "),
            html.Div([dbc.Button("Minimiser", id="chat-close", size="sm", className="btn-primary")])
        ]),
        html.Div(id="chat-messages", className="chat-body"),
        html.Div(style={"padding": "10px"}, children=[
            dcc.Upload(
                id="chat-upload",
                children=html.Div(["📤 Glisser-déposer un CSV ici ou ", html.B("cliquer pour sélectionner")]),
                multiple=False,
                className="upload-box"
            ),
            html.Small(id="upload-status", style={"color": "#6b7280"})
        ]),
        html.Div(className="chat-input-wrap", children=[
            dbc.Textarea(
                id="chat-input",
                className="chat-textarea",
                placeholder="Écrire une question… (Enter = envoyer, Shift+Enter = nouvelle ligne)",
                rows=2
            ),
            dbc.Button("Envoyer", id="chat-send", className="btn-primary", n_clicks=0)
        ])
    ]),
])


# ============================================================
# 🔐 CALLBACKS D'AUTHENTIFICATION
# ============================================================

@app.callback(
    Output("app-container", "children"),
    [Input("auth-state", "data"),
     Input("url", "pathname")],
    prevent_initial_call=False
)
def render_page_content(auth_state, pathname):
    """
    Callback principal qui gère l'affichage:
    - Si non authentifié → page de connexion
    - Si authentifié → dashboard avec barre utilisateur
    """
    print(f"🔐 render_page_content: auth_state={auth_state}, pathname={pathname}")

    # Vérifier si l'utilisateur est authentifié
    if not auth_state or not auth_state.get("authenticated"):
        print("   → Affichage page de connexion")
        return create_login_layout()

    # Utilisateur connecté - afficher le dashboard
    username = auth_state.get("username", "")
    print(f"   → Utilisateur connecté: {username}")

    # Créer la barre utilisateur
    user_navbar = create_user_navbar(username)

    # Créer la sidebar
    sidebar = make_sidebar()

    # Déterminer la page
    page_name = "overview"  # Par défaut

    # Sélectionner la page appropriée selon l'URL
    # Récupérer le rôle de l'utilisateur
    user_info = get_user_info(username) or {}
    user_role = user_info.get("role", "user")

    if pathname is None or pathname == "/" or pathname == "/overview":
        page_content = page_overview(initial_df, username, user_role)
        page_name = "overview"
    elif pathname == "/myspace":
        page_content = page_my_space(username, initial_df)
        page_name = "myspace"
    elif pathname == "/analytics":
        page_content = page_analytics(initial_df)
        page_name = "analytics"
    elif pathname == "/predictions":
        page_content = page_predictive(initial_df)
        page_name = "predictions"
    elif pathname == "/promotions":
        page_content = page_promotions()
        page_name = "promotions"
    elif pathname == "/about":
        page_content = page_about()
        page_name = "about"
    else:
        # Page par défaut
        page_content = page_overview(initial_df)
        page_name = "overview"

    # ✅ TRACKING SUPABASE: Page view
    if supabase_client and username:
        try:
            if username in ACTIVE_SESSIONS:
                session_info = ACTIVE_SESSIONS[username]
                track_activity(
                    session_info.get("user_id"),
                    session_info.get("session_id"),
                    "page_view",
                    page=page_name,
                    details={"pathname": pathname}
                )
        except Exception as e:
            print(f"⚠️ Erreur tracking page_view: {e}")

    # Retourner le layout complet avec navbar + sidebar + page
    return html.Div([
        user_navbar,
        sidebar,
        html.Div(
            id="page-container",
            style={
                "marginLeft": "240px",  # Largeur du sidebar
                "marginTop": "50px",     # Hauteur de la navbar
                "padding": "20px",
                "minHeight": "calc(100vh - 56px)",
                "background": "#f8fafc"
            },
            children=page_content
        )
    ])


@app.callback(
    [Output("auth-state", "data"),
     Output("login-error", "children"),
     Output("login-error", "style")],
    [Input("login-button", "n_clicks")],
    [State("login-username", "value"),
     State("login-password", "value"),
     State("auth-state", "data")],
    prevent_initial_call=True
)
def handle_login(n_clicks, username, password, current_auth):
    """
    🚀 VERSION OPTIMISÉE - Login ultra-rapide
    Le tracking Supabase se fait en arrière-plan pour ne pas bloquer
    """
    login_start = time.time()

    if not n_clicks or n_clicks == 0:
        raise dash.exceptions.PreventUpdate

    # Validation des champs (instantané)
    if not username or not password:
        return (
            current_auth or {"authenticated": False, "username": None},
            html.Div([
                html.Span("⚠️", style={"marginRight": "8px"}),
                "Veuillez remplir tous les champs"
            ], style={
                "background": "rgba(239, 68, 68, 0.15)",
                "border": "1px solid #ef4444",
                "borderRadius": "10px",
                "padding": "12px 16px",
                "marginBottom": "20px",
                "color": "#fca5a5",
                "fontSize": "13px"
            }),
            {"display": "block"}
        )

    # Vérification des identifiants (local = instantané)
    username = username.lower().strip()

    if verify_password(username, password):
        # ✅ Connexion réussie - réponse IMMÉDIATE
        user_info = get_user_info(username)

        # Générer un session_id local temporaire
        import uuid
        temp_session_id = str(uuid.uuid4())

        auth_data = {
            "authenticated": True,
            "username": username,
            "user_info": user_info,
            "session_id": temp_session_id,
            "user_id": None,
            "login_time": time.time()
        }

        # 🚀 TRACKING SUPABASE EN ARRIÈRE-PLAN (non-bloquant)
        def async_supabase_login():
            try:
                if supabase_client:
                    db_user = get_or_create_user(username)
                    if db_user:
                        user_id = db_user["id"]
                        session_id = create_session(user_id)
                        ACTIVE_SESSIONS[username] = {
                            "user_id": user_id,
                            "session_id": session_id
                        }
                        track_activity(user_id, session_id, "login", page="login")
                        print(f"   ✅ Supabase tracking OK pour {username}")
            except Exception as e:
                print(f"   ⚠️ Supabase tracking async: {e}")

        # Lancer en arrière-plan
        _async_executor.submit(async_supabase_login)

        elapsed = time.time() - login_start
        print(f"🔐 Login {username} réussi en {elapsed:.3f}s")

        return (auth_data, "", {"display": "none"})
    else:
        # Échec de connexion
        print(f"   ❌ Échec de connexion pour {username}")

        # Tracking échec en arrière-plan aussi
        def async_track_failed():
            try:
                if supabase_client:
                    track_activity(None, None, "login_failed", page="login", details={"username_attempted": username})
            except:
                pass

        _async_executor.submit(async_track_failed)

        return (
            {"authenticated": False, "username": None},
            html.Div([
                html.Span("❌", style={"marginRight": "8px"}),
                "Identifiant ou mot de passe incorrect"
            ], style={
                "background": "rgba(239, 68, 68, 0.15)",
                "border": "1px solid #ef4444",
                "borderRadius": "10px",
                "padding": "12px 16px",
                "marginBottom": "20px",
                "color": "#fca5a5",
                "fontSize": "13px"
            }),
            {"display": "block"}
        )


# Callback logout séparé avec allow_duplicate
@app.callback(
    Output("auth-state", "data", allow_duplicate=True),
    [Input("logout-button", "n_clicks")],
    [State("auth-state", "data")],
    prevent_initial_call=True
)
def handle_logout(n_clicks, current_auth):
    """Gère la déconnexion"""
    if not n_clicks or n_clicks == 0:
        raise dash.exceptions.PreventUpdate

    print("🔐 Déconnexion utilisateur")

    # ✅ TRACKING SUPABASE: Enregistrer la déconnexion
    if current_auth:
        username = current_auth.get("username")
        if username:
            track_logout(username)
            print(f"   📊 Déconnexion trackée: {username}")

    return {"authenticated": False, "username": None}


# ============================================================
# 🔐 FIN CALLBACKS AUTHENTIFICATION
# ============================================================


# NOTE: Ce callback de debug est désactivé car il cause des conflits
'''
@app.callback(
    Output('action-feedback', 'children'),
    Input('some-input', 'value'),
    prevent_initial_call=True
)
def update_feedback(input_value):
    ctx = dash.callback_context
    print(f"Callback triggered by: {ctx.triggered}")

    if input_value is None:
        return "No input provided"
    return f"Input value: {input_value}"
'''

@app.callback(
    Output('edit-product-output', 'children'),
    Input('edit-product', 'value'),
)
def update_product_name(value):
    if value:
        return f"Produit modifié: {value}"
    return no_update

# Callback pour stocker les données locales
@app.callback(
    Output({'type': 'storage', 'index': MATCH}, 'data'),
    Input({'type': 'button-storage', 'index': MATCH}, 'n_clicks'),
    State({'type': 'storage', 'index': MATCH}, 'data')
)
def store_data_in_local_storage(n_clicks, data):
    if n_clicks is None:
        return no_update

    data = data or {'clicks': 0}
    data['clicks'] = data['clicks'] + 1
    return data


# Callback pour afficher le nombre de clics dans chaque zone de stockage
@app.callback(
    Output({'type': 'output-storage', 'index': MATCH}, 'children'),
    Input({'type': 'storage', 'index': MATCH}, 'modified_timestamp'),
    State({'type': 'storage', 'index': MATCH}, 'data')
)
def display_storage_data(ts, data):
    if ts is None:
        return no_update
    data = data or {}
    return data.get('clicks', 0)


# ==================== CALLBACK 1 : CAPTURER LES MODIFICATIONS EN TEMPS RÉEL ====================
'''@app.callback(
    Output("qac-edits-store", "data"),
    Input("main-table", "data"),
    State("main-table", "data_previous"),
    State("qac-edits-store", "data"),
    prevent_initial_call=True
)
def capture_qac_edits(table_data, stored_edits):
    """
    Capture automatiquement les modifications de QAC edited
    et les sauvegarde dans localStorage
    """
    if not table_data:
        return stored_edits or {}

    stored_edits = stored_edits or {}

    # Parcourir les lignes du tableau
    for row in table_data:
        product_name = row.get("product_name")
        #qac_edited = row.get("QAC edited", "")
        qac_edited = row.get("QAC edited", "")
        if qac_edited is not None:
            qac_edited = qac_edited.strip()
        else:
            qac_edited = ""

        # Vérifier si qac_edited est None et le convertir en chaîne vide si nécessaire
        if qac_edited is not None:
            qac_edited = qac_edited.strip()
        else:
            qac_edited = ""  # Si c'est None, on le remplace par une chaîne vide

        # Si QAC edited n'est pas vide, sauvegarder
        if product_name and qac_edited:
            stored_edits[product_name] = {
                "qac_edited": qac_edited,
                "supplier": row.get("Supplier", ""),
                "timestamp": datetime.now().isoformat()
            }

    print(f"💾 [localStorage] {len(stored_edits)} QAC sauvegardés")
    return stored_edits
'''



# ==================== CALLBACK 2 : RESTAURER LES QAC AU CHARGEMENT ====================
@app.callback(
    Output("main-table", "data", allow_duplicate=True),
    Input("qac-edits-store", "modified_timestamp"),
    State("qac-edits-store", "data"),
    State("main-table", "data"),
    prevent_initial_call=True
)
def restore_qac_from_storage(timestamp, stored_edits, table_data):
    """
    Restaure les QAC édités depuis localStorage au chargement de la page
    """
    if not timestamp or not stored_edits or not table_data:
        return no_update

    restored_count = 0

    # Appliquer les QAC sauvegardés
    for row in table_data:
        product_name = row.get("product_name")
        if product_name in stored_edits:
            row["QAC edited"] = stored_edits[product_name]["qac_edited"]
            restored_count += 1

    if restored_count > 0:
        print(f"✅ [localStorage] {restored_count} QAC restaurés")

    return table_data


# ==================== CALLBACK 3 : SAUVEGARDER EN CSV (BOUTON) ====================
'''
@app.callback(
    [Output("action-feedback", "children", allow_duplicate=True),
     Output("main-table", "data", allow_duplicate=True)],
    Input('btn-save-qac', 'n_clicks'),
    [State("main-table", "data"),
     State("qac-edits-store", "data")],
    prevent_initial_call=True
)
def save_qac_to_csv(n_clicks, table_data, stored_edits):
    """
    Sauvegarde les QAC édités dans un fichier CSV
    Utilise les données du localStorage pour garantir la cohérence
    """
    if not n_clicks:
        return no_update, no_update

    try:
        file_path = 'qac_data.csv'

        # Préparer les données à sauvegarder
        rows_to_save = []

        for row in table_data:
            product_name = row.get('product_name')
            qac_edited = row.get('QAC edited', None)

            # Vérifier si qac_edited est None et le convertir en chaîne vide si nécessaire
            if qac_edited is None:
                qac_edited = ""  # Si c'est None, on le remplace par une chaîne vide
            else:
                qac_edited = qac_edited.strip()  # Applique strip() si ce n'est pas None

            # Sauvegarder seulement les lignes avec QAC edited non vide
            if qac_edited and qac_edited != " ":
                rows_to_save.append({
                    "product_name": product_name,
                    "supplier": row.get('Supplier', ''),
                    "total_stock": row.get('total_stock', 0),
                    "Average Daily Sales": row.get('Average Daily Sales', 0),
                    "QAC": row.get('QAC', 0),
                    "QAC edited": qac_edited,
                    "timestamp": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                })

        if not rows_to_save:
            return dbc.Alert(
                "⚠️ Aucune donnée QAC à sauvegarder",
                color="warning",
                duration=3000
            ), table_data

        # Écrire dans le CSV (mode écrasement pour éviter les doublons)
        with open(file_path, mode='w', newline='', encoding='utf-8') as f:
            fieldnames = ["product_name", "supplier", "total_stock",
                          "Average Daily Sales", "QAC", "QAC edited", "timestamp"]
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows_to_save)

        print(f"✅ [CSV] {len(rows_to_save)} lignes sauvegardées dans {file_path}")

        return dbc.Alert(
            f"✅ {len(rows_to_save)} QAC sauvegardés dans {file_path}",
            color="success",
            duration=4000
        ), table_data

    except Exception as e:
        print(f"❌ Erreur sauvegarde CSV : {e}")
        return dbc.Alert(
            f"❌ Erreur : {str(e)}",
            color="danger",
            duration=5000
        ), table_data


#========================Callback =================================#
@app.callback(
    Output("qac-edits-store", "data"),
    Input("main-table", "data"),
    State("main-table", "data_previous"),
    State("qac-edits-store", "data"),
    prevent_initial_call=True
)
def capture_qac_edits(current_data, previous_data, current_edits):
    """
    Détecte les modifications de QAC SANS BLOQUER l'édition.

    IMPORTANT : Ce callback NE TOUCHE PAS au tableau, seulement au Store.
    """
    from datetime import datetime
    import threading

    # ✅ Guards légers
    if not current_data or not previous_data:
        return current_edits or {}

    if not isinstance(current_data, list) or not isinstance(previous_data, list):
        return current_edits or {}

    try:
        df_current = pd.DataFrame(current_data)
        df_previous = pd.DataFrame(previous_data)
    except Exception as e:
        print(f"⚠️ Erreur création DataFrame: {e}")
        return current_edits or {}

    # Vérifier que QAC existe
    if 'QAC edited' not in df_current.columns or 'QAC edited' not in df_previous.columns:
        return current_edits or {}

    # Initialiser
    current_edits = current_edits or {}
    new_edits = {}
    has_changes = False

    # Helper pour conversion safe en float
    def safe_float(value, default=0.0):
        """Convertit n'importe quelle valeur en float de manière sûre"""
        if value is None or value == '':
            return default
        try:
            return float(value)
        except (ValueError, TypeError):
            return default

    # Détecter SEULEMENT les vraies modifications
    for idx, row in df_current.iterrows():
        try:
            product_name = str(row.get('product_name', '')).strip()
            supplier = str(row.get('Supplier', '')).strip()

            if not product_name or not supplier:
                continue

            product_key = f"{product_name}|{supplier}"

            # ✅ CORRECTION: Conversion sûre avec safe_float
            current_qac = safe_float(row.get('QAC edited'))

            # Trouver ligne correspondante dans previous
            prev_row = df_previous[
                (df_previous['product_name'] == product_name) &
                (df_previous['Supplier'] == supplier)
                ]

            if prev_row.empty:
                continue

            # ✅ CORRECTION: Conversion sûre avec safe_float
            previous_qac = safe_float(prev_row.iloc[0].get('QAC edited'))

            # ✅ Détecter changement RÉEL (avec tolérance pour éviter faux positifs)
            if abs(current_qac - previous_qac) > 0.01 and current_qac > 0:
                new_edits[product_key] = {
                    'QAC edited': current_qac,
                    'timestamp': datetime.now().isoformat(),
                    'product_name': product_name,
                    'Supplier': supplier
                }
                has_changes = True

                print(f"✏️ QAC modifiée : {product_name} ({supplier}) : {previous_qac:.1f} → {current_qac:.1f}")

        except Exception as e:
            print(f"⚠️ Erreur traitement ligne {idx}: {e}")
            continue  # Passer à la ligne suivante sans crasher

    # ✅ Si modifications détectées
    if has_changes:
        updated_edits = {**current_edits, **new_edits}

        # Sauvegarde CSV + Supabase asynchrone (non-bloquante)
        def save_async():
            try:
                save_qac_to_csv(new_edits)

                # ✅ NOUVEAU: Sauvegarder dans Supabase
                for product_key, edit_data in new_edits.items():
                    try:
                        # Récupérer user_id depuis la session active
                        user_id = None
                        session_id = None
                        for username, session_data in ACTIVE_SESSIONS.items():
                            user_id = session_data.get("user_id")
                            session_id = session_data.get("session_id")
                            break

                        if user_id:
                            # Trouver l'ancienne valeur
                            old_value = 0
                            product_name = edit_data.get('product_name', '')
                            supplier = edit_data.get('Supplier', '')

                            # Chercher dans previous_data
                            prev_row = df_previous[
                                (df_previous['product_name'] == product_name) &
                                (df_previous['Supplier'] == supplier)
                            ]
                            if not prev_row.empty:
                                old_value = safe_float(prev_row.iloc[0].get('QAC edited'))

                            new_value = edit_data.get('QAC edited', 0)

                            # Appeler track_qac_edit pour Supabase
                            track_qac_edit(
                                user_id=user_id,
                                session_id=session_id,
                                product_name=product_name,
                                old_value=int(old_value),
                                new_value=int(new_value),
                                supplier=supplier
                            )
                    except Exception as e:
                        print(f"⚠️ Erreur track_qac_edit Supabase: {e}")

            except Exception as e:
                print(f"⚠️ Erreur sauvegarde CSV (non bloquante) : {e}")

        # ✅ Exécuter dans un thread séparé pour ne pas bloquer l'UI
        thread = threading.Thread(target=save_async, daemon=True)
        thread.start()

        print(f"💾 {len(new_edits)} modif(s) QAC en cours de sauvegarde (total: {len(updated_edits)})")

        return updated_edits

    # ✅ Aucun changement → retourner l'état actuel
    return current_edits or {}


# ==================== CALLBACK 4 : CHARGER QAC DEPUIS CSV AU DÉMARRAGE =====================
@app.callback(
    Output("qac-edits-store", "data", allow_duplicate=True),
    Input("url", "pathname"),
    prevent_initial_call='initial_duplicate'
)
def load_qac_from_csv_on_startup(pathname):
    """
    Charge les QAC depuis le CSV au démarrage de l'application
    """
    if pathname not in ["/", None]:
        return no_update

    file_path = 'qac_data.csv'

    try:
        if not os.path.exists(file_path):
            return {}

        stored_edits = {}

        with open(file_path, mode='r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                product_name = row.get('product_name')
                if product_name:
                    stored_edits[product_name] = {
                        "qac_edited": row.get('QAC edited', ''),
                        "supplier": row.get('supplier', ''),
                        "timestamp": row.get('timestamp', '')
                    }

        if stored_edits:
            print(f"✅ [CSV] {len(stored_edits)} QAC chargés depuis {file_path}")

        return stored_edits

    except Exception as e:
        print(f"⚠️ Erreur chargement CSV : {e}")
        return {}
'''

# Callback pour mettre à jour le Store 'selected-product-for-notes'
@app.callback(
    Output('selected-product-for-notes', 'data', allow_duplicate = True),
    Input('main-table', 'active_cell'),
    State('main-table', 'data'),
    prevent_initial_call=True
)
def update_selected_product(active_cell, table_data):
    if active_cell:
        row = active_cell.get("row")
        product_name = table_data[row].get("product_name", "") if row is not None else ""
        return product_name  # Mettre à jour avec le nom du produit sélectionné
    return None


# Validation layout
#app.validation_layout = html.Div([
#   dcc.Location(id="url"),
#  dcc.Store(id="master-data"),
# dcc.Store(id="filtered-data"),
#dcc.Dropdown(id="filter-supplier"),
#   dcc.Dropdown(id="filter-category"),
#  dcc.Dropdown(id="filter-need"),  # ✅ IMPORTANT
# dcc.Dropdown(
#    id='filter-status',
#   options=[
#      {'label': 'Status 1', 'value': 'status1'},
#     {'label': 'Status 2', 'value': 'status2'}
#],
#       value='status1'
#  ),

# dcc.Input(id="search-input"),
#dbc.Checklist(id="toggle-options"),
#dcc.Store(id="uploaded-csv"),
#   dcc.Store(id="chat-store"),
#  dcc.Store(id="chat-open"),
# make_sidebar(),
#page_overview(initial_df),
#  page_analytics(),
#  page_predictive(),
# page_about(),
#html.Div(id="page-container"),
#  html.Button(id="chat-fab"),
# html.Div(id="chat-window"),
#html.Div(id="chat-messages"),
#  dbc.Textarea(id="chat-input"),
# dcc.Upload(id="chat-upload"),
#  html.Small(id="upload-status"),
# dbc.Button(id="chat-send"),
#  dbc.Button(id="chat-close"),
# dcc.Download(id="download-data"),
#dcc.Download(id="download-po"),
# dbc.Button(id="btn-add-row"),
#dbc.Modal(id="edit-modal"),
# dbc.Input(id="edit-product"),
#dbc.Input(id="edit-supplier"),
#  dbc.Input(id="edit-category"),
# dbc.Input(id="edit-stock"),
#dbc.Modal(id="notes-modal"),
# html.Div(id="notes-modal-title"),
# ✅ Ajouter les nouveaux composants
# html.Button(id="notes-fab"),
#   dcc.Dropdown(id="note-product-selector"),
#  dbc.Modal(id="notes-modal-new"),
# html.Div(id="notes-display-list"),
#dbc.Textarea(id="note-text-input"),
#  dbc.Input(id="note-author-input"),
# html.Div(id="note-feedback-new"),
#dbc.Button(id="note-modal-send"),
#  dbc.Button(id="note-modal-close"),
# dcc.Store(id="selected-product-for-notes"),
#])
# Validation layout
# Validation layout
app.validation_layout = html.Div([
    # ==================== NAVIGATION & STORES ====================
    dcc.Location(id="url"),
    dcc.Store(id="master-data"),
    dcc.Store(id="filtered-data"),
    dcc.Store(id="uploaded-csv"),
    dcc.Store(id="chat-store"),
    dcc.Store(id="chat-open"),
    dcc.Store(id={'type': 'selected-product-for-notes', 'index': '2'}),
    dcc.Store(id="edit-mode"),
    dcc.Store(id="edit-original-product"),
    #dcc.Store(id="qac-edits"),
    dcc.Store(id="qac-edits-store"),

    # ==================== COMPOSANTS CACHÉS (pour callbacks) ====================
    html.Div(id="action-feedback", style={"display": "none"}),
    html.Div(id="selection-counter", style={"display": "none"}),
    html.Span(id="selection-count-inline", style={"display": "none"}),
    dcc.Input(id="search-input", style={"display": "none"}),
    dcc.Dropdown(id="filter-supplier", style={"display": "none"}),
    dcc.Dropdown(id="filter-category", style={"display": "none"}),
    dcc.Dropdown(id="filter-need", style={"display": "none"}),
    dbc.Checklist(id="toggle-options", style={"display": "none"}),
    dbc.Button(id="btn-refresh", style={"display": "none"}),
    dbc.Button(id="btn-reload-data", style={"display": "none"}),
    dbc.Button(id="btn-add-row", style={"display": "none"}),
    dbc.Button(id="btn-po-pdf", style={"display": "none"}),
    dbc.Button(id="btn-run-agent-ia", style={"display": "none"}),
    dbc.Button(id="btn-select-all", style={"display": "none"}),
    dbc.Button(id="btn-clear-selection", style={"display": "none"}),
    dbc.Button(id="btn-fill-qac-target", style={"display": "none"}),
    dbc.Button(id="btn-save-qac", style={"display": "none"}),
    dcc.Download(id="download-data"),
    dcc.Download(id="download-po"),
    html.Div(id="debug-info", style={"display": "none"}),

    # Alertes (pour messages IA)
    dbc.Alert(id="alert-ia-analysis", is_open=False, dismissable=True),
    dbc.Alert(id="alert-bc-error", is_open=False, dismissable=True),

    # ==================== NAVIGATION ====================
    dbc.NavLink(id="nav-overview"),
    dbc.NavLink(id="nav-analytics"),
    dbc.NavLink(id="nav-pred"),
    dbc.NavLink(id="nav-about"),

    # ==================== PAGE CONTAINER ====================
    html.Div(id="page-container"),
    html.Div(id="action-feedback"),
    html.Div(id="qac-save-feedback"),  # ✅ AJOUTER CECI
    html.Div(id="selection-counter"),
    # ==================== MAIN TABLE (CRITIQUE) ====================
    dash_table.DataTable(
        id="main-table",
        columns=[
            {"name": "ID", "id": "product_id"},  # ✅ PREMIÈRE COLONNE = product_id
            {"name": "product_name", "id": "product_name"},
            {"name": "Supplier", "id": "Supplier"},
            {"name": "total_stock", "id": "total_stock"},
            {"name": "QAC", "id": "QAC"},
            {"name": "QAC edited", "id": "QAC edited", "editable": True},
        ],
        data=[],
        row_selectable="multi",
        selected_rows=[],
        editable=True
    ),

    # ==================== EDIT MODAL (VERSION COMPLÈTE) ====================
    html.Div(
        id="global-edit-modal-holder",
        children=[
            dbc.Modal(
                id="edit-modal",
                children=[
                    dbc.ModalHeader(dbc.ModalTitle("", id="edit-modal-title")),
                    dbc.ModalBody([
                        html.Label("Nom du produit", style={"fontWeight": "600", "marginBottom": "5px"}),
                        dbc.Input(id="edit-product-name", placeholder="Nom du produit", type="text"),
                        # dcc.Input(id='edit-product', type='text', placeholder='Modifier produit'),
                        html.Br(),

                        html.Label("Fournisseur", style={"fontWeight": "600", "marginBottom": "5px"}),
                        dbc.Input(id="edit-supplier", placeholder="Fournisseur", type="text"),
                        html.Br(),

                        html.Label("Catégorie", style={"fontWeight": "600", "marginBottom": "5px"}),
                        dbc.Input(id="edit-category", placeholder="Catégorie ABC/XYZ", type="text"),
                        html.Br(),

                        html.Label("Stock actuel", style={"fontWeight": "600", "marginBottom": "5px"}),
                        dbc.Input(id="edit-stock", placeholder="Stock", type="number", min=0),
                        html.Br(),

                        html.Label("QAC (Quantité à commander)", style={"fontWeight": "600", "marginBottom": "5px"}),
                        dbc.Input(id="edit-qac", placeholder="QAC", type="number", min=0),
                        html.Br(),

                        html.Label("Average Daily Sales", style={"fontWeight": "600", "marginBottom": "5px"}),
                        dbc.Input(id="edit-ads", placeholder="Ventes journalières moyennes", type="number", min=0),
                        html.Br(),

                        html.Label("Jours de crédit", style={"fontWeight": "600", "marginBottom": "5px"}),
                        dbc.Input(id="edit-credit", placeholder="Crédit (jours)", type="number", min=0),
                        html.Br(),

                        html.Label("Lead Time ajusté", style={"fontWeight": "600", "marginBottom": "5px"}),
                        dbc.Input(id="edit-leadtime", placeholder="Lead time (jours)", type="number", min=0),
                        html.Br(),

                        html.Label("Buffer ajusté", style={"fontWeight": "600", "marginBottom": "5px"}),
                        dbc.Input(id="edit-buffer", placeholder="Buffer (jours)", type="number", min=0),
                        html.Br(),

                        html.Label("Max Daily Sales", style={"fontWeight": "600", "marginBottom": "5px"}),
                        dbc.Input(id="edit-max-sales", placeholder="Ventes max journalières", type="number", min=0),
                        html.Br(),

                        html.Div(id="edit-modal-feedback", style={"color": "#10b981", "marginTop": "10px"})
                    ]),
                    dbc.ModalFooter([
                        dbc.Button("Annuler", id="edit-modal-close", outline=True, size="sm"),
                        dbc.Button("💾 Enregistrer", id="edit-modal-save", color="primary", size="sm", n_clicks=0),
                    ])
                ],
                is_open=False,
                backdrop="static",
            )
        ]
    ),

    # ==================== NOTES SYSTEM ====================
    html.Button(id="notes-fab"),
    dbc.Modal(
        id="notes-modal-new",
        is_open=False,
        children=[
            dbc.ModalHeader("Notes"),
            dbc.ModalBody([
                dcc.Dropdown(id="note-product-selector"),
                html.Div(id="notes-display-list"),
                dbc.Textarea(id="note-text-input"),
                dbc.Input(id="note-author-input"),
                html.Div(id="note-feedback-new")
            ]),
            dbc.ModalFooter([
                dbc.Button("Fermer", id="note-modal-close"),
                dbc.Button("Envoyer", id="note-modal-send")
            ])
        ]
    ),

    # ==================== CHAT ====================
    html.Button(id="chat-fab"),
    html.Div(id="chat-window", children=[
        html.Div(id="chat-close"),
        html.Div(id="chat-messages"),
        dcc.Upload(id="chat-upload"),
        html.Small(id="upload-status"),
        dbc.Textarea(id="chat-input"),
        dbc.Button(id="chat-send")
    ]),

    # ==================== ANALYTICS ====================
    dcc.Graph(id="analytics-scatter"),
    dcc.Dropdown(id="analytics-filter-supplier"),
    dcc.Dropdown(id="analytics-filter-category"),

    # ==================== PREDICTIONS ====================
    dcc.Graph(id="predictive-bar"),
    dash_table.DataTable(id="predictive-table", data=[], columns=[]),
    dcc.Dropdown(id="predictive-filter-supplier"),
    dcc.Dropdown(id="predictive-filter-category"),

    # ==================== PROMOTIONS ====================
    dash_table.DataTable(id="promo-table", data=[], columns=[]),

    # ==================== AUTRES COMPOSANTS ====================
    html.Div(id="conflict-alert"),
    dcc.ConfirmDialog(id="confirm-dialog"),
])
# ------------------------------ Routing ------------------------------------------
# ⚠️ ANCIEN CALLBACK DÉSACTIVÉ - Remplacé par render_page_content (avec authentification)
# @app.callback(
#     Output("page-container", "children"),
#     Input("url", "pathname"),
#     State("master-data", "data"),
#     prevent_initial_call=False
# )
# def render_page(path, master_json):
#     """Route vers les différentes pages selon l'URL"""
#     base = pd.DataFrame(json.loads(master_json)) if master_json else get_df_cached()
#
#     print(f"🔀 Routing vers : {path}")  # Debug
#
#     if path == "/analytics":
#         return page_analytics()
#     elif path == "/predictions":
#         return page_predictive()
#     elif path == "/promotions":  # ✅ AJOUTER CETTE CONDITION
#         return page_promotions()
#     elif path == "/about":
#         return page_about()
#     else:  # "/" ou autre
#         return page_overview(base)


# ------------------------------ Filtering logic ----------------------------------
def filter_dataframe(df: pd.DataFrame, query: str, suppliers: list, statuses: list, cats: list, options: list):
    """Filtre le DataFrame de manière cumulative et robuste"""

    # Toujours partir d'une copie
    out = df.copy()

    print(f"[filter_dataframe] Départ: {len(out)} lignes")
    print(f"[filter_dataframe] Filtres: query='{query}', suppliers={suppliers}, cats={cats}, options={options}")

    # ========== FILTRE 1: RECHERCHE TEXTUELLE ==========
    if query and query.strip():
        q = str(query).strip().lower()
        if 'product_name' in out.columns:
            out = out[out['product_name'].astype(str).str.lower().str.contains(q, na=False, regex=False)]
            print(f"[filter_dataframe] Après recherche '{q}': {len(out)} lignes")

    # ========== FILTRE 2: FOURNISSEUR (CORRIGÉ) ==========
    if suppliers and len(suppliers) > 0:
        if 'Supplier' in out.columns:
            # ✅ CORRECTION: Utiliser .isin() au lieu de any()
            # Normaliser les fournisseurs pour comparaison insensible à la casse
            suppliers_lower = [str(s).lower() for s in suppliers]
            out = out[out['Supplier'].astype(str).str.lower().isin(suppliers_lower)]
            print(f"[filter_dataframe] Après filtre fournisseurs {suppliers}: {len(out)} lignes")

    # ========== FILTRE 3: CATÉGORIE ==========
    if cats and len(cats) > 0:
        if 'Product Category' in out.columns:
            cats_lower = [str(c).lower() for c in cats]
            out = out[out['Product Category'].astype(str).str.lower().isin(cats_lower)]
            print(f"[filter_dataframe] Après filtre catégories {cats}: {len(out)} lignes")

    # ========== FILTRE 4: STATUTS (si utilisé) ==========
    if statuses and len(statuses) > 0:
        if 'Stock Status' in out.columns:
            statuses_lower = [str(s).lower() for s in statuses]
            out = out[out['Stock Status'].astype(str).str.lower().isin(statuses_lower)]
            print(f"[filter_dataframe] Après filtre statuts {statuses}: {len(out)} lignes")

    # ========== OPTIONS SPÉCIALES ==========
    options = options or []

    # Agrégation par produit
    if 'by_product' in options:
        print(f"[filter_dataframe] Agrégation par produit...")
        out = aggregate_by_product(out)
        print(f"[filter_dataframe] Après agrégation: {len(out)} lignes")

    print(f"[filter_dataframe] ✅ Résultat final: {len(out)} lignes")
    return out

def validate_core_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Garantit que les colonnes critiques existent et sont valides"""
    df = df.copy()

    # 0. product_id - CRITIQUE: doit être préservé
    if "product_id" in df.columns:
        # Garder les valeurs existantes, convertir en int
        df["product_id"] = pd.to_numeric(df["product_id"], errors='coerce').fillna(0).astype(int)
    else:
        df["product_id"] = 0

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
    if "Average Daily Sales" not in df.columns:
        df["Average Daily Sales"] = 0.1

    df["Average Daily Sales"] = pd.to_numeric(
        df["Average Daily Sales"],
        errors='coerce'
    ).fillna(0.1).clip(lower=0.1)

    return df


# ------------------------------ Callbacks: filtering / banner --------------------
# ==================================================================================
# SECTION CALLBACKS - VERSION NETTOYÉE (SUPPRIMER TOUS LES ANCIENS CALLBACKS)
# ==================================================================================

# NOTE: Ce callback est désactivé car dupliqué avec apply_filters (ligne ~7059)
'''
# ==================== CALLBACK 1 : INITIALISATION (PRIORITÉ 1) ====================
@app.callback(
    [Output("filtered-data", "data", allow_duplicate=True),
     Output("main-table", "data", allow_duplicate=True),
     Output("main-table", "selected_rows", allow_duplicate=True)],
    Input("master-data", "data"),
    prevent_initial_call=True  # ✅ CORRIGÉ: Ne pas s'exécuter avant que main-table existe
)
def initialize_table(master_json):
    """
    ⚡ Initialise le tableau - SE DÉCLENCHE aussi après rotation
    Temps d'exécution : < 0.5 seconde
    """
    start_time = time.time()

    print(f"\n{'=' * 60}")
    print(f"⚡ INITIALISATION TABLE")
    print(f"{'=' * 60}")

    try:
        # ✅ Charger données
        if master_json:
            print(f"📥 Chargement depuis master-data...")
            df = pd.DataFrame(json.loads(master_json))
        else:
            print(f"📥 Chargement depuis cache...")
            df = get_df_cached()

        elapsed_load = time.time() - start_time
        print(f"   ✅ {len(df)} produits en {elapsed_load:.3f}s")

        # ✅ Validation
        print(f"⚡ Validation...")
        df = validate_core_columns(df)

        # ✅ Supprimer colonnes promo + ADS temporaires
        promo_cols_to_remove = [
            'promo_status', 'days_remaining', 'uplift_pct', 'roi_pct',
            'promo_recommendation', 'promo_priority', 'net_profit_per_day',
            'discount_pct', 'sales_with_promo', 'sales_without_promo',
            'additional_sales_per_day', 'revenue_loss_per_day',
            'additional_profit_per_day',
            'Average Daily Sales (7d)', 'Average Daily Sales (30d)',
            'Average Daily Sales (3d)', 'Daily OOS Rate (7d)',
            'Daily OOS Rate (30d)', 'Stockout Probability',
            'Credit Adequacy Score'
        ]
        df = df.drop(
            columns=[c for c in promo_cols_to_remove if c in df.columns],
            errors='ignore'
        )

        elapsed_clean = time.time() - start_time - elapsed_load
        print(f"   ✅ Nettoyage en {elapsed_clean:.3f}s")

        # ✅ Formatage valeurs pour affichage
        print(f"⚡ Formatage...")

        for col in ['total_stock', 'QAC', 'target_quantity']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

        if 'Max Coverage Day' in df.columns:
            df['Max Coverage Day'] = pd.to_numeric(df['Max Coverage Day'], errors='coerce').fillna(0).round(1)

        if 'Average Daily Sales' in df.columns:
            df['Average Daily Sales'] = pd.to_numeric(df['Average Daily Sales'], errors='coerce').fillna(0.1).round(2)

        elapsed_format = time.time() - start_time - elapsed_load - elapsed_clean
        print(f"   ✅ Formatage en {elapsed_format:.3f}s")

        # ✅ Ajouter colonnes actions
        print(f"⚡ Ajout colonnes actions...")
        df = add_action_cols(df)

        # ✅ Statistiques
        if 'Average Daily Sales' in df.columns:
            print(f"\n📊 STATS ADS :")
            print(f"   Min     : {df['Average Daily Sales'].min():.2f}")
            print(f"   Médiane : {df['Average Daily Sales'].median():.2f}")
            print(f"   Max     : {df['Average Daily Sales'].max():.2f}")
            print(f"   Moyenne : {df['Average Daily Sales'].mean():.2f}")

        elapsed_total = time.time() - start_time

        print(f"\n⚡ INITIALISATION TERMINÉE en {elapsed_total:.3f}s")
        print(f"{'=' * 60}\n")

        return (
            df.to_json(orient="records"),
            df.to_dict("records"),
            []
        )

    except Exception as e:
        print(f"\n❌ ERREUR INITIALISATION : {e}")
        import traceback
        traceback.print_exc()
        print(f"{'=' * 60}\n")

        # Retourner données vides en cas d'erreur
        return "[]", [], []
'''

'''@app.callback(
    [Output("filtered-data", "data", allow_duplicate=True),
     Output("main-table", "data", allow_duplicate=True),
     Output("main-table", "selected_rows", allow_duplicate=True)],
    Input("url", "pathname"),
    State("master-data", "data"),
    prevent_initial_call='initial_duplicate'  # ✅ Corrigé
)
def initial_load_table(pathname, master_json):
    """S'exécute au chargement initial avec duplication autorisée."""
    if not master_json:
        raise dash.exceptions.PreventUpdate

    if pathname not in ["/", None]:
        raise dash.exceptions.PreventUpdate

    df = pd.DataFrame(json.loads(master_json))
    df = validate_core_columns(df)

    promo_cols_to_remove = [
        'promo_status', 'days_remaining', 'uplift_pct', 'roi_pct',
        'promo_recommendation', 'promo_priority', 'net_profit_per_day',
        'discount_pct', 'sales_with_promo', 'sales_without_promo'
    ]

    df = df.drop(columns=[c for c in promo_cols_to_remove if c in df.columns], errors='ignore')

    print(f"🚀 [initial_load] Chargement initial : {len(df)} lignes, 0 sélections")

    return df.to_json(orient="records"), df.to_dict("records"), []
'''
# ------------------------------ Notes System Callbacks ----------------------------

# ------------------------------ Export CSV ---------------------------------------
@app.callback(
    Output("download-data", "data"),
    Input("btn-export", "n_clicks"),
    State("filtered-data", "data"),
    prevent_initial_call=True
)
def export_csv(n, data_json):
    if not n or not data_json: return no_update
    df = pd.DataFrame(json.loads(data_json))
    for c in ["edit Edit", "delete Delete"]:
        if c in df.columns: df.drop(columns=[c], inplace=True)
    return dcc.send_data_frame(df.to_csv, f"supply_filtered_{datetime.now().strftime('%Y%m%d_%H%M')}.csv", index=False)


# ------------------------------ Export Purchase Order PDF ------------------------


# Fonction pour récupérer les données CSV depuis Google Sheets
def fetch_product_data_from_csv():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTrpcAiktxAPBiwznGOh35kVetc4O8-z5rQdFDgBaDE4OC3Jnb7JDGm59c55Cwm2pWCktcsBirWT_0b/pub?gid=751531326&single=true&output=csv"
    response = requests.get(url)
    if response.status_code == 200:
        product_data = pd.read_csv(io.StringIO(response.text))
        print(
            f"Colonnes du DataFrame : {product_data.columns.tolist()}")  # Afficher les colonnes pour vérifier le nom exact
        return product_data
    else:
        print(f"Erreur lors du téléchargement des données : {response.status_code}")
        return pd.DataFrame()


# Fonction pour générer un code de référence basé sur le nom du produit
def ref_from_name(name: str) -> str:
    """Génère un code de référence basé sur le nom du produit."""
    if not name:
        return ""
    parts = [p for p in str(name).split() if p]
    if not parts:
        return str(name)[:6].upper()
    left = (parts[0][:3] if len(parts[0]) >= 3 else parts[0]).upper()
    right = (parts[1][:3] if len(parts) > 1 and len(parts[1]) >= 3 else (
        parts[0][3:6] if len(parts[0]) > 3 else "")).upper()
    return "-".join([left, right]) if right else left


# Fonction pour calculer la quantité cible
def calculate_formula_based_target(df):
    """
    Formule de fallback pour produits sans historique de ventes.
    Calcule target_quantity basée sur la demande projetée.
    """
    ads = df.get('Average Daily Sales', pd.Series(0, index=df.index))
    leadtime = df.get('ADJUSTED_LEADTIME', pd.Series(7, index=df.index))
    credit = df.get('credit_days', pd.Series(14, index=df.index))
    buffer = df.get('AJUSTER_BUFFER', pd.Series(0, index=df.index))
    stock = df.get('total_stock', pd.Series(0, index=df.index))

    target = (
            ads * (leadtime + credit) +
            buffer * ads -
            stock
    )

    return np.maximum(0, target)


# Fonction pour récupérer le logo en base64 ou fichier local
def _get_logo_data_uri():
    try:
        logo_path = Path("logo_maad.jpg")
        if logo_path.exists():
            b64 = base64.b64encode(logo_path.read_bytes()).decode("utf-8")
            return f"data:image/jpeg;base64,{b64}"
    except Exception as e:
        print(f"❌ Erreur lors de la récupération du logo en base64: {e}")
    return None


LOGO_DATA_URI = _get_logo_data_uri()


def get_logo_for_reportlab():
    try:
        # Priorité 1 : Fichier local
        logo_path = Path("logo_maad.jpg")
        if logo_path.exists() and logo_path.stat().st_size > 0:
            print(f"✅ Logo trouvé : {logo_path.absolute()}")
            return str(logo_path)

        # Priorité 2 : Base64
        if LOGO_DATA_URI and "," in LOGO_DATA_URI:
            try:
                b64 = LOGO_DATA_URI.split(",", 1)[1] if LOGO_DATA_URI.startswith("data:") else LOGO_DATA_URI
                raw = base64.b64decode(b64)
                print(f"Logo chargé depuis base64 ({len(raw)} octets)")
                return ImageReader(io.BytesIO(raw))
            except Exception as e:
                print(f"❌ Échec du décodage base64 du logo : {e}")

        print("❌ Aucun logo disponible, continue sans")
        return None
    except Exception as e:
        print(f"❌ Erreur lors du chargement du logo : {e}")
        return None


# Remplacer TOUTE la section PDF (lignes ~1850-2000) par ceci :
# ============================== CACHE CATALOGUE & PACKAGING ==============================

# ============================================================
# 🚀 CACHE PRÉ-CHARGÉ POUR GÉNÉRATION BC ULTRA-RAPIDE
# ============================================================
_BC_CACHE = {
    "prices": None,
    "packaging": None,
    "catalog_packaging": None,
    "heroku_packaging": None,
    "logo_path": None,
    "loaded": False,
    "loading": False,
    "timestamp": 0
}
_BC_CACHE_LOCK = threading.Lock()


def preload_bc_cache():
    """
    Pré-charge toutes les données nécessaires pour la génération de BC.
    Appelé au démarrage et en arrière-plan.
    """
    global _BC_CACHE

    with _BC_CACHE_LOCK:
        if _BC_CACHE["loading"]:
            return
        _BC_CACHE["loading"] = True

    try:
        print("🔄 Pré-chargement cache BC...")
        start = time.time()

        # 1. Prix catalogue
        try:
            cat_url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTrpcAiktxAPBiwznGOh35kVetc4O8-z5rQdFDgBaDE4OC3Jnb7JDGm59c55Cwm2pWCktcsBirWT_0b/pub?gid=751531326&single=true&output=csv"
            catalog = load_csv_fast(cat_url, skiprows=1)
            if not catalog.empty and len(catalog.columns) >= 11:
                catalog_subset = catalog.iloc[:, [0, 1, 3, 10]].copy()
                catalog_subset.columns = ['product_id', 'product_name', 'selling_price', 'purchase_price']
                catalog_subset['product_name_clean'] = catalog_subset['product_name'].astype(str).str.lower().str.strip()
                prices = dict(zip(
                    catalog_subset['product_name_clean'],
                    pd.to_numeric(catalog_subset['purchase_price'], errors='coerce').fillna(1000)
                ))
                _BC_CACHE["prices"] = prices
                print(f"   ✅ Prix: {len(prices)} produits")
        except Exception as e:
            print(f"   ⚠️ Prix: {e}")
            _BC_CACHE["prices"] = {}

        # 2. Packaging map
        try:
            packaging = load_packaging_map()
            _BC_CACHE["packaging"] = packaging
            print(f"   ✅ Packaging: {len(packaging)} produits")
        except:
            _BC_CACHE["packaging"] = {}

        # 3. Catalog packaging (conditionnement)
        try:
            cat_pack = load_catalog_packaging()
            _BC_CACHE["catalog_packaging"] = cat_pack
        except:
            _BC_CACHE["catalog_packaging"] = {}

        # 4. Heroku packaging (fractions)
        try:
            heroku_pack = load_heroku_packaging()
            _BC_CACHE["heroku_packaging"] = heroku_pack
        except:
            _BC_CACHE["heroku_packaging"] = {}

        # 5. Chemin du logo (pré-résolu)
        logo_paths = [
            "assets/logo_maad.png",
            "assets/logo_maad.jpg",
            "assets/logo.png",
            "logo_maad.png",
            "logo.png"
        ]
        for path in logo_paths:
            if os.path.exists(path):
                _BC_CACHE["logo_path"] = path
                print(f"   ✅ Logo: {path}")
                break

        _BC_CACHE["loaded"] = True
        _BC_CACHE["timestamp"] = time.time()

        elapsed = time.time() - start
        print(f"✅ Cache BC pré-chargé en {elapsed:.2f}s")

    except Exception as e:
        print(f"❌ Erreur pré-chargement BC: {e}")
    finally:
        with _BC_CACHE_LOCK:
            _BC_CACHE["loading"] = False


def get_bc_cache():
    """Retourne le cache BC, le charge si nécessaire"""
    global _BC_CACHE

    # Si pas chargé ou expiré (>30 min), recharger en arrière-plan
    if not _BC_CACHE["loaded"] or (time.time() - _BC_CACHE["timestamp"]) > 1800:
        if not _BC_CACHE["loading"]:
            _async_executor.submit(preload_bc_cache)

        # Si jamais chargé, charger maintenant
        if not _BC_CACHE["loaded"]:
            preload_bc_cache()

    return _BC_CACHE


# Lancer le pré-chargement au démarrage (en arrière-plan)
_async_executor.submit(preload_bc_cache)


@cache.memoize(timeout=3600)  # Cache 1h
def get_catalog_prices():
    """Version rapide utilisant le cache pré-chargé"""
    bc_cache = get_bc_cache()
    if bc_cache["prices"]:
        return bc_cache["prices"]

    # Fallback: charger directement
    try:
        cat_url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTrpcAiktxAPBiwznGOh35kVetc4O8-z5rQdFDgBaDE4OC3Jnb7JDGm59c55Cwm2pWCktcsBirWT_0b/pub?gid=751531326&single=true&output=csv"
        catalog = load_csv_fast(cat_url, skiprows=1)

        if catalog.empty:
            return {}

        catalog_subset = catalog.iloc[:, [0, 1, 3, 10]].copy()
        catalog_subset.columns = ['product_id', 'product_name', 'selling_price', 'purchase_price']
        catalog_subset['product_name_clean'] = catalog_subset['product_name'].astype(str).str.lower().str.strip()

        return dict(zip(
            catalog_subset['product_name_clean'],
            pd.to_numeric(catalog_subset['purchase_price'], errors='coerce').fillna(1000)
        ))
    except Exception as e:
        print(f"❌ Erreur catalogue : {e}")
        return {}


@cache.memoize(timeout=3600)
def get_packaging_map_cached():
    """Version rapide utilisant le cache pré-chargé"""
    bc_cache = get_bc_cache()
    if bc_cache["packaging"]:
        return bc_cache["packaging"]

    try:
        return load_packaging_map()
    except Exception as e:
        print(f"❌ Erreur packaging : {e}")
        return {}
# ====== IMPORTS NÉCESSAIRES ======

# ===== Helpers packaging (ROBUSTES) =====
import re, math, csv, os, io
from datetime import datetime

PACKAGING_URL = "https://data.heroku.com/dataclips/cnmhrqqjneeunkbqibxklyxwcsrl.csv"

# ✅ URL du catalogue avec Conditionnement (AG) et Product Name propre (AH)
CATALOG_PACKAGING_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRTyAxh6v8o0FXV0r7f6ALPDgmeJNkjTZITjrEoKBHo2gs_f3iyV8sFk8fOzcAsUSkJMXBJCpJnhQKi/pub?gid=990028336&single=true&output=csv"

# Cache global pour le catalogue conditionnement
_catalog_packaging_cache = {
    "data": None,
    "timestamp": 0
}

# Cache pour le packaging Heroku
_heroku_packaging_cache = {
    "data": None,
    "timestamp": 0
}


def load_heroku_packaging():
    """
    🚀 VERSION OPTIMISÉE - Charge le packaging depuis Heroku dataclip
    Utilise la session HTTP poolée pour des connexions plus rapides

    Retourne: {product_name_lower: packaging_text}
    """
    global _heroku_packaging_cache

    # Vérifier le cache (10 minutes)
    if _heroku_packaging_cache["data"] is not None:
        age = time.time() - _heroku_packaging_cache["timestamp"]
        if age < 600:
            return _heroku_packaging_cache["data"]

    try:
        # Utiliser la session HTTP poolée
        dfp = load_csv_fast(PACKAGING_URL)

        if dfp.empty:
            return {}

        name_col = next((c for c in dfp.columns if c.lower() in ("name", "product_name", "designation")), None)
        pack_col = next((c for c in dfp.columns if ("pack" in c.lower()) or (c.lower() in ("packaging", "conditionnement"))), None)

        if not name_col or not pack_col:
            print("   ⚠️ Colonnes name/packaging non trouvées")
            return _heroku_packaging_cache["data"] or {}

        dfp["__key__"] = dfp[name_col].astype(str).str.lower().str.strip()
        dfp["__pack__"] = dfp[pack_col].astype(str).str.lower().str.strip()

        result = dict(zip(dfp["__key__"], dfp["__pack__"]))

        print(f"   ✅ {len(result)} produits avec packaging chargés")

        # Mettre en cache
        _heroku_packaging_cache["data"] = result
        _heroku_packaging_cache["timestamp"] = time.time()

        return result

    except Exception as e:
        print(f"   ❌ Erreur Heroku packaging: {e}")
        return _heroku_packaging_cache["data"] or {}


def load_catalog_packaging():
    """
    Charge le catalogue Google Sheet avec:
    - Colonne AG (index 32): Conditionnement (ex: 12, 24, 6)
    - Colonne AH (index 33): Product Name propre (sans "1/2 carton")
    - Colonne B (index 1): Nom original du produit

    Retourne: {product_name_lower: {"conditionnement": int, "clean_name": str}}
    """
    global _catalog_packaging_cache

    # Vérifier le cache (10 minutes)
    if _catalog_packaging_cache["data"] is not None:
        age = time.time() - _catalog_packaging_cache["timestamp"]
        if age < 600:
            return _catalog_packaging_cache["data"]

    try:
        print("📦 Chargement catalogue conditionnement (Google Sheet)...")
        df = load_csv_fast(CATALOG_PACKAGING_URL)  # Utilise load_csv_fast avec timeout intégré

        if df is None or df.empty:
            print("   ⚠️ Catalogue vide ou non chargé")
            return _catalog_packaging_cache["data"] or {}

        print(f"   📊 {df.shape[0]} lignes, {df.shape[1]} colonnes")

        result = {}

        # Vérifier qu'on a assez de colonnes (au moins 34 pour AH)
        if df.shape[1] >= 34:
            for idx, row in df.iterrows():
                try:
                    # Nom original (colonne B, index 1)
                    original_name = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                    if not original_name or original_name.lower() == "nan":
                        continue

                    key = original_name.lower().strip()

                    # Conditionnement (colonne AG, index 32)
                    conditionnement = 1
                    if pd.notna(row.iloc[32]):
                        try:
                            conditionnement = int(float(str(row.iloc[32]).strip()))
                            if conditionnement <= 0:
                                conditionnement = 1
                        except (ValueError, TypeError):
                            conditionnement = 1

                    # Nom propre (colonne AH, index 33)
                    clean_name = original_name
                    if pd.notna(row.iloc[33]):
                        temp_name = str(row.iloc[33]).strip()
                        if temp_name and temp_name.lower() != "nan":
                            clean_name = temp_name

                    result[key] = {
                        "conditionnement": conditionnement,
                        "clean_name": clean_name
                    }

                except Exception:
                    continue

            print(f"   ✅ {len(result)} produits avec conditionnement chargés")
        else:
            print(f"   ⚠️ Pas assez de colonnes ({df.shape[1]} < 34)")

        # Mettre en cache
        _catalog_packaging_cache["data"] = result
        _catalog_packaging_cache["timestamp"] = time.time()

        return result

    except Exception as e:
        print(f"   ❌ Erreur catalogue: {e}")
        return _catalog_packaging_cache["data"] or {}


def get_clean_product_name(product_name: str, catalog_packaging: dict) -> str:
    """
    Retourne le nom propre du produit (sans "1/2 carton", "1/2 demi carton", etc.)
    Utilise d'abord le catalogue, sinon nettoie manuellement.
    """
    if not product_name:
        return product_name

    key = product_name.lower().strip()

    # 1. Chercher dans le catalogue Google Sheet
    if key in catalog_packaging:
        clean = catalog_packaging[key].get("clean_name", "")
        if clean and clean.lower() != "nan":
            return clean

    # 2. Sinon, nettoyer manuellement
    clean = product_name
    patterns_to_remove = [
        "1/2 demi carton ", "1/2 demi-carton ",
        "1/2 carton ", "1/2carton ",
        "demi carton ", "demi-carton ",
        "1/4 carton ", "1/4carton ",
        "1/3 carton ", "1/3carton ",
        "1/2 demi carton", "1/2 demi-carton",
        "1/2 carton", "1/2carton",
        "demi carton", "demi-carton",
        "1/4 carton", "1/4carton",
        "1/3 carton", "1/3carton",
    ]

    name_lower = clean.lower()
    for pattern in patterns_to_remove:
        if name_lower.startswith(pattern):
            clean = clean[len(pattern):]
            break
        elif pattern in name_lower:
            idx = name_lower.find(pattern)
            clean = clean[:idx] + clean[idx + len(pattern):]
            break

    # Nettoyer les espaces multiples
    clean = " ".join(clean.split()).strip()

    return clean if clean else product_name


def get_conditionnement(product_name: str, catalog_packaging: dict) -> int:
    """
    Retourne le conditionnement du produit (nombre d'unités par carton).
    """
    if not product_name:
        return 1

    key = product_name.lower().strip()

    if key in catalog_packaging:
        cond = catalog_packaging[key].get("conditionnement", 1)
        if cond and cond > 0:
            return cond

    return 1


def get_fraction_from_packaging(product_name: str, heroku_packaging: dict) -> float:
    """
    Détecte la fraction depuis le packaging Heroku (1/2, 1/4, etc.)

    Exemples:
    - "1/2 carton" → 0.5
    - "1/4 carton" → 0.25
    - "carton" ou vide → 1.0
    """
    if not product_name:
        return 1.0

    key = product_name.lower().strip()

    # 1. Chercher dans Heroku
    packaging_text = heroku_packaging.get(key, "")

    if packaging_text:
        frac = detect_fraction_from_text(packaging_text)
        if frac != 1.0:
            return frac

    # 2. Chercher dans le nom du produit
    frac = detect_fraction_from_text(product_name)

    return frac


def calculate_quantity_for_po(qac_edited: float, product_name: str, catalog_packaging: dict, heroku_packaging: dict) -> int:
    """
    Calcule la quantité finale pour le bon de commande.

    Logique:
    1. Détecter si c'est un demi-carton (fraction depuis Heroku)
    2. Convertir en unités majeures (cartons entiers)
    3. Arrondir au conditionnement supérieur

    Exemple:
    - Produit "1/2 carton Biscuit", QAC edited = 10, Conditionnement = 12
    - Fraction = 0.5 (demi-carton)
    - Unités majeures = 10 * 0.5 = 5 cartons
    - Arrondi au conditionnement: ceil(5) = 5 (pas d'arrondi au conditionnement ici car c'est en cartons)

    Pour les produits NON fractionnés:
    - Produit "Biscuit", QAC edited = 25, Conditionnement = 12
    - Fraction = 1.0
    - Quantité = ceil(25 / 12) * 12 = 36
    """
    if qac_edited <= 0:
        return 0

    # 1. Obtenir la fraction (1/2, 1/4, etc.)
    fraction = get_fraction_from_packaging(product_name, heroku_packaging)

    # 2. Obtenir le conditionnement
    conditionnement = get_conditionnement(product_name, catalog_packaging)

    # 3. Calculer la quantité
    if fraction < 1.0:
        # C'est un demi/quart carton → convertir en cartons entiers
        # QAC edited est en demi-cartons, on convertit en cartons
        qty_cartons = qac_edited * fraction
        # Arrondir au supérieur
        qty_final = max(1, int(math.ceil(qty_cartons)))
    else:
        # Produit normal → arrondir au conditionnement
        if conditionnement > 1:
            nb_cartons = math.ceil(qac_edited / conditionnement)
            qty_final = nb_cartons * conditionnement
        else:
            qty_final = int(math.ceil(qac_edited))

    return qty_final


def load_packaging_map():
    """Map: product_name_lower -> packaging (ex: '1/2 carton')."""
    try:
        dfp = pd.read_csv(PACKAGING_URL)
        name_col = next((c for c in dfp.columns if c.lower() in ("name", "product_name", "designation")), None)
        pack_col = next((c for c in dfp.columns if ("pack" in c.lower()) or (c.lower() in ("packaging", "conditionnement"))), None)
        if not name_col or not pack_col:
            print("⚠️ Dataclip: colonnes name/packaging non trouvées.")
            return {}
        dfp["__key__"]  = dfp[name_col].astype(str).str.lower().str.strip()
        dfp["__pack__"] = dfp[pack_col].astype(str).str.lower().str.strip()
        return dict(zip(dfp["__key__"], dfp["__pack__"]))
    except Exception as e:
        print(f"❌ load_packaging_map: {e}")
        return {}

FRACTION_FINDER = re.compile(r"(\d+)\s*/\s*(\d+)", re.IGNORECASE)

def detect_fraction_from_text(txt: str) -> float:
    """Retourne la fraction trouvée (ex: '1/2' -> 0.5) ou 1.0 si rien."""
    if not txt:
        return 1.0
    s = str(txt).lower()
    m = FRACTION_FINDER.search(s)
    if m:
        try:
            num = int(m.group(1)); den = int(m.group(2))
            if den > 0:
                return num / den
        except (ValueError, TypeError):
            pass
    return 1.0

def consolidate_to_major(qty_units: float, packaging_text: str, product_name: str) -> int:
    """Convertit la QAC (éventuellement en sous-unité) → unités majeures entières (ceil)."""
    frac = detect_fraction_from_text(packaging_text)
    if frac == 1.0 and product_name:
        frac = detect_fraction_from_text(product_name)
    if frac <= 0:
        frac = 1.0
    maj = float(qty_units) * float(frac)
    return max(0, int(math.ceil(maj)))

def safe_float(v, default=0.0):
    try:
        if v is None or str(v).strip() == "":
            return default
        return float(str(v).strip())
    except (ValueError, TypeError):
        return default


# ============================================
# AGENT IA - CALCUL QAC AVEC GEMINI
# ============================================
def agent_ia_calculer_qac(table_data, use_gemini=True):
    """
    Calcul intelligent des QAC avec Gemini ou algorithme local

    Args:
        table_data: Données du tableau
        use_gemini: True pour utiliser Gemini, False pour algorithme local

    Returns:
        dict: {index: qac_calcule, ...}
    """
    if not table_data:
        return {}

    # Préparation des données
    produits = []
    for idx, prod in enumerate(table_data):
        prod_name = str(prod.get("product_name", "")).strip()
        if not prod_name:
            continue

        current_stock = safe_float(prod.get("current_stock", 0), 0)
        target_qty = safe_float(prod.get("target_quantity", 0), 0)
        supplier = str(prod.get("Supplier", "N/A")).strip()

        if target_qty <= 0:
            continue

        stock_pct = (current_stock / target_qty * 100) if target_qty > 0 else 100

        produits.append({
            'index': idx,
            'nom': prod_name[:50],
            'stock_actuel': current_stock,
            'quantite_cible': target_qty,
            'stock_pct': round(stock_pct, 1),
            'fournisseur': supplier
        })

    if not produits:
        return {}

    # Si Gemini désactivé ou échec, utiliser algorithme local
    if not use_gemini:
        return _agent_local(produits)

    # Tentative avec Gemini
    try:
        return _agent_gemini(produits)
    except Exception as e:
        print(f"❌ Erreur Gemini : {e}")
        print("⚠️ Basculement sur algorithme local")
        return _agent_local(produits)


def _agent_gemini(produits):
    """Analyse avec Gemini API"""
    print(f"\n🤖 Agent IA (Gemini 2.5 Flash) - Analyse de {len(produits)} produits...")

    # Prompt optimisé
    prompt = f"""Tu es un expert en gestion de stock. Analyse ces produits et détermine lesquels commander.

DONNÉES ({len(produits)} produits) :
{json.dumps(produits, indent=2, ensure_ascii=False)}

RÈGLES DE DÉCISION :
1. Stock < 15% de la cible → URGENT (commander pour atteindre 100% de la cible)
2. Stock entre 15-30% → IMPORTANT (commander pour atteindre 100%)
3. Stock entre 30-50% → NORMAL (commander pour atteindre 100%)
4. Stock > 50% → PAS DE COMMANDE

CALCUL QAC :
- QAC = quantite_cible - stock_actuel
- Arrondir à l'entier supérieur
- Minimum = 1 si commande nécessaire

FORMAT DE RÉPONSE :
Réponds UNIQUEMENT avec ce JSON exact (sans texte, sans markdown) :
{{
  "produits": [
    {{"index": 0, "qac": 150, "commander": true, "priorite": "URGENT"}},
    {{"index": 1, "qac": 0, "commander": false, "priorite": "OK"}}
  ],
  "statistiques": {{
    "urgents": 2,
    "importants": 3,
    "normaux": 1,
    "total_commander": 6
  }}
}}"""

    try:
        # Utilisation du modèle Flash (le plus rapide et gratuit)
        model = genai.GenerativeModel(GEMINI_FLASH)

        # Configuration optimisée pour JSON
        generation_config = genai.types.GenerationConfig(
            temperature=0.1,  # Très bas pour cohérence
            top_p=0.8,
            top_k=20,
            max_output_tokens=4096,
            response_mime_type="application/json"  # Force JSON
        )

        # Appel API
        response = model.generate_content(
            prompt,
            generation_config=generation_config,
            safety_settings={
                'HARASSMENT': 'block_none',
                'HATE_SPEECH': 'block_none',
                'SEXUALLY_EXPLICIT': 'block_none',
                'DANGEROUS_CONTENT': 'block_none'
            }
        )

        # Extraction réponse
        texte = response.text.strip()

        # Nettoyage
        texte = texte.replace('```json', '').replace('```', '').strip()

        # Chercher le JSON s'il y a du texte avant/après
        start = texte.find('{')
        end = texte.rfind('}') + 1
        if start != -1 and end > start:
            texte = texte[start:end]

        # Parse
        data = json.loads(texte)

        # Extraction résultats
        resultats = {}
        for prod in data.get('produits', []):
            if prod.get('commander', False):
                idx = prod['index']
                qac = prod.get('qac', 0)
                if qac > 0:
                    resultats[idx] = qac

        # Stats
        stats = data.get('statistiques', {})
        print(f"✅ Gemini : {len(resultats)} produits à commander")
        print(
            f"   🔴 {stats.get('urgents', 0)} urgents | 🟠 {stats.get('importants', 0)} importants | 🟢 {stats.get('normaux', 0)} normaux")

        return resultats

    except json.JSONDecodeError as e:
        print(f"❌ JSON invalide : {e}")
        print(f"   Réponse : {texte[:300]}...")
        raise
    except Exception as e:
        print(f"❌ Erreur API Gemini : {e}")
        raise


def _agent_local(produits):
    """Algorithme local de secours (sans API)"""
    print(f"\n🤖 Agent IA (Local) - Analyse de {len(produits)} produits...")

    resultats = {}
    stats = {'urgents': 0, 'importants': 0, 'normaux': 0}

    for p in produits:
        idx = p['index']
        stock_pct = p['stock_pct']
        current = p['stock_actuel']
        target = p['quantite_cible']

        qac = 0
        priorite = None

        if current <= 0:
            qac = target
            priorite = "URGENT"
            stats['urgents'] += 1
        elif stock_pct < 15:
            qac = target - current
            priorite = "URGENT"
            stats['urgents'] += 1
        elif stock_pct < 30:
            qac = target - current
            priorite = "IMPORTANT"
            stats['importants'] += 1
        elif stock_pct < 50:
            qac = target - current
            priorite = "NORMAL"
            stats['normaux'] += 1

        if qac > 0:
            resultats[idx] = int(math.ceil(qac))
            # Log détaillé
            print(
                f"   [{idx:3d}] {p['nom'][:35]:35} | {current:6.0f}/{target:6.0f} ({stock_pct:5.1f}%) → QAC: {resultats[idx]:5d} | {priorite}")

    print(f"\n✅ Local : {len(resultats)} produits à commander")
    print(f"   🔴 {stats['urgents']} urgents | 🟠 {stats['importants']} importants | 🟢 {stats['normaux']} normaux")

    return resultats


# ===== CALLBACK : AGENT IA =====
# ===== CALLBACK : AGENT IA (OPTIMISÉ) =====
@app.callback(
    [Output("main-table", "data", allow_duplicate=True),
     Output("main-table", "selected_rows", allow_duplicate=True),
     Output("btn-po-pdf", "disabled", allow_duplicate=True)],
    Input("btn-run-agent-ia", "n_clicks"),
    State("main-table", "data"),
    prevent_initial_call=True
)
def run_agent_ia_calcul(n_clicks, table_data):
    """
    🤖 AGENT IA - Génère automatiquement les commandes pour les produits en rupture

    LOGIQUE:
    1. Identifie les produits en rupture (stock <= 0) ou à risque (couverture < 7j)
    2. Remplit automatiquement QAC edited avec QAC calculé
    3. Sélectionne ces produits pour le bon de commande
    4. Active le bouton BC
    """
    if not n_clicks or not table_data:
        return no_update, no_update, no_update

    print(f"\n{'='*60}")
    print(f"🤖 AGENT IA - Analyse automatique des besoins")
    print(f"{'='*60}")

    updated_data = []
    indices_selection = []
    stats = {"rupture": 0, "risque": 0, "order_now": 0, "total_qac": 0}

    for idx, prod in enumerate(table_data):
        row = dict(prod)  # Copie

        prod_name = str(row.get("product_name", "")).strip()
        supplier = str(row.get("Supplier", "")).strip()

        if not prod_name or not supplier or supplier.lower() == "nan":
            updated_data.append(row)
            continue

        # Récupérer les données
        stock = safe_float(row.get("total_stock"), safe_float(row.get("current_stock"), 0))
        qac_auto = safe_float(row.get("QAC"), 0)
        coverage = safe_float(row.get("Max Coverage Day"), safe_float(row.get("coverage_days"), 999))
        ajusted_need = str(row.get("Ajusted_total_need", "")).upper()
        stockout_prob = safe_float(row.get("Stockout Probability"), 0)

        # ========================================
        # 📋 CRITÈRES DE SÉLECTION
        # ========================================
        need_order = False
        reason = ""

        # Critère 1: Rupture de stock (stock = 0)
        if stock <= 0 and qac_auto > 0:
            need_order = True
            reason = "RUPTURE"
            stats["rupture"] += 1

        # Critère 2: Couverture < 7 jours
        elif coverage < 7 and qac_auto > 0:
            need_order = True
            reason = "RISQUE (<7j)"
            stats["risque"] += 1

        # Critère 3: ORDER NOW
        elif "ORDER NOW" in ajusted_need and qac_auto > 0:
            need_order = True
            reason = "ORDER NOW"
            stats["order_now"] += 1

        # ========================================
        # 📝 REMPLIR QAC EDITED SI BESOIN
        # ========================================
        if need_order and qac_auto > 0:
            # Remplir QAC edited avec QAC calculé
            row["QAC edited"] = str(int(qac_auto))
            indices_selection.append(idx)
            stats["total_qac"] += qac_auto
            print(f"   ✅ [{reason}] {prod_name[:35]} → QAC: {int(qac_auto)} ({supplier})")

        updated_data.append(row)

    # ========================================
    # 📊 RÉSUMÉ
    # ========================================
    print(f"\n{'='*60}")
    print(f"📊 RÉSUMÉ AGENT IA")
    print(f"{'='*60}")
    print(f"   🔴 Ruptures: {stats['rupture']}")
    print(f"   🟠 Risque (<7j): {stats['risque']}")
    print(f"   🟡 Order Now: {stats['order_now']}")
    print(f"   📦 Total produits à commander: {len(indices_selection)}")
    print(f"   🔢 QAC total: {int(stats['total_qac'])} unités")
    print(f"{'='*60}\n")

    if not indices_selection:
        print("⚠️ Aucun produit à commander")
        return updated_data, [], True  # Désactiver bouton

    # Activer le bouton BC
    return updated_data, sorted(indices_selection), False


# ============================================
# FONCTION 2 : VALIDATION QAC
# ============================================
def valider_qac_selection(selected_rows, table_data):
    """
    Valide que toutes les lignes ont QAC edited > 0
    Retourne : (is_valid, error_message, missing_indices)
    """
    missing = []
    produits_manquants = []

    for idx in selected_rows:
        if 0 <= idx < len(table_data):
            qac = safe_float(table_data[idx].get("QAC edited", 0), 0)
            if qac <= 0:
                missing.append(idx)
                nom = str(table_data[idx].get("product_name", ""))[:40]
                produits_manquants.append(nom)

    if missing:
        msg = f"❌ {len(missing)} produit(s) sans QAC edited :\n"
        for i, nom in enumerate(produits_manquants[:5], 1):
            msg += f"  {i}. {nom}\n"
        if len(produits_manquants) > 5:
            msg += f"\n... et {len(produits_manquants) - 5} autre(s)\n"
        msg += "\n⚠️ Utilisez '🤖 Agent IA' ou remplissez manuellement"
        return False, msg, missing

    return True, "", []


# ============================================
# FONCTION 3 : GÉNÉRATION EXCEL AVEC FORMULES + LOGO + SIGNATURE
# ============================================
def is_riz_cereales(product_name: str) -> bool:
    """
    Détecte si un produit est du riz ou des céréales.
    Ces produits ont leur prix HT (TVA non incluse) contrairement aux autres.
    """
    name_lower = str(product_name).lower()
    keywords = ['riz', 'cereale', 'céréale', 'cereal', 'semoule', 'couscous', 'mil', 'maïs', 'mais', 'fonio']
    return any(kw in name_lower for kw in keywords)


def generer_bc_excel_avec_formules(selected_products, price_map, packaging_map):
    """
    🚀 VERSION OPTIMISÉE - Génère un Excel ultra-rapide
    - Cache pré-chargé
    - Styles pré-compilés
    - Logo pré-résolu

    📋 LOGIQUE TVA:
    - Produits normaux: Prix catalogue = TTC → on calcule le HT (PU/1.18)
    - Riz et céréales: Prix catalogue = HT → on ajoute la TVA (PU*0.18)
    """
    if not selected_products:
        raise ValueError("Aucun produit sélectionné")

    gen_start = time.time()

    # ========================================
    # 🚀 CHARGER DEPUIS CACHE PRÉ-CHARGÉ
    # ========================================
    bc_cache = get_bc_cache()
    catalog_packaging = bc_cache.get("catalog_packaging") or {}
    heroku_packaging = bc_cache.get("heroku_packaging") or {}
    logo_path = bc_cache.get("logo_path")

    # Regroupement par fournisseur
    suppliers = {}
    for prod in selected_products:
        sup = str(prod.get("Supplier", "")).strip()
        if not sup or sup.lower() == "nan":
            sup = "Fournisseur non spécifié"
        suppliers.setdefault(sup, []).append(prod)

    # Création workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Bon de Commande"

    # ========================================
    # 🎨 STYLES PRÉ-COMPILÉS (une seule fois)
    # ========================================
    header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    remise_fill = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
    escompte_fill = PatternFill(start_color="FFE0B2", end_color="FFE0B2", fill_type="solid")
    total_fill = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_align = Alignment(horizontal='center', vertical='center')
    left_align = Alignment(horizontal='left', vertical='center')
    right_align = Alignment(horizontal='right', vertical='center')
    bold_font = Font(bold=True)
    title_font = Font(size=18, bold=True, color="003366")
    supplier_font = Font(size=14, bold=True, color="228B22")
    supplier_fill = PatternFill(start_color="F0FFF0", end_color="F0FFF0", fill_type="solid")

    # ========================================
    # 🎨 LOGO MAAD (pré-résolu)
    # ========================================
    current_row = 1

    if logo_path:
        try:
            from openpyxl.drawing.image import Image as XLImage
            img = XLImage(logo_path)
            original_width = img.width
            original_height = img.height
            img.width = 180
            img.height = int(180 * original_height / original_width) if original_width > 0 else 60
            ws.add_image(img, 'A1')
            ws.row_dimensions[1].height = 25
            ws.row_dimensions[2].height = 25
            ws.row_dimensions[3].height = 25
            current_row = 5
        except:
            ws['A1'] = "MAAD"
            ws['A1'].font = Font(size=28, bold=True, color="1E40AF")
            current_row = 3
    else:
        ws['A1'] = "MAAD"
        ws['A1'].font = Font(size=28, bold=True, color="1E40AF")
        ws['A2'] = "Marketplace Africain de Distribution"
        ws['A2'].font = Font(size=11, italic=True, color="64748B")
        ws.row_dimensions[1].height = 40
        current_row = 4

    # ========================================
    # 📋 EN-TÊTE DU BON DE COMMANDE
    # ========================================
    po_number = get_next_po_number()
    current_row += 1

    # Titre
    ws[f'A{current_row}'] = f'BON DE COMMANDE N° {po_number}'
    ws[f'A{current_row}'].font = title_font
    ws.merge_cells(f'A{current_row}:I{current_row}')
    ws[f'A{current_row}'].alignment = center_align
    ws.row_dimensions[current_row].height = 30
    current_row += 2

    # Infos entreprise
    ws[f'A{current_row}'] = COMPANY_NAME
    ws[f'A{current_row}'].font = Font(bold=True, size=12, color="1E40AF")
    ws[f'G{current_row}'] = f"Date : {datetime.now().strftime('%d/%m/%Y')}"
    ws[f'G{current_row}'].font = Font(bold=True, size=11)
    ws[f'G{current_row}'].alignment = right_align
    current_row += 1

    ws[f'A{current_row}'] = COMPANY_ADDRESS
    ws[f'A{current_row}'].font = Font(size=10, color="64748B")
    current_row += 1

    ws[f'A{current_row}'] = f"Tél : {COMPANY_PHONE} | Email : {COMPANY_EMAIL}"
    ws[f'A{current_row}'].font = Font(size=10, color="64748B")
    current_row += 2

    TVA_RATE = 0.18

    # ========================================
    # 📊 TABLEAUX PAR FOURNISSEUR
    # ========================================
    headers = ['Réf.', 'Désignation', 'Qté', 'Cond.', 'Unité', 'PU HT', 'TVA Unit.', 'PU TTC', 'Remise %', 'Total HT',
               'Total TTC']

    for supplier, products in suppliers.items():
        # Titre fournisseur
        ws[f'A{current_row}'] = f'📦 FOURNISSEUR : {supplier.upper()}'
        ws[f'A{current_row}'].font = supplier_font
        ws[f'A{current_row}'].fill = supplier_fill
        ws.merge_cells(f'A{current_row}:K{current_row}')
        ws.row_dimensions[current_row].height = 25
        current_row += 1

        # En-têtes colonnes
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = border
        ws.cell(row=current_row, column=9).fill = remise_fill
        current_row += 1
        first_data_row = current_row

        # ========================================
        # 📦 LIGNES PRODUITS (optimisé)
        # ========================================
        for prod in products:
            prod_name = str(prod.get("product_name", "")).strip()
            prod_key = prod_name.lower()

            # 🔄 PRENDRE QAC edited SI disponible, SINON QAC calculé
            qac_edited = safe_float(prod.get("QAC edited"), 0.0)
            qac_auto = safe_float(prod.get("QAC"), 0.0)

            # Utiliser QAC edited si rempli, sinon QAC auto
            qac_final = qac_edited if qac_edited > 0 else qac_auto

            if qac_final <= 0:
                continue

            # Calculs avec cache pré-chargé
            qty_final = calculate_quantity_for_po(qac_final, prod_name, catalog_packaging, heroku_packaging)
            clean_name = get_clean_product_name(prod_name, catalog_packaging)
            catalog_price = safe_float(price_map.get(prod_key, 1000.0), 1000.0)
            if catalog_price <= 0:
                catalog_price = 1000.0

            # Obtenir conditionnement et fraction pour affichage
            conditionnement = get_conditionnement(prod_name, catalog_packaging)
            fraction = get_fraction_from_packaging(prod_name, heroku_packaging)

            # Déterminer l'unité affichée
            if fraction < 1.0:
                unite = "Carton"  # Demi-cartons convertis en cartons
            elif conditionnement > 1:
                unite = f"×{conditionnement}"  # Carton de X
            else:
                unite = "Unité"

            product_id = prod.get("product_id", "")
            if product_id:
                try:
                    product_id = int(float(product_id))
                except:
                    product_id = str(product_id)

            # ========================================
            # 💰 CALCUL TVA SELON TYPE PRODUIT
            # ========================================
            # Riz/Céréales: prix catalogue = HT → ajouter TVA
            # Autres: prix catalogue = TTC → extraire TVA

            if is_riz_cereales(prod_name):
                # Prix catalogue = HT
                pu_ht = catalog_price
                tva_unit = round(catalog_price * TVA_RATE, 0)
                pu_ttc = round(catalog_price * (1 + TVA_RATE), 0)
            else:
                # Prix catalogue = TTC → calculer HT
                pu_ttc = catalog_price
                pu_ht = round(catalog_price / (1 + TVA_RATE), 0)
                tva_unit = round(pu_ttc - pu_ht, 0)

            # Remplissage cellules
            row = current_row

            # Col 1: Réf
            c = ws.cell(row=row, column=1, value=product_id)
            c.border = border
            c.alignment = center_align
            c.font = bold_font

            # Col 2: Désignation
            c = ws.cell(row=row, column=2, value=clean_name[:50])
            c.border = border
            c.alignment = left_align

            # Col 3: Qté
            c = ws.cell(row=row, column=3, value=qty_final)
            c.border = border
            c.alignment = center_align
            c.font = bold_font

            # Col 4: Conditionnement
            cond_display = f"{conditionnement}" if conditionnement > 1 else "-"
            c = ws.cell(row=row, column=4, value=cond_display)
            c.border = border
            c.alignment = center_align
            c.font = Font(size=9, color="666666")

            # Col 5: Unité
            c = ws.cell(row=row, column=5, value=unite)
            c.border = border
            c.alignment = center_align

            # Col 6: PU HT (valeur calculée)
            c = ws.cell(row=row, column=6, value=pu_ht)
            c.border = border
            c.number_format = '#,##0'

            # Colonne 7 : TVA Unitaire (valeur calculée)
            cell = ws.cell(row=current_row, column=7, value=tva_unit)
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border
            cell.font = Font(italic=True, color="666666")

            # Colonne 8 : PU TTC (valeur calculée)
            cell = ws.cell(row=current_row, column=8, value=pu_ttc)
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border
            cell.font = Font(bold=True, color="0066CC")

            # Colonne 9 : Remise % (ÉDITABLE par produit)
            cell = ws.cell(row=current_row, column=9, value=0)
            cell.number_format = '0.00'
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = remise_fill
            cell.border = border
            cell.font = Font(bold=True)

            # ========================================
            # FORMULES AUTOMATIQUES
            # ========================================

            # Colonne 10 : Total HT = (PU HT × Qté) × (1-Remise/100)
            formula_ht = f"=(F{current_row}*C{current_row})*(1-I{current_row}/100)"
            cell = ws.cell(row=current_row, column=10, value=formula_ht)
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            # Colonne 11 : Total TTC = (PU TTC × Qté) × (1-Remise/100)
            formula_ttc = f"=(H{current_row}*C{current_row})*(1-I{current_row}/100)"
            cell = ws.cell(row=current_row, column=11, value=formula_ttc)
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.fill = total_fill
            cell.border = border
            cell.font = Font(bold=True)

            current_row += 1

        last_data_row = current_row - 1

        # ========================================
        # 💰 SOUS-TOTAL PAR FOURNISSEUR + ESCOMPTE
        # ========================================

        if last_data_row >= first_data_row:
            # Ligne Sous-total brut
            ws.merge_cells(f'A{current_row}:I{current_row}')
            cell = ws.cell(row=current_row, column=1, value=f"Sous-total {supplier.upper()}")
            cell.font = Font(bold=True, size=10)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            # Total HT brut (colonne 10)
            cell = ws.cell(row=current_row, column=10, value=f"=SUM(J{first_data_row}:J{last_data_row})")
            cell.number_format = '#,##0'
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            # TTC brut (colonne 11)
            cell = ws.cell(row=current_row, column=11, value=f"=SUM(K{first_data_row}:K{last_data_row})")
            cell.number_format = '#,##0'
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            subtotal_row = current_row
            current_row += 1

            # ========================================
            # 🎯 LIGNE ESCOMPTE FOURNISSEUR (ÉDITABLE)
            # ========================================
            ws.merge_cells(f'A{current_row}:H{current_row}')
            cell = ws.cell(row=current_row, column=1, value=f"🎯 Escompte {supplier.upper()}")
            cell.font = Font(bold=True, size=10, color="E65100")
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            # Cellule escompte % (ÉDITABLE - colonne I)
            escompte_cell = ws.cell(row=current_row, column=9, value=0)
            escompte_cell.number_format = '0.00"%"'
            escompte_cell.alignment = Alignment(horizontal='center', vertical='center')
            escompte_cell.fill = escompte_fill
            escompte_cell.border = border
            escompte_cell.font = Font(bold=True, size=11)

            # Montant escompte HT (formule - colonne 10)
            cell = ws.cell(row=current_row, column=10, value=f"=-J{subtotal_row}*I{current_row}/100")
            cell.number_format = '#,##0'
            cell.font = Font(bold=True, color="E65100")
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            # Montant escompte TTC (colonne 11)
            cell = ws.cell(row=current_row, column=11, value=f"=-K{subtotal_row}*I{current_row}/100")
            cell.number_format = '#,##0'
            cell.font = Font(bold=True, color="E65100")
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            escompte_row = current_row
            current_row += 1

            # ========================================
            # 💚 TOTAL NET FOURNISSEUR (après escompte)
            # ========================================
            ws.merge_cells(f'A{current_row}:I{current_row}')
            cell = ws.cell(row=current_row, column=1, value=f"TOTAL NET {supplier.upper()}")
            cell.font = Font(bold=True, size=11)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border
            cell.fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")

            # Total HT net (sous-total + escompte - colonne 10)
            cell = ws.cell(row=current_row, column=10, value=f"=J{subtotal_row}+J{escompte_row}")
            cell.number_format = '#,##0'
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border
            cell.fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")

            # TTC net (colonne 11)
            cell = ws.cell(row=current_row, column=11, value=f"=K{subtotal_row}+K{escompte_row}")
            cell.number_format = '#,##0'
            cell.font = Font(bold=True, size=11)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
            cell.border = border

            current_row += 1

        current_row += 2

    # ========================================
    # 💰 TOTAUX GLOBAUX
    # ========================================

    current_row += 1

    # Total HT (colonne J maintenant)
    ws[f'I{current_row}'] = "TOTAL HT :"
    ws[f'I{current_row}'].font = Font(bold=True, size=12)
    ws[f'I{current_row}'].alignment = Alignment(horizontal='right')
    ws[f'J{current_row}'] = '=SUMIF(A:A,"TOTAL NET*",J:J)'
    ws[f'J{current_row}'].number_format = '#,##0 "FCFA"'
    ws[f'J{current_row}'].font = Font(bold=True, size=12)
    ws[f'J{current_row}'].border = Border(bottom=Side(style='thin'))

    current_row += 1

    # TVA Totale (différence TTC - HT)
    ws[f'I{current_row}'] = "TVA (18%) :"
    ws[f'I{current_row}'].font = Font(bold=True, size=12)
    ws[f'I{current_row}'].alignment = Alignment(horizontal='right')
    ws[f'J{current_row}'] = f'=J{current_row + 1}-J{current_row - 1}'
    ws[f'J{current_row}'].number_format = '#,##0 "FCFA"'
    ws[f'J{current_row}'].font = Font(bold=True, size=12)
    ws[f'J{current_row}'].border = Border(bottom=Side(style='thin'))

    current_row += 1

    # Total TTC (colonne K maintenant)
    ws[f'I{current_row}'] = "TOTAL TTC :"
    ws[f'I{current_row}'].font = Font(bold=True, size=14, color="006400")
    ws[f'I{current_row}'].alignment = Alignment(horizontal='right')
    ws[f'J{current_row}'] = '=SUMIF(A:A,"TOTAL NET*",K:K)'
    ws[f'J{current_row}'].number_format = '#,##0 "FCFA"'
    ws[f'J{current_row}'].font = Font(bold=True, size=14, color="FFFFFF")
    ws[f'J{current_row}'].fill = PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid")
    ws[f'J{current_row}'].border = Border(
        top=Side(style='double'),
        bottom=Side(style='double')
    )

    # ========================================
    # ✍️ SIGNATURE EN BAS
    # ========================================

    current_row += 4  # Espace entre total et signature
    signature_row = current_row

    try:
        from openpyxl.drawing.image import Image as XLImage

        # Chercher la signature
        signature_paths = [
            "assets/signature.png",
            "assets/signature.jpg",
            "signature.png",
            "signature.jpg"
        ]

        signature_path = None
        for path in signature_paths:
            if os.path.exists(path):
                signature_path = path
                break

        if signature_path:
            # Ajouter la signature
            sig_img = XLImage(signature_path)

            # Redimensionner signature (largeur 150px)
            sig_img.width = 150
            sig_img.height = int(150 / sig_img.width * sig_img.height) if sig_img.width > 0 else 80

            # Positionner en bas à droite (colonne H)
            ws.add_image(sig_img, f'H{signature_row}')

            print(f"   ✅ Signature ajoutée dans Excel : {signature_path}")

            current_row += 6  # Espace pour la signature
        else:
            print("   ⚠️ Signature non trouvée - ajout cadre signature")

            # Alternative : Zone pour signature manuscrite
            ws.merge_cells(f'H{current_row}:J{current_row}')
            ws[f'H{current_row}'] = "Signature et Cachet"
            ws[f'H{current_row}'].font = Font(bold=True, size=11)
            ws[f'H{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'H{current_row}'].border = Border(
                bottom=Side(style='thin'),
                top=Side(style='thin'),
                left=Side(style='thin'),
                right=Side(style='thin')
            )
            ws.row_dimensions[current_row].height = 60

            current_row += 1

    except Exception as e:
        print(f"   ⚠️ Erreur ajout signature Excel : {e}")

        # Fallback : Ligne pour signature
        ws.merge_cells(f'H{current_row}:J{current_row}')
        ws[f'H{current_row}'] = "_" * 30
        ws[f'H{current_row}'].alignment = Alignment(horizontal='center')
        current_row += 1
        ws[f'H{current_row}'] = "Signature et Cachet"
        ws[f'H{current_row}'].font = Font(bold=True, size=10)
        ws[f'H{current_row}'].alignment = Alignment(horizontal='center')

    # ========================================
    # 📐 LARGEURS COLONNES
    # ========================================
    ws.column_dimensions['A'].width = 8      # Réf
    ws.column_dimensions['B'].width = 38     # Désignation
    ws.column_dimensions['C'].width = 7      # Qté
    ws.column_dimensions['D'].width = 6      # Cond.
    ws.column_dimensions['E'].width = 8      # Unité
    ws.column_dimensions['F'].width = 11     # PU HT
    ws.column_dimensions['G'].width = 10     # TVA Unit
    ws.column_dimensions['H'].width = 11     # PU TTC
    ws.column_dimensions['I'].width = 9      # Remise %
    ws.column_dimensions['J'].width = 13     # Total HT
    ws.column_dimensions['K'].width = 13     # Total TTC

    # ========================================
    # 💾 EXPORT
    # ========================================
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    gen_elapsed = time.time() - gen_start
    print(f"   ⚡ Excel généré en {gen_elapsed:.2f}s : BC-{po_number}")

    return buf, po_number

# ===== CALLBACK : GÉNÉRATION BON DE COMMANDE EXCEL AVEC VALIDATION =====
@app.callback(
    [Output("download-po", "data"),
     Output("po-loading-overlay", "is_open")],
    Input("btn-po-pdf", "n_clicks"),
    [State("main-table", "selected_rows"),
     State("main-table", "data"),
     State("auth-state", "data")],  # ✅ AJOUT: auth-state pour tracking
    prevent_initial_call=True
)
def export_po_excel_validated(n_clicks, selected_rows, table_data, auth_state):
    """
    🚀 VERSION ULTRA-RAPIDE - Génère un BC Excel en <2s
    Utilise le cache pré-chargé pour éviter les chargements réseau
    + Tracking des commandes par agent
    """
    callback_start = time.time()

    if not n_clicks or not table_data or not selected_rows:
        return no_update, False

    print(f"\n{'='*60}")
    print(f"📄 GÉNÉRATION BON DE COMMANDE EXCEL")
    print(f"{'='*60}")

    # Récupérer l'utilisateur connecté
    username = auth_state.get("username", "unknown") if auth_state else "unknown"

    # ========== VALIDATION STRICTE QAC ==========
    is_valid, error_msg, missing = valider_qac_selection(selected_rows, table_data)

    if not is_valid:
        print(f"❌ VALIDATION ÉCHOUÉE: {len(missing)} produits sans QAC")
        return no_update, False

    # ========== GÉNÉRATION RAPIDE ==========
    selected_products = [table_data[idx] for idx in selected_rows]
    print(f"📦 {len(selected_products)} produits sélectionnés par {username}")

    # Utiliser le cache pré-chargé (instantané)
    bc_cache = get_bc_cache()
    price_map = bc_cache.get("prices") or get_catalog_prices()
    packaging_map = bc_cache.get("packaging") or get_packaging_map_cached()

    try:
        # Génération Excel avec formules
        excel_buffer, po_number = generer_bc_excel_avec_formules(
            selected_products, price_map, packaging_map
        )

        fname = f"BC_{po_number}_{len(selected_products)}p.xlsx"

        # ========== TRACKING PAR FOURNISSEUR ==========
        # Calculer les stats par fournisseur pour le tracking
        suppliers_stats = {}
        for prod in selected_products:
            supplier = str(prod.get("Supplier", "N/A")).strip()
            if supplier not in suppliers_stats:
                suppliers_stats[supplier] = {"products": 0, "amount": 0}

            qac = safe_float(prod.get("QAC edited"), safe_float(prod.get("QAC"), 0))
            price = safe_float(price_map.get(str(prod.get("product_name", "")).lower(), 1000), 1000)

            suppliers_stats[supplier]["products"] += 1
            suppliers_stats[supplier]["amount"] += qac * price

        # Tracker chaque fournisseur séparément
        for supplier, stats in suppliers_stats.items():
            track_agent_order(
                username=username,
                supplier=supplier,
                order_total=stats["amount"],
                products_count=stats["products"],
                po_number=po_number
            )

        callback_elapsed = time.time() - callback_start
        print(f"⚡ TOTAL: {callback_elapsed:.2f}s - {fname}")
        print(f"   📊 Tracking: {len(suppliers_stats)} fournisseurs pour {username}")
        print(f"{'='*60}\n")

        return dcc.send_bytes(excel_buffer.read(), filename=fname), False

    except Exception as e:
        print(f"❌ Erreur: {e}")
        import traceback
        traceback.print_exc()
        return no_update, False


# ===== CALLBACK : AGENT IA - CALCULER QAC =====
# ===== CALLBACK 1 : TOGGLE BOUTON (VERSION CORRIGÉE) =====
@app.callback(
    Output("btn-po-pdf", "disabled"),
    [Input("main-table", "selected_rows"),
     Input("main-table", "data")],
    prevent_initial_call=True  # ✅ CORRIGÉ: main-table doit exister d'abord
)
def toggle_po_button(selected_rows, data):
    if not data or not selected_rows or len(selected_rows) == 0:
        return True

    for idx in selected_rows:
        if idx < 0 or idx >= len(data):
            continue

        row = data[idx]
        if not row:
            return True

        if not row.get("product_name") or not row.get("Supplier"):
            return True

        qac = safe_float(row.get("QAC edited", 0), 0)
        if qac <= 0:
            return True

    return False


# ===== CALLBACK 2 : AGENT IA (DÉSACTIVÉ - UTILISE LE CALLBACK PRINCIPAL) =====
# NOTE: Le callback principal run_agent_ia_calcul est défini plus haut et modifie main-table directement
'''
@app.callback(
    Output("master-data", "data", allow_duplicate=True),
    Input("btn-run-agent-ia", "n_clicks"),
    State("master-data", "data"),
    prevent_initial_call=True
)
def run_agent_ia_calcul_OLD(n_clicks, master_json):
    """
    Lance l'Agent IA et met à jour master-data
    → apply_filters se déclenchera automatiquement après
    """
    if not n_clicks or not master_json:
        return no_update

    print("\n" + "=" * 60)
    print("🤖 AGENT IA - CALCUL DES QAC")
    print("=" * 60)

    # Charger les données depuis master-data
    master_df = pd.DataFrame(json.loads(master_json))

    # Convertir en format attendu par l'agent
    table_data = master_df.to_dict('records')

    # Appel Agent IA
    qac_par_index = agent_ia_calculer_qac(table_data, use_gemini=True)

    if not qac_par_index:
        print("⚠️ Aucun produit à commander")
        print("=" * 60 + "\n")
        return no_update

    # Mise à jour du DataFrame
    for idx, qac_value in qac_par_index.items():
        if 0 <= idx < len(master_df):
            master_df.at[idx, 'QAC edited'] = qac_value

    print(f"✅ {len(qac_par_index)} QAC calculées et injectées dans master-data")
    print("=" * 60 + "\n")

    # Retourner master-data mis à jour
    # → apply_filters se déclenchera automatiquement et propagera à main-table
    return master_df.to_json(orient="records")
'''

# ===== CALLBACK : PRÉSÉLECTION APRÈS AGENT IA (DÉSACTIVÉ - INTÉGRÉ AU CALLBACK PRINCIPAL) =====
'''
@app.callback(
    Output("main-table", "selected_rows", allow_duplicate=True),
    [Input("main-table", "data"),
     Input("btn-run-agent-ia", "n_clicks")],
    State("btn-run-agent-ia", "n_clicks"),
    prevent_initial_call=True
)
def preselect_qac_rows(table_data, ia_clicks, ia_clicks_state):
    """
    Présélectionne automatiquement les lignes avec QAC edited > 0
    après que l'Agent IA ait calculé
    """
    # Vérifier que c'est l'Agent IA qui a déclenché
    ctx = dash.callback_context
    if not ctx.triggered:
        return no_update

    trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]

    # Ne présélectionner QUE si c'est après l'Agent IA
    if trigger_id != "btn-run-agent-ia":
        return no_update

    if not table_data:
        return no_update

    # Trouver les lignes avec QAC edited > 0
    indices_selection = []
    for idx, row in enumerate(table_data):
        qac = safe_float(row.get("QAC edited", 0), 0)
        if qac > 0:
            indices_selection.append(idx)

    if indices_selection:
        print(f"✅ Présélection automatique de {len(indices_selection)} lignes")
        return sorted(indices_selection)

    return no_update
'''

# ===== CALLBACK 3 : REMPLIR QAC =====
# ===== CALLBACK : REMPLIR QAC DEPUIS TARGET (CORRIGÉ) =====
@app.callback(
    Output("master-data", "data", allow_duplicate=True),
    Input("btn-fill-qac-target", "n_clicks"),
    [State("main-table", "selected_rows"),
     State("master-data", "data")],
    prevent_initial_call=True
)
def fill_qac_from_target(n_clicks, selected_rows, master_json):
    """
    Copie target_quantity → QAC edited
    """
    if not n_clicks or not selected_rows or not master_json:
        return no_update

    master_df = pd.DataFrame(json.loads(master_json))
    count = 0

    for idx in selected_rows:
        if 0 <= idx < len(master_df):
            target = safe_float(master_df.at[idx, "target_quantity"], 0)
            if target > 0:
                master_df.at[idx, "QAC edited"] = target
                count += 1

    print(f"✅ {count} QAC remplies depuis target_quantity")

    return master_df.to_json(orient="records")

# ====== EXPORT BON DE COMMANDE WORD ======
'''
# ===== Génération BON DE COMMANDE (DOCX) avec conversion QAC edited -> unités majeures =====
@app.callback(
    [Output("download-po", "data"),
     Output("po-loading-overlay", "is_open")],  # ✅ Feedback visuel
    Input("btn-po-pdf", "n_clicks"),
    [State("main-table", "selected_rows"),
     State("main-table", "data")],
    prevent_initial_call=True
)
def export_po_word_optimized(n_clicks, selected_rows, table_data):
    """
    Génère un bon de commande WORD (.docx) OPTIMISÉ (rapide).
    Qté = QAC edited convertie en unités MAJEURES (ceil).
    """
    if not DOCX_AVAILABLE:
        print(f"❌ python-docx indisponible: {DOCX_IMPORT_ERROR}")
        return no_update, False

    if not n_clicks or not table_data or not selected_rows or len(selected_rows) == 0:
        return no_update, False

    print("\n" + "=" * 60)
    print("📄 GÉNÉRATION BON DE COMMANDE")
    print("=" * 60)

    # ========== 1. DONNÉES SÉLECTIONNÉES ==========
    selected_products = [table_data[idx] for idx in selected_rows if idx < len(table_data)]
    if not selected_products:
        return no_update, False

    print(f"📦 Produits sélectionnés : {len(selected_products)}")

    # ========== 2. CHARGEMENT CACHE (RAPIDE) ==========
    price_map = get_catalog_prices()  # ✅ Cache Redis
    packaging_map = get_packaging_map_cached()  # ✅ Cache Redis

    # ========== 3. REGROUPEMENT PAR FOURNISSEUR ==========
    suppliers = {}
    for prod in selected_products:
        sup = str(prod.get("Supplier", "")).strip()
        if not sup or sup.lower() == "nan":
            sup = "Fournisseur non spécifié"
        suppliers.setdefault(sup, []).append(prod)

    print(f"🏭 Fournisseurs : {len(suppliers)}")

    # ========== 4. CRÉATION DOCX RAPIDE ==========
    doc = Document()

    # Marges réduites (gain de temps)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # Logo (optionnel, rapide si existe)
    logo_path = Path("logo_maad.jpg")
    if logo_path.exists():
        try:
            doc.add_picture(str(logo_path), width=Inches(1.2))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.LEFT
        except Exception as e:
            print(f"Warning: Could not add logo: {e}")

    # En-tête
    po_number = get_next_po_number()
    title = doc.add_heading(f'BON DE COMMANDE N° {po_number}', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].font.color.rgb = RGBColor(0, 51, 102)

    company_info = doc.add_paragraph()
    company_info.add_run(f"{COMPANY_NAME}\n").bold = True
    if COMPANY_ADDRESS: company_info.add_run(f"{COMPANY_ADDRESS}\n")
    if COMPANY_PHONE: company_info.add_run(f"Tél : {COMPANY_PHONE}\n")
    if COMPANY_EMAIL: company_info.add_run(f"Email : {COMPANY_EMAIL}\n")

    doc.add_paragraph().add_run(f"Date : {datetime.now().strftime('%d/%m/%Y')}").bold = True
    doc.add_paragraph()

    # ========== 5. TOTAUX GLOBAUX ==========
    TVA_RATE = 0.18
    total_ht_global = 0.0
    total_tva_global = 0.0

    # ========== 6. TABLEAU PAR FOURNISSEUR (OPTIMISÉ) ==========
    for supplier, products in suppliers.items():
        doc.add_heading(f'📦 Fournisseur : {supplier}', level=2).runs[0].font.color.rgb = RGBColor(34, 139, 34)

        # ✅ Tableau simplifié (moins de formatage = plus rapide)
        table = doc.add_table(rows=1, cols=8)
        table.style = 'Light Grid Accent 1'

        headers = ['Réf.', 'Désignation', 'Qté', 'PU HT', 'Total HT', 'Remise (%)', 'TVA 18%', 'Total TTC']
        hdr_cells = table.rows[0].cells
        for i, header in enumerate(headers):
            hdr_cells[i].text = header
            for paragraph in hdr_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                    run.font.size = Pt(10)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # ✅ Largeurs fixes (pas de calcul dynamique)
        widths = [Inches(0.6), Inches(2.5), Inches(0.5), Inches(0.8), Inches(0.9), Inches(0.8), Inches(0.8),
                  Inches(1.0)]
        for i, width in enumerate(widths):
            for cell in table.columns[i].cells:
                cell.width = width

        total_ht_supplier = 0.0
        total_tva_supplier = 0.0

        # ========== 7. LIGNES PRODUITS (RAPIDE) ==========
        for prod in products:
            prod_name = str(prod.get("product_name", "")).strip()
            prod_key = prod_name.lower()

            # Quantité (QAC edited prioritaire)
            qty_units = safe_float(prod.get("QAC edited"), 0.0)
            if qty_units <= 0:
                qty_units = safe_float(prod.get("target_quantity"), 0.0)
            if qty_units <= 0:
                qty_units = 1.0

            # Conversion unités majeures
            packaging_text = packaging_map.get(prod_key, "")
            qty_major = consolidate_to_major(qty_units, packaging_text, prod_name)

            # Prix & calculs
            unit_price = safe_float(price_map.get(prod_key, 1000.0), 1000.0)
            if unit_price <= 0:
                unit_price = 1000.0

            total_ht = qty_major * unit_price
            total_tva = total_ht * TVA_RATE
            total_ttc = total_ht + total_tva

            total_ht_supplier += total_ht
            total_tva_supplier += total_tva

            # ✅ Ajout ligne (formatage minimal)
            ref_code = ref_from_name(prod_name)
            row_cells = table.add_row().cells
            row_cells[0].text = ref_code
            row_cells[1].text = prod_name[:50]
            row_cells[2].text = f"{qty_major}"
            row_cells[3].text = f"{unit_price:,.0f}"
            row_cells[4].text = f"{total_ht:,.0f}"
            row_cells[5].text = ""
            row_cells[6].text = f"{total_tva:,.0f}"
            row_cells[7].text = f"{total_ttc:,.0f}"

            # Alignement rapide
            for i_col in [2, 3, 4, 5, 6, 7]:
                row_cells[i_col].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # ========== 8. SOUS-TOTAL FOURNISSEUR ==========
        subtotal_row = table.add_row().cells
        subtotal_row[0].merge(subtotal_row[3])
        subtotal_row[0].text = f"SOUS-TOTAL {supplier.upper()}"
        subtotal_row[0].paragraphs[0].runs[0].font.bold = True

        subtotal_row[4].text = f"{total_ht_supplier:,.0f}"
        subtotal_row[5].text = "-"
        subtotal_row[6].text = f"{total_tva_supplier:,.0f}"
        subtotal_row[7].text = f"{(total_ht_supplier + total_tva_supplier):,.0f}"

        for i in [4, 6, 7]:
            subtotal_row[i].paragraphs[0].runs[0].font.bold = True
            subtotal_row[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        total_ht_global += total_ht_supplier
        total_tva_global += total_tva_supplier

        doc.add_paragraph()

    # ========== 9. TOTAUX GLOBAUX ==========
    total_ttc_global = total_ht_global + total_tva_global

    doc.add_paragraph()
    totals_para = doc.add_paragraph()
    totals_para.add_run(f"TOTAL HT : {total_ht_global:,.0f} FCFA\n").bold = True
    totals_para.add_run(f"TVA (18%) : {total_tva_global:,.0f} FCFA\n").bold = True
    ttc_run = totals_para.add_run(f"TOTAL TTC : {total_ttc_global:,.0f} FCFA")
    ttc_run.bold = True
    ttc_run.font.size = Pt(14)
    ttc_run.font.color.rgb = RGBColor(0, 102, 204)

    # ========== 10. NOTES (MINIMAL) ==========
    doc.add_paragraph()
    notes_heading = doc.add_heading('Notes importantes :', level=3)
    notes_list = doc.add_paragraph(style='List Bullet')
    notes_list.add_run("La colonne 'Remise (%)' est à remplir manuellement\n")
    notes_list.add_run("Formule : Total TTC = (Total HT × (1 - Remise/100)) × 1.18")

    doc.add_paragraph()
    conditions = doc.add_paragraph()
    conditions.add_run("Conditions : ").bold = True
    conditions.add_run("Selon termes contractuels")

    # Pied de page
    doc.add_paragraph()
    footer_para = doc.add_paragraph()
    footer_run = footer_para.add_run(f"Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}")
    footer_run.font.size = Pt(8)
    footer_run.font.color.rgb = RGBColor(128, 128, 128)
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ========== 11. EXPORT ==========
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    fname = f"BC_{po_number}_{len(selected_products)}p.docx"

    print("=" * 60)
    print(f"✅ Document généré : {fname}")
    print(f"   - Produits : {len(selected_products)}")
    print(f"   - Fournisseurs : {len(suppliers)}")
    print(f"   - Total TTC : {total_ttc_global:,.0f} FCFA")
    print("=" * 60 + "\n")

    # ========== 12. TRACKING SUPABASE ==========
    try:
        # Récupérer l'utilisateur actif
        for username, session_info in ACTIVE_SESSIONS.items():
            user_id = session_info.get("user_id")
            session_id = session_info.get("session_id")
            if user_id and session_id:
                # Tracker chaque fournisseur
                for supplier_name, supplier_products in suppliers.items():
                    supplier_total = sum(
                        safe_float(p.get("QAC edited", 0), 0) * safe_float(price_map.get(str(p.get("product_name", "")).lower(), 1000), 1000)
                        for p in supplier_products
                    )
                    track_po_generated(
                        user_id=user_id,
                        session_id=session_id,
                        supplier=supplier_name,
                        products_count=len(supplier_products),
                        total_amount=supplier_total * 1.18,
                        po_type="docx"
                    )
                break
    except Exception as e:
        print(f"⚠️ Erreur tracking PO: {e}")

    return dcc.send_bytes(buf.read(), filename=fname), False  # ✅ Fermer overlay
'''


'''
@app.callback(
    Output("download-po", "data"),
    Input("btn-po-pdf", "n_clicks"),
    [State("main-table", "selected_rows"),
     State("main-table", "data")],
    prevent_initial_call=True
)
def export_po_word(n_clicks, selected_rows, table_data):
    """
    Génère un bon de commande WORD (.docx) pour PLUSIEURS produits sélectionnés.

    IMPORTANT : Utilise la QAC ÉDITÉE (colonne QAC du tableau, pas target_quantity).
    Le document Word permet d'ajouter les remises manuellement après génération.
    """
    if not DOCX_AVAILABLE:
        # Log clair + stopper proprement cette feature
        print(f"❌ python-docx indisponible: {DOCX_IMPORT_ERROR}")
        raise RuntimeError("La génération Word est indisponible sur cet environnement.")
    if not n_clicks or not table_data or not selected_rows or len(selected_rows) == 0:
        return no_update

    # ========================================
    # 1. RÉCUPÉRER LES PRODUITS SÉLECTIONNÉS
    # ========================================
    selected_products = [table_data[idx] for idx in selected_rows if idx < len(table_data)]

    if not selected_products:
        return no_update

    print(f"\n{'=' * 60}")
    print(f"📄 GÉNÉRATION BON DE COMMANDE WORD")
    print(f"{'=' * 60}")
    print(f"Produits sélectionnés : {len(selected_products)}")

    # ========================================
    # 2. CHARGER LE CATALOGUE POUR LES PRIX
    # ========================================
    try:
        cat_url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTrpcAiktxAPBiwznGOh35kVetc4O8-z5rQdFDgBaDE4OC3Jnb7JDGm59c55Cwm2pWCktcsBirWT_0b/pub?gid=751531326&single=true&output=csv"
        catalog = pd.read_csv(cat_url, skiprows=1)

        # Colonnes : A=id, B=name, D=selling_price, K=purchase_price
        catalog_subset = catalog.iloc[:, [0, 1, 3, 10]].copy()
        catalog_subset.columns = ['product_id', 'product_name', 'selling_price', 'purchase_price']

        catalog_subset['product_name_clean'] = (
            catalog_subset['product_name']
            .astype(str)
            .str.lower()
            .str.strip()
        )

        # Mapping produit → prix d'achat
        price_map = dict(zip(
            catalog_subset['product_name_clean'],
            pd.to_numeric(catalog_subset['purchase_price'], errors='coerce').fillna(1000)
        ))

        print(f"✅ Catalogue chargé : {len(price_map)} produits avec prix")

    except Exception as e:
        print(f"❌ Erreur chargement catalogue : {e}")
        price_map = {}

    # ========================================
    # 3. REGROUPER PAR FOURNISSEUR
    # ========================================
    suppliers = {}
    for prod in selected_products:
        supplier = str(prod.get("Supplier", "")).strip()
        if not supplier or supplier.lower() == 'nan':
            supplier = "Fournisseur non spécifié"

        if supplier not in suppliers:
            suppliers[supplier] = []
        suppliers[supplier].append(prod)

    print(f"✅ Fournisseurs détectés : {len(suppliers)}")
    for sup, prods in suppliers.items():
        print(f"   - {sup}: {len(prods)} produit(s)")

    # ========================================
    # 4. CRÉER LE DOCUMENT WORD
    # ========================================

    doc = Document()

    # Configuration des marges (en inches)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # ========================================
    # 5. EN-TÊTE DU DOCUMENT
    # ========================================

    # Logo (si disponible)
    logo_path = Path("logo_maad.jpg")
    if logo_path.exists():
        try:
            doc.add_picture(str(logo_path), width=Inches(1.2))
            last_paragraph = doc.paragraphs[-1]
            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        except Exception as e:
            print(f"⚠️ Logo non ajouté : {e}")

    # Titre
    po_number = get_next_po_number()
    title = doc.add_heading(f'BON DE COMMANDE N° {po_number}', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.runs[0]
    title_run.font.color.rgb = RGBColor(0, 51, 102)  # Bleu foncé

    # Informations entreprise
    company_info = doc.add_paragraph()
    company_info.add_run(f"{COMPANY_NAME}\n").bold = True
    if COMPANY_ADDRESS:
        company_info.add_run(f"{COMPANY_ADDRESS}\n")
    if COMPANY_PHONE:
        company_info.add_run(f"Tél : {COMPANY_PHONE}\n")
    if COMPANY_EMAIL:
        company_info.add_run(f"Email : {COMPANY_EMAIL}\n")

    # Date
    date_para = doc.add_paragraph()
    date_para.add_run(f"Date : {datetime.now().strftime('%d/%m/%Y')}").bold = True

    # Résumé
    summary = doc.add_paragraph()
    summary.add_run(f"Nombre de produits : {len(selected_products)} • ").bold = True
    summary.add_run(f"Fournisseurs : {len(suppliers)}").bold = True

    doc.add_paragraph()  # Ligne vide

    # ========================================
    # 6. CONSTANTES TVA
    # ========================================
    TVA_RATE = 0.18
    total_ht_global = 0
    total_tva_global = 0

    # ========================================
    # 7. CRÉER TABLEAU PAR FOURNISSEUR
    # ========================================

    for supplier, products in suppliers.items():

        # Titre fournisseur
        supplier_heading = doc.add_heading(f'📦 Fournisseur : {supplier}', level=2)
        supplier_run = supplier_heading.runs[0]
        supplier_run.font.color.rgb = RGBColor(34, 139, 34)  # Vert

        # Créer tableau (8 colonnes : Réf, Produit, Qté, PU HT, Total HT, Remise, TVA, Total TTC)
        table = doc.add_table(rows=1, cols=8)
        table.style = 'Light Grid Accent 1'
        table.autofit = False
        table.allow_autofit = False

        # En-têtes
        headers = ['Réf.', 'Désignation', 'Qté', 'PU HT', 'Total HT', 'Remise (%)', 'TVA 18%', 'Total TTC']
        hdr_cells = table.rows[0].cells

        for i, header in enumerate(headers):
            hdr_cells[i].text = header
            # Style header
            for paragraph in hdr_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                    run.font.size = Pt(10)
                    run.font.color.rgb = RGBColor(255, 255, 255)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Couleur de fond header (bleu)
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), '4472C4')
            hdr_cells[i]._element.get_or_add_tcPr().append(shading_elm)

        # Largeurs de colonnes (en inches)
        widths = [Inches(0.6), Inches(2.5), Inches(0.5), Inches(0.8), Inches(0.9), Inches(0.8), Inches(0.8),
                  Inches(1.0)]
        for i, width in enumerate(widths):
            for cell in table.columns[i].cells:
                cell.width = width

        # ========================================
        # 8. REMPLIR LES LIGNES DE PRODUITS
        # ========================================

        total_ht_supplier = 0
        total_tva_supplier = 0

        for prod in products:
            prod_name = str(prod.get("product_name", "")).strip()

            # ✅ ✅ ✅ UTILISER QAC ÉDITÉE (priorité absolue) ✅ ✅ ✅
            qty = float(prod.get("QAC", 0))

            # Si QAC est vide ou 0, fallback sur target_quantity
            if qty <= 0:
                qty = float(prod.get("target_quantity", 0))

            # Si toujours 0, mettre 1 par défaut
            if qty <= 0:
                qty = 1

            # Chercher prix dans le catalogue
            unit_price = price_map.get(prod_name.lower(), 1000.0)
            if unit_price <= 0:
                unit_price = 1000.0

            # Calculs (SANS remise pour l'instant)
            total_ht = qty * unit_price
            total_tva = total_ht * TVA_RATE
            total_ttc = total_ht + total_tva

            total_ht_supplier += total_ht
            total_tva_supplier += total_tva

            # Référence produit
            ref_code = ref_from_name(prod_name)

            # Ajouter ligne
            row_cells = table.add_row().cells

            row_cells[0].text = ref_code
            row_cells[1].text = prod_name[:50]  # Limiter longueur
            row_cells[2].text = f"{qty:.0f}"
            row_cells[3].text = f"{unit_price:,.0f}"
            row_cells[4].text = f"{total_ht:,.0f}"
            row_cells[5].text = ""  # ✅ REMISE VIDE (à remplir manuellement)
            row_cells[6].text = f"{total_tva:,.0f}"
            row_cells[7].text = f"{total_ttc:,.0f}"

            # Alignement
            for i in [2, 3, 4, 5, 6, 7]:  # Colonnes numériques
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

            # Taille de police
            for cell in row_cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(9)

        # ========================================
        # 9. SOUS-TOTAL PAR FOURNISSEUR
        # ========================================

        subtotal_row = table.add_row().cells
        subtotal_row[0].merge(subtotal_row[3])
        subtotal_row[0].text = f"SOUS-TOTAL {supplier.upper()}"
        subtotal_row[0].paragraphs[0].runs[0].font.bold = True

        subtotal_row[4].text = f"{total_ht_supplier:,.0f}"
        subtotal_row[5].text = "-"
        subtotal_row[6].text = f"{total_tva_supplier:,.0f}"
        subtotal_row[7].text = f"{(total_ht_supplier + total_tva_supplier):,.0f}"

        # Style sous-total
        for i in [4, 6, 7]:
            subtotal_row[i].paragraphs[0].runs[0].font.bold = True
            subtotal_row[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # Couleur de fond sous-total (gris clair)
        for cell in subtotal_row:
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), 'E7E6E6')
            cell._element.get_or_add_tcPr().append(shading_elm)

        total_ht_global += total_ht_supplier
        total_tva_global += total_tva_supplier

        doc.add_paragraph()  # Espacement

    # ========================================
    # 10. TOTAUX GLOBAUX
    # ========================================

    total_ttc_global = total_ht_global + total_tva_global

    doc.add_paragraph()  # Ligne vide

    totals_para = doc.add_paragraph()
    totals_para.add_run(f"TOTAL HT : {total_ht_global:,.0f} FCFA\n").bold = True
    totals_para.add_run(f"TVA (18%) : {total_tva_global:,.0f} FCFA\n").bold = True

    ttc_run = totals_para.add_run(f"TOTAL TTC : {total_ttc_global:,.0f} FCFA")
    ttc_run.bold = True
    ttc_run.font.size = Pt(14)
    ttc_run.font.color.rgb = RGBColor(0, 102, 204)

    # ========================================
    # 11. NOTES ET CONDITIONS
    # ========================================

    doc.add_paragraph()

    notes_heading = doc.add_heading('Notes importantes :', level=3)
    notes_list = doc.add_paragraph(style='List Bullet')
    notes_list.add_run("La colonne 'Remise (%)' est à remplir manuellement selon négociations\n")
    notes_list.add_run("Les totaux seront recalculés après application des remises\n")
    notes_list.add_run("Formule : Total TTC = (Total HT × (1 - Remise/100)) × 1.18")

    doc.add_paragraph()

    conditions = doc.add_paragraph()
    conditions.add_run("Conditions de livraison : ").bold = True
    conditions.add_run("À convenir avec les fournisseurs\n")

    conditions.add_run("Modalités de paiement : ").bold = True
    conditions.add_run("Selon termes contractuels")

    # ========================================
    # 12. PIED DE PAGE
    # ========================================

    doc.add_paragraph()
    footer_para = doc.add_paragraph()
    footer_run = footer_para.add_run(f"Document généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}")
    footer_run.font.size = Pt(8)
    footer_run.font.color.rgb = RGBColor(128, 128, 128)
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ========================================
    # 13. SAUVEGARDER ET ENVOYER
    # ========================================

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    fname = f"bon_commande_{po_number}_{len(selected_products)}_produits.docx"

    print(f"✅ Document Word généré : {fname}")
    print(f"   - {len(selected_products)} produits")
    print(f"   - {len(suppliers)} fournisseur(s)")
    print(f"   - Total HT : {total_ht_global:,.0f} FCFA")
    print(f"   - Total TTC : {total_ttc_global:,.0f} FCFA")
    print(f"{'=' * 60}\n")

    return dcc.send_bytes(buf.read(), filename=fname)


# Activer le bouton PO si au moins une ligne valide est sélectionnée
@app.callback(
    Output("btn-po-pdf", "disabled"),
    Input("main-table", "selected_rows"),     # <-- seul déclencheur
    State("main-table", "data"),              # <-- lecture, ne déclenche pas
    prevent_initial_call=True
)
def toggle_po_button(selected_rows, data):
    # Désactiver si pas de données
    if not data:
        return True

    # Désactiver si aucune ligne sélectionnée
    if not selected_rows:
        return True

    # Activer si AU MOINS une ligne sélectionnée possède product_name + Supplier
    for idx in selected_rows:
        if 0 <= idx < len(data):
            row = data[idx] or {}
            if row.get("product_name") and row.get("Supplier"):
                return False  # bouton activé

    # Sinon, désactiver
    return True
'''
# ==================== CALLBACK 4 : COMPTEUR DE SÉLECTION ====================
@app.callback(
    Output("selection-counter", "children"),
    Input("main-table", "selected_rows"),
    prevent_initial_call=True  # ✅ CORRIGÉ: main-table doit exister d'abord
)
def update_selection_counter(selected_rows):
    """Affiche le nombre de lignes sélectionnées"""
    count = len(selected_rows) if selected_rows else 0

    if count == 0:
        return html.Div([
            html.Small("Aucune sélection", style={"color": "#6b7280", "fontSize": "12px"})
        ])

    return dbc.Badge(
        f"☑️ {count} produit{'s' if count > 1 else ''} sélectionné{'s' if count > 1 else ''}",
        color="primary",
        pill=True,
        style={"fontSize": "13px", "padding": "8px 12px"}
    )


# NOTE: Callback select_all_rows désactivé - remplacé par toggle_select_all plus haut
'''
@app.callback(
    Output("main-table", "selected_rows", allow_duplicate=True),
    Input("btn-select-all", "n_clicks"),
    State("main-table", "data"),
    prevent_initial_call=True
)
def select_all_rows(n_clicks, table_data):
    """
    Sélectionne toutes les lignes actuellement affichées dans le DataTable (vue filtrée).
    """
    if not n_clicks:
        return no_update
    if not table_data:
        return []

    # Indices 0..N-1 de la vue affichée
    return list(range(len(table_data)))
'''

@app.callback(
    Output("main-table", "selected_rows", allow_duplicate=True),
    Input("btn-clear-selection", "n_clicks"),
    prevent_initial_call=True
)
def clear_selection(n_clicks):
    """Vide la sélection."""
    if not n_clicks:
        return no_update
    return []

# ------------------------------ Edit/Add/Delete rows -----------------------------
# NOTE: Ce callback est désactivé car remplacé par open_add_product_modal
'''
@app.callback(
   # Output("edit-modal", "is_open"),
    #Output("edit-input", "value"),
    #Output("edit-supplier", "value"),
    Output("edit-category", "value"),
    Output("edit-stock", "value"),
    Output("main-table", "active_cell"),
    Input("btn-add-row", "n_clicks"),
    Input("main-table", "active_cell"),
    State("main-table", "data"),
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
        if col == "edit Edit" and data and 0 <= row < len(data):
            r = data[row]
            return True, r.get("product_name", ""), r.get("Supplier", ""), r.get("Product Category", ""), r.get(
                "total_stock", 0), None
    return dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update
'''

import json

@app.callback(
    [Output("master-data", "data", allow_duplicate=True),
     Output("filtered-data", "data", allow_duplicate=True),
     Output("main-table", "data", allow_duplicate=True),
     Output("main-table", "selected_rows", allow_duplicate=True),
     Output("edit-modal", "is_open", allow_duplicate=True),
     Output("action-feedback", "children", allow_duplicate=True),
     Output("adding-new-product-flag", "data", allow_duplicate=True)],
    Input("edit-modal-save", "n_clicks"),
    State("edit-product-name", "value"),
    State("edit-supplier", "value"),
    State("edit-category", "value"),
    State("edit-stock", "value"),
    State("main-table", "active_cell"),
    State("main-table", "data"),
    State("search-input", "value"),
    State("filter-supplier", "value"),
    State("filter-category", "value"),
    State("toggle-options", "value"),
    State("master-data", "data"),
    State("adding-new-product-flag", "data"),
    State("auth-state", "data"),
    prevent_initial_call=True
)
def save_edit(n_clicks, prod, sup, cat, stock, active_cell, table_data, q, fs, fc, filter_opts, master_json, is_adding, auth_state):
    """Sauvegarde les modifications ou ajoute un nouveau produit avec intégration Supabase"""
    global initial_df

    if not n_clicks:
        return no_update, no_update, no_update, no_update, no_update, no_update, no_update

    # Validation
    if not prod or not prod.strip():
        feedback = dbc.Alert("⚠️ Le nom du produit est requis", color="warning", duration=3000)
        return no_update, no_update, no_update, no_update, True, feedback, no_update

    # Récupérer l'utilisateur actif
    username = auth_state.get("username", "") if auth_state else ""
    user_id = None
    session_id = None
    if username and username in ACTIVE_SESSIONS:
        session_info = ACTIVE_SESSIONS[username]
        user_id = session_info.get("user_id")
        session_id = session_info.get("session_id")

    # If master_json is already a list (not a JSON string), use it directly
    if isinstance(master_json, str):
        base = pd.DataFrame(json.loads(master_json)) if master_json else get_df_cached()
    else:
        base = pd.DataFrame(master_json) if master_json else get_df_cached()

    df = base.copy()

    # Mode ajout ou édition
    if is_adding:
        # ✅ AJOUT D'UN NOUVEAU PRODUIT
        print(f"\n{'='*60}")
        print(f"➕ AJOUT D'UN NOUVEAU PRODUIT")
        print(f"{'='*60}")

        new_row = {c: np.nan for c in df.columns}
        new_row["product_name"] = prod.strip()
        new_row["Supplier"] = sup.strip() if sup else ""
        new_row["Product Category"] = cat.strip() if cat else ""
        new_row["QAC edited"] = " "

        try:
            new_row["total_stock"] = float(stock or 0)
        except (ValueError, TypeError):
            new_row["total_stock"] = 0.0

        # Valeurs par défaut pour les colonnes numériques
        numeric_defaults = {
            "Average Daily Sales": 0.0,
            "optimal stock ": 0.0,
            "Max Lead Time": 7.0,
            "Max Avg Daily Sales": 0.0,
            "Max Coverage Day": 30.0,
            "Daily OOS Rate (30d)": 0.0,
            "Predicted Order Quantity": 0.0,
            "QAC": 0.0,
            "Avg Lead Time": 5.0,
            "Purchase Need": 0.0,
            "days_since_reception": None,
            "last_reception_qty": 0,
        }

        for col, default_val in numeric_defaults.items():
            if col in df.columns:
                new_row[col] = default_val

        # Status basé sur le stock
        if "Stock Status" in df.columns:
            if new_row["total_stock"] <= 0:
                new_row["Stock Status"] = "Out of Stock"
            elif new_row["total_stock"] < 10:
                new_row["Stock Status"] = "Order Soon"
            else:
                new_row["Stock Status"] = "In Stock"

        if "Ajusted_total_need" in df.columns:
            new_row["Ajusted_total_need"] = "ORDER NOW" if new_row["total_stock"] <= 0 else "STOCK OK"

        # Ajouter au DataFrame
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

        print(f"   📦 Produit: {prod}")
        print(f"   🏭 Fournisseur: {sup}")
        print(f"   🏷️ Catégorie: {cat}")
        print(f"   📊 Stock initial: {stock}")

        # ✅ SAUVEGARDER DANS SUPABASE
        if user_id and session_id:
            # Sauvegarder le produit ajouté
            save_added_product(user_id, session_id, {
                "product_name": prod.strip(),
                "supplier": sup.strip() if sup else "",
                "category": cat.strip() if cat else "",
                "stock": float(stock or 0),
                "added_by": username
            })

            # Tracker l'action
            track_product_action(user_id, session_id, "added", prod.strip(), {
                "supplier": sup,
                "category": cat,
                "initial_stock": float(stock or 0)
            })
            print(f"   ✅ Sauvegardé dans Supabase")

        print(f"{'='*60}\n")

        feedback = dbc.Alert([
            html.I(className="fas fa-check-circle me-2"),
            f"Produit '{prod}' ajouté avec succès"
        ], color="success", duration=4000)

    elif active_cell and active_cell.get("column_id") == "edit Edit" and active_cell.get("row") is not None and table_data:
        # ✅ ÉDITION D'UN PRODUIT EXISTANT
        row = active_cell["row"]
        r = table_data[row]
        key_p = r.get("product_name")
        key_s = r.get("Supplier")
        idx = df[(df["product_name"].astype(str) == str(key_p)) & (df["Supplier"].astype(str) == str(key_s))].index

        if len(idx) > 0:
            i = idx[0]
            old_values = {
                "product_name": df.at[i, "product_name"],
                "Supplier": df.at[i, "Supplier"],
                "Product Category": df.at[i, "Product Category"] if "Product Category" in df.columns else "",
                "total_stock": df.at[i, "total_stock"] if "total_stock" in df.columns else 0
            }

            if prod is not None: df.at[i, "product_name"] = prod.strip()
            if sup is not None: df.at[i, "Supplier"] = sup.strip()
            if cat is not None: df.at[i, "Product Category"] = cat.strip()
            if stock is not None:
                try:
                    df.at[i, "total_stock"] = float(stock)
                except (ValueError, TypeError):
                    pass

            # Tracker la modification
            if user_id and session_id:
                track_product_action(user_id, session_id, "edited", prod.strip(), {
                    "old_values": old_values,
                    "new_values": {
                        "product_name": prod,
                        "supplier": sup,
                        "category": cat,
                        "stock": stock
                    }
                })

        feedback = dbc.Alert([
            html.I(className="fas fa-edit me-2"),
            f"Produit '{prod}' modifié avec succès"
        ], color="info", duration=4000)
        print(f"✏️ Produit modifié: {prod} ({sup})")
    else:
        # Cas où on n'est ni en ajout ni en édition valide
        feedback = dbc.Alert("⚠️ Aucune action effectuée", color="warning", duration=3000)
        return no_update, no_update, no_update, no_update, False, feedback, False

    # ✅ Mettre à jour le DataFrame global (thread-safe)
    with _data_lock:
        initial_df = df.copy()

    # Application des filtres
    sup_list = fs or []
    stat_list = []
    cat_list = fc or []
    options = filter_opts or []

    if not df.empty:
        fdf = filter_dataframe(df, q, sup_list, stat_list, cat_list, options)
    else:
        fdf = pd.DataFrame()

    # Application des actions sur fdf
    fdf_actions = add_action_cols(fdf)

    print(f"📊 Table mise à jour: {len(fdf_actions)} produits affichés (total: {len(df)})")

    # Retour des résultats (fermer le modal et reset le flag)
    return df.to_json(orient="records"), fdf_actions.to_json(orient="records"), fdf_actions.to_dict("records"), [], False, feedback, False

'''
@app.callback(
    [Output("filtered-data", "data", allow_duplicate=True),
     Output("main-table", "data", allow_duplicate=True),
     Output("main-table", "selected_rows", allow_duplicate=True),
     Output("risk-banner", "children", allow_duplicate=True)],  # Allow duplicate outputs
    Input("master-data", "data"),
    prevent_initial_call=True
)
def delete_row(master_json):
    """Function to delete row based on master_json"""

    # Check if master_json is None or empty
    if not master_json:
        # If master_json is None, fallback to an empty DataFrame or a cached DataFrame
        print("master_json is None or empty. Using fallback data.")
        base = get_df_cached()  # Replace with appropriate fallback DataFrame function
    else:
        try:
            # If master_json is a valid JSON string, parse it; if it's already a list, use it directly
            base = pd.DataFrame(json.loads(master_json)) if isinstance(master_json, str) else pd.DataFrame(master_json)
        except ValueError as e:
            # Fallback if JSON is invalid
            print(f"Error parsing master_json: {e}. Using fallback data.")
            base = get_df_cached()  # Fallback if JSON is invalid

    # Process the base DataFrame (you can add more processing logic here)
    df = base.copy()

    # Example of filtering and updating
    # Apply any necessary filters based on your actual requirements
    banner = "Updated successfully"  # Set your banner message here

    # Return the 4 outputs required by the callback:
    # filtered-data, main-table, selected_rows (empty list), and risk-banner
    return df.to_json(orient="records"), df.to_dict("records"), [], banner
'''
# ------------------------------ Floating Chat callbacks ---------------------------
@app.callback(
    Output("chat-open", "data"),
    [Input("chat-fab", "n_clicks"), Input("chat-close", "n_clicks")],
    State("chat-open", "data"),
    prevent_initial_call=False
)
def toggle_chat(n_fab, n_close, is_open):
    is_open = bool(is_open)
    ctx = dash.callback_context
    if not ctx.triggered:
        return is_open
    trig = ctx.triggered[0]["prop_id"].split(".")[0]
    if trig == "chat-close":
        return False
    if trig == "chat-fab":
        return not is_open
    return is_open


@app.callback(
    Output("chat-window", "style"),
    Input("chat-open", "data")
)
def show_hide_chat(is_open):
    return {"display": "flex" if is_open else "none"}


def parse_contents(content):
    import base64
    content_type, content_string = content.split(',')
    decoded = base64.b64decode(content_string)
    try:
        df = pd.read_csv(io.BytesIO(decoded));
        return df
    except Exception:
        try:
            df = pd.read_excel(io.BytesIO(decoded));
            return df
        except Exception:
            return pd.DataFrame()


@app.callback(
    Output("uploaded-csv", "data"),
    Output("upload-status", "children"),
    Output("chat-messages", "children", allow_duplicate=True),
    Output("chat-store", "data", allow_duplicate=True),
    Input("chat-upload", "contents"),
    State("chat-upload", "filename"),
    State("chat-store", "data"),
    State("chat-messages", "children"),
    prevent_initial_call=True
)
def on_upload(contents, filename, history, rendered):
    if not contents:
        raise dash.exceptions.PreventUpdate

    # Validation du fichier
    try:
        df_u = parse_contents(contents)
        if df_u.empty:
            return no_update, f"⚠️ Échec de lecture du fichier {filename or ''}.", rendered, history
    except Exception as e:
        return no_update, f"❌ Erreur lors du téléchargement du fichier : {str(e)}", rendered, history

    # Processus d'upload réussi
    history = history or []
    msg = f"📥 Fichier chargé : **{filename}** — {len(df_u)} lignes détectées."
    history.append({"role": "assistant", "text": msg, "ts": datetime.now().isoformat()})
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
        return no_update, no_update, no_update  # Empêche un traitement inutile

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
    reply = chatbot_reply(user_text, df, history[:-1])  # Exclut le dernier message pour éviter doublon

    # Ajouter la réponse
    history.append({"role": "assistant", "text": reply, "ts": datetime.now().isoformat()})

    return _render_messages(history), history, ""  # ✅ Vider l'input


# ------------------------------ Run ------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(debug=True, host="0.0.0.0", port=port)
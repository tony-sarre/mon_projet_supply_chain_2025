# ========================= PREMIUM SUPPLY CHAIN DASHBOARD =========================
# Requirements:
# pip install dash==2.17.1 dash-bootstrap-components==1.6.0 plotly==5.22.0
# pip install pandas scikit-learn flask-caching numpy
# pip install reportlab supabase
# Optional: pip install openai
import csv
import os, sys
from functools import lru_cache
import hashlib  # ✅ NOUVEAU: Pour l'authentification

from dash.exceptions import PreventUpdate

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

# ✅ NOTE: L'authentification est intégrée directement dans ce fichier (voir section AUTH_USERS plus bas)
# Pas besoin d'importer depuis auth.py

import re
from dotenv import load_dotenv
import io
import base64
from datetime import datetime
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
warnings.filterwarnings("ignore", message="Parsing dates.*ambiguous", category=DeprecationWarning)

# Cache setup
cache = Cache(app.server, config={"CACHE_TYPE": "SimpleCache", "CACHE_DEFAULT_TIMEOUT": 3600})


@cache.memoize()  # Exemple de mise en cache pour la fonction
def get_df_cached():
    return load_supply_data()  # Fonction pour charger vos données

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
# 📊 FONCTIONS DE TRACKING SUPABASE
# ============================================================

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
    """Enregistre une activité utilisateur"""
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
    """Enregistre une modification QAC"""
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
            "at_risk_count": 0,  # Sera calculé plus tard
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
                    background: #22d3ee;
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
                    border-left: 4px solid #22d3ee;
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
                    background: #22d3ee;
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
    "Arame": {"name": "Arame Toure", "email": "arame.toure@maad.io"},
    "Coumba": {"name": "Coumba Cisse", "email": "ndeyecoumba.cisse@maad.io"},
    "Fallou": {"name": "Fallou Diop", "email": "serigne.diop@maad.io"},
    "Ravane": {"name": "Ravane Diop", "email": "pr.diop@maad.io"},
    "Insa": {"name": "Insa Niang", "email": "insa.niang@maad.io"},
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
        "role": "manager"
    },
    "insa": {
        "name": "Insa Niang",
        "email": "insa.niang@maad.io",
        "password_hash": DEFAULT_PASSWORD_HASH,
        "role": "user"
    },
}


def verify_password(username: str, password: str) -> bool:
    """Vérifie le mot de passe d'un utilisateur"""
    username = username.lower().strip()
    print(f"   🔍 DEBUG verify_password: username='{username}'")

    if username not in AUTH_USERS:
        print(f"   ❌ Utilisateur '{username}' non trouvé dans AUTH_USERS")
        print(f"   📋 Utilisateurs disponibles: {list(AUTH_USERS.keys())}")
        return False

    password_hash = hashlib.sha256(password.encode()).hexdigest()
    expected_hash = AUTH_USERS[username]["password_hash"]

    print(f"   🔐 Hash fourni: {password_hash[:20]}...")
    print(f"   🔐 Hash attendu: {expected_hash[:20]}...")

    if password_hash == expected_hash:
        print(f"   ✅ Mot de passe correct!")
        return True
    else:
        print(f"   ❌ Mot de passe incorrect")
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

/* ===== CARTE LOGIN ===== */
.login-card {
    background: rgba(30, 41, 59, 0.95);
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
    background: linear-gradient(135deg, #22d3ee 0%, #a78bfa 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    margin-bottom: 8px;
}

.login-subtitle {
    color: #94a3b8;
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

.login-input {
    width: 100%;
    padding: 16px 16px 16px 48px;
    background: rgba(15, 23, 42, 0.8);
    border: 2px solid #334155;
    border-radius: 12px;
    color: #f0f4f8;
    font-size: 15px;
    transition: all 0.3s ease;
}

.login-input:focus {
    outline: none;
    border-color: #22d3ee;
    box-shadow: 0 0 0 4px rgba(34, 211, 238, 0.15);
    background: rgba(15, 23, 42, 1);
}

.login-input::placeholder {
    color: #64748b;
}

/* ===== BOUTON ===== */
.login-button {
    width: 100%;
    padding: 16px;
    background: linear-gradient(135deg, #22d3ee 0%, #06b6d4 100%);
    border: none;
    border-radius: 12px;
    color: #0f172a;
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
    color: #fca5a5;
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
    background: linear-gradient(135deg, #22d3ee, #a78bfa);
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
    """Crée le layout de la page de connexion"""
    return html.Div([
        # Styles CSS injectés via une balise style dans un Iframe srcdoc ou via style inline
        # On utilise un Div avec dangerouslySetInnerHTML n'existe pas en Dash, donc on met les styles inline

        # Conteneur principal avec styles inline
        html.Div(className="login-background", style={
            "minHeight": "100vh",
            "background": "linear-gradient(-45deg, #0f172a, #1e293b, #0f172a, #1a1a2e)",
            "backgroundSize": "400% 400%",
            "display": "flex",
            "alignItems": "center",
            "justifyContent": "center",
            "padding": "20px",
            "fontFamily": "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif"
        }, children=[
            html.Div(className="login-card", style={
                "background": "rgba(30, 41, 59, 0.95)",
                "backdropFilter": "blur(20px)",
                "borderRadius": "24px",
                "padding": "48px",
                "width": "100%",
                "maxWidth": "420px",
                "boxShadow": "0 25px 50px -12px rgba(0, 0, 0, 0.5), 0 0 0 1px rgba(255, 255, 255, 0.1)",
                "border": "1px solid rgba(255, 255, 255, 0.1)"
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
                        "fontSize": "28px",
                        "fontWeight": "800",
                        "background": "linear-gradient(135deg, #22d3ee 0%, #a78bfa 100%)",
                        "WebkitBackgroundClip": "text",
                        "WebkitTextFillColor": "transparent",
                        "marginBottom": "8px"
                    }),
                    html.P("Connectez-vous pour accéder au dashboard", style={
                        "color": "#94a3b8",
                        "fontSize": "14px"
                    })
                ]),

                # Zone d'erreur (cachée par défaut)
                html.Div(id="login-error", style={"display": "none"}),

                # Formulaire
                html.Div(className="login-form", style={"marginTop": "24px"}, children=[

                    # Champ identifiant
                    html.Div(style={"position": "relative", "marginBottom": "20px"}, children=[
                        html.Span("👤", style={
                            "position": "absolute",
                            "left": "16px",
                            "top": "50%",
                            "transform": "translateY(-50%)",
                            "color": "#64748b",
                            "fontSize": "18px",
                            "zIndex": "10"
                        }),
                        dcc.Input(
                            id="login-username",
                            type="text",
                            placeholder="Identifiant (ex: tony, samuel...)",
                            style={
                                "width": "100%",
                                "padding": "16px 16px 16px 48px",
                                "background": "rgba(15, 23, 42, 0.8)",
                                "border": "2px solid #334155",
                                "borderRadius": "12px",
                                "color": "#f0f4f8",
                                "fontSize": "15px",
                                "boxSizing": "border-box"
                            },
                            autoComplete="username"
                        )
                    ]),

                    # Champ mot de passe
                    html.Div(style={"position": "relative", "marginBottom": "20px"}, children=[
                        html.Span("🔒", style={
                            "position": "absolute",
                            "left": "16px",
                            "top": "50%",
                            "transform": "translateY(-50%)",
                            "color": "#64748b",
                            "fontSize": "18px",
                            "zIndex": "10"
                        }),
                        dcc.Input(
                            id="login-password",
                            type="password",
                            placeholder="Mot de passe",
                            style={
                                "width": "100%",
                                "padding": "16px 16px 16px 48px",
                                "background": "rgba(15, 23, 42, 0.8)",
                                "border": "2px solid #334155",
                                "borderRadius": "12px",
                                "color": "#f0f4f8",
                                "fontSize": "15px",
                                "boxSizing": "border-box"
                            },
                            autoComplete="current-password"
                        )
                    ]),

                    # Bouton de connexion - utiliser dbc.Button pour meilleure compatibilité
                    dbc.Button(
                        "🚀 Se connecter",
                        id="login-button",
                        n_clicks=0,
                        color="info",
                        className="w-100 mt-2",
                        style={
                            "padding": "16px",
                            "background": "linear-gradient(135deg, #22d3ee 0%, #06b6d4 100%)",
                            "border": "none",
                            "borderRadius": "12px",
                            "color": "#0f172a",
                            "fontSize": "16px",
                            "fontWeight": "700",
                            "textTransform": "uppercase",
                            "letterSpacing": "1px"
                        }
                    )
                ]),

                # Footer
                html.Div([
                    html.P([
                        "Mot de passe oublié? Contactez ",
                        html.A("l'administrateur", href="mailto:tony.sarre@maad.io",
                               style={"color": "#22d3ee", "textDecoration": "none"})
                    ], style={"color": "#64748b", "fontSize": "12px", "textAlign": "center", "marginTop": "32px"}),
                    html.P("© 2024 MAAD - Supply Chain Analytics",
                           style={"color": "#475569", "fontSize": "11px", "textAlign": "center", "marginTop": "8px"})
                ])
            ])
        ])
    ])


def create_user_navbar(username: str):
    """Crée la barre utilisateur connecté"""
    if not username or username.lower() not in AUTH_USERS:
        return html.Div()

    user = AUTH_USERS[username.lower()]
    initials = "".join([n[0].upper() for n in user["name"].split()[:2]])

    role_colors = {"admin": "#ef4444", "manager": "#f59e0b", "user": "#22d3ee"}
    role_labels = {"admin": "Admin", "manager": "Manager", "user": "User"}

    # Prénom seulement pour gagner de la place
    first_name = user["name"].split()[0] if user["name"] else username

    return html.Div([
        html.Div(style={
            "position": "fixed",
            "top": "8px",
            "right": "12px",
            "zIndex": "9999",
            "display": "flex",
            "alignItems": "center",
            "gap": "8px",
            "padding": "6px 12px",
            "background": "rgba(30, 41, 59, 0.95)",
            "backdropFilter": "blur(10px)",
            "borderRadius": "50px",
            "border": "1px solid #334155",
            "boxShadow": "0 4px 20px rgba(0, 0, 0, 0.3)",
            "maxWidth": "180px"
        }, children=[
            # Avatar (plus petit)
            html.Div(initials, style={
                "width": "30px",
                "height": "30px",
                "minWidth": "30px",
                "background": f"linear-gradient(135deg, {role_colors.get(user['role'], '#22d3ee')}, #a78bfa)",
                "borderRadius": "50%",
                "display": "flex",
                "alignItems": "center",
                "justifyContent": "center",
                "fontWeight": "700",
                "color": "#0f172a",
                "fontSize": "12px"
            }),

            # Infos (prénom seulement + rôle court)
            html.Div([
                html.Div(first_name, style={
                    "color": "#f0f4f8",
                    "fontWeight": "600",
                    "fontSize": "12px",
                    "whiteSpace": "nowrap",
                    "overflow": "hidden",
                    "textOverflow": "ellipsis",
                    "maxWidth": "80px"
                }),
                html.Div(role_labels.get(user["role"], "User"), style={
                    "color": role_colors.get(user["role"], "#64748b"),
                    "fontSize": "9px",
                    "textTransform": "uppercase",
                    "letterSpacing": "0.5px"
                })
            ]),

            # Bouton déconnexion (plus petit)
            html.Button("🚪", id="logout-button", n_clicks=0, style={
                "background": "rgba(239, 68, 68, 0.2)",
                "border": "1px solid rgba(239, 68, 68, 0.4)",
                "borderRadius": "6px",
                "padding": "5px 8px",
                "cursor": "pointer",
                "fontSize": "14px",
                "color": "#f87171"
            }, title="Se déconnecter")
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
#      <h2 style="color: #22d3ee;">📌 Nouvelle mention</h2>
#     <p style="color: #e5e7eb;">Bonjour {to_name},</p>
#    <p style="color: #e5e7eb;">
#       <strong>{author}</strong> vous a mentionné dans une note sur le produit
#      <strong style="color: #22d3ee;">{product_name}</strong> :
# </p>
# <blockquote style="background: #1f2937; padding: 15px; border-left: 4px solid #22d3ee; margin: 20px 0;">
#   <p style="color: #e5e7eb; font-style: italic;">{message}</p>
# </blockquote>
# <p style="color: #9ca3af; font-size: 12px;">
#   Date : {datetime.now().strftime('%d/%m/%Y à %H:%M')}
# </p>
# <a href="https://your-dashboard-url.com"
#  style="display: inline-block; background: #22d3ee; color: #001018;
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


def add_note(product_name: str, author: str, message: str, mentions: list = None):
    """Ajoute une nouvelle note"""
    notes = _load_notes()
    new_note = {
        "id": f"note_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{len(notes)}",
        "product_name": product_name,
        "author": author,
        "message": message,
        "mentions": mentions or [],
        "timestamp": datetime.now().isoformat(),
        "read_by": []
    }
    notes.append(new_note)
    _save_notes(notes)
    return new_note


def get_notes_for_product(product_name: str):
    """Récupère toutes les notes d'un produit"""
    notes = _load_notes()
    return [n for n in notes if n.get("product_name", "").lower() == product_name.lower()]


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


#import pandas as pd
#import numpy as np
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import OneHotEncoder
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.impute import SimpleImputer
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
    roi_df = promo_pivot.merge(
        current_df[['product_name', 'product_id']],
        on='product_name',
        how='left'
    )

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
    except Exception as e:
        print(f"Warning: Could not load delisting data: {e}")
        delisting_df = pd.DataFrame(columns=["product_name", "delisting_status"])

    try:
        parametres_replenish_df = pd.read_csv(PARAMETRES_REPLENISH_URL)
        print("'Parametres Replenish' data loaded successfully.")
    except Exception as e:
        print(f"Warning: Could not load parametres replenish: {e}")
        parametres_replenish_df = pd.DataFrame()

    # ✅ NOUVEAU: Charger les dernières réceptions (Heroku Dataclip)
    # ⚠️ DÉSACTIVÉ temporairement car peut bloquer le chargement
    # Pour réactiver, décommenter le bloc try ci-dessous
    last_receptions_df = pd.DataFrame(columns=['product_name', 'last_reception_qty', 'last_reception_date'])
    print("\n📦 Chargement des dernières réceptions: DÉSACTIVÉ (pour éviter les timeouts)")

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

    # Charger le catalogue pour product_id
    # Charger le catalogue pour product_id ET is_active
    try:
        # IMPORTANT : skiprows=1 pour sauter la ligne d'en-tête
        catalog_df = pd.read_csv(CATALOG_URL, skiprows=1)

        print(f"📋 Catalogue chargé : {catalog_df.shape[0]} lignes, {catalog_df.shape[1]} colonnes")

        # Vérifier qu'il y a au moins 8 colonnes (index 0-7)
        if catalog_df.shape[1] > 7:
            # Extraire colonnes : A (id), B (name), H (is_active)
            # Index :              0        1          7
            catalog_subset = catalog_df.iloc[:, [0, 1, 7]].copy()
            catalog_subset.columns = ['product_id', 'product_name', 'is_active']

            # Nettoyer product_name
            catalog_subset['product_name_clean'] = (
                catalog_subset['product_name']
                .astype(str)
                .str.lower()
                .str.strip()
            )

            # Nettoyer is_active (TRUE/FALSE depuis Google Sheets)
            catalog_subset['is_active'] = (
                    catalog_subset['is_active']
                    .astype(str)
                    .str.upper()
                    .str.strip()
                    == 'TRUE'
            )

            # Préparer le mapping
            product_map = catalog_subset[['product_id', 'product_name_clean', 'is_active']].drop_duplicates(
                subset=['product_name_clean']
            )
            product_map.columns = ['product_id', 'product_name', 'is_active']

            # Convertir product_id en int
            product_map['product_id'] = pd.to_numeric(product_map['product_id'], errors='coerce').fillna(0).astype(int)

            print(f"✅ Catalogue traité :")
            print(f"   - Total produits : {len(product_map)}")
            print(f"   - Actifs (TRUE) : {product_map['is_active'].sum()}")
            print(f"   - Inactifs (FALSE) : {(~product_map['is_active']).sum()}")
            print(f"   - Échantillon actifs : {product_map[product_map['is_active']]['product_name'].head(3).tolist()}")

            product_id_map = product_map

        else:
            print(f"⚠️ Catalogue incomplet : {catalog_df.shape[1]} colonnes (8 minimum requis)")
            product_id_map = pd.DataFrame(columns=['product_id', 'product_name', 'is_active'])

    except Exception as e:
        print(f"❌ Erreur chargement catalogue : {e}")
        import traceback
        print(traceback.format_exc())
        product_id_map = pd.DataFrame(columns=['product_id', 'product_name', 'is_active'])
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

    # ✅ MERGER product_id ET is_active
    if not product_id_map.empty:
        # Supprimer colonnes existantes si présentes
        for col in ['product_id', 'is_active']:
            if col in total_stock_df.columns:
                print(f"⚠️ Colonne '{col}' déjà présente, elle sera remplacée")
                total_stock_df.drop(columns=[col], inplace=True)

        # Merge
        total_stock_df = total_stock_df.merge(
            product_id_map,
            on="product_name",
            how="left",
            validate="m:1"
        )

        # Gestion des valeurs manquantes
        total_stock_df['product_id'] = total_stock_df['product_id'].fillna(0).astype(int)
        total_stock_df['is_active'] = total_stock_df['is_active'].fillna(False)  # Par défaut inactif si absent

        print(f"✅ Merge catalogue effectué :")
        print(f"   - Produits avec ID valide : {(total_stock_df['product_id'] > 0).sum()}")
        print(f"   - Produits actifs : {total_stock_df['is_active'].sum()}")
        print(f"   - Produits sans match catalogue : {(total_stock_df['product_id'] == 0).sum()}")
    else:
        print("⚠️ Catalogue vide, product_id=0 et is_active=True par défaut")
        total_stock_df['product_id'] = 0
        total_stock_df['is_active'] = True

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
        print("⚠️ Pas de données de réception à fusionner")
        final_stock_sales_df['last_reception_qty'] = 0
        final_stock_sales_df['last_reception_date'] = None
        final_stock_sales_df['days_since_reception'] = None

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
        final_stock_sales_df, ml_model = create_supervised_target_from_sales(
            sales_history,
            final_stock_sales_df
        )
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

    try:
        promo_history = load_and_analyze_promotions()

        if not sales_history.empty and not promo_history.empty:
            print("Calcul ROI des promotions...")
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
    # ML: Stockout Probability
    # =========================
    try:
        feature_cols = [
            "total_stock", "Average Daily Sales",
            "Max Daily Sales (Pikine)", "optimal stock",
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

    if "Average Daily Sales" not in final_stock_sales_df.columns:
        final_stock_sales_df["Average Daily Sales"] = 0.1

    final_stock_sales_df["Average Daily Sales"] = (
        pd.to_numeric(final_stock_sales_df["Average Daily Sales"], errors="coerce")
        .fillna(0.1)
        .clip(lower=0.1)
    )

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
    # ✅ FILTRE FINAL : Produits actifs uniquement
    # =========================
    if 'is_active' in final_stock_sales_df.columns:
        initial_count = len(final_stock_sales_df)
        active_count = final_stock_sales_df['is_active'].sum()

        # Filtrer
        final_stock_sales_df = final_stock_sales_df[final_stock_sales_df['is_active'] == True].copy()

        print(f"\n{'=' * 60}")
        print(f"✅ FILTRE PRODUITS ACTIFS")
        print(f"{'=' * 60}")
        print(f"   Avant filtre : {initial_count} produits")
        print(f"   Actifs détectés : {active_count}")
        print(f"   Après filtre : {len(final_stock_sales_df)} produits")
        print(f"   Exclus : {initial_count - len(final_stock_sales_df)} produits")
        print(f"{'=' * 60}\n")
    else:
        print("\n⚠️ Colonne 'is_active' absente, TOUS les produits sont affichés")

    # =========================
    # ✅ FILTRE : Exclure les produits "cadeau"
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

        # ✅ AJOUTER JUSTE AVANT LE RETURN FINAL
        print(f"\n{'=' * 70}")
        print(f"🎉 LOAD_SUPPLY_DATA TERMINÉ")
        print(f"{'=' * 70}")
        print(f"   Période : {period_days}")
        print(f"   Produits : {len(final_stock_sales_df)}")
        print(f"   Colonnes : {len(final_stock_sales_df.columns)}")

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

        # product_id
        if "product_id" not in df.columns:
            df["product_id"] = 0
        else:
            df["product_id"] = (
                pd.to_numeric(df["product_id"], errors="coerce")
                .fillna(0)
                .astype(int)
            )

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
               VARIABLES AMÉLIORÉES
               ======================================== */
            :root {
                --brand-accent: #22d3ee;
                --brand-hover: #06b6d4;
                --badge-danger: #ef4444;
                --badge-warning: #f59e0b;
                --badge-ok: #10b981;
                --bg-primary: #0b1220;
                --bg-secondary: #0f1625;
                --bg-tertiary: #1a2332;
                --bg-input: #0a1320;
                --text-primary: #f0f4f8;
                --text-secondary: #cbd5e0;
                --text-muted: #94a3b8;
                --border-color: #2d3748;
                --border-hover: #4a5568;
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
                background: #0b1220; 
                border-right: 1px solid var(--border-color); 
                overflow-y: auto; 
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
                border: 1px solid #243244; 
                padding: 10px 12px; 
                border-radius: 12px; 
                background: #0f1828; 
                color: var(--text-primary) !important;
            }

            .kpi { 
                border-radius: 16px; 
                padding: 20px; 
                border: 1px solid var(--border-color); 
                background: linear-gradient(135deg, var(--bg-primary) 0%, var(--bg-secondary) 100%);
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
                transition: all 0.3s ease;
            }

            .kpi:hover {
                transform: translateY(-2px);
                box-shadow: 0 8px 24px rgba(34, 211, 238, 0.2);
                border-color: rgba(34, 211, 238, 0.3);
            }

            /* ========================================
               CARTES PRODUITS À RISQUE (CLIQUABLES)
               ======================================== */
            .risk-product-card {
                transition: all 0.2s ease !important;
            }

            .risk-product-card:hover {
                background: #1a2332 !important;
                transform: translateX(4px);
                box-shadow: 0 4px 12px rgba(34, 211, 238, 0.2);
                border-color: #22d3ee !important;
            }

            .risk-product-card:active {
                transform: translateX(2px);
                background: #243244 !important;
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

            /* Container principal */
            #rotation-period {
                background: var(--bg-input) !important;
            }

            /* Contrôle du dropdown */
            .dark-dropdown .Select-control,
            #rotation-period .Select-control {
                background: var(--bg-input) !important;
                border: 2px solid var(--border-color) !important;
                border-radius: 12px !important;
                color: var(--text-primary) !important;
                transition: all 0.3s ease !important;
            }

            .dark-dropdown .Select-control:hover,
            #rotation-period .Select-control:hover {
                border-color: var(--border-hover) !important;
            }

            /* Focus state */
            .dark-dropdown .is-focused:not(.is-open) > .Select-control,
            #rotation-period .is-focused:not(.is-open) > .Select-control {
                border-color: var(--brand-accent) !important;
                box-shadow: 0 0 0 3px rgba(34, 211, 238, 0.15) !important;
            }

            /* Texte sélectionné */
            .dark-dropdown .Select-value-label,
            .dark-dropdown .Select-placeholder,
            #rotation-period .Select-value-label,
            #rotation-period .Select-placeholder {
                color: var(--text-primary) !important;
                font-weight: 600 !important;
                font-size: 14px !important;
            }

            /* Flèche */
            .dark-dropdown .Select-arrow-zone,
            #rotation-period .Select-arrow-zone {
                color: var(--text-primary) !important;
            }

            .dark-dropdown .Select-arrow,
            #rotation-period .Select-arrow {
                border-color: var(--text-primary) transparent transparent !important;
            }

            /* Menu déroulant */
            .dark-dropdown .Select-menu-outer,
            #rotation-period .Select-menu-outer {
                background: var(--bg-tertiary) !important;
                border: 2px solid var(--border-color) !important;
                border-radius: 12px !important;
                box-shadow: 0 8px 24px rgba(0, 0, 0, 0.4) !important;
                margin-top: 4px !important;
            }

            /* Options */
            .dark-dropdown .Select-option,
            #rotation-period .Select-option {
                background: var(--bg-tertiary) !important;
                color: var(--text-primary) !important;
                padding: 12px 16px !important;
                font-weight: 500 !important;
                transition: all 0.2s ease !important;
            }

            .dark-dropdown .Select-option:hover,
            .dark-dropdown .Select-option.is-focused,
            #rotation-period .Select-option:hover,
            #rotation-period .Select-option.is-focused {
                background: rgba(34, 211, 238, 0.15) !important;
                color: var(--brand-accent) !important;
            }

            .dark-dropdown .Select-option.is-selected,
            #rotation-period .Select-option.is-selected {
                background: rgba(34, 211, 238, 0.25) !important;
                color: var(--brand-accent) !important;
                font-weight: 700 !important;
            }

            /* Input dans le dropdown */
            .dark-dropdown .Select-input > input,
            #rotation-period .Select-input > input {
                color: var(--text-primary) !important;
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
                box-shadow: 0 4px 16px rgba(0, 0, 0, 0.3) !important;
            }

            .dash-header {
                background: linear-gradient(135deg, var(--bg-tertiary) 0%, #243447 100%) !important;
                color: var(--text-primary) !important;
                font-weight: 700 !important;
                text-transform: uppercase !important;
                letter-spacing: 0.5px !important;
                font-size: 11px !important;
                border-bottom: 2px solid var(--brand-accent) !important;
            }

            .dash-cell {
                background: var(--bg-input) !important;
                color: var(--text-primary) !important;
                border-color: var(--border-color) !important;
                font-size: 13px !important;
            }

            .dash-table-container tr:nth-child(even) .dash-cell {
                background: rgba(26, 35, 50, 0.6) !important;
            }

            .dash-table-container tr:hover .dash-cell {
                background: rgba(34, 211, 238, 0.08) !important;
                border-color: rgba(34, 211, 238, 0.3) !important;
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
                background: var(--brand-accent); 
                color: #001018; 
                font-weight: 800;
                box-shadow: 0 12px 24px rgba(0, 0, 0, 0.35);
                display: flex; 
                align-items: center; 
                gap: 8px; 
                cursor: pointer;
                transition: all 0.3s ease;
            }

            .chat-fab:hover {
                transform: translateY(-2px);
                box-shadow: 0 16px 32px rgba(34, 211, 238, 0.4);
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
               RESPONSIVE
               ======================================== */
            @media (max-width: 768px) {
                .content {
                    margin-left: 0;
                    padding: 12px;
                }

                .sidebar {
                    display: none;
                }

                #search-input {
                    font-size: 16px !important; /* Évite le zoom sur mobile */
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
'''@cache.memoize()
def get_df_cached():
    return load_supply_data()
'''

@lru_cache(maxsize=10)
def get_df_cached(period: str = "7d"):
    """Cache avec support période rotation"""
    return load_supply_data(period_days=period)

# ------------------------------ Sidebar ------------------------------------------
def make_sidebar():
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
        #html.Div(className="banner-risk", id="risk-banner", children="Chargement..."),
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
            html.Small(" Agent"),
            dcc.Dropdown(
                id="filter-agent",
                options=[{"label": "Tous", "value": "all"}] + [{"label": a, "value": a} for a in agents_list],
                value="all",
                placeholder="Tous les agents",
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
            dbc.Button("📄 Bon de commande", id="btn-po-pdf", className="btn-primary", size="sm", disabled=True), # Le bouton est désactivé par défaut
            # ✅ AJOUTER : Modal de chargement
            dbc.Modal(
                [
                    dbc.ModalBody([
                        html.Div([
                            dbc.Spinner(color="primary", size="lg"),
                            html.H5("Génération du bon de commande...", className="mt-3", style={"color": "#22d3ee"}),
                            html.P("Veuillez patienter (5-10 secondes)", style={"color": "#9ca3af"})
                        ], style={"textAlign": "center", "padding": "30px"})
                    ], style={"background": "#0b1220"})
                ],
                id="po-loading-overlay",
                is_open=False,
                centered=True,
                backdrop="static",  # ✅ Empêche fermeture manuelle
                keyboard=False
            ),
            dcc.Download(id="download-data"),
            dcc.Download(id="download-po"),
        ]),
        html.Br(),
        html.Div([
            dbc.Nav([
                dbc.NavLink("Overview", href="/", id="nav-overview", active="exact"),
                dbc.NavLink("Analyses", href="/analytics", id="nav-analytics", active="exact"),
                dbc.NavLink("Prédictions", href="/predictions", id="nav-pred", active="exact"),
                dbc.NavLink(" Agents", href="/agents", id="nav-agents", active="exact"),
                dbc.NavLink(" Promotions", href="/promotions", active="exact"),
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

    # Filtrer les produits delisted
    if 'delisting_status' in df.columns:
        delist = df['delisting_status'].astype(str).str.lower()
        mask_valid = delist != 'delisted'
        df_clean = df.loc[mask_valid].copy()
    else:
        df_clean = df.copy()

    # KPI 1 : SKUs uniques
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

    print(f"   📊 calculate_stock_kpis: SKUs={total_skus}, Ruptures={out_of_stock}, À risque={at_risk}, Fournisseurs={suppliers}")

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
                        "background": "#0f1625",
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
                                "color": "#22d3ee",
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
                    "color": "#e5e7eb",
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
                        "background": "#0f1625",
                        "border": "1px solid #1f2937",
                        "borderLeft": f"4px solid {color_border}",
                        "borderRadius": "8px",
                    },
                    children=[
                        html.Div([
                            html.Strong(prod['product_name'], style={"color": "#22d3ee", "fontSize": "13px", "marginRight": "8px"}),
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
                    "color": "#e5e7eb",
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


def page_overview(master_df: pd.DataFrame = None):
    # Charger les données

    print(master_df.head())
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

        # Average Daily Sales non vide et numérique
        if "Average Daily Sales" not in df.columns:
            df["Average Daily Sales"] = 0.1

        df["Average Daily Sales"] = (
            pd.to_numeric(df["Average Daily Sales"], errors="coerce")
            .fillna(0.1)
            .clip(lower=0.1)
        )

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

    # Vérifier que product_name est dans available_cols
    if "product_name" not in available_cols and "product_name" in df_overview.columns:
        available_cols.insert(0, "product_name")

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

    # Dans page_overview(), AVANT le tableau

    # ✅ DROPDOWN SIMPLIFIÉ (sans indicateur)
    rotation_dropdown = html.Div([
        html.Label(" Période de rotation ADS :", style={
            "fontWeight": "700",
            "marginRight": "12px",
            "color": "#f0f4f8",
            "fontSize": "14px"
        }),
        dcc.Dropdown(
            id="rotation-period",
            options=[
                {"label": " 3 jours ", "value": "3d"},
                {"label": " 7 jours ", "value": "7d"},
                {"label": " 30 jours ", "value": "30d"},
            ],
            value="7d",
            clearable=False,
            searchable=False,
            style={"width": "400px"},
            className="dark-dropdown"
        ),
        # ❌ SUPPRIMER rotation-indicator
    ], style={
        "padding": "16px 18px",
        "background": "linear-gradient(135deg, #1a2332 0%, #141b2d 100%)",
        "borderRadius": "12px",
        "border": "1px solid #334155",
        "marginBottom": "20px",
        "boxShadow": "0 4px 12px rgba(0,0,0,0.2)"
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
        page_size=15,
        filter_action="native",
        sort_action="native",
        sort_mode="multi",
        column_selectable="single",
        editable=True,
        active_cell=None,
        dropdown_conditional=[],
        row_selectable="multi",
        selected_rows=[],

        # ========================================
        # TABLE STYLES
        # ========================================
        style_table={
            "overflowX": "auto",
            "maxWidth": "100%"
        },

        style_header={
            "backgroundColor": "#0f1625",
            "border": "1px solid #2d3748",
            "fontWeight": "700",
            "textAlign": "center",
            "color": "#f0f4f8",
            "fontSize": "12px",
            "textTransform": "uppercase",
            "letterSpacing": "0.5px",
            "padding": "12px 8px"
        },

        style_cell={
            "backgroundColor": "#0b1220",
            "color": "#e5e7eb",
            "border": "1px solid #1f2937",
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
                "backgroundColor": "#1a2332",
                "border": "2px dashed #4a5568",
                "color": "#22d3ee"
            },

            # QAC - valeur calculée
            {
                "if": {"column_id": "QAC"},
                "backgroundColor": "#1a2332",
                "cursor": "default",
                "fontWeight": "700",
                "fontSize": "13px",
                "color": "#22d3ee"
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
                "color": "#94a3b8",
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
                "color": "#a78bfa",
                "fontSize": "13px"
            },

            # Max Daily Sales
            {
                "if": {"column_id": "Max Daily Sales (Pikine)"},
                "fontWeight": "600",
                "color": "#fbbf24"
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
                "backgroundColor": "rgba(239, 68, 68, 0.15)",
                "color": "#fecaca",
                "fontWeight": "600"
            },

            # Colonne Ajusted_total_need - ORDER NOW (badge fort)
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'ORDER NOW'",
                    "column_id": "Ajusted_total_need"
                },
                "backgroundColor": "rgba(239, 68, 68, 0.4)",
                "color": "#ffffff",
                "fontWeight": "800",
                "fontSize": "12px",
                "border": "2px solid #ef4444",
                "borderRadius": "6px",
                "textTransform": "uppercase"
            },

            # product_name - ORDER NOW (bordure rouge)
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'ORDER NOW'",
                    "column_id": "product_name"
                },
                "backgroundColor": "rgba(239, 68, 68, 0.25)",
                "color": "#fee2e2",
                "fontWeight": "700",
                "borderLeft": "4px solid #ef4444"
            },

            # Stock - ORDER NOW (highlight)
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'ORDER NOW'",
                    "column_id": "total_stock"
                },
                "backgroundColor": "rgba(239, 68, 68, 0.3)",
                "color": "#fecaca",
                "fontWeight": "700",
                "fontSize": "14px"
            },

            # QAC - ORDER NOW (highlight)
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'ORDER NOW'",
                    "column_id": "QAC"
                },
                "backgroundColor": "rgba(239, 68, 68, 0.3)",
                "color": "#ffffff",
                "fontWeight": "800",
                "fontSize": "15px",
                "border": "2px solid #ef4444"
            },

            # ========================================
            # 🟠 ORDER NOT URGENT - LIGNE ORANGE
            # ========================================
            {
                "if": {"filter_query": "{Ajusted_total_need} = 'ORDER NOT URGENT'"},
                "backgroundColor": "rgba(245, 158, 11, 0.12)",
                "color": "#fde68a",
                "fontWeight": "500"
            },

            # Colonne Ajusted_total_need - ORDER NOT URGENT
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'ORDER NOT URGENT'",
                    "column_id": "Ajusted_total_need"
                },
                "backgroundColor": "rgba(245, 158, 11, 0.35)",
                "color": "#ffffff",
                "fontWeight": "700",
                "fontSize": "12px",
                "border": "2px solid #f59e0b",
                "borderRadius": "6px",
                "textTransform": "uppercase"
            },

            # product_name - ORDER NOT URGENT
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'ORDER NOT URGENT'",
                    "column_id": "product_name"
                },
                "backgroundColor": "rgba(245, 158, 11, 0.2)",
                "color": "#fef3c7",
                "fontWeight": "600",
                "borderLeft": "4px solid #f59e0b"
            },

            # QAC - ORDER NOT URGENT
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'ORDER NOT URGENT'",
                    "column_id": "QAC"
                },
                "backgroundColor": "rgba(245, 158, 11, 0.25)",
                "color": "#ffffff",
                "fontWeight": "700",
                "fontSize": "14px"
            },

            # ========================================
            # 🟢 NO NEED - LIGNE VERTE
            # ========================================
            {
                "if": {"filter_query": "{Ajusted_total_need} = 'NO NEED'"},
                "backgroundColor": "rgba(16, 185, 129, 0.08)",
                "color": "#d1fae5"
            },

            # Colonne Ajusted_total_need - NO NEED
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'NO NEED'",
                    "column_id": "Ajusted_total_need"
                },
                "backgroundColor": "rgba(16, 185, 129, 0.3)",
                "color": "#ffffff",
                "fontWeight": "700",
                "fontSize": "12px",
                "border": "2px solid #10b981",
                "borderRadius": "6px",
                "textTransform": "uppercase"
            },

            # product_name - NO NEED (bordure verte subtile)
            {
                "if": {
                    "filter_query": "{Ajusted_total_need} = 'NO NEED'",
                    "column_id": "product_name"
                },
                "borderLeft": "4px solid #10b981"
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
                "backgroundColor": "rgba(239, 68, 68, 0.25)",
                "color": "#ffffff",
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
                "backgroundColor": "rgba(245, 158, 11, 0.2)",
                "color": "#fde68a",
                "fontWeight": "700",
                "fontSize": "13px"
            },

            # 14-30 jours - BON
            {
                "if": {
                    "filter_query": "{Max Coverage Day} >= 14 && {Max Coverage Day} < 30",
                    "column_id": "Max Coverage Day"
                },
                "backgroundColor": "rgba(34, 197, 94, 0.15)",
                "color": "#bbf7d0",
                "fontWeight": "600"
            },

            # Plus de 30 jours - EXCELLENT
            {
                "if": {
                    "filter_query": "{Max Coverage Day} >= 30",
                    "column_id": "Max Coverage Day"
                },
                "backgroundColor": "rgba(16, 185, 129, 0.2)",
                "color": "#a7f3d0",
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
                "backgroundColor": "rgba(239, 68, 68, 0.2)",
                "color": "#fecaca"
            },
            {
                "if": {
                    "filter_query": "{Product Category} contains 'b'",
                    "column_id": "Product Category"
                },
                "backgroundColor": "rgba(245, 158, 11, 0.2)",
                "color": "#fde68a"
            },
            {
                "if": {
                    "filter_query": "{Product Category} contains 'c'",
                    "column_id": "Product Category"
                },
                "backgroundColor": "rgba(16, 185, 129, 0.15)",
                "color": "#a7f3d0"
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
                "backgroundColor": "rgba(16, 185, 129, 0.25)",
                "color": "#34d399",
                "fontWeight": "700"
            },
            # Réception moyenne (7-14 jours) - Jaune
            {
                "if": {
                    "filter_query": "{days_since_reception} >= 7 && {days_since_reception} < 14",
                    "column_id": "days_since_reception"
                },
                "backgroundColor": "rgba(245, 158, 11, 0.2)",
                "color": "#fbbf24",
                "fontWeight": "600"
            },
            # Réception ancienne (14-30 jours) - Orange
            {
                "if": {
                    "filter_query": "{days_since_reception} >= 14 && {days_since_reception} < 30",
                    "column_id": "days_since_reception"
                },
                "backgroundColor": "rgba(249, 115, 22, 0.2)",
                "color": "#fb923c",
                "fontWeight": "600"
            },
            # Réception très ancienne (> 30 jours) - Rouge
            {
                "if": {
                    "filter_query": "{days_since_reception} >= 30",
                    "column_id": "days_since_reception"
                },
                "backgroundColor": "rgba(239, 68, 68, 0.2)",
                "color": "#f87171",
                "fontWeight": "700"
            },

            # ========================================
            # 🎯 ÉTATS INTERACTIFS
            # ========================================

            # Lignes sélectionnées
            {
                "if": {"state": "selected"},
                "backgroundColor": "rgba(34, 211, 238, 0.3)",
                "border": "2px solid #22d3ee",
                "fontWeight": "700",
                "color": "#ffffff"
            },

            # Cellule active (en cours d'édition)
            {
                "if": {"state": "active"},
                "backgroundColor": "#1e293b",
                "border": "2px solid #22d3ee",
                "outline": "none",
                "color": "#ffffff",
                "fontWeight": "700"
            },

            # QAC edited - toujours visible
            {
                "if": {"column_id": "QAC edited"},
                "backgroundColor": "#1a2332"
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

    action_buttons = dbc.ButtonGroup([
        dbc.Button("🔄 Actualiser", id="btn-refresh", className="btn-outline-secondary", size="sm"),
        dbc.Button("➕ Ajouter produit", id="btn-add-row", className="btn-primary", size="sm"),
        dbc.Button("💾 Enregistrer QAC", id={'type': 'btn-save-qac', 'index': 'dbc'}, className="btn-success", size="sm"), # ✅ nouveau
    ], style={"marginBottom": "15px"})


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
            "borderBottom": "2px solid #22d3ee",
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
                    "background": "linear-gradient(135deg, #22d3ee 0%, #06b6d4 100%)",
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
            dbc.Col(html.H2(" Overview", style={"color": "#22d3ee", "fontWeight": "800"}), md=8),
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
    ⚡ CALLBACK UNIFIÉ - Initialisation + Filtrage
    S'exécute au démarrage ET quand les filtres changent
    """
    start_time = time.time()

    print(f"\n{'=' * 60}")
    print(f"⚡ UPDATE TABLE")
    print(f"{'=' * 60}")

    # ✅ Charger données
    if not master_json:
        print("📥 Chargement depuis cache...")
        base = get_df_cached()
    else:
        print("📥 Chargement depuis master-data...")
        base = pd.DataFrame(json.loads(master_json))

    print(f"📊 Base : {len(base)} produits")

    # ✅ Validation
    base = validate_core_columns(base)

    # ✅ Supprimer colonnes bannies
    promo_cols = [
        'promo_status', 'uplift_pct', 'roi_pct',
        'Average Daily Sales (7d)', 'Average Daily Sales (30d)',
        'Average Daily Sales (3d)', 'Daily OOS Rate (7d)',
        'Daily OOS Rate (30d)', 'Stockout Probability',
        'Credit Adequacy Score'
    ]
    base = base.drop(columns=[c for c in promo_cols if c in base.columns], errors='ignore')

    # ✅ APPLIQUER FILTRES (seulement si des filtres sont actifs)
    fdf = base.copy()

    if search and search.strip():
        search_lower = search.strip().lower()
        mask = pd.Series([False] * len(fdf), index=fdf.index)
        for col in ['product_name', 'Supplier', 'Product Category']:
            if col in fdf.columns:
                mask |= fdf[col].astype(str).str.lower().str.contains(search_lower, na=False, regex=False)
        fdf = fdf[mask].reset_index(drop=True)
        print(f"   🔎 Recherche : {len(fdf)} produits")

    if sup and len(sup) > 0 and 'Supplier' in fdf.columns:
        fdf = fdf[fdf['Supplier'].isin(sup)]
        print(f"   🏭 Fournisseurs : {len(fdf)} produits")

    if cat and len(cat) > 0 and 'Product Category' in fdf.columns:
        fdf = fdf[fdf['Product Category'].isin(cat)]
        print(f"   🏷️ Catégories : {len(fdf)} produits")

    if need and len(need) > 0 and 'Ajusted_total_need' in fdf.columns:
        fdf = fdf[fdf['Ajusted_total_need'].isin(need)]
        print(f"   📦 Besoins : {len(fdf)} produits")

    if options and len(options) > 0 and 'Stock Status' in fdf.columns:
        for opt in options:
            if opt == 'show_out_of_stock':
                fdf = fdf[fdf['Stock Status'] == 'Out of Stock']
            elif opt == 'show_predicted_stockout':
                fdf = fdf[fdf['Stock Status'] == 'Predicted Stockout Soon']
            elif opt == 'show_order_soon':
                fdf = fdf[fdf['Stock Status'] == 'Order Soon']

    # ✅ FORMATAGE
    for col in ['total_stock', 'QAC', 'target_quantity']:
        if col in fdf.columns:
            fdf[col] = pd.to_numeric(fdf[col], errors='coerce').fillna(0).astype(int)

    if 'Max Coverage Day' in fdf.columns:
        fdf['Max Coverage Day'] = pd.to_numeric(fdf['Max Coverage Day'], errors='coerce').fillna(0).round(1)

    if 'Average Daily Sales' in fdf.columns:
        fdf['Average Daily Sales'] = pd.to_numeric(fdf['Average Daily Sales'], errors='coerce').fillna(0.1).round(2)

    # ✅ Actions
    fdf = add_action_cols(fdf)

    elapsed_total = time.time() - start_time

    print(f"⚡ UPDATE TERMINÉ en {elapsed_total:.3f}s - {len(fdf)} produits")
    print(f"{'=' * 60}\n")

    return (
        fdf.to_json(orient="records"),
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
                f'<span style="color:#22d3ee;font-weight:600">@{mention}</span>'
            )

        notes_display.append(
            html.Div([
                html.Div([
                    html.Span("👤", style={"marginRight": "6px"}),
                    html.Strong(note["author"], style={"color": "#22d3ee"}),
                    html.Span(f" • {timestamp}",
                              style={"color": "#6b7280", "fontSize": "11px", "marginLeft": "6px"})
                ], style={"marginBottom": "6px"}),

                dcc.Markdown(
                    message_html,
                    dangerously_allow_html=True,
                    style={"color": "#e5e7eb", "fontSize": "13px", "lineHeight": "1.5"}
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
    Output('main-table', 'data', allow_duplicate=True),  # Cela dépend de ce que tu veux actualiser
    Input('btn-refresh', 'n_clicks'),
    prevent_initial_call=True
)
def refresh_data(n_clicks):
    if n_clicks:
        # Rafraîchir les données ici (par exemple, recharger les données depuis la source)
        updated_data = load_supply_data()  # Assure-toi d'avoir une fonction load_data() qui recharge les données
        return updated_data
    return no_update
@app.callback(
    Output('main-table', 'data', allow_duplicate=True),
    Input('btn-add-row', 'n_clicks'),
    State('main-table', 'data'),
    prevent_initial_call=True
)
def add_new_product(n_clicks, current_data):
    if n_clicks:
        new_product = {
            "product_name": "Nouveau produit",
            "total_stock": 0,
            "Average Daily Sales": 0.1,
            # Remplir avec d'autres valeurs par défaut si nécessaire
        }
        current_data.append(new_product)
        return current_data
    return no_update
'''

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
                    "color": "#f0f4f8",
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
                    "color": "#f0f4f8",
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
                    "color": "#f0f4f8",
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
            "color": "#22d3ee",
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
            dbc.Col(html.H2(" Analyses Avancées", style={"color": "#22d3ee"}), md=12)
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
                plot_bgcolor="#0b1220",
                paper_bgcolor="#0b1220",
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
                plot_bgcolor="#0b1220",
                paper_bgcolor="#0b1220",
                font=dict(color="#e5e7eb", size=12),
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
                color_discrete_sequence=["#22d3ee"]
            )

            fig_hist_stock.update_layout(
                plot_bgcolor="#0b1220",
                paper_bgcolor="#0b1220",
                font=dict(color="#e5e7eb", size=12),
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
                plot_bgcolor="#0b1220",
                paper_bgcolor="#0b1220",
                font=dict(color="#e5e7eb", size=12),
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
                    plot_bgcolor="#0b1220",
                    paper_bgcolor="#0b1220",
                    font=dict(color="#e5e7eb", size=12),
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
                plot_bgcolor="#0b1220",
                paper_bgcolor="#0b1220",
                font=dict(color="#e5e7eb", size=12),
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
                plot_bgcolor="#0b1220",
                paper_bgcolor="#0b1220",
                font=dict(color="#e5e7eb", size=12),
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
            plot_bgcolor="#0b1220",
            paper_bgcolor="#0b1220"
        )

        error_indicator = html.Div([
            html.Span(f"❌ Erreur : {str(e)[:100]}", style={
                "fontWeight": "700",
                "color": "#ef4444"
            })
        ])

        return (empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, error_indicator)

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
                    "color": "#f0f4f8",
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
                    "color": "#f0f4f8",
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
                    "color": "#f0f4f8",
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
            "color": "#22d3ee",
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
            dbc.Col(html.H3("🔮 Prédictions ML", style={"color": "#22d3ee"}), md=8),
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
                    "backgroundColor": "#0f1625",
                    "border": "1px solid #2d3748",
                    "fontWeight": "700",
                    "textAlign": "center",
                    "color": "#f0f4f8"
                },
                style_cell={
                    "backgroundColor": "#0b1220",
                    "color": "#e5e7eb",
                    "border": "1px solid #1f2937",
                    "fontSize": 12,
                    "textAlign": "center"
                },
                style_data_conditional=[
                    {
                        "if": {"column_id": "target_quantity"},
                        "fontWeight": "700",
                        "color": "#22d3ee",
                        "fontSize": "14px"
                    }
                ]
            )
        ])
    ])


# ==========================================
# 🔮 CALLBACKS PRÉDICTIONS DYNAMIQUES
# ==========================================

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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb"),
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb")
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb"),
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb"),
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
                "backgroundColor": "#0f1625",
                "border": "1px solid #1f2937",
                "fontWeight": "700",
                "textAlign": "center"
            },
            style_cell={
                "backgroundColor": "#0b1220",
                "color": "#e5e7eb",
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb"),
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
                "color": "#22d3ee",
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

        # ✅ Placeholder pour éviter erreurs
        create_hidden_table_placeholder()
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb", size=11),
        height=400
    )

    # ========================================
    # 📊 GRAPHIQUE 2 : HISTOGRAMME STOCK
    # ========================================
    fig_hist = px.histogram(
        df,
        x="total_stock",
        nbins=30,
        labels={"total_stock": "Stock Total"},
        color_discrete_sequence=["#22d3ee"]
    )

    fig_hist.update_layout(
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb", size=11),
        showlegend=False,
        height=400
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb", size=11),
        showlegend=False,
        height=400
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb", size=11),
        xaxis=dict(tickangle=-45),
        showlegend=False,
        height=400
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb", size=11),
        height=400
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb", size=11),
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb", size=11),
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb", size=11),
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
                "color": "#22d3ee",
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
                    "backgroundColor": "#0f1625",
                    "border": "1px solid #2d3748",
                    "fontWeight": "700",
                    "textAlign": "center",
                    "color": "#f0f4f8"
                },
                style_cell={
                    "backgroundColor": "#0b1220",
                    "color": "#e5e7eb",
                    "border": "1px solid #1f2937",
                    "fontSize": 12,
                    "textAlign": "center"
                },
                style_data_conditional=[
                    {
                        "if": {"column_id": "target_quantity"},
                        "fontWeight": "700",
                        "color": "#22d3ee",
                        "fontSize": "14px"
                    }
                ]
            )
        ]),

        # ✅ Placeholder
        create_hidden_table_placeholder()
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb", size=11),
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb", size=11),
        height=400
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb", size=11),
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb", size=11),
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
        plot_bgcolor="#0b1220",
        paper_bgcolor="#0b1220",
        font=dict(color="#e5e7eb", size=11),
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
            html.H5("Auteur", style={"color": "#e5e7eb", "marginTop": "20px"}),
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
            plot_bgcolor="#0b1220",
            paper_bgcolor="#0b1220",
            font=dict(color="#e5e7eb"),
            xaxis=dict(tickangle=-45),
            height=400
        )
    else:
        fig_roi = {'data': [],
                   'layout': {'title': 'Pas de données ROI', 'plot_bgcolor': '#0b1220', 'paper_bgcolor': '#0b1220',
                              'font': {'color': '#e5e7eb'}}}

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
            plot_bgcolor="#0b1220",
            paper_bgcolor="#0b1220",
            font=dict(color="#e5e7eb"),
            xaxis=dict(tickangle=-45),
            height=400
        )
    else:
        fig_uplift = {'data': [], 'layout': {'title': 'Pas de données Uplift', 'plot_bgcolor': '#0b1220',
                                             'paper_bgcolor': '#0b1220', 'font': {'color': '#e5e7eb'}}}

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
                    "backgroundColor": "#0f1625",
                    "color": "white",
                    "fontWeight": "bold",
                    "textAlign": "center"
                },
                style_cell={
                    "backgroundColor": "#0b1220",
                    "color": "#e5e7eb",
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
        ])
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
        "background": "linear-gradient(135deg, #1e293b 0%, #0f172a 100%)", "borderRadius": "12px",
        "padding": "16px 20px", "minWidth": "130px", "border": "1px solid #334155",
        "boxShadow": "0 4px 12px rgba(0,0,0,0.3)"
    }
    kpi_value_style = {"fontSize": "28px", "fontWeight": "700", "lineHeight": "1.2"}
    kpi_label_style = {"color": "#94a3b8", "fontSize": "11px", "marginTop": "6px", "textTransform": "uppercase", "letterSpacing": "0.5px"}

    kpis_row = html.Div([
        html.Div([
            html.Div(f"{total_agents}", style={**kpi_value_style, "color": "#22d3ee"}),
            html.Div(" Agents", style=kpi_label_style)
        ], style=kpi_card_style),
        html.Div([
            html.Div(f"{avg_score:.0f}", style={**kpi_value_style, "color": "#34d399" if avg_score >= 60 else "#f87171"}),
            html.Div(" Score Moyen", style=kpi_label_style)
        ], style=kpi_card_style),
        html.Div([
            html.Div(f"{total_order_value/1e6:.1f}M", style={**kpi_value_style, "color": "#a78bfa"}),
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
            'backgroundColor': '#1e293b', 'color': '#94a3b8', 'fontWeight': '600',
            'fontSize': '12px', 'border': 'none', 'padding': '12px 8px', 'textAlign': 'center'
        },
        style_cell={
            'backgroundColor': '#0f172a', 'color': '#e2e8f0', 'border': 'none',
            'padding': '10px 8px', 'fontSize': '13px', 'textAlign': 'center', 'minWidth': '70px'
        },
        style_data_conditional=[
            {'if': {'filter_query': '{rank} = 1'}, 'borderLeft': '3px solid #fbbf24', 'backgroundColor': 'rgba(251, 191, 36, 0.05)'},
            {'if': {'filter_query': '{rank} = 2'}, 'borderLeft': '3px solid #94a3b8'},
            {'if': {'filter_query': '{rank} = 3'}, 'borderLeft': '3px solid #b45309'},
            {'if': {'column_id': 'sc'}, 'fontWeight': '700', 'color': '#22d3ee', 'fontSize': '14px'},
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
                html.Span("Performance des Agents", style={"fontSize": "16px", "fontWeight": "600", "color": "#e2e8f0"})
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
            'backgroundColor': '#1e293b', 'color': '#94a3b8', 'fontWeight': '600',
            'fontSize': '12px', 'border': 'none', 'padding': '12px 8px', 'textAlign': 'left'
        },
        style_cell={
            'backgroundColor': '#0f172a', 'color': '#e2e8f0', 'border': 'none',
            'padding': '10px 8px', 'fontSize': '13px', 'textAlign': 'left',
            'maxWidth': '200px', 'overflow': 'hidden', 'textOverflow': 'ellipsis'
        },
        style_data_conditional=[
            {'if': {'filter_query': '{status} contains "Rupture"'}, 'backgroundColor': 'rgba(239, 68, 68, 0.1)'},
            {'if': {'filter_query': '{status} contains "Critique"'}, 'backgroundColor': 'rgba(251, 146, 60, 0.08)'},
            {'if': {'column_id': 'agt'}, 'fontWeight': '600', 'color': '#22d3ee'},
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
                html.Span(f"Produits à Risque ({len(risk_products)})", style={"fontSize": "16px", "fontWeight": "600", "color": "#e2e8f0"})
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
            html.H4(" Performance Agents", style={"margin": "0", "fontSize": "20px", "fontWeight": "700", "color": "#22d3ee"}),
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
        detail
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
        html.B(agent_name, style={"color": "#22d3ee"}),
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
    Version Gemini – conserve toute la logique métier (détection langue, contexte, fallback).
    """
    try:
        if not user_text or not str(user_text).strip():
            return "Je peux analyser tes stocks, les risques de rupture et suggérer des quantités à commander."

        lang = _detect_lang(user_text)
        sample = _clean_df_for_advice(df if isinstance(df, pd.DataFrame) else pd.DataFrame())
        fields = [
            'product_name_display', 'supplier_name', 'abc_class', 'xyz_class', 'current_stock',
            'avg_daily_sales', 'coverage_days', 'leadtime_days', 'credit_days', 'rupture_ml',
            'delisting_product', 'risk'
        ]
        view = sample[[c for c in fields if c in sample.columns]].copy() if not sample.empty else pd.DataFrame()

        cov_med = None
        rows_digest = []
        try:
            if not view.empty and 'coverage_days' in view.columns:
                cov_med = float(view['coverage_days'].median())
                at_risk = view[
                    view['rupture_ml'].str.upper() == 'OUI'] if 'rupture_ml' in view.columns else pd.DataFrame()
                top_list = (
                    at_risk.sort_values('coverage_days', ascending=True).head(6)
                    if not at_risk.empty else
                    view.sort_values('coverage_days', ascending=True).head(6)
                )

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

        # Détecte si l'utilisateur demande une analyse
        needs_analysis = any(keyword in user_text.lower() for keyword in [
            'analyse', 'recommande', 'conseil', 'rupture', 'commande', 'stock',
            'produit', 'quels', 'combien', 'urgent', 'priorité', 'fournisseur',
            'risque', 'order', 'achat', 'besoin', 'coverage', 'lead time'
        ])

        if lang == "fr":
            system_msg = (
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
            )

            if needs_analysis and table_json:
                ctx = "Données disponibles pour analyse :\n"
                ctx += f"- Extrait produits (20 premiers) : {json.dumps(table_json[:20], ensure_ascii=False)}\n"
                if cov_med is not None:
                    ctx += f"- KPI global : couverture médiane ≈ {cov_med:.1f} jours\n"
                if rows_digest:
                    ctx += "- Produits prioritaires : " + "; ".join(
                        [
                            "{name} (fournisseur {sup}, couverture {cov:.1f}j, lead time {lt}j)".format(**d)
                            for d in rows_digest[:3]
                        ]
                    ) + "\n"

                else:
                    ctx = "Conversation générale. Données disponibles si besoin d'analyse détaillée.\n"

            user_q = f"Question : {user_text.strip()}"

        else:
            system_msg = (
                "You are maad_assistante, a Supply Chain expert with 15 years of experience. "
                "You're conversational and respond naturally to questions.\n\n"
                "CONVERSATION RULES:\n"
                "- If the user greets you or chats, respond in a friendly and natural way\n"
                "- If the user asks a general supply chain question, explain clearly without forcing data analysis\n"
                "- ONLY if the user explicitly requests analysis, recommendations, or advice on their inventory, use the provided data\n"
                "- Adapt your detail level to the question: simple question = short answer, analysis requested = detailed with figures\n\n"
                "WHEN ANALYZING DATA:\n"
                "- Cite concrete SKUs from the table\n"
                "- Propose quantified amounts\n"
                "- Mention priority suppliers\n"
                "- Suggest levers (credit, promo clearance, ABC/XYZ categories)\n"
                "- Be action-oriented with precise recommendations"
            )

            if needs_analysis and table_json:
                ctx = "Available data for analysis:\n"
                ctx += f"- Product excerpt (20 first): {json.dumps(table_json[:20], ensure_ascii=False)}\n"
                if cov_med is not None:
                    ctx += f"- Global KPI: median coverage ≈ {cov_med:.1f} days\n"
                if rows_digest:
                    ctx += "- Produits prioritaires : " + "; ".join(
                        [
                            "{name} (fournisseur {sup}, couverture {cov:.1f}j, lead time {lt}j)".format(**d)
                            for d in rows_digest[:3]
                        ]
                    ) + "\n"

            else:
                ctx = "General conversation. Data available if detailed analysis needed.\n"

            user_q = f"Question: {user_text.strip()}"

        # ------------------------- Appel Gemini (unique prompt) -------------------------
        # Historique compact (10 derniers messages)
        hist_lines = []
        for m in (history_messages or [])[-10:]:
            role = "Utilisateur" if m.get("role") == "user" else "Assistant"
            hist_lines.append(f"{role}: {m.get('text', '').strip()}")
        history_txt = "\n".join(hist_lines) if hist_lines else "—"

        prompt_text = (
            f"[SYSTEM]\n{system_msg}\n\n"
            f"[CONTEXTE]\n{ctx}\n\n"
            f"[HISTORIQUE (dernier·e·s 10)]\n{history_txt}\n\n"
            f"[QUESTION]\n{user_q}\n\n"
            f"[INSTRUCTIONS DE SORTIE]\n"
            f"- Réponds en **{'français' if lang == 'fr' else 'anglais'}**.\n"
            f"- Sois concis, clair, structuré en puces si nécessaire.\n"
            f"- N'invente pas de calculs : appuie-toi uniquement sur les données fournies dans le CONTEXTE.\n"
        )

        out = ask_ai(prompt_text).strip()
        return out if out else _chatbot_fallback(user_text, view)
        # -------------------------------------------------------------------------------

    except Exception as e:
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
    # ========== STORES (données partagées entre callbacks) ==========
    dcc.Store(id='agent-ia-recommendations', data=None),
    dcc.Store(id='bc-data-store', data=None),

    # ========== CONTENEUR PRINCIPAL (login ou dashboard) ==========
    html.Div(id="app-container"),

    # ========== COMPOSANTS CACHÉS (pour callbacks) ==========
    html.Div(id='edit-product-output', style={"display": "none"}),


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

    # Dans la page concernée, ajoutez :
    html.Div([
        dbc.Button(
            "🤖 Lancer Agent IA",
            id="btn-run-agent-ia",
            color="primary",
            className="mb-3"
        ),
        html.Div(id="agent-ia-output")  # Output du callback
    ], style={"display": "none"}),  # Caché par défaut, visible après login

    # ========== FEEDBACKS & COMPTEURS ==========
    # html.Div(id="action-feedback", style={"position": "fixed", "top": "80px", "right": "20px", "zIndex": 10000}),
    #html.Div(id="selection-counter"),

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
    if pathname is None or pathname == "/" or pathname == "/overview":
        page_content = page_overview(initial_df)
        page_name = "overview"
    elif pathname == "/analytics":
        page_content = page_analytics(initial_df)
        page_name = "analytics"
    elif pathname == "/predictions":
        page_content = page_predictive(initial_df)
        page_name = "predictions"
    elif pathname == "/promotions":
        page_content = page_promotions()
        page_name = "promotions"
    elif pathname == "/agents":
        page_content = page_agents(initial_df)
        page_name = "agents"
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
            className="content",  # Utiliser la classe CSS existante
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
    """Gère la tentative de connexion"""
    print(f"🔐 handle_login: n_clicks={n_clicks}, username={username}")

    if not n_clicks or n_clicks == 0:
        raise dash.exceptions.PreventUpdate

    # Validation des champs
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

    # Vérification des identifiants
    username = username.lower().strip()

    if verify_password(username, password):
        # Connexion réussie
        print(f"   ✅ Connexion réussie pour {username}")
        user_info = get_user_info(username)

        # ✅ TRACKING SUPABASE: Créer session
        session_id = None
        user_id = None
        if supabase_client:
            try:
                # Récupérer ou créer l'utilisateur
                db_user = get_or_create_user(username)
                if db_user:
                    user_id = db_user["id"]
                    # Créer une session
                    session_id = create_session(user_id)
                    # Stocker la session active
                    ACTIVE_SESSIONS[username] = {
                        "user_id": user_id,
                        "session_id": session_id
                    }
                    # Tracker l'activité de login
                    track_activity(user_id, session_id, "login", page="login")
            except Exception as e:
                print(f"⚠️ Erreur tracking login: {e}")

        return (
            {
                "authenticated": True,
                "username": username,
                "user_info": user_info,
                "session_id": session_id,
                "user_id": user_id
            },
            "",
            {"display": "none"}
        )
    else:
        # Échec de connexion
        print(f"   ❌ Échec de connexion pour {username}")

        # ✅ TRACKING: Logger les tentatives échouées
        if supabase_client:
            try:
                track_activity(None, None, "login_failed", page="login", details={"username_attempted": username})
            except:
                pass

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

    # ✅ TRACKING SUPABASE: Terminer la session
    if current_auth and supabase_client:
        try:
            username = current_auth.get("username")
            if username and username in ACTIVE_SESSIONS:
                session_info = ACTIVE_SESSIONS[username]
                # Tracker l'activité de logout
                track_activity(
                    session_info.get("user_id"),
                    session_info.get("session_id"),
                    "logout",
                    page="logout"
                )
                # Terminer la session
                end_session(session_info.get("session_id"))
                # Nettoyer
                del ACTIVE_SESSIONS[username]
        except Exception as e:
            print(f"⚠️ Erreur tracking logout: {e}")

    return {"authenticated": False, "username": None}


# ============================================================
# 🔐 FIN CALLBACKS AUTHENTIFICATION
# ============================================================


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

    # ==================== SIDEBAR COMPONENTS ====================
    #html.Div(id="risk-banner"),
    html.Div(id="action-feedback"),  # ✅ AJOUTER
    html.Div(id="selection-counter"),  # ✅ AJOUTER
    dcc.Input(id="search-input"),
    dcc.Dropdown(id="filter-supplier"),
    dcc.Dropdown(id="filter-category"),
    dcc.Dropdown(id="filter-need"),
    #dcc.Dropdown(id="filter-status", options=[
    #   {'label': 'Tous', 'value': 'all'},
    #  {'label': 'Stock OK', 'value': 'ok'},
    # {'label': 'Rupture', 'value': 'oos'}
    #], value='all'),
    dbc.Checklist(id="toggle-options"),
    dbc.Button(id="btn-refresh"),
    dbc.Button(id="btn-add-row"),
    dbc.Button(id="btn-po-pdf"),
    dbc.Button(id="btn-save-qac"),  # ✅ AJOUT
    dcc.Download(id="download-data"),
    dcc.Download(id="download-po"),
    html.Div(id="debug-info"),
    #html.Div(id="action-feedback"),
    #html.Div(id="selection-counter"),
    #dbc.Button("✅ Tout sélectionner (vue filtrée)", id="btn-select-all", size="sm", color="secondary", className="me-2"),
    ## Remplacez votre section boutons par celle-ci :
    html.Div([
        dbc.Button("🤖 Lancer Agent IA", id="btn-run-agent-ia", color="success", size="sm", className="me-2"),
        dbc.Button("📋 Remplir QAC = Target", id="btn-fill-qac-target", color="info", size="sm", className="me-2"),
        dbc.Button("☑️ Tout sélectionner", id="btn-select-all", color="secondary", size="sm", className="me-2"),
        dbc.Button("✖️ Désélectionner", id="btn-clear-selection", color="secondary", size="sm", className="me-2"),
        dbc.Button("📄 Bon de commande", id="btn-po-pdf", color="primary", size="sm", disabled=True),
        html.Div(id="selection-counter", className="d-inline-block ms-3")
    ], className="mb-3"),
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
            {"name": "product_name", "id": "product_name"},
            {"name": "Supplier", "id": "Supplier"},
            {"name": "total_stock", "id": "total_stock"},
            {"name": "QAC", "id": "QAC"},
            {"name": "QAC edited", "id": "QAC edited", "editable": True},  # ✅ AJOUT
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

@cache.memoize(timeout=3600)  # Cache 1h
def get_catalog_prices():
    """Charge les prix du catalogue (avec cache Redis)"""
    try:
        cat_url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTrpcAiktxAPBiwznGOh35kVetc4O8-z5rQdFDgBaDE4OC3Jnb7JDGm59c55Cwm2pWCktcsBirWT_0b/pub?gid=751531326&single=true&output=csv"

        print("🔄 Chargement catalogue prix...")
        catalog = pd.read_csv(cat_url, skiprows=1)
        catalog_subset = catalog.iloc[:, [0, 1, 3, 10]].copy()
        catalog_subset.columns = ['product_id', 'product_name', 'selling_price', 'purchase_price']
        catalog_subset['product_name_clean'] = catalog_subset['product_name'].astype(str).str.lower().str.strip()

        price_map = dict(zip(
            catalog_subset['product_name_clean'],
            pd.to_numeric(catalog_subset['purchase_price'], errors='coerce').fillna(1000)
        ))

        print(f"✅ Catalogue chargé : {len(price_map)} prix")
        return price_map

    except Exception as e:
        print(f"❌ Erreur catalogue : {e}")
        return {}


@cache.memoize(timeout=3600)
def get_packaging_map_cached():
    """Charge le packaging map (avec cache Redis)"""
    try:
        print("🔄 Chargement packaging map...")
        packaging = load_packaging_map()  # Ta fonction existante
        print(f"✅ Packaging map chargé : {len(packaging)} produits")
        return packaging
    except Exception as e:
        print(f"❌ Erreur packaging : {e}")
        return {}
# ====== IMPORTS NÉCESSAIRES ======

# ===== Helpers packaging (ROBUSTES) =====
import re, math, csv, os, io
from datetime import datetime

PACKAGING_URL = "https://data.heroku.com/dataclips/cnmhrqqjneeunkbqibxklyxwcsrl.csv"

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
'''
# ===== CALLBACK : AGENT IA (MODIFIÉ) =====
@app.callback(
    [Output("main-table", "data", allow_duplicate=True),
     Output("main-table", "selected_rows", allow_duplicate=True),
     Output("btn-po-pdf", "disabled", allow_duplicate=True)],  # ✅ Ajout Output
    Input("btn-run-agent-ia", "n_clicks"),
    State("main-table", "data"),
    prevent_initial_call=True
)
def run_agent_ia_calcul(n_clicks, table_data):
    """
    Lance l'Agent IA et active le bouton si succès
    """
    if not n_clicks or not table_data:
        return no_update, no_update, no_update

    # Calcul QAC
    USE_GEMINI = True
    qac_par_index = agent_ia_calculer_qac(table_data, use_gemini=USE_GEMINI)

    if not qac_par_index:
        print("⚠️ Aucun produit à commander")
        return no_update, no_update, True  # Désactiver bouton

    # Mise à jour tableau
    updated_data = table_data.copy()
    indices_selection = []

    for idx, qac_value in qac_par_index.items():
        if 0 <= idx < len(updated_data):
            updated_data[idx]['QAC edited'] = qac_value
            indices_selection.append(idx)

    print(f"✅ {len(qac_par_index)} QAC calculées, {len(indices_selection)} lignes sélectionnées")

    # Vérifier si on peut activer le bouton
    can_enable = all(
        updated_data[idx].get("product_name") and
        updated_data[idx].get("Supplier") and
        safe_float(updated_data[idx].get("QAC edited", 0), 0) > 0
        for idx in indices_selection
    )

    button_disabled = not can_enable

    return updated_data, sorted(indices_selection), button_disabled
'''
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
# FONCTION 3 : GÉNÉRATION EXCEL AVEC FORMULES
# ============================================
# ============================================
# FONCTION 3 : GÉNÉRATION EXCEL AVEC FORMULES + LOGO + SIGNATURE
# ============================================
def generer_bc_excel_avec_formules(selected_products, price_map, packaging_map):
    """
    Génère un Excel avec colonnes Remise/Escompte éditables,
    formules automatiques, LOGO en haut et SIGNATURE en bas
    """
    if not selected_products:
        raise ValueError("Aucun produit sélectionné")

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
    # 🎨 STYLES
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

    # ========================================
    # 🎨 LOGO MAAD EN HAUT
    # ========================================
    current_row = 1

    try:
        import os
        from openpyxl.drawing.image import Image as XLImage

        # Chercher le logo MAAD
        logo_paths = [
            "assets/logo_maad.png",
            "assets/logo_maad.jpg",
            "assets/logo.png",
            "assets/logo.jpg",
            "logo_maad.png",
            "logo.png",
            "logo.jpg"
        ]

        logo_path = None
        for path in logo_paths:
            if os.path.exists(path):
                logo_path = path
                break

        if logo_path:
            # Ajouter le logo
            img = XLImage(logo_path)

            # Redimensionner (largeur environ 180px, hauteur proportionnelle)
            original_width = img.width
            original_height = img.height
            img.width = 180
            img.height = int(180 * original_height / original_width) if original_width > 0 else 60

            # Positionner en A1
            ws.add_image(img, 'A1')

            print(f"   ✅ Logo MAAD ajouté : {logo_path}")

            # Ajuster hauteur des lignes pour le logo
            ws.row_dimensions[1].height = 25
            ws.row_dimensions[2].height = 25
            ws.row_dimensions[3].height = 25

            # Laisser de l'espace pour le logo
            current_row = 5
        else:
            # Si pas de logo, créer un en-tête texte stylé
            ws['A1'] = "MAAD"
            ws['A1'].font = Font(size=28, bold=True, color="1E40AF")
            ws['A2'] = "Marketplace Africain de Distribution"
            ws['A2'].font = Font(size=11, italic=True, color="64748B")
            ws.row_dimensions[1].height = 40
            print("   ⚠️ Logo non trouvé - En-tête texte créé")
            current_row = 4

    except Exception as e:
        print(f"   ⚠️ Erreur ajout logo Excel : {e}")
        ws['A1'] = "MAAD"
        ws['A1'].font = Font(size=28, bold=True, color="1E40AF")
        current_row = 3

    # ========================================
    # 📋 EN-TÊTE DU BON DE COMMANDE
    # ========================================

    po_number = get_next_po_number()

    # Ligne séparatrice
    current_row += 1

    # Titre du bon de commande
    ws[f'A{current_row}'] = f'BON DE COMMANDE N° {po_number}'
    ws[f'A{current_row}'].font = Font(size=18, bold=True, color="003366")
    ws.merge_cells(f'A{current_row}:I{current_row}')
    ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[current_row].height = 30
    current_row += 1

    # Ligne vide
    current_row += 1

    # Informations entreprise (à gauche)
    ws[f'A{current_row}'] = COMPANY_NAME
    ws[f'A{current_row}'].font = Font(bold=True, size=12, color="1E40AF")
    # Date (à droite)
    ws[f'G{current_row}'] = f"Date : {datetime.now().strftime('%d/%m/%Y')}"
    ws[f'G{current_row}'].font = Font(bold=True, size=11)
    ws[f'G{current_row}'].alignment = Alignment(horizontal='right')
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

    for supplier, products in suppliers.items():
        # Titre fournisseur EN MAJUSCULES
        supplier_upper = supplier.upper()
        ws[f'A{current_row}'] = f'📦 FOURNISSEUR : {supplier_upper}'
        ws[f'A{current_row}'].font = Font(size=14, bold=True, color="228B22")
        ws[f'A{current_row}'].fill = PatternFill(start_color="F0FFF0", end_color="F0FFF0", fill_type="solid")
        ws.merge_cells(f'A{current_row}:I{current_row}')
        ws.row_dimensions[current_row].height = 25
        current_row += 1

        # En-têtes colonnes (SANS Escompte par ligne)
        headers = [
            'Réf.',
            'Désignation',
            'Qté',
            'Unité',
            'PU HT',
            'Remise %',
            'Total HT',
            'TVA 18%',
            'Total TTC'
        ]

        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border

        # Colorier colonne Remise éditable
        ws.cell(row=current_row, column=6).fill = remise_fill

        current_row += 1
        first_data_row = current_row

        # ========================================
        # 📦 LIGNES PRODUITS
        # ========================================

        for prod in products:
            prod_name = str(prod.get("product_name", "")).strip()
            prod_key = prod_name.lower()

            # QAC edited OBLIGATOIRE
            qac_edited = safe_float(prod.get("QAC edited"), 0.0)
            if qac_edited <= 0:
                continue

            # Récupération données
            packaging_text = packaging_map.get(prod_key, "")
            qty_major = consolidate_to_major(qac_edited, packaging_text, prod_name)

            unit_price = safe_float(price_map.get(prod_key, 1000.0), 1000.0)
            if unit_price <= 0:
                unit_price = 1000.0

            # ✅ Récupérer product_id
            product_id = prod.get("product_id", "")

            if product_id:
                try:
                    product_id = int(float(product_id))
                except (ValueError, TypeError):
                    product_id = str(product_id)
            else:
                product_id = ""

            # ========================================
            # REMPLISSAGE CELLULES
            # ========================================

            # Colonne 1 : Réf. (product_id)
            cell = ws.cell(row=current_row, column=1, value=product_id)
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = Font(bold=True, size=10)

            # Colonne 2 : Désignation
            cell = ws.cell(row=current_row, column=2, value=prod_name[:50])
            cell.border = border
            cell.alignment = Alignment(horizontal='left', vertical='center')

            # Colonne 3 : Qté
            cell = ws.cell(row=current_row, column=3, value=qty_major)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
            cell.font = Font(bold=True)

            # Colonne 4 : Unité
            cell = ws.cell(row=current_row, column=4, value="unité")
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border

            # Colonne 5 : PU HT
            cell = ws.cell(row=current_row, column=5, value=unit_price)
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            # Colonne 6 : Remise % (ÉDITABLE par produit)
            cell = ws.cell(row=current_row, column=6, value=0)
            cell.number_format = '0.00'
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = remise_fill
            cell.border = border
            cell.font = Font(bold=True)

            # ========================================
            # FORMULES AUTOMATIQUES (sans escompte ligne)
            # ========================================

            # Colonne 7 : Total HT (formule avec remise seulement)
            # Total HT = (PU × Qté) × (1-Remise/100)
            formula_ht = f"=(E{current_row}*C{current_row})*(1-F{current_row}/100)"
            cell = ws.cell(row=current_row, column=7, value=formula_ht)
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            # Colonne 8 : TVA 18% (formule)
            formula_tva = f"=G{current_row}*0.18"
            cell = ws.cell(row=current_row, column=8, value=formula_tva)
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            # Colonne 9 : Total TTC (formule)
            formula_ttc = f"=G{current_row}+H{current_row}"
            cell = ws.cell(row=current_row, column=9, value=formula_ttc)
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
            ws.merge_cells(f'A{current_row}:F{current_row}')
            cell = ws.cell(row=current_row, column=1, value=f"Sous-total {supplier_upper}")
            cell.font = Font(bold=True, size=10)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            # Total HT brut
            cell = ws.cell(row=current_row, column=7, value=f"=SUM(G{first_data_row}:G{last_data_row})")
            cell.number_format = '#,##0'
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            # TVA brute
            cell = ws.cell(row=current_row, column=8, value=f"=SUM(H{first_data_row}:H{last_data_row})")
            cell.number_format = '#,##0'
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            # TTC brut
            cell = ws.cell(row=current_row, column=9, value=f"=SUM(I{first_data_row}:I{last_data_row})")
            cell.number_format = '#,##0'
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            subtotal_row = current_row
            current_row += 1

            # ========================================
            # 🎯 LIGNE ESCOMPTE FOURNISSEUR (ÉDITABLE)
            # ========================================
            ws.merge_cells(f'A{current_row}:E{current_row}')
            cell = ws.cell(row=current_row, column=1, value=f"🎯 Escompte {supplier_upper}")
            cell.font = Font(bold=True, size=10, color="E65100")
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            # Cellule escompte % (ÉDITABLE - colonne F)
            escompte_cell = ws.cell(row=current_row, column=6, value=0)
            escompte_cell.number_format = '0.00"%"'
            escompte_cell.alignment = Alignment(horizontal='center', vertical='center')
            escompte_cell.fill = escompte_fill
            escompte_cell.border = border
            escompte_cell.font = Font(bold=True, size=11)

            # Montant escompte HT (formule)
            cell = ws.cell(row=current_row, column=7, value=f"=-G{subtotal_row}*F{current_row}/100")
            cell.number_format = '#,##0'
            cell.font = Font(bold=True, color="E65100")
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            # Montant escompte TVA
            cell = ws.cell(row=current_row, column=8, value=f"=-H{subtotal_row}*F{current_row}/100")
            cell.number_format = '#,##0'
            cell.font = Font(bold=True, color="E65100")
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            # Montant escompte TTC
            cell = ws.cell(row=current_row, column=9, value=f"=-I{subtotal_row}*F{current_row}/100")
            cell.number_format = '#,##0'
            cell.font = Font(bold=True, color="E65100")
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border

            escompte_row = current_row
            current_row += 1

            # ========================================
            # 💚 TOTAL NET FOURNISSEUR (après escompte)
            # ========================================
            ws.merge_cells(f'A{current_row}:F{current_row}')
            cell = ws.cell(row=current_row, column=1, value=f"TOTAL NET {supplier.upper()}")
            cell.font = Font(bold=True, size=11)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border
            cell.fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")

            # Total HT net (sous-total + escompte)
            cell = ws.cell(row=current_row, column=7, value=f"=G{subtotal_row}+G{escompte_row}")
            cell.number_format = '#,##0'
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border
            cell.fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")

            # TVA nette
            cell = ws.cell(row=current_row, column=8, value=f"=H{subtotal_row}+H{escompte_row}")
            cell.number_format = '#,##0'
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = border
            cell.fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")

            # TTC net
            cell = ws.cell(row=current_row, column=9, value=f"=I{subtotal_row}+I{escompte_row}")
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

    # Total HT (colonne G maintenant)
    ws[f'F{current_row}'] = "TOTAL HT :"
    ws[f'F{current_row}'].font = Font(bold=True, size=12)
    ws[f'F{current_row}'].alignment = Alignment(horizontal='right')
    ws[f'G{current_row}'] = '=SUMIF(A:A,"TOTAL NET*",G:G)'
    ws[f'G{current_row}'].number_format = '#,##0 "FCFA"'
    ws[f'G{current_row}'].font = Font(bold=True, size=12)
    ws[f'G{current_row}'].border = Border(bottom=Side(style='thin'))

    current_row += 1

    # TVA (colonne H maintenant)
    ws[f'F{current_row}'] = "TVA (18%) :"
    ws[f'F{current_row}'].font = Font(bold=True, size=12)
    ws[f'F{current_row}'].alignment = Alignment(horizontal='right')
    ws[f'G{current_row}'] = '=SUMIF(A:A,"TOTAL NET*",H:H)'
    ws[f'G{current_row}'].number_format = '#,##0 "FCFA"'
    ws[f'G{current_row}'].font = Font(bold=True, size=12)
    ws[f'G{current_row}'].border = Border(bottom=Side(style='thin'))

    current_row += 1

    # Total TTC (colonne I maintenant)
    ws[f'F{current_row}'] = "TOTAL TTC :"
    ws[f'F{current_row}'].font = Font(bold=True, size=14, color="006400")
    ws[f'F{current_row}'].alignment = Alignment(horizontal='right')
    ws[f'G{current_row}'] = '=SUMIF(A:A,"TOTAL NET*",I:I)'
    ws[f'G{current_row}'].number_format = '#,##0 "FCFA"'
    ws[f'G{current_row}'].font = Font(bold=True, size=14, color="FFFFFF")
    ws[f'G{current_row}'].fill = PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid")
    ws[f'G{current_row}'].border = Border(
        top=Side(style='double'),
        bottom=Side(style='double')
    )

    # ========================================
    # ✍️ SIGNATURE EN BAS
    # ========================================

    current_row += 4  # Espace entre total et signature
    signature_row = current_row

    try:
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
        import traceback
        traceback.print_exc()

        # Fallback : Ligne pour signature
        ws.merge_cells(f'H{current_row}:J{current_row}')
        ws[f'H{current_row}'] = "_" * 30
        ws[f'H{current_row}'].alignment = Alignment(horizontal='center')
        current_row += 1
        ws[f'H{current_row}'] = "Signature et Cachet"
        ws[f'H{current_row}'].font = Font(bold=True, size=10)
        ws[f'H{current_row}'].alignment = Alignment(horizontal='center')

    # ========================================
    # 📐 MISE EN PAGE
    # ========================================

    # Largeurs colonnes (9 colonnes maintenant)
    ws.column_dimensions['A'].width = 10  # Réf
    ws.column_dimensions['B'].width = 40  # Désignation
    ws.column_dimensions['C'].width = 8   # Qté
    ws.column_dimensions['D'].width = 10  # Unité
    ws.column_dimensions['E'].width = 12  # PU HT
    ws.column_dimensions['F'].width = 12  # Remise % (par produit) / Escompte % (par fournisseur)
    ws.column_dimensions['G'].width = 15  # Total HT
    ws.column_dimensions['H'].width = 15  # TVA
    ws.column_dimensions['I'].width = 15  # Total TTC

    # Hauteur des premières lignes (pour le logo)
    ws.row_dimensions[1].height = 60
    ws.row_dimensions[2].height = 20

    # ========================================
    # 💾 EXPORT
    # ========================================

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    print(f"   ✅ Excel généré : BC-{po_number}")

    return buf, po_number

# ===== CALLBACK : GÉNÉRATION BON DE COMMANDE EXCEL AVEC VALIDATION =====
@app.callback(
    [Output("download-po", "data"),
     Output("po-loading-overlay", "is_open")],
    Input("btn-po-pdf", "n_clicks"),
    [State("main-table", "selected_rows"),
     State("main-table", "data")],
    prevent_initial_call=True
)
def export_po_excel_validated(n_clicks, selected_rows, table_data):
    """
    Génère un BC Excel SEULEMENT si toutes les QAC edited sont remplies
    """
    if not n_clicks or not table_data or not selected_rows:
        return no_update, False

    print("\n" + "=" * 60)
    print("📄 GÉNÉRATION BON DE COMMANDE EXCEL")
    print("=" * 60)

    # ========== VALIDATION STRICTE QAC ==========
    is_valid, error_msg, missing = valider_qac_selection(selected_rows, table_data)

    if not is_valid:
        print("❌ VALIDATION ÉCHOUÉE")
        print(error_msg)
        print("=" * 60 + "\n")
        # TODO : Afficher une alerte à l'utilisateur
        return no_update, False

    print("✅ VALIDATION OK - Toutes les QAC sont remplies")

    # ========== GÉNÉRATION ==========
    selected_products = [table_data[idx] for idx in selected_rows]
    print(f"📦 Produits : {len(selected_products)}")

    # Chargement données
    price_map = get_catalog_prices()
    packaging_map = get_packaging_map_cached()

    try:
        # Génération Excel avec formules
        excel_buffer, po_number = generer_bc_excel_avec_formules(
            selected_products, price_map, packaging_map
        )

        fname = f"BC_{po_number}_{len(selected_products)}p.xlsx"

        print("=" * 60)
        print(f"✅ Excel généré : {fname}")
        print(f"   ⚡ Colonnes Remise/Escompte éditables avec calculs auto")
        print("=" * 60 + "\n")

        return dcc.send_bytes(excel_buffer.read(), filename=fname), False

    except Exception as e:
        print(f"❌ Erreur génération : {e}")
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


# ===== CALLBACK 2 : AGENT IA =====
# ===== CALLBACK : AGENT IA (VERSION CORRIGÉE POUR VOTRE ARCHITECTURE) =====
@app.callback(
    Output("master-data", "data", allow_duplicate=True),
    Input("btn-run-agent-ia", "n_clicks"),
    State("master-data", "data"),
    prevent_initial_call=True
)
def run_agent_ia_calcul(n_clicks, master_json):
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


# ===== CALLBACK : PRÉSÉLECTION APRÈS AGENT IA =====
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
'''@app.callback(
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
)'''
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


import json

@app.callback(
    [Output("master-data", "data", allow_duplicate=True),
     Output("filtered-data", "data", allow_duplicate=True),
     Output("main-table", "data", allow_duplicate=True),
     Output("main-table", "selected_rows", allow_duplicate=True)],
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
    prevent_initial_call=True
)
def save_edit(n_clicks, prod, sup, cat, stock, active_cell, table_data, q, fs, fc, filter_opts, master_json):
    # If master_json is already a list (not a JSON string), use it directly
    if isinstance(master_json, str):
        base = pd.DataFrame(json.loads(master_json)) if master_json else get_df_cached()
    else:
        base = pd.DataFrame(master_json) if master_json else get_df_cached()

    df = base.copy()

    if active_cell and active_cell.get("column_id") == "edit Edit" and active_cell.get(
            "row") is not None and table_data:
        row = active_cell["row"]
        r = table_data[row]
        key_p = r.get("product_name")
        key_s = r.get("Supplier")
        idx = df[(df["product_name"].astype(str) == str(key_p)) & (df["Supplier"].astype(str) == str(key_s))].index
        if len(idx) > 0:
            i = idx[0]
            if prod is not None: df.at[i, "product_name"] = prod
            if sup is not None: df.at[i, "Supplier"] = sup
            if cat is not None: df.at[i, "Product Category"] = cat
            if stock is not None:
                try:
                    df.at[i, "total_stock"] = float(stock)
                except (ValueError, TypeError):
                    pass
    else:
        new_row = {c: np.nan for c in df.columns}
        new_row["product_name"] = prod or ""
        new_row["Supplier"] = sup or ""
        new_row["Product Category"] = cat or ""
        try:
            new_row["total_stock"] = float(stock or 0)
        except (ValueError, TypeError):
            new_row["total_stock"] = 0.0
        for c in ["optimal stock ", "Max Lead Time", "Max Avg Daily Sales", "Max Coverage Day",
                  "Daily OOS Rate (30d)", "Predicted Order Quantity"]:
            if c in df.columns and pd.isna(new_row.get(c)):
                new_row[c] = 0.0
        if "Predicted Stockout" in df.columns and pd.isna(new_row.get("Predicted Stockout")):
            new_row["Predicted Stockout"] = False
        if "Stock Status" in df.columns and pd.isna(new_row.get("Stock Status")):
            new_row["Stock Status"] = "Order Soon" if new_row["total_stock"] else "Out of Stock"
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    # Application des filtres
    sup_list = fs or []
    stat_list = []  # 'fst' n'était pas utilisé dans la fonction précédente
    cat_list = fc or []
    options = filter_opts or []

    # Vérification si fdf n'est pas vide avant application des filtres
    if not df.empty:
        fdf = filter_dataframe(df, q, sup_list, stat_list, cat_list, options)

        # Calcul du risque
        risk_count = int((fdf['Stock Status'].isin(
            ['Out of Stock', 'Predicted Stockout Soon']).sum())) if 'Stock Status' in fdf.columns else 0
    else:
        fdf = pd.DataFrame()  # DataFrame vide
        risk_count = 0

    # Application des actions sur fdf
    fdf_actions = add_action_cols(fdf)

    # Retour des résultats sous forme de JSON (4 outputs = 4 valeurs)
    return df.to_json(orient="records"), fdf_actions.to_json(orient="records"), fdf_actions.to_dict("records"), []

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
'''from waitress import serve
from supply_chain_analytics import app  # Assurez-vous que 'app' est bien l'instance de votre application Dash/Flask

if __name__ == "__main__":
    serve(app, host="0.0.0.0", port=10000)  # Lance Waitress sur le port 10000
'''
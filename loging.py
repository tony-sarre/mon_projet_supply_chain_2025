# login.py - Page de connexion indépendante
import os
from dash import Dash, html, dcc, Input, Output, State
import dash_bootstrap_components as dbc
from flask import session, redirect
import hashlib
from datetime import datetime

# Configuration
THEME = dbc.themes.CYBORG
SESSION_TIMEOUT = 28800  # 8 heures

# Base de données utilisateurs
USERS_DB = {
    "tony.sarre@maad.io": {
        "password_hash": hashlib.sha256("TonyMaad2025!".encode()).hexdigest(),
        "name": "Tony SARRE",
        "role": "admin"
    },
    "essodeke@maad.io": {
        "password_hash": hashlib.sha256("SamuelMaad2025!".encode()).hexdigest(),
        "name": "Samuel Essodeke",
        "role": "user"
    },
    "maimouna@maad.io": {
        "password_hash": hashlib.sha256("MaimounaMaad2025!".encode()).hexdigest(),
        "name": "Maimouna Dagois",
        "role": "user"
    },
    "seydouna@maad.io": {
        "password_hash": hashlib.sha256("SeydounaMaad2025!".encode()).hexdigest(),
        "name": "Seydouna Oumar Niang",
        "role": "user"
    },
    "arame.toure@maad.io": {
        "password_hash": hashlib.sha256("ArameMaad2025!".encode()).hexdigest(),
        "name": "Arame Toure",
        "role": "user"
    },
    "ndeyecoumba.cisse@maad.io": {
        "password_hash": hashlib.sha256("CoumbaMaad2025!".encode()).hexdigest(),
        "name": "Coumba Cisse",
        "role": "user"
    },
    "pr.diop@maad.io": {
        "password_hash": hashlib.sha256("RavaneMaad2025!".encode()).hexdigest(),
        "name": "Ravane Diop",
        "role": "user"
    },
    "serigne.diop@maad.io": {
        "password_hash": hashlib.sha256("FallouMaad2025!".encode()).hexdigest(),
        "name": "Fallou Diop",
        "role": "user"
    },
    "insa.niang@maad.io": {
        "password_hash": hashlib.sha256("InsaMaad2025!".encode()).hexdigest(),
        "name": "Insa Niang",
        "role": "user"
    }
}


# Fonctions d'authentification
def hash_password(password: str) -> str:
    """Hash un mot de passe avec SHA-256"""
    return hashlib.sha256(password.encode()).hexdigest()


def verify_password(email: str, password: str) -> bool:
    """Vérifie les credentials"""
    user = USERS_DB.get(email)
    if not user:
        return False
    return user["password_hash"] == hash_password(password)


def is_authenticated() -> bool:
    """Vérifie si l'utilisateur est connecté"""
    return session.get("authenticated", False)


def get_current_user():
    """Retourne les infos de l'utilisateur connecté"""
    email = session.get("user_email")
    if email and email in USERS_DB:
        return {
            "email": email,
            "name": USERS_DB[email]["name"],
            "role": USERS_DB[email]["role"]
        }
    return None


# Layout de la page de connexion
def login_layout():
    return html.Div(
        style={
            "minHeight": "100vh",
            "display": "flex",
            "alignItems": "center",
            "justifyContent": "center",
            "background": "linear-gradient(135deg, #0b1220 0%, #1a2332 100%)"
        },
        children=[
            html.Div(
                style={
                    "maxWidth": "420px",
                    "width": "100%",
                    "padding": "32px",
                    "background": "#0b1220",
                    "border": "1px solid #1f2937",
                    "borderRadius": "16px",
                    "boxShadow": "0 20px 60px rgba(0,0,0,.4)"
                },
                children=[
                    # Logo
                    html.Div(
                        style={"textAlign": "center", "marginBottom": "24px"},
                        children=[
                            html.Div("🔐", style={"fontSize": "48px", "marginBottom": "16px"}),
                            html.H2("Connexion",
                                    style={"color": "#e5e7eb", "marginBottom": "8px", "fontWeight": "700"}),
                            html.P("Supply Chain Command Center", style={"color": "#6b7280", "fontSize": "14px"})
                        ]
                    ),

                    # Formulaire
                    html.Div([
                        html.Label("Email", style={"color": "#9ca3af", "fontSize": "13px", "marginBottom": "6px",
                                                   "display": "block"}),
                        dbc.Input(
                            id="login-email",
                            placeholder="votre.email@maad.io",
                            type="email",
                            style={
                                "marginBottom": "16px",
                                "background": "#0a1320",
                                "border": "1px solid #1f2937",
                                "color": "#e5e7eb",
                                "borderRadius": "8px"
                            }
                        ),

                        html.Label("Mot de passe", style={"color": "#9ca3af", "fontSize": "13px", "marginBottom": "6px",
                                                          "display": "block"}),
                        dbc.Input(
                            id="login-password",
                            placeholder="••••••••",
                            type="password",
                            style={
                                "marginBottom": "20px",
                                "background": "#0a1320",
                                "border": "1px solid #1f2937",
                                "color": "#e5e7eb",
                                "borderRadius": "8px"
                            }
                        ),

                        dbc.Button(
                            "Se connecter",
                            id="login-button",
                            className="w-100",
                            style={
                                "background": "linear-gradient(135deg, #22d3ee 0%, #0ea5e9 100%)",
                                "border": "none",
                                "fontWeight": "700",
                                "padding": "12px",
                                "borderRadius": "8px",
                                "color": "#001018",
                                "fontSize": "15px"
                            }
                        ),

                        html.Div(id="login-error", style={"color": "#ef4444", "marginTop": "16px", "fontSize": "13px",
                                                          "textAlign": "center"})
                    ]),

                    # Footer
                    html.Div(
                        style={"marginTop": "32px", "textAlign": "center", "paddingTop": "24px",
                               "borderTop": "1px solid #1f2937"},
                        children=[
                            html.Small("© 2025 Maad SaSu • Supply Chain Analytics",
                                       style={"color": "#6b7280", "fontSize": "11px"})
                        ]
                    )
                ]
            )
        ]
    )


# Créer l'app de login
def create_login_app():
    app = Dash(__name__, external_stylesheets=[THEME], suppress_callback_exceptions=True)
    app.server.secret_key = os.getenv("FLASK_SECRET_KEY", "change-me-in-production")

    app.layout = html.Div([
        dcc.Location(id="url", refresh=True),
        dcc.Store(id="auth-redirect", storage_type='session'),
        html.Div(id="login-content", children=login_layout())
    ])

    # Callback de connexion
    @app.callback(
        Output("url", "pathname"),
        Output("login-error", "children"),
        [Input("login-button", "n_clicks"),
         Input("login-password", "n_submit")],
        [State("login-email", "value"),
         State("login-password", "value")],
        prevent_initial_call=True
    )
    def handle_login(n_clicks, n_submit, email, password):
        if not n_clicks and not n_submit:
            return "/login", ""

        if not email or not password:
            return "/login", "⚠️ Email et mot de passe requis"

        if verify_password(email, password):
            # Authentification réussie
            session["authenticated"] = True
            session["user_email"] = email
            session["login_time"] = datetime.now().isoformat()

            print(f"✅ Connexion réussie : {email}")
            return "/dashboard", ""
        else:
            print(f"❌ Échec connexion : {email}")
            return "/login", "❌ Email ou mot de passe incorrect"

    return app


if __name__ == "__main__":
    app = create_login_app()
    port = int(os.environ.get("PORT", 8050))
    app.run(debug=True, host="0.0.0.0", port=port)
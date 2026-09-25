import os
import sqlite3
import base64
import mimetypes
from datetime import date, datetime, timedelta
from contextlib import contextmanager
from io import BytesIO
from pathlib import Path
import urllib.parse

import pandas as pd
import streamlit as st

try:
    import openpyxl  # nécessaire pour les exports Excel
except ImportError:
    openpyxl = None

try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
    from psycopg2.pool import ThreadedConnectionPool
except ImportError:
    psycopg2 = None
    RealDictCursor = None
    ThreadedConnectionPool = None

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage

try:
    from twilio.rest import Client
except ImportError:
    Client = None

DB_PATH = Path("epargne_etudiant.db")

ADMIN_NAME = "Abdou Latif ALD"
ADMIN_USERNAME = "iy@2012"
ADMIN_PASSWORD = "issayoume2026"

WHATSAPP = "777521969"
COUNTRY_CODE = "221"
CURRENCY = "FCFA"

def secret_or_env(name, default=""):
    """Lit d'abord Streamlit Secrets, puis les variables d'environnement."""
    try:
        value = st.secrets.get(name, "")
    except Exception:
        value = ""
    return value or os.getenv(name, default)


TWILIO_ACCOUNT_SID = secret_or_env("TWILIO_ACCOUNT_SID")
TWILIO_AUTH_TOKEN = secret_or_env("TWILIO_AUTH_TOKEN")
TWILIO_WHATSAPP_FROM = secret_or_env(
    "TWILIO_WHATSAPP_FROM",
    "whatsapp:+221777521969"
)

# Configuration Supabase / PostgreSQL.
# On lit d'abord .streamlit/secrets.toml, puis les variables d'environnement.
# Supabase/PostgreSQL : accepte les deux formats de secrets :
#   [postgres] host/port/database/user/password/sslmode
# ou des clés plates SUPABASE_DB_* / SUPABASE_* .
def postgres_secret(name, default=""):
    try:
        section = st.secrets.get("postgres", {})
        if hasattr(section, "get"):
            value = section.get(name, "")
            if value not in (None, ""):
                return value
    except Exception:
        pass
    return secret_or_env(name, default)

SUPABASE_DB_URL = secret_or_env("SUPABASE_DB_URL", "")
SUPABASE_HOST = postgres_secret("host", "") or secret_or_env("SUPABASE_HOST", "")
SUPABASE_PORT = str(postgres_secret("port", "5432") or "5432")
SUPABASE_DATABASE = postgres_secret("database", "postgres") or secret_or_env("SUPABASE_DATABASE", "postgres")
SUPABASE_USER = postgres_secret("user", "") or secret_or_env("SUPABASE_USER", "")
SUPABASE_PASSWORD = postgres_secret("password", "") or secret_or_env("SUPABASE_PASSWORD", "")
SUPABASE_SSLMODE = postgres_secret("sslmode", "require") or "require"
SUPABASE_CONNECT_TIMEOUT = int(postgres_secret("connect_timeout", "10") or 10)

# Visuel de marque fourni pour l'application et les bulletins PDF.
ASSET_IMAGE = Path(__file__).with_name("peo.png")

BRAND_NAVY = "#122A55"
BRAND_BLUE = "#A9D4F5"
BRAND_PINK = "#F5A7B8"
BRAND_GREEN = "#1F7A6E"
BRAND_RED = "#A64B4B"
BRAND_CREAM = "#FBF8F2"
BRAND_GOLD = "#D89A2B"

st.set_page_config(
    page_title="Épargne Étudiant",
    page_icon="💰",
    layout="wide",
    menu_items={"Get help": None, "Report a bug": None, "About": None},
)

# Réduit la barre d'outils Streamlit lorsque cette option est disponible.
try:
    st.set_option("client.toolbarMode", "minimal")
except Exception:
    pass


# ============================================================
# IDENTITÉ VISUELLE
# ============================================================

def inject_brand_css():
    st.markdown(
        f"""
        <style>
        :root {{
            --brand-navy: {BRAND_NAVY};
            --brand-blue: {BRAND_BLUE};
            --brand-pink: {BRAND_PINK};
            --brand-green: {BRAND_GREEN};
            --brand-red: {BRAND_RED};
            --brand-cream: {BRAND_CREAM};
            --brand-gold: {BRAND_GOLD};
        }}

        .stApp {{
            background:
                radial-gradient(circle at 95% 0%, rgba(169,212,245,.22), transparent 28%),
                linear-gradient(180deg, #ffffff 0%, {BRAND_CREAM} 100%);
        }}

        [data-testid="stHeader"] {{
            background: rgba(255,255,255,.82);
        }}

        /* Masque la barre d'outils Streamlit (Share, GitHub, Edit, etc.) dans l'application. */
        [data-testid="stToolbar"],
        [data-testid="stDecoration"],
        [data-testid="stStatusWidget"],
        [data-testid="stAppDeployButton"] {{
            display: none !important;
            visibility: hidden !important;
        }}

        /* Masque le bouton « Gérer l'application » selon les versions de Streamlit. */
        button[kind="headerNoPadding"],
        [data-testid="stAppDeployButton"],
        [data-testid="stDeployButton"] {{
            display: none !important;
            visibility: hidden !important;
        }}

        [data-testid="stSidebar"] {{
            background: linear-gradient(180deg, {BRAND_NAVY} 0%, #1D3B68 62%, #274C79 100%);
        }}

        [data-testid="stSidebar"] * {{
            color: white !important;
        }}

        [data-testid="stSidebar"] .stButton > button {{
            border: 1px solid rgba(255,255,255,.22);
            background: rgba(255,255,255,.08);
            color: white !important;
        }}

        .brand-hero {{
            border-radius: 24px;
            padding: 28px 30px;
            margin: 4px 0 22px 0;
            background: linear-gradient(135deg, rgba(255,255,255,.97), rgba(238,247,253,.96));
            border: 1px solid rgba(18,42,85,.10);
            box-shadow: 0 14px 36px rgba(18,42,85,.09);
        }}

        .brand-kicker {{
            color: {BRAND_GREEN};
            font-weight: 800;
            letter-spacing: .08em;
            text-transform: uppercase;
            font-size: .82rem;
            margin-bottom: 6px;
        }}

        .brand-title {{
            color: {BRAND_NAVY};
            font-size: clamp(2rem, 4vw, 3.2rem);
            line-height: 1.02;
            font-weight: 900;
            margin: 0;
        }}

        .brand-subtitle {{
            color: #40536C;
            font-size: 1.08rem;
            line-height: 1.5;
            margin-top: 12px;
            margin-bottom: 0;
        }}

        /* Cadre unique pour toutes les photos : même hauteur, même ratio, même alignement. */
        .brand-photo-wrap {{
            width: 100%;
            aspect-ratio: 16 / 9;
            border-radius: 22px;
            overflow: hidden;
            background: linear-gradient(135deg, #eaf5fc, #fff5f7);
            border: 1px solid rgba(18,42,85,.10);
            box-shadow: 0 12px 28px rgba(18,42,85,.12);
        }}

        .brand-photo-wrap img {{
            width: 100%;
            height: 100%;
            display: block;
            object-fit: cover;
            object-position: center center;
        }}

        .photo-card {{
            width: 100%;
            aspect-ratio: 16 / 9;
            border-radius: 22px;
            overflow: hidden;
            background: linear-gradient(135deg, #eaf5fc, #fff5f7);
            box-shadow: 0 14px 32px rgba(18,42,85,.14);
            border: 1px solid rgba(18,42,85,.10);
        }}

        .photo-card img {{
            width: 100%;
            height: 100%;
            object-fit: cover;
            object-position: center;
            display: block;
        }}

        .section-title {{
            color: {BRAND_NAVY};
            font-weight: 850;
            font-size: 1.35rem;
            margin: 8px 0 12px 0;
        }}

        .metric-card {{
            background: rgba(255,255,255,.94);
            border: 1px solid rgba(18,42,85,.09);
            border-radius: 20px;
            padding: 18px 20px;
            min-height: 108px;
            box-shadow: 0 8px 24px rgba(18,42,85,.07);
        }}

        .metric-label {{
            color: #607089;
            font-size: .88rem;
            font-weight: 700;
        }}

        .metric-value {{
            color: {BRAND_NAVY};
            font-size: 1.65rem;
            font-weight: 900;
            margin-top: 5px;
        }}

        .info-card {{
            background: linear-gradient(135deg, rgba(169,212,245,.24), rgba(245,167,184,.18));
            border: 1px solid rgba(18,42,85,.08);
            border-radius: 18px;
            padding: 16px 18px;
            margin: 8px 0 18px 0;
        }}

        div[data-testid="stMetric"] {{
            background: rgba(255,255,255,.92);
            border: 1px solid rgba(18,42,85,.08);
            border-radius: 18px;
            padding: 12px 16px;
            box-shadow: 0 7px 22px rgba(18,42,85,.06);
        }}

        .stButton > button, .stDownloadButton > button {{
            border-radius: 12px;
            font-weight: 750;
            border: 1px solid rgba(18,42,85,.14);
        }}

        .stButton > button[kind="primary"] {{
            background: {BRAND_NAVY};
        }}

        .stDataFrame {{
            border-radius: 14px;
            overflow: hidden;
        }}

        div[data-testid="stExpander"] {{
            border-radius: 16px;
            border: 1px solid rgba(18,42,85,.10);
            background: rgba(255,255,255,.75);
        }}

        /* ---------- Écran de connexion ---------- */
        .login-shell {{
            max-width: 1160px;
            margin: 4vh auto 0 auto;
        }}

        .login-card {{
            background: rgba(255,255,255,.97);
            border: 1px solid rgba(18,42,85,.10);
            border-radius: 28px;
            padding: 28px;
            box-shadow: 0 22px 60px rgba(18,42,85,.12);
            height: 100%;
        }}

        .login-photo {{
            width: 100%;
            aspect-ratio: 4 / 3;
            border-radius: 22px;
            overflow: hidden;
            background: linear-gradient(135deg, #eaf5fc, #fff5f7);
            border: 1px solid rgba(18,42,85,.10);
        }}

        .login-photo img {{
            width: 100%;
            height: 100%;
            object-fit: cover;
            object-position: center;
            display: block;
        }}

        .login-badge {{
            display:inline-block;
            padding:7px 12px;
            border-radius:999px;
            background:rgba(31,122,110,.10);
            color:#1F7A6E;
            font-weight:800;
            font-size:.78rem;
            letter-spacing:.05em;
            text-transform:uppercase;
            margin-bottom:12px;
        }}

        .login-title {{
            color:#122A55;
            font-size:clamp(2.1rem,4vw,3.4rem);
            line-height:1.02;
            font-weight:950;
            margin-bottom:10px;
        }}

        .login-subtitle {{
            color:#53657d;
            font-size:1.05rem;
            line-height:1.55;
            margin-bottom:25px;
        }}

        .login-panel {{
            padding: 30px;
        }}

        .login-panel .stButton > button {{
            min-height:48px;
            border-radius:14px;
            font-size:1rem;
        }}

        .global-report-card {{
            border-radius: 22px;
            padding: 22px;
            background: linear-gradient(135deg, rgba(18,42,85,.97), rgba(39,76,121,.94));
            color: white;
            box-shadow: 0 16px 36px rgba(18,42,85,.15);
            margin: 8px 0 20px 0;
        }}

        @media (max-width: 800px) {{
            .login-shell {{ margin-top: 2vh; }}
            .login-card {{ padding:18px; border-radius:22px; }}
            .login-panel {{ padding:20px; }}
            .login-photo {{ aspect-ratio: 16 / 10; }}
            .brand-hero {{ padding:22px; }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def asset_image_html(css_class="brand-photo-wrap", alt="Épargne Étudiant"):
    """Retourne l'image de marque en HTML pour conserver exactement le même cadrage partout."""
    if not ASSET_IMAGE.exists():
        return ""
    try:
        mime = mimetypes.guess_type(str(ASSET_IMAGE))[0] or "image/jpeg"
        data = base64.b64encode(ASSET_IMAGE.read_bytes()).decode("ascii")
        return f'<div class="{css_class}"><img src="data:{mime};base64,{data}" alt="{alt}"></div>'
    except Exception:
        return ""


def brand_hero(title="Épargne Étudiant", subtitle="Petits efforts, grands projets !", compact=False):
    if ASSET_IMAGE.exists():
        left, right = st.columns([1.15, 0.85] if not compact else [1.5, 0.5])
        with left:
            st.markdown(
                f"""
                <div class="brand-hero">
                    <div class="brand-kicker">Épargne Étudiant</div>
                    <div class="brand-title">{title}</div>
                    <p class="brand-subtitle">{subtitle}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with right:
            st.image(str(ASSET_IMAGE), use_container_width=True)
    else:
        st.markdown(
            f"""
            <div class="brand-hero">
                <div class="brand-kicker">Épargne Étudiant</div>
                <div class="brand-title">{title}</div>
                <p class="brand-subtitle">{subtitle}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )


inject_brand_css()


# ============================================================
# BASE DE DONNÉES
# ============================================================


class PostgresCursorAdapter:
    """Adaptateur compatible avec sqlite3 tout en utilisant psycopg2."""

    def __init__(self, cursor):
        self.cursor = cursor

    def execute(self, query, params=None):
        self.cursor.execute(query.replace("?", "%s"), params)
        return self

    def executemany(self, query, params=None):
        self.cursor.executemany(query.replace("?", "%s"), params)
        return self

    def fetchone(self):
        return self.cursor.fetchone()

    def fetchmany(self, size=None):
        return self.cursor.fetchmany() if size is None else self.cursor.fetchmany(size)

    def fetchall(self):
        return self.cursor.fetchall()

    def close(self):
        return self.cursor.close()

    @property
    def description(self):
        return self.cursor.description

    @property
    def rowcount(self):
        return self.cursor.rowcount

    def __iter__(self):
        return iter(self.cursor)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()

    def __getattr__(self, name):
        return getattr(self.cursor, name)


class PostgresConnectionAdapter:
    """Connexion PostgreSQL avec une API proche de sqlite3."""

    def __init__(self, con):
        self._con = con

    def cursor(self, *args, **kwargs):
        return PostgresCursorAdapter(self._con.cursor(*args, **kwargs))

    def execute(self, query, params=None):
        cur = self.cursor()
        cur.execute(query, params)
        return cur

    def executemany(self, query, params=None):
        cur = self.cursor()
        cur.executemany(query, params)
        return cur

    def commit(self):
        return self._con.commit()

    def rollback(self):
        return self._con.rollback()

    def close(self):
        return self._con.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()

    def __getattr__(self, name):
        return getattr(self._con, name)


def postgres_dsn():
    """Construit la connexion Supabase/PostgreSQL."""
    if SUPABASE_DB_URL:
        return SUPABASE_DB_URL
    return (
        f"host={SUPABASE_HOST} "
        f"port={SUPABASE_PORT} "
        f"dbname={SUPABASE_DATABASE} "
        f"user={SUPABASE_USER} "
        f"password={SUPABASE_PASSWORD} "
        f"sslmode={SUPABASE_SSLMODE} "
        f"connect_timeout={SUPABASE_CONNECT_TIMEOUT}"
    )


def use_supabase():
    """Retourne True uniquement si PostgreSQL Supabase est réellement configuré."""
    configured = bool(
        SUPABASE_DB_URL
        or (SUPABASE_HOST and SUPABASE_USER and SUPABASE_PASSWORD)
    )
    return configured and psycopg2 is not None


def require_supabase():
    """Empêche silencieusement l'application de basculer vers SQLite."""
    if not use_supabase():
        missing = []
        if not SUPABASE_HOST: missing.append("host")
        if not SUPABASE_USER: missing.append("user")
        if not SUPABASE_PASSWORD: missing.append("password")
        if psycopg2 is None: missing.append("psycopg2-binary")
        details = ", ".join(missing) if missing else "configuration PostgreSQL"
        raise RuntimeError(
            "Supabase PostgreSQL n'est pas disponible. Vérifiez les secrets [postgres] "
            f"et requirements.txt ({details})."
        )


def supabase_health_check():
    """Teste réellement la connexion et la présence du schéma Supabase."""
    require_supabase()
    with db() as con:
        row = con.execute(
            "SELECT current_database() AS db, current_schema() AS schema, NOW() AS server_time"
        ).fetchone()
        tables = {}
        for table in ("members", "contributions", "loans", "loan_installments"):
            found = con.execute(
                "SELECT to_regclass(?) AS name", (f"public.{table}",)
            ).fetchone()
            tables[table] = bool(found and found.get("name"))
        return row, tables


@st.cache_resource(show_spinner=False)
def get_pg_pool(dsn=None):
    """
    Pool PostgreSQL partagé par l'application.
    Évite d'ouvrir une nouvelle connexion Supabase à chaque requête.
    """
    if not use_supabase() or ThreadedConnectionPool is None:
        return None

    dsn = dsn or postgres_dsn()
    return ThreadedConnectionPool(
        minconn=1,
        maxconn=5,
        dsn=dsn,
        cursor_factory=RealDictCursor,
        connect_timeout=8,
        application_name="epargne-etudiant",
        keepalives=1,
        keepalives_idle=30,
        keepalives_interval=10,
        keepalives_count=3,
    )


@contextmanager
def db():
    """
    Connexion PostgreSQL Supabase réutilisée via un pool.
    SQLite reste disponible uniquement comme secours local.
    """
    if use_supabase():
        pool = get_pg_pool(postgres_dsn())
        if pool is None:
            raise RuntimeError("Le pool PostgreSQL n'est pas disponible.")

        raw_con = pool.getconn()
        con = PostgresConnectionAdapter(raw_con)
        try:
            raw_cur = raw_con.cursor()
            raw_cur.execute("SET search_path TO public")
            raw_cur.close()
            yield con
            raw_con.commit()
        except Exception:
            raw_con.rollback()
            raise
        finally:
            # Nettoyage de la transaction avant de remettre la connexion
            # dans le pool. On ne ferme surtout pas la connexion ici.
            try:
                raw_con.rollback()
            except Exception:
                pass
            pool.putconn(raw_con)
    else:
        con = sqlite3.connect(DB_PATH)
        con.row_factory = sqlite3.Row
        try:
            yield con
            con.commit()
        except Exception:
            con.rollback()
            raise
        finally:
            con.close()

def sql(query):
    """Adapte les placeholders SQLite (?) vers PostgreSQL (%s)."""
    if use_supabase():
        return query.replace("?", "%s")
    return query


def read_sql(query, params=None):
    """Lecture SQL robuste sans passer par pandas.read_sql_query."""
    with db() as con:
        cur = con.cursor()
        cur.execute(sql(query), params or [])
        rows = cur.fetchall()
        description = cur.description or []
        columns = [d[0] for d in description]
        cur.close()
    if not rows:
        return pd.DataFrame(columns=columns)
    if isinstance(rows[0], dict):
        return pd.DataFrame(rows, columns=columns)
    return pd.DataFrame([tuple(r) for r in rows], columns=columns)


def migrate_database(con):
    """Migration de l'ancien schéma SQLite."""
    if use_supabase():
        return

    def columns(table):
        return {row["name"] for row in con.execute(f"PRAGMA table_info({table})").fetchall()}

    specs = {
        "members": {
            "phone": "TEXT",
            "monthly_target": "REAL DEFAULT 0",
            "notes": "TEXT",
            "active": "INTEGER DEFAULT 1",
            "member_username": "TEXT",
            "member_password": "TEXT",
            "member_login_active": "INTEGER DEFAULT 0",
            "created_at": "TEXT",
        },
        "contributions": {
            "payment_date": "TEXT",
            "month_label": "TEXT",
            "note": "TEXT",
            "created_at": "TEXT",
        },
        "loans": {
            "loan_date": "TEXT",
            "interest_rate": "REAL DEFAULT 0",
            "total_due": "REAL DEFAULT 0",
            "duration_months": "INTEGER DEFAULT 1",
            "first_due_date": "TEXT",
            "status": "TEXT DEFAULT 'Actif'",
            "note": "TEXT",
            "created_at": "TEXT",
            "total_interest_rate": "REAL DEFAULT 0",
            "installments_count": "INTEGER DEFAULT 1",
        },
        "loan_installments": {
            "installment_number": "INTEGER DEFAULT 1",
            "due_date": "TEXT",
            "amount_due": "REAL DEFAULT 0",
            "amount_paid": "REAL DEFAULT 0",
            "payment_date": "TEXT",
            "note": "TEXT",
            "created_at": "TEXT",
            "expected_amount": "REAL DEFAULT 0",
            "paid_date": "TEXT",
            "paid_amount": "REAL DEFAULT 0",
        },
        "loan_votes": {
            "loan_id": "INTEGER",
            "voter_member_id": "INTEGER",
            "decision": "TEXT DEFAULT 'En attente'",
            "comment": "TEXT",
            "created_at": "TEXT",
        },
        "member_messages": {
            "member_id": "INTEGER",
            "sender_role": "TEXT DEFAULT 'member'",
            "sender_member_id": "INTEGER",
            "subject": "TEXT",
            "message": "TEXT",
            "message_type": "TEXT DEFAULT 'message'",
            "is_read": "INTEGER DEFAULT 0",
            "created_at": "TEXT",
        },
        "member_reminders": {
            "member_id": "INTEGER",
            "reminder_type": "TEXT",
            "title": "TEXT",
            "message": "TEXT",
            "due_date": "TEXT",
            "whatsapp_sent": "INTEGER DEFAULT 0",
            "created_at": "TEXT",
        },
    }

    for table, fields in specs.items():
        existing = columns(table)
        for field, definition in fields.items():
            if field not in existing:
                con.execute(f"ALTER TABLE {table} ADD COLUMN {field} {definition}")

    contribution_cols = columns("contributions")
    if "payment_date" in contribution_cols:
        if "date" in contribution_cols:
            con.execute(
                "UPDATE contributions SET payment_date = date "
                "WHERE payment_date IS NULL OR payment_date = ''"
            )
        elif "contribution_date" in contribution_cols:
            con.execute(
                "UPDATE contributions SET payment_date = contribution_date "
                "WHERE payment_date IS NULL OR payment_date = ''"
            )
        con.execute(
            "UPDATE contributions SET payment_date = date('now') "
            "WHERE payment_date IS NULL OR payment_date = ''"
        )
        con.execute(
            "UPDATE contributions SET month_label = "
            "strftime('%m/%Y', payment_date) "
            "WHERE month_label IS NULL OR month_label = ''"
        )


@st.cache_resource(show_spinner=False)
def create_supabase_schema():
    """
    Prépare le schéma Supabase une seule fois par processus Streamlit.

    Correction importante :
    l'ancienne version créait l'index member_username AVANT d'ajouter
    member_username aux anciennes tables members. Si la table members
    existait déjà sans cette colonne, PostgreSQL arrêtait toute l'initialisation.
    """
    if not use_supabase():
        return

    with db() as con:
        cur = con.cursor()

        # ------------------------------------------------------------
        # 1. Création des tables si elles n'existent pas.
        # ------------------------------------------------------------
        table_statements = [
            """
            CREATE TABLE IF NOT EXISTS public.admins (
                id BIGSERIAL PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                full_name TEXT NOT NULL,
                active BOOLEAN NOT NULL DEFAULT TRUE,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS public.members (
                id BIGSERIAL PRIMARY KEY,
                full_name TEXT NOT NULL,
                phone TEXT,
                monthly_target NUMERIC(14,2) NOT NULL DEFAULT 0,
                notes TEXT,
                active BOOLEAN NOT NULL DEFAULT TRUE,
                member_username TEXT UNIQUE,
                member_password TEXT,
                member_login_active BOOLEAN NOT NULL DEFAULT FALSE,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS public.contributions (
                id BIGSERIAL PRIMARY KEY,
                member_id BIGINT NOT NULL REFERENCES public.members(id) ON DELETE CASCADE,
                payment_date DATE NOT NULL,
                month_label TEXT,
                amount NUMERIC(14,2) NOT NULL,
                note TEXT,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS public.loans (
                id BIGSERIAL PRIMARY KEY,
                member_id BIGINT NOT NULL REFERENCES public.members(id) ON DELETE CASCADE,
                loan_date DATE NOT NULL,
                principal NUMERIC(14,2) NOT NULL,
                interest_rate NUMERIC(8,4) NOT NULL DEFAULT 0,
                total_due NUMERIC(14,2) NOT NULL DEFAULT 0,
                duration_months INTEGER NOT NULL DEFAULT 1,
                first_due_date DATE NOT NULL,
                status TEXT NOT NULL DEFAULT 'Actif',
                note TEXT,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS public.loan_installments (
                id BIGSERIAL PRIMARY KEY,
                loan_id BIGINT NOT NULL REFERENCES public.loans(id) ON DELETE CASCADE,
                installment_number INTEGER NOT NULL DEFAULT 1,
                due_date DATE NOT NULL,
                amount_due NUMERIC(14,2) NOT NULL DEFAULT 0,
                amount_paid NUMERIC(14,2) NOT NULL DEFAULT 0,
                payment_date DATE,
                note TEXT,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS public.loan_votes (
                id BIGSERIAL PRIMARY KEY,
                loan_id BIGINT NOT NULL REFERENCES public.loans(id) ON DELETE CASCADE,
                voter_member_id BIGINT NOT NULL REFERENCES public.members(id) ON DELETE CASCADE,
                decision TEXT NOT NULL DEFAULT 'En attente',
                comment TEXT,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                UNIQUE(loan_id, voter_member_id)
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS public.member_messages (
                id BIGSERIAL PRIMARY KEY,
                member_id BIGINT NOT NULL REFERENCES public.members(id) ON DELETE CASCADE,
                sender_role TEXT NOT NULL DEFAULT 'member',
                sender_member_id BIGINT REFERENCES public.members(id) ON DELETE SET NULL,
                subject TEXT,
                message TEXT NOT NULL,
                message_type TEXT NOT NULL DEFAULT 'message',
                is_read BOOLEAN NOT NULL DEFAULT FALSE,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS public.member_reminders (
                id BIGSERIAL PRIMARY KEY,
                member_id BIGINT NOT NULL REFERENCES public.members(id) ON DELETE CASCADE,
                reminder_type TEXT NOT NULL,
                title TEXT NOT NULL,
                message TEXT NOT NULL,
                due_date DATE,
                whatsapp_sent BOOLEAN NOT NULL DEFAULT FALSE,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
            """,
        ]

        for statement in table_statements:
            cur.execute(statement)

        # ------------------------------------------------------------
        # 2. Migration douce des anciennes tables.
        # IMPORTANT : les colonnes sont ajoutées AVANT les index.
        # ------------------------------------------------------------
        alter_statements = [
            # members
            "ALTER TABLE public.members ADD COLUMN IF NOT EXISTS full_name TEXT",
            "ALTER TABLE public.members ADD COLUMN IF NOT EXISTS phone TEXT",
            "ALTER TABLE public.members ADD COLUMN IF NOT EXISTS monthly_target NUMERIC(14,2) DEFAULT 0",
            "ALTER TABLE public.members ADD COLUMN IF NOT EXISTS notes TEXT",
            "ALTER TABLE public.members ADD COLUMN IF NOT EXISTS active BOOLEAN DEFAULT TRUE",
            "ALTER TABLE public.members ADD COLUMN IF NOT EXISTS member_username TEXT",
            "ALTER TABLE public.members ADD COLUMN IF NOT EXISTS member_password TEXT",
            "ALTER TABLE public.members ADD COLUMN IF NOT EXISTS member_login_active BOOLEAN DEFAULT FALSE",
            "ALTER TABLE public.members ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ DEFAULT NOW()",

            # contributions
            "ALTER TABLE public.contributions ADD COLUMN IF NOT EXISTS member_id BIGINT",
            "ALTER TABLE public.contributions ADD COLUMN IF NOT EXISTS payment_date DATE",
            "ALTER TABLE public.contributions ADD COLUMN IF NOT EXISTS month_label TEXT",
            "ALTER TABLE public.contributions ADD COLUMN IF NOT EXISTS amount NUMERIC(14,2) DEFAULT 0",
            "ALTER TABLE public.contributions ADD COLUMN IF NOT EXISTS note TEXT",
            "ALTER TABLE public.contributions ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ DEFAULT NOW()",

            # loans
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS member_id BIGINT",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS loan_date DATE",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS principal NUMERIC(14,2) DEFAULT 0",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS interest_rate NUMERIC(8,4) DEFAULT 0",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS total_due NUMERIC(14,2) DEFAULT 0",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS duration_months INTEGER DEFAULT 1",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS first_due_date DATE",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS status TEXT DEFAULT 'Actif'",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS note TEXT",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ DEFAULT NOW()",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS total_interest_rate NUMERIC(8,4) DEFAULT 0",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS installments_count INTEGER DEFAULT 1",

            # loan_installments
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS loan_id BIGINT",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS installment_number INTEGER DEFAULT 1",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS due_date DATE",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS amount_due NUMERIC(14,2) DEFAULT 0",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS amount_paid NUMERIC(14,2) DEFAULT 0",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS payment_date DATE",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS note TEXT",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ DEFAULT NOW()",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS expected_amount NUMERIC(14,2) DEFAULT 0",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS paid_date DATE",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS paid_amount NUMERIC(14,2) DEFAULT 0",
        ]

        for statement in alter_statements:
            cur.execute(statement)

        # Valeurs par défaut pour les anciennes lignes.
        cur.execute("""
            UPDATE public.members
            SET active = COALESCE(active, TRUE),
                monthly_target = COALESCE(monthly_target, 0),
                member_login_active = COALESCE(member_login_active, FALSE)
            WHERE active IS NULL
               OR monthly_target IS NULL
               OR member_login_active IS NULL
        """)

        # Nettoyage des anciennes lignes d'en-tête importées par erreur
        # (ex. full_name / phone / monthly_target / notes). Elles ne sont
        # pas de vrais membres et provoquent des sélections impossibles.
        cur.execute("""
            DELETE FROM public.members
            WHERE lower(trim(COALESCE(full_name, ''))) IN ('full_name', 'nom complet', 'nom_complet')
              AND lower(trim(COALESCE(phone, ''))) IN ('phone', 'telephone', 'téléphone')
              AND lower(trim(COALESCE(notes, ''))) IN ('notes', 'note')
        """)

        # ------------------------------------------------------------
        # 3. Index APRES création/migration des colonnes.
        # ------------------------------------------------------------
        index_statements = [
            """
            CREATE UNIQUE INDEX IF NOT EXISTS idx_members_member_username_unique
            ON public.members(member_username)
            WHERE member_username IS NOT NULL
            """,
            """
            CREATE INDEX IF NOT EXISTS idx_members_active_name
            ON public.members(active, full_name)
            """,
            """
            CREATE INDEX IF NOT EXISTS idx_contributions_member_date
            ON public.contributions(member_id, payment_date DESC)
            """,
            """
            CREATE INDEX IF NOT EXISTS idx_loans_member_date
            ON public.loans(member_id, loan_date DESC)
            """,
            """
            CREATE INDEX IF NOT EXISTS idx_installments_loan_due
            ON public.loan_installments(loan_id, due_date)
            """,
            """
            CREATE UNIQUE INDEX IF NOT EXISTS idx_loan_votes_unique
            ON public.loan_votes(loan_id, voter_member_id)
            """,
            """
            CREATE INDEX IF NOT EXISTS idx_member_messages_member_created
            ON public.member_messages(member_id, created_at DESC)
            """,
            """
            CREATE INDEX IF NOT EXISTS idx_member_reminders_member_created
            ON public.member_reminders(member_id, created_at DESC)
            """,
        ]

        for statement in index_statements:
            cur.execute(statement)

        # ------------------------------------------------------------
        # 4. Administrateur initial.
        # ------------------------------------------------------------
        cur.execute(
            """
            INSERT INTO public.admins (username, password, full_name)
            VALUES (?, ?, ?)
            ON CONFLICT (username) DO NOTHING
            """,
            (ADMIN_USERNAME, ADMIN_PASSWORD, ADMIN_NAME),
        )

        # ------------------------------------------------------------
        # 5. Reprise des anciennes colonnes de prêts.
        # Une seule exécution au démarrage grâce au cache resource.
        # ------------------------------------------------------------
        try:
            # SAVEPOINT : une éventuelle erreur de compatibilité legacy
            # ne doit pas annuler les CREATE TABLE / ALTER TABLE / INDEX.
            cur.execute("SAVEPOINT legacy_upgrade")
            cur.execute("""
                UPDATE public.loans
                SET interest_rate = CASE
                        WHEN COALESCE(interest_rate, 0) = 0
                        THEN COALESCE(total_interest_rate, 0)
                        ELSE interest_rate
                    END,
                    duration_months = CASE
                        WHEN COALESCE(duration_months, 0) <= 0
                        THEN COALESCE(installments_count, 1)
                        ELSE duration_months
                    END,
                    total_due = CASE
                        WHEN COALESCE(total_due, 0) = 0
                        THEN principal * (1 + COALESCE(total_interest_rate, 0) / 100.0)
                        ELSE total_due
                    END,
                    status = COALESCE(NULLIF(status, ''), 'Actif')
            """)

            cur.execute("""
                UPDATE public.loan_installments
                SET amount_due = CASE
                        WHEN COALESCE(amount_due, 0) = 0
                        THEN COALESCE(expected_amount, 0)
                        ELSE amount_due
                    END,
                    amount_paid = CASE
                        WHEN COALESCE(amount_paid, 0) = 0
                        THEN COALESCE(paid_amount, 0)
                        ELSE amount_paid
                    END,
                    payment_date = COALESCE(payment_date, paid_date)
            """)
            cur.execute("RELEASE SAVEPOINT legacy_upgrade")
        except Exception:
            # Une ancienne structure peut ne pas posséder ces colonnes.
            cur.execute("ROLLBACK TO SAVEPOINT legacy_upgrade")
            cur.execute("RELEASE SAVEPOINT legacy_upgrade")

        cur.close()


def database_status():
    if use_supabase():
        return "Supabase PostgreSQL"
    return "SQLite local (secours)"



def sqlite_table_exists(con, table_name):
    return con.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,),
    ).fetchone() is not None


def auto_migrate_sqlite_to_supabase():
    """Importe automatiquement l'ancienne SQLite vers Supabase une seule fois."""
    if not use_supabase() or not DB_PATH.exists():
        return

    with db() as con:
        cur=con.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS public.app_migrations (
                key TEXT PRIMARY KEY,
                completed_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
        """)
        cur.execute("SELECT 1 FROM public.app_migrations WHERE key=?",
                    ("sqlite_to_supabase_v1",))
        if cur.fetchone():
            cur.close()
            return

    s=sqlite3.connect(DB_PATH)
    s.row_factory=sqlite3.Row
    try:
        with db() as pg:
            cur=pg.cursor()
            member_map={}
            if sqlite_table_exists(s,"members"):
                for m in s.execute("SELECT * FROM members ORDER BY id").fetchall():
                    legacy_name = str(m["full_name"] or "").strip().lower()
                    legacy_phone = str(m["phone"] or "").strip().lower() if "phone" in m.keys() else ""
                    legacy_notes = str(m["notes"] or "").strip().lower() if "notes" in m.keys() else ""
                    if legacy_name in {"", "full_name", "nom complet", "nom_complet"}:
                        continue
                    if legacy_name == "id" or (legacy_phone in {"phone", "telephone", "téléphone"} and legacy_notes in {"notes", "note"}):
                        continue
                    cur.execute("""
                        SELECT id FROM public.members
                        WHERE lower(trim(full_name))=lower(trim(?)) AND COALESCE(phone,'')=COALESCE(?,'')
                        LIMIT 1
                    """,(m["full_name"],m["phone"] if "phone" in m.keys() else None))
                    found=cur.fetchone()
                    if found:
                        member_map[m["id"]]=found["id"]
                    else:
                        cur.execute("""
                            INSERT INTO public.members
                            (full_name,phone,monthly_target,notes,active)
                            VALUES (?,?,?,?,?) RETURNING id
                        """,(
                            m["full_name"],
                            m["phone"] if "phone" in m.keys() else None,
                            m["monthly_target"] if "monthly_target" in m.keys() else 0,
                            m["notes"] if "notes" in m.keys() else None,
                            bool(m["active"]) if "active" in m.keys() else True,
                        ))
                        member_map[m["id"]]=cur.fetchone()["id"]

            if sqlite_table_exists(s,"contributions"):
                for c in s.execute("SELECT * FROM contributions ORDER BY id").fetchall():
                    mid=member_map.get(c["member_id"])
                    if not mid: continue
                    cur.execute("""
                        SELECT 1 FROM public.contributions
                        WHERE member_id=? AND payment_date=? AND amount=?
                        LIMIT 1
                    """,(mid,c["payment_date"],c["amount"]))
                    if not cur.fetchone():
                        cur.execute("""
                            INSERT INTO public.contributions
                            (member_id,payment_date,month_label,amount,note)
                            VALUES (?,?,?,?,?)
                        """,(
                            mid,c["payment_date"],c["month_label"],c["amount"],
                            c["note"] if "note" in c.keys() else None,
                        ))

            loan_map={}
            if sqlite_table_exists(s,"loans"):
                for l in s.execute("SELECT * FROM loans ORDER BY id").fetchall():
                    mid=member_map.get(l["member_id"])
                    if not mid: continue
                    cur.execute("""
                        SELECT id FROM public.loans
                        WHERE member_id=? AND loan_date=? AND principal=?
                        LIMIT 1
                    """,(mid,l["loan_date"],l["principal"]))
                    found=cur.fetchone()
                    if found:
                        loan_map[l["id"]]=found["id"]
                    else:
                        rate = float(l["total_interest_rate"] or 0) if "total_interest_rate" in l.keys() else 0
                        duration = int(l["installments_count"] or 1) if "installments_count" in l.keys() else 1
                        principal = float(l["principal"] or 0)
                        total_due = principal * (1 + rate / 100.0)
                        cur.execute("""
                            INSERT INTO public.loans
                            (member_id,loan_date,principal,interest_rate,total_due,
                             duration_months,first_due_date,status,note,total_interest_rate,installments_count)
                            VALUES (?,?,?,?,?,?,?,?,?,?,?) RETURNING id
                        """,(
                            mid,l["loan_date"],principal,rate,total_due,
                            duration,l["first_due_date"],"Actif",
                            l["note"] if "note" in l.keys() else None,
                            rate,duration,
                        ))
                        loan_map[l["id"]]=cur.fetchone()["id"]

            if sqlite_table_exists(s,"loan_installments"):
                for i in s.execute("SELECT * FROM loan_installments ORDER BY id").fetchall():
                    lid=loan_map.get(i["loan_id"])
                    if not lid: continue
                    cur.execute("""
                        SELECT 1 FROM public.loan_installments
                        WHERE loan_id=? AND installment_number=? LIMIT 1
                    """,(lid,i["installment_number"]))
                    if not cur.fetchone():
                        cur.execute("""
                            INSERT INTO public.loan_installments
                            (loan_id,installment_number,due_date,amount_due,
                             amount_paid,payment_date,note,expected_amount,paid_date,paid_amount)
                            VALUES (?,?,?,?,?,?,?,?,?,?)
                        """,(
                            lid,i["installment_number"],i["due_date"],
                            i["expected_amount"],i["paid_amount"],i["paid_date"],
                            i["note"] if "note" in i.keys() else None,
                            i["expected_amount"],i["paid_date"],i["paid_amount"],
                        ))

            cur.execute("""
                INSERT INTO public.app_migrations(key)
                VALUES (?)
                ON CONFLICT(key) DO NOTHING
            """,("sqlite_to_supabase_v1",))
            cur.close()
    finally:
        s.close()


@st.cache_resource(show_spinner=False)
def init_db(config_fingerprint=None):
    """Initialise exclusivement Supabase PostgreSQL."""
    require_supabase()
    create_supabase_schema()
    # La migration SQLite n'est lancée que si elle est explicitement activée.
    # Cela évite qu'une vieille base locale détourne ou ralentisse l'application.
    migrate_flag = str(secret_or_env("MIGRATE_SQLITE_TO_SUPABASE", "1")).strip().lower()
    if migrate_flag in {"1", "true", "yes", "oui"}:
        auto_migrate_sqlite_to_supabase()
    return

def safe_int_id(value):
    """Convertit un identifiant PostgreSQL/Pandas en entier sans faire planter l'application.

    Certaines lignes peuvent arriver sous forme de None, NaN, pd.NA ou texte
    après une lecture PostgreSQL/Pandas. Dans ce cas, on retourne None et la
    ligne concernée peut être ignorée dans les listes de sélection.
    """
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    try:
        if isinstance(value, str):
            value = value.strip()
            if not value:
                return None
        number = float(value)
        if not number.is_integer():
            return None
        return int(number)
    except (TypeError, ValueError, OverflowError):
        return None


def money(value):
    try:
        value = float(value or 0)
    except Exception:
        value = 0

    return f"{value:,.0f}".replace(",", " ") + f" {CURRENCY}"


def normalize_phone(phone):
    digits = "".join(
        ch for ch in str(phone or "")
        if ch.isdigit()
    )

    if digits.startswith("00"):
        digits = digits[2:]

    if len(digits) == 9:
        digits = COUNTRY_CODE + digits

    elif digits.startswith("0") and len(digits) == 10:
        digits = COUNTRY_CODE + digits[1:]

    return digits


def month_label(d):
    names = [
        "Janvier",
        "Février",
        "Mars",
        "Avril",
        "Mai",
        "Juin",
        "Juillet",
        "Août",
        "Septembre",
        "Octobre",
        "Novembre",
        "Décembre",
    ]

    return f"{names[d.month - 1]} {d.year}"


def add_months(d, months):
    month = d.month - 1 + months
    year = d.year + month // 12
    month = month % 12 + 1
    day = min(d.day, 28)

    return date(year, month, day)


# ============================================================
# AUTHENTIFICATION
# ============================================================

def authenticate(username, password):

    with db() as con:

        row = con.execute(
            """
            SELECT id, username, full_name
            FROM admins
            WHERE username=?
              AND password=?
              AND active=TRUE
            """,
            (
                username.strip(),
                password,
            )
        ).fetchone()

    return dict(row) if row else None


# Authentification d'un membre : le compte est strictement lié à un seul membre.
def authenticate_member(username, password):
    with db() as con:
        row = con.execute(
            """
            SELECT id, full_name, member_username
            FROM members
            WHERE member_username=?
              AND member_password=?
              AND member_login_active=TRUE
              AND active=TRUE
            LIMIT 1
            """,
            (username.strip(), password),
        ).fetchone()
    if not row:
        return None
    result = dict(row)
    result["role"] = "member"
    result["member_id"] = result["id"]
    return result


def set_member_login(member_id, username, password, active=True):
    username = username.strip()
    password = password.strip()
    if not username or not password:
        raise ValueError("Le nom d'utilisateur et le mot de passe du membre sont obligatoires.")
    with db() as con:
        # Empêche qu'un même identifiant soit attribué à deux membres.
        row = con.execute(
            "SELECT id FROM members WHERE member_username=? AND id<>? LIMIT 1",
            (username, member_id),
        ).fetchone()
        if row:
            raise ValueError("Cet identifiant est déjà utilisé par un autre membre.")
        con.execute(
            """
            UPDATE members
            SET member_username=?, member_password=?, member_login_active=?
            WHERE id=?
            """,
            (username, password, bool(active), member_id),
        )
        con.commit()


def get_member_login_status():
    return read_sql(
        """
        SELECT id, full_name, phone, member_username, member_login_active
        FROM members
        ORDER BY full_name
        """
    )


@st.cache_data(ttl=5, show_spinner=False)
def member_account_data_cached(member_id):
    member_id = int(member_id)
    mt = 'public.members' if use_supabase() else 'members'
    ct = 'public.contributions' if use_supabase() else 'contributions'
    lt = 'public.loans' if use_supabase() else 'loans'
    it = 'public.loan_installments' if use_supabase() else 'loan_installments'

    member_df = read_sql(f"SELECT id, full_name, phone, monthly_target, notes, active, member_username, member_login_active, created_at FROM {mt} WHERE id=? LIMIT 1", [member_id])
    cdf = read_sql(f"SELECT c.id, c.member_id, m.full_name, c.payment_date, c.month_label, c.amount, c.note FROM {ct} c LEFT JOIN {mt} m ON m.id=c.member_id WHERE c.member_id=? ORDER BY c.payment_date DESC, c.id DESC", [member_id])
    ldf = read_sql(f"SELECT l.id, l.member_id, m.full_name, l.loan_date, l.principal, l.interest_rate, l.total_due, l.duration_months, l.first_due_date, l.status, l.note FROM {lt} l LEFT JOIN {mt} m ON m.id=l.member_id WHERE l.member_id=? ORDER BY l.loan_date DESC, l.id DESC", [member_id])
    idf = read_sql(f"SELECT i.id, i.loan_id, i.installment_number, i.due_date, i.amount_due, i.amount_paid, i.payment_date, i.note FROM {it} i JOIN {lt} l ON l.id=i.loan_id WHERE l.member_id=? ORDER BY i.due_date, i.id", [member_id])

    cdf = _clean_contributions_df(cdf)
    for col in ('amount_due','amount_paid'):
        if col in idf.columns:
            idf[col] = pd.to_numeric(idf[col], errors='coerce').fillna(0)
    if not idf.empty:
        idf['reste'] = (idf['amount_due'] - idf['amount_paid']).clip(lower=0)

    total_contributed = float(pd.to_numeric(cdf.get('amount', pd.Series(dtype=float)), errors='coerce').fillna(0).sum())
    total_borrowed = float(pd.to_numeric(ldf.get('principal', pd.Series(dtype=float)), errors='coerce').fillna(0).sum())
    total_received = float(pd.to_numeric(idf.get('amount_paid', pd.Series(dtype=float)), errors='coerce').fillna(0).sum())
    total_due = float(pd.to_numeric(idf.get('amount_due', pd.Series(dtype=float)), errors='coerce').fillna(0).sum())

    return {'member': member_df, 'contributions': cdf, 'loans': ldf, 'installments': idf,
            'total_contributed': total_contributed, 'total_borrowed': total_borrowed,
            'total_received': total_received, 'outstanding': max(total_due-total_received,0)}

def member_account_data(member_id):
    """Retourne uniquement les données du membre authentifié."""
    return member_account_data_cached(int(member_id))

def member_account_page(member_id):
    data = member_account_data(member_id)
    if data["member"].empty:
        st.error("Compte membre introuvable.")
        return
    name = str(data["member"].iloc[0]["full_name"])
    brand_hero("Mon espace membre", f"Bienvenue {name}. Retrouvez vos opérations, vos rappels et participez aux validations.", compact=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total cotisé", money(data["total_contributed"]))
    c2.metric("Total emprunté", money(data["total_borrowed"]))
    c3.metric("Remboursements reçus", money(data["total_received"]))
    c4.metric("Reste à payer", money(data["outstanding"]))

    reminders = get_member_reminders(member_id)
    messages = get_member_messages(member_id)
    votes = loan_votes_for_member(member_id)

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "💰 Mes cotisations", "💳 Mes emprunts", "📅 Mes échéances",
        "🔔 Mes rappels", "💬 Messages & réclamations", "🗳️ Votes / droit de veto"
    ])
    with tab1:
        st.dataframe(data["contributions"], use_container_width=True, hide_index=True)
    with tab2:
        st.dataframe(data["loans"], use_container_width=True, hide_index=True)
    with tab3:
        st.dataframe(data["installments"], use_container_width=True, hide_index=True)

    with tab4:
        if reminders.empty:
            st.info("Aucun rappel pour le moment.")
        else:
            for _, r in reminders.iterrows():
                with st.container(border=True):
                    st.markdown(f"**{r['title']}**")
                    if pd.notna(r.get("due_date")) and str(r.get("due_date")):
                        st.caption(f"Échéance : {r['due_date']}")
                    st.write(str(r["message"]))
                    if bool(r.get("whatsapp_sent")):
                        st.caption("📱 Rappel également envoyé sur WhatsApp.")

    with tab5:
        st.subheader("✉️ Écrire à l'administration")
        with st.form("member_message_form"):
            subject = st.text_input("Objet", placeholder="Question, remarque, demande...")
            msg_type = st.selectbox("Type", ["Message", "Réclamation"])
            message = st.text_area("Votre message", height=150)
            send = st.form_submit_button("📨 Envoyer")
            if send:
                try:
                    send_member_message_to_admin(
                        member_id, subject, message,
                        "reclamation" if msg_type == "Réclamation" else "message"
                    )
                    st.success("Votre message a été transmis à l'administration.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))

        st.divider()
        st.subheader("📥 Échanges avec l'administration")
        if messages.empty:
            st.info("Aucun échange.")
        else:
            for _, r in messages.iterrows():
                sender = "👨‍💼 Administration" if r["sender_role"] == "admin" else "👤 Vous"
                label = "Réclamation" if r["message_type"] == "reclamation" else "Message"
                st.markdown(f"**{sender} — {label} — {r['subject'] or 'Sans objet'}**")
                st.write(str(r["message"]))
                st.caption(str(r["created_at"]))
                if not bool(r["is_read"]):
                    mark_message_read(r["id"])

    with tab6:
        st.info("Chaque membre actif, sauf le bénéficiaire, dispose d'un vote sur les nouveaux prêts. Un veto bloque la confirmation du prêt.")
        if votes.empty:
            st.success("Aucun prêt d'un autre membre ne vous attend actuellement.")
        else:
            for _, r in votes.iterrows():
                with st.container(border=True):
                    st.markdown(f"**Prêt #{int(r['loan_id'])} — {r['borrower_name']} — {money(r['principal'])}**")
                    st.caption(f"Date : {r['loan_date']} · Statut : {r['status']}")
                    current = str(r["decision"])
                    if current == "En attente":
                        with st.form(f"vote_form_{int(r['loan_id'])}"):
                            decision = st.radio("Votre décision", ["Approuvé", "Veto"], horizontal=True)
                            comment = st.text_area("Commentaire (facultatif)", height=90)
                            vote_btn = st.form_submit_button("Valider mon vote", type="primary")
                            if vote_btn:
                                try:
                                    new_status = cast_loan_vote(int(r["loan_id"]), member_id, decision, comment)
                                    st.success(f"Vote enregistré. Statut du prêt : {new_status}.")
                                    st.rerun()
                                except Exception as exc:
                                    st.error(str(exc))
                    else:
                        icon = "🛑" if current == "Veto" else "✅"
                        st.write(f"{icon} Votre vote : **{current}**")
                        if r.get("comment"):
                            st.caption(f"Commentaire : {r['comment']}")


# ============================================================
# MEMBRES
# ============================================================

def _clean_members_df(df):
    """Nettoie les lignes membres corrompues sans supprimer les vrais membres."""
    expected = [
        "id", "full_name", "phone", "monthly_target", "notes",
        "active", "member_username", "member_login_active", "created_at"
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=expected)

    df = df.copy()
    for col in expected:
        if col not in df.columns:
            df[col] = None

    def norm(v):
        if v is None:
            return ""
        try:
            if pd.isna(v):
                return ""
        except Exception:
            pass
        return str(v).strip().lower()

    # Supprime les anciennes lignes qui correspondent littéralement aux
    # noms de colonnes (visible dans Supabase : id/full_name/phone/...).
    bad_header_rows = (
        df["full_name"].map(norm).isin({"", "full_name", "nom complet", "nom_complet"})
        & df["phone"].map(norm).isin({"", "phone", "telephone", "téléphone"})
        & df["notes"].map(norm).isin({"", "notes", "note"})
    )

    # Une ligne sans nom ne doit jamais être proposée comme membre.
    missing_name = df["full_name"].map(norm).eq("")
    df = df.loc[~(bad_header_rows | missing_name)].copy()

    if df.empty:
        return pd.DataFrame(columns=expected)

    df["full_name"] = df["full_name"].astype(str).str.strip()
    df["phone"] = df["phone"].fillna("").astype(str).str.strip()
    return df.sort_values("full_name", kind="stable").reset_index(drop=True)


def build_member_options(df, include_phone=True):
    """Construit des choix stables et uniques : le nom affiché n'est jamais la clé DB."""
    options = {}
    if df is None or df.empty:
        return options

    for _, row in df.iterrows():
        member_id = safe_int_id(row.get("id"))
        name = str(row.get("full_name") or "").strip()
        if member_id is None or not name:
            continue
        phone = str(row.get("phone") or "").strip()
        label = f"{name} — {phone}" if include_phone and phone else name
        # L'ID rend le libellé unique même si deux membres ont le même nom/téléphone.
        label = f"{label} · ID {member_id}"
        options[label] = member_id
    return options


def refresh_application_data():
    """Force Streamlit à relire les données Supabase immédiatement."""
    try:
        member_account_data_cached.clear()
    except Exception:
        pass
    try:
        st.cache_data.clear()
    except Exception:
        pass


def get_members(active_only=False):
    """
    Lecture fraîche des membres depuis la base réellement utilisée.
    Important : cette fonction n'est pas mise en cache, afin qu'un membre
    ajouté dans Supabase apparaisse immédiatement après le rerun Streamlit.
    """
    table = "public.members" if use_supabase() else "members"

    if use_supabase():
        active_clause = "WHERE active IS NOT FALSE" if active_only else ""
    else:
        active_clause = "WHERE COALESCE(active, 1)=1" if active_only else ""

    query = f"""
        SELECT
            id,
            full_name,
            phone,
            monthly_target,
            notes,
            active,
            member_username,
            member_login_active,
            created_at
        FROM {table}
        {active_clause}
        ORDER BY lower(COALESCE(full_name, '')), id
    """

    return _clean_members_df(read_sql(query))


def add_member(name, phone, target, notes):
    """Ajoute un membre et vérifie immédiatement qu'il est réellement présent."""
    clean_name = str(name or '').strip()
    clean_phone = normalize_phone(phone)
    clean_notes = str(notes or '').strip()

    if not clean_name:
        raise ValueError("Le nom complet du membre est obligatoire.")
    if not clean_phone or len(clean_phone) != 12 or not clean_phone.startswith(COUNTRY_CODE):
        raise ValueError("Le numéro WhatsApp doit être un numéro sénégalais valide de 9 chiffres.")

    target_value = float(target or 0)
    if target_value < 0:
        raise ValueError("L'objectif mensuel ne peut pas être négatif.")

    table = 'public.members' if use_supabase() else 'members'
    with db() as con:
        existing_rows = con.execute(
            f"SELECT id, full_name, phone FROM {table} WHERE lower(trim(full_name))=lower(trim(?))",
            (clean_name,),
        ).fetchall()
        for existing in existing_rows:
            if normalize_phone(existing.get('phone') if hasattr(existing, 'get') else existing['phone']) == clean_phone:
                raise ValueError(f"Ce membre existe déjà (ID {existing['id']}).")

        if use_supabase():
            row = con.execute(
                f"""
                INSERT INTO {table}
                    (full_name, phone, monthly_target, notes, active, member_login_active)
                VALUES (?, ?, ?, ?, TRUE, FALSE)
                RETURNING id, full_name, phone, monthly_target, notes, active, created_at
                """,
                (clean_name, clean_phone, target_value, clean_notes),
            ).fetchone()
            new_id = safe_int_id(row.get('id') if row else None)
        else:
            cursor = con.execute(
                f"""INSERT INTO {table}
                    (full_name, phone, monthly_target, notes, active, member_login_active)
                   VALUES (?, ?, ?, ?, 1, 0)""",
                (clean_name, clean_phone, target_value, clean_notes),
            )
            new_id = safe_int_id(cursor.lastrowid)

        if new_id is None:
            raise RuntimeError("Le membre n'a pas pu être créé dans la base de données.")

        verify = con.execute(
            f"SELECT id FROM {table} WHERE id=? LIMIT 1",
            (new_id,),
        ).fetchone()
        if not verify:
            raise RuntimeError("La base n'a pas confirmé l'enregistrement du membre.")

        con.commit()
        return new_id


def update_member(member_id, name, phone, target, notes, active):
    member_id = safe_int_id(member_id)
    if member_id is None:
        raise ValueError("Identifiant membre invalide.")

    clean_name = str(name or '').strip()
    clean_phone = normalize_phone(phone)
    if not clean_name:
        raise ValueError("Le nom complet du membre est obligatoire.")
    if not clean_phone or len(clean_phone) != 12 or not clean_phone.startswith(COUNTRY_CODE):
        raise ValueError("Le numéro WhatsApp doit être un numéro sénégalais valide de 9 chiffres.")

    table = 'public.members' if use_supabase() else 'members'
    with db() as con:
        if use_supabase():
            row = con.execute(
                f"""UPDATE {table}
                    SET full_name=?, phone=?, monthly_target=?, notes=?, active=?
                    WHERE id=?
                    RETURNING id""",
                (clean_name, clean_phone, float(target or 0), str(notes or '').strip(), bool(active), member_id),
            ).fetchone()
            updated_id = safe_int_id(row.get('id') if row else None)
        else:
            cursor = con.execute(
                f"""UPDATE {table}
                    SET full_name=?, phone=?, monthly_target=?, notes=?, active=?
                    WHERE id=?""",
                (clean_name, clean_phone, float(target or 0), str(notes or '').strip(), int(bool(active)), member_id),
            )
            updated_id = member_id if cursor.rowcount else None

        if updated_id is None:
            raise ValueError("Membre introuvable.")
        con.commit()
        member_account_data_cached.clear()
        return updated_id


# ============================================================
# COTISATIONS
# ============================================================

def add_contribution(member_id, payment_date, amount, note):
    member_id = safe_int_id(member_id)
    amount = float(amount)
    if member_id is None:
        raise ValueError('Membre invalide.')
    if amount <= 0:
        raise ValueError('Le montant de la cotisation doit être supérieur à 0.')
    mt = 'public.members' if use_supabase() else 'members'
    ct = 'public.contributions' if use_supabase() else 'contributions'
    with db() as con:
        member = con.execute(f'SELECT id, full_name FROM {mt} WHERE id=? LIMIT 1', (member_id,)).fetchone()
        if not member:
            raise ValueError('Le membre sélectionné est introuvable dans la base utilisée par l’application.')
        if use_supabase():
            row = con.execute(f"INSERT INTO {ct}(member_id,payment_date,amount,month_label,note) VALUES (?,?,?,?,?) RETURNING id", (member_id,payment_date.isoformat(),amount,month_label(payment_date),str(note or '').strip())).fetchone()
            contribution_id = safe_int_id(row.get('id') if row else None)
        else:
            cur = con.execute(f"INSERT INTO {ct}(member_id,payment_date,amount,month_label,note) VALUES (?,?,?,?,?)", (member_id,payment_date.isoformat(),amount,month_label(payment_date),str(note or '').strip()))
            contribution_id = safe_int_id(cur.lastrowid)
        if contribution_id is None:
            raise RuntimeError('La cotisation a été écrite mais son identifiant n’a pas pu être récupéré.')
        verify = con.execute(f'SELECT id FROM {ct} WHERE id=? LIMIT 1', (contribution_id,)).fetchone()
        if not verify:
            raise RuntimeError('La base n’a pas confirmé la cotisation.')
        con.commit()
    refresh_application_data()
    return contribution_id


def update_contribution(contribution_id, payment_date, amount, note):
    contribution_id = safe_int_id(contribution_id)
    amount = float(amount)
    if contribution_id is None or amount <= 0:
        raise ValueError("Cotisation ou montant invalide.")
    table = 'public.contributions' if use_supabase() else 'contributions'
    with db() as con:
        cur = con.execute(
            f"UPDATE {table} SET payment_date=?, amount=?, month_label=?, note=? WHERE id=?",
            (payment_date.isoformat(), amount, month_label(payment_date), str(note or '').strip(), contribution_id)
        )
        if cur.rowcount == 0:
            raise ValueError("Cotisation introuvable.")
        con.commit()
    refresh_application_data()


def delete_contribution(contribution_id):
    contribution_id = safe_int_id(contribution_id)
    if contribution_id is None:
        raise ValueError("Cotisation invalide.")
    table = 'public.contributions' if use_supabase() else 'contributions'
    with db() as con:
        con.execute(f"DELETE FROM {table} WHERE id=?", (contribution_id,))
        con.commit()
    refresh_application_data()

def _clean_contributions_df(df):
    """Nettoie les éventuelles lignes d'en-tête importées par erreur."""
    columns = [
        "id", "member_id", "full_name", "payment_date",
        "month_label", "amount", "note"
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=columns)

    df = df.copy()
    for col in columns:
        if col not in df.columns:
            df[col] = None

    def norm(value):
        if value is None:
            return ""
        try:
            if pd.isna(value):
                return ""
        except Exception:
            pass
        return str(value).strip().lower()

    # Cette ligne apparaît lorsque les noms de colonnes ont été enregistrés
    # accidentellement comme une vraie cotisation. Elle ne doit jamais être
    # affichée comme une cotisation réelle.
    bad_header = (
        df["id"].map(norm).eq("id")
        & df["member_id"].map(norm).eq("member_id")
        & df["full_name"].map(norm).eq("full_name")
        & df["payment_date"].map(norm).eq("payment_date")
        & df["month_label"].map(norm).eq("month_label")
        & df["amount"].map(norm).eq("amount")
        & df["note"].map(norm).eq("note")
    )
    df = df.loc[~bad_header].copy()

    # Une cotisation doit obligatoirement être rattachée à un membre réel.
    valid_member_ids = set()
    try:
        members_df = get_members(False)
        valid_member_ids = {
            mid for mid in (safe_int_id(v) for v in members_df["id"].tolist())
            if mid is not None
        }
    except Exception:
        valid_member_ids = set()

    if valid_member_ids:
        parsed_ids = df["member_id"].map(safe_int_id)
        df = df.loc[parsed_ids.isin(valid_member_ids)].copy()
        df["member_id"] = parsed_ids.loc[df.index].astype(int)

    df["full_name"] = df["full_name"].fillna("").astype(str).str.strip()
    df["payment_date"] = df["payment_date"].fillna("").astype(str).str.strip()
    df["month_label"] = df["month_label"].fillna("").astype(str).str.strip()
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
    df["note"] = df["note"].fillna("").astype(str)

    return df.reset_index(drop=True)


def contributions(member_id=None):
    ct = 'public.contributions' if use_supabase() else 'contributions'
    mt = 'public.members' if use_supabase() else 'members'
    query = f"SELECT c.id,c.member_id,m.full_name,c.payment_date,c.month_label,c.amount,c.note FROM {ct} c LEFT JOIN {mt} m ON m.id=c.member_id"
    params=[]
    if member_id is not None:
        query += ' WHERE c.member_id=?'
        params.append(int(member_id))
    query += ' ORDER BY c.payment_date DESC,c.id DESC'
    return _clean_contributions_df(read_sql(query, params))


# ============================================================
# EMPRUNTS
# ============================================================

def create_loan(
    member_id, loan_date, principal, rate, duration, first_due_date, note
):
    member_id = safe_int_id(member_id)
    principal = float(principal)
    rate = float(rate)
    duration = int(duration)

    if member_id is None:
        raise ValueError("Membre invalide.")
    if principal <= 0:
        raise ValueError("Le montant du prêt doit être supérieur à 0.")
    if rate < 0:
        raise ValueError("Le taux ne peut pas être négatif.")
    if duration < 1 or duration > 60:
        raise ValueError("Le nombre d'échéances doit être compris entre 1 et 60.")

    total_due = principal * (1 + rate / 100.0)
    installment = total_due / duration

    with db() as con:
        member_table = 'public.members' if use_supabase() else 'members'
        active_clause = "COALESCE(active, TRUE)=TRUE" if use_supabase() else "COALESCE(active, 1)=1"
        member = con.execute(
            f"SELECT id FROM {member_table} WHERE id=? AND {active_clause} LIMIT 1",
            (member_id,),
        ).fetchone()
        if not member:
            raise ValueError("Le membre sélectionné n'existe pas ou est inactif.")

        if use_supabase():
            cursor = con.execute(
                """
                INSERT INTO public.loans(
                    member_id, loan_date, principal, interest_rate, total_due,
                    duration_months, first_due_date, status, note,
                    total_interest_rate, installments_count
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, 'Actif', ?, ?, ?)
                RETURNING id
                """,
                (member_id, loan_date.isoformat(), principal, rate, total_due, duration,
                 first_due_date.isoformat(), str(note or '').strip(), rate, duration),
            )
            loan_row = cursor.fetchone()
            loan_id = safe_int_id(loan_row['id']) if loan_row else None
        else:
            cursor = con.execute(
                """
                INSERT INTO loans(
                    member_id, loan_date, principal, interest_rate, total_due,
                    duration_months, first_due_date, status, note
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, 'Actif', ?)
                """,
                (member_id, loan_date.isoformat(), principal, rate, total_due, duration,
                 first_due_date.isoformat(), str(note or '').strip()),
            )
            loan_id = safe_int_id(cursor.lastrowid)

        if loan_id is None:
            raise RuntimeError("Impossible de récupérer l'identifiant du prêt créé.")

        installment_table = 'public.loan_installments' if use_supabase() else 'loan_installments'
        for i in range(duration):
            due_date = add_months(first_due_date, i)
            amount = total_due - installment * (duration - 1) if i == duration - 1 else installment
            con.execute(
                f"""INSERT INTO {installment_table}(
                    loan_id, installment_number, due_date, amount_due, amount_paid
                ) VALUES (?, ?, ?, ?, 0)""",
                (loan_id, i + 1, due_date.isoformat(), round(amount, 2)),
            )

        con.commit()
        loan_table = 'public.loans' if use_supabase() else 'loans'
        verify = con.execute(f"SELECT id FROM {loan_table} WHERE id=? LIMIT 1", (loan_id,)).fetchone()
        if not verify:
            raise RuntimeError('La base n’a pas confirmé le prêt.')
    ensure_loan_votes(loan_id)
    _, vote_status = loan_vote_status(loan_id)
    with db() as con:
        con.execute(f"UPDATE {loan_table} SET status=? WHERE id=?", (vote_status, loan_id))
        con.commit()
    refresh_application_data()
    return loan_id


def update_loan(loan_id, member_id, loan_date, principal, rate, duration, first_due_date, note):
    loan_id = safe_int_id(loan_id); member_id = safe_int_id(member_id)
    principal = float(principal); rate = float(rate); duration = int(duration)
    if loan_id is None or member_id is None:
        raise ValueError("Prêt ou membre invalide.")
    if principal <= 0 or rate < 0 or duration < 1 or duration > 60:
        raise ValueError("Valeurs du prêt invalides.")
    lt = 'public.loans' if use_supabase() else 'loans'
    it = 'public.loan_installments' if use_supabase() else 'loan_installments'
    total_due = principal * (1 + rate / 100.0)
    with db() as con:
        loan = con.execute(f"SELECT id FROM {lt} WHERE id=?", (loan_id,)).fetchone()
        if not loan:
            raise ValueError("Prêt introuvable.")
        paid_row = con.execute(f"SELECT COALESCE(SUM(amount_paid),0) AS paid, COUNT(*) AS n FROM {it} WHERE loan_id=? AND COALESCE(amount_paid,0)>0", (loan_id,)).fetchone()
        paid_total = float(paid_row['paid'] or 0)
        paid_count = int(paid_row['n'] or 0)
        old_count_row = con.execute(f"SELECT COUNT(*) AS n FROM {it} WHERE loan_id=?", (loan_id,)).fetchone()
        old_count = int(old_count_row['n'] or 0)
        if paid_count and duration != old_count:
            raise ValueError("Le nombre d'échéances ne peut pas être modifié après le début des remboursements. Modifiez les échéances une par une.")
        if total_due + 0.01 < paid_total:
            raise ValueError("Le nouveau total dû ne peut pas être inférieur aux remboursements déjà enregistrés.")
        con.execute(f"UPDATE {lt} SET member_id=?, loan_date=?, principal=?, interest_rate=?, total_due=?, duration_months=?, first_due_date=?, note=?, total_interest_rate=?, installments_count=? WHERE id=?",
                    (member_id, loan_date.isoformat(), principal, rate, total_due, duration, first_due_date.isoformat(), str(note or '').strip(), rate, duration, loan_id))
        rows = con.execute(f"SELECT id, installment_number, amount_paid FROM {it} WHERE loan_id=? ORDER BY installment_number", (loan_id,)).fetchall()
        # Si aucun remboursement n'a commencé, la modification du nombre ou du début
        # du calendrier reconstruit proprement les échéances.
        if paid_count == 0:
            con.execute(f"DELETE FROM {it} WHERE loan_id=?", (loan_id,))
            installment = total_due / duration
            for i in range(duration):
                due_date = add_months(first_due_date, i)
                amt = total_due - installment * (duration - 1) if i == duration - 1 else installment
                con.execute(f"INSERT INTO {it}(loan_id, installment_number, due_date, amount_due, amount_paid) VALUES (?,?,?,?,0)",
                            (loan_id, i + 1, due_date.isoformat(), round(amt,2)))
        else:
            # Après le début des remboursements, les lignes déjà payées restent intactes.
            unpaid = [r for r in rows if float(r['amount_paid'] or 0) <= 0.01]
            remaining = max(total_due - paid_total, 0)
            if unpaid:
                each = remaining / len(unpaid)
                for idx, r in enumerate(unpaid):
                    amt = remaining - each * (len(unpaid)-1) if idx == len(unpaid)-1 else each
                    con.execute(f"UPDATE {it} SET amount_due=? WHERE id=?", (round(amt,2), r['id']))
        con.commit()
    refresh_application_data()


def update_installment(installment_id, due_date, amount_due, amount_paid, payment_date, note):
    installment_id = safe_int_id(installment_id)
    if installment_id is None:
        raise ValueError("Échéance invalide.")
    amount_due = float(amount_due); amount_paid = float(amount_paid)
    if amount_due <= 0 or amount_paid < 0:
        raise ValueError("Les montants de l'échéance sont invalides.")
    if amount_paid > amount_due + 0.01:
        raise ValueError("Le montant remboursé ne peut pas dépasser le montant de l'échéance.")
    it = 'public.loan_installments' if use_supabase() else 'loan_installments'
    lt = 'public.loans' if use_supabase() else 'loans'
    with db() as con:
        row = con.execute(f"SELECT loan_id FROM {it} WHERE id=?", (installment_id,)).fetchone()
        if not row:
            raise ValueError("Échéance introuvable.")
        loan_id = row['loan_id']
        con.execute(f"UPDATE {it} SET due_date=?, amount_due=?, amount_paid=?, payment_date=?, note=? WHERE id=?",
                     (due_date.isoformat(), amount_due, amount_paid, payment_date.isoformat() if payment_date else None, str(note or '').strip(), installment_id))
        total = con.execute(f"SELECT COALESCE(SUM(amount_due),0) AS due, COALESCE(SUM(amount_paid),0) AS paid FROM {it} WHERE loan_id=?", (loan_id,)).fetchone()
        status = "Remboursé" if float(total['paid'] or 0) >= float(total['due'] or 0)-0.01 else "Actif"
        con.execute(f"UPDATE {lt} SET status=? WHERE id=?", (status, loan_id))
        con.commit()
    refresh_application_data()

def loans(member_id=None):
    lt = 'public.loans' if use_supabase() else 'loans'
    mt = 'public.members' if use_supabase() else 'members'
    query = f"SELECT l.id,l.member_id,m.full_name,l.loan_date,l.principal,l.interest_rate,l.total_due,l.duration_months,l.first_due_date,l.status,l.note FROM {lt} l LEFT JOIN {mt} m ON m.id=l.member_id"
    params=[]
    if member_id is not None:
        query += ' WHERE l.member_id=?'
        params.append(int(member_id))
    query += ' ORDER BY l.loan_date DESC,l.id DESC'
    return read_sql(query, params)


def get_installments(loan_id):
    it = 'public.loan_installments' if use_supabase() else 'loan_installments'
    return read_sql(f"SELECT id,loan_id,installment_number,due_date,amount_due,amount_paid,payment_date,note FROM {it} WHERE loan_id=? ORDER BY due_date,id", [int(loan_id)])


def register_installment_payment(
    installment_id,
    amount_paid,
    payment_date,
    note
):

    with db() as con:

        it = 'public.loan_installments' if use_supabase() else 'loan_installments'
        lt = 'public.loans' if use_supabase() else 'loans'
        row = con.execute(
            f"SELECT loan_id FROM {it} WHERE id=?",
            (installment_id,)
        ).fetchone()

        if not row:
            raise ValueError(
                "Échéance introuvable."
            )

        loan_id = row["loan_id"]

        current = con.execute(f"SELECT amount_due FROM {it} WHERE id=?", (installment_id,)).fetchone()
        if not current:
            raise ValueError("Échéance introuvable.")
        if float(amount_paid) < 0 or float(amount_paid) > float(current['amount_due'] or 0) + 0.01:
            raise ValueError("Le montant remboursé doit être compris entre 0 et le montant de l'échéance.")

        con.execute(
            f"UPDATE {it} SET amount_paid=?, payment_date=?, note=? WHERE id=?",
            (
                float(amount_paid),
                payment_date.isoformat(),
                note.strip(),
                installment_id,
            )
        )

        total = con.execute(
            f"SELECT SUM(amount_due) AS due, SUM(amount_paid) AS paid FROM {it} WHERE loan_id=?",
            (loan_id,)
        ).fetchone()

        due = float(total["due"] or 0)
        paid = float(total["paid"] or 0)

        status = (
            "Remboursé"
            if paid >= due - 0.01
            else "Actif"
        )

        con.execute(
            f"UPDATE {lt} SET status=? WHERE id=?",
            (
                status,
                loan_id,
            )
        )

        con.commit()
    refresh_application_data()



# ============================================================
# COMMUNICATION MEMBRES / VALIDATION DES PRÊTS
# ============================================================

def _table(name):
    return f"public.{name}" if use_supabase() else name


def create_member_message(member_id, message, subject="", message_type="message",
                          sender_role="member", sender_member_id=None):
    member_id = safe_int_id(member_id)
    if member_id is None or not str(message or "").strip():
        raise ValueError("Le membre et le message sont obligatoires.")
    table = _table("member_messages")
    with db() as con:
        con.execute(
            f"""INSERT INTO {table}
                (member_id, sender_role, sender_member_id, subject, message, message_type, is_read)
                VALUES (?, ?, ?, ?, ?, ?, FALSE)""",
            (member_id, sender_role, safe_int_id(sender_member_id),
             str(subject or "").strip(), str(message).strip(), message_type),
        )
        con.commit()


def get_member_messages(member_id=None, unread_only=False):
    table = _table("member_messages")
    mt = _table("members")
    query = f"""
        SELECT mm.id, mm.member_id, m.full_name AS member_name,
               mm.sender_role, mm.sender_member_id, mm.subject, mm.message,
               mm.message_type, mm.is_read, mm.created_at
        FROM {table} mm
        LEFT JOIN {mt} m ON m.id=mm.sender_member_id
    """
    clauses, params = [], []
    if member_id is not None:
        clauses.append("mm.member_id=?")
        params.append(int(member_id))
    if unread_only:
        clauses.append("COALESCE(mm.is_read,FALSE)=FALSE")
    if clauses:
        query += " WHERE " + " AND ".join(clauses)
    query += " ORDER BY mm.created_at DESC, mm.id DESC"
    return read_sql(query, params)


def mark_message_read(message_id):
    table = _table("member_messages")
    with db() as con:
        con.execute(f"UPDATE {table} SET is_read=TRUE WHERE id=?", (safe_int_id(message_id),))
        con.commit()


def create_member_reminder(member_id, reminder_type, title, message, due_date=None,
                           whatsapp_sent=False):
    member_id = safe_int_id(member_id)
    if member_id is None or not str(message or "").strip():
        raise ValueError("Le membre et le rappel sont obligatoires.")
    table = _table("member_reminders")
    with db() as con:
        con.execute(
            f"""INSERT INTO {table}
                (member_id, reminder_type, title, message, due_date, whatsapp_sent)
                VALUES (?, ?, ?, ?, ?, ?)""",
            (member_id, str(reminder_type), str(title), str(message).strip(),
             due_date.isoformat() if due_date else None, bool(whatsapp_sent)),
        )
        con.commit()


def get_member_reminders(member_id):
    table = _table("member_reminders")
    return read_sql(
        f"""SELECT id, member_id, reminder_type, title, message, due_date,
                   whatsapp_sent, created_at
            FROM {table}
            WHERE member_id=?
            ORDER BY created_at DESC, id DESC""",
        [int(member_id)],
    )


def send_member_message_to_admin(member_id, subject, message, message_type="message"):
    """Le membre écrit à l'administration; la conversation est visible côté admin."""
    create_member_message(
        member_id, message, subject, message_type,
        sender_role="member", sender_member_id=member_id
    )


def send_admin_message_to_member(member_id, subject, message, message_type="message"):
    """L'administration répond ou informe un membre."""
    create_member_message(
        member_id, message, subject, message_type,
        sender_role="admin", sender_member_id=None
    )


def ensure_loan_votes(loan_id):
    """Crée un bulletin pour chaque autre membre actif."""
    loan_id = safe_int_id(loan_id)
    if loan_id is None:
        return
    lt, mt, vt = _table("loans"), _table("members"), _table("loan_votes")
    with db() as con:
        loan = con.execute(f"SELECT member_id FROM {lt} WHERE id=?", (loan_id,)).fetchone()
        if not loan:
            return
        borrower_id = safe_int_id(loan["member_id"])
        members = con.execute(
            f"SELECT id FROM {mt} WHERE active=TRUE AND id<>?", (borrower_id,)
        ).fetchall()
        for row in members:
            if use_supabase():
                con.execute(
                    f"""INSERT INTO {vt}(loan_id, voter_member_id, decision)
                        VALUES (?, ?, 'En attente')
                        ON CONFLICT (loan_id, voter_member_id) DO NOTHING""",
                    (loan_id, safe_int_id(row["id"])),
                )
            else:
                con.execute(
                    f"""INSERT OR IGNORE INTO {vt}(loan_id, voter_member_id, decision)
                        VALUES (?, ?, 'En attente')""",
                    (loan_id, safe_int_id(row["id"])),
                )
        con.commit()


def loan_vote_status(loan_id):
    """Un seul veto bloque; le prêt est confirmé lorsque tous les autres membres ont voté oui."""
    ensure_loan_votes(loan_id)
    vt, mt = _table("loan_votes"), _table("members")
    with db() as con:
        rows = con.execute(
            f"""SELECT lv.id, lv.loan_id, lv.voter_member_id, m.full_name AS voter_name,
                       lv.decision, lv.comment, lv.created_at
                FROM {vt} lv
                LEFT JOIN {mt} m ON m.id=lv.voter_member_id
                WHERE lv.loan_id=?
                ORDER BY m.full_name, lv.id""",
            (safe_int_id(loan_id),),
        ).fetchall()
    df = pd.DataFrame([dict(r) for r in rows]) if rows else pd.DataFrame(
        columns=["id","loan_id","voter_member_id","voter_name","decision","comment","created_at"]
    )
    total = len(df)
    vetoes = int((df["decision"] == "Veto").sum()) if not df.empty else 0
    approvals = int((df["decision"] == "Approuvé").sum()) if not df.empty else 0
    pending = int((df["decision"] == "En attente").sum()) if not df.empty else 0
    if vetoes:
        status = "Bloqué par un veto"
    elif total == 0 or (approvals == total and pending == 0):
        status = "Confirmé"
    else:
        status = "En attente de validation"
    return df, status


def cast_loan_vote(loan_id, voter_member_id, decision, comment=""):
    loan_id = safe_int_id(loan_id)
    voter_member_id = safe_int_id(voter_member_id)
    decision = str(decision or "").strip()
    if decision not in {"Approuvé", "Veto"}:
        raise ValueError("Décision invalide.")
    if loan_id is None or voter_member_id is None:
        raise ValueError("Prêt ou membre invalide.")
    lt, mt, vt = _table("loans"), _table("members"), _table("loan_votes")
    with db() as con:
        loan = con.execute(f"SELECT member_id FROM {lt} WHERE id=?", (loan_id,)).fetchone()
        if not loan:
            raise ValueError("Prêt introuvable.")
        if safe_int_id(loan["member_id"]) == voter_member_id:
            raise ValueError("Le bénéficiaire du prêt ne peut pas voter sur son propre prêt.")
        voter = con.execute(
            f"SELECT id FROM {mt} WHERE id=? AND active=TRUE", (voter_member_id,)
        ).fetchone()
        if not voter:
            raise ValueError("Membre votant introuvable ou inactif.")
        if use_supabase():
            con.execute(
                f"""INSERT INTO {vt}(loan_id, voter_member_id, decision, comment)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT (loan_id, voter_member_id)
                    DO UPDATE SET decision=EXCLUDED.decision, comment=EXCLUDED.comment,
                                  created_at=NOW()""",
                (loan_id, voter_member_id, decision, str(comment or "").strip()),
            )
        else:
            con.execute(
                f"""INSERT OR REPLACE INTO {vt}
                    (loan_id, voter_member_id, decision, comment, created_at)
                    VALUES (?, ?, ?, ?, datetime('now'))""",
                (loan_id, voter_member_id, decision, str(comment or "").strip()),
            )
        con.commit()
    _, status = loan_vote_status(loan_id)
    with db() as con:
        con.execute(f"UPDATE {lt} SET status=? WHERE id=?", (status, loan_id))
        con.commit()
    refresh_application_data()
    return status


def loan_votes_for_member(member_id):
    lt, vt, mt = _table("loans"), _table("loan_votes"), _table("members")
    return read_sql(
        f"""SELECT l.id AS loan_id, l.member_id AS borrower_id,
                   borrower.full_name AS borrower_name, l.loan_date, l.principal,
                   l.status, lv.decision, lv.comment, lv.created_at
            FROM {lt} l
            JOIN {vt} lv ON lv.loan_id=l.id
            JOIN {mt} borrower ON borrower.id=l.member_id
            WHERE lv.voter_member_id=? AND l.member_id<>?
            ORDER BY l.loan_date DESC, l.id DESC""",
        [int(member_id), int(member_id)],
    )


# ============================================================
# DATES / RAPPELS
# ============================================================

def _format_reminder_date(value):
    """Formate une date venant de PostgreSQL, SQLite, Pandas ou datetime."""
    if value is None:
        return "Date non définie"
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    try:
        ts = pd.to_datetime(value, errors="coerce")
        if pd.notna(ts):
            return ts.strftime("%d/%m/%Y")
    except Exception:
        pass
    return str(value)

def installment_status(due_date, amount_due, amount_paid):
    """Statut lisible d'une échéance."""
    due = float(amount_due or 0)
    paid = float(amount_paid or 0)
    if paid >= due - 0.01:
        return "Payée"
    if paid > 0:
        return "Partiellement payée"
    try:
        due_d = pd.to_datetime(due_date).date()
        if due_d < date.today():
            return "En retard"
    except Exception:
        pass
    return "À venir"

# ============================================================
# WHATSAPP
# ============================================================

def contribution_message(member_name, d=None):

    d = d or date.today()

    return (
        f"Bonjour {member_name},\n\n"
        f"Petit rappel concernant votre cotisation "
        f"du mois de {month_label(d)}.\n"
        f"Merci d'effectuer votre versement "
        f"dès que possible.\n\n"
        f"Cordialement,\n"
        f"{ADMIN_NAME}\n"
        f"Épargne Étudiant"
    )


def loan_message(
    member_name,
    amount,
    due_date
):

    return (
        f"Bonjour {member_name},\n\n"
        f"Rappel concernant votre échéance de prêt "
        f"de {money(amount)}, prévue le "
        f"{due_date.strftime('%d/%m/%Y')}.\n\n"
        f"Merci d'effectuer votre remboursement "
        f"dans les délais.\n\n"
        f"Cordialement,\n"
        f"{ADMIN_NAME}\n"
        f"Épargne Étudiant"
    )


def whatsapp_link(phone, message):

    phone = normalize_phone(phone)

    return (
        "https://wa.me/"
        + phone
        + "?text="
        + urllib.parse.quote(message)
    )


def send_whatsapp(phone, message):

    if Client is None:
        raise RuntimeError(
            "Twilio n'est pas installé."
        )

    if not TWILIO_ACCOUNT_SID:
        raise RuntimeError(
            "TWILIO_ACCOUNT_SID n'est pas configuré."
        )

    if not TWILIO_AUTH_TOKEN:
        raise RuntimeError(
            "TWILIO_AUTH_TOKEN n'est pas configuré."
        )

    client = Client(
        TWILIO_ACCOUNT_SID,
        TWILIO_AUTH_TOKEN
    )

    return client.messages.create(
        from_=TWILIO_WHATSAPP_FROM,
        to="whatsapp:+" + normalize_phone(phone),
        body=message,
    )


def send_monthly_reminders():

    df = get_members(True)

    results = []

    for _, row in df.iterrows():

        message = contribution_message(
            row["full_name"]
        )

        try:

            sent = send_whatsapp(
                row["phone"],
                message
            )

            results.append({
                "Membre": row["full_name"],
                "Téléphone": row["phone"],
                "Statut": "Envoyé",
                "SID": getattr(sent, "sid", ""),
            })

        except Exception as exc:

            results.append({
                "Membre": row["full_name"],
                "Téléphone": row["phone"],
                "Statut": f"Erreur : {exc}",
                "SID": "",
            })

    return pd.DataFrame(results)


# ============================================================
# ADMINISTRATEURS
# ============================================================

def get_admins():
    table = 'public.admins' if use_supabase() else 'admins'
    return read_sql(f"SELECT id,username,full_name,active,created_at FROM {table} ORDER BY full_name")


def add_admin(username, password, full_name):

    with db() as con:
        con.execute(
            """
            INSERT INTO admins(
                username,
                password,
                full_name
            )
            VALUES (?, ?, ?)
            """,
            (
                username.strip(),
                password,
                full_name.strip(),
            )
        )
        con.commit()


# ============================================================
# TABLEAU DE BORD
# ============================================================

def dashboard():

    mdf = get_members(True)
    cdf = contributions()
    ldf = loans()

    total_saved = (
        float(cdf["amount"].sum())
        if not cdf.empty
        else 0
    )

    total_borrowed = (
        float(ldf["principal"].sum())
        if not ldf.empty
        else 0
    )

    total_due = (
        float(ldf["total_due"].sum())
        if not ldf.empty
        else 0
    )

    brand_hero(
        "Construire son avenir, un versement à la fois",
        "Suivez les cotisations, les prêts et les projets des étudiants en toute simplicité."
    )

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("Membres actifs", len(mdf))
    with col2:
        st.metric("Total épargné", money(total_saved))
    with col3:
        st.metric("Total emprunté", money(total_borrowed))
    with col4:
        st.metric("Total à rembourser", money(total_due))

    st.markdown('<div class="section-title">📊 Résumé des membres</div>', unsafe_allow_html=True)

    rows = []

    for _, member in mdf.iterrows():

        member_id = safe_int_id(member.get("id"))
        if member_id is None:
            continue

        saved = cdf[
            cdf["member_id"] == member_id
        ]["amount"].sum()

        borrowed = ldf[
            ldf["member_id"] == member_id
        ]["principal"].sum()

        rows.append({
            "Membre": member["full_name"],
            "Téléphone": member["phone"],
            "Épargne": money(saved),
            "Emprunts": money(borrowed),
        })

    if rows:
        st.dataframe(
            pd.DataFrame(rows),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# PDF MEMBRE
# ============================================================

def generate_member_pdf(member_id):
    """Génère un bulletin PDF inspiré du visuel Épargne Étudiant."""
    mdf = get_members(False)
    member_rows = mdf[mdf["id"] == member_id]

    if member_rows.empty:
        raise ValueError("Membre introuvable.")

    member = member_rows.iloc[0]
    cdf = contributions(member_id)
    ldf = loans(member_id)
    idf = pd.DataFrame()
    try:
        loan_ids = [safe_int_id(v) for v in ldf.get("id", pd.Series(dtype=object)).tolist()] if not ldf.empty else []
        frames = [get_installments(lid) for lid in loan_ids if lid is not None]
        if frames:
            idf = pd.concat(frames, ignore_index=True)
    except Exception:
        idf = pd.DataFrame()

    # Les colonnes peuvent être absentes ou contenir des valeurs texte/NULL.
    # On normalise toujours les montants avant le calcul pour éviter le ValueError.
    amount_series = pd.to_numeric(cdf.get("amount", pd.Series(dtype=float)), errors="coerce").fillna(0)
    principal_series = pd.to_numeric(ldf.get("principal", pd.Series(dtype=float)), errors="coerce").fillna(0)
    due_series = pd.to_numeric(ldf.get("total_due", pd.Series(dtype=float)), errors="coerce").fillna(0)
    total_saved = float(amount_series.sum())
    total_borrowed = float(principal_series.sum())
    total_due = float(due_series.sum())
    paid_series = pd.to_numeric(idf.get("amount_paid", pd.Series(dtype=float)), errors="coerce").fillna(0)
    installment_due_series = pd.to_numeric(idf.get("amount_due", pd.Series(dtype=float)), errors="coerce").fillna(0)
    total_repaid = float(paid_series.sum())
    total_installment_due = float(installment_due_series.sum())
    total_remaining = max(total_installment_due - total_repaid, 0)

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=13 * mm,
        leftMargin=13 * mm,
        topMargin=13 * mm,
        bottomMargin=14 * mm,
        title=f"Bulletin d'épargne - {member['full_name']}",
        author=ADMIN_NAME,
    )

    styles = getSampleStyleSheet()
    title_style = styles["Title"].clone("BrandTitle")
    title_style.fontName = "Helvetica-Bold"
    title_style.fontSize = 19
    title_style.leading = 22
    title_style.textColor = colors.HexColor(BRAND_NAVY)

    subtitle_style = styles["Normal"].clone("BrandSubtitle")
    subtitle_style.fontSize = 9.5
    subtitle_style.leading = 13
    subtitle_style.textColor = colors.HexColor("#53657D")

    heading_style = styles["Heading2"].clone("BrandHeading")
    heading_style.fontName = "Helvetica-Bold"
    heading_style.fontSize = 12.5
    heading_style.leading = 15
    heading_style.textColor = colors.HexColor(BRAND_NAVY)
    heading_style.spaceBefore = 5
    heading_style.spaceAfter = 5

    small_style = styles["Normal"].clone("BrandSmall")
    small_style.fontSize = 8.5
    small_style.leading = 11
    small_style.textColor = colors.HexColor("#53657D")

    story = []

    # En-tête visuel avec la photo fournie.
    header_cells = []
    if ASSET_IMAGE.exists():
        image = RLImage(str(ASSET_IMAGE), width=42 * mm, height=42 * mm)
        header_cells.append(image)
    else:
        header_cells.append(Spacer(42 * mm, 42 * mm))

    header_text = [
        Paragraph("ÉPARGNE ÉTUDIANT", title_style),
        Spacer(1, 2 * mm),
        Paragraph("Petits efforts, grands projets !", subtitle_style),
        Spacer(1, 6 * mm),
        Paragraph(f"<b>{member['full_name']}</b>", styles["Heading2"]),
        Paragraph(
            f"WhatsApp : +{normalize_phone(member['phone'])}<br/>"
            f"Administrateur : {ADMIN_NAME}",
            small_style,
        ),
    ]

    header_table = Table([[header_cells[0], header_text]], colWidths=[47 * mm, 132 * mm])
    header_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F7FBFE")),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#D7E5F0")),
        ("ROUNDEDCORNERS", [8, 8, 8, 8]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 7 * mm))

    # Cartes de synthèse.
    summary_data = [[
        Paragraph("<b>ÉPARGNE</b><br/><font size=14>%s</font>" % money(total_saved), styles["Normal"]),
        Paragraph("<b>EMPRUNTÉ</b><br/><font size=14>%s</font>" % money(total_borrowed), styles["Normal"]),
        Paragraph("<b>REMBOURSÉ</b><br/><font size=14>%s</font>" % money(total_repaid), styles["Normal"]),
        Paragraph("<b>RESTE À PAYER</b><br/><font size=14>%s</font>" % money(total_remaining), styles["Normal"]),
    ]]
    summary = Table(summary_data, colWidths=[44.6 * mm] * 4)
    summary.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#EAF5FB")),
        ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#FFF0F3")),
        ("BACKGROUND", (2, 0), (2, 0), colors.HexColor("#EAF6F3")),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor(BRAND_NAVY)),
        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#D7E5F0")),
        ("INNERGRID", (0, 0), (-1, -1), 0.6, colors.white),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(summary)
    story.append(Spacer(1, 6 * mm))

    # Cotisations.
    story.append(Paragraph("Historique des cotisations", heading_style))
    contribution_data = [["Date", "Membre", "Mois", "Montant", "Note"]]
    for _, row in cdf.iterrows():
        contribution_data.append([
            str(row.get("payment_date", "")),
            str(member["full_name"]),
            str(row.get("month_label", "")),
            money(pd.to_numeric(row.get("amount", 0), errors="coerce") if pd.notna(row.get("amount", 0)) else 0),
            str(row.get("note", "") or ""),
        ])
    if len(contribution_data) == 1:
        contribution_data.append(["-", str(member["full_name"]), "-", "-", "Aucune cotisation enregistrée"])

    table = Table(
        contribution_data,
        repeatRows=1,
        colWidths=[25 * mm, 40 * mm, 25 * mm, 32 * mm, 57 * mm],
    )
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(BRAND_NAVY)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D8E0E8")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(table)
    story.append(Spacer(1, 6 * mm))

    # Emprunts.
    story.append(Paragraph("Historique des emprunts", heading_style))
    loan_data = [["Date", "Principal", "Taux", "Total dû", "Durée", "Statut"]]
    for _, row in ldf.iterrows():
        loan_data.append([
            str(row["loan_date"]),
            money(row["principal"]),
            f"{row['interest_rate']} %",
            money(row["total_due"]),
            f"{row['duration_months']} mois",
            str(row["status"]),
        ])
    if len(loan_data) == 1:
        loan_data.append(["-", "-", "-", "-", "-", "Aucun emprunt enregistré"])

    loan_table = Table(
        loan_data,
        repeatRows=1,
        colWidths=[27 * mm, 31 * mm, 20 * mm, 31 * mm, 27 * mm, 43 * mm],
    )
    loan_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(BRAND_GREEN)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.8),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D8E0E8")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(loan_table)
    story.append(Spacer(1, 6 * mm))

    # Historique détaillé des échéances et remboursements du membre.
    story.append(Paragraph("Historique des remboursements et solde", heading_style))
    payment_data = [["Échéance", "Date prévue", "Prévu", "Remboursé", "Reste", "Date paiement", "Statut"]]
    for _, r in idf.sort_values(["due_date", "id"] if not idf.empty and "id" in idf.columns else ["due_date"]).iterrows():
        due = float(r.get("amount_due", 0) or 0)
        paid = float(r.get("amount_paid", 0) or 0)
        payment_data.append([
            str(r.get("installment_number", "")),
            _format_reminder_date(r.get("due_date")),
            money(due),
            money(paid),
            money(max(due-paid, 0)),
            _format_reminder_date(r.get("payment_date")) if r.get("payment_date") else "—",
            installment_status(r.get("due_date"), due, paid),
        ])
    if len(payment_data) == 1:
        payment_data.append(["—", "—", money(0), money(0), money(0), "—", "Aucun remboursement"])
    payment_table = Table(payment_data, repeatRows=1, colWidths=[15*mm, 24*mm, 28*mm, 28*mm, 25*mm, 27*mm, 32*mm])
    payment_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor(BRAND_NAVY)),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 7.3),
        ("GRID", (0,0), (-1,-1), 0.35, colors.HexColor("#D8E0E8")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7FAFC")]),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))
    story.append(payment_table)
    story.append(Spacer(1, 7 * mm))

    story.append(
        Paragraph(
            "« Mon avenir se construit aujourd’hui. »",
            ParagraphStyle(
                "Quote",
                parent=styles["Normal"],
                fontName="Helvetica-Oblique",
                fontSize=10,
                textColor=colors.HexColor(BRAND_NAVY),
                alignment=1,
            ),
        )
    )
    story.append(Spacer(1, 2 * mm))
    story.append(
        Paragraph(
            f"Document généré le {date.today().strftime('%d/%m/%Y')} · {ADMIN_NAME} · Épargne Étudiant",
            small_style,
        )
    )

    def draw_footer(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor(BRAND_BLUE))
        canvas.setLineWidth(1)
        canvas.line(13 * mm, 9 * mm, A4[0] - 13 * mm, 9 * mm)
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(colors.HexColor("#687A90"))
        canvas.drawString(13 * mm, 5.5 * mm, "Épargne Étudiant — Petits efforts, grands projets")
        canvas.drawRightString(
            A4[0] - 13 * mm,
            5.5 * mm,
            f"Page {doc.page}",
        )
        canvas.restoreState()

    doc.build(story, onFirstPage=draw_footer, onLaterPages=draw_footer)
    buffer.seek(0)
    return buffer.getvalue()




# ============================================================
# RAPPORT GLOBAL PDF + EXCEL
# ============================================================

def all_installments():
    """Récupère toutes les échéances en une seule requête (évite le N+1)."""
    if use_supabase():
        tables = ('public.loan_installments', 'public.loans', 'public.members')
    else:
        tables = ('loan_installments', 'loans', 'members')
    return read_sql(f"""
        SELECT i.id, i.loan_id, i.installment_number, i.due_date,
               i.amount_due, i.amount_paid, i.payment_date, i.note,
               l.member_id, m.full_name
        FROM {tables[0]} i
        JOIN {tables[1]} l ON l.id=i.loan_id
        JOIN {tables[2]} m ON m.id=l.member_id
        ORDER BY i.due_date, i.id
    """)


def global_report_data():
    """Construit une vue globale et détaillée de toute l'épargne enregistrée."""
    mdf = get_members(False).copy()
    cdf = contributions().copy()
    ldf = loans().copy()

    if cdf.empty:
        cdf = pd.DataFrame(columns=["id", "member_id", "full_name", "payment_date", "month_label", "amount", "note"])
    if ldf.empty:
        ldf = pd.DataFrame(columns=["id", "member_id", "full_name", "loan_date", "principal", "interest_rate", "total_due", "duration_months", "first_due_date", "status", "note"])

    # Échéances et remboursements enregistrés : une seule requête.
    try:
        idf = all_installments().copy()
    except Exception:
        idf = pd.DataFrame(columns=["id", "loan_id", "installment_number", "due_date", "amount_due", "amount_paid", "payment_date", "note", "member_id", "full_name"])

    total_contributed = float(pd.to_numeric(cdf.get("amount", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
    total_borrowed = float(pd.to_numeric(ldf.get("principal", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
    total_received = float(pd.to_numeric(idf.get("amount_paid", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
    total_due = float(pd.to_numeric(idf.get("amount_due", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
    outstanding = max(total_due - total_received, 0.0)
    available = total_contributed + total_received - total_borrowed

    if not cdf.empty:
        cdf["amount"] = pd.to_numeric(cdf["amount"], errors="coerce").fillna(0)
        cdf["payment_date"] = cdf["payment_date"].astype(str)
    if not ldf.empty:
        for col in ["principal", "interest_rate", "total_due"]:
            ldf[col] = pd.to_numeric(ldf[col], errors="coerce").fillna(0)
        ldf["loan_date"] = ldf["loan_date"].astype(str)
        ldf["first_due_date"] = ldf["first_due_date"].astype(str)
    if not idf.empty:
        for col in ["amount_due", "amount_paid"]:
            idf[col] = pd.to_numeric(idf[col], errors="coerce").fillna(0)
        idf["remaining"] = (idf["amount_due"] - idf["amount_paid"]).clip(lower=0)
        idf["due_date"] = idf["due_date"].astype(str)
        idf["payment_date"] = idf["payment_date"].fillna("").astype(str)

    rows = []
    for _, member in mdf.iterrows():
        mid = safe_int_id(member.get("id"))
        if mid is None:
            continue
        mc = cdf[cdf["member_id"] == mid] if not cdf.empty else cdf
        ml = ldf[ldf["member_id"] == mid] if not ldf.empty else ldf
        mi = idf[idf["member_id"] == mid] if not idf.empty else idf
        contributed = float(mc["amount"].sum()) if not mc.empty else 0
        borrowed = float(ml["principal"].sum()) if not ml.empty else 0
        received = float(mi["amount_paid"].sum()) if not mi.empty else 0
        due = float(mi["amount_due"].sum()) if not mi.empty else 0
        months = sorted(set(str(x) for x in mc["month_label"].dropna())) if not mc.empty else []
        rows.append({
            "Membre": member["full_name"],
            "Téléphone": member["phone"],
            "Mois de cotisation": ", ".join(months),
            "Nombre de mois": len(months),
            "Total cotisé": contributed,
            "Total emprunté": borrowed,
            "Somme reçue": received,
            "Échéances prévues": due,
            "Reste à recevoir": max(due - received, 0),
            "Disponible net": contributed + received - borrowed,
        })

    summary_df = pd.DataFrame(rows)
    return {
        "members": mdf,
        "contributions": cdf,
        "loans": ldf,
        "installments": idf,
        "summary": summary_df,
        "total_contributed": total_contributed,
        "total_borrowed": total_borrowed,
        "total_received": total_received,
        "total_due": total_due,
        "outstanding": outstanding,
        "available": available,
        "contribution_months": sorted(set(str(x) for x in cdf["month_label"].dropna())) if not cdf.empty else [],
    }


def generate_global_excel():
    """Génère un classeur Excel complet : synthèse, membres, cotisations, prêts et échéances."""
    if openpyxl is None:
        raise RuntimeError(
            "Le module openpyxl est requis pour l'export Excel. Ajoutez openpyxl dans requirements.txt puis redéployez l'application."
        )
    data = global_report_data()
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        summary = pd.DataFrame([
            ["Total cotisé", data["total_contributed"]],
            ["Total emprunté", data["total_borrowed"]],
            ["Somme reçue sur remboursements", data["total_received"]],
            ["Total des échéances prévues", data["total_due"]],
            ["Reste à recevoir", data["outstanding"]],
            ["Somme disponible nette", data["available"]],
            ["Nombre de membres", len(data["members"])],
            ["Nombre de mois de cotisation", len(data["contribution_months"])],
            ["Mois de cotisation", ", ".join(data["contribution_months"])],
            ["Généré le", date.today().isoformat()],
        ], columns=["Indicateur", "Valeur"])
        summary.to_excel(writer, sheet_name="Synthèse", index=False)
        data["summary"].to_excel(writer, sheet_name="Par membre", index=False)
        data["contributions"].to_excel(writer, sheet_name="Cotisations", index=False)
        data["loans"].to_excel(writer, sheet_name="Emprunts", index=False)
        data["installments"].to_excel(writer, sheet_name="Échéances", index=False)

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                max_len = max((len(str(cell.value)) if cell.value is not None else 0) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 42)
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)

    buffer.seek(0)
    return buffer.getvalue()


def _pdf_image(path, max_width, max_height):
    """Retourne une image PDF sans déformation, avec conservation de son ratio."""
    try:
        reader = ImageReader(str(path))
        width, height = reader.getSize()
        if not width or not height:
            return None
        ratio = min(max_width / width, max_height / height)
        img = RLImage(str(path), width=width * ratio, height=height * ratio)
        return img
    except Exception:
        return None


def generate_global_pdf():
    """Génère le rapport global PDF de toutes les opérations enregistrées."""
    data = global_report_data()
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=10 * mm,
        leftMargin=10 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title="Rapport global - Épargne Étudiant",
        author=ADMIN_NAME,
    )

    styles = getSampleStyleSheet()
    title = styles["Title"].clone("GlobalTitle")
    title.fontName = "Helvetica-Bold"
    title.fontSize = 21
    title.leading = 24
    title.textColor = colors.HexColor(BRAND_NAVY)

    h2 = styles["Heading2"].clone("GlobalH2")
    h2.fontName = "Helvetica-Bold"
    h2.fontSize = 12.5
    h2.textColor = colors.HexColor(BRAND_NAVY)
    h2.spaceBefore = 5
    h2.spaceAfter = 5

    small = styles["Normal"].clone("GlobalSmall")
    small.fontSize = 8
    small.leading = 10
    small.textColor = colors.HexColor("#53657D")

    story = []
    header_cells = []
    img = _pdf_image(ASSET_IMAGE, 32 * mm, 25 * mm) if ASSET_IMAGE.exists() else None
    header_cells.append(img if img else Spacer(32 * mm, 20 * mm))
    header_cells.append([
        Paragraph("ÉPARGNE ÉTUDIANT", title),
        Spacer(1, 1 * mm),
        Paragraph("Rapport global de l'épargne, des cotisations, des prêts et des échéances", small),
        Paragraph(f"Édité le {date.today().strftime('%d/%m/%Y')} — Administrateur : {ADMIN_NAME}", small),
    ])
    ht = Table([header_cells], colWidths=[40 * mm, 235 * mm])
    ht.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#F7FBFE")),
        ("BOX", (0,0), (-1,-1), 0.7, colors.HexColor("#D7E5F0")),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 7),
        ("RIGHTPADDING", (0,0), (-1,-1), 7),
        ("TOPPADDING", (0,0), (-1,-1), 7),
        ("BOTTOMPADDING", (0,0), (-1,-1), 7),
    ]))
    story += [ht, Spacer(1, 5 * mm)]

    metrics = [
        [Paragraph(f"<b>TOTAL COTISÉ</b><br/><font size=15>{money(data['total_contributed'])}</font>", small),
         Paragraph(f"<b>TOTAL EMPRUNTÉ</b><br/><font size=15>{money(data['total_borrowed'])}</font>", small),
         Paragraph(f"<b>SOMME REÇUE</b><br/><font size=15>{money(data['total_received'])}</font>", small),
         Paragraph(f"<b>ÉCHÉANCES RESTANTES</b><br/><font size=15>{money(data['outstanding'])}</font>", small),
         Paragraph(f"<b>DISPONIBLE NET</b><br/><font size=15>{money(data['available'])}</font>", small)]
    ]
    mt = Table(metrics, colWidths=[54 * mm] * 5)
    mt.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (0,0), colors.HexColor("#EAF5FB")),
        ("BACKGROUND", (1,0), (1,0), colors.HexColor("#FFF0F3")),
        ("BACKGROUND", (2,0), (2,0), colors.HexColor("#EAF6F3")),
        ("BACKGROUND", (3,0), (3,0), colors.HexColor("#FFF7E8")),
        ("BACKGROUND", (4,0), (4,0), colors.HexColor("#EDF0FA")),
        ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor("#D7E5F0")),
        ("INNERGRID", (0,0), (-1,-1), 0.5, colors.white),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 7),
        ("RIGHTPADDING", (0,0), (-1,-1), 7),
        ("TOPPADDING", (0,0), (-1,-1), 7),
        ("BOTTOMPADDING", (0,0), (-1,-1), 7),
    ]))
    story += [mt, Spacer(1, 4 * mm)]

    story.append(Paragraph("1. Synthèse par membre", h2))
    summary_data = [["Membre", "Mois", "Cotisé", "Emprunté", "Reçu", "Échéances", "Reste", "Disponible net"]]
    for _, r in data["summary"].iterrows():
        summary_data.append([
            str(r["Membre"]), str(r["Nombre de mois"]), money(r["Total cotisé"]), money(r["Total emprunté"]),
            money(r["Somme reçue"]), money(r["Échéances prévues"]), money(r["Reste à recevoir"]), money(r["Disponible net"])
        ])
    if len(summary_data) == 1:
        summary_data.append(["Aucun membre", "0", money(0), money(0), money(0), money(0), money(0), money(0)])
    stbl = Table(summary_data, repeatRows=1, colWidths=[57*mm, 15*mm, 27*mm, 27*mm, 27*mm, 27*mm, 27*mm, 31*mm])
    stbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor(BRAND_NAVY)),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 7.5),
        ("GRID", (0,0), (-1,-1), 0.3, colors.HexColor("#D8E0E8")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7FAFC")]),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))
    story += [stbl, Spacer(1, 4 * mm)]

    story.append(Paragraph("2. Toutes les cotisations enregistrées", h2))
    ctable = [["Membre", "Date", "Mois", "Montant", "Note"]]
    for _, r in data["contributions"].iterrows():
        ctable.append([str(r.get("full_name", "")), str(r.get("payment_date", "")), str(r.get("month_label", "")), money(r.get("amount", 0)), str(r.get("note", "") or "")])
    if len(ctable) == 1:
        ctable.append(["Aucune", "-", "-", money(0), ""])
    ct = Table(ctable, repeatRows=1, colWidths=[55*mm, 28*mm, 28*mm, 32*mm, 98*mm])
    ct.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor(BRAND_NAVY)),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 7.5),
        ("GRID", (0,0), (-1,-1), 0.3, colors.HexColor("#D8E0E8")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7FAFC")]),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
    ]))
    story += [ct, Spacer(1, 4 * mm)]

    story.append(Paragraph("3. Tous les emprunts enregistrés", h2))
    ltable = [["Membre", "Date", "Principal", "Taux", "Total dû", "Durée", "1ère échéance", "Statut"]]
    for _, r in data["loans"].iterrows():
        ltable.append([
            str(r.get("full_name", "")), str(r.get("loan_date", "")), money(r.get("principal", 0)),
            f"{float(r.get('interest_rate', 0)):.2f} %", money(r.get("total_due", 0)),
            f"{r.get('duration_months', 0)} mois", str(r.get("first_due_date", "")), str(r.get("status", ""))
        ])
    if len(ltable) == 1:
        ltable.append(["Aucun", "-", money(0), "0 %", money(0), "-", "-", "-"])
    lt = Table(ltable, repeatRows=1, colWidths=[52*mm, 25*mm, 31*mm, 20*mm, 31*mm, 25*mm, 32*mm, 45*mm])
    lt.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor(BRAND_GREEN)),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 7.3),
        ("GRID", (0,0), (-1,-1), 0.3, colors.HexColor("#D8E0E8")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7FAFC")]),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
    ]))
    story += [lt, Spacer(1, 4 * mm)]

    story.append(Paragraph("4. Échéances et remboursements", h2))
    itable = [["Membre", "N° prêt", "Échéance", "Montant prévu", "Montant reçu", "Reste", "Date réception", "Note"]]
    for _, r in data["installments"].iterrows():
        itable.append([
            str(r.get("full_name", "")), str(r.get("loan_id", "")), str(r.get("due_date", "")),
            money(r.get("amount_due", 0)), money(r.get("amount_paid", 0)), money(r.get("remaining", 0)),
            str(r.get("payment_date", "")), str(r.get("note", "") or "")
        ])
    if len(itable) == 1:
        itable.append(["Aucune", "-", "-", money(0), money(0), money(0), "-", ""])
    it = Table(itable, repeatRows=1, colWidths=[48*mm, 22*mm, 28*mm, 32*mm, 32*mm, 27*mm, 30*mm, 60*mm])
    it.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor(BRAND_NAVY)),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 7.2),
        ("GRID", (0,0), (-1,-1), 0.3, colors.HexColor("#D8E0E8")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7FAFC")]),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
    ]))
    story.append(it)

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.HexColor("#687A90"))
        canvas.drawString(10 * mm, 6 * mm, "Épargne Étudiant — Rapport global")
        canvas.drawRightString(landscape(A4)[0] - 10 * mm, 6 * mm, f"Page {doc.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# INTERFACE
# ============================================================

_db_fingerprint = "|".join([SUPABASE_HOST, SUPABASE_PORT, SUPABASE_DATABASE, SUPABASE_USER, SUPABASE_PASSWORD[:8] if SUPABASE_PASSWORD else ""])
try:
    init_db(_db_fingerprint)
except Exception as _db_exc:
    st.error("❌ Connexion Supabase impossible")
    st.code(str(_db_exc))
    st.info("Vérifiez les secrets [postgres] et le paquet psycopg2-binary dans requirements.txt, puis redémarrez l'application.")
    st.stop()

# Les prêts déjà présents reçoivent également leur bulletin de vote.
try:
    existing_loans = loans()
    for _lid in existing_loans["id"].tolist() if not existing_loans.empty else []:
        ensure_loan_votes(_lid)
except Exception:
    pass

if "user" not in st.session_state:
    st.session_state.user = None


if st.session_state.user is None:

    st.markdown(
        """<style>
        [data-testid="stSidebar"] {display:none;}
        [data-testid="stMainBlockContainer"] {max-width:1180px;}
        </style>""",
        unsafe_allow_html=True,
    )

    st.markdown('<div class="login-shell">', unsafe_allow_html=True)
    left, right = st.columns([1.08, 0.92], gap="large")

    with left:
        st.markdown('<div class="login-card">', unsafe_allow_html=True)
        if ASSET_IMAGE.exists():
            st.markdown(asset_image_html("login-photo"), unsafe_allow_html=True)
        else:
            st.markdown(
                """<div class="login-photo"><div style="text-align:center;padding:30px;">
                <div style="font-size:5rem;">💰</div>
                <div style="font-size:1.4rem;font-weight:900;color:#122A55;">Épargne Étudiant</div>
                </div></div>""",
                unsafe_allow_html=True,
            )
        st.markdown('</div>', unsafe_allow_html=True)

    with right:
        st.markdown('<div class="login-card login-panel">', unsafe_allow_html=True)
        st.markdown('<div class="login-badge">Gestion financière étudiante</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-title">Petits efforts,<br>grands projets !</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="login-subtitle">Gérez simplement les membres, les cotisations, les prêts et les rappels de votre groupe étudiant.</div>',
            unsafe_allow_html=True,
        )
        login_tab_admin, login_tab_member = st.tabs(["👨‍💼 Administrateur", "👤 Membre"])

        with login_tab_admin:
            with st.form("login_form_admin", clear_on_submit=False):
                username = st.text_input("Nom d'utilisateur", placeholder="Identifiant administrateur", key="admin_login_username")
                password = st.text_input("Mot de passe", type="password", placeholder="Mot de passe", key="admin_login_password")
                submitted = st.form_submit_button("Se connecter", type="primary", use_container_width=True)
                if submitted:
                    try:
                        user = authenticate(username, password)
                    except Exception as exc:
                        st.error("Connexion impossible. Vérifiez la configuration Supabase et les tables.")
                        st.code(str(exc))
                        user = None
                    if user:
                        user["role"] = "admin"
                        st.session_state.user = user
                        st.rerun()
                    else:
                        st.error("Identifiants administrateur incorrects.")

        with login_tab_member:
            with st.form("login_form_member", clear_on_submit=False):
                member_username = st.text_input("Identifiant membre", placeholder="Identifiant qui vous a été remis", key="member_login_username")
                member_password = st.text_input("Mot de passe membre", type="password", placeholder="Votre mot de passe", key="member_login_password")
                member_submitted = st.form_submit_button("Accéder à mon compte", type="primary", use_container_width=True)
                if member_submitted:
                    try:
                        user = authenticate_member(member_username, member_password)
                    except Exception as exc:
                        st.error("Connexion membre impossible. Vérifiez la configuration de la base.")
                        st.code(str(exc))
                        user = None
                    if user:
                        st.session_state.user = user
                        st.rerun()
                    else:
                        st.error("Identifiant ou mot de passe membre incorrect, ou compte désactivé.")

        st.markdown(
            '<div class="info-card">🔒 Vos données d’épargne, de cotisations et de prêts sont enregistrées dans la base configurée par l’administrateur.</div>',
            unsafe_allow_html=True,
        )
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# ============================================================
# BARRE DE NAVIGATION SUPÉRIEURE
# ============================================================

user_role = st.session_state.user.get("role", "admin")

# Navigation principale en haut de page : aucun menu latéral.
st.markdown(
    """
    <style>
    /* Supprime complètement l'ancien panneau latéral. */
    [data-testid="stSidebar"] {
        display: none !important;
    }

    /* La zone principale reprend toute la largeur. */
    [data-testid="stAppViewContainer"] > .main {
        margin-left: 0 !important;
    }

    .top-navigation {
        position: sticky;
        top: 0;
        z-index: 999;
        margin: -1rem -1rem 1.25rem -1rem;
        padding: 12px 20px 10px 20px;
        background: rgba(255,255,255,.96);
        backdrop-filter: blur(16px);
        -webkit-backdrop-filter: blur(16px);
        border-bottom: 1px solid rgba(18,42,85,.10);
        box-shadow: 0 8px 24px rgba(18,42,85,.08);
    }

    .top-brand {
        display: flex;
        align-items: center;
        gap: 10px;
        min-height: 42px;
    }

    .top-brand-icon {
        font-size: 1.9rem;
        line-height: 1;
    }

    .top-brand-name {
        color: #122A55;
        font-size: 1.12rem;
        font-weight: 900;
        line-height: 1.1;
    }

    .top-brand-subtitle {
        color: #64748B;
        font-size: .72rem;
        margin-top: 2px;
    }

    /* Style du radio horizontal utilisé comme menu. */
    div[data-testid="stRadio"] > label {
        display: none;
    }

    div[data-testid="stRadio"] > div {
        gap: 5px !important;
        flex-wrap: wrap !important;
    }

    div[data-testid="stRadio"] [role="radiogroup"] {
        gap: 5px !important;
        flex-wrap: wrap !important;
    }

    div[data-testid="stRadio"] label {
        border-radius: 999px !important;
        padding: 7px 13px !important;
        border: 1px solid transparent !important;
        transition: all .18s ease;
    }

    div[data-testid="stRadio"] label:hover {
        background: #EEF7FD !important;
        border-color: rgba(18,42,85,.10) !important;
    }

    .top-user {
        text-align: right;
        color: #475569;
        font-size: .78rem;
        padding-top: 4px;
    }

    @media (max-width: 900px) {
        .top-navigation {
            margin-left: -0.5rem;
            margin-right: -0.5rem;
            padding-left: 10px;
            padding-right: 10px;
        }
        .top-user {
            text-align: left;
            margin-top: 4px;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="top-navigation">', unsafe_allow_html=True)

brand_col, user_col = st.columns([2.4, 1.2], vertical_alignment="center")

with brand_col:
    st.markdown(
        """
        <div class="top-brand">
            <div class="top-brand-icon">🐷</div>
            <div>
                <div class="top-brand-name">Épargne Étudiant</div>
                <div class="top-brand-subtitle">Petits efforts, grands projets</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with user_col:
    st.markdown(
        f'<div class="top-user">👤 <strong>{st.session_state.user["full_name"]}</strong></div>',
        unsafe_allow_html=True,
    )

if user_role == "member":
    page = "Mon compte"
    st.info("👤 Espace membre — lecture seule")
else:
    pages = [
        "Tableau de bord",
        "Membres",
        "Cotisations",
        "Emprunts",
        "Rappels WhatsApp",
        "Communication",
        "Rapport global",
        "Bulletins PDF",
        "Administrateurs",
    ]

    page = st.radio(
        "Menu principal",
        pages,
        horizontal=True,
        label_visibility="collapsed",
        key="top_navigation_page",
    )

# Actions globales dans la barre supérieure.
action_col1, action_col2, action_col3 = st.columns([1, 1, 5])
with action_col1:
    if st.button("🔄 Actualiser", use_container_width=True, key="top_refresh"):
        refresh_application_data()
        st.rerun()
with action_col2:
    if st.button("↪ Déconnexion", use_container_width=True, key="top_logout"):
        st.session_state.user = None
        st.rerun()

if use_supabase():
    st.caption("🟢 Supabase PostgreSQL connecté")
else:
    st.caption("🔴 Supabase PostgreSQL non configuré")

st.markdown('</div>', unsafe_allow_html=True)


# ============================================================
# TABLEAU DE BORD
# ============================================================

if page == "Mon compte" and user_role == "member":
    member_account_page(int(st.session_state.user["member_id"]))

elif page == "Tableau de bord":

    dashboard()


# ============================================================
# MEMBRES
# ============================================================

elif page == "Membres":

    brand_hero("Les membres", "Chaque étudiant avance à son rythme, avec un objectif clair.")

    with st.expander(
        "➕ Ajouter un membre",
        expanded=True
    ):

        with st.form("add_member_form"):

            name = st.text_input(
                "Nom complet"
            )

            phone = st.text_input(
                "Numéro WhatsApp",
                placeholder="77 752 19 69"
            )

            target = st.number_input(
                "Objectif mensuel",
                min_value=0.0,
                step=1000.0
            )

            notes = st.text_area(
                "Note"
            )

            submit = st.form_submit_button(
                "Ajouter"
            )

            if submit:

                try:

                    new_member_id = add_member(
                        name,
                        phone,
                        target,
                        notes
                    )
                    refresh_application_data()
                    st.session_state["member_flash"] = (
                        f"Membre ajouté avec succès — ID {new_member_id}."
                    )
                    st.rerun()

                except Exception as exc:

                    st.error(str(exc))

    member_flash = st.session_state.pop("member_flash", None)
    if member_flash:
        st.success(member_flash)

    st.divider()

    df = get_members(False)

    if df.empty:
        st.warning("Aucun membre trouvé dans la base utilisée par l'application.")
        if use_supabase():
            try:
                with db() as con:
                    count_row = con.execute(
                        "SELECT COUNT(*) AS total FROM public.members"
                    ).fetchone()
                total_db = int(count_row["total"]) if count_row else 0
                st.caption(
                    f"Diagnostic Supabase : {total_db} ligne(s) présente(s) "
                    "dans public.members."
                )
            except Exception as exc:
                st.error(f"Lecture de public.members impossible : {exc}")

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )

    if not df.empty:

        st.subheader("Modifier un membre")

        options = build_member_options(df)

        option_labels = list(options.keys())
        if not option_labels:
            st.warning("Aucun membre sélectionnable.")
            st.stop()

        selected = st.selectbox(
            "Membre",
            option_labels,
            index=0
        )

        member_id = options.get(selected)
        if member_id is None:
            st.warning("Le membre sélectionné n'est plus disponible. Actualisez la page.")
            st.stop()

        row = df[
            df["id"] == member_id
        ].iloc[0]

        with st.form("edit_member_form"):

            name = st.text_input(
                "Nom",
                value=str(row["full_name"])
            )

            phone = st.text_input(
                "WhatsApp",
                value=str(row["phone"])
            )

            target = st.number_input(
                "Objectif mensuel",
                min_value=0.0,
                value=float(row["monthly_target"] or 0),
                step=1000.0
            )

            notes = st.text_area(
                "Note",
                value=str(row["notes"] or "")
            )

            active = st.checkbox(
                "Membre actif",
                value=bool(row["active"])
            )

            save = st.form_submit_button(
                "Enregistrer"
            )

            if save:

                update_member(
                    member_id,
                    name,
                    phone,
                    target,
                    notes,
                    active
                )

                st.success(
                    "Membre modifié."
                )

                st.rerun()


# ============================================================
# COTISATIONS
# ============================================================

elif page == "Cotisations":
    brand_hero("Les cotisations", "Enregistrez, consultez et modifiez chaque versement.", compact=True)
    mdf = get_members(True)
    if mdf.empty:
        st.warning("Aucun membre actif trouvé.")
    else:
        member_options = build_member_options(mdf)
        with st.form("contribution_form"):
            selected = st.selectbox("Membre", list(member_options.keys()), key="new_contribution_member")
            payment_date = st.date_input("Date réelle du paiement", value=date.today())
            amount = st.number_input("Montant réellement versé", min_value=0.01, step=500.0)
            note = st.text_input("Note")
            submit = st.form_submit_button("Enregistrer la cotisation")
            if submit:
                try:
                    add_contribution(member_options[selected], payment_date, amount, note)
                    st.success("Cotisation enregistrée.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))
        st.divider()
        cdf = contributions()
        st.subheader("📋 Historique des cotisations")
        display = cdf.copy()
        if not display.empty:
            display['Montant'] = pd.to_numeric(display['amount'], errors='coerce').fillna(0).map(money)
            display = display.rename(columns={'full_name':'Membre','payment_date':'Date','month_label':'Mois','note':'Note'})[['id','Membre','Date','Mois','Montant','Note']]
        st.dataframe(display, use_container_width=True, hide_index=True)
        if not cdf.empty:
            copts = {f"#{safe_int_id(r['id'])} — {r['full_name']} — {money(r['amount'])} — {r['payment_date']}": safe_int_id(r['id']) for _,r in cdf.iterrows() if safe_int_id(r.get('id')) is not None}
            selected_c = st.selectbox("Cotisation à modifier", list(copts.keys()), key="edit_contribution_select")
            cid = copts[selected_c]
            crow = cdf[cdf['id'].map(safe_int_id)==cid].iloc[0]
            with st.form("edit_contribution_form"):
                edit_date = st.date_input("Date", value=pd.to_datetime(crow['payment_date']).date())
                edit_amount = st.number_input("Montant", min_value=0.01, value=float(crow['amount']), step=500.0)
                edit_note = st.text_input("Note", value=str(crow['note'] or ''))
                save_c = st.form_submit_button("💾 Modifier la cotisation")
                if save_c:
                    try:
                        update_contribution(cid, edit_date, edit_amount, edit_note)
                        st.success("Cotisation modifiée.")
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))


# ============================================================
# EMPRUNTS
# ============================================================

elif page == "Emprunts":
    brand_hero("Les emprunts", "Prêts, échéances, remboursements et historique de paiement au même endroit.", compact=True)
    mdf = get_members(True)
    if mdf.empty:
        st.warning("Aucun membre actif trouvé.")
    else:
        options = build_member_options(mdf)
        with st.form("loan_form"):
            selected = st.selectbox("Membre", list(options.keys()), key="new_loan_member")
            member_id = options.get(selected)
            loan_date = st.date_input("Date du prêt", value=date.today())
            principal = st.number_input("Montant du prêt", min_value=1.0, step=1000.0)
            rate = st.number_input("Taux d'intérêt total (%)", min_value=0.0, step=0.5)
            duration = st.number_input("Nombre d'échéances", min_value=1, max_value=60, value=1)
            first_due_date = st.date_input("Première échéance", value=date.today())
            note = st.text_input("Note")
            submit = st.form_submit_button("Enregistrer le prêt")
            if submit:
                try:
                    create_loan(member_id, loan_date, principal, rate, duration, first_due_date, note)
                    st.success("Prêt enregistré. Il est maintenant soumis aux votes des autres membres actifs.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))

        st.divider()
        ldf = loans()
        st.subheader("📋 Historique des emprunts")
        st.dataframe(ldf, use_container_width=True, hide_index=True)

        if not ldf.empty:
            st.subheader("🗳️ État des validations / vetos")
            vote_rows = []
            for _, loan in ldf.iterrows():
                lid = safe_int_id(loan.get("id"))
                if lid is None:
                    continue
                vdf, vstatus = loan_vote_status(lid)
                vote_rows.append({
                    "Prêt": lid,
                    "Bénéficiaire": loan.get("full_name", ""),
                    "Montant": money(loan.get("principal", 0)),
                    "Statut": vstatus,
                    "Approuvés": int((vdf["decision"] == "Approuvé").sum()) if not vdf.empty else 0,
                    "Vetos": int((vdf["decision"] == "Veto").sum()) if not vdf.empty else 0,
                    "En attente": int((vdf["decision"] == "En attente").sum()) if not vdf.empty else 0,
                })
            st.dataframe(pd.DataFrame(vote_rows), use_container_width=True, hide_index=True)

        if not ldf.empty:
            loan_options = {f"#{safe_int_id(r['id'])} — {r['full_name']} — {money(r['principal'])}": safe_int_id(r['id']) for _,r in ldf.iterrows() if safe_int_id(r.get('id')) is not None}
            selected_loan_label = st.selectbox("Prêt à gérer", list(loan_options.keys()), key="manage_loan_select")
            loan_id = loan_options[selected_loan_label]
            loan_row = ldf[ldf['id'].map(safe_int_id)==loan_id].iloc[0]

            with st.expander("✏️ Modifier le prêt", expanded=False):
                member_opts = build_member_options(mdf)
                current_member_label = next((k for k,v in member_opts.items() if v == safe_int_id(loan_row['member_id'])), list(member_opts.keys())[0])
                with st.form("edit_loan_form"):
                    edit_member_label = st.selectbox("Membre", list(member_opts.keys()), index=list(member_opts.keys()).index(current_member_label))
                    edit_loan_date = st.date_input("Date du prêt", value=pd.to_datetime(loan_row['loan_date']).date())
                    edit_principal = st.number_input("Montant du prêt", min_value=1.0, value=float(loan_row['principal']), step=1000.0)
                    edit_rate = st.number_input("Taux d'intérêt total (%)", min_value=0.0, value=float(loan_row['interest_rate']), step=0.5)
                    edit_duration = st.number_input("Nombre d'échéances", min_value=1, max_value=60, value=int(loan_row['duration_months']))
                    edit_first_due = st.date_input("Première échéance", value=pd.to_datetime(loan_row['first_due_date']).date())
                    edit_note = st.text_input("Note", value=str(loan_row['note'] or ''))
                    save_loan = st.form_submit_button("💾 Enregistrer les modifications")
                    if save_loan:
                        try:
                            update_loan(loan_id, member_opts[edit_member_label], edit_loan_date, edit_principal, edit_rate, edit_duration, edit_first_due, edit_note)
                            st.success("Prêt modifié.")
                            st.rerun()
                        except Exception as exc:
                            st.error(str(exc))

            idf = get_installments(loan_id).copy()
            if not idf.empty:
                idf['Statut'] = [installment_status(r['due_date'], r['amount_due'], r['amount_paid']) for _,r in idf.iterrows()]
                idf['Reste'] = (pd.to_numeric(idf['amount_due'], errors='coerce').fillna(0)-pd.to_numeric(idf['amount_paid'], errors='coerce').fillna(0)).clip(lower=0)
                st.subheader("📅 Historique des paiements selon les échéances")
                hist = idf.rename(columns={'installment_number':'Échéance #','due_date':'Date prévue','amount_due':'Montant prévu','amount_paid':'Montant remboursé','payment_date':'Date remboursement','note':'Note'})
                hist['Montant prévu'] = hist['Montant prévu'].map(money); hist['Montant remboursé'] = hist['Montant remboursé'].map(money); hist['Reste'] = hist['Reste'].map(money)
                st.dataframe(hist[['Échéance #','Date prévue','Montant prévu','Montant remboursé','Reste','Date remboursement','Statut','Note']], use_container_width=True, hide_index=True)

                inst_options = {f"Échéance #{int(r['installment_number'])} — {r['due_date']} — {money(r['amount_due'])}": safe_int_id(r['id']) for _,r in idf.iterrows() if safe_int_id(r.get('id')) is not None}
                selected_inst_label = st.selectbox("Échéance à modifier / rembourser", list(inst_options.keys()), key="manage_installment_select")
                installment_id = inst_options[selected_inst_label]
                selected_row = idf[idf['id'].map(safe_int_id)==installment_id].iloc[0]

                with st.form("edit_installment_form"):
                    dval = pd.to_datetime(selected_row['due_date']).date()
                    edit_due = st.date_input("Échéance (date modifiable)", value=dval)
                    edit_due_amount = st.number_input("Montant de l'échéance", min_value=0.01, value=float(selected_row['amount_due']), step=500.0)
                    edit_paid = st.number_input("Montant remboursé", min_value=0.0, value=float(selected_row['amount_paid'] or 0), step=500.0)
                    current_pay_date = pd.to_datetime(selected_row['payment_date']).date() if pd.notna(selected_row['payment_date']) and str(selected_row['payment_date']).strip() else date.today()
                    edit_pay_date = st.date_input("Date du remboursement", value=current_pay_date)
                    edit_inst_note = st.text_input("Note", value=str(selected_row['note'] or ''))
                    c1,c2 = st.columns(2)
                    save_inst = c1.form_submit_button("💾 Enregistrer échéance / paiement")
                    clear_payment = c2.form_submit_button("↩️ Marquer non payé")
                    if save_inst or clear_payment:
                        try:
                            if clear_payment:
                                update_installment(installment_id, edit_due, edit_due_amount, 0, None, edit_inst_note)
                            else:
                                update_installment(installment_id, edit_due, edit_due_amount, edit_paid, edit_pay_date, edit_inst_note)
                            st.success("Échéance et remboursement mis à jour.")
                            st.rerun()
                        except Exception as exc:
                            st.error(str(exc))

                total_due = float(pd.to_numeric(idf['amount_due'], errors='coerce').fillna(0).sum())
                total_paid = float(pd.to_numeric(idf['amount_paid'], errors='coerce').fillna(0).sum())
                c1,c2,c3 = st.columns(3)
                c1.metric("Total prévu", money(total_due))
                c2.metric("Total remboursé", money(total_paid))
                c3.metric("Reste à payer", money(max(total_due-total_paid,0)))


# ============================================================
# RAPPELS WHATSAPP
# ============================================================

elif page == "Rappels WhatsApp":
    brand_hero("Rappels WhatsApp & espace membre", "Envoyez les rappels sur WhatsApp et déposez automatiquement une copie dans l'espace membre.", compact=True)
    st.info(f"Numéro administratif configuré : +{WHATSAPP}")

    st.subheader("📅 Rappel mensuel des cotisations")
    if date.today().day == 8:
        st.success("Nous sommes le 8 : journée prévue pour les rappels mensuels.")
    if st.button("📨 Envoyer les rappels mensuels WhatsApp + espace membre"):
        mdf = get_members(True)
        results = []
        for _, row in mdf.iterrows():
            msg = contribution_message(row["full_name"])
            wa_status = "Non envoyé"
            try:
                send_whatsapp(row["phone"], msg)
                wa_status = "Envoyé"
                wa_sent = True
            except Exception as exc:
                wa_sent = False
                wa_status = f"Erreur : {exc}"
            create_member_reminder(row["id"], "cotisation", "Rappel de cotisation", msg, date.today(), wa_sent)
            results.append({"Membre": row["full_name"], "WhatsApp": wa_status, "Espace membre": "Ajouté"})
        st.dataframe(pd.DataFrame(results), use_container_width=True, hide_index=True)

    st.divider()
    mdf = get_members(True)
    if not mdf.empty:
        member_options = build_member_options(mdf)
        selected = st.selectbox("Membre", list(member_options.keys()), key="reminder_member")
        member_id = member_options[selected]
        member_row = mdf[mdf['id'].map(safe_int_id)==member_id].iloc[0]

        tab_c, tab_r, tab_p = st.tabs(["💰 Cotisation", "💳 Remboursement", "✍️ Remarque personnelle"])
        with tab_c:
            msg = contribution_message(member_row['full_name'])
            st.text_area("Message prérempli", value=msg, height=170, key="contribution_reminder_msg")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("📱 WhatsApp", key="wa_contribution"):
                    try:
                        send_whatsapp(member_row["phone"], msg)
                        create_member_reminder(member_id, "cotisation", "Rappel de cotisation", msg, date.today(), True)
                        st.success("Rappel envoyé sur WhatsApp et ajouté à l'espace membre.")
                    except Exception as exc:
                        st.error(str(exc))
            with c2:
                if st.button("👤 Espace membre", key="space_contribution"):
                    create_member_reminder(member_id, "cotisation", "Rappel de cotisation", msg, date.today(), False)
                    st.success("Rappel ajouté à l'espace membre.")

        with tab_r:
            rdf = all_installments()
            rdf = rdf[rdf['member_id'].map(safe_int_id)==member_id].copy() if not rdf.empty else rdf
            if rdf.empty:
                st.info("Aucune échéance de remboursement pour ce membre.")
            else:
                rdf['remaining'] = (pd.to_numeric(rdf['amount_due'], errors='coerce').fillna(0)-pd.to_numeric(rdf['amount_paid'], errors='coerce').fillna(0)).clip(lower=0)
                pending = rdf[rdf['remaining'] > 0.01].copy()
                if pending.empty:
                    st.success("Toutes les échéances de ce membre sont soldées.")
                else:
                    inst_options = {f"Échéance #{int(r['installment_number'])} — {r['due_date']} — reste {money(r['remaining'])}": r for _,r in pending.iterrows()}
                    chosen = st.selectbox("Échéance à rappeler", list(inst_options.keys()), key="reminder_installment")
                    rr = inst_options[chosen]
                    msg = loan_message(member_row['full_name'], float(rr['remaining']), pd.to_datetime(rr['due_date']).date())
                    st.text_area("Message prérempli de remboursement", value=msg, height=190, key="loan_reminder_msg")
                    c1, c2 = st.columns(2)
                    with c1:
                        if st.button("📱 WhatsApp", key="wa_loan"):
                            try:
                                send_whatsapp(member_row["phone"], msg)
                                create_member_reminder(member_id, "remboursement", "Rappel de remboursement", msg, pd.to_datetime(rr["due_date"]).date(), True)
                                st.success("Rappel envoyé sur WhatsApp et ajouté à l'espace membre.")
                            except Exception as exc:
                                st.error(str(exc))
                    with c2:
                        if st.button("👤 Espace membre", key="space_loan"):
                            create_member_reminder(member_id, "remboursement", "Rappel de remboursement", msg, pd.to_datetime(rr["due_date"]).date(), False)
                            st.success("Rappel ajouté à l'espace membre.")

        with tab_p:
            default = f"Bonjour {member_row['full_name']},\n\n"
            custom = st.text_area("Votre remarque", value=default, height=200, key="personal_remark")
            st.caption("Le message est entièrement modifiable avant l'envoi.")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("📱 WhatsApp avec cette remarque", key="wa_remark"):
                    try:
                        send_whatsapp(member_row["phone"], custom)
                        create_member_reminder(member_id, "remarque", "Remarque de l'administration", custom, None, True)
                        st.success("Remarque envoyée sur WhatsApp et ajoutée à l'espace membre.")
                    except Exception as exc:
                        st.error(str(exc))
            with c2:
                if st.button("👤 Publier dans l'espace membre", key="space_remark"):
                    create_member_reminder(member_id, "remarque", "Remarque de l'administration", custom, None, False)
                    st.success("Remarque publiée dans l'espace membre.")


# ============================================================
# RAPPORT GLOBAL
# ============================================================

# ============================================================
# COMMUNICATION ADMIN ↔ MEMBRES
# ============================================================

elif page == "Communication":
    brand_hero("Communication", "Un espace de dialogue entre l'administration et chaque membre.", compact=True)
    ctab1, ctab2 = st.tabs(["📥 Messages reçus", "📨 Écrire à un membre"])

    with ctab1:
        inbox = get_member_messages()
        if inbox.empty:
            st.info("Aucun message ou réclamation.")
        else:
            for _, r in inbox.iterrows():
                sender = r["member_name"] or "Membre"
                kind = "Réclamation" if r["message_type"] == "reclamation" else "Message"
                with st.container(border=True):
                    st.markdown(f"**👤 {sender} — {kind} — {r['subject'] or 'Sans objet'}**")
                    st.write(str(r["message"]))
                    st.caption(f"Reçu le {r['created_at']}")
                    if not bool(r["is_read"]):
                        mark_message_read(r["id"])
                    with st.form(f"reply_form_{int(r['id'])}"):
                        reply = st.text_area("Réponse", height=100, key=f"reply_{int(r['id'])}")
                        send_reply = st.form_submit_button("📨 Répondre au membre")
                        if send_reply:
                            try:
                                send_admin_message_to_member(int(r["member_id"]), f"Re: {r['subject'] or 'Message'}", reply)
                                st.success("Réponse envoyée dans l'espace membre.")
                                st.rerun()
                            except Exception as exc:
                                st.error(str(exc))

    with ctab2:
        mdf = get_members(True)
        if mdf.empty:
            st.info("Aucun membre actif.")
        else:
            options = build_member_options(mdf, include_phone=False)
            selected = st.selectbox("Membre destinataire", list(options.keys()), key="communication_member")
            target_id = options[selected]
            with st.form("admin_message_form"):
                subject = st.text_input("Objet")
                message_type = st.selectbox("Type", ["Message", "Réclamation / suivi", "Information"])
                message = st.text_area("Message", height=160)
                send = st.form_submit_button("📨 Envoyer dans l'espace membre", type="primary")
                if send:
                    try:
                        send_admin_message_to_member(
                            target_id, subject, message,
                            "reclamation" if message_type == "Réclamation / suivi" else "message"
                        )
                        st.success("Message envoyé au membre.")
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))


elif page == "Rapport global":

    brand_hero(
        "Rapport global",
        "Une vue complète de toutes les cotisations, tous les emprunts, les remboursements et les échéances.",
        compact=True,
    )

    data = global_report_data()

    st.markdown(
        """<div class="global-report-card">
        <div style="font-size:1.25rem;font-weight:900;">📊 Situation globale</div>
        <div style="margin-top:8px;opacity:.92;line-height:1.55;">
        Le disponible net est calculé selon les opérations enregistrées : <b>cotisations + remboursements reçus − montants empruntés</b>.
        Les données correspondent uniquement aux opérations présentes dans la base.
        </div></div>""",
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total cotisé", money(data["total_contributed"]))
    c2.metric("Total emprunté", money(data["total_borrowed"]))
    c3.metric("Somme reçue", money(data["total_received"]))
    c4.metric("Reste à recevoir", money(data["outstanding"]))
    c5.metric("Disponible net", money(data["available"]))

    st.write(f"**Mois de cotisation enregistrés :** {', '.join(data['contribution_months']) if data['contribution_months'] else 'Aucun'}")
    st.write(f"**Membres :** {len(data['members'])}  •  **Échéances :** {len(data['installments'])}")

    st.markdown('<div class="section-title">👥 Synthèse par membre</div>', unsafe_allow_html=True)
    display_summary = data["summary"].copy()
    if not display_summary.empty:
        for col in ["Total cotisé", "Total emprunté", "Somme reçue", "Échéances prévues", "Reste à recevoir", "Disponible net"]:
            display_summary[col] = display_summary[col].map(money)
    st.dataframe(display_summary, use_container_width=True, hide_index=True)

    b1, b2 = st.columns(2)
    with b1:
        st.download_button(
            "📥 Télécharger le rapport global PDF",
            data=generate_global_pdf(),
            file_name=f"rapport_global_epargne_{date.today().isoformat()}.pdf",
            mime="application/pdf",
            use_container_width=True,
            type="primary",
        )
    with b2:
        st.download_button(
            "📊 Télécharger le rapport global Excel",
            data=generate_global_excel(),
            file_name=f"rapport_global_epargne_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with st.expander("Voir toutes les cotisations"):
        st.dataframe(data["contributions"], use_container_width=True, hide_index=True)
    with st.expander("Voir tous les emprunts"):
        st.dataframe(data["loans"], use_container_width=True, hide_index=True)
    with st.expander("Voir toutes les échéances"):
        st.dataframe(data["installments"], use_container_width=True, hide_index=True)


# ============================================================
# BULLETINS PDF
# ============================================================

elif page == "Bulletins PDF":

    brand_hero("Bulletins PDF", "Un relevé clair et élégant pour chaque membre.", compact=True)

    df = get_members(False)

    if df.empty:

        st.info(
            "Aucun membre."
        )

    else:

        options = {
            f"{r['full_name']} — {r['phone']}": rid
            for _, r in df.iterrows()
            if (rid := safe_int_id(r.get("id"))) is not None
        }

        option_labels = list(options.keys())
        if not option_labels:
            st.warning("Aucun membre disponible pour générer un bulletin.")
            st.stop()

        selected = st.selectbox(
            "Membre",
            option_labels,
            index=0,
            key="bulletin_member_select"
        )

        member_id = options.get(selected)
        if member_id is None:
            st.warning("Le membre sélectionné n'est plus disponible. Actualisez la page.")
            st.stop()

        pdf = generate_member_pdf(
            member_id
        )

        st.download_button(
            "📥 Télécharger le bulletin PDF",
            data=pdf,
            file_name="bulletin_epargne.pdf",
            mime="application/pdf"
        )


# ============================================================
# ADMINISTRATEURS
# ============================================================

elif page == "Administrateurs":

    brand_hero("Administrateurs", "Gérez les administrateurs et les comptes personnels des membres.", compact=True)

    tab_admins, tab_members_accounts = st.tabs(["👨‍💼 Administrateurs", "👤 Comptes membres"])

    with tab_admins:
        st.subheader("Ajouter un administrateur")
        with st.form("admin_form"):
            full_name = st.text_input("Nom complet")
            username = st.text_input("Nom d'utilisateur")
            password = st.text_input("Mot de passe", type="password")
            submit = st.form_submit_button("Ajouter")
            if submit:
                try:
                    add_admin(username, password, full_name)
                    st.success("Administrateur ajouté.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Impossible d'ajouter l'administrateur : {exc}")
        st.divider()
        st.dataframe(get_admins(), use_container_width=True, hide_index=True)

    with tab_members_accounts:
        st.subheader("Authentification des membres")
        st.info("Chaque compte est lié à un seul membre. Un membre connecté ne peut voir ni modifier les données des autres membres.")
        members_df = get_members(False)
        if members_df.empty:
            st.warning("Ajoutez d'abord un membre dans le menu Membres.")
        else:
            member_options = build_member_options(members_df, include_phone=False)
            if not member_options:
                st.info("Aucun membre avec un nom et un identifiant valides. Ajoutez d'abord un membre.")
                st.stop()
            selected_label = st.selectbox(
                "Membre à authentifier",
                list(member_options.keys()),
                key="admin_member_account_select"
            )
            selected_member_id = member_options.get(selected_label)
            if selected_member_id is None:
                st.warning("Le membre sélectionné n'est plus disponible. Actualisez la page.")
                st.stop()
            current = members_df[members_df["id"] == selected_member_id].iloc[0]
            with st.form("member_account_form"):
                member_login = st.text_input("Identifiant membre", value=str(current.get("member_username") or ""), placeholder="ex. issa.membre")
                member_pass = st.text_input("Mot de passe membre", type="password", placeholder="Nouveau mot de passe")
                member_active = st.checkbox("Autoriser la connexion", value=bool(current.get("member_login_active") or False))
                save_member_login = st.form_submit_button("💾 Enregistrer le compte membre", type="primary")
                if save_member_login:
                    try:
                        set_member_login(selected_member_id, member_login, member_pass, member_active)
                        st.success("Compte membre enregistré. Le membre peut maintenant se connecter depuis l'onglet Membre de la page d'accueil.")
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))

            st.markdown("### État des comptes membres")
            status_df = get_member_login_status().copy()
            status_df["État"] = status_df["member_login_active"].map(lambda x: "Actif" if bool(x) else "Désactivé")
            status_df = status_df.rename(columns={"full_name": "Membre", "member_username": "Identifiant"})
            status_df = status_df[["id", "Membre", "phone", "Identifiant", "État"]]
            status_df = status_df.rename(columns={"id": "ID", "phone": "Téléphone"})
            st.dataframe(status_df, use_container_width=True, hide_index=True)

import os
import sqlite3
import base64
import mimetypes
from datetime import date, timedelta
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
ASSET_IMAGE = Path(__file__).with_name("pe.jpeg")

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
        [data-testid="stStatusWidget"] {{
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
    """Lecture SQL compatible avec les deux moteurs."""
    with db() as con:
        return pd.read_sql_query(sql(query), con, params=params or [])


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


@st.cache_data(ttl=15, show_spinner=False)
def member_account_data_cached(member_id):
    """Lecture regroupée et temporairement mise en cache pour un membre."""
    member_id = int(member_id)

    with db() as con:
        member_df = pd.read_sql_query(
            """
            SELECT id, full_name, phone, monthly_target, notes,
                   active, member_username, member_login_active, created_at
            FROM members
            WHERE id=?
            LIMIT 1
            """,
            con,
            params=[member_id],
        )

        cdf = pd.read_sql_query(
            """
            SELECT c.id, c.member_id, m.full_name, c.payment_date,
                   c.month_label, c.amount, c.note
            FROM contributions c
            JOIN members m ON m.id = c.member_id
            WHERE c.member_id=?
            ORDER BY c.payment_date DESC, c.id DESC
            """,
            con,
            params=[member_id],
        )

        ldf = pd.read_sql_query(
            """
            SELECT l.id, l.member_id, m.full_name, l.loan_date,
                   l.principal, l.interest_rate, l.total_due,
                   l.duration_months, l.first_due_date, l.status, l.note
            FROM loans l
            JOIN members m ON m.id=l.member_id
            WHERE l.member_id=?
            ORDER BY l.loan_date DESC, l.id DESC
            """,
            con,
            params=[member_id],
        )

        idf = pd.read_sql_query(
            """
            SELECT i.id, i.loan_id, i.installment_number, i.due_date,
                   i.amount_due, i.amount_paid, i.payment_date, i.note
            FROM loan_installments i
            JOIN loans l ON l.id=i.loan_id
            WHERE l.member_id=?
            ORDER BY i.due_date, i.id
            """,
            con,
            params=[member_id],
        )

    cdf = _clean_contributions_df(cdf)

    if not idf.empty:
        idf["amount_due"] = pd.to_numeric(idf["amount_due"], errors="coerce").fillna(0)
        idf["amount_paid"] = pd.to_numeric(idf["amount_paid"], errors="coerce").fillna(0)
        idf["reste"] = (idf["amount_due"] - idf["amount_paid"]).clip(lower=0)

    total_contributed = float(
        pd.to_numeric(cdf.get("amount", pd.Series(dtype=float)), errors="coerce")
        .fillna(0).sum()
    )
    total_borrowed = float(
        pd.to_numeric(ldf.get("principal", pd.Series(dtype=float)), errors="coerce")
        .fillna(0).sum()
    )
    total_received = float(
        pd.to_numeric(idf.get("amount_paid", pd.Series(dtype=float)), errors="coerce")
        .fillna(0).sum()
    )
    total_due = float(
        pd.to_numeric(idf.get("amount_due", pd.Series(dtype=float)), errors="coerce")
        .fillna(0).sum()
    )

    return {
        "member": member_df,
        "contributions": cdf,
        "loans": ldf,
        "installments": idf,
        "total_contributed": total_contributed,
        "total_borrowed": total_borrowed,
        "total_received": total_received,
        "outstanding": max(total_due - total_received, 0),
    }


def member_account_data(member_id):
    """Retourne uniquement les données du membre authentifié."""
    return member_account_data_cached(int(member_id))

def member_account_page(member_id):
    data = member_account_data(member_id)
    if data["member"].empty:
        st.error("Compte membre introuvable.")
        return
    name = str(data["member"].iloc[0]["full_name"])
    brand_hero("Mon compte", f"Bienvenue {name}. Cette page est personnelle et en lecture seule.", compact=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total cotisé", money(data["total_contributed"]))
    c2.metric("Total emprunté", money(data["total_borrowed"]))
    c3.metric("Remboursements reçus", money(data["total_received"]))
    c4.metric("Reste à payer", money(data["outstanding"]))

    st.info("🔒 Vous ne pouvez consulter que vos propres cotisations, emprunts et échéances. Aucune modification n'est autorisée depuis cet espace.")
    t1, t2, t3 = st.tabs(["💰 Mes cotisations", "💳 Mes emprunts", "📅 Mes échéances"])
    with t1:
        st.dataframe(data["contributions"], use_container_width=True, hide_index=True)
    with t2:
        st.dataframe(data["loans"], use_container_width=True, hide_index=True)
    with t3:
        st.dataframe(data["installments"], use_container_width=True, hide_index=True)


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

    with db() as con:
        if use_supabase():
            cur = con.cursor()
            cur.execute(query)
            rows = cur.fetchall()
            columns = [desc[0] for desc in cur.description]
            cur.close()
            df = pd.DataFrame(rows, columns=columns)
        else:
            df = pd.read_sql_query(query, con)

    return _clean_members_df(df)


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

def add_contribution(
    member_id,
    payment_date,
    amount,
    note
):

    with db() as con:
        con.execute(
            """
            INSERT INTO contributions(
                member_id,
                payment_date,
                amount,
                month_label,
                note
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                member_id,
                payment_date.isoformat(),
                float(amount),
                month_label(payment_date),
                note.strip(),
            )
        )
        con.commit()
        member_account_data_cached.clear()
        refresh_application_data()


def update_contribution(
    contribution_id,
    payment_date,
    amount,
    note
):

    with db() as con:
        con.execute(
            """
            UPDATE contributions
            SET
                payment_date=?,
                amount=?,
                month_label=?,
                note=?
            WHERE id=?
            """,
            (
                payment_date.isoformat(),
                float(amount),
                month_label(payment_date),
                note.strip(),
                contribution_id,
            )
        )
        con.commit()
        member_account_data_cached.clear()


def delete_contribution(contribution_id):

    with db() as con:
        con.execute(
            "DELETE FROM contributions WHERE id=?",
            (contribution_id,)
        )
        con.commit()


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
    """Retourne les cotisations avec le nom réel du membre, sans lignes parasites."""
    query = """
        SELECT
            c.id,
            c.member_id,
            m.full_name,
            c.payment_date,
            c.month_label,
            c.amount,
            c.note
        FROM contributions c
        JOIN members m ON m.id = c.member_id
    """

    params = []
    if member_id is not None:
        query += " WHERE c.member_id=? "
        params.append(member_id)

    query += " ORDER BY c.payment_date DESC, c.id DESC "

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

        for i in range(duration):
            due_date = add_months(first_due_date, i)
            amount = total_due - installment * (duration - 1) if i == duration - 1 else installment
            con.execute(
                """
                INSERT INTO loan_installments(
                    loan_id, installment_number, due_date, amount_due, amount_paid
                )
                VALUES (?, ?, ?, ?, 0)
                """,
                (loan_id, i + 1, due_date.isoformat(), round(amount, 2)),
            )

        con.commit()
        return loan_id


def loans(member_id=None):

    query = """
        SELECT
            l.id,
            l.member_id,
            m.full_name,
            l.loan_date,
            l.principal,
            l.interest_rate,
            l.total_due,
            l.duration_months,
            l.first_due_date,
            l.status,
            l.note
        FROM loans l
        JOIN members m
            ON m.id=l.member_id
    """

    params = []

    if member_id is not None:
        query += " WHERE l.member_id=? "
        params.append(member_id)

    query += """
        ORDER BY
            l.loan_date DESC,
            l.id DESC
    """

    with db() as con:
        return pd.read_sql_query(
            query,
            con,
            params=params
        )


def get_installments(loan_id):

    with db() as con:
        return pd.read_sql_query(
            """
            SELECT
                id,
                loan_id,
                due_date,
                amount_due,
                amount_paid,
                payment_date,
                note
            FROM loan_installments
            WHERE loan_id=?
            ORDER BY due_date, id
            """,
            con,
            params=[loan_id]
        )


def register_installment_payment(
    installment_id,
    amount_paid,
    payment_date,
    note
):

    with db() as con:

        row = con.execute(
            """
            SELECT loan_id
            FROM loan_installments
            WHERE id=?
            """,
            (installment_id,)
        ).fetchone()

        if not row:
            raise ValueError(
                "Échéance introuvable."
            )

        loan_id = row["loan_id"]

        con.execute(
            """
            UPDATE loan_installments
            SET
                amount_paid=?,
                payment_date=?,
                note=?
            WHERE id=?
            """,
            (
                float(amount_paid),
                payment_date.isoformat(),
                note.strip(),
                installment_id,
            )
        )

        total = con.execute(
            """
            SELECT
                SUM(amount_due) AS due,
                SUM(amount_paid) AS paid
            FROM loan_installments
            WHERE loan_id=?
            """,
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
            """
            UPDATE loans
            SET status=?
            WHERE id=?
            """,
            (
                status,
                loan_id,
            )
        )

        con.commit()


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

    with db() as con:
        return pd.read_sql_query(
            """
            SELECT
                id,
                username,
                full_name,
                active,
                created_at
            FROM admins
            ORDER BY full_name
            """,
            con
        )


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

    # Les colonnes peuvent être absentes ou contenir des valeurs texte/NULL.
    # On normalise toujours les montants avant le calcul pour éviter le ValueError.
    amount_series = pd.to_numeric(cdf.get("amount", pd.Series(dtype=float)), errors="coerce").fillna(0)
    principal_series = pd.to_numeric(ldf.get("principal", pd.Series(dtype=float)), errors="coerce").fillna(0)
    due_series = pd.to_numeric(ldf.get("total_due", pd.Series(dtype=float)), errors="coerce").fillna(0)
    total_saved = float(amount_series.sum())
    total_borrowed = float(principal_series.sum())
    total_due = float(due_series.sum())

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
    summary_data = [
        [
            Paragraph("<b>ÉPARGNE</b><br/><font size=15>%s</font>" % money(total_saved), styles["Normal"]),
            Paragraph("<b>EMPRUNTS</b><br/><font size=15>%s</font>" % money(total_borrowed), styles["Normal"]),
            Paragraph("<b>À REMBOURSER</b><br/><font size=15>%s</font>" % money(total_due), styles["Normal"]),
        ]
    ]
    summary = Table(summary_data, colWidths=[59.5 * mm] * 3)
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
    with db() as con:
        return pd.read_sql_query(
            f"""
            SELECT
                i.id, i.loan_id, i.installment_number, i.due_date,
                i.amount_due, i.amount_paid, i.payment_date, i.note,
                l.member_id, m.full_name
            FROM {tables[0]} i
            JOIN {tables[1]} l ON l.id=i.loan_id
            JOIN {tables[2]} m ON m.id=l.member_id
            ORDER BY i.due_date, i.id
            """,
            con,
        )


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
# SIDEBAR
# ============================================================

user_role = st.session_state.user.get("role", "admin")

st.sidebar.markdown(
    """
    <div style="text-align:center; padding:10px 0 18px 0;">
        <div style="font-size:2.4rem;">🐷</div>
        <div style="font-size:1.25rem; font-weight:900;">Épargne Étudiant</div>
        <div style="opacity:.82; font-size:.85rem;">Petits efforts, grands projets</div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.sidebar.success(
    f"Connecté : {st.session_state.user['full_name']}"
)

if st.sidebar.button("Se déconnecter"):
    st.session_state.user = None
    st.rerun()

if st.sidebar.button("🔄 Actualiser les données"):
    refresh_application_data()
    st.rerun()

# État réel de la base utilisée par l'application.
if use_supabase():
    st.sidebar.success("🟢 Supabase PostgreSQL actif")
    if st.sidebar.button("🔎 Tester Supabase"):
        try:
            row, tables = supabase_health_check()
            missing_tables = [name for name, ok in tables.items() if not ok]
            if missing_tables:
                st.sidebar.error("Tables manquantes : " + ", ".join(missing_tables))
            else:
                st.sidebar.success("Supabase OK — base : " + str(row.get("db")))
        except Exception as exc:
            st.sidebar.error("Test Supabase échoué : " + str(exc))
else:
    st.sidebar.error("🔴 Supabase non configuré")

st.sidebar.divider()

if user_role == "member":
    page = "Mon compte"
    st.sidebar.info("👤 Espace membre — lecture seule")
else:
    page = st.sidebar.radio(
        "Menu",
        [
            "Tableau de bord",
            "Membres",
            "Cotisations",
            "Emprunts",
            "Rappels WhatsApp",
            "Rapport global",
            "Bulletins PDF",
            "Administrateurs",
        ]
    )


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

    brand_hero("Les cotisations", "Enregistrez chaque versement réel et gardez une trace précise de l'épargne.")

    mdf = get_members(True)

    if mdf.empty:

        st.warning(
            "Aucun membre actif trouvé."
        )
        st.info(
            "Si le membre apparaît dans Supabase mais pas ici, utilisez "
            "« 🔄 Actualiser les données » dans la barre latérale. "
            "Les membres avec active = TRUE ou NULL sont considérés comme actifs."
        )

    else:

        member_options = build_member_options(mdf)

        with st.form("contribution_form"):

            selected = st.selectbox(
                "Membre",
                list(member_options.keys())
            )

            member_id = member_options.get(selected)
            if member_id is None:
                st.warning("Le membre sélectionné n'est plus disponible. Actualisez la page.")
                st.stop()

            payment_date = st.date_input(
                "Date réelle du paiement",
                value=date.today()
            )

            amount = st.number_input(
                "Montant réellement versé",
                min_value=0.0,
                step=500.0
            )

            note = st.text_input(
                "Note"
            )

            submit = st.form_submit_button(
                "Enregistrer la cotisation"
            )

            if submit:

                add_contribution(
                    member_id,
                    payment_date,
                    amount,
                    note
                )

                st.success(
                    "Cotisation enregistrée."
                )

                st.rerun()

        st.divider()

        cdf = contributions()

        # Tableau lisible : le membre et son montant sont visibles directement.
        display_cdf = cdf.copy()
        if not display_cdf.empty:
            display_cdf = display_cdf.rename(columns={
                "full_name": "Membre",
                "payment_date": "Date",
                "month_label": "Mois",
                "amount": "Montant",
                "note": "Note",
            })
            display_cdf = display_cdf[["Membre", "Date", "Mois", "Montant", "Note"]]
            display_cdf["Montant"] = pd.to_numeric(
                display_cdf["Montant"], errors="coerce"
            ).fillna(0).map(money)
        else:
            display_cdf = pd.DataFrame(columns=["Membre", "Date", "Mois", "Montant", "Note"])

        st.dataframe(
            display_cdf,
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# EMPRUNTS
# ============================================================

elif page == "Emprunts":

    brand_hero("Les emprunts", "Suivez les prêts, les échéances et les remboursements sans perdre le fil.")

    mdf = get_members(True)

    if mdf.empty:

        st.warning(
            "Aucun membre actif trouvé."
        )
        st.info(
            "Si le membre apparaît dans Supabase mais pas ici, utilisez "
            "« 🔄 Actualiser les données » dans la barre latérale. "
            "Les membres avec active = TRUE ou NULL sont considérés comme actifs."
        )

    else:

        options = {
            f"{r['full_name']} — {r['phone']}": rid
            for _, r in mdf.iterrows()
            if (rid := safe_int_id(r.get("id"))) is not None
        }

        with st.form("loan_form"):

            option_labels = list(options.keys())
            if not option_labels:
                st.warning("Aucun membre sélectionnable.")
                st.stop()

            selected = st.selectbox(
                "Membre",
                option_labels,
                index=0,
                key="loan_member_select"
            )

            member_id = options.get(selected)
            if member_id is None:
                st.warning("Le membre sélectionné n'est plus disponible. Actualisez la page.")
                st.stop()

            loan_date = st.date_input(
                "Date du prêt",
                value=date.today()
            )

            principal = st.number_input(
                "Montant du prêt",
                min_value=1.0,
                step=1000.0
            )

            rate = st.number_input(
                "Taux d'intérêt total (%)",
                min_value=0.0,
                step=0.5
            )

            duration = st.number_input(
                "Nombre d'échéances",
                min_value=1,
                max_value=60,
                value=1
            )

            first_due_date = st.date_input(
                "Première échéance",
                value=date.today()
            )

            note = st.text_input(
                "Note"
            )

            submit = st.form_submit_button(
                "Enregistrer le prêt"
            )

            if submit:

                create_loan(
                    member_id,
                    loan_date,
                    principal,
                    rate,
                    duration,
                    first_due_date,
                    note
                )

                st.success(
                    "Prêt enregistré avec ses échéances."
                )

                st.rerun()

        st.divider()

        ldf = loans()

        st.dataframe(
            ldf,
            use_container_width=True,
            hide_index=True
        )

        if not ldf.empty:

            st.subheader(
                "Remboursement d'une échéance"
            )

            loan_options = {
                f"#{rid} — {r['full_name']} — {money(r['principal'])}": rid
                for _, r in ldf.iterrows()
                if (rid := safe_int_id(r.get("id"))) is not None
            }

            loan_label = st.selectbox(
                "Prêt",
                list(loan_options.keys())
            )

            loan_id = loan_options.get(loan_label)
            if loan_id is None:
                st.warning("Le prêt sélectionné n'est plus disponible. Actualisez les données.")
                st.stop()

            idf = get_installments(loan_id)

            st.dataframe(
                idf,
                use_container_width=True,
                hide_index=True
            )

            if not idf.empty:

                installment_options = {
                    f"#{rid} — {r['due_date']} — dû {money(r['amount_due'])}": rid
                    for _, r in idf.iterrows()
                    if (rid := safe_int_id(r.get("id"))) is not None
                }

                selected_installment = st.selectbox(
                    "Échéance",
                    list(installment_options.keys())
                )

                installment_id = installment_options.get(selected_installment)
                if installment_id is None:
                    st.warning("L'échéance sélectionnée n'est plus disponible. Actualisez les données.")
                    st.stop()

                selected_row = idf[
                    idf["id"] == installment_id
                ].iloc[0]

                with st.form("payment_form"):

                    amount_paid = st.number_input(
                        "Montant payé",
                        min_value=0.0,
                        value=float(
                            selected_row["amount_paid"] or 0
                        ),
                        step=500.0
                    )

                    payment_date = st.date_input(
                        "Date du remboursement",
                        value=date.today()
                    )

                    note = st.text_input(
                        "Note du remboursement"
                    )

                    submit = st.form_submit_button(
                        "Enregistrer le remboursement"
                    )

                    if submit:

                        register_installment_payment(
                            installment_id,
                            amount_paid,
                            payment_date,
                            note
                        )

                        st.success(
                            "Remboursement enregistré."
                        )

                        st.rerun()


# ============================================================
# RAPPELS WHATSAPP
# ============================================================

elif page == "Rappels WhatsApp":

    brand_hero("Rappels WhatsApp", "Des messages simples pour garder le groupe régulier et organisé.", compact=True)

    st.info(
        f"Numéro administratif configuré : +{WHATSAPP}"
    )

    st.subheader(
        "Rappel mensuel du 8"
    )

    st.write(
        "Le programme automatique peut être lancé "
        "le 8 de chaque mois afin d'envoyer un "
        "message privé à chaque membre."
    )

    if date.today().day == 8:
        st.success(
            "Nous sommes le 8 : c'est la journée prévue "
            "pour les rappels mensuels."
        )

    st.subheader(
        "Envoyer maintenant à tous les membres"
    )

    if st.button(
        "📨 Envoyer les rappels maintenant"
    ):

        try:

            results = send_monthly_reminders()

            st.dataframe(
                results,
                use_container_width=True,
                hide_index=True
            )

        except Exception as exc:

            st.error(str(exc))

    st.divider()

    st.subheader(
        "Messages WhatsApp préremplis"
    )

    mdf = get_members(True)

    if not mdf.empty:

        options = {}
        for _, r in mdf.iterrows():
            member_id = safe_int_id(r.get("id"))
            name = str(r.get("full_name") or "").strip()
            if member_id is None or not name:
                continue
            phone = normalize_phone(r.get("phone"))
            label = f"{name} — +{phone}" if phone else name
            options[f"{label} · ID {member_id}"] = r

        if not options:
            st.info("Aucun membre actif avec un nom valide.")
            st.stop()

        selected = st.selectbox(
            "Membre",
            list(options.keys()),
            key="whatsapp_member_select"
        )

        member = options.get(selected)
        if member is None:
            st.warning("Le membre sélectionné n'est plus disponible. Actualisez la page.")
            st.stop()

        message = contribution_message(
            member["full_name"]
        )

        st.text_area(
            "Message",
            value=message,
            height=180,
            key="whatsapp_preview"
        )

        st.link_button(
            "💬 Ouvrir WhatsApp avec le message",
            whatsapp_link(
                member["phone"],
                message
            )
        )


# ============================================================
# RAPPORT GLOBAL
# ============================================================

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

import os
import sqlite3
from datetime import date, timedelta
from contextlib import contextmanager
from io import BytesIO
from pathlib import Path
import urllib.parse

import pandas as pd
import streamlit as st

try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
except ImportError:
    psycopg2 = None
    RealDictCursor = None

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
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
SUPABASE_DB_URL = secret_or_env("SUPABASE_DB_URL", "")
SUPABASE_HOST = secret_or_env("SUPABASE_HOST", "")
SUPABASE_PORT = secret_or_env("SUPABASE_PORT", "5432")
SUPABASE_DATABASE = secret_or_env("SUPABASE_DATABASE", "postgres")
SUPABASE_USER = secret_or_env("SUPABASE_USER", "")
SUPABASE_PASSWORD = secret_or_env("SUPABASE_PASSWORD", "")

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
)


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
                radial-gradient(circle at 95% 0%, rgba(169,212,245,.25), transparent 28%),
                linear-gradient(180deg, #ffffff 0%, {BRAND_CREAM} 100%);
        }}

        [data-testid="stHeader"] {{
            background: rgba(255,255,255,.82);
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
            border-radius: 28px;
            padding: 30px 34px;
            margin: 4px 0 26px 0;
            background: linear-gradient(135deg, rgba(255,255,255,.97), rgba(238,247,253,.96));
            border: 1px solid rgba(18,42,85,.10);
            box-shadow: 0 16px 40px rgba(18,42,85,.10);
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
            font-size: clamp(2rem, 4vw, 3.3rem);
            line-height: 1.02;
            font-weight: 900;
            margin: 0;
        }}

        .brand-subtitle {{
            color: #40536C;
            font-size: 1.12rem;
            margin-top: 12px;
            margin-bottom: 0;
        }}

        .photo-card {{
            border-radius: 24px;
            overflow: hidden;
            box-shadow: 0 14px 32px rgba(18,42,85,.14);
            border: 1px solid rgba(18,42,85,.10);
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
            max-width: 1120px;
            margin: 2vh auto 0 auto;
        }}
        .login-card {{
            background: rgba(255,255,255,.97);
            border: 1px solid rgba(18,42,85,.10);
            border-radius: 30px;
            padding: 34px;
            box-shadow: 0 24px 65px rgba(18,42,85,.14);
        }}
        .login-photo {{
            border-radius: 24px;
            overflow: hidden;
            background: linear-gradient(135deg, #eaf5fc, #fff5f7);
            border: 1px solid rgba(18,42,85,.10);
            min-height: 430px;
            display:flex;
            align-items:center;
            justify-content:center;
        }}
        .login-photo img {{
            width:100%;
            height:430px;
            object-fit:cover;
            display:block;
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
            padding:6px 4px;
        }}
        .login-panel .stButton > button {{
            min-height:48px;
            border-radius:14px;
            font-size:1rem;
        }}
        @media (max-width: 800px) {{
            .login-card {{ padding:20px; border-radius:22px; }}
            .login-photo, .login-photo img {{ min-height:250px; height:250px; }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


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
    def __init__(self, cursor):
        self.cursor = cursor

    def execute(self, query, params=None):
        return self.cursor.execute(query.replace("?", "%s"), params)

    def executemany(self, query, params=None):
        return self.cursor.executemany(query.replace("?", "%s"), params)

    def fetchone(self):
        return self.cursor.fetchone()

    def fetchall(self):
        return self.cursor.fetchall()

    def __iter__(self):
        return iter(self.cursor)


class PostgresConnectionAdapter:
    def __init__(self, con):
        self._con = con

    def cursor(self, *args, **kwargs):
        return PostgresCursorAdapter(self._con.cursor(*args, **kwargs))

    def execute(self, query, params=None):
        return self._con.cursor().execute(query.replace("?", "%s"), params)

    def executemany(self, query, params=None):
        return self._con.cursor().executemany(query.replace("?", "%s"), params)

    def commit(self):
        return self._con.commit()

    def rollback(self):
        return self._con.rollback()

    def close(self):
        return self._con.close()


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
        f"sslmode=require"
    )


def use_supabase():
    """Indique si une connexion Supabase/PostgreSQL peut être utilisée."""
    configured = bool(
        SUPABASE_DB_URL
        or (
            SUPABASE_HOST
            and SUPABASE_USER
            and SUPABASE_PASSWORD
        )
    )
    return configured and psycopg2 is not None


@contextmanager
def db():
    """Connexion PostgreSQL Supabase si psycopg2 est installé, sinon SQLite."""
    if use_supabase():
        raw_con = psycopg2.connect(postgres_dsn(), cursor_factory=RealDictCursor)
        con = PostgresConnectionAdapter(raw_con)
        try:
            yield con
            con.commit()
        except Exception:
            con.rollback()
            raise
        finally:
            con.close()
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
            "total_interest_rate": "REAL DEFAULT 0",
            "installments_count": "INTEGER DEFAULT 1",
            "first_due_date": "TEXT",
            "note": "TEXT",
            "created_at": "TEXT",
        },
        "loan_installments": {
            "due_date": "TEXT",
            "paid_date": "TEXT",
            "paid_amount": "REAL DEFAULT 0",
            "note": "TEXT",
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


def create_supabase_schema():
    """Crée automatiquement toutes les tables Supabase au démarrage."""
    if not use_supabase():
        return

    statements = [
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
            total_interest_rate NUMERIC(8,4) NOT NULL DEFAULT 0,
            installments_count INTEGER NOT NULL DEFAULT 1,
            first_due_date DATE NOT NULL,
            note TEXT,
            created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS public.loan_installments (
            id BIGSERIAL PRIMARY KEY,
            loan_id BIGINT NOT NULL REFERENCES public.loans(id) ON DELETE CASCADE,
            installment_number INTEGER NOT NULL,
            due_date DATE NOT NULL,
            expected_amount NUMERIC(14,2) NOT NULL DEFAULT 0,
            paid_date DATE,
            paid_amount NUMERIC(14,2) NOT NULL DEFAULT 0,
            note TEXT,
            created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
        )
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

    with db() as con:
        cur = con.cursor()
        for statement in statements:
            cur.execute(statement)

        # Administrateur initial.
        cur.execute(
            """
            INSERT INTO public.admins (username, password, full_name)
            VALUES (?, ?, ?)
            ON CONFLICT (username) DO NOTHING
            """,
            (ADMIN_USERNAME, ADMIN_PASSWORD, ADMIN_NAME),
        )

        # Migration douce : ajoute les colonnes qui pourraient manquer
        # si les tables existaient déjà dans une ancienne version.
        alter_statements = [
            "ALTER TABLE public.members ADD COLUMN IF NOT EXISTS phone TEXT",
            "ALTER TABLE public.members ADD COLUMN IF NOT EXISTS monthly_target NUMERIC(14,2) DEFAULT 0",
            "ALTER TABLE public.members ADD COLUMN IF NOT EXISTS notes TEXT",
            "ALTER TABLE public.members ADD COLUMN IF NOT EXISTS active BOOLEAN DEFAULT TRUE",
            "ALTER TABLE public.members ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ DEFAULT NOW()",

            "ALTER TABLE public.contributions ADD COLUMN IF NOT EXISTS payment_date DATE",
            "ALTER TABLE public.contributions ADD COLUMN IF NOT EXISTS month_label TEXT",
            "ALTER TABLE public.contributions ADD COLUMN IF NOT EXISTS note TEXT",
            "ALTER TABLE public.contributions ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ DEFAULT NOW()",

            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS loan_date DATE",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS total_interest_rate NUMERIC(8,4) DEFAULT 0",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS installments_count INTEGER DEFAULT 1",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS first_due_date DATE",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS note TEXT",
            "ALTER TABLE public.loans ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ DEFAULT NOW()",

            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS installment_number INTEGER DEFAULT 1",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS due_date DATE",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS expected_amount NUMERIC(14,2) DEFAULT 0",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS paid_date DATE",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS paid_amount NUMERIC(14,2) DEFAULT 0",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS note TEXT",
            "ALTER TABLE public.loan_installments ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ DEFAULT NOW()",
        ]

        for statement in alter_statements:
            try:
                cur.execute(statement)
            except Exception:
                # Une colonne peut être incompatible avec une ancienne structure.
                # Les tables principales ont déjà été créées avec le bon schéma.
                con.rollback()
                cur = con.cursor()

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
                    cur.execute("""
                        SELECT id FROM public.members
                        WHERE full_name=? AND COALESCE(phone,'')=COALESCE(?,'')
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
                        cur.execute("""
                            INSERT INTO public.loans
                            (member_id,loan_date,principal,total_interest_rate,
                             installments_count,first_due_date,note)
                            VALUES (?,?,?,?,?,?,?) RETURNING id
                        """,(
                            mid,l["loan_date"],l["principal"],
                            l["total_interest_rate"],l["installments_count"],
                            l["first_due_date"],l["note"] if "note" in l.keys() else None,
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
                            (loan_id,installment_number,due_date,expected_amount,
                             paid_date,paid_amount,note)
                            VALUES (?,?,?,?,?,?,?)
                        """,(
                            lid,i["installment_number"],i["due_date"],
                            i["expected_amount"],i["paid_date"],i["paid_amount"],
                            i["note"] if "note" in i.keys() else None,
                        ))

            cur.execute("""
                INSERT INTO public.app_migrations(key)
                VALUES (?)
                ON CONFLICT(key) DO NOTHING
            """,("sqlite_to_supabase_v1",))
            cur.close()
    finally:
        s.close()


def init_db():
    if use_supabase():
        try:
            create_supabase_schema()
            auto_migrate_sqlite_to_supabase()
        except Exception as exc:
            st.error("Erreur de connexion ou de préparation Supabase.")
            st.code(str(exc))
            st.stop()
        return

    with db() as con:
        con.executescript(
            """
            CREATE TABLE IF NOT EXISTS admins (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                full_name TEXT NOT NULL,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS members (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                phone TEXT NOT NULL DEFAULT '777521969',
                monthly_target REAL NOT NULL DEFAULT 0,
                notes TEXT,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS contributions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                member_id INTEGER NOT NULL,
                payment_date TEXT NOT NULL,
                amount REAL NOT NULL DEFAULT 0,
                month_label TEXT NOT NULL,
                note TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS loans (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                member_id INTEGER NOT NULL,
                loan_date TEXT NOT NULL,
                principal REAL NOT NULL,
                total_interest_rate REAL NOT NULL DEFAULT 0,
                installments_count INTEGER NOT NULL DEFAULT 1,
                first_due_date TEXT NOT NULL,
                note TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS loan_installments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                loan_id INTEGER NOT NULL,
                installment_number INTEGER NOT NULL,
                due_date TEXT NOT NULL,
                expected_amount REAL NOT NULL DEFAULT 0,
                paid_date TEXT,
                paid_amount REAL NOT NULL DEFAULT 0,
                note TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            """
        )
        migrate_database(con)
        con.execute(
            """
            INSERT OR IGNORE INTO admins (username, password, full_name)
            VALUES (?, ?, ?)
            """,
            (ADMIN_USERNAME, ADMIN_PASSWORD, ADMIN_NAME),
        )



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


# ============================================================
# MEMBRES
# ============================================================

def get_members(active_only=False):

    query = """
        SELECT
            id,
            full_name,
            phone,
            monthly_target,
            notes,
            active,
            created_at
        FROM members
    """

    if active_only:
        query += " WHERE active=TRUE "

    query += " ORDER BY full_name "

    with db() as con:
        return pd.read_sql_query(query, con)


def add_member(name, phone, target, notes):

    phone = normalize_phone(phone)

    if not name.strip():
        raise ValueError("Le nom du membre est obligatoire.")

    if not phone:
        raise ValueError("Le numéro WhatsApp est obligatoire.")

    with db() as con:
        con.execute(
            """
            INSERT INTO members(
                full_name,
                phone,
                monthly_target,
                notes
            )
            VALUES (?, ?, ?, ?)
            """,
            (
                name.strip(),
                phone,
                float(target or 0),
                notes.strip(),
            )
        )
        con.commit()


def update_member(
    member_id,
    name,
    phone,
    target,
    notes,
    active
):

    with db() as con:
        con.execute(
            """
            UPDATE members
            SET
                full_name=?,
                phone=?,
                monthly_target=?,
                notes=?,
                active=?
            WHERE id=?
            """,
            (
                name.strip(),
                normalize_phone(phone),
                float(target or 0),
                notes.strip(),
                bool(active),
                member_id,
            )
        )
        con.commit()


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


def delete_contribution(contribution_id):

    with db() as con:
        con.execute(
            "DELETE FROM contributions WHERE id=?",
            (contribution_id,)
        )
        con.commit()


def contributions(member_id=None):

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
        JOIN members m
            ON m.id = c.member_id
    """

    params = []

    if member_id is not None:
        query += " WHERE c.member_id=? "
        params.append(member_id)

    query += """
        ORDER BY
            c.payment_date DESC,
            c.id DESC
    """

    with db() as con:
        return pd.read_sql_query(
            query,
            con,
            params=params
        )


# ============================================================
# EMPRUNTS
# ============================================================

def create_loan(
    member_id,
    loan_date,
    principal,
    rate,
    duration,
    first_due_date,
    note
):

    principal = float(principal)
    rate = float(rate)
    duration = int(duration)

    total_due = principal * (1 + rate / 100)
    installment = total_due / duration

    with db() as con:

        cursor = con.execute(
            """
            INSERT INTO loans(
                member_id,
                loan_date,
                principal,
                interest_rate,
                total_due,
                duration_months,
                first_due_date,
                note
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                member_id,
                loan_date.isoformat(),
                principal,
                rate,
                total_due,
                duration,
                first_due_date.isoformat(),
                note.strip(),
            )
        )

        loan_id = cursor.lastrowid

        for i in range(duration):

            due_date = add_months(
                first_due_date,
                i
            )

            if i == duration - 1:
                amount = (
                    total_due
                    - installment * (duration - 1)
                )
            else:
                amount = installment

            con.execute(
                """
                INSERT INTO loan_installments(
                    loan_id,
                    due_date,
                    amount_due,
                    amount_paid
                )
                VALUES (?, ?, ?, 0)
                """,
                (
                    loan_id,
                    due_date.isoformat(),
                    round(amount, 2),
                )
            )

        con.commit()


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

        member_id = int(member["id"])

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

    total_saved = float(cdf["amount"].sum()) if not cdf.empty else 0
    total_borrowed = float(ldf["principal"].sum()) if not ldf.empty else 0
    total_due = float(ldf["total_due"].sum()) if not ldf.empty else 0

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
    contribution_data = [["Date", "Mois", "Montant", "Note"]]
    for _, row in cdf.iterrows():
        contribution_data.append([
            str(row["payment_date"]),
            str(row["month_label"]),
            money(row["amount"]),
            str(row["note"] or ""),
        ])
    if len(contribution_data) == 1:
        contribution_data.append(["-", "-", "-", "Aucune cotisation enregistrée"])

    table = Table(
        contribution_data,
        repeatRows=1,
        colWidths=[29 * mm, 29 * mm, 35 * mm, 86 * mm],
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
# INTERFACE
# ============================================================

init_db()

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
            st.markdown('<div class="login-photo">', unsafe_allow_html=True)
            st.image(str(ASSET_IMAGE), use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
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
        st.markdown('<div class="section-title">🔐 Se connecter</div>', unsafe_allow_html=True)

        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Nom d'utilisateur", placeholder="Votre nom d'utilisateur")
            password = st.text_input("Mot de passe", type="password", placeholder="Votre mot de passe")
            submitted = st.form_submit_button("Se connecter", type="primary", use_container_width=True)

            if submitted:
                try:
                    user = authenticate(username, password)
                except Exception as exc:
                    st.error("Connexion impossible. Vérifiez la configuration Supabase et les tables.")
                    st.code(str(exc))
                    user = None

                if user:
                    st.session_state.user = user
                    st.rerun()
                elif username or password:
                    st.error("Identifiants incorrects.")

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


st.sidebar.divider()

page = st.sidebar.radio(
    "Menu",
    [
        "Tableau de bord",
        "Membres",
        "Cotisations",
        "Emprunts",
        "Rappels WhatsApp",
        "Bulletins PDF",
        "Administrateurs",
    ]
)


# ============================================================
# TABLEAU DE BORD
# ============================================================

if page == "Tableau de bord":

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

                    add_member(
                        name,
                        phone,
                        target,
                        notes
                    )

                    st.success(
                        "Membre ajouté."
                    )

                    st.rerun()

                except Exception as exc:

                    st.error(str(exc))

    st.divider()

    df = get_members(False)

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )

    if not df.empty:

        st.subheader("Modifier un membre")

        options = {
            f"{r['full_name']} — {r['phone']}":
            int(r["id"])
            for _, r in df.iterrows()
        }

        selected = st.selectbox(
            "Membre",
            list(options.keys())
        )

        member_id = options[selected]

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
            "Ajoutez d'abord un membre."
        )

    else:

        member_options = {
            f"{r['full_name']} — {r['phone']}":
            int(r["id"])
            for _, r in mdf.iterrows()
        }

        with st.form("contribution_form"):

            selected = st.selectbox(
                "Membre",
                list(member_options.keys())
            )

            member_id = member_options[selected]

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

        st.dataframe(
            cdf,
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
            "Ajoutez d'abord un membre."
        )

    else:

        options = {
            f"{r['full_name']} — {r['phone']}":
            int(r["id"])
            for _, r in mdf.iterrows()
        }

        with st.form("loan_form"):

            selected = st.selectbox(
                "Membre",
                list(options.keys())
            )

            member_id = options[selected]

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
                f"#{int(r['id'])} — {r['full_name']} — {money(r['principal'])}":
                int(r["id"])
                for _, r in ldf.iterrows()
            }

            loan_label = st.selectbox(
                "Prêt",
                list(loan_options.keys())
            )

            loan_id = loan_options[loan_label]

            idf = get_installments(loan_id)

            st.dataframe(
                idf,
                use_container_width=True,
                hide_index=True
            )

            if not idf.empty:

                installment_options = {
                    f"#{int(r['id'])} — {r['due_date']} — dû {money(r['amount_due'])}":
                    int(r["id"])
                    for _, r in idf.iterrows()
                }

                selected_installment = st.selectbox(
                    "Échéance",
                    list(installment_options.keys())
                )

                installment_id = installment_options[
                    selected_installment
                ]

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

        options = {
            f"{r['full_name']} — +{normalize_phone(r['phone'])}":
            r
            for _, r in mdf.iterrows()
        }

        selected = st.selectbox(
            "Membre",
            list(options.keys())
        )

        member = options[selected]

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
            f"{r['full_name']} — {r['phone']}":
            int(r["id"])
            for _, r in df.iterrows()
        }

        selected = st.selectbox(
            "Membre",
            list(options.keys())
        )

        member_id = options[selected]

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

    brand_hero("Administrateurs", "Gérez les accès à l’espace de suivi.", compact=True)

    st.subheader(
        "Ajouter un administrateur"
    )

    with st.form("admin_form"):

        full_name = st.text_input(
            "Nom complet"
        )

        username = st.text_input(
            "Nom d'utilisateur"
        )

        password = st.text_input(
            "Mot de passe",
            type="password"
        )

        submit = st.form_submit_button(
            "Ajouter"
        )

        if submit:

            try:

                add_admin(
                    username,
                    password,
                    full_name
                )

                st.success(
                    "Administrateur ajouté."
                )

                st.rerun()

            except sqlite3.IntegrityError:

                st.error(
                    "Ce nom d'utilisateur existe déjà."
                )

    st.divider()

    st.dataframe(
        get_admins(),
        use_container_width=True,
        hide_index=True
    )

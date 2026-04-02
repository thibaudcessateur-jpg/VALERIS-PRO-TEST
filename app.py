from __future__ import annotations

import json
import os
import sys
import textwrap
import importlib.util
import itertools
import uuid  # FIXED: stable UUID keys for fund lines (Bug 5)
from datetime import date
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

if importlib.util.find_spec("matplotlib") is not None:
    import matplotlib.pyplot as plt
    MATPLOTLIB_AVAILABLE = True
    MATPLOTLIB_ERROR = ""
else:
    plt = None
    MATPLOTLIB_AVAILABLE = False
    MATPLOTLIB_ERROR = "matplotlib non installé"
import numpy as np
import pandas as pd
import requests
import streamlit as st
import altair as alt
if importlib.util.find_spec("pypfopt") is not None:
    from pypfopt import EfficientFrontier, risk_models, expected_returns
    PYPFOPT_AVAILABLE = True
    PYPFOPT_ERROR = ""
else:
    EfficientFrontier = None
    risk_models = None
    expected_returns = None
    PYPFOPT_AVAILABLE = False
    PYPFOPT_ERROR = "pyportfolioopt non installé"
if importlib.util.find_spec("reportlab") is not None:
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfgen import canvas
    REPORTLAB_AVAILABLE = True
    REPORTLAB_ERROR = ""
else:
    SimpleDocTemplate = Paragraph = Spacer = Image = Table = TableStyle = PageBreak = None
    A4 = None
    getSampleStyleSheet = None
    colors = None
    ParagraphStyle = None
    canvas = None
    REPORTLAB_AVAILABLE = False
    REPORTLAB_ERROR = "reportlab non installé"

try:
    import mstarpy
    MSTARPY_AVAILABLE = True
except ImportError:
    mstarpy = None  # type: ignore[assignment]
    MSTARPY_AVAILABLE = False

try:
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    go = None  # type: ignore[assignment]
    PLOTLY_AVAILABLE = False

# ------------------------------------------------------------
# Constantes & univers de fonds recommandés
# ------------------------------------------------------------
TODAY = pd.Timestamp.today().normalize()
APP_TITLE = "Comparateur de portefeuilles"
ANNUAL_FEE_EURO_PCT = 0.9
ANNUAL_FEE_UC_PCT = 1.2
RISK_FREE_RATE_FALLBACK = 0.026  # 2.6 % — valeur de repli si API Bund indisponible

# Registre des contrats disponibles
# Clé = label affiché à l'utilisateur
# Valeur = dict avec chemin des fichiers
CONTRACTS_REGISTRY: Dict[str, Any] = {
    "Linxea Avenir 2": {
        "path": "linxea/avenir2",
        "funds_filename": "Liste des fonds LINXEA AVENIR 2.csv",
        "euro_funds": {
            "Suravenir Opportunités 2": "suravenir_opportunites2_historique.csv",
            "Suravenir Rendement 2":    "suravenir_rendement2_historique.csv",
        },
        "assureur": "Suravenir",
        "entry_fee_max_pct": 0.0,
    },
    "Linxea Spirit 2": {
        "path": "linxea/spirit2",
        "funds_filename": "SPIRIT2_UC_frais_avec_categories.csv",
        "euro_funds": {
            "Euro Nouvelle Génération": "spirit2_euro_nouvelle_generation.csv",
            "Euro Objectif Climat":     "spirit2_euro_objectif_climat.csv",
        },
        "assureur": "Spirica",
        "entry_fee_max_pct": 0.0,
    },
}

RECO_FUNDS_CORE = [
    ("R-co Valor C EUR", "FR0011253624"),
    ("Vivalor International", "FR0014001LS1"),
    ("CARMIGNAC Investissement A EUR", "FR0010148981"),
    ("FIDELITY FUNDS - WORLD FUND", "LU0069449576"),
    ("CLARTAN VALEURS", "LU1100076550"),
    ("CARMIGNAC PATRIMOINE", "FR0010135103"),
]

RECO_FUNDS_DEF = [
    ("Fonds en euros (EUROFUND)", "EUROFUND"),
    ("SYCOYIELD 2030 RC", "FR001400MCQ6"),
    ("R-Co Target 2029 HY", "FR0014002XJ3"),
    ("Euro Bond 1-3 Years", "LU0321462953"),
]

FUND_NAME_MAP = {isin: name for name, isin in RECO_FUNDS_CORE + RECO_FUNDS_DEF}

# Univers UC explicites (hors fonds en euros)
EQUITY_FUNDS = [isin for _, isin in RECO_FUNDS_CORE]
BOND_FUNDS = [isin for _, isin in RECO_FUNDS_DEF if isin != "EUROFUND"]

# Libellés FR -> codes internes pour l'affectation des versements
ALLOC_LABELS = {
    "Répartition égale": "equal",
    "Personnalisé": "custom",
    "Tout sur une ligne": "single",
}


# ------------------------------------------------------------
# Utils format
# ------------------------------------------------------------

def params_hash(values: Tuple[Any, ...]) -> str:
    try:
        payload = json.dumps(values, default=str, sort_keys=True)
    except Exception:
        payload = repr(values)
    return str(hash(payload))

def to_eur(x: Any) -> str:
    try:
        v = float(x)
    except Exception:
        return "—"
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")
    return s + " €"


def fmt_date(x: Any) -> str:
    try:
        return pd.Timestamp(x).strftime("%d/%m/%Y")
    except Exception:
        return "—"


def fmt_eur_fr(x: Any) -> str:
    try:
        v = float(x)
    except Exception:
        return "—"
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")
    return f"{s} €"


def fmt_pct_fr(x: Any) -> str:
    try:
        v = float(x)
    except Exception:
        return "—"
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")
    return f"{s} %"


# ------------------------------------------------------------
# XIRR
# ------------------------------------------------------------

def _npv(rate: float, cfs: List[Tuple[pd.Timestamp, float]]) -> float:
    t0 = cfs[0][0]
    return sum(cf / ((1 + rate) ** ((t - t0).days / 365.25)) for t, cf in cfs)


def xirr(cash_flows: List[Tuple[pd.Timestamp, float]], guess: float = 0.1) -> Optional[float]:
    if not cash_flows:
        return None
    cfs = sorted(cash_flows, key=lambda x: x[0])
    try:
        r = guess
        for _ in range(100):
            f = _npv(r, cfs)
            h = 1e-6
            f1 = _npv(r + h, cfs)
            d = (f1 - f) / h
            if abs(d) < 1e-12:
                break
            r2 = r - f / d
            if abs(r2 - r) < 1e-9:
                r = r2
                break
            r = r2
        if abs(_npv(r, cfs)) > 1.0:  # NPV non nulle = pas convergé
            return None
        return r
    except Exception:
        return None


# ------------------------------------------------------------
# API EODHD
# ------------------------------------------------------------

def _get_api_key() -> str:
    return st.secrets.get("EODHD_API_KEY", "")


@st.cache_data(show_spinner=False, ttl=3600)
def eodhd_get(path: str, params: Dict[str, Any] | None = None) -> Any:
    base = "https://eodhd.com/api"
    token = _get_api_key()
    p = {"api_token": token, "fmt": "json"}
    if params:
        p.update(params)
    with st.spinner("Chargement EODHD..."):
        r = requests.get(f"{base}{path}", params=p, timeout=20)
    r.raise_for_status()
    try:
        return r.json()
    except Exception:
        return None


@st.cache_data(show_spinner=False, ttl=3600)
def eodhd_search(q: str) -> List[Dict[str, Any]]:
    try:
        js = eodhd_get(f"/search/{q}")
        if isinstance(js, list):
            return js
    except Exception:
        pass
    return []


@st.cache_data(show_spinner=False, ttl=86400)
def _fetch_bund_rate_from_api() -> Optional[float]:
    """
    Tente de récupérer le rendement du Bund allemand 10 ans via EODHD.
    Retourne le taux en décimal (ex: 0.026) ou None si indisponible.
    Essaie plusieurs tickers dans l'ordre en cas d'échec.
    TTL : 24 h (le taux change peu en intra-journalier).
    """
    tickers = ["DE10Y.GBOND", "BUND10Y.INDX", "IRLTLT01DEM156N.FRED"]
    for ticker in tickers:
        try:
            data = eodhd_get(
                f"/eod/{ticker}",
                {"fmt": "json", "order": "d", "limit": 1},
            )
            if isinstance(data, list) and len(data) > 0:
                rate_pct = float(data[0].get("close", 0.0))
                if 0.0 < rate_pct < 15.0:
                    return rate_pct / 100.0
        except Exception:
            continue
    return None


def get_risk_free_rate() -> float:
    """
    Retourne le taux sans risque actif en décimal.
    Priorité :
    1. API EODHD (Bund 10 ans, si disponible)
    2. Saisie manuelle dans la sidebar (RISK_FREE_RATE_MANUAL, stocké en décimal)
    3. Fallback RISK_FREE_RATE_FALLBACK
    """
    api_rate = _fetch_bund_rate_from_api()
    if api_rate is not None:
        return api_rate
    return float(
        st.session_state.get("RISK_FREE_RATE_MANUAL", RISK_FREE_RATE_FALLBACK)
    )


@st.cache_data(show_spinner=False, ttl=86400)
def eodhd_prices_daily(symbol: str) -> pd.DataFrame:
    try:
        js = eodhd_get(f"/eod/{symbol}", params={"period": "d"})
        if not isinstance(js, list) or len(js) == 0:
            return pd.DataFrame()
        df = pd.DataFrame(js)
        df["date"] = pd.to_datetime(df["date"])
        df.set_index("date", inplace=True)
        if "adjusted_close" in df.columns and pd.notnull(df["adjusted_close"]).any():
            df["Close"] = df["adjusted_close"].astype(float)
        elif "close" in df.columns:
            df["Close"] = df["close"].astype(float)
        else:
            return pd.DataFrame()
        return df[["Close"]].sort_index()
    except Exception:
        return pd.DataFrame()


def _symbol_candidates(isin_or_name: str) -> List[str]:
    val = str(isin_or_name).strip()
    if not val:
        return []
    if val.upper() == "EUROFUND":
        return ["EUROFUND"]
    candidates = [f"{val}.EUFUND", f"{val}.FUND", val]
    try:
        res = eodhd_search(val)
        for it in res:
            code = it.get("Code")
            exch = it.get("Exchange")
            if code and exch:
                candidates.append(f"{code}.{exch}")
    except Exception:
        pass
    seen = set()
    uniq = []
    for c in candidates:
        if c not in seen:
            seen.add(c)
            uniq.append(c)
    return uniq


def _get_close_on(df: pd.DataFrame, d: pd.Timestamp) -> float:
    if df.empty:
        return np.nan
    if d in df.index:
        return float(df.loc[d, "Close"])
    after = df.loc[df.index >= d]
    if not after.empty:
        return float(after.iloc[0]["Close"])
    return float(df.iloc[-1]["Close"])


def apply_annual_fee(
    df: pd.DataFrame,
    annual_fee_pct: float,
    buy_date: Optional[pd.Timestamp] = None,  # FIXED: anchor fee deduction to buy date, not fund inception (Bug 1)
) -> pd.DataFrame:
    # ⚠️ USAGE CORRECT de cette fonction :
    # - Toujours fournir buy_date (date d'achat du client), jamais None
    # - Réservée aux calculs de simulation et de performance client
    #   (XIRR, val_now = qty × last_px_net)
    # - NE PAS utiliser pour afficher une VL de marché absolue
    # - Pour la VL de marché brute : utiliser get_current_nav(isin)
    #   ou get_price_series(isin, None, 0.0)
    # - Pour les calculs de corrélation/vol/Sharpe : utiliser
    #   get_series_for_line(..., apply_fees=False)
    if df.empty or annual_fee_pct == 0:
        return df
    df = df.copy()
    fee_rate = float(annual_fee_pct) / 100.0
    # FIXED: use buy_date as base to avoid applying 20+ years of fictitious fees (Bug 1)
    base_date = pd.Timestamp(buy_date).normalize() if buy_date is not None else df.index[0]
    # FIXED: clamp to zero so pre-buy-date rows (if any) are never penalised (Bug 1)
    day_offsets = np.maximum(0.0, (df.index - base_date).days.astype(float))
    fee_factors = (1.0 - fee_rate) ** (day_offsets / 365.25)
    df["Close"] = df["Close"].astype(float).to_numpy() * fee_factors
    return df


# ------------------------------------------------------------
# Loaders données contrat
# ------------------------------------------------------------

# Chemin racine des données contrats
CONTRACT_DATA_ROOT = "data"


@st.cache_data(show_spinner=False, ttl=3600)
def load_contract_funds(contract_path: str, funds_filename: str) -> pd.DataFrame:
    """
    Charge le fichier Liste des fonds d'un contrat.
    Nettoie les frais (virgule→point, supprime %) et les convertit en float.
    Retourne un DataFrame avec colonnes normalisées :
      isin, name, manager, category, sri, fee_uc_pct, fee_contract_pct, fee_total_pct
    """
    path = os.path.join(CONTRACT_DATA_ROOT, contract_path, funds_filename)
    try:
        df = pd.read_csv(path, sep=";", encoding="utf-8-sig", dtype=str)
    except FileNotFoundError:
        st.error(
            f"Fichier référentiel introuvable : {path}\n"
            f"Vérifiez que le fichier existe bien dans le dépôt GitHub "
            f"avec ce nom exact (attention aux espaces et majuscules)."
        )
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Erreur chargement référentiel : {e}")
        return pd.DataFrame()
    df.columns = df.columns.str.strip()

    def _parse_pct(val: str) -> float:
        try:
            return float(str(val).replace("%", "").replace(",", ".").strip())
        except Exception:
            return 0.0

    result = pd.DataFrame({
        "isin":             df["Code ISIN"].str.strip(),
        "name":             df["Libellé du fonds"].str.strip(),
        "manager":          df["Société de gestion"].str.strip(),
        "category":         df["Catégorie Morningstar"].str.strip(),
        "sri":              pd.to_numeric(df["Risque (SRI)"], errors="coerce").fillna(0).astype(int),
        "fee_uc_pct":       df["Frais UC (B) %"].apply(_parse_pct),
        "fee_contract_pct": df["Frais contrat (C) %"].apply(_parse_pct),
        "fee_total_pct":    df["Frais totaux (B+C) %"].apply(_parse_pct),
    })
    return result.dropna(subset=["isin"]).reset_index(drop=True)


@st.cache_data(show_spinner=False, ttl=3600)
def load_euro_fund_history(contract_path: str, fund_filename: str) -> pd.DataFrame:
    """
    Charge l'historique d'un fonds euros.
    Retourne un DataFrame avec colonnes : annee (int), taux_net_publie_pct (float).
    La colonne utilisée pour la simulation est taux_net_publie_% (net de frais,
    avant prélèvements sociaux — c'est le rendement servi sur le contrat).
    """
    path = os.path.join(CONTRACT_DATA_ROOT, contract_path, fund_filename)
    df = pd.read_csv(path, sep=";", encoding="utf-8-sig")
    df.columns = df.columns.str.strip()
    df["annee"] = pd.to_numeric(df["annee"], errors="coerce").astype("Int64")
    df["taux_net_publie_pct"] = pd.to_numeric(
        df["taux_net_publie_%"], errors="coerce"
    )
    return df[["annee", "taux_net_publie_pct"]].dropna().reset_index(drop=True)


def get_euro_fund_avg_rate(history_df: pd.DataFrame, years: int = 5) -> float:
    """
    Calcule la moyenne du taux net publié sur les N dernières années disponibles.
    Utilisé comme taux par défaut du fonds euros dans la simulation.
    """
    if history_df.empty:
        return 2.0
    recent = history_df.nlargest(years, "annee")
    return round(float(recent["taux_net_publie_pct"].mean()), 2)


def _compute_auto_euro_rate(
    history_df: pd.DataFrame,
    start: pd.Timestamp,
    end: pd.Timestamp,
) -> Optional[float]:
    """
    Calcule le taux moyen du fonds en euros sur les années couvertes par la fenêtre
    [start, end]. Retourne None si l'historique est insuffisant.
    """
    if history_df.empty:
        return None
    start_year = int(start.year)
    end_year = int(end.year)
    mask = (history_df["annee"] >= start_year) & (history_df["annee"] <= end_year)
    filtered = history_df[mask]
    if filtered.empty:
        return None
    return round(float(filtered["taux_net_publie_pct"].mean()), 2)


@st.cache_data(show_spinner=False, ttl=3600)
@st.cache_data(show_spinner=False, ttl=86400)
def _mstarpy_nav_series(isin: str) -> pd.DataFrame:
    """
    Tente de récupérer l'historique VL via mstarpy (fallback EODHD).
    Retourne un DataFrame avec colonne 'Close' indexé par DatetimeIndex.
    Retourne pd.DataFrame() si échec ou mstarpy non disponible.
    TTL 24h — la VL est publiée quotidiennement.
    """
    if not MSTARPY_AVAILABLE:
        return pd.DataFrame()
    try:
        fund = mstarpy.Funds(term=isin, pageSize=1)
        if not fund.isin:
            return pd.DataFrame()
        nav_data = fund.nav(start_date="2005-01-01", frequency="daily")
        if not nav_data:
            return pd.DataFrame()
        df = pd.DataFrame(nav_data)
        if "date" not in df.columns or df.empty:
            return pd.DataFrame()
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df = df.dropna(subset=["date"])
        df = df.set_index("date")
        # mstarpy retourne "nav" ou "close" selon la version
        nav_col = next((c for c in df.columns if c.lower() in ("nav", "close")), None)
        if nav_col is None:
            return pd.DataFrame()
        df = df.rename(columns={nav_col: "Close"})
        df["Close"] = pd.to_numeric(df["Close"], errors="coerce")
        return df[["Close"]].dropna().sort_index()
    except Exception:
        return pd.DataFrame()


def get_price_series(
    isin_or_name: str, start: Optional[pd.Timestamp], euro_rate: float
) -> Tuple[pd.DataFrame, str, str]:
    """
    Retourne la série de prix BRUTE (sans frais appliqués).
    • EUROFUND : start et euro_rate comptent dans la clé de cache (ils définissent la série).
    • Instruments EODHD : toujours appeler avec start=None pour maximiser le taux de cache ;
      le filtrage et l'application des frais sont délégués à get_price_series_with_fees().
    """
    # FIXED: fees and start-filtering for EODHD removed from cached fn (Correction 1 — cache explosion)
    debug = {"cands": []}
    val = str(isin_or_name).strip()
    if not val:
        return pd.DataFrame(), "", json.dumps(debug)

    # ✅ Fonds en euros — capitalisation annualisée (jours calendaires)
    if val.upper() == "EUROFUND":
        start_dt = (
            pd.Timestamp(start).normalize()
            if start is not None
            else pd.Timestamp("2000-01-03")
        )
        start_dt = max(start_dt, pd.Timestamp("2000-01-03"))

        idx = pd.bdate_range(start=start_dt, end=TODAY, freq="B")
        if len(idx) == 0:
            return pd.DataFrame(), "", "{}"

        # FIXED: vectorise with NumPy (Bug 8) — fees applied in get_price_series_with_fees()
        days = (idx - idx[0]).days.values.astype(float)
        r = float(euro_rate) / 100.0
        close_values = (1.0 + r) ** (days / 365.25)
        df = pd.DataFrame({"Close": close_values}, index=idx)
        return df, "EUROFUND", "{}"

    # Source prioritaire : VL importée manuellement (prend le dessus sur EODHD et mstarpy)
    # get_price_series() n'est pas décorée @st.cache_data — accès direct à st.session_state autorisé
    _manual_store = st.session_state.get("MANUAL_NAV_STORE", {})
    if val in _manual_store:
        return _manual_store[val], val, json.dumps({"source": "manual_upload"})

    # ✅ Instruments EODHD — historique complet, sans filtrage ni frais
    # Passer start=None depuis les appelants pour qu'un seul enregistrement de cache
    # serve toutes les lignes du même fonds quelle que soit leur date d'achat.
    cands = _symbol_candidates(val)
    debug["cands"] = cands

    for sym in cands:
        df = eodhd_prices_daily(sym)
        if not df.empty:
            return df, sym, json.dumps(debug)

    # Fallback mstarpy si EODHD n'a rien retourné
    # Uniquement pour les ISIN valides (12 caractères, 2 premières lettres alpha)
    _is_isin = len(val) == 12 and val[:2].isalpha()
    if _is_isin:
        _mstar_df = _mstarpy_nav_series(val)
        if not _mstar_df.empty:
            return _mstar_df, val, json.dumps({"source": "mstarpy_nav"})

    return pd.DataFrame(), "", json.dumps(debug)


def get_price_series_with_fees(
    isin_or_name: str,
    start: Optional[pd.Timestamp],
    euro_rate: float,
    annual_uc_fee_pct: float = ANNUAL_FEE_UC_PCT,  # frais UC de la ligne (contract-aware)
) -> Tuple[pd.DataFrame, str, str]:
    """
    Wrapper non-caché : appelle get_price_series() (start=None pour les instruments EODHD
    afin de maximiser le taux de cache) puis applique les frais annuels ancrés sur la
    date d'achat.  La fonction cachée est ainsi partagée entre toutes les lignes d'un
    même fonds indépendamment de leur date d'achat individuelle.

    ⚠️ USAGE CORRECT de cette fonction :
    - Toujours fournir start = date d'achat du client (jamais None)
    - Réservée aux calculs de simulation et de performance client
    - NE PAS utiliser pour afficher une VL de marché absolue
    - Pour la VL de marché brute : utiliser get_current_nav(isin)
    - Pour les calculs de corrélation/vol/Sharpe : utiliser
      get_series_for_line(..., apply_fees=False)
    """
    # FIXED: fee application and start-trimming moved outside the cache (Correction 1)
    val = str(isin_or_name).strip()

    if val.upper() == "EUROFUND":
        # EUROFUND : euro_rate est déjà net de tous frais de gestion assureur
        # → pas d'apply_annual_fee pour ne pas doubler-compter les frais
        df, sym, dbg = get_price_series(isin_or_name, start, euro_rate)
        return df, sym, dbg

    # Instrument EODHD : start=None → un seul enregistrement de cache par ISIN
    df, sym, dbg = get_price_series(isin_or_name, None, euro_rate)
    if df.empty:
        return df, sym, dbg
    if start is not None:
        df = df.loc[df.index >= pd.Timestamp(start)]
    if not df.empty:
        df = apply_annual_fee(df, annual_uc_fee_pct, buy_date=start)
    return df, sym, dbg


def get_current_nav(
    isin_or_name: str,
) -> Tuple[float, Optional[pd.Timestamp]]:
    """
    Retourne (dernière_VL_brute, date_de_cette_VL).
    Retourne (np.nan, None) si indisponible.
    Ne jamais appeler apply_annual_fee() ici.
    """
    val = str(isin_or_name).strip().upper()
    if val in ("EUROFUND", "STRUCTURED", ""):
        return np.nan, None
    df, _, _ = get_price_series(isin_or_name, None, 0.0)
    if df.empty:
        return np.nan, None
    last_date = df.index[-1]
    return float(df["Close"].iloc[-1]), pd.Timestamp(last_date)


@st.cache_data(show_spinner=False, ttl=3600)
def structured_series(
    start: pd.Timestamp,
    end: pd.Timestamp,
    annual_rate_pct: float,
    redemption_years: int,
) -> pd.DataFrame:
    """
    Série synthétique autocall (simplifiée) :
    - Prix d'achat = 1.0
    - Plat jusqu'à la date de remboursement estimée
    - Saut à 1 + (rate * years) le jour de remboursement, puis figé
    """
    start_dt = pd.Timestamp(start).normalize()
    end_dt = pd.Timestamp(end).normalize()
    idx = pd.bdate_range(start=start_dt, end=end_dt, freq="B")
    if len(idx) == 0:
        return pd.DataFrame()

    df = pd.DataFrame(index=idx, columns=["Close"], dtype=float)
    df.iloc[0, 0] = 1.0

    r = float(annual_rate_pct) / 100.0
    yrs = int(redemption_years)

    redemption_dt = start_dt + pd.DateOffset(years=yrs)

    # série plate + saut à partir du 1er jour ouvré >= redemption_dt
    redeemed = False
    for i in range(1, len(df)):
        d = df.index[i]
        df.iloc[i, 0] = df.iloc[i - 1, 0]

        if (not redeemed) and (d >= redemption_dt):
            df.iloc[i, 0] = 1.0 + r * yrs
            df.iloc[i:, 0] = df.iloc[i, 0]
            redeemed = True
            break

    # sécurité : propagation si besoin
    for i in range(1, len(df)):
        if pd.isna(df.iloc[i, 0]):
            df.iloc[i, 0] = df.iloc[i - 1, 0]

    return df


def _warn_once(key: str, msg: str) -> None:
    seen = st.session_state.get("WARN_ONCE")
    if not isinstance(seen, set):
        if isinstance(seen, list):
            seen = set(seen)
        else:
            seen = set()
    if key in seen:
        return
    seen.add(key)
    st.session_state["WARN_ONCE"] = seen
    st.warning(msg)


def _safe_struct_params(line: Dict[str, Any]) -> Tuple[float, int]:
    try:
        years = int(line.get("struct_years", 6))
    except Exception:
        years = 6
    years = max(1, years)
    try:
        rate = float(line.get("struct_rate", 8.0))
    except Exception:
        rate = 0.0
    rate = max(0.0, min(rate, 25.0))
    return rate, years


def get_series_for_line(
    line: Dict[str, Any],
    start: Optional[pd.Timestamp],
    euro_rate: float,
    apply_fees: bool = True,  # FIXED: route to fee-aware wrapper or raw cache fn (Correction 1)
) -> Tuple[pd.DataFrame, str]:
    isin_or_name = str(line.get("isin") or line.get("name") or "").strip()
    sym_upper = isin_or_name.upper()

    if sym_upper == "EUROFUND":
        if apply_fees:
            df, _, _ = get_price_series_with_fees("EUROFUND", start, euro_rate)
        else:
            df, _, _ = get_price_series("EUROFUND", start, euro_rate)
        if df.empty:
            _warn_once("series_empty:EUROFUND", "Série indisponible pour le fonds en euros (EUROFUND).")
        return df, "EUROFUND"

    if sym_upper == "STRUCTURED":
        buy_ts = pd.Timestamp(line.get("buy_date") or start or TODAY).normalize()
        rate, years = _safe_struct_params(line)
        df = structured_series(
            start=buy_ts,
            end=TODAY,
            annual_rate_pct=rate,
            redemption_years=years,
        )
        if start is not None:
            df = df.loc[df.index >= pd.Timestamp(start)]
        if df.empty:
            label = line.get("name") or "Produit structuré"
            _warn_once(
                f"series_empty:STRUCTURED:{label}",
                f"Série indisponible pour {label} (STRUCTURED) : ligne ignorée.",
            )
        return df, "STRUCTURED"

    # Appliquer uniquement fee_contract_pct (C) — la VL EODHD est déjà
    # nette de fee_uc_pct (B/TER). Utiliser fee_total_pct uniquement en
    # fallback si fee_contract_pct est absent (ligne sans référentiel contrat).
    _raw_fee_gsl = line.get("fee_contract_pct")
    if _raw_fee_gsl is None or _raw_fee_gsl == "":
        _raw_fee_gsl = line.get("fee_total_pct")
        if _raw_fee_gsl is None or _raw_fee_gsl == "":
            fee_uc = ANNUAL_FEE_UC_PCT
        else:
            try:
                fee_uc = float(_raw_fee_gsl)
            except (TypeError, ValueError):
                fee_uc = ANNUAL_FEE_UC_PCT
    else:
        try:
            fee_uc = float(_raw_fee_gsl)
        except (TypeError, ValueError):
            fee_uc = ANNUAL_FEE_UC_PCT
    # fee_uc == 0.0 explicite → pas de frais appliqués (ETF sans frais contrat)

    if apply_fees:
        df, sym, _ = get_price_series_with_fees(isin_or_name, start, euro_rate,
                                                 annual_uc_fee_pct=fee_uc)
    else:
        df, sym, _ = get_price_series(isin_or_name, None, euro_rate)
        if start is not None and not df.empty:
            df = df.loc[df.index >= pd.Timestamp(start)]
    if df.empty:
        label = line.get("name") or isin_or_name or "Ligne"
        _warn_once(
            f"series_empty:{isin_or_name}",
            f"Série indisponible pour {label} ({isin_or_name}) : ligne ignorée.",
        )
    return df, sym or isin_or_name

# ------------------------------------------------------------
# Alternatives si date < 1ère VL
# ------------------------------------------------------------

def suggest_alternative_funds(buy_date: pd.Timestamp, euro_rate: float) -> List[Tuple[str, str, pd.Timestamp]]:
    """
    Propose des fonds recommandés (core + défensifs) dont la première VL
    est antérieure ou égale à la date d'achat donnée.
    Retourne (nom, isin, date_inception).
    """
    alternatives: List[Tuple[str, str, pd.Timestamp]] = []
    universe = RECO_FUNDS_CORE + RECO_FUNDS_DEF

    for name, isin in universe:
        df, _, _ = get_price_series(isin, None, euro_rate)
        if df.empty:
            continue
        inception = df.index.min()
        if inception <= buy_date:
            alternatives.append((name, isin, inception))

    return alternatives


# ------------------------------------------------------------
# Calendrier versements & poids
# ------------------------------------------------------------

def _month_schedule(d0: pd.Timestamp, d1: pd.Timestamp) -> List[pd.Timestamp]:
    if d0 > d1:
        return []
    out = []
    cur = pd.Timestamp(d0.year, d0.month, 1)
    stop = pd.Timestamp(d1.year, d1.month, 1)
    while cur <= stop:
        bdays = pd.bdate_range(start=cur, end=cur + pd.offsets.MonthEnd(0))
        if len(bdays) > 0:
            out.append(bdays[0])
        cur = cur + pd.offsets.MonthBegin(1)
    return out


def _weights_for(
    lines: List[Dict[str, Any]],
    alloc_mode: str,
    custom_weights: Dict[Any, float],
    single_target: Optional[Any],
) -> Dict[Any, float]:
    # FIXED: use stable UUID key (ln.get("id")) with id(ln) fallback for old lines (Bug 5)
    keys = [ln.get("id") or id(ln) for ln in lines]
    if len(keys) == 0:
        return {}
    if alloc_mode == "equal":
        w = 1.0 / len(keys)
        return {k: w for k in keys}
    if alloc_mode == "custom":
        tot = sum(max(0.0, float(custom_weights.get(ln.get("id") or id(ln), 0.0))) for ln in lines)
        if tot <= 0:
            w = 1.0 / len(keys)
            return {k: w for k in keys}
        return {ln.get("id") or id(ln): max(0.0, float(custom_weights.get(ln.get("id") or id(ln), 0.0))) / tot for ln in lines}
    if alloc_mode == "single":
        target = single_target
        return {ln.get("id") or id(ln): (1.0 if (ln.get("id") or id(ln)) == target else 0.0) for ln in lines}
    w = 1.0 / len(keys)
    return {k: w for k in keys}


# ------------------------------------------------------------
# Calcul par ligne (avec frais)
# ------------------------------------------------------------

def compute_line_metrics(
    line: Dict[str, Any], fee_pct: float, euro_rate: float
) -> Tuple[float, float, float]:
    amt_brut = float(line.get("amount_gross", 0.0))
    net_amt = amt_brut * (1.0 - fee_pct / 100.0)
    buy_ts = pd.Timestamp(line.get("buy_date"))
    px_manual = line.get("buy_px", None)

    dfp, sym_used = get_series_for_line(line, buy_ts, euro_rate)
    if dfp.empty:
        return float(net_amt), np.nan, 0.0

    sym_upper = str(sym_used or line.get("isin") or "").upper()
    if sym_upper == "EUROFUND":
        px = _get_close_on(dfp, buy_ts)
    else:
        if px_manual not in (None, "", 0, "0"):
            try:
                px = float(px_manual)
            except Exception:
                px = _get_close_on(dfp, buy_ts)
        else:
            px = _get_close_on(dfp, buy_ts)

    qty = (net_amt / px) if px and px > 0 else 0.0
    return float(net_amt), float(px), float(qty)


# ------------------------------------------------------------
# Simulation d'un portefeuille (avec contrôle 1ère VL)
# + distinction poids mensuels / ponctuels
# ------------------------------------------------------------

def simulate_portfolio(
    lines: List[Dict[str, Any]],
    monthly_amt_gross: float,
    one_amt_gross: float,
    one_date: date,
    alloc_mode: str,
    custom_weights_monthly: Optional[Dict[int, float]],
    custom_weights_oneoff: Optional[Dict[int, float]],
    single_target: Optional[int],
    euro_rate: float,
    fee_pct: float,
    portfolio_label: str = "",
) -> Tuple[pd.DataFrame, float, float, float, Optional[float], pd.Timestamp, pd.Timestamp]:
    if not lines:
        return pd.DataFrame(), 0.0, 0.0, 0.0, None, TODAY, TODAY

    price_map: Dict[Any, pd.Series] = {}
    eff_buy_date: Dict[Any, pd.Timestamp] = {}
    buy_price_used: Dict[Any, float] = {}

    invalid_found = False
    date_warnings = st.session_state.setdefault("DATE_WARNINGS", [])

    for ln in lines:
        key_id = ln.get("id") or id(ln)  # FIXED: stable UUID key, avoids rerun ID changes (Bug 5)
        # FIXED (Bug B): lire la série BRUTE (sans frais) pour déterminer l'inception
        # et vérifier la date d'achat ; les frais seront appliqués ci-dessous
        # ancrés sur d_buy et non sur l'inception du fonds.
        df_full, sym = get_series_for_line(ln, None, euro_rate, apply_fees=False)

        # Sécurité
        if df_full.empty:
            continue

        inception = df_full.index.min()
        d_buy = pd.Timestamp(ln["buy_date"])

        if d_buy < inception:
            invalid_found = True
            ln["invalid_date"] = True
            ln["inception_date"] = inception

            alts = suggest_alternative_funds(d_buy, euro_rate)
            if alts:
                alt_lines = [
                    f"- {name} ({isin}), historique depuis le {fmt_date(incep)}"
                    for name, isin, incep in alts
                ]
                alt_msg = "\n".join(alt_lines)
            else:
                alt_msg = "Aucun fonds recommandé ne dispose d'un historique suffisant pour cette date."

            date_warnings.append(
                f"[{portfolio_label}] {ln.get('name','(sans nom)')} "
                f"({ln.get('isin','—')}) :\n"
                f"- Date d'achat saisie : {fmt_date(d_buy)}\n"
                f"- 1ère VL disponible : {fmt_date(inception)}\n\n"
                f"Impossible de simuler ce fonds sur toute la période demandée.\n"
                f"Propositions d'alternatives pour l'analyse historique :\n{alt_msg}"
            )
            continue

        # FIXED (Bug B): appliquer les frais ancrés sur d_buy (date client réelle)
        # FIXED (Bug E): distinguer fee_total_pct absent (fallback) de 0.0 explicite
        # Appliquer uniquement fee_contract_pct (C) — VL EODHD déjà nette du TER.
        # Fallback fee_total_pct si fee_contract_pct absent, puis ANNUAL_FEE_UC_PCT.
        _isin_sim = str(ln.get("isin", "")).upper()
        _raw_fee_sim = ln.get("fee_contract_pct")
        if _raw_fee_sim is None or _raw_fee_sim == "":
            _raw_fee_sim = ln.get("fee_total_pct")
            if _raw_fee_sim is None or _raw_fee_sim == "":
                _fee_sim = ANNUAL_FEE_UC_PCT if _isin_sim not in ("EUROFUND", "STRUCTURED") else 0.0
            else:
                try:
                    _fee_sim = float(_raw_fee_sim)
                except (TypeError, ValueError):
                    _fee_sim = ANNUAL_FEE_UC_PCT if _isin_sim not in ("EUROFUND", "STRUCTURED") else 0.0
        else:
            try:
                _fee_sim = float(_raw_fee_sim)
            except (TypeError, ValueError):
                _fee_sim = ANNUAL_FEE_UC_PCT if _isin_sim not in ("EUROFUND", "STRUCTURED") else 0.0
        # EUROFUND et STRUCTURED restent à 0.0 (inchangé)
        if _fee_sim > 0 and _isin_sim not in ("EUROFUND", "STRUCTURED"):
            df = apply_annual_fee(df_full.copy(), _fee_sim, buy_date=d_buy)
        else:
            df = df_full

        ln["sym_used"] = sym

        if d_buy in df.index:
            px_buy = float(df.loc[d_buy, "Close"])
            eff_dt = d_buy
        else:
            after = df.loc[df.index >= d_buy]
            if after.empty:
                px_buy = float(df.iloc[-1]["Close"])
                eff_dt = df.index[-1]
            else:
                px_buy = float(after.iloc[0]["Close"])
                eff_dt = after.index[0]

        px_manual = ln.get("buy_px", None)
        px_for_qty = float(px_manual) if (px_manual not in (None, "", 0, "0")) else px_buy

        price_map[key_id] = df["Close"].astype(float)
        eff_buy_date[key_id] = eff_dt
        buy_price_used[key_id] = px_for_qty

    if invalid_found and not price_map:
        return pd.DataFrame(), 0.0, 0.0, 0.0, None, TODAY, TODAY
    if not price_map:
        return pd.DataFrame(), 0.0, 0.0, 0.0, None, TODAY, TODAY

    start_min = min(eff_buy_date.values())
    start_full = max(eff_buy_date.values())

    bidx = pd.bdate_range(start=start_min, end=TODAY, freq="B")
    prices = pd.DataFrame(index=bidx)
    for key_id, s in price_map.items():
        prices[key_id] = s.reindex(bidx).ffill()

    qty_events = pd.DataFrame(0.0, index=bidx, columns=prices.columns)
    total_brut = 0.0
    total_net = 0.0
    cash_flows: List[Tuple[pd.Timestamp, float]] = []

    # Achats initiaux
    for ln in lines:
        key_id = ln.get("id") or id(ln)  # FIXED: stable UUID key (Bug 5)
        if key_id not in prices.columns:
            continue
        brut = float(ln.get("amount_gross", 0.0))
        net = brut * (1.0 - fee_pct / 100.0)
        px = float(buy_price_used[key_id])
        dt = eff_buy_date[key_id]
        if brut > 0 and px > 0:
            q = net / px
            tgt = dt if dt in qty_events.index else qty_events.index[qty_events.index >= dt][0]
            qty_events.loc[tgt, key_id] += q
            total_brut += brut
            total_net += net
            cash_flows.append((tgt, -brut))

    # Poids pour versements mensuels / ponctuels
    weights_monthly = _weights_for(
        lines,
        alloc_mode,
        custom_weights_monthly or {},
        single_target,
    )
    weights_oneoff = _weights_for(
        lines,
        alloc_mode,
        custom_weights_oneoff or {},
        single_target,
    )

    # Versement ponctuel
    if one_amt_gross > 0:
        dt = pd.Timestamp(one_date)
        if dt not in qty_events.index:
            after = qty_events.index[qty_events.index >= dt]
            if len(after) > 0:
                dt = after[0]
            else:
                dt = None
        if dt is not None:
            net_amt = one_amt_gross * (1.0 - fee_pct / 100.0)
            for ln in lines:
                key_id = ln.get("id") or id(ln)  # FIXED: stable UUID key (Bug 5)
                w = weights_oneoff.get(key_id, 0.0)
                if w <= 0 or key_id not in prices.columns:
                    continue
                px = float(prices.loc[dt, key_id])
                if px > 0:
                    qty_events.loc[dt, key_id] += (net_amt * w) / px
            total_brut += float(one_amt_gross)
            total_net += float(net_amt)
            cash_flows.append((dt, -float(one_amt_gross)))

    # Mensuels
    if monthly_amt_gross > 0:
        sched = _month_schedule(start_min, TODAY)
        for dt in sched:
            if dt not in qty_events.index:
                after = qty_events.index[qty_events.index >= dt]
                if len(after) == 0:
                    continue
                dt = after[0]
            net_m = monthly_amt_gross * (1.0 - fee_pct / 100.0)
            for ln in lines:
                key_id = ln.get("id") or id(ln)  # FIXED: stable UUID key (Bug 5)
                w = weights_monthly.get(key_id, 0.0)
                if w <= 0 or key_id not in prices.columns:
                    continue
                px = float(prices.loc[dt, key_id])
                if px > 0:
                    qty_events.loc[dt, key_id] += (net_m * w) / px
            total_brut += float(monthly_amt_gross)
            total_net += float(net_m)
            cash_flows.append((dt, -float(monthly_amt_gross)))

    qty_cum = qty_events.cumsum()
    values = (qty_cum * prices).sum(axis=1)
    df_val = pd.DataFrame({"Valeur": values})
    final_val = float(df_val["Valeur"].iloc[-1]) if not df_val.empty else 0.0

    cash_flows.append((TODAY, final_val))
    irr = xirr(cash_flows)

    return df_val, total_brut, total_net, final_val, (irr * 100.0 if irr is not None else None), start_min, start_full


@st.cache_data(show_spinner=False, ttl=3600)
def _simulate_portfolio_cached(
    lines_json: str,
    monthly_amt_gross: float,
    one_amt_gross: float,
    one_date_str: str,
    alloc_mode: str,
    custom_weights_monthly_json: str,
    custom_weights_oneoff_json: str,
    single_target: Optional[int],
    euro_rate: float,
    fee_pct: float,
    portfolio_label: str = "",
) -> Tuple[pd.DataFrame, float, float, float, Optional[float], pd.Timestamp, pd.Timestamp]:
    """Wrapper caché de simulate_portfolio. Tous les paramètres sont hashables.
    La clé de cache est le hash de la sérialisation JSON des lignes + scalaires.
    Si le portefeuille n'a pas changé entre deux reruns, résultat instantané."""
    lines = json.loads(lines_json)
    custom_weights_monthly = json.loads(custom_weights_monthly_json)
    custom_weights_oneoff = json.loads(custom_weights_oneoff_json)
    one_date = pd.Timestamp(one_date_str).date()
    return simulate_portfolio(
        lines=lines,
        monthly_amt_gross=monthly_amt_gross,
        one_amt_gross=one_amt_gross,
        one_date=one_date,
        alloc_mode=alloc_mode,
        custom_weights_monthly=custom_weights_monthly or None,
        custom_weights_oneoff=custom_weights_oneoff or None,
        single_target=single_target,
        euro_rate=euro_rate,
        fee_pct=fee_pct,
        portfolio_label=portfolio_label,
    )


def _make_sim_args(
    lines: List[Dict[str, Any]],
    monthly_amt: float,
    one_amt: float,
    one_date,
    alloc_mode: str,
    custom_weights_monthly: Optional[Dict],
    custom_weights_oneoff: Optional[Dict],
    single_target: Optional[int],
    euro_rate: float,
    fee_pct: float,
    label: str,
) -> dict:
    """Sérialise les arguments pour _simulate_portfolio_cached.
    Convertit les clés en str pour la sérialisation JSON."""
    def _str_keys(d: Optional[Dict]) -> Dict[str, float]:
        if not d:
            return {}
        return {str(k): v for k, v in d.items()}
    return dict(
        lines_json=json.dumps(lines, default=str, sort_keys=True),
        monthly_amt_gross=float(monthly_amt),
        one_amt_gross=float(one_amt),
        one_date_str=str(pd.Timestamp(one_date).date()),
        alloc_mode=str(alloc_mode),
        custom_weights_monthly_json=json.dumps(_str_keys(custom_weights_monthly)),
        custom_weights_oneoff_json=json.dumps(_str_keys(custom_weights_oneoff)),
        single_target=single_target,
        euro_rate=float(euro_rate),
        fee_pct=float(fee_pct),
        portfolio_label=str(label),
    )


# ------------------------------------------------------------
# Cartes lignes (édition / suppression)
# ------------------------------------------------------------

def _line_card(line: Dict[str, Any], idx: int, port_key: str):
    # FIXED: use stable UUID-based key so widget state survives st.rerun() (Bug 5)
    _card_id = line.get("id", str(idx))
    state_key = f"edit_mode_{port_key}_{_card_id}"
    if state_key not in st.session_state:
        st.session_state[state_key] = False

    fee_pct = st.session_state.get("FEE_A", 0.0) if port_key == "A_lines" else st.session_state.get("FEE_B", 0.0)
    # FIXED: use per-portfolio euro rate instead of shared EURO_RATE_PREVIEW (Résidu Bug 4)
    euro_rate = (
        st.session_state.get("EURO_RATE_A", 2.0)
        if port_key == "A_lines"
        else st.session_state.get("EURO_RATE_B", 2.5)
    )
    net_amt, buy_px, qty_disp = compute_line_metrics(line, fee_pct, euro_rate)

    with st.container(border=True):
        cols = st.columns([3, 2, 2, 2, 1])
        with cols[0]:
            st.markdown(f"**{line.get('name','—')}**")
            st.caption(f"ISIN / Code : `{line.get('isin','—')}`")
            st.caption(f"Symbole EODHD : `{line.get('sym_used','—')}`")
            if line.get("invalid_date"):
                st.markdown(
                    f"⚠️ Date d'achat antérieure à la 1ère VL ({fmt_date(line.get('inception_date'))}).",
                )
        with cols[1]:
            st.markdown(f"Investi (brut)\n\n**{to_eur(line.get('amount_gross', 0.0))}**")
            st.caption(f"Net après frais {fee_pct:.1f}% : **{to_eur(net_amt)}**")
            st.caption(f"Date d'achat : {fmt_date(line.get('buy_date'))}")
            if line.get("date_overridden"):
                st.caption("📌 Date individuelle (non liée à la date globale)")
            _buy_ts_check = pd.Timestamp(line.get("buy_date", TODAY))
            if _buy_ts_check > pd.Timestamp(TODAY):
                st.caption("⚠️ Date d'achat dans le futur — simulation non disponible.")
        with cols[2]:
            st.markdown(f"VL d'achat\n\n**{to_eur(buy_px)}**")
            st.caption(f"Quantité : {qty_disp:.6f}")
            if line.get("note"):
                st.caption(line["note"])
        with cols[3]:
            try:
                # FIXED (Bug A + P2): VL brute sans frais + date de cotation
                last, nav_date = get_current_nav(
                    line.get("isin") or line.get("name") or ""
                )
                if last == last:  # not nan
                    date_str = fmt_date(nav_date) if nav_date is not None else "—"
                    st.markdown(f"Dernière VL connue ({date_str}) : **{to_eur(last)}**")
                    st.caption(
                        "VL publiée par EODHD (clôture J-1 à J-3 selon le fonds)"
                    )
                else:
                    st.markdown("Dernière VL connue : —")
            except Exception:
                st.markdown("Dernière VL connue : —")
        with cols[4]:
            if not st.session_state[state_key]:
                # FIXED: use stable UUID-based widget key (Bug 5)
                if st.button("✏️", key=f"edit_{port_key}_{_card_id}", help="Modifier"):
                    st.session_state[state_key] = True
                    st.rerun()  # FIXED: st.experimental_rerun() deprecated since 1.27 (Bug 6)
            # FIXED: use stable UUID-based widget key (Bug 5)
            if st.button("🗑️", key=f"del_{port_key}_{_card_id}", help="Supprimer"):
                st.session_state[port_key].pop(idx)
                st.rerun()  # FIXED: st.experimental_rerun() deprecated since 1.27 (Bug 6)

        if st.session_state[state_key]:
            # FIXED: use stable UUID-based form key (Bug 5)
            with st.form(key=f"form_edit_{port_key}_{_card_id}", clear_on_submit=False):
                c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
                with c1:
                    new_amount = st.text_input("Montant investi (brut) €", value=str(line.get("amount_gross", "")))
                with c2:
                    new_date = st.date_input("Date d’achat", value=pd.Timestamp(line.get("buy_date")).date())
                with c3:
                    new_px = st.text_input("Prix d’achat (optionnel)", value=str(line.get("buy_px", "")))
                with c4:
                    st.caption(" ")
                    submitted = st.form_submit_button("💾 Enregistrer")
                if submitted:
                    try:
                        amt_gross = float(str(new_amount).replace(" ", "").replace(",", "."))
                        assert amt_gross > 0
                    except Exception:
                        st.warning("Montant brut invalide.")
                        st.stop()
                    buy_ts = pd.Timestamp(new_date)
                    line["amount_gross"] = float(amt_gross)
                    line["buy_date"] = buy_ts
                    line["date_overridden"] = True
                    if new_px.strip():
                        try:
                            line["buy_px"] = float(str(new_px).replace(",", "."))
                        except Exception:
                            line["buy_px"] = ""
                    else:
                        line["buy_px"] = ""
                    line.pop("invalid_date", None)
                    line.pop("inception_date", None)
                    st.session_state[state_key] = False
                    st.success("Ligne mise à jour.")
                    st.rerun()  # FIXED: st.experimental_rerun() deprecated since 1.27 (Bug 6)

    # Import VL manuelle si la série est vide (EODHD + mstarpy ont échoué)
    _isin_card = str(line.get("isin") or "").strip()
    _has_data = not get_series_for_line(
        line, pd.Timestamp(line.get("buy_date", TODAY)), euro_rate, apply_fees=False
    )[0].empty
    _has_manual = _isin_card in st.session_state.get("MANUAL_NAV_STORE", {})
    if not _has_data and not _has_manual and _isin_card not in ("EUROFUND", "STRUCTURED", ""):
        with st.expander(
            f"📥 Importer l'historique VL — {line.get('name', _isin_card)}", expanded=False
        ):
            st.caption(
                "Ce fonds n'est pas disponible via EODHD ni Morningstar. "
                "Importez son historique de VL au format CSV (colonnes : date ; vl). "
                "Sources recommandées : OPCVM360, Fundkis, Morningstar Direct."
            )
            _uploaded_nav = st.file_uploader(
                "CSV historique VL (séparateur ; décimale ,)",
                type=["csv"],
                key=f"nav_upload_{_isin_card}_{_card_id}",
            )
            if _uploaded_nav is not None:
                try:
                    _nav_raw = pd.read_csv(
                        _uploaded_nav, sep=";", decimal=",",
                        encoding="utf-8-sig", dtype=str,
                    )
                    _nav_raw.columns = _nav_raw.columns.str.strip().str.lower()
                    _vl_col = next(
                        (c for c in _nav_raw.columns
                         if c in ("vl", "nav", "close", "valeur liquidative", "price")),
                        None,
                    )
                    if _vl_col and "date" in _nav_raw.columns:
                        _nav_raw["date"] = pd.to_datetime(
                            _nav_raw["date"], dayfirst=True, errors="coerce"
                        )
                        _nav_raw[_vl_col] = pd.to_numeric(
                            _nav_raw[_vl_col].str.replace(",", "."), errors="coerce"
                        )
                        _nav_raw = _nav_raw.dropna(subset=["date", _vl_col])
                        _nav_df = (
                            _nav_raw.set_index("date")
                            .rename(columns={_vl_col: "Close"})[["Close"]]
                            .sort_index()
                        )
                        if not _nav_df.empty:
                            _store = st.session_state.get("MANUAL_NAV_STORE", {})
                            _store[_isin_card] = _nav_df
                            st.session_state["MANUAL_NAV_STORE"] = _store
                            st.success(
                                f"✅ {len(_nav_df)} VL importées pour "
                                f"{line.get('name', _isin_card)}. Relancez le calcul."
                            )
                        else:
                            st.error("Fichier valide mais aucune donnée exploitable.")
                    else:
                        st.error(
                            "Colonnes attendues : 'date' et 'vl' (ou 'nav' / 'close'). "
                            f"Colonnes détectées : {list(_nav_raw.columns)}"
                        )
                except Exception as _e:
                    st.error(f"Erreur lecture CSV : {_e}")


def build_positions_dataframe(port_key: str) -> pd.DataFrame:
    """
    Construit un DataFrame par ligne :
    Nom, ISIN, Date d'achat, Net investi, Valeur actuelle, Perf € et Perf %.
    """
    fee_pct = (
        st.session_state.get("FEE_A", 0.0)
        if port_key == "A_lines"
        else st.session_state.get("FEE_B", 0.0)
    )

    euro_rate = (
        st.session_state.get("EURO_RATE_A", 2.0)
        if port_key == "A_lines"
        else st.session_state.get("EURO_RATE_B", 2.5)
    )

    lines = st.session_state.get(port_key, [])
    rows: List[Dict[str, Any]] = []

    for ln in lines:
        buy_ts = pd.Timestamp(ln.get("buy_date"))
        net_amt, buy_px, qty = compute_line_metrics(ln, fee_pct, euro_rate)
        dfl, _ = get_series_for_line(ln, buy_ts, euro_rate)
        if dfl.empty:
            continue

        # FIXED: removed dead `if not dfl.empty` block — unreachable after `continue` above (Bug 9)
        last_px = float(dfl["Close"].iloc[-1])

        val_now = qty * last_px if last_px == last_px else 0.0
        perf_abs = val_now - net_amt
        perf_pct = (val_now / net_amt - 1.0) * 100.0 if net_amt > 0 else np.nan

        # FIXED (Bug D + P2): VL marché brute séparée de val_now ; date ignorée ici
        nav_brute, _ = get_current_nav(ln.get("isin") or "")

        rows.append(
            {
                "Nom": ln.get("name", ""),
                "ISIN / Code": ln.get("isin", ""),
                "Date d'achat": fmt_date(ln.get("buy_date")),
                "Date individuelle": "📌" if ln.get("date_overridden") else "",
                "Net investi €": net_amt,
                "Valeur actuelle €": val_now,
                "VL marché": round(nav_brute, 4) if nav_brute == nav_brute else None,
                "Perf €": perf_abs,
                "Perf %": perf_pct,
            }
        )

    return pd.DataFrame(rows)

# ------------------------------------------------------------
# Tableau synthétique par ligne (un seul tableau par portefeuille)
# ------------------------------------------------------------

def positions_table(title: str, port_key: str):
    """
    Affiche un seul tableau synthétique par portefeuille :
    Nom, ISIN, Date d'achat, Net investi, Valeur actuelle, Perf € et Perf %.
    """
    fee_pct = (
        st.session_state.get("FEE_A", 0.0)
        if port_key == "A_lines"
        else st.session_state.get("FEE_B", 0.0)
    )

    # ✅ Taux fonds euros par portefeuille (au lieu de EURO_RATE_PREVIEW)
    euro_rate = (
        st.session_state.get("EURO_RATE_A", 2.0)
        if port_key == "A_lines"
        else st.session_state.get("EURO_RATE_B", 2.5)
    )

    lines = st.session_state.get(port_key, [])
    rows: List[Dict[str, Any]] = []

    for ln in lines:
        buy_ts = pd.Timestamp(ln.get("buy_date"))

        # Montant net investi, VL d'achat et quantité
        net_amt, buy_px, qty = compute_line_metrics(ln, fee_pct, euro_rate)

        # ✅ IMPORTANT : on récupère la série "depuis buy_ts" pour éviter le mismatch EUROFUND
        dfl, _ = get_series_for_line(ln, buy_ts, euro_rate)
        if dfl.empty:
            continue

        # FIXED: removed dead `if not dfl.empty` block — unreachable after `continue` above (Bug 9)
        last_px = float(dfl["Close"].iloc[-1])

        # Valeur actuelle et performance
        val_now = qty * last_px if last_px == last_px else 0.0
        perf_abs = val_now - net_amt
        perf_pct = (val_now / net_amt - 1.0) * 100.0 if net_amt > 0 else np.nan

        # FIXED (Bug D + P2): VL marché brute séparée de val_now ; date ignorée ici
        nav_brute, _ = get_current_nav(ln.get("isin") or "")

        rows.append(
            {
                "Nom": ln.get("name", ""),
                "ISIN / Code": ln.get("isin", ""),
                "Date d'achat": fmt_date(ln.get("buy_date")),
                "Date individuelle": "📌" if ln.get("date_overridden") else "",
                "Net investi €": net_amt,
                "Valeur actuelle €": val_now,
                "VL marché": round(nav_brute, 4) if nav_brute == nav_brute else None,
                "Perf €": perf_abs,
                "Perf %": perf_pct,
            }
        )

    st.markdown(f"### {title}")
    df = pd.DataFrame(rows)
    if df.empty:
        st.info("Aucune ligne.")
    else:
        st.dataframe(
            df.style.format(
                {
                    "Net investi €": to_eur,
                    "Valeur actuelle €": to_eur,
                    "VL marché": "{:.4f}".format,
                    "Perf €": to_eur,
                    "Perf %": "{:,.2f}%".format,
                }
            ),
            hide_index=True,
            use_container_width=True,
        )
        st.caption("VL marché : dernière cotation brute EODHD (hors frais de gestion)")


def _prepare_pie_df(df_positions: pd.DataFrame, max_items: int = 8, min_pct: float = 0.03) -> pd.DataFrame:
    if df_positions.empty:
        return df_positions
    df = df_positions.copy()
    df = df[df["Valeur actuelle €"] > 0]
    if df.empty:
        return df
    total = df["Valeur actuelle €"].sum()
    df["Part %"] = df["Valeur actuelle €"] / total
    df = df.sort_values("Valeur actuelle €", ascending=False)
    if len(df) > max_items:
        df_main = df.iloc[:max_items].copy()
        df_other = df.iloc[max_items:]
        df_main = pd.concat(
            [
                df_main,
                pd.DataFrame(
                    {
                        "Nom": ["Autres"],
                        "Valeur actuelle €": [df_other["Valeur actuelle €"].sum()],
                        "Part %": [df_other["Valeur actuelle €"].sum() / total],
                    }
                ),
            ],
            ignore_index=True,
        )
        df = df_main
    else:
        small = df[df["Part %"] < min_pct]
        if not small.empty and len(df) > 1:
            df_main = df[df["Part %"] >= min_pct]
            df_other = pd.DataFrame(
                {
                    "Nom": ["Autres"],
                    "Valeur actuelle €": [small["Valeur actuelle €"].sum()],
                    "Part %": [small["Valeur actuelle €"].sum() / total],
                }
            )
            df = pd.concat([df_main, df_other], ignore_index=True)
    df["Part %"] = df["Part %"] * 100.0
    return df


# ------------------------------------------------------------
# Analytics internes : retours, corrélation, volatilité
# ------------------------------------------------------------


def _build_returns_df(
    lines: List[Dict[str, Any]],
    euro_rate: float,
    years: int = 3,
    min_points: int = 60,
) -> pd.DataFrame:
    """
    Construit un DataFrame de rendements journaliers (pct_change)
    pour toutes les lignes du portefeuille avec un historique suffisant.
    Index = dates, colonnes = "Nom (ISIN)".
    """
    cutoff = TODAY - pd.Timedelta(days=365 * years)
    series_map: Dict[str, pd.Series] = {}

    for ln in lines:
        label = (ln.get("name") or ln.get("isin") or "Ligne").strip()
        isin = (ln.get("isin") or "").strip()
        key = f"{label} ({isin})" if isin else label

        # FIXED (Bug F): apply_fees=False — les frais biaisent les rendements si
        # ancrés sur l'inception du fonds ; pct_change() ne nécessite pas de frais nets
        df, _ = get_series_for_line(ln, None, euro_rate, apply_fees=False)
        if df.empty:
            continue

        s = df["Close"].astype(float)
        s = s[s.index >= cutoff]
        if s.size < min_points:
            continue

        series_map[key] = s

    if not series_map:
        return pd.DataFrame()

    df_prices = pd.DataFrame(series_map).dropna(how="any")
    if df_prices.shape[0] < min_points:
        return pd.DataFrame()

    returns = df_prices.pct_change().dropna(how="any")
    return returns



def correlation_matrix_from_lines(
    lines: List[Dict[str, Any]],
    euro_rate: float,
    years: int = 3,
    min_points: int = 60,
) -> pd.DataFrame:
    """
    Matrice de corrélation entre les lignes du portefeuille,
    basée sur les rendements journaliers.
    """
    rets = _build_returns_df(lines, euro_rate, years=years, min_points=min_points)
    if rets.empty:
        return pd.DataFrame()
    return rets.corr()


def volatility_table_from_lines(
    lines: List[Dict[str, Any]],
    euro_rate: float,
    years: int = 3,
    min_points: int = 60,
) -> pd.DataFrame:
    """
    Volatilité annuelle par ligne (et écart-type quotidien).
    """
    rets = _build_returns_df(lines, euro_rate, years=years, min_points=min_points)
    if rets.empty:
        return pd.DataFrame()

    rows = []
    for col in rets.columns:
        r = rets[col]
        daily_std = float(r.std())
        ann_std = daily_std * np.sqrt(252.0)
        rows.append(
            {
                "Ligne": col,
                "Écart-type quotidien %": daily_std * 100.0,
                "Volatilité annuelle %": ann_std * 100.0,
                "Nombre de points": int(r.count()),
            }
        )
    df = pd.DataFrame(rows)
    return df.sort_values("Volatilité annuelle %", ascending=False)


def portfolio_risk_stats(
    lines: List[Dict[str, Any]],
    euro_rate: float,
    years: int = 3,
    min_points: int = 60,
    fee_pct: float = 0.0,  # FIXED: explicit fee rate eliminates max(net_A, net_B) proxy (Bug 3)
) -> Optional[Dict[str, float]]:
    """
    Calcule quelques stats globales de risque pour le portefeuille :
    - volatilité annuelle
    - max drawdown sur la période.
    Pondération par montant net investi.
    """
    rets = _build_returns_df(lines, euro_rate, years=years, min_points=min_points)
    if rets.empty:
        return None

    # FIXED: use caller-supplied fee_pct instead of guessing A vs B from session state (Bug 3)
    net_by_col: Dict[str, float] = {}
    for ln in lines:
        label = (ln.get("name") or ln.get("isin") or "Ligne").strip()
        isin = (ln.get("isin") or "").strip()
        key = f"{label} ({isin})" if isin else label

        net, _, _ = compute_line_metrics(ln, fee_pct, euro_rate)
        if net > 0:
            net_by_col[key] = net

    # normalisation des poids
    weights = {}
    tot = sum(net_by_col.get(c, 0.0) for c in rets.columns)
    if tot <= 0:
        return None
    for c in rets.columns:
        w = net_by_col.get(c, 0.0) / tot
        weights[c] = w

    # série de rendement du portefeuille
    w_vec = np.array([weights[c] for c in rets.columns])
    rp = rets.to_numpy().dot(w_vec)
    rp = pd.Series(rp, index=rets.index)

    daily_std = float(rp.std())
    vol_ann = daily_std * np.sqrt(252.0)

    # max drawdown
    cum = (1.0 + rp).cumprod()
    running_max = cum.cummax()
    dd = cum / running_max - 1.0
    max_dd = float(dd.min())  # négatif

    return {
        "vol_ann_pct": vol_ann * 100.0,
        "max_dd_pct": max_dd * 100.0,
    }


def _corr_heatmap_chart(corr: pd.DataFrame, title: str) -> Optional[alt.Chart]:
    """
    Construit une heatmap Altair pour visualiser la matrice de corrélation.
    """
    if corr.empty or corr.shape[0] < 2:
        return None

    df_corr = corr.copy()
    df_corr["Ligne1"] = df_corr.index
    df_melt = df_corr.melt(id_vars="Ligne1", var_name="Ligne2", value_name="corr")

    base = (
        alt.Chart(df_melt)
        .encode(
            x=alt.X("Ligne1:O", sort=None, title=""),
            y=alt.Y("Ligne2:O", sort=None, title=""),
        )
    )

    heat = base.mark_rect().encode(
        color=alt.Color("corr:Q", scale=alt.Scale(domain=[-1, 1])),
        tooltip=[
            alt.Tooltip("Ligne1:N", title="Ligne 1"),
            alt.Tooltip("Ligne2:N", title="Ligne 2"),
            alt.Tooltip("corr:Q", title="Corrélation", format=".2f"),
        ],
    )

    text = base.mark_text(baseline="middle").encode(
        text=alt.Text("corr:Q", format=".2f"),
    )

    return (heat + text).properties(title=title, height=300)

# ------------------------------------------------------------
# Blocs de saisie : soit fonds recommandés, soit saisie libre
# FIXED: removed 2 duplicate comment blocks (Bug 9)
# ------------------------------------------------------------

def _add_from_reco_block(port_key: str, label: str):
    st.subheader(label)

    cat = st.selectbox(
        "Catégorie",
        ["Core (référence)", "Défensif", "Produits structurés"],
        key=f"reco_cat_{port_key}",
    )

    # ✅ Date d'achat centralisée (versement initial uniquement)
    buy_date = (
        st.session_state.get("INIT_A_DATE", pd.Timestamp("2024-01-02").date())
        if port_key == "A_lines"
        else st.session_state.get("INIT_B_DATE", pd.Timestamp("2024-01-02").date())
    )

    # ============================
    # CAS 1 — PRODUIT STRUCTURÉ
    # ============================
    if cat == "Produits structurés":
        st.markdown("### Produit structuré (Autocall)")

        c1, c2 = st.columns(2)
        with c1:
            amount = st.text_input(
                "Montant investi (brut) €",
                value="",
                key=f"struct_amt_{port_key}",
            )
        with c2:
            struct_years = st.number_input(
                "Durée estimée avant remboursement (années)",
                min_value=1,
                max_value=12,
                value=6,
                step=1,
                key=f"struct_years_{port_key}",
            )

        struct_rate = st.number_input(
            "Rendement annuel estimé (%)",
            min_value=0.0,
            max_value=25.0,
            value=8.0,
            step=0.10,
            key=f"struct_rate_{port_key}",
        )

        st.caption(
            f"Date d’investissement initiale : {pd.Timestamp(buy_date).strftime('%d/%m/%Y')}"
        )

        if st.button("➕ Ajouter le produit structuré", key=f"struct_add_{port_key}"):
            try:
                amt = float(str(amount).replace(" ", "").replace(",", "."))
                assert amt > 0
            except Exception:
                st.warning("Montant invalide.")
                return

            ln = {
                "id": str(uuid.uuid4()),  # FIXED: stable UUID key for line identity (Bug 5)
                "name": f"Produit structuré ({struct_rate:.2f}% / {int(struct_years)} ans)",
                "isin": "STRUCTURED",
                "amount_gross": float(amt),
                "buy_date": pd.Timestamp(buy_date),
                "buy_px": 1.0,
                "struct_rate": float(struct_rate),
                "struct_years": int(struct_years),
                "note": "",
                "sym_used": "STRUCTURED",
            }
            st.session_state[port_key].append(ln)
            st.success("Produit structuré ajouté.")
        return  # ✅ IMPORTANT : on sort de la fonction pour ne pas afficher la partie fonds

    # ============================
    # CAS 2 — FONDS CLASSIQUES
    # ============================
    if cat == "Core (référence)":
        fonds_list = RECO_FUNDS_CORE
    else:
        fonds_list = RECO_FUNDS_DEF

    options = [f"{nm} ({isin})" for nm, isin in fonds_list]
    choice = st.selectbox("Fonds recommandé", options, key=f"reco_choice_{port_key}")
    idx = options.index(choice) if choice in options else 0
    name, isin = fonds_list[idx]

    c1, c2 = st.columns([2, 2])
    with c1:
        amount = st.text_input("Montant investi (brut) €", value="", key=f"reco_amt_{port_key}")
    with c2:
        st.caption(f"Date d’achat (versement initial) : {pd.Timestamp(buy_date).strftime('%d/%m/%Y')}")

    px = st.text_input("Prix d’achat (optionnel)", value="", key=f"reco_px_{port_key}")

    if st.button("➕ Ajouter ce fonds recommandé", key=f"reco_add_{port_key}"):
        try:
            amt = float(str(amount).replace(" ", "").replace(",", "."))
            assert amt > 0
        except Exception:
            st.warning("Montant invalide.")
            return

        ln = {
            "id": str(uuid.uuid4()),  # FIXED: stable UUID key for line identity (Bug 5)
            "name": name,
            "isin": isin,
            "amount_gross": float(amt),
            "buy_date": pd.Timestamp(buy_date),
            "buy_px": float(str(px).replace(",", ".")) if px.strip() else "",
            "note": "",
            "sym_used": "",
        }
        st.session_state[port_key].append(ln)
        st.success("Fonds recommandé ajouté.")


def _add_line_form_free(port_key: str, label: str):
    st.subheader(label)

    # ✅ Date d'achat centralisée (versement initial)
    buy_date_central = (
        st.session_state.get("INIT_A_DATE", pd.Timestamp("2024-01-02").date())
        if port_key == "A_lines"
        else st.session_state.get("INIT_B_DATE", pd.Timestamp("2024-01-02").date())
    )

    with st.form(key=f"form_add_free_{port_key}", clear_on_submit=False):
        c1, c2 = st.columns([3, 2])

        with c1:
            name = st.text_input("Nom du fonds (libre)", value="")
            isin = st.text_input("ISIN ou code (peut être 'EUROFUND')", value="")

        with c2:
            amount = st.text_input("Montant investi (brut) €", value="")
            st.caption(
                f"Date d’achat (versement initial) : "
                f"{pd.Timestamp(buy_date_central).strftime('%d/%m/%Y')}"
            )

        px = st.text_input("Prix d’achat (optionnel)", value="")
        note = st.text_input("Note (optionnel)", value="")
        add_btn = st.form_submit_button("➕ Ajouter cette ligne")

    if not add_btn:
        return

    isin_final = isin.strip()
    name_final = name.strip()

    # Si nom vide mais ISIN renseigné : tentative de récupération du nom
    if not name_final and isin_final:
        res = eodhd_search(isin_final)
        match = None
        for it in res:
            if it.get("ISIN") == isin_final:
                match = it
                break
        if match is None and res:
            match = res[0]
        if match:
            name_final = match.get("Name", isin_final)

    if not name_final and isin_final.upper() == "EUROFUND":
        name_final = "Fonds en euros (EUROFUND)"

    if not name_final:
        name_final = isin_final or "—"

    try:
        amt = float(str(amount).replace(" ", "").replace(",", "."))
        assert amt > 0
    except Exception:
        st.warning("Montant invalide.")
        return

    ln = {
        "id": str(uuid.uuid4()),  # FIXED: stable UUID key for line identity (Bug 5)
        "name": name_final,
        "isin": isin_final or name_final,
        "amount_gross": float(amt),
        "buy_date": pd.Timestamp(buy_date_central),  # ✅ applique la date centrale
        "buy_px": float(str(px).replace(",", ".")) if px.strip() else "",
        "note": note.strip(),
        "sym_used": "",
    }

    st.session_state[port_key].append(ln)
    st.success("Ligne ajoutée.")


@st.cache_data(show_spinner=False, ttl=3600)
def _build_returns_matrix(
    isins: Tuple[str, ...],
    euro_rate: float,
    start: pd.Timestamp,
    end: pd.Timestamp,
    min_points: int = 60,
) -> Tuple[pd.DataFrame, List[str], Dict[str, str]]:
    series_map: Dict[str, pd.Series] = {}
    warnings: List[str] = []
    status: Dict[str, str] = {}

    for isin in isins:
        df, _, _ = get_price_series(isin, None, euro_rate)
        if df.empty or df["Close"].dropna().shape[0] < min_points:
            warnings.append(f"{isin} : historique insuffisant")
            status[isin] = "insufficient"
            continue
        s = df["Close"].astype(float)
        s = s[(s.index >= start) & (s.index <= end)]
        if s.dropna().shape[0] < min_points:
            warnings.append(f"{isin} : historique insuffisant sur la fenêtre")
            status[isin] = "insufficient"
            continue
        series_map[isin] = s
        status[isin] = "ok"

    if not series_map:
        return pd.DataFrame(), warnings, status

    prices = pd.DataFrame(series_map).ffill().dropna(how="any")
    if prices.shape[0] < min_points:
        return pd.DataFrame(), warnings, status
    returns = prices.pct_change().dropna(how="any")
    return returns, warnings, status


def _suggest_weights(
    returns: pd.DataFrame,
    max_weight: float,
    min_funds: int,
) -> Dict[str, float]:
    if returns.empty:
        return {}
    corr = returns.corr()
    vols = returns.std() * np.sqrt(252.0)
    avg_corr = corr.mean()
    ranked = avg_corr.sort_values().index.tolist()
    min_funds = max(1, min(min_funds, len(ranked)))
    selected = ranked[:min_funds]
    inv_vol = 1 / vols[selected]
    weights = (inv_vol / inv_vol.sum()).to_dict()

    # cap weights if needed
    if max_weight > 0 and max_weight < 1:
        for _ in range(10):
            over = {k: v for k, v in weights.items() if v > max_weight}
            if not over:
                break
            excess = sum(v - max_weight for v in over.values())
            for k in over:
                weights[k] = max_weight
            remaining = {k: v for k, v in weights.items() if k not in over}
            if not remaining:
                break
            total_remaining = sum(remaining.values())
            for k in remaining:
                weights[k] += excess * (remaining[k] / total_remaining)

    total = sum(weights.values())
    if total > 0:
        weights = {k: v / total for k, v in weights.items()}
    return weights


def _round_allocations(amounts: Dict[str, float]) -> Dict[str, int]:
    floors = {k: int(np.floor(v)) for k, v in amounts.items()}
    remainder = int(round(sum(amounts.values()) - sum(floors.values())))
    if remainder <= 0:
        return floors
    fractions = sorted(
        ((k, amounts[k] - floors[k]) for k in amounts),
        key=lambda x: x[1],
        reverse=True,
    )
    for i in range(min(remainder, len(fractions))):
        k, _ = fractions[i]
        floors[k] += 1
    return floors


def _apply_weight_caps(weights: Dict[str, float], max_weight: float) -> Dict[str, float]:
    if not weights or max_weight <= 0:
        return weights
    weights = weights.copy()
    for _ in range(10):
        over = {k: v for k, v in weights.items() if v > max_weight}
        if not over:
            break
        excess = sum(v - max_weight for v in over.values())
        for k in over:
            weights[k] = max_weight
        remaining = {k: v for k, v in weights.items() if k not in over}
        if not remaining or excess <= 0:
            break
        total_remaining = sum(remaining.values())
        for k in remaining:
            weights[k] += excess * (remaining[k] / total_remaining)
    total = sum(weights.values())
    if total > 0:
        weights = {k: v / total for k, v in weights.items()}
    return weights


def _apply_min_floor_preserve_count(
    weights: Dict[str, float],
    floor: float,
) -> Dict[str, float]:
    if not weights:
        return {}
    keys = list(weights.keys())
    n = len(keys)
    if n == 0:
        return {}

    cleaned: Dict[str, float] = {}
    for k, v in weights.items():
        try:
            w = float(v)
        except Exception:
            w = 0.0
        if not np.isfinite(w) or w < 0:
            w = 0.0
        cleaned[k] = w

    total = sum(cleaned.values())
    if total <= 0:
        base = {k: 1.0 / n for k in keys}
    else:
        base = {k: cleaned[k] / total for k in keys}

    max_floor = (1.0 / n) - 1e-9
    eff_floor = min(max(floor, 0.0), max_floor) if max_floor > 0 else 0.0
    residual = 1.0 - n * eff_floor
    if residual < 0:
        residual = 0.0

    out = {k: eff_floor + residual * base[k] for k in keys}
    return out


def _round_weights_to_step_preserve_count(
    weights: Dict[str, float],
    step: float,
    min_w: float,
    max_w: Optional[float] = None,
) -> Dict[str, float]:
    if not weights:
        return {}
    if step <= 0:
        return weights

    keys = list(weights.keys())
    n = len(keys)
    if n == 0:
        return {}

    def _round_half_up(v: float) -> int:
        return int(np.floor((v / step) + 0.5 + 1e-12))

    eff_min = min_w
    if n > 0:
        max_floor = (1.0 / n) - 1e-9
        if max_floor > 0:
            eff_min = min(min_w, max_floor)
        else:
            eff_min = 0.0
    min_units = int(np.ceil(eff_min / step)) if eff_min > 0 else 0
    max_units = int(np.floor(max_w / step)) if (max_w is not None and max_w > 0) else None
    if max_units is not None and max_units < min_units:
        max_units = min_units
    target_units = int(round(1.0 / step))

    units: Dict[str, int] = {}
    for k in keys:
        u = _round_half_up(weights.get(k, 0.0))
        if u < min_units:
            u = min_units
        if max_units is not None and u > max_units:
            u = max_units
        units[k] = u

    delta = target_units - sum(units.values())
    if delta > 0:
        while delta > 0:
            candidates = [k for k in keys if max_units is None or units[k] < max_units]
            if not candidates:
                break
            k = max(candidates, key=lambda x: units[x])
            units[k] += 1
            delta -= 1
    elif delta < 0:
        while delta < 0:
            candidates = [k for k in keys if units[k] > min_units]
            if not candidates:
                break
            k = max(candidates, key=lambda x: units[x])
            units[k] -= 1
            delta += 1

    if delta != 0:
        k = max(keys, key=lambda x: units[x])
        units[k] = max(units[k] + delta, min_units)

    result = {k: units[k] * step for k in keys}
    total = sum(result.values())
    if total > 0 and abs(total - 1.0) > 1e-9:
        result = {k: v / total for k, v in result.items()}
    return result


def _apply_practical_constraints(
    weights: Dict[str, float],
    min_w: float = 0.10,
    step: float = 0.05,
    max_w: Optional[float] = None,
) -> Dict[str, float]:
    if not weights:
        return {}

    floored = _apply_min_floor_preserve_count(weights, min_w)

    if max_w is not None and max_w > 0:
        floored = _apply_weight_caps(floored, max_w)
        total = sum(floored.values())
        if total > 0:
            floored = {k: v / total for k, v in floored.items()}

    rounded = _round_weights_to_step_preserve_count(floored, step, min_w, max_w=max_w)
    return rounded


def _returns_for_isins(
    isins: List[str],
    start: pd.Timestamp,
    end: pd.Timestamp,
    euro_rate: float,
    min_points: int = 60,
) -> Tuple[pd.DataFrame, Dict[str, str]]:
    series_map: Dict[str, pd.Series] = {}
    status: Dict[str, str] = {}
    for isin in isins:
        df, _, _ = get_price_series(isin, None, euro_rate)
        if df.empty:
            status[isin] = "insufficient"
            continue
        s = df["Close"].astype(float)
        s = s[(s.index >= start) & (s.index <= end)]
        if s.dropna().shape[0] < min_points:
            status[isin] = "insufficient"
            continue
        series_map[isin] = s
        status[isin] = "ok"
    if not series_map:
        return pd.DataFrame(), status
    prices = pd.DataFrame(series_map).ffill().dropna(how="any")
    if prices.empty or prices.shape[0] < min_points:
        return pd.DataFrame(), status
    returns = prices.pct_change().dropna(how="any")
    return returns, status


def _annualized_stats(returns: pd.DataFrame) -> Tuple[pd.Series, pd.Series]:
    if returns.empty:
        return pd.Series(dtype=float), pd.Series(dtype=float)
    # FIXED: arithmetic compound formula — was mean()*252 (linear, not compounded) (Résidu Bug 2)
    ann_return = (1 + returns.mean()) ** 252 - 1
    ann_vol = returns.std() * np.sqrt(252.0)
    return ann_return, ann_vol


def _avg_correlation(corr: pd.DataFrame) -> float:
    if corr.empty or corr.shape[0] < 2:
        return np.nan
    vals = corr.values[np.triu_indices_from(corr, 1)]
    return float(np.nanmean(vals)) if vals.size else np.nan


def _avg_offdiag_corr(corr: pd.DataFrame) -> float:
    if corr.empty or corr.shape[0] < 2:
        return np.nan
    vals = corr.values[np.triu_indices_from(corr, 1)]
    return float(np.nanmean(vals)) if vals.size else np.nan


def _select_min_corr_subset(
    candidates: List[str],
    returns: pd.DataFrame,
    k: int,
    anchor: Optional[str] = None,
) -> List[str]:
    if k <= 0 or not candidates:
        return []
    if anchor and anchor not in candidates:
        return []
    if k >= len(candidates):
        return candidates
    corr = returns.corr()
    if corr.empty:
        return candidates[:k]
    if anchor:
        pool = [c for c in candidates if c != anchor]
        best_combo = [anchor]
        best_score = None
        max_checks = 2000
        combos = itertools.combinations(pool, k - 1)
        for idx, combo in enumerate(combos):
            if idx >= max_checks:
                break
            subset = [anchor, *combo]
            score = _avg_offdiag_corr(corr.loc[subset, subset])
            if best_score is None or score < best_score:
                best_score = score
                best_combo = subset
        if best_score is None:
            best_combo = [anchor] + pool[: k - 1]
        return best_combo
    max_checks = 2000
    best_combo: List[str] = []
    best_score = None
    for idx, combo in enumerate(itertools.combinations(candidates, k)):
        if idx >= max_checks:
            break
        subset = list(combo)
        score = _avg_offdiag_corr(corr.loc[subset, subset])
        if best_score is None or score < best_score:
            best_score = score
            best_combo = subset
    if not best_combo:
        best_combo = candidates[:k]
    return best_combo


def _greedy_select(
    candidates: List[str],
    returns: pd.DataFrame,
    target_count: int,
    forced: Optional[str] = None,
    corr_penalty: float = 0.35,
) -> List[str]:
    selected: List[str] = []
    if forced and forced in candidates:
        selected.append(forced)
    remaining = [c for c in candidates if c not in selected]
    if returns.empty or target_count <= 0:
        return selected
    ann_return, ann_vol = _annualized_stats(returns[remaining + selected] if remaining else returns)
    rfr = get_risk_free_rate()  # taux sans risque dynamique (Bund 10 ans ou saisie manuelle)
    sharpe = (ann_return - rfr) / ann_vol.replace(0, np.nan)
    corr = returns.corr() if not returns.empty else pd.DataFrame()

    while len(selected) < target_count and remaining:
        best = None
        best_score = None
        for cand in remaining:
            base_sharpe = float(sharpe.get(cand, np.nan))
            if np.isnan(base_sharpe):
                base_sharpe = -1e9
            if selected and not corr.empty:
                avg_corr = float(corr.loc[cand, selected].mean())
            else:
                avg_corr = 0.0
            score = base_sharpe - corr_penalty * avg_corr
            if best_score is None or score > best_score:
                best_score = score
                best = cand
        if best is None:
            break
        selected.append(best)
        remaining = [c for c in remaining if c != best]
    return selected


def _optimize_uc_weights(
    returns: pd.DataFrame,
    objective: str,
    min_weight: float,
    max_weight: float,
    target_vol: Optional[float],
    target_return: Optional[float],
) -> Dict[str, float]:
    if returns.empty:
        return {}
    n_assets = returns.shape[1]
    if n_assets == 0:
        return {}
    bounds = (min_weight, max_weight)
    try:
        if not PYPFOPT_AVAILABLE:
            raise RuntimeError(PYPFOPT_ERROR)
        mu = expected_returns.mean_historical_return(returns, frequency=252)
        cov = risk_models.sample_cov(returns, frequency=252)
        ef = EfficientFrontier(mu, cov, weight_bounds=bounds)
        if objective == "max_sharpe":
            weights = ef.max_sharpe(risk_free_rate=get_risk_free_rate())
        elif objective == "target_vol" and target_vol is not None:
            weights = ef.efficient_risk(target_vol)
        elif objective == "target_return" and target_return is not None:
            weights = ef.efficient_return(target_return)
        else:
            weights = ef.max_sharpe(risk_free_rate=get_risk_free_rate())
        cleaned = ef.clean_weights()
        total = sum(cleaned.values())
        if total > 0:
            return {k: v / total for k, v in cleaned.items()}
    except Exception:
        pass
    equal_weight = 1.0 / n_assets
    return {col: equal_weight for col in returns.columns}


def _select_min_corr_combo(
    returns: pd.DataFrame,
    k: int,
    anchor: Optional[str] = None,
) -> List[str]:
    if returns.empty:
        return []
    cols = list(returns.columns)
    if anchor and anchor not in cols:
        return []
    if k <= 0:
        return []
    if k > len(cols):
        return cols
    corr = returns.corr()
    best_combo: List[str] = []
    best_score = None
    if anchor:
        others = [c for c in cols if c != anchor]
        for combo in itertools.combinations(others, k - 1):
            candidate = [anchor, *combo]
            sub = corr.loc[candidate, candidate].to_numpy()
            triu = sub[np.triu_indices_from(sub, 1)]
            score = float(triu.mean()) if triu.size else 1.0
            if best_score is None or score < best_score:
                best_score = score
                best_combo = candidate
    else:
        for combo in itertools.combinations(cols, k):
            sub = corr.loc[list(combo), list(combo)].to_numpy()
            triu = sub[np.triu_indices_from(sub, 1)]
            score = float(triu.mean()) if triu.size else 1.0
            if best_score is None or score < best_score:
                best_score = score
                best_combo = list(combo)
    return best_combo


def _compute_drawdown(returns: pd.Series) -> Optional[float]:
    if returns.empty:
        return None
    cum = (1.0 + returns).cumprod()
    running_max = cum.cummax()
    dd = cum / running_max - 1.0
    return float(dd.min())


def _fund_name(isin: str) -> str:
    return FUND_NAME_MAP.get(isin, isin)


def _safe_fund_label(name: str, isin: str) -> str:
    cleaned = str(name or "").strip()
    if cleaned:
        return f"{cleaned} ({isin})"
    if isin:
        res = eodhd_search(isin)
        if res:
            maybe = res[0].get("Name") or res[0].get("name")
            if maybe:
                return f"{maybe} ({isin})"
    return f"{isin} ({isin})"


def render_portfolio_builder():
    st.title("Construction de portefeuille optimisé")

    st.session_state.setdefault("PP_RUN", False)
    st.session_state.setdefault("PP_PARAMS_HASH", "")
    st.session_state.setdefault("PP_OPT_START_DATE", (TODAY - pd.DateOffset(years=3)).date())
    st.session_state.setdefault("PP_OPT_END_DATE", TODAY.date())
    st.session_state.setdefault("PP_CONTRACT_LABEL", list(CONTRACTS_REGISTRY.keys())[0])
    st.session_state.setdefault("PP_TOTAL_UC_COUNT", 4)

    # Charger le référentiel du contrat sélectionné (indépendant de render_app)
    pp_contract_label = st.session_state.get("PP_CONTRACT_LABEL", list(CONTRACTS_REGISTRY.keys())[0])
    pp_contract_cfg = CONTRACTS_REGISTRY.get(pp_contract_label, list(CONTRACTS_REGISTRY.values())[0])
    funds_df = load_contract_funds(
        pp_contract_cfg["path"],
        pp_contract_cfg["funds_filename"],
    )
    CONTRACT_FUND_NAMES = dict(zip(funds_df["isin"], funds_df["name"])) if not funds_df.empty else FUND_NAME_MAP.copy()
    CONTRACT_FUND_FEES  = dict(zip(funds_df["isin"], funds_df["fee_total_pct"])) if not funds_df.empty else {}

    profile_map = {
        "Prudent": 50,
        "Equilibre": 30,
        "Dynamique": 20,
        "Agressif": 10,
    }

    with st.sidebar:
        st.header("Paramètres clés")

        # ── Contrat & fonds en euros ───────────────────────────
        pp_contract_label = st.selectbox(
            "Contrat",
            list(CONTRACTS_REGISTRY.keys()),
            key="PP_CONTRACT_LABEL",
        )
        pp_contract_cfg = CONTRACTS_REGISTRY[pp_contract_label]
        pp_euro_options = list(pp_contract_cfg.get("euro_funds", {}).keys())
        pp_euro_fund_label = st.selectbox(
            "Fonds en euros",
            pp_euro_options if pp_euro_options else ["—"],
            key="PP_EURO_FUND_LABEL",
        )
        st.markdown("---")

        # ── Profil & budget ────────────────────────────────────
        profile = st.selectbox(
            "Profil de risque client",
            list(profile_map.keys()),
            index=1,
            key="PP_PROFILE",
        )
        _recommended_euro_pct = profile_map[profile]
        st.info(
            f"📋 Allocation fonds en euros recommandée "
            f"pour un profil **{profile}** : **{_recommended_euro_pct}%**"
        )
        euro_pct = int(st.number_input(
            "Part fonds en euros (%)",
            min_value=0,
            max_value=100,
            value=int(st.session_state.get("PP_EURO_PCT", _recommended_euro_pct)),
            step=5,
            key="PP_EURO_PCT",
            help="Valeur recommandée selon le profil. Modifiable selon la situation du client.",
        ))

        total_budget = st.number_input(
            "Budget total (EUR)",
            min_value=0,
            max_value=10_000_000,
            value=100_000,
            step=10,
            key="PP_TOTAL_BUDGET",
        )

        # ── Fenêtre d'analyse (détermine opt_start / opt_end) ──
        opt_window_mode = st.radio(
            "Fenêtre d'analyse",
            ["1 an", "3 ans", "5 ans", "10 ans", "Dates personnalisées"],
            horizontal=False,
            key="PP_WINDOW_MODE",
        )

        if opt_window_mode == "Dates personnalisées":
            opt_start_date = st.date_input(
                "Date de début",
                value=st.session_state.get("PP_OPT_START_DATE", (TODAY - pd.DateOffset(years=3)).date()),
                key="PP_OPT_START_DATE",
            )
            opt_end_date = st.date_input(
                "Date de fin",
                value=st.session_state.get("PP_OPT_END_DATE", TODAY.date()),
                key="PP_OPT_END_DATE",
            )
            opt_start = pd.Timestamp(opt_start_date)
            opt_end = pd.Timestamp(opt_end_date)
        else:
            years_map = {"1 an": 1, "3 ans": 3, "5 ans": 5, "10 ans": 10}
            opt_years = years_map[opt_window_mode]
            opt_start = TODAY - pd.DateOffset(years=opt_years)
            opt_end = TODAY

        # ── Taux fonds en euros (auto depuis historique) ───────
        st.markdown("---")
        st.markdown("**Rendement fonds en euros**")
        try:
            _pp_ef_filename = pp_contract_cfg["euro_funds"].get(pp_euro_fund_label, "")
            _pp_ef_history = (
                load_euro_fund_history(pp_contract_cfg["path"], _pp_ef_filename)
                if _pp_ef_filename else pd.DataFrame()
            )
        except Exception:
            _pp_ef_history = pd.DataFrame()

        _auto_euro_rate = _compute_auto_euro_rate(_pp_ef_history, opt_start, opt_end)
        if _auto_euro_rate is not None:
            st.success(
                f"Taux net moyen **{pp_euro_fund_label}** "
                f"sur la fenêtre : **{_auto_euro_rate:.2f}%**"
            )
        else:
            st.info("Historique insuffisant sur la fenêtre — saisissez le taux manuellement.")

        _override_euro = st.checkbox(
            "Saisir un taux personnalisé",
            value=False,
            key="PP_OVERRIDE_EURO_RATE",
        )
        if _override_euro:
            euro_rate = float(st.number_input(
                "Rendement annuel du fonds en euros (%)",
                min_value=0.0, max_value=10.0,
                value=float(_auto_euro_rate if _auto_euro_rate is not None else 2.0),
                step=0.10,
                key="PP_EURO_RATE",
            ))
        else:
            euro_rate = float(_auto_euro_rate if _auto_euro_rate is not None else 2.0)
            st.session_state["PP_EURO_RATE"] = euro_rate
            st.caption(f"Taux utilisé : **{euro_rate:.2f}%**")

        # ── Taux sans risque (Sharpe) ──────────────────────────
        st.markdown("---")
        st.markdown("**Taux sans risque (Sharpe)**")
        rfr_api = _fetch_bund_rate_from_api()
        if rfr_api is not None:
            st.success(
                f"📡 Bund 10 ans : **{rfr_api * 100:.2f}%** — récupéré automatiquement"
            )
            st.session_state["RISK_FREE_RATE_SOURCE"] = "api"
            st.session_state["RISK_FREE_RATE_VALUE"] = rfr_api
        else:
            st.warning("⚠️ Bund 10 ans indisponible via API")
        default_pct = (
            rfr_api * 100.0
            if rfr_api is not None
            else float(st.session_state.get("RISK_FREE_RATE_MANUAL_PCT",
                                            RISK_FREE_RATE_FALLBACK * 100.0))
        )
        manual_rfr_pct = st.number_input(
            "Taux sans risque (%)",
            min_value=0.0,
            max_value=15.0,
            value=default_pct,
            step=0.05,
            format="%.2f",
            help=(
                "Taux de référence : rendement du Bund allemand 10 ans. "
                "Valeur actuelle : https://www.investing.com/rates-bonds/"
                "germany-10-year-bond-yield"
            ),
            key="RFR_MANUAL_INPUT",
        )
        st.session_state["RISK_FREE_RATE_MANUAL_PCT"] = manual_rfr_pct
        st.session_state["RISK_FREE_RATE_MANUAL"] = manual_rfr_pct / 100.0
        if rfr_api is None:
            st.session_state["RISK_FREE_RATE_SOURCE"] = "manual"

    # ── Fonds UC configurables ─────────────────────────────────
    st.markdown("### Fonds UC")

    # Utiliser _is_bond_category (définie au niveau module) pour une
    # classification cohérente avec le comparateur.
    _all_cats = sorted(funds_df["category"].dropna().unique().tolist()) if not funds_df.empty else []
    _bond_cats = [c for c in _all_cats if _is_bond_category(c)]
    _other_cats = [c for c in _all_cats if not _is_bond_category(c)]
    _bond_cat_options  = ["(Toutes obligataires)"] + _bond_cats
    _other_cat_options = ["(Toutes)"] + _other_cats

    # ── Fonds obligataires ─────────────────────────────────────
    st.markdown("#### Fonds obligataires")
    n_bond = int(st.number_input(
        "Nombre de fonds obligataires",
        min_value=0, max_value=8,
        value=int(st.session_state.get("PP_N_BOND", 1)),
        step=1,
        key="PP_N_BOND",
        help="Nombre de fonds UC à dominante obligataire à inclure dans le portefeuille.",
    ))
    bond_slot_configs: List[Dict[str, Any]] = []
    for _i in range(n_bond):
        with st.container(border=True):
            st.caption(f"Fonds obligataire {_i + 1}")
            _sc1, _sc2 = st.columns([3, 1])
            with _sc1:
                _cat = st.selectbox(
                    "Catégorie",
                    _bond_cat_options,
                    key=f"PP_BOND_CAT_{_i}",
                )
                _cat_normalized = "(Toutes)" if _cat == "(Toutes obligataires)" else _cat
            with _sc2:
                _sri = st.selectbox(
                    "SRI max",
                    ["Tous"] + [str(j) for j in range(1, 8)],
                    key=f"PP_BOND_SRI_{_i}",
                )
        bond_slot_configs.append({"cat": _cat_normalized, "sri": _sri})

    st.markdown("---")

    # ── Autres fonds UC ────────────────────────────────────────
    st.markdown("#### Autres fonds UC")
    n_other = int(st.number_input(
        "Nombre d'autres fonds UC",
        min_value=0, max_value=10,
        value=int(st.session_state.get("PP_N_OTHER", 3)),
        step=1,
        key="PP_N_OTHER",
        help="Fonds actions, diversifiés, thématiques, etc.",
    ))
    other_slot_configs: List[Dict[str, Any]] = []
    for _i in range(n_other):
        with st.container(border=True):
            st.caption(f"Fonds UC {_i + 1}")
            _sc1, _sc2 = st.columns([3, 1])
            with _sc1:
                _cat = st.selectbox(
                    "Catégorie",
                    _other_cat_options,
                    key=f"PP_OTHER_CAT_{_i}",
                )
            with _sc2:
                _sri = st.selectbox(
                    "SRI max",
                    ["Tous"] + [str(j) for j in range(1, 8)],
                    key=f"PP_OTHER_SRI_{_i}",
                )
        other_slot_configs.append({"cat": _cat, "sri": _sri})

    # ── Fusion pour compatibilité avec le reste du code ────────
    uc_slot_configs: List[Dict[str, Any]] = bond_slot_configs + other_slot_configs
    total_uc_count = len(uc_slot_configs)
    st.session_state["PP_TOTAL_UC_COUNT"] = total_uc_count

    objective_options = [
        "Maximiser Sharpe",
        "Minimiser volatilite",
        "Maximiser rendement annualise",
        "Meilleur compromis (Sharpe + diversification)",
        "Diversification maximale (decorrelation)",
    ]
    if st.session_state.get("PP_OBJECTIVE") == "Risk parity":
        st.session_state["PP_OBJECTIVE"] = objective_options[0]
    objective_choice = st.selectbox(
        "Objectif",
        objective_options,
        key="PP_OBJECTIVE",
    )

    practical_mode = st.checkbox(
        "Optimisation terrain (min 10% + arrondi 5%)",
        value=True,
        key="PP_PRACTICAL_MODE",
    )

    force_fund = st.checkbox("Forcer un fonds (ancre)", value=False, key="PP_FORCE_ANCHOR")
    forced_isin = None
    if force_fund:
        if not funds_df.empty:
            _anchor_options = [
                f"{row['name']} ({row['isin']})"
                for _, row in funds_df.iterrows()
            ]
            _anchor_lookup = {
                f"{row['name']} ({row['isin']})": row["isin"]
                for _, row in funds_df.iterrows()
            }
            _forced_choice = st.selectbox("Fonds ancre", _anchor_options, key="PP_ANCHOR_ISIN")
            forced_isin = _anchor_lookup.get(_forced_choice)

    params = (
        profile,
        int(euro_pct),
        float(euro_rate),
        int(total_budget),
        opt_window_mode,
        str(opt_start.date()),
        str(opt_end.date()),
        tuple((s["cat"], s["sri"]) for s in uc_slot_configs),
        objective_choice,
        bool(force_fund),
        forced_isin,
    )
    current_hash = params_hash(params)

    if st.session_state.get("PP_RUN") and st.session_state.get("PP_PARAMS_HASH") != current_hash:
        st.session_state["PP_RUN"] = False
        st.info("Paramètres modifiés. Cliquez de nouveau sur '✅ Lancer l'optimisation'.")

    run_clicked = st.button("✅ Lancer l'optimisation", type="primary")
    if run_clicked:
        st.session_state["PP_RUN"] = True
        st.session_state["PP_PARAMS_HASH"] = current_hash

    if not st.session_state.get("PP_RUN"):
        st.info("Renseignez tous les paramètres puis cliquez sur '✅ Lancer l'optimisation'.")
        return

    if opt_start > opt_end:
        st.warning("La date de debut doit etre anterieure a la date de fin.")
        return

    try:
        if funds_df.empty:
            st.warning(
                "Référentiel de fonds non chargé. "
                "Sélectionnez un contrat dans la barre latérale."
            )
            return

        # Construire l'univers de candidats depuis les slots UC
        def _slot_candidates(slot: Dict[str, Any]) -> List[str]:
            """Retourne la liste des ISIN du contrat correspondant au filtre slot."""
            df = funds_df.copy()
            if slot["cat"] != "(Toutes)":
                df = df[df["category"] == slot["cat"]]
            if slot["sri"] != "Tous":
                df = df[df["sri"] <= int(slot["sri"])]
            return df["isin"].tolist()

        all_slot_isins: List[str] = []
        for slot in uc_slot_configs:
            all_slot_isins.extend(_slot_candidates(slot))

        all_candidates = sorted(set(all_slot_isins))
        if not all_candidates:
            st.info("Aucun fonds UC disponible avec les critères de slots choisis.")
            return

        returns_all, status_all = _returns_for_isins(all_candidates, opt_start, opt_end, euro_rate=float(euro_rate))

        insufficient = [isin for isin, status in status_all.items() if status != "ok"]
        valid_all = [isin for isin in all_candidates if status_all.get(isin) == "ok"]

        if insufficient:
            st.warning("Certains fonds ont été exclus (historique insuffisant sur la fenêtre).")

        if forced_isin and forced_isin not in valid_all:
            st.warning("Le fonds ancre est indisponible sur la période et a été ignoré.")
            forced_isin = None

        if returns_all.empty:
            st.info("Historique insuffisant pour calculer les rendements sur la fenêtre choisie.")
            return

        def _zscore(s: pd.Series) -> pd.Series:
            if s.empty:
                return s
            std = float(s.std())
            if std == 0 or np.isnan(std):
                return pd.Series(0.0, index=s.index)
            return (s - s.mean()) / std

        def _stats_for(candidates: List[str], ret_df: pd.DataFrame) -> Tuple[pd.Series, pd.Series, pd.Series, pd.DataFrame, pd.Series]:
            if not candidates or ret_df.empty:
                return pd.Series(dtype=float), pd.Series(dtype=float), pd.Series(dtype=float), pd.DataFrame(), pd.Series(dtype=float)
            # FIXED: arithmetic annualisation formula — was incorrectly treating pct_change() returns
            # as log-returns then applying exp(), double-compounding the result (Bug 2)
            rfr = get_risk_free_rate()  # taux sans risque dynamique (Bund 10 ans ou saisie manuelle)
            ann_return = (1 + ret_df[candidates].mean()) ** 252 - 1
            ann_vol = ret_df[candidates].std() * np.sqrt(252.0)
            sharpe = ((ann_return - rfr) / ann_vol.replace(0, np.nan)).replace([np.inf, -np.inf], np.nan)
            corr = ret_df[candidates].corr()
            avg_corr = corr.apply(lambda row: row.drop(labels=row.name, errors="ignore").mean(), axis=1).fillna(0.0) if not corr.empty else pd.Series(0.0, index=candidates)
            return ann_return, ann_vol, sharpe, corr, avg_corr

        def _greedy_min_corr(candidates: List[str], corr: pd.DataFrame, k: int, seed: Optional[List[str]] = None) -> List[str]:
            if k <= 0 or not candidates:
                return []
            selected = [x for x in (seed or []) if x in candidates]
            remaining = [c for c in candidates if c not in selected]
            while len(selected) < k and remaining:
                best, best_score = None, None
                for cand in remaining:
                    score = float(corr.loc[cand, selected].mean()) if selected and not corr.empty else 0.0
                    if best_score is None or score < best_score:
                        best, best_score = cand, score
                if best is None:
                    break
                selected.append(best)
                remaining = [c for c in remaining if c != best]
            return selected

        def _select_by_objective(candidates: List[str], ret_df: pd.DataFrame, k: int, objective: str, forced: Optional[str] = None) -> List[str]:
            if k <= 0 or not candidates:
                return []
            if ret_df.empty:
                base = candidates[:k]
                if forced and forced in candidates and forced not in base:
                    base = [forced] + [x for x in base if x != forced]
                    base = base[:k]
                return base

            ann_ret, ann_vol, sharpe, corr, avg_corr = _stats_for(candidates, ret_df)
            sharpe_rank = sharpe.fillna(-1e9).sort_values(ascending=False)

            if objective == "Maximiser rendement annualise":
                selected = ann_ret.sort_values(ascending=False).index.tolist()[:k]
            elif objective == "Minimiser volatilite":
                selected = ann_vol.sort_values(ascending=True).index.tolist()[:k]
            elif objective == "Maximiser Sharpe":
                selected = sharpe_rank.index.tolist()[:k]
            elif objective in ("Diversification maximale (decorrelation)", "Risk parity"):
                seed = sharpe_rank.index.tolist()[:1]
                selected = _greedy_min_corr(candidates, corr, k, seed=seed)
            elif objective == "Meilleur compromis (Sharpe + diversification)":
                score = _zscore(sharpe.fillna(0.0)) + _zscore(1.0 - avg_corr)
                selected = score.sort_values(ascending=False).index.tolist()[:k]
            else:
                selected = sharpe_rank.index.tolist()[:k]

            if forced and forced in candidates:
                if forced not in selected:
                    if objective in ("Diversification maximale (decorrelation)", "Risk parity"):
                        selected = _greedy_min_corr(candidates, corr, k, seed=[forced])
                    else:
                        rest = [x for x in selected if x != forced]
                        selected = [forced] + rest
                        selected = selected[:k]
                else:
                    selected = [forced] + [x for x in selected if x != forced]
            return selected[:k]

        # Sélection par slot : 1 fonds par slot, dédupliqué
        _seen: set = set()
        selected_isins: List[str] = []
        _isin_slot_label: Dict[str, str] = {}  # isin → label catégorie pour affichage
        _anchor_placed = False
        for _si, _slot in enumerate(uc_slot_configs):
            _slot_valid = [
                isin for isin in _slot_candidates(_slot)
                if isin in valid_all and isin not in _seen
            ]
            if not _slot_valid:
                continue
            _forced_for_slot = forced_isin if (not _anchor_placed and forced_isin in _slot_valid) else None
            _slot_returns = returns_all[_slot_valid] if _slot_valid else pd.DataFrame()
            _picked = _select_by_objective(_slot_valid, _slot_returns, 1, objective_choice, forced=_forced_for_slot)
            for _isin in _picked:
                if _isin not in _seen:
                    selected_isins.append(_isin)
                    _seen.add(_isin)
                    _cat_label = _slot["cat"] if _slot["cat"] != "(Toutes)" else (
                        funds_df.loc[funds_df["isin"] == _isin, "category"].values[0]
                        if not funds_df.loc[funds_df["isin"] == _isin].empty else "UC"
                    )
                    _isin_slot_label[_isin] = _cat_label
                    if _forced_for_slot and _isin == forced_isin:
                        _anchor_placed = True
        if not selected_isins:
            st.info("Aucun fonds selectionne apres filtrage.")
            return

        returns_selected = returns_all[selected_isins].dropna(how="any")
        if returns_selected.empty:
            st.info("Historique insuffisant apres selection des UC.")
            return

        uc_total = max(0.0, 1.0 - (float(euro_pct) / 100.0))
        if uc_total <= 0.0:
            st.info("Part UC nulle avec le profil choisi.")
            return

        cap_uc_final = 0.25
        uc_max_bound = min(cap_uc_final / uc_total, 1.0)

        def _risk_parity_weights(ret_df: pd.DataFrame) -> Dict[str, float]:
            cov = ret_df.cov()
            if cov.empty:
                return {}
            vols = np.sqrt(np.diag(cov.values))
            inv = np.where(vols > 0, 1.0 / vols, 0.0)
            if float(inv.sum()) <= 0:
                return {c: 1.0 / len(ret_df.columns) for c in ret_df.columns}
            w = inv / inv.sum()
            return {c: float(w[i]) for i, c in enumerate(ret_df.columns)}

        weights_uc_raw: Dict[str, float] = {}
        if PYPFOPT_AVAILABLE:
            try:
                mu = expected_returns.mean_historical_return(returns_selected, returns_data=True, frequency=252)
                cov = risk_models.sample_cov(returns_selected, returns_data=True, frequency=252)
                ef = EfficientFrontier(mu, cov, weight_bounds=(0.0, uc_max_bound))

                if objective_choice == "Maximiser Sharpe":
                    ef.max_sharpe(risk_free_rate=get_risk_free_rate())
                    weights_uc_raw = ef.clean_weights()
                elif objective_choice == "Minimiser volatilite":
                    ef.min_volatility()
                    weights_uc_raw = ef.clean_weights()
                elif objective_choice == "Maximiser rendement annualise":
                    if not mu.empty:
                        target = float(mu.max()) - 1e-6
                        try:
                            ef.efficient_return(target_return=target)
                        except Exception:
                            ef.max_sharpe(risk_free_rate=get_risk_free_rate())
                    else:
                        ef.max_sharpe(risk_free_rate=get_risk_free_rate())
                    weights_uc_raw = ef.clean_weights()
                elif objective_choice == "Risk parity":
                    weights_uc_raw = _risk_parity_weights(returns_selected)
                else:
                    ef.min_volatility() if objective_choice == "Diversification maximale (decorrelation)" else ef.max_sharpe(risk_free_rate=get_risk_free_rate())
                    weights_uc_raw = ef.clean_weights()
            except Exception:
                weights_uc_raw = {}

        if not weights_uc_raw:
            if objective_choice == "Risk parity":
                weights_uc_raw = _risk_parity_weights(returns_selected)
            elif objective_choice == "Minimiser volatilite":
                vol = returns_selected.std() * np.sqrt(252.0)
                score = (1.0 / vol.replace(0, np.nan)).fillna(0.0)
                weights_uc_raw = (score / score.sum()).to_dict() if float(score.sum()) > 0 else {}
            elif objective_choice == "Maximiser rendement annualise":
                # FIXED: arithmetic compound annualisation (Résidu Bug 2)
                ann_ret = (1 + returns_selected.mean()) ** 252 - 1
                score = ann_ret.clip(lower=0.0)
                weights_uc_raw = (score / score.sum()).to_dict() if float(score.sum()) > 0 else {}
            elif objective_choice == "Meilleur compromis (Sharpe + diversification)":
                # FIXED: arithmetic compound annualisation for Sharpe (Résidu Bug 2)
                rfr = get_risk_free_rate()
                ann_ret = (1 + returns_selected.mean()) ** 252 - 1
                vol = returns_selected.std() * np.sqrt(252.0)
                sharpe = ((ann_ret - rfr) / vol.replace(0, np.nan)).fillna(0.0)
                corr = returns_selected.corr()
                avg_corr = corr.apply(lambda row: row.drop(labels=row.name, errors="ignore").mean(), axis=1).fillna(0.0) if not corr.empty else pd.Series(0.0, index=returns_selected.columns)
                score = _zscore(sharpe) + _zscore(1.0 - avg_corr)
                score = score.clip(lower=0.0)
                weights_uc_raw = (score / score.sum()).to_dict() if float(score.sum()) > 0 else {}
            elif objective_choice == "Diversification maximale (decorrelation)":
                corr = returns_selected.corr()
                avg_corr = corr.apply(lambda row: row.drop(labels=row.name, errors="ignore").mean(), axis=1).fillna(0.0) if not corr.empty else pd.Series(0.0, index=returns_selected.columns)
                score = (1.0 - avg_corr).clip(lower=0.0)
                weights_uc_raw = (score / score.sum()).to_dict() if float(score.sum()) > 0 else {}
            else:
                # FIXED: arithmetic compound annualisation for default Sharpe (Résidu Bug 2)
                rfr = get_risk_free_rate()
                ann_ret = (1 + returns_selected.mean()) ** 252 - 1
                vol = returns_selected.std() * np.sqrt(252.0)
                score = ((ann_ret - rfr) / vol.replace(0, np.nan)).fillna(0.0).clip(lower=0.0)
                weights_uc_raw = (score / score.sum()).to_dict() if float(score.sum()) > 0 else {}

        if not weights_uc_raw:
            weights_uc_raw = {isin: 1.0 / len(selected_isins) for isin in selected_isins}

        weights_uc_raw = {k: float(v) for k, v in weights_uc_raw.items() if k in selected_isins}
        if not weights_uc_raw:
            weights_uc_raw = {isin: 1.0 / len(selected_isins) for isin in selected_isins}

        weights_uc_raw = _apply_weight_caps(weights_uc_raw, uc_max_bound)
        total_uc_raw = float(sum(weights_uc_raw.values()))
        if total_uc_raw <= 0:
            weights_uc_raw = {isin: 1.0 / len(selected_isins) for isin in selected_isins}
            weights_uc_raw = _apply_weight_caps(weights_uc_raw, uc_max_bound)
            total_uc_raw = float(sum(weights_uc_raw.values()))
        if total_uc_raw > 0:
            weights_uc_raw = {k: v / total_uc_raw for k, v in weights_uc_raw.items()}

        if practical_mode:
            weights_uc_raw = _apply_practical_constraints(
                weights_uc_raw,
                min_w=0.10,
                step=0.05,
                max_w=uc_max_bound,
            )

        weights_final = {"EUROFUND": float(euro_pct) / 100.0}
        for isin in selected_isins:
            weights_final[isin] = float(weights_uc_raw.get(isin, 0.0)) * uc_total

        # Security clamp + exact renormalization
        weights_final = {k: max(0.0, float(v)) for k, v in weights_final.items()}
        total_weight = float(sum(weights_final.values()))
        if total_weight > 0:
            weights_final = {k: v / total_weight for k, v in weights_final.items()}

        def _round_allocations_to_step(amounts: Dict[str, float], step: int, total: int) -> Dict[str, int]:
            if step <= 0:
                step = 1
            rounded = {k: int(np.floor(max(0.0, v) / step)) * step for k, v in amounts.items()}
            remaining = int(total - sum(rounded.values()))
            if remaining > 0 and rounded:
                frac_rank = sorted(
                    ((k, (max(0.0, amounts[k]) / step) - np.floor(max(0.0, amounts[k]) / step)) for k in rounded.keys()),
                    key=lambda x: x[1],
                    reverse=True,
                )
                idx = 0
                while remaining >= step and frac_rank:
                    k = frac_rank[idx % len(frac_rank)][0]
                    rounded[k] += step
                    remaining -= step
                    idx += 1
                if remaining != 0:
                    last_key = list(rounded.keys())[-1]
                    rounded[last_key] = max(0, rounded[last_key] + remaining)
            elif remaining < 0 and rounded:
                last_key = list(rounded.keys())[-1]
                rounded[last_key] = max(0, rounded[last_key] + remaining)

            delta = int(total - sum(rounded.values()))
            if rounded and delta != 0:
                last_key = list(rounded.keys())[-1]
                rounded[last_key] = max(0, rounded[last_key] + delta)
            return rounded

        amounts_raw = {k: float(total_budget) * float(w) for k, w in weights_final.items()}
        amounts = _round_allocations_to_step(amounts_raw, step=10, total=int(total_budget))

        rows = [
            {
                "Support": "Fonds en euros (EUROFUND)",
                "Categorie": "EUROFUND",
                "%": weights_final.get("EUROFUND", 0.0) * 100.0,
                "Montant EUR": amounts.get("EUROFUND", 0),
            }
        ]
        for isin in selected_isins:
            rows.append(
                {
                    "Support": CONTRACT_FUND_NAMES.get(isin, FUND_NAME_MAP.get(isin, isin)),
                    "Categorie": _isin_slot_label.get(isin, "UC"),
                    "%": weights_final.get(isin, 0.0) * 100.0,
                    "Montant EUR": amounts.get(isin, 0),
                }
            )
        df_alloc = pd.DataFrame(rows)

        st.markdown("**Allocation finale**")
        st.dataframe(
            df_alloc.style.format({"%": "{:,.2f}%".format, "Montant EUR": to_eur}),
            use_container_width=True,
            hide_index=True,
        )

        if MATPLOTLIB_AVAILABLE:
            fig, ax = plt.subplots(figsize=(5.0, 3.2))
            ax.pie(df_alloc["Montant EUR"], labels=df_alloc["Support"], autopct="%1.1f%%")
            ax.set_title("Repartition finale")
            st.pyplot(fig)
            plt.close(fig)
        else:
            st.warning(f"Graphique indisponible ({MATPLOTLIB_ERROR}).")

        if len(selected_isins) >= 2:
            st.markdown("**Heatmap correlation UC**")
            corr_uc = returns_selected.corr().copy()
            corr_uc["Ligne1"] = corr_uc.index
            heat_df = corr_uc.melt(id_vars="Ligne1", var_name="Ligne2", value_name="corr")
            heat = (
                alt.Chart(heat_df)
                .mark_rect()
                .encode(
                    x=alt.X("Ligne1:O", sort=None, title=""),
                    y=alt.Y("Ligne2:O", sort=None, title=""),
                    color=alt.Color("corr:Q", scale=alt.Scale(domain=[-1, 1])),
                    tooltip=[
                        alt.Tooltip("Ligne1:N", title="Ligne 1"),
                        alt.Tooltip("Ligne2:N", title="Ligne 2"),
                        alt.Tooltip("corr:Q", title="Correlation", format=".2f"),
                    ],
                )
                .properties(height=260)
            )
            st.altair_chart(heat, use_container_width=True)

        uc_weights_norm = pd.Series({k: weights_uc_raw.get(k, 0.0) for k in selected_isins}, index=selected_isins)
        uc_weights_norm = uc_weights_norm / uc_weights_norm.sum() if float(uc_weights_norm.sum()) > 0 else pd.Series(1.0 / len(selected_isins), index=selected_isins)

        port_log_ret = returns_selected[selected_isins].dot(uc_weights_norm.values)
        # FIXED: arithmetic compound annualisation — was exp(mean*252)-1 on arithmetic returns (Résidu Bug 2)
        rfr = get_risk_free_rate()  # taux sans risque dynamique (Bund 10 ans ou saisie manuelle)
        ann_ret_uc = float((1 + port_log_ret.mean()) ** 252 - 1)
        ann_vol_uc = float(port_log_ret.std() * np.sqrt(252.0))
        sharpe_uc = (ann_ret_uc - rfr) / ann_vol_uc if ann_vol_uc > 0 else np.nan

        # FIXED: use fee-aware wrapper so management fees are included in the return metric (Correction 1)
        euro_df, _, _ = get_price_series_with_fees("EUROFUND", None, float(euro_rate))
        euro_df = euro_df.loc[(euro_df.index >= opt_start) & (euro_df.index <= opt_end)]
        if euro_df.empty:
            euro_total_ret = 0.0
        else:
            euro_total_ret = float(euro_df["Close"].iloc[-1] / euro_df["Close"].iloc[0] - 1.0)

        uc_path = (1.0 + port_log_ret).cumprod()
        uc_total_ret = float(uc_path.iloc[-1] - 1.0) if not uc_path.empty else 0.0 if not uc_path.empty else 0.0

        total_ret = (float(euro_pct) / 100.0) * euro_total_ret + uc_total * uc_total_ret

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.metric("Rendement annualise UC", fmt_pct_fr(ann_ret_uc * 100.0))
        with k2:
            st.metric("Volatilite annualisee UC", fmt_pct_fr(ann_vol_uc * 100.0))
        with k3:
            st.metric("Sharpe UC", f"{sharpe_uc:.2f}" if sharpe_uc == sharpe_uc else "-")
        with k4:
            st.metric("Rendement total (EUROFUND+UC)", fmt_pct_fr(total_ret * 100.0))

        rfr_display = get_risk_free_rate()
        rfr_src = st.session_state.get("RISK_FREE_RATE_SOURCE", "manual")
        rfr_label = (
            f"📡 Bund 10 ans en temps réel ({rfr_display * 100:.2f}%)"
            if rfr_src == "api"
            else f"✏️ Taux saisi manuellement ({rfr_display * 100:.2f}%)"
        )
        st.caption(
            f"Ratio de Sharpe calculé avec un taux sans risque de "
            f"**{rfr_display * 100:.2f}%** — {rfr_label}"
        )

        st.markdown(
            """
- Allocation calibree sur la fenetre d'analyse et l'objectif choisi.
- EUROFUND est traite uniquement via son taux annuel parametre.
- Les UC respectent un cap strict de 25% du portefeuille final par fonds.
- En mode ancre, le fonds impose est conserve dans la poche actions.
- En cas de donnees insuffisantes, la selection est reduite automatiquement.
"""
        )

        if practical_mode:
            st.caption(
                "Les ponderations UC sont ajustees pour rester coherentes en pratique "
                "(min 10 % et arrondi par paliers de 5 %)."
            )

        if insufficient:
            st.caption("Fonds exclus : " + ", ".join(insufficient))
        st.caption(f"Fenetre utilisee : {fmt_date(opt_start)} a {fmt_date(opt_end)}")

    except Exception as e:
        st.error("Une erreur est survenue dans le builder.")
        st.exception(e)


# ------------------------------------------------------------
# Sélection fonds depuis le référentiel contrat
# ------------------------------------------------------------

_BOND_CATEGORY_PREFIXES = ("OBLIGATIONS",)
_BOND_CATEGORY_KEYWORDS = (
    "CONVERTIBLES",
    "SUBORDINATED BOND",
    "REVENUS FIXES",
)


def _is_bond_category(cat: str) -> bool:
    """
    Retourne True si la catégorie Morningstar est de nature obligataire.
    Couvre :
    - Toutes catégories commençant par "OBLIGATIONS" (ex : OBLIGATIONS EUR...)
    - Convertibles (CONVERTIBLES EUR, CONVERTIBLES EUROPE...)
    - Dettes subordonnées (EUR SUBORDINATED BOND)
    - Revenus fixes (REVENUS FIXES EUR — fonds à échéance)
    - Produits leveragés/inversés sur obligations
      (TRADING - LEVERAGED/INVERSE OBLIGATIONS)
    """
    c = str(cat).upper().strip()
    return (
        c.startswith(_BOND_CATEGORY_PREFIXES)
        or any(k in c for k in _BOND_CATEGORY_KEYWORDS)
        or "OBLIGATIONS" in c
    )


def _get_assureur(contract_label: str) -> str:
    cfg = CONTRACTS_REGISTRY.get(contract_label, {})
    return cfg.get("assureur", "—")


def _add_fund_from_contract(port_key: str, label: str):
    """
    Interface d'ajout de fonds structurée en deux blocs :
      Bloc A — Actifs défensifs (fonds euros + fonds obligataires)
      Bloc B — Autres UC (actions, diversifiés, etc.)
    Utilise le contrat spécifique au portefeuille (A ou B).
    """
    st.subheader(label)

    is_a = port_key == "A_lines"
    contract_label = st.session_state.get(
        "CONTRACT_LABEL_A" if is_a else "CONTRACT_LABEL_B", ""
    )
    contract_cfg = CONTRACTS_REGISTRY.get(contract_label, {})
    funds_df = st.session_state.get(
        "CONTRACT_FUNDS_DF_A" if is_a else "CONTRACT_FUNDS_DF_B", pd.DataFrame()
    )
    buy_date_central = (
        st.session_state.get("INIT_A_DATE", pd.Timestamp("2024-01-02").date())
        if is_a
        else st.session_state.get("INIT_B_DATE", pd.Timestamp("2024-01-02").date())
    )

    if not contract_label or not contract_cfg:
        st.warning("Contrat non sélectionné — configurez-le dans les paramètres.")
        return

    assureur = contract_cfg.get("assureur", "—")
    st.caption(
        f"Contrat : **{contract_label}** ({assureur}) "
        f"— {len(funds_df)} UC disponibles "
        f"| Date d'achat : {pd.Timestamp(buy_date_central).strftime('%d/%m/%Y')}"
    )

    # ═══════════════════════════════════════════════════════════
    # BLOC A — Actifs défensifs
    # ═══════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("### 🛡️ Actifs défensifs")

    # ── Sous-section 1 : Fonds en euros ──────────────────────
    st.markdown("#### Fonds en euros — capital garanti à 98%")

    ef_options = list(contract_cfg.get("euro_funds", {}).keys())
    if ef_options:
        ef_selected = st.selectbox("Fonds en euros", ef_options, key=f"ef_sel_{port_key}")
        ef_filename = contract_cfg["euro_funds"][ef_selected]
        try:
            ef_history = load_euro_fund_history(contract_cfg["path"], ef_filename)
        except Exception:
            ef_history = pd.DataFrame()
        ef_avg = get_euro_fund_avg_rate(ef_history, years=5)
        ef_last = (
            float(ef_history.iloc[-1]["taux_net_publie_pct"])
            if not ef_history.empty else ef_avg
        )
        with st.container(border=True):
            hc1, hc2 = st.columns([3, 1])
            with hc1:
                st.markdown(f"**{ef_selected}**")
                st.caption(
                    f"Assureur : {assureur} | SRI 1 | Capital garanti à 98%"
                    f" | Taux net {int(ef_history['annee'].max()) if not ef_history.empty else '—'}"
                    f" : **{ef_last:.2f}%**"
                )
            with hc2:
                st.metric("Moy. 5 ans", f"{ef_avg:.2f}%")
        euro_amt = st.number_input(
            "Montant (€)",
            min_value=0.0, max_value=10_000_000.0,
            value=0.0, step=1000.0,
            key=f"euro_amt_{port_key}",
        )
        if st.button("➕ Ajouter fonds en euros", key=f"add_euro_{port_key}"):
            if euro_amt <= 0:
                st.warning("Montant invalide.")
            else:
                st.session_state[port_key].append({
                    "id":            str(uuid.uuid4()),
                    "name":          ef_selected,
                    "isin":          "EUROFUND",
                    "amount_gross":  float(euro_amt),
                    "buy_date":      pd.Timestamp(buy_date_central),
                    "buy_px":        "",
                    "note":          f"Contrat : {contract_label}",
                    "sym_used":      "EUROFUND",
                    "fee_total_pct": 0.0,
                })
                st.success(f"{ef_selected} ajouté.")

    # ── Sous-section 2 : Fonds obligataires ──────────────────
    st.markdown("#### Fonds obligataires")

    if funds_df.empty:
        st.info("Référentiel de fonds non chargé.")
    else:
        bond_df = funds_df[funds_df["category"].apply(_is_bond_category)].copy()

        if bond_df.empty:
            st.info("Aucun fonds obligataire dans le référentiel.")
        else:
            bf1, bf2, bf3 = st.columns([3, 2, 1])
            with bf1:
                bond_search = st.text_input(
                    "Rechercher (nom, société de gestion)",
                    value="", key=f"bond_search_{port_key}",
                    placeholder="Ex : AXA, Amundi...",
                )
            with bf2:
                bond_cats = ["Toutes"] + sorted(bond_df["category"].dropna().unique().tolist())
                bond_cat = st.selectbox("Catégorie", bond_cats, key=f"bond_cat_{port_key}")
            with bf3:
                bond_sri = st.selectbox("SRI", ["Tous"] + [str(i) for i in range(1, 8)], key=f"bond_sri_{port_key}")

            if bond_search.strip():
                q = bond_search.strip().upper()
                bond_df = bond_df[
                    bond_df["name"].str.upper().str.contains(q, na=False) |
                    bond_df["manager"].str.upper().str.contains(q, na=False)
                ]
            if bond_cat != "Toutes":
                bond_df = bond_df[bond_df["category"] == bond_cat]
            if bond_sri != "Tous":
                bond_df = bond_df[bond_df["sri"] == int(bond_sri)]

            if bond_df.empty:
                st.info("Aucun résultat.")
            else:
                # Téléchargement liste filtrée
                _bond_xlsx = BytesIO()
                bond_df[["isin", "name", "manager", "category", "sri", "fee_total_pct"]].rename(
                    columns={
                        "isin": "ISIN", "name": "Libellé", "manager": "Société",
                        "category": "Catégorie", "sri": "SRI", "fee_total_pct": "Frais totaux %",
                    }
                ).to_excel(_bond_xlsx, index=False, engine="openpyxl")
                st.download_button(
                    f"⬇️ Télécharger la liste filtrée ({len(bond_df)} fonds)",
                    data=_bond_xlsx.getvalue(),
                    file_name=f"obligataires_{contract_label.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"bond_dl_{port_key}",
                )
                # Sélection directe par selectbox
                bond_options = [
                    f"{row['name']}  |  {row['manager']}  |  SRI {row['sri']}  |  {row['fee_total_pct']:.2f}%"
                    for _, row in bond_df.iterrows()
                ]
                bond_sel_idx = st.selectbox(
                    f"{len(bond_df)} fonds obligataires disponibles — sélectionner",
                    options=range(len(bond_options)),
                    format_func=lambda i: bond_options[i],
                    key=f"bond_sel_{port_key}",
                )
                bond_row = bond_df.iloc[bond_sel_idx]
                with st.container(border=True):
                    bc1, bc2 = st.columns([3, 1])
                    with bc1:
                        st.markdown(f"**{bond_row['name']}**")
                        st.caption(
                            f"{bond_row['manager']}  ·  {bond_row['category']}"
                            f"  ·  ISIN : `{bond_row['isin']}`"
                        )
                    with bc2:
                        st.metric("Frais totaux", f"{bond_row['fee_total_pct']:.2f}%")
                bb1, bb2, bb3 = st.columns([2, 2, 1])
                with bb1:
                    bond_amt = st.text_input(
                        "Montant (€)", value="", key=f"bond_amt_{port_key}",
                        placeholder="Ex : 10000",
                    )
                with bb2:
                    bond_px = st.text_input(
                        "Prix d'achat (optionnel)", value="", key=f"bond_px_{port_key}",
                    )
                with bb3:
                    if st.button("➕ Ajouter", key=f"bond_add_{port_key}"):
                        try:
                            amt = float(str(bond_amt).replace(" ", "").replace(",", "."))
                            assert amt > 0
                        except Exception:
                            st.warning("Montant invalide.")
                        else:
                            st.session_state[port_key].append({
                                "id":               str(uuid.uuid4()),
                                "name":             bond_row["name"],
                                "isin":             bond_row["isin"],
                                "amount_gross":     float(amt),
                                "buy_date":         pd.Timestamp(buy_date_central),
                                "buy_px":           float(str(bond_px).replace(",", ".")) if bond_px.strip() else "",
                                "note":             f"Contrat : {contract_label} | SRI {bond_row['sri']}",
                                "sym_used":         "",
                                "fee_uc_pct":       bond_row["fee_uc_pct"],
                                "fee_contract_pct": bond_row["fee_contract_pct"],
                                "fee_total_pct":    bond_row["fee_total_pct"],
                            })
                            st.success(f"{bond_row['name']} ajouté.")

    # ═══════════════════════════════════════════════════════════
    # BLOC B — Autres UC
    # ═══════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("### 📈 Autres UC — Actions, diversifiés, etc.")

    if funds_df.empty:
        st.info("Référentiel de fonds non chargé.")
        return

    other_df = funds_df[~funds_df["category"].apply(_is_bond_category)].copy()

    if other_df.empty:
        st.info("Aucun fonds UC dans le référentiel.")
        return

    oc1, oc2, oc3 = st.columns([3, 2, 1])
    with oc1:
        other_search = st.text_input(
            "Rechercher (nom, ISIN)",
            value="", key=f"other_search_{port_key}",
            placeholder="Ex : Carmignac, FR0010135103...",
        )
    with oc2:
        other_cats = ["Toutes"] + sorted(other_df["category"].dropna().unique().tolist())
        other_cat = st.selectbox("Catégorie", other_cats, key=f"other_cat_{port_key}")
    with oc3:
        other_sri = st.selectbox("SRI", ["Tous"] + [str(i) for i in range(1, 8)], key=f"other_sri_{port_key}")

    if other_search.strip():
        q = other_search.strip().upper()
        other_df = other_df[
            other_df["isin"].str.upper().str.contains(q, na=False) |
            other_df["name"].str.upper().str.contains(q, na=False)
        ]
    if other_cat != "Toutes":
        other_df = other_df[other_df["category"] == other_cat]
    if other_sri != "Tous":
        other_df = other_df[other_df["sri"] == int(other_sri)]

    if other_df.empty:
        st.info("Aucun fonds ne correspond aux critères.")
        return

    # Téléchargement liste filtrée
    _other_xlsx = BytesIO()
    other_df[["isin", "name", "manager", "category", "sri", "fee_total_pct"]].rename(
        columns={
            "isin": "ISIN", "name": "Libellé", "manager": "Société",
            "category": "Catégorie", "sri": "SRI", "fee_total_pct": "Frais totaux %",
        }
    ).to_excel(_other_xlsx, index=False, engine="openpyxl")
    st.download_button(
        f"⬇️ Télécharger la liste filtrée ({len(other_df)} UC)",
        data=_other_xlsx.getvalue(),
        file_name=f"uc_{contract_label.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"other_dl_{port_key}",
    )
    # Sélection directe par selectbox
    other_options = [
        f"{row['name']}  |  {row['manager']}  |  SRI {row['sri']}  |  {row['fee_total_pct']:.2f}%"
        for _, row in other_df.iterrows()
    ]
    other_sel_idx = st.selectbox(
        f"{len(other_df)} UC disponibles — sélectionner",
        options=range(len(other_options)),
        format_func=lambda i: other_options[i],
        key=f"other_sel_{port_key}",
    )
    other_row = other_df.iloc[other_sel_idx]
    with st.container(border=True):
        oc1, oc2 = st.columns([3, 1])
        with oc1:
            st.markdown(f"**{other_row['name']}**")
            st.caption(
                f"{other_row['manager']}  ·  {other_row['category']}"
                f"  ·  ISIN : `{other_row['isin']}`"
            )
        with oc2:
            st.metric("Frais totaux", f"{other_row['fee_total_pct']:.2f}%")
    ob1, ob2, ob3 = st.columns([2, 2, 1])
    with ob1:
        other_amt = st.text_input(
            "Montant (€)", value="", key=f"other_amt_{port_key}",
            placeholder="Ex : 10000",
        )
    with ob2:
        other_px = st.text_input(
            "Prix d'achat (optionnel)", value="", key=f"other_px_{port_key}",
        )
    with ob3:
        if st.button("➕ Ajouter", key=f"other_add_{port_key}"):
            try:
                amt = float(str(other_amt).replace(" ", "").replace(",", "."))
                assert amt > 0
            except Exception:
                st.warning("Montant invalide.")
            else:
                st.session_state[port_key].append({
                    "id":               str(uuid.uuid4()),
                    "name":             other_row["name"],
                    "isin":             other_row["isin"],
                    "amount_gross":     float(amt),
                    "buy_date":         pd.Timestamp(buy_date_central),
                    "buy_px":           float(str(other_px).replace(",", ".")) if other_px.strip() else "",
                    "note":             f"Contrat : {contract_label} | SRI {other_row['sri']}",
                    "sym_used":         "",
                    "fee_uc_pct":       other_row["fee_uc_pct"],
                    "fee_contract_pct": other_row["fee_contract_pct"],
                    "fee_total_pct":    other_row["fee_total_pct"],
                })
                st.success(f"{other_row['name']} ajouté.")


def render_app(run_page_config: bool = True):
    # ------------------------------------------------------------
    # Layout principal
    # ------------------------------------------------------------
    if run_page_config:
        st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.info(f"App chargée, statut {st.session_state.get('APP_STATUS', 'OK')}")
    # Init state
    st.session_state.setdefault("A_lines", [])
    st.session_state.setdefault("B_lines", [])
    st.session_state.setdefault("FEE_A", 3.0)
    st.session_state.setdefault("FEE_B", 2.0)
    st.session_state.setdefault("M_A", 0.0)
    st.session_state.setdefault("M_B", 0.0)
    st.session_state.setdefault("ONE_A", 0.0)
    st.session_state.setdefault("ONE_B", 0.0)
    st.session_state.setdefault("ONE_A_DATE", pd.Timestamp("2024-07-01").date())
    st.session_state.setdefault("ONE_B_DATE", pd.Timestamp("2024-07-01").date())
    st.session_state.setdefault("ALLOC_MODE", "equal")
    st.session_state.setdefault("MANUAL_NAV_STORE", {})
    st.session_state.setdefault("DATE_WARNINGS", [])
    st.session_state.setdefault("INIT_A_DATE", pd.Timestamp("2024-01-02").date())
    st.session_state.setdefault("INIT_B_DATE", pd.Timestamp("2024-01-02").date())
    st.session_state.setdefault("EURO_RATE_A", 2.0)
    st.session_state.setdefault("EURO_RATE_B", 2.5)

    # -------------------------------------------------------------------
    # Sidebar : paramètres globaux
    # -------------------------------------------------------------------
    with st.sidebar:
        # ── Contrat client ─────────────────────────────────────────
        st.header("Contrat client")
        contract_label_A = st.selectbox(
            "Contrat client",
            list(CONTRACTS_REGISTRY.keys()),
            key="CONTRACT_LABEL_A",
        )
        contract_cfg_A = CONTRACTS_REGISTRY[contract_label_A]
        try:
            funds_df_A = load_contract_funds(
                contract_cfg_A["path"], contract_cfg_A["funds_filename"]
            )
        except Exception:
            funds_df_A = pd.DataFrame()
        st.session_state["CONTRACT_FUNDS_DF_A"] = funds_df_A

        euro_fund_label_A = st.selectbox(
            "Fonds en euros client",
            list(contract_cfg_A["euro_funds"].keys()),
            key="EURO_FUND_LABEL_A",
        )
        try:
            euro_history_df_A = load_euro_fund_history(
                contract_cfg_A["path"], contract_cfg_A["euro_funds"][euro_fund_label_A]
            )
        except Exception:
            euro_history_df_A = pd.DataFrame()
        avg_rate_A = get_euro_fund_avg_rate(euro_history_df_A, years=5)
        with st.expander(f"Historique {euro_fund_label_A}", expanded=False):
            if not euro_history_df_A.empty:
                st.dataframe(
                    euro_history_df_A.rename(columns={
                        "annee": "Année", "taux_net_publie_pct": "Taux net publié %"
                    }),
                    hide_index=True, use_container_width=True,
                )
            st.caption(f"Moyenne sur 5 dernières années : **{avg_rate_A:.2f}%**")

        st.markdown("---")

        # ── Contrat cabinet ────────────────────────────────────────
        st.header("Contrat cabinet")
        contract_label_B = st.selectbox(
            "Contrat cabinet",
            list(CONTRACTS_REGISTRY.keys()),
            key="CONTRACT_LABEL_B",
        )
        contract_cfg_B = CONTRACTS_REGISTRY[contract_label_B]
        try:
            funds_df_B = load_contract_funds(
                contract_cfg_B["path"], contract_cfg_B["funds_filename"]
            )
        except Exception:
            funds_df_B = pd.DataFrame()
        st.session_state["CONTRACT_FUNDS_DF_B"] = funds_df_B

        euro_fund_label_B = st.selectbox(
            "Fonds en euros cabinet",
            list(contract_cfg_B["euro_funds"].keys()),
            key="EURO_FUND_LABEL_B",
        )
        try:
            euro_history_df_B = load_euro_fund_history(
                contract_cfg_B["path"], contract_cfg_B["euro_funds"][euro_fund_label_B]
            )
        except Exception:
            euro_history_df_B = pd.DataFrame()
        avg_rate_B = get_euro_fund_avg_rate(euro_history_df_B, years=5)
        with st.expander(f"Historique {euro_fund_label_B}", expanded=False):
            if not euro_history_df_B.empty:
                st.dataframe(
                    euro_history_df_B.rename(columns={
                        "annee": "Année", "taux_net_publie_pct": "Taux net publié %"
                    }),
                    hide_index=True, use_container_width=True,
                )
            st.caption(f"Moyenne sur 5 dernières années : **{avg_rate_B:.2f}%**")

        st.markdown("---")

        # Rétrocompatibilité : render_portfolio_builder() lit CONTRACT_LABEL / CONTRACT_FUNDS_DF
        st.session_state["CONTRACT_LABEL"]    = contract_label_A
        st.session_state["CONTRACT_CFG"]      = contract_cfg_A
        st.session_state["CONTRACT_FUNDS_DF"] = funds_df_A
        st.session_state["EURO_FUND_LABEL"]   = euro_fund_label_A
        st.session_state["EURO_FUND_HISTORY"] = euro_history_df_A
        st.session_state["EURO_FUND_AVG_RATE"] = avg_rate_A

        # ── Taux fonds en euros ───────────────────────────────────
        st.header("Taux fonds en euros")
        st.caption(
            f"Client : **{avg_rate_A:.2f}%** ({euro_fund_label_A}) | "
            f"Cabinet : **{avg_rate_B:.2f}%** ({euro_fund_label_B})"
        )
        override_euro = st.checkbox(
            "Utiliser un taux personnalisé (fonds euros boosté, etc.)",
            value=False,
            key="OVERRIDE_EURO_RATE",
        )
        if override_euro:
            euro_rate_value_A = st.number_input(
                "Portefeuille Client — taux annuel (%)",
                min_value=0.0, max_value=10.0,
                value=float(st.session_state.get("EURO_RATE_A", avg_rate_A)),
                step=0.10, key="EURO_RATE_A",
            )
            euro_rate_value_B = st.number_input(
                "Portefeuille Cabinet — taux annuel (%)",
                min_value=0.0, max_value=10.0,
                value=float(st.session_state.get("EURO_RATE_B", avg_rate_B)),
                step=0.10, key="EURO_RATE_B",
            )
        else:
            euro_rate_value_A = avg_rate_A
            euro_rate_value_B = avg_rate_B
            st.session_state["EURO_RATE_A"] = avg_rate_A
            st.session_state["EURO_RATE_B"] = avg_rate_B
            if avg_rate_A == avg_rate_B:
                st.caption(f"Taux appliqué aux deux portefeuilles : **{avg_rate_A:.2f}%**")
            else:
                st.caption(
                    f"Taux client : **{avg_rate_A:.2f}%** | Taux cabinet : **{avg_rate_B:.2f}%**"
                )
        st.caption(
            "Frais de gestion contrat fonds euros inclus dans le taux net publié. "
            "Frais UC contrat : lus depuis le référentiel fonds."
        )

        # Frais d’entrée
        st.header("Frais d’entrée (%)")

        FEE_A = st.number_input(
            "Frais d’entrée — Portefeuille 1 (Client)",
            0.0,
            10.0,
            st.session_state.get("FEE_A", 3.0),
            0.10,
            key="FEE_A",
        )

        FEE_B = st.number_input(
            "Frais d’entrée — Portefeuille 2 (Cabinet)",
            0.0,
            10.0,
            st.session_state.get("FEE_B", 2.0),
            0.10,
            key="FEE_B",
        )

        st.caption("Les frais s’appliquent sur chaque investissement (initial, mensuel, ponctuel).")

        # Date du versement initial (centralisée)
        st.header("Date du versement initial")

        st.date_input(
            "Portefeuille 1 (Client) — date d’investissement initiale",
            value=st.session_state.get("INIT_A_DATE", pd.Timestamp("2024-01-02").date()),
            key="INIT_A_DATE",
        )

        st.date_input(
            "Portefeuille 2 (Cabinet) — date d’investissement initiale",
            value=st.session_state.get("INIT_B_DATE", pd.Timestamp("2024-01-02").date()),
            key="INIT_B_DATE",
        )

        # Paramètres de versement
        st.header("Paramètres de versement")

        with st.expander("Portefeuille 1 — Client"):
            M_A = st.number_input(
                "Mensuel brut (€)",
                0.0,
                1_000_000.0,
                st.session_state.get("M_A", 0.0),
                100.0,
                key="M_A",
            )
            ONE_A = st.number_input(
                "Ponctuel brut (€)",
                0.0,
                1_000_000.0,
                st.session_state.get("ONE_A", 0.0),
                100.0,
                key="ONE_A",
            )
            ONE_A_DATE = st.date_input(
                "Date du ponctuel",
                value=st.session_state.get("ONE_A_DATE", pd.Timestamp("2024-07-01").date()),
                key="ONE_A_DATE",
            )

        with st.expander("Portefeuille 2 — Cabinet"):
            M_B = st.number_input(
                "Mensuel brut (€)",
                0.0,
                1_000_000.0,
                st.session_state.get("M_B", 0.0),
                100.0,
                key="M_B",
            )
            ONE_B = st.number_input(
                "Ponctuel brut (€)",
                0.0,
                1_000_000.0,
                st.session_state.get("ONE_B", 0.0),
                100.0,
                key="ONE_B",
            )
            ONE_B_DATE = st.date_input(
                "Date du ponctuel",
                value=st.session_state.get("ONE_B_DATE", pd.Timestamp("2024-07-01").date()),
                key="ONE_B_DATE",
            )

        # Règle d’affectation
        st.header("Règle d’affectation des versements")

        current_code = st.session_state.get("ALLOC_MODE", "equal")
        inv_labels = {v: k for k, v in ALLOC_LABELS.items()}
        current_label = inv_labels.get(current_code, "Répartition égale")

        mode_label = st.selectbox(
            "Mode",
            list(ALLOC_LABELS.keys()),
            index=list(ALLOC_LABELS.keys()).index(current_label),
            help="Répartition des versements entre les lignes.",
        )

        st.session_state["ALLOC_MODE"] = ALLOC_LABELS[mode_label]

        st.divider()
        st.header("Mode d’analyse")

        mode_ui = st.radio(
            "Choix",
            ["Comparer Client vs Cabinet", "Analyser uniquement Cabinet", "Analyser uniquement Client"],
            index=0,
            key="MODE_ANALYSE_UI",
        )

        if "Comparer" in mode_ui:
            st.session_state["MODE_ANALYSE"] = "compare"
        elif "Cabinet" in mode_ui:
            st.session_state["MODE_ANALYSE"] = "valority"
        else:
            st.session_state["MODE_ANALYSE"] = "client"

        st.divider()
        debug_mode = st.checkbox("Mode debug", value=False)
        if debug_mode:
            st.subheader("Debug")
            st.caption("Versions & état")
            st.code(
                f"Python: {sys.version.split()[0]}\n"
                f"Streamlit: {st.__version__}\n"
                f"Pandas: {pd.__version__}"
            )
            st.caption("Modules")
            st.code(
                f"Matplotlib: {MATPLOTLIB_AVAILABLE} ({MATPLOTLIB_ERROR})\n"
                f"Reportlab: {REPORTLAB_AVAILABLE} ({REPORTLAB_ERROR})"
            )
            st.caption("Session state (clés)")
            st.write(sorted(list(st.session_state.keys())))
            st.caption("Dernière exception")
            st.write(st.session_state.get("LAST_EXCEPTION", "—"))
            st.caption("Test rapide EODHD")
            if _get_api_key():
                try:
                    res = eodhd_get("/status")
                    st.write(res if res is not None else "Réponse vide")
                except Exception as e:
                    st.write(f"Erreur EODHD: {e}")
            else:
                st.write("Token EODHD absent")


    mode = st.session_state.get("MODE_ANALYSE", "compare")
    show_client = mode in ("compare", "client")
    show_valority = mode in ("compare", "valority")

    # Onglets principaux : Client / Cabinet (conditionnés au mode)
    tab_labels = []
    if show_client:
        tab_labels.append("Portefeuille Client")
    if show_valority:
        tab_labels.append("Portefeuille Cabinet")
    tabs = st.tabs(tab_labels) if tab_labels else []

    idx = 0
    if show_client:
        with tabs[idx]:
            _add_fund_from_contract("A_lines", "Ajouter un fonds — Portefeuille Client")
            st.markdown("#### Lignes actuelles — Portefeuille Client")
            for i, ln in enumerate(st.session_state.get("A_lines", [])):
                _line_card(ln, i, "A_lines")
        idx += 1

    if show_valority:
        with tabs[idx]:
            _add_fund_from_contract("B_lines", "Ajouter un fonds — Portefeuille Cabinet")
            st.markdown("#### Lignes actuelles — Portefeuille Cabinet")
            for i, ln in enumerate(st.session_state.get("B_lines", [])):
                _line_card(ln, i, "B_lines")

    # ------------------------------------------------------------
    # Simulation (selon mode)
    # ------------------------------------------------------------
    mode = st.session_state.get("MODE_ANALYSE", "compare")
    show_client = mode in ("compare", "client")
    show_valority = mode in ("compare", "valority")

    # FIXED: use stable UUID key instead of id() which changes on st.rerun() (Bug 5)
    _ln_a0 = st.session_state["A_lines"][0] if (show_client and st.session_state["A_lines"]) else None
    single_target_A = (_ln_a0.get("id") or id(_ln_a0)) if _ln_a0 is not None else None
    _ln_b0 = st.session_state["B_lines"][0] if (show_valority and st.session_state["B_lines"]) else None
    single_target_B = (_ln_b0.get("id") or id(_ln_b0)) if _ln_b0 is not None else None

    alloc_mode_code = st.session_state.get("ALLOC_MODE", "equal")

    custom_month_weights_A: Optional[Dict[int, float]] = None
    custom_oneoff_weights_A: Optional[Dict[int, float]] = None
    custom_month_weights_B: Optional[Dict[int, float]] = None
    custom_oneoff_weights_B: Optional[Dict[int, float]] = None

    if alloc_mode_code == "custom":
        if show_client:
            cmA = st.session_state.get("CUSTOM_M_A", {}) or {}
            coA = st.session_state.get("CUSTOM_O_A", {}) or {}
            tot_mA = sum(v for v in cmA.values() if v > 0)
            tot_oA = sum(v for v in coA.values() if v > 0)
            if tot_mA > 0:
                custom_month_weights_A = {k: v / tot_mA for k, v in cmA.items() if v > 0}
            if tot_oA > 0:
                custom_oneoff_weights_A = {k: v / tot_oA for k, v in coA.items() if v > 0}

        if show_valority:
            cmB = st.session_state.get("CUSTOM_M_B", {}) or {}
            coB = st.session_state.get("CUSTOM_O_B", {}) or {}
            tot_mB = sum(v for v in cmB.values() if v > 0)
            tot_oB = sum(v for v in coB.values() if v > 0)
            if tot_mB > 0:
                custom_month_weights_B = {k: v / tot_mB for k, v in cmB.items() if v > 0}
            if tot_oB > 0:
                custom_oneoff_weights_B = {k: v / tot_oB for k, v in coB.items() if v > 0}

    # Reset warnings avant chaque run
    st.session_state["DATE_WARNINGS"] = []

    # Valeurs par défaut (si on ne simule pas un des portefeuilles)
    dfA, brutA, netA, valA, xirrA, startA_min, fullA = pd.DataFrame(), 0.0, 0.0, 0.0, None, TODAY, TODAY
    dfB, brutB, netB, valB, xirrB, startB_min, fullB = pd.DataFrame(), 0.0, 0.0, 0.0, None, TODAY, TODAY

    # ── Synchronisation date globale → lignes (pour cohérence carte + tableau + simulation) ──
    _global_date_A = pd.Timestamp(st.session_state.get("INIT_A_DATE", pd.Timestamp("2024-01-02").date()))
    _global_date_B = pd.Timestamp(st.session_state.get("INIT_B_DATE", pd.Timestamp("2024-01-02").date()))
    for _ln in st.session_state.get("A_lines", []):
        if not _ln.get("date_overridden"):
            _ln["buy_date"] = _global_date_A
    for _ln in st.session_state.get("B_lines", []):
        if not _ln.get("date_overridden"):
            _ln["buy_date"] = _global_date_B

    if show_client:
        _args_A = _make_sim_args(
            lines=st.session_state.get("A_lines", []),
            monthly_amt=st.session_state.get("M_A", 0.0),
            one_amt=st.session_state.get("ONE_A", 0.0),
            one_date=st.session_state.get("ONE_A_DATE", pd.Timestamp("2024-07-01").date()),
            alloc_mode=alloc_mode_code,
            custom_weights_monthly=custom_month_weights_A,
            custom_weights_oneoff=custom_oneoff_weights_A,
            single_target=single_target_A,
            euro_rate=st.session_state.get("EURO_RATE_A", 2.0),
            fee_pct=st.session_state.get("FEE_A", 0.0),
            label="Client",
        )
        dfA, brutA, netA, valA, xirrA, startA_min, fullA = _simulate_portfolio_cached(**_args_A)

    if show_valority:
        _args_B = _make_sim_args(
            lines=st.session_state.get("B_lines", []),
            monthly_amt=st.session_state.get("M_B", 0.0),
            one_amt=st.session_state.get("ONE_B", 0.0),
            one_date=st.session_state.get("ONE_B_DATE", pd.Timestamp("2024-07-01").date()),
            alloc_mode=alloc_mode_code,
            custom_weights_monthly=custom_month_weights_B,
            custom_weights_oneoff=custom_oneoff_weights_B,
            single_target=single_target_B,
            euro_rate=st.session_state.get("EURO_RATE_B", 2.5),
            fee_pct=st.session_state.get("FEE_B", 0.0),
            label="Cabinet",
        )
        dfB, brutB, netB, valB, xirrB, startB_min, fullB = _simulate_portfolio_cached(**_args_B)

    # ------------------------------------------------------------
    # Avertissements sur les dates / 1ère VL
    # ------------------------------------------------------------
    if st.session_state.get("DATE_WARNINGS"):
        with st.expander("⚠️ Problèmes d'historique / dates de VL"):
            for msg in st.session_state["DATE_WARNINGS"]:
                st.warning(msg)

    # ------------------------------------------------------------
    # Graphique (évolution des portefeuilles)
    # ------------------------------------------------------------
    st.subheader("Évolution de la valeur des portefeuilles")

    # Déterminer le start_plot uniquement sur les portefeuilles affichés
    full_dates: List[pd.Timestamp] = []
    if show_client and isinstance(fullA, pd.Timestamp):
        full_dates.append(fullA)
    if show_valority and isinstance(fullB, pd.Timestamp):
        full_dates.append(fullB)

    start_plot = max(full_dates) if full_dates else TODAY

    idx = pd.bdate_range(start=start_plot, end=TODAY, freq="B")
    chart_df = pd.DataFrame(index=idx)

    if show_client and not dfA.empty:
        chart_df["Client"] = dfA.reindex(idx)["Valeur"].ffill()

    if show_valority and not dfB.empty:
        chart_df["Cabinet"] = dfB.reindex(idx)["Valeur"].ffill()

    # Passage en format long pour Altair
    chart_long = chart_df.reset_index().rename(columns={"index": "Date"})
    chart_long = chart_long.melt("Date", var_name="variable", value_name="Valeur (€)")

    if chart_long.dropna().empty:
        st.info("Ajoutez des lignes et/ou vérifiez vos paramètres pour afficher le graphique.")
    else:
        # Calcul du domaine Y dynamique avec marges
        valid_vals = chart_long["Valeur (€)"].dropna()
        if not valid_vals.empty:
            y_min = float(valid_vals.min())
            y_max = float(valid_vals.max())
            padding = (y_max - y_min) * 0.05 if y_max > y_min else y_max * 0.05
            y_domain = [max(0.0, y_min - padding), y_max + padding]
        else:
            y_domain = [0, 1]
        base = (
            alt.Chart(chart_long)
            .mark_line()
            .encode(
                x=alt.X("Date:T", title="Date"),
                y=alt.Y(
                    "Valeur (€):Q",
                    title="Valeur (€)",
                    scale=alt.Scale(domain=y_domain),
                    axis=alt.Axis(format=",.0f"),
                ),
                color=alt.Color("variable:N", title="Portefeuille"),
                tooltip=[
                    alt.Tooltip("Date:T", title="Date"),
                    alt.Tooltip("variable:N", title="Portefeuille"),
                    alt.Tooltip("Valeur (€):Q", title="Valeur", format=",.2f"),
                ],
            )
            .properties(height=360, width="container")
        )
        st.altair_chart(base, use_container_width=True)

    # ------------------------------------------------------------
    # Synthèse chiffrée : cartes Client / Cabinet
    # ------------------------------------------------------------
    st.subheader("Synthèse chiffrée")

    mode = st.session_state.get("MODE_ANALYSE", "compare")

    # Période analysée (uniquement sur ce qui est affiché)
    period_dates: List[pd.Timestamp] = []
    if mode in ("compare", "client") and isinstance(startA_min, pd.Timestamp):
        period_dates.append(startA_min)
    if mode in ("compare", "valority") and isinstance(startB_min, pd.Timestamp):
        period_dates.append(startB_min)

    if period_dates:
        start_global = min(period_dates)
        st.caption(f"Période analysée : du **{fmt_date(start_global)}** au **{fmt_date(TODAY)}**")

    perf_tot_client = (valA / netA - 1.0) * 100.0 if (show_client and netA > 0) else None
    perf_tot_valority = (valB / netB - 1.0) * 100.0 if (show_valority and netB > 0) else None

    # ✅ 2 colonnes si compare, sinon 1 colonne (container)
    if mode == "compare":
        col_client, col_valority = st.columns(2)
    else:
        col_client = st.container()
        col_valority = st.container()

    # ----- Carte Client -----
    if mode in ("compare", "client"):
        with col_client:
            with st.container(border=True):
                st.markdown("#### 🧍 Situation actuelle — Client")
                st.metric("Valeur actuelle", to_eur(valA))
                st.markdown(
                    f"""
- Montants réellement investis (après frais) : **{to_eur(netA)}**
- Montants versés (brut) : {to_eur(brutA)}
- Rendement total depuis le début : **{perf_tot_client:.2f}%**
"""
                    if perf_tot_client is not None
                    else f"""
- Montants réellement investis (après frais) : **{to_eur(netA)}**
- Montants versés (brut) : {to_eur(brutA)}
- Rendement total depuis le début : **—**
"""
                )
                st.markdown(
                    f"- Rendement annualisé (XIRR) : **{xirrA:.2f}%**"
                    if xirrA is not None
                    else "- Rendement annualisé (XIRR) : **—**"
                )
                _fee_detail_A = []
                for _ln in st.session_state.get("A_lines", []):
                    _isin = str(_ln.get("isin", "")).upper()
                    if _isin in ("EUROFUND", "STRUCTURED"):
                        continue
                    _name = _ln.get("name") or _isin
                    _fc = _ln.get("fee_contract_pct")
                    _fb = _ln.get("fee_uc_pct")
                    if _fc is not None and _fb is not None:
                        try:
                            _fee_detail_A.append(
                                f"{_name[:30]} : frais contrat {float(_fc):.2f}%/an "
                                f"(+ TER {float(_fb):.2f}%/an intégré dans la VL)"
                            )
                        except Exception:
                            pass
                _fee_note_A = (
                    "ℹ️ **Performance nette des frais de gestion du contrat assureur** "
                    f"({st.session_state.get('CONTRACT_LABEL_A', 'contrat')}). "
                    "Les frais de gestion internes des fonds (TER) sont déjà intégrés "
                    "dans les valeurs liquidatives publiées et ne sont pas déduits en plus. "
                    "À titre de comparaison, Quantalys et Morningstar affichent les "
                    "performances brutes de frais contrat."
                )
                with st.expander("ℹ️ Détail du calcul de performance", expanded=False):
                    st.markdown(_fee_note_A)
                    if _fee_detail_A:
                        st.markdown("**Frais appliqués par fonds :**")
                        for _d in _fee_detail_A:
                            st.markdown(f"- {_d}")
                    st.markdown(
                        "**Formule :** XIRR calculé sur les flux réels "
                        "(versements initiaux, mensuels, ponctuels) et la valeur "
                        "liquidative nette à ce jour."
                    )


    # ----- Carte Cabinet -----
    if mode in ("compare", "valority"):
        with col_valority:
            with st.container(border=True):
                st.markdown("#### 🏢 Simulation — Allocation Cabinet")
                st.metric("Valeur actuelle simulée", to_eur(valB))
                st.markdown(
                    f"""
- Montants réellement investis (après frais) : **{to_eur(netB)}**
- Montants versés (brut) : {to_eur(brutB)}
- Rendement total depuis le début : **{perf_tot_valority:.2f}%**
"""
                    if perf_tot_valority is not None
                    else f"""
- Montants réellement investis (après frais) : **{to_eur(netB)}**
- Montants versés (brut) : {to_eur(brutB)}
- Rendement total depuis le début : **—**
"""
                )
                st.markdown(
                    f"- Rendement annualisé (XIRR) : **{xirrB:.2f}%**"
                    if xirrB is not None
                    else "- Rendement annualisé (XIRR) : **—**"
                )
                _fee_detail_B = []
                for _ln in st.session_state.get("B_lines", []):
                    _isin = str(_ln.get("isin", "")).upper()
                    if _isin in ("EUROFUND", "STRUCTURED"):
                        continue
                    _name = _ln.get("name") or _isin
                    _fc = _ln.get("fee_contract_pct")
                    _fb = _ln.get("fee_uc_pct")
                    if _fc is not None and _fb is not None:
                        try:
                            _fee_detail_B.append(
                                f"{_name[:30]} : frais contrat {float(_fc):.2f}%/an "
                                f"(+ TER {float(_fb):.2f}%/an intégré dans la VL)"
                            )
                        except Exception:
                            pass
                _fee_note_B = (
                    "ℹ️ **Performance nette des frais de gestion du contrat assureur** "
                    f"({st.session_state.get('CONTRACT_LABEL_B', 'contrat')}). "
                    "Les frais de gestion internes des fonds (TER) sont déjà intégrés "
                    "dans les valeurs liquidatives publiées et ne sont pas déduits en plus. "
                    "À titre de comparaison, Quantalys et Morningstar affichent les "
                    "performances brutes de frais contrat."
                )
                with st.expander("ℹ️ Détail du calcul de performance", expanded=False):
                    st.markdown(_fee_note_B)
                    if _fee_detail_B:
                        st.markdown("**Frais appliqués par fonds :**")
                        for _d in _fee_detail_B:
                            st.markdown(f"- {_d}")
                    st.markdown(
                        "**Formule :** XIRR calculé sur les flux réels "
                        "(versements initiaux, mensuels, ponctuels) et la valeur "
                        "liquidative nette à ce jour."
                    )


    def build_html_report(report: Dict[str, Any]) -> str:
        """
        Construit un rapport HTML exportable pour le client.
        Le contenu repose sur 'report', préparé plus bas dans le code.
        """
        as_of = report.get("as_of", "")
        mode = report.get("mode", "compare")
        synthA = report.get("client_summary", {})
        synthB = report.get("valority_summary", {})
        comp = report.get("comparison", {})

        dfA_lines = report.get("df_client_lines")
        dfB_lines = report.get("df_valority_lines")
        dfA_val = report.get("dfA_val")
        dfB_val = report.get("dfB_val")

        def _fmt_eur(x):
            try:
                return f"{x:,.2f} €".replace(",", " ").replace(".", ",")
            except Exception:
                return str(x)

        # Tables en HTML
        html_client_lines = dfA_lines.to_html(index=False, border=0, justify="left") if dfA_lines is not None else ""
        html_valority_lines = dfB_lines.to_html(index=False, border=0, justify="left") if dfB_lines is not None else ""

        if dfA_val is not None:
            html_A_val = dfA_val.to_html(index=False, border=0, justify="left")
        else:
            html_A_val = ""

        if dfB_val is not None:
            html_B_val = dfB_val.to_html(index=False, border=0, justify="left")
        else:
            html_B_val = ""

        if mode == "compare":
            synth_html = f"""
<div class="block">
  <h3>Situation actuelle — Client</h3>
  <ul>
    <li>Valeur actuelle : <b>{_fmt_eur(synthA.get("val", 0))}</b></li>
    <li>Montants réellement investis (net) : {_fmt_eur(synthA.get("net", 0))}</li>
    <li>Montants versés (brut) : {_fmt_eur(synthA.get("brut", 0))}</li>
    <li>Rendement total depuis le début : <b>{synthA.get("perf_tot_pct", 0):.2f} %</b></li>
    <li>Rendement annualisé (XIRR) : <b>{synthA.get("irr_pct", 0):.2f} %</b></li>
  </ul>
</div>

<div class="block">
  <h3>Simulation — Allocation Cabinet</h3>
  <ul>
    <li>Valeur actuelle simulée : <b>{_fmt_eur(synthB.get("val", 0))}</b></li>
    <li>Montants réellement investis (net) : {_fmt_eur(synthB.get("net", 0))}</li>
    <li>Montants versés (brut) : {_fmt_eur(synthB.get("brut", 0))}</li>
    <li>Rendement total depuis le début : <b>{synthB.get("perf_tot_pct", 0):.2f} %</b></li>
    <li>Rendement annualisé (XIRR) : <b>{synthB.get("irr_pct", 0):.2f} %</b></li>
  </ul>
</div>

<div class="block">
  <h3>Comparaison Client vs Cabinet</h3>
  <ul>
    <li>Différence de valeur finale : <b>{_fmt_eur(comp.get("delta_val", 0))}</b></li>
    <li>Écart de performance totale (Cabinet – Client) :
        <b>{comp.get("delta_perf_pct", 0):.2f} %</b></li>
  </ul>
</div>
"""
            detail_html = f"""
<h2>2. Détail des lignes</h2>

<h3>Portefeuille Client</h3>
{html_client_lines}

<h3>Portefeuille Cabinet</h3>
{html_valority_lines}
"""
            hist_html = f"""
<h2>3. Historique de la valeur des portefeuilles</h2>

<h3>Client – Valeur du portefeuille par date</h3>
{html_A_val}

<h3>Cabinet – Valeur du portefeuille par date</h3>
{html_B_val}
"""
        elif mode == "valority":
            synth_html = f"""
<div class="block">
  <h3>Simulation — Allocation Cabinet</h3>
  <ul>
    <li>Valeur actuelle simulée : <b>{_fmt_eur(synthB.get("val", 0))}</b></li>
    <li>Montants réellement investis (net) : {_fmt_eur(synthB.get("net", 0))}</li>
    <li>Montants versés (brut) : {_fmt_eur(synthB.get("brut", 0))}</li>
    <li>Rendement total depuis le début : <b>{synthB.get("perf_tot_pct", 0):.2f} %</b></li>
    <li>Rendement annualisé (XIRR) : <b>{synthB.get("irr_pct", 0):.2f} %</b></li>
  </ul>
</div>
"""
            detail_html = f"""
<h2>2. Détail des lignes</h2>

<h3>Portefeuille Cabinet</h3>
{html_valority_lines}
"""
            hist_html = f"""
<h2>3. Historique de la valeur du portefeuille</h2>

<h3>Cabinet – Valeur du portefeuille par date</h3>
{html_B_val}
"""
        else:
            synth_html = f"""
<div class="block">
  <h3>Situation actuelle — Client</h3>
  <ul>
    <li>Valeur actuelle : <b>{_fmt_eur(synthA.get("val", 0))}</b></li>
    <li>Montants réellement investis (net) : {_fmt_eur(synthA.get("net", 0))}</li>
    <li>Montants versés (brut) : {_fmt_eur(synthA.get("brut", 0))}</li>
    <li>Rendement total depuis le début : <b>{synthA.get("perf_tot_pct", 0):.2f} %</b></li>
    <li>Rendement annualisé (XIRR) : <b>{synthA.get("irr_pct", 0):.2f} %</b></li>
  </ul>
</div>
"""
            detail_html = f"""
<h2>2. Détail des lignes</h2>

<h3>Portefeuille Client</h3>
{html_client_lines}
"""
            hist_html = f"""
<h2>3. Historique de la valeur du portefeuille</h2>

<h3>Client – Valeur du portefeuille par date</h3>
{html_A_val}
"""

        # ── Section analyse fondamentale ────────────────────────────────────
        fa_html = ""
        _fa = report.get("fa_result")
        if _fa and not _fa.get("error"):
            _FA_ALLOC_LBL = {
                "AssetAllocEquity": "Actions",
                "AssetAllocBond": "Obligations",
                "AssetAllocCash": "Cash",
                "AssetAllocOther": "Autres",
                "AssetAllocNotClassified": "Non classifié",
            }
            _FA_GEO_LBL = {
                "northAmerica": "Amérique du Nord",
                "europeDeveloped": "Europe développée",
                "asiaDeveloped": "Asie développée",
                "asiaEmerging": "Asie émergente",
                "japan": "Japon",
                "latinAmerica": "Amérique latine",
                "unitedKingdom": "Royaume-Uni",
                "europeEmerging": "Europe émergente",
                "africaMiddleEast": "Afrique / Moyen-Orient",
                "australasia": "Australasie",
            }

            # Couverture
            _cov = _fa.get("covered_pct", 0.0)
            _nf = _fa.get("not_found", [])
            _nf_str = (
                " — Fonds non couverts : "
                + ", ".join(
                    f"{d.get('name', '—')} ({d.get('isin', '—')})" for d in _nf
                )
            ) if _nf else ""

            # Allocation
            _alloc = _fa.get("allocation") or {}
            _alloc_rows = "".join(
                f"<tr><td>{_FA_ALLOC_LBL.get(k, k)}</td><td>{v:.1f}%</td></tr>"
                for k, v in _alloc.items()
                if v > 0.5
            )
            _alloc_table = (
                f"<table><tr><th>Classe d'actif</th><th>Poids</th></tr>"
                f"{_alloc_rows}</table>"
            ) if _alloc_rows else "<p>Données non disponibles.</p>"

            # Géographie top 8
            _geo = _fa.get("geography") or {}
            _geo_sorted = sorted(
                [(k, v) for k, v in _geo.items() if v > 0.5],
                key=lambda x: x[1],
                reverse=True,
            )[:8]
            _geo_rows = "".join(
                f"<tr><td>{_FA_GEO_LBL.get(k, k)}</td><td>{v:.1f}%</td></tr>"
                for k, v in _geo_sorted
            )
            _geo_table = (
                f"<table><tr><th>Zone</th><th>Poids</th></tr>"
                f"{_geo_rows}</table>"
            ) if _geo_rows else ""

            # Secteurs actions top 8
            _sec_eq = _fa.get("sectors_equity") or {}
            _sec_eq_sorted = sorted(_sec_eq.items(), key=lambda x: x[1], reverse=True)[:8]
            _sec_eq_rows = "".join(
                f"<tr><td>{k}</td><td>{v:.1f}%</td></tr>"
                for k, v in _sec_eq_sorted
            )
            _sec_eq_table = (
                f"<table><tr><th>Secteur</th><th>Poids</th></tr>"
                f"{_sec_eq_rows}</table>"
            ) if _sec_eq_rows else ""

            # Top 10 holdings
            _holdings = (_fa.get("top_holdings") or [])[:10]
            _hold_rows = "".join(
                f"<tr><td>{h.get('name', '—')[:40]}</td>"
                f"<td>{h.get('isin', '—')}</td>"
                f"<td>{h['weight_portfolio'] * 100:.2f}%</td>"
                f"<td>{h.get('country', '—')}</td>"
                f"<td>{h.get('sector', '—')[:25] if h.get('sector') else '—'}</td></tr>"
                for h in _holdings
            )
            _hold_table = (
                f"<table><tr><th>Titre</th><th>ISIN</th>"
                f"<th>Poids %</th><th>Pays</th><th>Secteur</th></tr>"
                f"{_hold_rows}</table>"
            ) if _hold_rows else ""

            # ESG
            _esg = _fa.get("esg_score")
            _esg_html = ""
            if _esg is not None:
                _esg_cat = (
                    "Négligeable" if _esg <= 10
                    else "Faible" if _esg <= 20
                    else "Moyen" if _esg <= 30
                    else "Élevé" if _esg <= 40
                    else "Sévère"
                )
                _esg_html = (
                    f"<h3>Score ESG agrégé (Morningstar Sustainalytics)</h3>"
                    f"<p>Score moyen pondéré : <b>{_esg:.1f}</b> — "
                    f"Catégorie : <b>{_esg_cat}</b></p>"
                )

            fa_html = f"""
<h2>4. Analyse fondamentale des sous-jacents</h2>
<p class="small">
  Couverture Morningstar : <b>{_cov:.0f}%</b> du portefeuille. Données pondérées par la valeur
  actuelle de chaque fonds. Fonds euros et produits structurés exclus.{_nf_str}
</p>
<h3>Allocation d'actifs consolidée</h3>
{_alloc_table}
<h3>Répartition géographique (top 8)</h3>
{_geo_table}
<h3>Secteurs actions (top 8)</h3>
{_sec_eq_table}
<h3>Top 10 positions consolidées</h3>
{_hold_table}
{_esg_html}
"""

        html = f"""
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8" />
<title>Rapport de portefeuille</title>
<style>
body {{
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  margin: 24px;
  color: #222;
}}
h1, h2, h3 {{
  margin-top: 24px;
}}
table {{
  border-collapse: collapse;
  width: 100%;
  margin: 8px 0 16px 0;
  font-size: 14px;
}}
th, td {{
  border: 1px solid #ddd;
  padding: 6px 8px;
}}
th {{
  background-color: #f5f5f5;
  text-align: left;
}}
.small {{
  font-size: 12px;
  color: #666;
}}
.block {{
  border: 1px solid #eee;
  border-radius: 8px;
  padding: 12px 16px;
  margin-bottom: 16px;
  background-color: #fafafa;
}}
</style>
</head>
<body>

<h1>Rapport de portefeuille</h1>
<p class="small">Date de génération : {as_of}</p>

<h2>1. Synthèse chiffrée</h2>

{synth_html}
{detail_html}
{hist_html}
{fa_html}

<p class="small">
Ce document est fourni à titre informatif uniquement et ne constitue pas un conseil en investissement
personnalisé.
</p>

</body>
    </html>
"""
        return html


    def _add_table_to_story(
        story: List[Any],
        df: pd.DataFrame,
        col_widths: Optional[List[float]] = None,
        font_size: int = 9,
    ):
        if df.empty:
            story.append(Paragraph("Données indisponibles.", getSampleStyleSheet()["Normal"]))
            return
        headers = list(df.columns)
        fmt_rows = []
        for _, row in df.iterrows():
            formatted = []
            for col, val in row.items():
                if "€" in col:
                    formatted.append(val if isinstance(val, str) else fmt_eur_fr(val))
                elif "%" in col:
                    formatted.append(val if isinstance(val, str) else fmt_pct_fr(val))
                else:
                    formatted.append(str(val))
            fmt_rows.append(formatted)
        data = [headers] + fmt_rows
        table = Table(data, repeatRows=1, colWidths=col_widths)
        style = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), font_size),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]
        for i in range(1, len(data)):
            if i % 2 == 0:
                style.append(("BACKGROUND", (0, i), (-1, i), colors.whitesmoke))
        table.setStyle(TableStyle(style))
        story.append(table)


    def _fig_to_rl_image(fig: plt.Figure, width: float = 480, height: float = 270) -> Image:
        if not MATPLOTLIB_AVAILABLE:
            return None
        buf = BytesIO()
        fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
        buf.seek(0)
        plt.close(fig)
        return Image(buf, width=width, height=height)


    def _build_value_chart(df_map: Dict[str, pd.DataFrame]) -> Optional[Image]:
        if not MATPLOTLIB_AVAILABLE:
            return None
        if not df_map:
            return None
        fig, ax = plt.subplots(figsize=(6, 3))
        has_data = False
        for label, df in df_map.items():
            if df is None or df.empty or "Valeur" not in df.columns:
                continue
            ax.plot(df.index, df["Valeur"], label=label)
            has_data = True
        if not has_data:
            plt.close(fig)
            return None
        ax.set_title("Évolution de la valeur du portefeuille")
        ax.set_xlabel("Date")
        ax.set_ylabel("Valeur (€)")
        ax.legend(loc="best")
        fig.autofmt_xdate()
        return _fig_to_rl_image(fig)


    def _wrap_label(label: str, width: int = 28) -> str:
        if not label:
            return "—"
        return "\n".join(textwrap.wrap(str(label), width=width)) or str(label)


    def _allocation_from_positions(df_positions: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
        df = df_positions.copy()
        df = df[df["Valeur actuelle €"] >= 0]
        df["Nom"] = df["Nom"].fillna("—")
        df["ISIN / Code"] = df["ISIN / Code"].fillna("—")

        total_value = df["Valeur actuelle €"].sum()
        if total_value > 0:
            df["Poids"] = df["Valeur actuelle €"] / total_value
            basis_label = "Valeur actuelle"
        else:
            total_net = df["Net investi €"].sum()
            if total_net > 0:
                df["Poids"] = df["Net investi €"] / total_net
            else:
                df["Poids"] = 0.0
                if len(df) > 0:
                    df.loc[df.index[0], "Poids"] = 1.0
            basis_label = "Net investi"

        df = df.sort_values("Poids", ascending=False)
        if len(df) > 8:
            df_main = df.iloc[:8].copy()
            df_other = df.iloc[8:]
            other_row = pd.DataFrame(
                {
                    "Nom": ["Autres"],
                    "ISIN / Code": ["—"],
                    "Net investi €": [df_other["Net investi €"].sum()],
                    "Valeur actuelle €": [df_other["Valeur actuelle €"].sum()],
                    "Poids": [df_other["Poids"].sum()],
                }
            )
            df = pd.concat([df_main, other_row], ignore_index=True)

        df["Part %"] = df["Poids"] * 100.0
        return df, basis_label


    def _build_allocation_donut(
        df_alloc: pd.DataFrame,
        title: str,
        figsize: Tuple[float, float] = (6.0, 3.4),
    ) -> Optional[Image]:
        if not MATPLOTLIB_AVAILABLE or df_alloc.empty:
            return None
        fig, ax = plt.subplots(figsize=figsize)
        wedges, _ = ax.pie(
            df_alloc["Poids"],
            startangle=90,
            labels=None,
            wedgeprops=dict(width=0.35, edgecolor="white"),
        )
        labels = [
            f"{_wrap_label(nm)} ({pct:.1f}%)"
            for nm, pct in zip(df_alloc["Nom"], df_alloc["Part %"])
        ]
        ax.legend(
            wedges,
            labels,
            loc="center left",
            bbox_to_anchor=(1.02, 0.5),
            frameon=False,
            fontsize=8,
        )
        ax.set_title(title)
        ax.set_aspect("equal")
        fig.tight_layout(rect=[0, 0, 0.78, 1])
        return _fig_to_rl_image(fig, width=460, height=280)


    def _build_envelope_breakdown(
        lines: List[Dict[str, Any]],
        title: str,
    ) -> Tuple[Optional[Image], Optional[str]]:
        if not lines:
            return None, "Répartition par enveloppe : —"
        categories = {"Fonds euros": 0.0, "UC": 0.0, "Structurés": 0.0}
        for ln in lines:
            isin = str(ln.get("isin", "")).upper()
            val = float(ln.get("value", 0.0))
            if isin == "EUROFUND":
                categories["Fonds euros"] += val
            elif isin == "STRUCTURED":
                categories["Structurés"] += val
            else:
                categories["UC"] += val
        total = sum(categories.values())
        if total <= 0:
            return None, "Répartition par enveloppe : —"

        shares = {k: v / total for k, v in categories.items() if v > 0}
        major = max(shares.items(), key=lambda x: x[1])
        if sum(1 for v in shares.values() if v >= 0.01) < 2:
            return None, f"Répartition par enveloppe : {major[1] * 100:.1f}% {major[0]}"

        if not MATPLOTLIB_AVAILABLE:
            return None, None
        labels = list(shares.keys())
        values = [shares[k] * 100 for k in labels]
        fig, ax = plt.subplots(figsize=(6.0, 1.8))
        ax.barh(labels, values, color="#4C78A8")
        ax.set_xlim(0, 100)
        ax.set_xlabel("%")
        ax.set_title(title)
        for i, v in enumerate(values):
            ax.text(min(v + 1, 98), i, f"{v:.1f}%", va="center", fontsize=8)
        fig.tight_layout()
        return _fig_to_rl_image(fig, width=460, height=140), None


    def _build_contribution_bar(df_positions: pd.DataFrame) -> Optional[Image]:
        if not MATPLOTLIB_AVAILABLE or df_positions.empty:
            return None
        df = df_positions.copy()
        if not {"Nom", "Valeur actuelle €", "Net investi €"}.issubset(df.columns):
            return None
        df["Contribution €"] = df["Valeur actuelle €"] - df["Net investi €"]
        df = df.sort_values("Contribution €", ascending=False)
        fig_height = max(2.0, min(4.2, 0.35 * len(df) + 1.2))
        fig, ax = plt.subplots(figsize=(6.2, fig_height))
        ax.barh(df["Nom"], df["Contribution €"], color="#2F6F9F")
        ax.invert_yaxis()
        ax.set_title("Contribution à la performance (€)")
        ax.axvline(0, color="black", linewidth=0.5)
        ax.tick_params(axis="y", labelsize=8)
        for i, v in enumerate(df["Contribution €"]):
            offset = 0.01 * abs(v) if v != 0 else 0.5
            x_pos = v + offset if v >= 0 else v - offset
            ax.text(x_pos, i, fmt_eur_fr(v), va="center", fontsize=8)
        fig.tight_layout()
        return _fig_to_rl_image(fig, width=460, height=200)


    def _build_single_line_bar(label: str, value: float, title: str) -> Optional[Image]:
        if not MATPLOTLIB_AVAILABLE:
            return None
        fig, ax = plt.subplots(figsize=(6.0, 1.3))
        ax.barh([label], [100], color="#4C78A8")
        ax.set_xlim(0, 100)
        ax.set_title(title)
        ax.set_xticks([])
        ax.tick_params(axis="y", labelsize=8)
        fig.tight_layout()
        return _fig_to_rl_image(fig, width=460, height=90)


    def fmt_eur_pdf(x: Any) -> str:
        return fmt_eur_fr(x)


    def generate_pdf_report(report: Dict[str, Any]) -> bytes:
        if not REPORTLAB_AVAILABLE:
            raise RuntimeError(f"PDF indisponible: {REPORTLAB_ERROR}")
        if not MATPLOTLIB_AVAILABLE:
            raise RuntimeError(f"PDF indisponible: {MATPLOTLIB_ERROR}")

        class NumberedCanvas(canvas.Canvas):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self._saved_page_states = []

            def showPage(self):
                self._saved_page_states.append(dict(self.__dict__))
                self._startPage()

            def save(self):
                page_count = len(self._saved_page_states)
                for state in self._saved_page_states:
                    self.__dict__.update(state)
                    self._draw_header_footer(page_count)
                    super().showPage()
                super().save()

            def _draw_header_footer(self, page_count: int):
                width, height = A4
                self.setFillColor(colors.HexColor("#1F3B6D"))
                self.setFont("Helvetica-Bold", 10)
                self.drawString(36, height - 30, "Rapport de portefeuille – Cabinet")
                self.setFillColor(colors.grey)
                self.setFont("Helvetica", 8)
                self.drawRightString(width - 36, height - 30, report.get("as_of", ""))
                self.setFillColor(colors.grey)
                self.setFont("Helvetica", 7)
                self.drawString(
                    36,
                    24,
                    "Document informatif, ne constitue pas un conseil en investissement.",
                )
                self.drawRightString(width - 36, 24, f"Page {self.getPageNumber()} / {page_count}")

        buffer = BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            leftMargin=36,
            rightMargin=36,
            topMargin=54,
            bottomMargin=48,
        )
        base_styles = getSampleStyleSheet()
        styles = {
            "title": ParagraphStyle(
                "Title",
                parent=base_styles["Title"],
                textColor=colors.HexColor("#1F3B6D"),
                fontSize=20,
                spaceAfter=12,
            ),
            "h1": ParagraphStyle(
                "H1",
                parent=base_styles["Heading1"],
                textColor=colors.HexColor("#1F3B6D"),
                fontSize=14,
                spaceAfter=8,
            ),
            "h2": ParagraphStyle(
                "H2",
                parent=base_styles["Heading2"],
                textColor=colors.HexColor("#4B5563"),
                fontSize=12,
                spaceAfter=6,
            ),
            "small": ParagraphStyle(
                "Small",
                parent=base_styles["Normal"],
                fontSize=8,
                textColor=colors.grey,
            ),
            "kpi": ParagraphStyle(
                "KPI",
                parent=base_styles["Normal"],
                fontSize=10,
                textColor=colors.HexColor("#111827"),
            ),
        }
        story: List[Any] = []

        def _kpi_table(title: str, rows: List[Tuple[str, str]]) -> Table:
            data = [[Paragraph(f"<b>{title}</b>", styles["h2"]), ""]]
            for label, value in rows:
                data.append([Paragraph(label, styles["small"]), Paragraph(value, styles["kpi"])])
            table = Table(data, colWidths=[160, 120])
            table.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EEF2F7")),
                        ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
                        ("BOX", (0, 0), (-1, -1), 0.5, colors.lightgrey),
                        ("INNERGRID", (0, 1), (-1, -1), 0.25, colors.lightgrey),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 6),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ]
                )
            )
            return table

        story.append(Paragraph("Rapport client", styles["title"]))
        story.append(Paragraph(f"Date de génération : {report.get('as_of', '')}", styles["small"]))
        story.append(Spacer(1, 12))

        mode = report.get("mode", "compare")
        synthA = report.get("client_summary", {})
        synthB = report.get("valority_summary", {})
        comp = report.get("comparison", {})

        story.append(Paragraph("Synthèse", styles["h1"]))
        if mode == "compare":
            client_rows = [
                ("Valeur actuelle", fmt_eur_fr(synthA.get("val", 0))),
                ("Net investi", fmt_eur_fr(synthA.get("net", 0))),
                ("Brut versé", fmt_eur_fr(synthA.get("brut", 0))),
                ("Perf totale", fmt_pct_fr(synthA.get("perf_tot_pct", 0))),
                ("XIRR", fmt_pct_fr(synthA.get("irr_pct", 0))),
            ]
            valority_rows = [
                ("Valeur actuelle", fmt_eur_fr(synthB.get("val", 0))),
                ("Net investi", fmt_eur_fr(synthB.get("net", 0))),
                ("Brut versé", fmt_eur_fr(synthB.get("brut", 0))),
                ("Perf totale", fmt_pct_fr(synthB.get("perf_tot_pct", 0))),
                ("XIRR", fmt_pct_fr(synthB.get("irr_pct", 0))),
            ]
            table = Table(
                [[_kpi_table("Client", client_rows), _kpi_table("Cabinet", valority_rows)]],
                colWidths=[240, 240],
            )
            story.append(table)
            story.append(Spacer(1, 4))
            story.append(Paragraph(
                "Performance nette des frais de gestion du contrat. "
                "Les frais internes des fonds (TER) sont intégrés dans les valeurs liquidatives.",
                styles["small"],
            ))
            story.append(Spacer(1, 8))
            comp_rows = [
                ("Différence de valeur", fmt_eur_fr(comp.get("delta_val", 0))),
                ("Écart de performance", fmt_pct_fr(comp.get("delta_perf_pct", 0))),
            ]
            story.append(_kpi_table("Comparaison", comp_rows))
        else:
            title = "Cabinet" if mode == "valority" else "Client"
            synth = synthB if mode == "valority" else synthA
            rows = [
                ("Valeur actuelle", fmt_eur_fr(synth.get("val", 0))),
                ("Net investi", fmt_eur_fr(synth.get("net", 0))),
                ("Brut versé", fmt_eur_fr(synth.get("brut", 0))),
                ("Perf totale", fmt_pct_fr(synth.get("perf_tot_pct", 0))),
                ("XIRR", fmt_pct_fr(synth.get("irr_pct", 0))),
            ]
            story.append(_kpi_table(title, rows))
            story.append(Spacer(1, 4))
            story.append(Paragraph(
                "Performance nette des frais de gestion du contrat. "
                "Les frais internes des fonds (TER) sont intégrés dans les valeurs liquidatives.",
                styles["small"],
            ))
            fees = report.get("fees_analysis", {})
            if fees:
                story.append(Spacer(1, 8))
                fees_rows = [
                    ("Frais d’entrée payés", fmt_eur_fr(fees.get("fees_paid", 0))),
                    ("Valeur créée", fmt_eur_fr(fees.get("value_created", 0))),
                    ("Valeur/an", fmt_eur_fr(fees.get("value_per_year", 0))),
                ]
                story.append(_kpi_table("Frais & valeur créée", fees_rows))

        story.append(Spacer(1, 12))
        story.append(Paragraph("Graphiques", styles["h1"]))

        value_chart = _build_value_chart(report.get("df_map", {}))
        if value_chart is not None:
            story.append(value_chart)
            story.append(Spacer(1, 8))
        else:
            story.append(Paragraph("Données indisponibles pour la courbe de valeur.", styles["small"]))

        def add_portfolio_details_section(
            label: str,
            positions_df: pd.DataFrame,
            lines_values: List[Dict[str, Any]],
        ):
            # PAGE 2/4 — Détail du portefeuille
            story.append(PageBreak())
            story.append(Paragraph(f"Détail du portefeuille — {label}", styles["h1"]))

            if isinstance(positions_df, pd.DataFrame) and not positions_df.empty:
                df_alloc, basis_label = _allocation_from_positions(positions_df)

                if len(df_alloc) >= 2:
                    story.append(Paragraph(f"Allocation par ligne ({basis_label})", styles["h2"]))
                    donut = _build_allocation_donut(df_alloc, "Allocation par ligne")
                    if donut is not None:
                        story.append(donut)
                        story.append(Spacer(1, 6))
                else:
                    line = df_alloc.iloc[0] if not df_alloc.empty else None
                    name = line["Nom"] if line is not None else "—"
                    story.append(
                        Paragraph(
                            f"Portefeuille concentré : 100% sur <b>{name}</b>.",
                            styles["kpi"],
                        )
                    )
                    bar = _build_single_line_bar(_wrap_label(name), 100.0, "Répartition 100%")
                    if bar is not None:
                        story.append(bar)
                        story.append(Spacer(1, 6))

                alloc_table = df_alloc[
                    ["Nom", "ISIN / Code", "Part %", "Net investi €", "Valeur actuelle €"]
                ].copy()
                _add_table_to_story(
                    story,
                    alloc_table,
                    col_widths=[170, 80, 55, 95, 95],
                    font_size=9,
                )
                story.append(Spacer(1, 6))

                envelope_chart, envelope_text = _build_envelope_breakdown(
                    lines_values,
                    "Répartition par enveloppe",
                )
                if envelope_chart is not None:
                    story.append(envelope_chart)
                elif envelope_text:
                    story.append(Paragraph(envelope_text, styles["small"]))
            else:
                story.append(Paragraph("Données indisponibles pour la composition.", styles["small"]))

            # PAGE 3/5 — Contribution & Positions
            story.append(PageBreak())
            story.append(Paragraph(f"Contribution & positions — {label}", styles["h1"]))

            if isinstance(positions_df, pd.DataFrame) and not positions_df.empty:
                if len(positions_df) == 1:
                    ln = positions_df.iloc[0]
                    story.append(
                        Paragraph(
                            f"Contribution : <b>{ln['Nom']}</b> = {fmt_eur_pdf(ln['Valeur actuelle €'] - ln['Net investi €'])}",
                            styles["kpi"],
                        )
                    )
                    bar = _build_single_line_bar(_wrap_label(ln["Nom"]), 100.0, "Contribution (ligne unique)")
                    if bar is not None:
                        story.append(bar)
                        story.append(Spacer(1, 6))
                else:
                    contrib_chart = _build_contribution_bar(positions_df)
                    if contrib_chart is not None:
                        story.append(contrib_chart)
                        story.append(Spacer(1, 6))
            else:
                story.append(Paragraph("Données indisponibles pour la contribution.", styles["small"]))

            story.append(Paragraph("Positions", styles["h2"]))
            if isinstance(positions_df, pd.DataFrame) and not positions_df.empty:
                positions_table = positions_df[
                    [
                        "Nom",
                        "ISIN / Code",
                        "Date d'achat",
                        "Net investi €",
                        "Valeur actuelle €",
                        "Perf €",
                        "Perf %",
                    ]
                ].copy()
                _add_table_to_story(
                    story,
                    positions_table,
                    col_widths=[150, 70, 65, 80, 80, 50, 45],
                    font_size=9,
                )
            else:
                story.append(Paragraph("Données indisponibles.", styles["small"]))

        if mode == "compare":
            add_portfolio_details_section(
                "Client",
                report.get("positions_df_client", pd.DataFrame()),
                report.get("lines_client", []),
            )
            add_portfolio_details_section(
                "Cabinet",
                report.get("positions_df_valority", pd.DataFrame()),
                report.get("lines_valority", []),
            )
        else:
            add_portfolio_details_section(
                "Cabinet" if mode == "valority" else "Client",
                report.get("positions_df", pd.DataFrame()),
                report.get("lines", []),
            )

        # ── Page Analyse fondamentale (si disponible) ─────────────────
        # FIXED (P1): lire depuis report (passé explicitement) avec fallback session_state
        _fa_result = report.get("fa_result") or st.session_state.get("FUND_ANALYSIS_RESULT")
        if _fa_result and not _fa_result.get("error"):
            try:
                story.append(PageBreak())
                story.append(Paragraph("Analyse fondamentale des sous-jacents", styles["h1"]))

                # Bandeau couverture
                _fa_cov = _fa_result.get("covered_pct", 0.0)
                story.append(
                    Paragraph(
                        f"Couverture Morningstar : {_fa_cov:.0f}% du portefeuille analysé. "
                        "Données pondérées par la valeur actuelle de chaque fonds. "
                        "Fonds euros et produits structurés exclus.",
                        styles["small"],
                    )
                )
                _fa_not_found = _fa_result.get("not_found", [])
                if _fa_not_found:
                    _nf_str = ", ".join(
                        f"{d.get('name','—')} ({d.get('isin','—')})"
                        for d in _fa_not_found
                    )
                    story.append(
                        Paragraph(f"Fonds non couverts : {_nf_str}", styles["small"])
                    )
                story.append(Spacer(1, 8))

                # Tableau allocation d'actifs
                _fa_alloc = _fa_result.get("allocation") or {}
                if _fa_alloc:
                    story.append(Paragraph("Allocation d'actifs consolidée", styles["h2"]))
                    _FA_ALLOC_LBL = {
                        "AssetAllocEquity":        "Actions",
                        "AssetAllocBond":          "Obligations",
                        "AssetAllocCash":          "Cash",
                        "AssetAllocOther":         "Autres",
                        "AssetAllocNotClassified": "Non classifié",
                    }
                    _alloc_rows = [
                        [_FA_ALLOC_LBL.get(k, k), f"{v:.1f}%"]
                        for k, v in _fa_alloc.items() if v > 0.5
                    ]
                    if _alloc_rows:
                        _add_table_to_story(
                            story,
                            pd.DataFrame(_alloc_rows, columns=["Classe d'actif", "Poids (%)"]),
                            col_widths=[200, 80],
                            font_size=9,
                        )
                        # FIXED (P3): note si somme < 95%
                        try:
                            _alloc_total_pdf = sum(
                                float(v.rstrip("%")) for _, v in _alloc_rows
                            )
                        except Exception:
                            _alloc_total_pdf = 100.0
                        if _alloc_total_pdf < 95.0:
                            story.append(Paragraph(
                                f"Note : total affiché {_alloc_total_pdf:.1f}% (expositions "
                                "nettes Morningstar, dérivés exclus).",
                                styles["small"],
                            ))
                    story.append(Spacer(1, 8))

                # Tableau géographie top 8
                _fa_geo = _fa_result.get("geography") or {}
                if _fa_geo:
                    story.append(Paragraph("Répartition géographique (top 8 zones)", styles["h2"]))
                    _FA_GEO_LBL = {
                        "northAmerica":     "Amérique du Nord",
                        "europeDeveloped":  "Europe développée",
                        "asiaDeveloped":    "Asie développée",
                        "asiaEmerging":     "Asie émergente",
                        "japan":            "Japon",
                        "latinAmerica":     "Amérique latine",
                        "unitedKingdom":    "Royaume-Uni",
                        "europeEmerging":   "Europe émergente",
                        "africaMiddleEast": "Afrique / Moyen-Orient",
                        "australasia":      "Australasie",
                    }
                    _geo_sorted = sorted(
                        [(k, v) for k, v in _fa_geo.items() if v > 0.5],
                        key=lambda x: x[1], reverse=True,
                    )[:8]
                    if _geo_sorted:
                        _add_table_to_story(
                            story,
                            pd.DataFrame(
                                [[_FA_GEO_LBL.get(k, k), f"{v:.1f}%"] for k, v in _geo_sorted],
                                columns=["Zone géographique", "Poids (%)"],
                            ),
                            col_widths=[200, 80],
                            font_size=9,
                        )
                    story.append(Spacer(1, 8))

                # Tableau secteurs actions top 8
                _fa_sec_eq = _fa_result.get("sectors_equity") or {}
                if _fa_sec_eq:
                    story.append(Paragraph("Secteurs actions (top 8)", styles["h2"]))
                    _fa_sec_eq_sorted = sorted(
                        _fa_sec_eq.items(), key=lambda x: x[1], reverse=True
                    )[:8]
                    _add_table_to_story(
                        story,
                        pd.DataFrame(
                            [[_FS_SECTOR_LABELS.get(k, k), f"{v:.1f}%"] for k, v in _fa_sec_eq_sorted],
                            columns=["Secteur", "Poids (%)"],
                        ),
                        col_widths=[200, 80],
                        font_size=9,
                    )
                    story.append(Spacer(1, 8))

                # Tableau top 10 holdings consolidés
                _fa_top_h = _fa_result.get("top_holdings") or []
                if _fa_top_h:
                    story.append(Paragraph("Top 10 positions consolidées", styles["h2"]))
                    _h_rows = [
                        [
                            h.get("name", "—")[:40],
                            h.get("isin", "—"),
                            f"{h['weight_portfolio'] * 100:.2f}%",
                            h.get("country", "—"),
                            h.get("sector", "—")[:20] if h.get("sector") else "—",
                        ]
                        for h in _fa_top_h[:10]
                    ]
                    _add_table_to_story(
                        story,
                        pd.DataFrame(
                            _h_rows,
                            columns=["Titre", "ISIN", "Poids %", "Pays", "Secteur"],
                        ),
                        col_widths=[160, 85, 50, 50, 80],
                        font_size=8,
                    )
                    story.append(Spacer(1, 8))

                # Score ESG
                _fa_esg = _fa_result.get("esg_score")
                if _fa_esg is not None:
                    if _fa_esg <= 10:
                        _fa_esg_cat = "Négligeable"
                    elif _fa_esg <= 20:
                        _fa_esg_cat = "Faible"
                    elif _fa_esg <= 30:
                        _fa_esg_cat = "Moyen"
                    elif _fa_esg <= 40:
                        _fa_esg_cat = "Élevé"
                    else:
                        _fa_esg_cat = "Sévère"
                    story.append(Paragraph("Score ESG agrégé (Morningstar Sustainalytics)", styles["h2"]))
                    story.append(
                        Paragraph(
                            f"Score moyen pondéré : {_fa_esg:.1f} — Catégorie : {_fa_esg_cat}",
                            styles["kpi"],
                        )
                    )
            except Exception:
                pass

        doc.build(story, canvasmaker=NumberedCanvas)
        buffer.seek(0)
        return buffer.read()


    def _years_between(d0: pd.Timestamp, d1: pd.Timestamp) -> float:
        return max(0.0, (d1 - d0).days / 365.25)


    report_data = {
        "as_of": fmt_date(TODAY),
        "mode": st.session_state.get("MODE_ANALYSE", "compare"),
        # FIXED (P1): transmettre le résultat analyse fondamentale dans report_data
        # pour éviter qu'il ne soit absent si le session_state est lu trop tôt
        "fa_result": st.session_state.get("FUND_ANALYSIS_RESULT"),
    }

    mode_report = report_data["mode"]
    df_client_lines = build_positions_dataframe("A_lines") if show_client else pd.DataFrame()
    df_valority_lines = build_positions_dataframe("B_lines") if show_valority else pd.DataFrame()

    report_data["df_client_lines"] = df_client_lines
    report_data["df_valority_lines"] = df_valority_lines
    report_data["positions_df_client"] = df_client_lines
    report_data["positions_df_valority"] = df_valority_lines
    report_data["dfA_val"] = (
        dfA.reset_index().rename(columns={"index": "Date"}) if (show_client and not dfA.empty) else pd.DataFrame()
    )
    report_data["dfB_val"] = (
        dfB.reset_index().rename(columns={"index": "Date"}) if (show_valority and not dfB.empty) else pd.DataFrame()
    )
    report_data["client_summary"] = {
        "val": valA,
        "net": netA,
        "brut": brutA,
        "perf_tot_pct": perf_tot_client or 0.0,
        "irr_pct": xirrA or 0.0,
    }
    report_data["valority_summary"] = {
        "val": valB,
        "net": netB,
        "brut": brutB,
        "perf_tot_pct": perf_tot_valority or 0.0,
        "irr_pct": xirrB or 0.0,
    }
    report_data["comparison"] = {
        "delta_val": (valB - valA) if (valA is not None and valB is not None) else 0.0,
        "delta_perf_pct": (perf_tot_valority - perf_tot_client)
        if (perf_tot_client is not None and perf_tot_valority is not None)
        else 0.0,
    }

    df_map: Dict[str, pd.DataFrame] = {}
    if mode_report in ("compare", "client") and not dfA.empty:
        df_map["Client"] = dfA
    if mode_report in ("compare", "valority") and not dfB.empty:
        df_map["Cabinet"] = dfB
    report_data["df_map"] = df_map

    if mode_report == "compare":
        positions_df = df_valority_lines if not df_valority_lines.empty else df_client_lines
        lines = st.session_state.get("B_lines", []) or st.session_state.get("A_lines", [])
    else:
        if mode_report == "valority":
            positions_df = df_valority_lines
            lines = st.session_state.get("B_lines", [])
            start_min = startB_min
            brut = brutB
            net = netB
            val = valB
        else:
            positions_df = df_client_lines
            lines = st.session_state.get("A_lines", [])
            start_min = startA_min
            brut = brutA
            net = netA
            val = valA

        years = _years_between(start_min, TODAY) if isinstance(start_min, pd.Timestamp) else 0.0
        fees_paid = max(0.0, brut - net) if brut is not None and net is not None else 0.0
        value_created = (val - net) if val is not None and net is not None else 0.0
        value_per_year = (value_created / years) if years > 0 else 0.0
        report_data["fees_analysis"] = {
            "fees_paid": fees_paid,
            "value_created": value_created,
            "value_per_year": value_per_year,
        }

    def _build_lines_with_values(
        lines_src: List[Dict[str, Any]],
        positions_src: pd.DataFrame,
    ) -> List[Dict[str, Any]]:
        items: List[Dict[str, Any]] = []
        if isinstance(positions_src, pd.DataFrame) and not positions_src.empty:
            for ln in lines_src:
                isin = ln.get("isin", "")
                name = ln.get("name", "")
                match = positions_src[
                    (positions_src["Nom"] == name) & (positions_src["ISIN / Code"] == isin)
                ]
                val = float(match["Valeur actuelle €"].iloc[0]) if not match.empty else 0.0
                items.append({"isin": isin, "value": val})
        return items

    report_data["positions_df"] = positions_df
    report_data["lines"] = _build_lines_with_values(lines, positions_df)
    report_data["lines_client"] = _build_lines_with_values(
        st.session_state.get("A_lines", []),
        df_client_lines,
    )
    report_data["lines_valority"] = _build_lines_with_values(
        st.session_state.get("B_lines", []),
        df_valority_lines,
    )

    st.session_state["REPORT_DATA"] = report_data

    # ------------------------------------------------------------
    # Bloc final : Comparaison OU "Frais & valeur créée"
    # ------------------------------------------------------------
    mode = st.session_state.get("MODE_ANALYSE", "compare")

    # ============================
    # CAS 1 — MODE COMPARAISON
    # ============================
    if mode == "compare":
        st.subheader("📌 Comparaison : Client vs Cabinet")

        gain_vs_client = (valB - valA) if (valA is not None and valB is not None) else 0.0
        delta_xirr = (xirrB - xirrA) if (xirrA is not None and xirrB is not None) else None
        perf_diff_tot = (
            (perf_tot_valority - perf_tot_client)
            if (perf_tot_client is not None and perf_tot_valority is not None)
            else None
        )

        with st.container(border=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("Gain en valeur", to_eur(gain_vs_client))
            with c2:
                st.metric(
                    "Surperformance totale",
                    f"{perf_diff_tot:+.2f}%" if perf_diff_tot is not None else "—",
                )
            with c3:
                st.metric(
                    "Surperformance annualisée (Î” XIRR)",
                    f"{delta_xirr:+.2f}%" if delta_xirr is not None else "—",
                )

            st.markdown(
                f"""
Aujourd’hui, avec votre allocation actuelle, votre portefeuille vaut **{to_eur(valA)}**.  
Avec l’allocation Cabinet, il serait autour de **{to_eur(valB)}**, soit environ **{to_eur(gain_vs_client)}** de plus.
"""
            )

    # ============================
    # CAS 2 — MODE ANALYSE SIMPLE
    # ============================
    else:
        # Sélection des variables selon le mode
        if mode == "valority":
            brut = brutB
            net = netB
            val = valB
            start_min = startB_min
            irr = xirrB
            fee_pct = st.session_state.get("FEE_B", 0.0)
            title = "🏢 Allocation Cabinet — Frais & valeur créée"
        else:  # mode == "client"
            brut = brutA
            net = netA
            val = valA
            start_min = startA_min
            irr = xirrA
            fee_pct = st.session_state.get("FEE_A", 0.0)
            title = "🧍 Portefeuille — Frais & valeur créée"

        st.subheader("📌 Analyse : frais & valeur créée")

        if brut > 0 and net >= 0 and val >= 0 and isinstance(start_min, pd.Timestamp):
            fees_paid = max(0.0, brut - net)     # frais d'entrée réellement payés
            value_created = val - net            # valeur créée vs capital réellement investi
            years = _years_between(start_min, TODAY)
            value_per_year = (value_created / years) if years > 0 else None

            with st.container(border=True):
                st.markdown(f"#### {title}")
                st.caption(
                    f"Période : **{fmt_date(start_min)} → {fmt_date(TODAY)}** "
                    f"• Frais d’entrée : **{fee_pct:.2f}%**"
                )

                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("Frais d’entrée payés", to_eur(fees_paid))
                with c2:
                    st.metric("Valeur créée (net)", to_eur(value_created))
                with c3:
                    st.metric(
                        "Valeur créée / an (moyenne)",
                        to_eur(value_per_year) if value_per_year is not None else "—",
                    )

                st.markdown(
                    f"""
- Montants versés (brut) : **{to_eur(brut)}**
- Montants réellement investis (après frais) : **{to_eur(net)}**
- Valeur actuelle : **{to_eur(val)}**
"""
                )

                if irr is not None:
                    st.markdown(f"- Rendement annualisé (XIRR) : **{irr:.2f}%**")
                else:
                    st.markdown("- Rendement annualisé (XIRR) : **—**")

                # Message factuel — différencié selon que la valeur créée est positive ou négative
                if fees_paid > 0 and value_created > 0:
                    ratio = value_created / fees_paid
                    st.markdown(
                        f"**Lecture :** {to_eur(fees_paid)} de frais d’entrée ont généré "
                        f"**{to_eur(value_created)}** de valeur nette créée à date "
                        f"(**×{ratio:.1f}**)."
                    )
                elif fees_paid > 0 and value_created <= 0:
                    st.markdown(
                        f"**Lecture :** {to_eur(fees_paid)} de frais d’entrée payés. "
                        f"Le portefeuille affiche une moins-value nette de **{to_eur(abs(value_created))}** à date."
                    )
        else:
            st.info("Ajoutez des lignes (et/ou des versements) pour afficher l’analyse frais & valeur créée.")


    # ------------------------------------------------------------
    # Tables positions
    # ------------------------------------------------------------
    if show_client:
        positions_table("Portefeuille 1 — Client", "A_lines")
    if show_valority:
        positions_table("Portefeuille 2 — Cabinet", "B_lines")

    st.subheader("Composition du portefeuille")

    def _render_portfolio_pie(port_key: str, title: str):
        if not MATPLOTLIB_AVAILABLE:
            st.warning(f"{title} : Camembert indisponible ({MATPLOTLIB_ERROR}).")
            return
        df_positions = build_positions_dataframe(port_key)
        if df_positions.empty:
            st.info(f"{title} : Données indisponibles.")
            return
        df_pie = _prepare_pie_df(df_positions)
        if df_pie.empty:
            st.info(f"{title} : Données indisponibles.")
            return
        fig, ax = plt.subplots(figsize=(5, 3))
        ax.pie(
            df_pie["Valeur actuelle €"],
            labels=df_pie["Nom"],
            autopct="%1.1f%%",
        )
        ax.set_title(title)
        st.pyplot(fig)
        plt.close(fig)
        st.dataframe(
            df_pie[["Nom", "Valeur actuelle €", "Part %"]].style.format(
                {
                    "Valeur actuelle €": to_eur,
                    "Part %": "{:,.2f}%".format,
                }
            ),
            hide_index=True,
            use_container_width=True,
        )

    if show_client and show_valority:
        col_a, col_b = st.columns(2)
        with col_a:
            _render_portfolio_pie("A_lines", "Portefeuille Client")
        with col_b:
            _render_portfolio_pie("B_lines", "Portefeuille Cabinet")
    elif show_client:
        _render_portfolio_pie("A_lines", "Portefeuille Client")
    elif show_valority:
        _render_portfolio_pie("B_lines", "Portefeuille Cabinet")

    # APP – Composition
    def _wrap_label_app(label: str, width: int = 28) -> str:
        if not label:
            return "—"
        return "\n".join(textwrap.wrap(str(label), width=width)) or str(label)

    def _render_valority_composition_section():
        if not MATPLOTLIB_AVAILABLE:
            st.warning(f"Cabinet : graphique indisponible ({MATPLOTLIB_ERROR}).")
            return
        df_positions = build_positions_dataframe("B_lines")
        if df_positions.empty:
            st.info("Aucune donnée pour le portefeuille Cabinet.")
            return

        df = df_positions.copy()
        total_val = df["Valeur actuelle €"].sum()
        if total_val > 0:
            df["Poids %"] = df["Valeur actuelle €"] / total_val * 100.0
        else:
            total_net = df["Net investi €"].sum()
            if total_net > 0:
                df["Poids %"] = df["Net investi €"] / total_net * 100.0
            else:
                df["Poids %"] = 0.0
                if len(df) > 0:
                    df.loc[df.index[0], "Poids %"] = 100.0
        df = df.sort_values("Poids %", ascending=False)

        if len(df) > 8:
            df_main = df.iloc[:8].copy()
            df_other = df.iloc[8:]
            other_row = pd.DataFrame(
                {
                    "Nom": ["Autres"],
                    "ISIN / Code": ["—"],
                    "Date d'achat": ["—"],
                    "Net investi €": [df_other["Net investi €"].sum()],
                    "Valeur actuelle €": [df_other["Valeur actuelle €"].sum()],
                    "Perf €": [df_other["Perf €"].sum()],
                    "Perf %": [np.nan],
                    "Poids %": [df_other["Poids %"].sum()],
                }
            )
            df = pd.concat([df_main, other_row], ignore_index=True)

        if len(df) >= 2:
            fig, ax = plt.subplots(figsize=(5.2, 3.2))
            wedges, _ = ax.pie(
                df["Poids %"],
                startangle=90,
                labels=None,
                wedgeprops=dict(width=0.35, edgecolor="white"),
            )
            labels = [
                f"{_wrap_label_app(nm)} ({pct:.1f}%)"
                for nm, pct in zip(df["Nom"], df["Poids %"])
            ]
            ax.legend(
                wedges,
                labels,
                loc="center left",
                bbox_to_anchor=(1.02, 0.5),
                frameon=False,
                fontsize=8,
            )
            ax.set_aspect("equal")
            fig.tight_layout(rect=[0, 0, 0.78, 1])
            st.pyplot(fig)
            plt.close(fig)
        else:
            st.info("Portefeuille concentré : 100% sur une seule ligne.")

        df_table = df[["Nom", "ISIN / Code", "Poids %", "Net investi €", "Valeur actuelle €"]]
        st.dataframe(
            df_table.style.format(
                {
                    "Poids %": "{:,.2f}%".format,
                    "Net investi €": to_eur,
                    "Valeur actuelle €": to_eur,
                }
            ),
            hide_index=True,
            use_container_width=True,
        )

    if show_valority:
        st.subheader("Composition du portefeuille (Cabinet)")
        _render_valority_composition_section()

    # ----------------------------------------------------------------
    # Section téléchargement des rapports
    # ----------------------------------------------------------------
    def _generate_pdf_safe(rd: Dict[str, Any]) -> bytes:
        try:
            return generate_pdf_report(rd)
        except Exception as e:
            st.warning(f"PDF indisponible : {e}")
            return b""

    st.markdown("---")
    st.subheader("📥 Télécharger le rapport")
    _report_data = st.session_state.get("REPORT_DATA")
    if _report_data is not None:
        # ── Rapport standard (sans analyse fondamentale) ──────────────────
        _report_data_standard = {**_report_data, "fa_result": None}
        _html_standard = build_html_report(_report_data_standard)
        _col1, _col2 = st.columns(2)
        with _col1:
            st.download_button(
                "📄 Rapport standard (PDF)",
                data=_generate_pdf_safe(_report_data_standard),
                file_name="rapport_portefeuille.pdf",
                mime="application/pdf",
                help="Rapport de performance et composition, sans analyse Morningstar",
            )
        with _col2:
            st.download_button(
                "📄 Rapport standard (HTML)",
                data=_html_standard.encode("utf-8"),
                file_name="rapport_portefeuille.html",
                mime="text/html",
                help="Version HTML du rapport standard, consultable dans un navigateur",
            )

        # ── Rapport complet avec analyse fondamentale ─────────────────────
        st.markdown("---")
        _fa_result_live = st.session_state.get("FUND_ANALYSIS_RESULT")
        _fa_available = (
            _fa_result_live is not None
            and not (_fa_result_live or {}).get("error")
            and bool(
                (_fa_result_live or {}).get("allocation")
                or (_fa_result_live or {}).get("geography")
                or (_fa_result_live or {}).get("sectors_equity")
            )
        )
        if _fa_available:
            _report_data_full = {
                **_report_data,
                "fa_result": _fa_result_live,
            }
            _html_full = build_html_report(_report_data_full)
            _col3, _col4 = st.columns(2)
            with _col3:
                st.download_button(
                    "🔬 Rapport complet avec analyse fondamentale (PDF)",
                    data=_generate_pdf_safe(_report_data_full),
                    file_name="rapport_complet_fondamentaux.pdf",
                    mime="application/pdf",
                    help="Rapport de performance + radiographie Morningstar des sous-jacents",
                )
            with _col4:
                st.download_button(
                    "🔬 Rapport complet avec analyse fondamentale (HTML)",
                    data=_html_full.encode("utf-8"),
                    file_name="rapport_complet_fondamentaux.html",
                    mime="text/html",
                    help="Version HTML du rapport complet avec analyse Morningstar",
                )
        else:
            st.info(
                "💡 Pour télécharger le rapport complet incluant l’analyse "
                "fondamentale Morningstar, lancez d’abord l’analyse dans "
                "l’expander 🔬 ci-dessous."
            )
    else:
        st.info("Les rapports seront disponibles après le calcul du portefeuille.")

    with st.expander("Aide rapide"):
        st.markdown(
            """
- Dans chaque portefeuille, vous pouvez **soit** ajouter des *fonds recommandés* (onglet dédié),
  **soit** utiliser la *saisie libre* avec ISIN / code.
- Pour le **fonds en euros**, utilisez le symbole **EUROFUND** (taux paramétrable dans la barre de gauche).
- Les frais d’entrée s’appliquent à chaque investissement.
- Le **rendement total** est la performance globale depuis l’origine (valeur actuelle / net investi).
- Le **rendement annualisé** utilise le XIRR (prise en compte des dates et montants).
- En mode **Personnalisé**, vous pouvez affecter précisément les versements mensuels et ponctuels à chaque ligne,
  avec un contrôle automatique de cohérence par rapport aux montants bruts saisis.
            """
        )

    # ------------------------------------------------------------
    # Analyse interne — Corrélation & volatilité (réservé conseiller)
    # ------------------------------------------------------------
    st.markdown("---")
    with st.expander("🔒 Analyse interne — Corrélation, volatilité et profil de risque", expanded=False):
        st.caption(
            "Section réservée au conseiller : analyse technique basée sur les valeurs liquidatives "
            "(corrélations, volatilités, drawdown)."
        )

        # FIXED: use per-portfolio euro rates instead of shared EURO_RATE_PREVIEW (Résidu Bug 4)
        euro_rate_A = st.session_state.get("EURO_RATE_A", 2.0)
        euro_rate_B = st.session_state.get("EURO_RATE_B", 2.5)
        linesA = st.session_state.get("A_lines", [])
        linesB = st.session_state.get("B_lines", [])

        # Portefeuille Client
        if show_client:
            st.markdown("### Portefeuille 1 — Client")
            corrA = correlation_matrix_from_lines(linesA, euro_rate_A)
            volA = volatility_table_from_lines(linesA, euro_rate_A)
            riskA = portfolio_risk_stats(linesA, euro_rate_A, fee_pct=st.session_state.get("FEE_A", 0.0))

            if corrA.empty and volA.empty:
                st.info("Pas assez d'historique ou de lignes pour analyser ce portefeuille.")
            else:
                if riskA is not None:
                    c1, c2 = st.columns(2)
                    with c1:
                        st.metric(
                            "Volatilité annuelle estimée",
                            f"{riskA['vol_ann_pct']:.2f} %",
                        )
                    with c2:
                        st.metric(
                            "Max drawdown (historique sur la période)",
                            f"{riskA['max_dd_pct']:.2f} %",
                        )

                if not volA.empty:
                    st.markdown("**Volatilité par ligne**")
                    st.dataframe(
                        volA.style.format(
                            {
                                "Écart-type quotidien %": "{:,.2f}%".format,
                                "Volatilité annuelle %": "{:,.2f}%".format,
                            }
                        ),
                        use_container_width=True,
                    )

                if not corrA.empty:
                    chartA = _corr_heatmap_chart(corrA, "Corrélation des lignes — Portefeuille Client")
                    if chartA is not None:
                        st.altair_chart(chartA, use_container_width=True)

        if show_client and show_valority:
            st.markdown("---")

        if show_valority:
            st.markdown("### Portefeuille 2 — Cabinet")
            corrB = correlation_matrix_from_lines(linesB, euro_rate_B)
            volB = volatility_table_from_lines(linesB, euro_rate_B)
            riskB = portfolio_risk_stats(linesB, euro_rate_B, fee_pct=st.session_state.get("FEE_B", 0.0))

            if corrB.empty and volB.empty:
                st.info("Pas assez d'historique ou de lignes pour analyser ce portefeuille.")
            else:
                if riskB is not None:
                    c1, c2 = st.columns(2)
                    with c1:
                        st.metric(
                            "Volatilité annuelle estimée",
                            f"{riskB['vol_ann_pct']:.2f} %",
                        )
                    with c2:
                        st.metric(
                            "Max drawdown (historique sur la période)",
                            f"{riskB['max_dd_pct']:.2f} %",
                        )

                if not volB.empty:
                    st.markdown("**Volatilité par ligne**")
                    st.dataframe(
                        volB.style.format(
                            {
                                "Écart-type quotidien %": "{:,.2f}%".format,
                                "Volatilité annuelle %": "{:,.2f}%".format,
                            }
                        ),
                        use_container_width=True,
                    )

                if not corrB.empty:
                    chartB = _corr_heatmap_chart(corrB, "Corrélation des lignes — Portefeuille Cabinet")
                    if chartB is not None:
                        st.altair_chart(chartB, use_container_width=True)

    # ----------------------------------------------------------------
    # Analyse fondamentale des sous-jacents
    # ----------------------------------------------------------------
    st.markdown("---")
    with st.expander("🔬 Analyse fondamentale des sous-jacents", expanded=False):
        st.caption(
            "Agrégation des données Morningstar pondérées par la valeur "
            "actuelle de chaque fonds. Fonds euros et produits structurés "
            "exclus. Données mises à jour toutes les 7 jours."
        )
        if not MSTARPY_AVAILABLE:
            st.warning("Module mstarpy non installé — analyse fondamentale indisponible.")
        else:
            # Sélection du portefeuille à analyser
            if show_client and show_valority:
                port_choice = st.radio(
                    "Portefeuille à analyser",
                    ["Client", "Cabinet"],
                    horizontal=True,
                    key="fa_port_choice",
                )
                lines_to_analyze = (
                    st.session_state.get("A_lines", [])
                    if port_choice == "Client"
                    else st.session_state.get("B_lines", [])
                )
                fee_to_use = (
                    st.session_state.get("FEE_A", 0.0)
                    if port_choice == "Client"
                    else st.session_state.get("FEE_B", 0.0)
                )
                euro_to_use = (
                    st.session_state.get("EURO_RATE_A", 2.0)
                    if port_choice == "Client"
                    else st.session_state.get("EURO_RATE_B", 2.5)
                )
            elif show_client:
                lines_to_analyze = st.session_state.get("A_lines", [])
                fee_to_use = st.session_state.get("FEE_A", 0.0)
                euro_to_use = st.session_state.get("EURO_RATE_A", 2.0)
            else:
                lines_to_analyze = st.session_state.get("B_lines", [])
                fee_to_use = st.session_state.get("FEE_B", 0.0)
                euro_to_use = st.session_state.get("EURO_RATE_B", 2.5)

            uc_lines = [
                ln for ln in lines_to_analyze
                if str(ln.get("isin") or "").upper() not in ("EUROFUND", "STRUCTURED")
            ]
            if not uc_lines:
                st.info(
                    "Aucun fonds UC dans ce portefeuille. "
                    "L'analyse fondamentale ne s'applique pas aux fonds euros "
                    "et produits structurés."
                )
            else:
                if st.button(
                    "🔬 Lancer l'analyse fondamentale",
                    key="fa_run_btn",
                ):
                    with st.spinner(
                        "Chargement des données Morningstar… "
                        "(première exécution : ~5-10 secondes par fonds)"
                    ):
                        agg = aggregate_portfolio_fundamentals(
                            lines_to_analyze,
                            euro_to_use,
                            fee_to_use,
                        )
                    st.session_state["FUND_ANALYSIS_RESULT"] = agg
                    st.rerun()
                fa_agg = st.session_state.get("FUND_ANALYSIS_RESULT")
                if fa_agg and not fa_agg.get("error"):
                    _render_fundamentals_dashboard(fa_agg)


# ============================================================
# MODULE ANALYSE FONDAMENTALE — fonctions module-level
# ============================================================

@st.cache_data(ttl=604800, show_spinner=False)
def _load_fund_fundamentals(isin: str) -> Dict[str, Any]:
    """Charge toutes les données fondamentales d'un fonds via mstarpy. TTL 7 jours."""
    if not MSTARPY_AVAILABLE:
        return {"found": False, "error": True, "isin": isin}
    try:
        fund = mstarpy.Funds(term=isin, pageSize=1)
        if not fund.isin:
            return {"found": False, "error": False, "isin": isin}
        result: Dict[str, Any] = {
            "found": True,
            "error": False,
            "name": getattr(fund, "name", isin),
            "isin": fund.isin,
            "allocation": None,
            "sectors_equity": None,
            "sectors_fi": None,
            "geography": None,
            "style_box": None,
            "credit_quality": None,
            "market_cap": None,
            "esg_score": None,
            "holdings": None,
        }

        # Allocation actions/obligations/cash
        try:
            raw = fund.allocationMap()

            # Stocker raw pour debug si activé (jamais en production)
            if st.session_state.get("FA_DEBUG_MODE", False):
                if "FA_DEBUG_ALLOC" not in st.session_state:
                    st.session_state["FA_DEBUG_ALLOC"] = raw

            def _find_alloc_dict(r: dict) -> dict:
                """Cherche récursivement le premier dict contenant
                AssetAllocEquity avec une valeur non nulle."""
                if not isinstance(r, dict):
                    return {}
                # Vérifier à la racine
                if "AssetAllocEquity" in r:
                    eq = r.get("AssetAllocEquity") or {}
                    if isinstance(eq, dict) and (
                        eq.get("netAllocation") or eq.get("longAllocation")
                    ):
                        return r
                # Vérifier une profondeur
                for v in r.values():
                    if isinstance(v, dict) and "AssetAllocEquity" in v:
                        eq = v.get("AssetAllocEquity") or {}
                        if isinstance(eq, dict) and (
                            eq.get("netAllocation") or eq.get("longAllocation")
                        ):
                            return v
                return {}

            dv = _find_alloc_dict(raw)
            ASSET_KEYS = [
                "AssetAllocEquity", "AssetAllocBond", "AssetAllocCash",
                "AssetAllocOther", "AssetAllocNotClassified",
            ]
            # Essayer netAllocation d'abord
            alloc: Dict[str, float] = {}
            for asset_key in ASSET_KEYS:
                entry = dv.get(asset_key) or {}
                net = entry.get("netAllocation") if isinstance(entry, dict) else None
                try:
                    alloc[asset_key] = float(net) if net is not None else 0.0
                except (TypeError, ValueError):
                    alloc[asset_key] = 0.0
            alloc_sum = sum(alloc.values())
            # Si netAllocation < 80%, essayer longAllocation
            if alloc_sum < 80.0 and dv:
                alloc_long: Dict[str, float] = {}
                for asset_key in ASSET_KEYS:
                    entry = dv.get(asset_key) or {}
                    long_val = entry.get("longAllocation") if isinstance(entry, dict) else None
                    try:
                        alloc_long[asset_key] = float(long_val) if long_val is not None else 0.0
                    except (TypeError, ValueError):
                        alloc_long[asset_key] = 0.0
                long_sum = sum(alloc_long.values())
                if long_sum > alloc_sum:
                    alloc = alloc_long
                    alloc_sum = long_sum
            # Fallback dualViewData si allocationMap toujours vide
            if alloc_sum < 5.0:
                dual = raw.get("dualViewData") or {}
                if not dual:
                    # chercher dualViewData en profondeur
                    for v in raw.values():
                        if isinstance(v, dict) and "marketValueStockNet" in v:
                            dual = v
                            break
                if dual:
                    eq_val = dual.get("marketValueStockNet") or dual.get("NetEquity") or 0.0
                    bo_val = dual.get("marketValueBondNet") or dual.get("NetBond") or 0.0
                    ca_val = dual.get("marketValueCashNet") or dual.get("NetCash") or 0.0
                    try:
                        alloc["AssetAllocEquity"] = float(eq_val)
                        alloc["AssetAllocBond"] = float(bo_val)
                        alloc["AssetAllocCash"] = float(ca_val)
                        alloc_sum = sum(alloc.values())
                    except (TypeError, ValueError):
                        pass
            # Dernier recours : signaler l'absence de données
            if alloc_sum < 5.0:
                result["allocation"] = None
                result["allocation_sum"] = 0.0
            else:
                result["allocation"] = alloc
                result["allocation_sum"] = alloc_sum
        except Exception:
            pass

        # Secteurs actions et obligations
        try:
            sectors = fund.sector()
            if isinstance(sectors, dict):
                eq = (sectors.get("EQUITY") or {}).get("fundPortfolio") or {}
                fi = (sectors.get("FIXEDINCOME") or {}).get("fundPortfolio") or {}
                result["sectors_equity"] = {
                    k: float(v) for k, v in eq.items()
                    if k != "portfolioDate" and isinstance(v, (int, float)) and float(v) > 0
                }
                result["sectors_fi"] = {
                    k: float(v) for k, v in fi.items()
                    if k != "portfolioDate" and isinstance(v, (int, float)) and float(v) > 0
                }
        except Exception:
            pass

        # Géographie
        try:
            geo = fund.regionalSector()
            fp = geo.get("fundPortfolio") or {}
            GEO_KEYS = [
                "northAmerica", "unitedKingdom", "europeDeveloped", "europeEmerging",
                "africaMiddleEast", "japan", "australasia", "asiaDeveloped",
                "asiaEmerging", "latinAmerica",
            ]
            result["geography"] = {k: float(fp.get(k) or 0.0) for k in GEO_KEYS}
        except Exception:
            pass

        # Style box
        try:
            sw = fund.allocationWeighting()
            STYLE_KEYS = [
                "largeValue", "largeBlend", "largeGrowth",
                "middleValue", "middleBlend", "middleGrowth",
                "smallValue", "smallBlend", "smallGrowth",
            ]
            result["style_box"] = {k: float(sw.get(k) or 0.0) for k in STYLE_KEYS}
        except Exception:
            pass

        # Qualité crédit
        try:
            cq = fund.creditQuality()
            fd = cq.get("fund") or {}
            CQ_KEYS = [
                "creditQualityAAA", "creditQualityAA", "creditQualityA", "creditQualityBBB",
                "creditQualityBB", "creditQualityB", "creditQualityBelowB", "creditQualityNotRated",
            ]
            result["credit_quality"] = {k: float(fd.get(k) or 0.0) for k in CQ_KEYS}
        except Exception:
            pass

        # Capitalisation boursière
        try:
            mc = fund.marketCapitalization()
            fd = mc.get("fund") or {}
            result["market_cap"] = {
                "giant":        float(fd.get("giant") or 0.0),
                "large":        float(fd.get("large") or 0.0),
                "medium":       float(fd.get("medium") or 0.0),
                "small":        float(fd.get("small") or 0.0),
                "micro":        float(fd.get("micro") or 0.0),
                "avgMarketCap": float(fd.get("avgMarketCap") or 0.0),
            }
        except Exception:
            pass

        # Score ESG
        try:
            esg = fund.esgRisk()
            score = esg.get("fundSustainabilityScore")
            if score is not None:
                result["esg_score"] = float(score)
        except Exception:
            pass

        # Holdings top 15
        try:
            df_h = fund.holdings(holdingType="all")
            if df_h is not None:
                if not isinstance(df_h, pd.DataFrame):
                    df_h = pd.DataFrame(df_h)
                if not df_h.empty:
                    cols = ["securityName", "weighting", "country", "sector", "holdingType", "isin"]
                    cols_ok = [c for c in cols if c in df_h.columns]
                    result["holdings"] = df_h[cols_ok].head(15).to_dict("records")
        except Exception:
            pass

        return result
    except Exception:
        return {"found": False, "error": True, "isin": isin}


def aggregate_portfolio_fundamentals(
    lines: List[Dict[str, Any]],
    euro_rate: float,
    fee_pct: float,
) -> Dict[str, Any]:
    """Agrège les métriques fondamentales Morningstar de tous les fonds UC du portefeuille."""
    # 1. Calculer la valeur actuelle de chaque ligne UC
    line_values: Dict[str, float] = {}
    for ln in lines:
        isin = str(ln.get("isin") or "").upper()
        if isin in ("EUROFUND", "STRUCTURED"):
            continue
        buy_ts = pd.Timestamp(ln.get("buy_date"))
        net_amt, buy_px, qty = compute_line_metrics(ln, fee_pct, euro_rate)
        dfl, _ = get_series_for_line(ln, buy_ts, euro_rate)
        if dfl.empty or qty <= 0:
            continue
        last_px = float(dfl["Close"].iloc[-1])
        val = qty * last_px
        if val > 0:
            line_id = ln.get("id") or ln.get("isin") or str(id(ln))
            line_values[line_id] = val

    total_val = sum(line_values.values())
    if total_val <= 0:
        return {"covered_pct": 0.0, "not_found": [], "error": "Aucune valeur disponible"}

    # 2. Charger les fondamentaux pour chaque ligne UC
    fund_data: Dict[str, Any] = {}
    weights: Dict[str, float] = {}
    for ln in lines:
        isin = str(ln.get("isin") or "").upper()
        if isin in ("EUROFUND", "STRUCTURED"):
            continue
        line_id = ln.get("id") or ln.get("isin") or str(id(ln))
        if line_id not in line_values:
            continue
        w = line_values[line_id] / total_val
        weights[line_id] = w
        data = _load_fund_fundamentals(isin)
        fund_data[line_id] = {**data, "line_name": ln.get("name", isin), "weight": w}

    # 3. Calculer covered_pct et not_found
    covered_weight = sum(
        weights[lid] for lid, d in fund_data.items() if d.get("found")
    )
    not_found = [
        {"name": d.get("line_name", "—"), "isin": d.get("isin", "—")}
        for d in fund_data.values()
        if not d.get("found")
    ]

    # 4. Agréger chaque dimension (pondération par poids fonds)
    def _agg(key: str, sub_key: Optional[str] = None) -> Dict[str, float]:
        result_d: Dict[str, float] = {}
        total_w = 0.0
        for lid, d in fund_data.items():
            if not d.get("found"):
                continue
            w = weights.get(lid, 0.0)
            src = d.get(key) or {}
            if sub_key:
                src = src.get(sub_key) or {}
            if not src:
                continue
            for k, v in src.items():
                try:
                    result_d[k] = result_d.get(k, 0.0) + float(v) * w
                except (TypeError, ValueError):
                    pass
            total_w += w
        # Renormaliser si couverture partielle
        if 0 < total_w < 1.0:
            result_d = {k: v / total_w for k, v in result_d.items()}
        return result_d

    allocation_agg = _agg("allocation")
    sectors_equity_agg = _agg("sectors_equity")
    sectors_fi_agg = _agg("sectors_fi")
    geography_agg = _agg("geography")
    style_box_agg = _agg("style_box")
    credit_quality_agg = _agg("credit_quality")
    market_cap_agg = _agg("market_cap")

    # ESG : moyenne pondérée
    esg_num = 0.0
    esg_w = 0.0
    for lid, d in fund_data.items():
        if d.get("found") and d.get("esg_score") is not None:
            w = weights.get(lid, 0.0)
            esg_num += float(d["esg_score"]) * w
            esg_w += w
    esg_agg: Optional[float] = (esg_num / esg_w) if esg_w > 0 else None

    # Holdings consolidés : agréger par ISIN ou par nom si ISIN absent
    holdings_agg: Dict[str, Dict[str, Any]] = {}
    for lid, d in fund_data.items():
        if not d.get("found") or not d.get("holdings"):
            continue
        w_fund = weights.get(lid, 0.0)
        for h in d["holdings"]:
            h_isin = str(h.get("isin") or "").strip()
            h_name = str(h.get("securityName") or "").strip()
            h_key = h_isin if h_isin else h_name
            if not h_key:
                continue
            h_weight_in_fund = float(h.get("weighting") or 0.0)
            contribution = h_weight_in_fund * w_fund / 100.0
            if h_key in holdings_agg:
                holdings_agg[h_key]["weight_portfolio"] += contribution
                if lid not in holdings_agg[h_key]["found_in"]:
                    holdings_agg[h_key]["found_in"].append(d.get("line_name", lid))
            else:
                holdings_agg[h_key] = {
                    "name": h_name,
                    "isin": h_isin,
                    "weight_portfolio": contribution,
                    "country": h.get("country", ""),
                    "sector": h.get("sector", ""),
                    "found_in": [d.get("line_name", lid)],
                }

    top_holdings = sorted(
        holdings_agg.values(),
        key=lambda x: x["weight_portfolio"],
        reverse=True,
    )[:20]

    return {
        "covered_pct":    covered_weight * 100.0,
        "not_found":      not_found,
        "allocation":     allocation_agg,
        "sectors_equity": sectors_equity_agg,
        "sectors_fi":     sectors_fi_agg,
        "geography":      geography_agg,
        "style_box":      style_box_agg,
        "credit_quality": credit_quality_agg,
        "market_cap":     market_cap_agg,
        "esg_score":      esg_agg,
        "top_holdings":   top_holdings,
    }


def _render_fundamentals_dashboard(agg: Dict[str, Any]) -> None:
    """Affiche le tableau de bord d'analyse fondamentale agrégée."""
    # Bandeau couverture
    cov = agg.get("covered_pct", 0.0)
    not_found = agg.get("not_found", [])
    try:
        if cov >= 80:
            st.success(f"✅ Couverture Morningstar : {cov:.0f}% du portefeuille analysé")
        elif cov >= 50:
            st.warning(f"⚠️ Couverture partielle : {cov:.0f}% du portefeuille analysé")
        else:
            st.error(f"❌ Couverture insuffisante : {cov:.0f}% — résultats peu représentatifs")
        if not_found:
            nf_list = ", ".join(f"{d['name']} ({d['isin']})" for d in not_found)
            st.caption(f"Fonds non couverts par Morningstar : {nf_list}")
    except Exception:
        pass

    st.markdown("---")

    # ── Section 1 : Allocation d'actifs ──────────────────────────────
    try:
        alloc = agg.get("allocation") or {}
        if alloc:
            st.subheader("📊 Allocation d'actifs consolidée")
            _FA_ALLOC_LABELS = {
                "AssetAllocEquity":         "Actions",
                "AssetAllocBond":           "Obligations",
                "AssetAllocCash":           "Cash",
                "AssetAllocOther":          "Autres",
                "AssetAllocNotClassified":  "Non classifié",
            }
            alloc_display = {
                _FA_ALLOC_LABELS.get(k, k): v for k, v in alloc.items() if v > 0.5
            }
            if alloc_display:
                cols_alloc = st.columns(len(alloc_display))
                for i, (label, val) in enumerate(alloc_display.items()):
                    cols_alloc[i].metric(label, f"{val:.1f}%")
                # FIXED (P3): note explicative si le total < 95%
                alloc_total = sum(alloc_display.values())
                if alloc_total < 95.0:
                    st.caption(
                        f"ℹ️ Total affiché : {alloc_total:.1f}% — Les positions nettes "
                        "(long − short) retournées par Morningstar peuvent ne pas sommer "
                        "à 100% pour les fonds utilisant des dérivés ou positions synthétiques. "
                        "Les pourcentages reflètent l'exposition nette réelle de chaque classe d'actif."
                    )
            st.markdown("---")
    except Exception:
        pass

    # ── Section 2 : Répartition géographique ─────────────────────────
    try:
        geo = agg.get("geography") or {}
        if geo:
            st.subheader("🌍 Répartition géographique consolidée")
            _FA_GEO_LABELS = {
                "northAmerica":     "Amérique du Nord",
                "europeDeveloped":  "Europe développée",
                "asiaDeveloped":    "Asie développée",
                "asiaEmerging":     "Asie émergente",
                "japan":            "Japon",
                "latinAmerica":     "Amérique latine",
                "unitedKingdom":    "Royaume-Uni",
                "europeEmerging":   "Europe émergente",
                "africaMiddleEast": "Afrique / Moyen-Orient",
                "australasia":      "Australasie",
            }
            geo_filtered = {
                _FA_GEO_LABELS.get(k, k): v for k, v in geo.items() if v > 0.5
            }
            if geo_filtered:
                sorted_geo = sorted(geo_filtered.items(), key=lambda x: x[1], reverse=True)
                _fs_bar_chart(
                    [x[0] for x in sorted_geo],
                    [x[1] for x in sorted_geo],
                    "#2196F3",
                    "Géographie",
                )
            st.markdown("---")
    except Exception:
        pass

    # ── Section 3 : Secteurs actions et obligations ───────────────────
    try:
        col_sec1, col_sec2 = st.columns(2)
        with col_sec1:
            sec_eq = agg.get("sectors_equity") or {}
            if sec_eq:
                st.subheader("📈 Secteurs actions")
                _fs_bar_chart(
                    [_FS_SECTOR_LABELS.get(k, k) for k in sec_eq],
                    list(sec_eq.values()),
                    "#1f77b4",
                    "Secteurs actions",
                )
        with col_sec2:
            sec_fi = agg.get("sectors_fi") or {}
            if sec_fi:
                st.subheader("📉 Secteurs obligations")
                _fs_bar_chart(
                    [_FS_FI_LABELS.get(k, k) for k in sec_fi],
                    list(sec_fi.values()),
                    "#ff7f0e",
                    "Secteurs obligations",
                )
        st.markdown("---")
    except Exception:
        pass

    # ── Section 4 : Top 20 holdings consolidés ───────────────────────
    try:
        top_h = agg.get("top_holdings") or []
        if top_h:
            st.subheader("🏦 Top 20 positions consolidées")
            st.caption(
                "Poids calculé : (poids du holding dans le fonds) × "
                "(poids du fonds dans le portefeuille). "
                "Les doublons (même titre dans plusieurs fonds) sont agrégés."
            )
            rows_h = [
                {
                    "Titre":              h.get("name", "—"),
                    "ISIN":               h.get("isin", "—"),
                    "Poids portef. (%)":  round(h["weight_portfolio"] * 100, 3),
                    "Pays":               h.get("country", "—"),
                    "Secteur":            h.get("sector", "—"),
                    "Présent dans":       ", ".join(h.get("found_in", [])),
                }
                for h in top_h
            ]
            df_top = pd.DataFrame(rows_h)
            overweight = df_top[df_top["Poids portef. (%)"] > 5.0]
            if not overweight.empty:
                names_ow = ", ".join(overweight["Titre"].tolist())
                st.warning(
                    f"⚠️ Surconcentration détectée (>5% du portefeuille) : {names_ow}"
                )
            st.dataframe(df_top, hide_index=True, use_container_width=True)
            st.markdown("---")
    except Exception:
        pass

    # ── Section 5 : Style box + Capitalisation ────────────────────────
    try:
        col_style, col_cap = st.columns(2)
        with col_style:
            sb = agg.get("style_box") or {}
            if sb:
                st.subheader("🎯 Style box actions")
                _STYLE_LABELS = {
                    "largeValue":   "Large Value",  "largeBlend":   "Large Blend",  "largeGrowth":   "Large Growth",
                    "middleValue":  "Mid Value",    "middleBlend":  "Mid Blend",    "middleGrowth":  "Mid Growth",
                    "smallValue":   "Small Value",  "smallBlend":   "Small Blend",  "smallGrowth":   "Small Growth",
                }
                sb_display = {
                    _STYLE_LABELS.get(k, k): v for k, v in sb.items() if v > 1.0
                }
                if sb_display:
                    sorted_sb = sorted(sb_display.items(), key=lambda x: x[1], reverse=True)
                    _fs_bar_chart(
                        [x[0] for x in sorted_sb],
                        [x[1] for x in sorted_sb],
                        "#9C27B0",
                        "Style box",
                    )
        with col_cap:
            mc = agg.get("market_cap") or {}
            if mc:
                st.subheader("📏 Capitalisation boursière")
                _CAP_LABELS = {
                    "giant": "Giant", "large": "Large", "medium": "Medium",
                    "small": "Small", "micro": "Micro",
                }
                mc_display = {
                    _CAP_LABELS.get(k, k): v for k, v in mc.items()
                    if k != "avgMarketCap" and v > 0.5
                }
                if mc_display:
                    sorted_mc = sorted(mc_display.items(), key=lambda x: x[1], reverse=True)
                    _fs_bar_chart(
                        [x[0] for x in sorted_mc],
                        [x[1] for x in sorted_mc],
                        "#4CAF50",
                        "Cap boursière",
                    )
                avg = mc.get("avgMarketCap")
                if avg and avg > 0:
                    st.caption(f"Capitalisation moyenne pondérée : {avg / 1000:.0f} Md$")
        st.markdown("---")
    except Exception:
        pass

    # ── Section 6 : Qualité crédit ────────────────────────────────────
    try:
        cq = agg.get("credit_quality") or {}
        if cq and sum(cq.values()) > 1.0:
            st.subheader("🏦 Qualité crédit obligataire")
            _CQ_LABELS = {
                "creditQualityAAA":      "AAA",
                "creditQualityAA":       "AA",
                "creditQualityA":        "A",
                "creditQualityBBB":      "BBB",
                "creditQualityBB":       "BB",
                "creditQualityB":        "B",
                "creditQualityBelowB":   "Below B",
                "creditQualityNotRated": "Non noté",
            }
            cq_display = {_CQ_LABELS.get(k, k): v for k, v in cq.items() if v > 0.5}
            if cq_display:
                sorted_cq = sorted(cq_display.items(), key=lambda x: x[1], reverse=True)
                _fs_bar_chart(
                    [x[0] for x in sorted_cq],
                    [x[1] for x in sorted_cq],
                    "#FF5722",
                    "Qualité crédit",
                )
    except Exception:
        pass

    # ── Section 7 : Score ESG agrégé ──────────────────────────────────
    try:
        esg = agg.get("esg_score")
        if esg is not None:
            st.markdown("---")
            st.subheader("🌱 Score ESG agrégé (Morningstar Sustainalytics)")
            if esg <= 10:
                esg_cat = "Négligeable"
            elif esg <= 20:
                esg_cat = "Faible"
            elif esg <= 30:
                esg_cat = "Moyen"
            elif esg <= 40:
                esg_cat = "Élevé"
            else:
                esg_cat = "Sévère"
            ec1, ec2 = st.columns([1, 3])
            ec1.metric("Score ESG moyen pondéré", f"{esg:.1f}")
            ec2.caption(
                f"Catégorie de risque : {esg_cat}\n\n"
                "Score calculé par moyenne pondérée des scores ESG "
                "Sustainalytics de chaque fonds."
            )
    except Exception:
        pass


def run_comparator():
    render_app(run_page_config=False)


# ============================================================
# MODULE FISCALITÉ ASSURANCE-VIE — fonctions de calcul pures
# ============================================================

def calc_quote_part_gains(
    valeur_contrat: float,
    versements_nets: float,
    montant_rachat: float,
) -> float:
    """Quote-part de gains dans le rachat."""
    if valeur_contrat <= 0 or versements_nets >= valeur_contrat:
        return 0.0
    return max(0.0, montant_rachat * (1.0 - versements_nets / valeur_contrat))


def calc_imposition_rachat(
    gains: float,
    anciennete_annees: float,
    situation_familiale: str,
    versements_nets_total: float,
    montant_rachat: float,
    option_ir: bool = False,
) -> Dict[str, Any]:
    PS_RATE = 0.172
    PFU_RATE = 0.128
    PFL_LOW_RATE = 0.075
    PFL_HIGH_RATE = 0.128
    ABATTEMENT_SEUL = 4_600.0
    ABATTEMENT_COUPLE = 9_200.0
    SEUIL_150K = 150_000.0

    abattement = 0.0
    base_ir = gains
    taux_ir = 0.0
    regime = ""
    tmi_applicable = ""

    if anciennete_annees < 8:
        taux_ir = PFU_RATE
        regime = "PFU 30% (contrat < 8 ans)"
        tmi_applicable = "Flat Tax 12,8% IR + 17,2% PS"
        if option_ir:
            tmi_applicable = "Option barème IR — à déclarer case 2CH"
    else:
        abattement = ABATTEMENT_COUPLE if "Couple" in situation_familiale else ABATTEMENT_SEUL
        base_ir = max(0.0, gains - abattement)
        if versements_nets_total <= SEUIL_150K:
            taux_ir = PFL_LOW_RATE
            regime = "PFL 7,5% + PS (contrat ≥8 ans, versements ≤150k€)"
            tmi_applicable = "Prélèvement forfaitaire libératoire 7,5%"
        else:
            taux_ir = PFL_HIGH_RATE
            regime = "PFU 12,8% + PS (contrat ≥8 ans, versements >150k€)"
            tmi_applicable = "Flat Tax 12,8% (versements post-27/09/2017 >150k€)"

    montant_ir = base_ir * taux_ir
    montant_ps = gains * PS_RATE

    return {
        "gains": gains,
        "regime": regime,
        "base_ir": base_ir,
        "abattement_applique": abattement,
        "taux_ir": taux_ir,
        "montant_ir": montant_ir,
        "montant_ps": montant_ps,
        "total_impots": montant_ir + montant_ps,
        "net_percu": montant_rachat - (montant_ir + montant_ps),
        "tmi_applicable": tmi_applicable,
    }


def calc_rachat_depuis_net(
    montant_net_souhaite: float,
    valeur_contrat: float,
    versements_nets: float,
    anciennete_annees: float,
    situation_familiale: str,
    versements_nets_total: float,
) -> Tuple[float, Dict[str, Any]]:
    """Bisection : trouve le brut tel que brut - impôts == net_souhaité."""
    if montant_net_souhaite <= 0:
        return 0.0, calc_imposition_rachat(0.0, anciennete_annees, situation_familiale, versements_nets_total, 0.0)

    lo, hi = montant_net_souhaite, min(valeur_contrat, montant_net_souhaite * 2.5)
    # s'assurer que hi couvre bien le cas
    for _ in range(50):
        g = calc_quote_part_gains(valeur_contrat, versements_nets, hi)
        d = calc_imposition_rachat(g, anciennete_annees, situation_familiale, versements_nets_total, hi)
        if hi - d["total_impots"] >= montant_net_souhaite:
            break
        hi = min(valeur_contrat, hi * 1.5)

    for _ in range(1000):
        mid = (lo + hi) / 2.0
        g = calc_quote_part_gains(valeur_contrat, versements_nets, mid)
        d = calc_imposition_rachat(g, anciennete_annees, situation_familiale, versements_nets_total, mid)
        net = mid - d["total_impots"]
        if abs(net - montant_net_souhaite) < 0.01:
            return mid, d
        if net < montant_net_souhaite:
            lo = mid
        else:
            hi = mid
    return mid, d


def calc_optimisation_abattement(
    valeur_contrat: float,
    versements_nets: float,
    anciennete_annees: float,
    situation_familiale: str,
    versements_nets_total: float,
    abattement_deja_utilise: float = 0.0,
) -> Dict[str, Any]:
    PS_RATE = 0.172
    ABATTEMENT_SEUL = 4_600.0
    ABATTEMENT_COUPLE = 9_200.0

    abattement = ABATTEMENT_COUPLE if "Couple" in situation_familiale else ABATTEMENT_SEUL
    abattement_restant = max(0.0, abattement - abattement_deja_utilise)

    if valeur_contrat <= 0 or versements_nets >= valeur_contrat:
        return {
            "rachat_optimal": 0.0, "gains_realises": 0.0,
            "ir_du": 0.0, "ps_du": 0.0, "economies_ir": 0.0,
            "abattement_utilise": 0.0, "abattement_restant_apres": abattement_restant,
        }

    taux_pv = 1.0 - versements_nets / valeur_contrat
    if taux_pv <= 0:
        return {
            "rachat_optimal": 0.0, "gains_realises": 0.0,
            "ir_du": 0.0, "ps_du": 0.0, "economies_ir": 0.0,
            "abattement_utilise": 0.0, "abattement_restant_apres": abattement_restant,
        }

    rachat_optimal = abattement_restant / taux_pv
    rachat_optimal = min(rachat_optimal, valeur_contrat)
    gains_realises = calc_quote_part_gains(valeur_contrat, versements_nets, rachat_optimal)

    # IR si on avait dépassé l'abattement (scénario sans optimisation)
    SEUIL_150K = 150_000.0
    taux_ir = 0.075 if versements_nets_total <= SEUIL_150K else 0.128
    economies_ir = gains_realises * taux_ir

    return {
        "rachat_optimal": rachat_optimal,
        "gains_realises": gains_realises,
        "ir_du": 0.0,
        "ps_du": gains_realises * PS_RATE,
        "economies_ir": economies_ir,
        "abattement_utilise": gains_realises,
        "abattement_restant_apres": 0.0,
    }


def calc_transmission_990I(
    capital_deces: float,
    nb_beneficiaires: int,
    parts_pct: List[float],
    types_beneficiaires: List[str],
) -> List[Dict[str, Any]]:
    ABATTEMENT_990I = 152_500.0
    TAUX_990I_T1 = 0.20
    TAUX_990I_T2 = 0.3125
    SEUIL_990I_T2 = 700_000.0

    results = []
    for i in range(nb_beneficiaires):
        t = types_beneficiaires[i] if i < len(types_beneficiaires) else "Enfant"
        p = parts_pct[i] if i < len(parts_pct) else 0.0
        part_eur = capital_deces * p / 100.0
        note = ""

        if "Conjoint" in t or "PACS" in t:
            taxe = 0.0
            taxable = 0.0
            abat = part_eur
        else:
            if "Démembré" in t:
                note = (
                    "Le démembrement de la clause bénéficiaire nécessite une évaluation "
                    "actuarielle des droits d'usufruit et nue-propriété selon le barème "
                    "fiscal de l'article 669 CGI (fonction de l'âge de l'usufruitier). "
                    "Cette simulation affiche la taxation globale sur la part démembrée "
                    "comme si elle revenait à un bénéficiaire standard, à affiner avec "
                    "un notaire."
                )
            abat = min(part_eur, ABATTEMENT_990I)
            taxable = max(0.0, part_eur - ABATTEMENT_990I)
            if taxable <= SEUIL_990I_T2:
                taxe = taxable * TAUX_990I_T1
            else:
                taxe = SEUIL_990I_T2 * TAUX_990I_T1 + (taxable - SEUIL_990I_T2) * TAUX_990I_T2

        results.append({
            "index": i + 1,
            "type": t,
            "part_pct": p,
            "part_eur": part_eur,
            "abattement": abat,
            "taxable": taxable if "Conjoint" not in t and "PACS" not in t else 0.0,
            "taxe": taxe,
            "net_recu": part_eur - taxe,
            "note": note,
        })
    return results


def calc_transmission_757B(
    primes_versees_apres_70: float,
    nb_beneficiaires: int,
    parts_pct: List[float],
    types_beneficiaires: List[str],
    capital_deces: float,
    liens_succession: List[str],
) -> List[Dict[str, Any]]:
    ABATTEMENT_757B = 30_500.0
    TAUX_LIEN = {
        "Enfant": 0.20,
        "Frère/Sœur": 0.40,
        "Neveu/Nièce": 0.55,
        "Tiers": 0.60,
        "Conjoint/PACS": 0.0,
    }

    # Parts des non-conjoints pour répartir l'abattement global
    non_conjoint_total_pct = sum(
        parts_pct[i] for i in range(nb_beneficiaires)
        if i < len(types_beneficiaires)
        and "Conjoint" not in types_beneficiaires[i]
        and "PACS" not in types_beneficiaires[i]
    )

    results = []
    for i in range(nb_beneficiaires):
        t = types_beneficiaires[i] if i < len(types_beneficiaires) else "Enfant"
        p = parts_pct[i] if i < len(parts_pct) else 0.0
        lien = liens_succession[i] if i < len(liens_succession) else "Enfant"
        part_eur = capital_deces * p / 100.0

        if "Conjoint" in t or "PACS" in t or lien == "Conjoint/PACS":
            results.append({
                "index": i + 1, "type": t, "part_pct": p, "part_eur": part_eur,
                "primes_part": 0.0, "abattement_part": 0.0,
                "taxable": 0.0, "taxe": 0.0, "net_recu": part_eur,
                "taux_applique": 0.0,
            })
            continue

        # Répartition des primes post-70 ans selon la quote-part
        primes_part = primes_versees_apres_70 * p / max(non_conjoint_total_pct, 1.0)
        # Abattement global 30 500 € réparti proportionnellement
        abat_part = ABATTEMENT_757B * p / max(non_conjoint_total_pct, 1.0)
        taxable = max(0.0, primes_part - abat_part)
        taux = TAUX_LIEN.get(lien, 0.60)
        taxe = taxable * taux

        results.append({
            "index": i + 1, "type": t, "part_pct": p, "part_eur": part_eur,
            "primes_part": primes_part, "abattement_part": abat_part,
            "taxable": taxable, "taxe": taxe, "net_recu": part_eur - taxe,
            "taux_applique": taux,
        })
    return results


# ============================================================
# MODULE FISCALITÉ ASSURANCE-VIE — onglets UI
# ============================================================

def _fmt_eur(v: float) -> str:
    """Formatage monétaire uniforme pour le module fiscal."""
    return f"{v:,.0f} €".replace(",", "\u202f")


def _tab_rachat() -> None:
    with st.expander("Paramètres du contrat", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            date_ouverture = st.date_input(
                "Date d'ouverture du contrat",
                value=st.session_state.get("tax_date_ouverture", date(2016, 1, 2)),
                key="tax_date_ouverture",
            )
            anciennete_jours = (date.today() - date_ouverture).days
            anciennete_annees = anciennete_jours / 365.25
            st.caption(f"Ancienneté : **{anciennete_annees:.1f} ans**")

            valeur_contrat = st.number_input(
                "Valeur actuelle du contrat (€)",
                min_value=0.0, max_value=10_000_000.0,
                value=float(st.session_state.get("tax_valeur_contrat", 100_000.0)),
                step=1_000.0, key="tax_valeur_contrat",
            )
            versements_nets = st.number_input(
                "Total versements nets (€)",
                min_value=0.0,
                max_value=float(max(valeur_contrat, 0.01)),
                value=float(min(
                    st.session_state.get("tax_versements_nets", 80_000.0),
                    max(valeur_contrat, 0.0),
                )),
                step=1_000.0, key="tax_versements_nets",
                help="Versements bruts − rachats antérieurs en capital",
            )
        with c2:
            versements_nets_total = st.number_input(
                "Total versements nets tous contrats (pour seuil 150k€) (€)",
                min_value=0.0, max_value=10_000_000.0,
                value=float(st.session_state.get("tax_versements_nets_total", 80_000.0)),
                step=1_000.0, key="tax_versements_nets_total",
                help="Utilisé pour déterminer si le taux IR est 7,5% ou 12,8% sur les contrats ≥8 ans avec versements post-27/09/2017",
            )
            situation_familiale = st.radio(
                "Situation familiale",
                ["Célibataire / veuf / divorcé", "Couple (imposition commune)"],
                key="tax_situation_familiale",
                horizontal=False,
            )
            ps_deja_preleves = st.number_input(
                "PS déjà prélevés sur fonds euros (€)",
                min_value=0.0, max_value=100_000.0,
                value=0.0, step=100.0,
                help="Sur les fonds en euros, les PS sont prélevés chaque année par l'assureur. Indiquez le cumul déjà prélevé pour éviter de les recompter au rachat.",
            )

    st.markdown("---")
    st.markdown("#### Montant du rachat")

    # Synchronisation brut ↔ net via session_state
    last_edited = st.session_state.get("tax_last_edited", "brut")

    col_brut, col_net = st.columns(2)
    with col_brut:
        brut_default = float(st.session_state.get("tax_montant_brut", 10_000.0))
        montant_brut = st.number_input(
            "Montant brut du rachat (€)",
            min_value=0.0,
            max_value=float(max(valeur_contrat, 0.01)),
            value=brut_default,
            step=500.0,
            key="_tax_brut_input",
        )
        if montant_brut != brut_default:
            st.session_state["tax_last_edited"] = "brut"
            st.session_state["tax_montant_brut"] = montant_brut

    with col_net:
        net_default = float(st.session_state.get("tax_montant_net", 9_000.0))
        montant_net_input = st.number_input(
            "Montant net souhaité (€)",
            min_value=0.0,
            max_value=float(max(valeur_contrat, 0.01)),
            value=net_default,
            step=500.0,
            key="_tax_net_input",
        )
        if montant_net_input != net_default:
            st.session_state["tax_last_edited"] = "net"
            st.session_state["tax_montant_net"] = montant_net_input

    # Résoudre le champ non-prioritaire
    if st.session_state.get("tax_last_edited", "brut") == "net":
        montant_brut, _ = calc_rachat_depuis_net(
            montant_net_input, valeur_contrat, versements_nets,
            anciennete_annees, situation_familiale, versements_nets_total,
        )
    else:
        montant_brut = st.session_state.get("tax_montant_brut", montant_brut)

    # Calcul et affichage
    if valeur_contrat > 0 and versements_nets > 0 and montant_brut > 0:
        gains = calc_quote_part_gains(valeur_contrat, versements_nets, montant_brut)
        result = calc_imposition_rachat(
            gains, anciennete_annees, situation_familiale,
            versements_nets_total, montant_brut,
        )
        st.session_state["tax_rachat_result"] = result

        # Bandeau régime
        if anciennete_annees < 8:
            st.error(f"⏱️ Contrat de {anciennete_annees:.1f} an(s) — Régime PFU 30%")
        else:
            st.success(f"✅ Contrat de {anciennete_annees:.1f} ans — Régime favorable ≥8 ans")

        # Tableau récapitulatif
        capital_pur = montant_brut - gains
        abat = result["abattement_applique"]
        ps_net = max(0.0, result["montant_ps"] - ps_deja_preleves)
        total_impots_net = result["montant_ir"] + ps_net
        net_percu_net = montant_brut - total_impots_net

        rows_display = [
            ("Montant brut du rachat", _fmt_eur(montant_brut)),
            ("  dont quote-part gains", _fmt_eur(gains)),
            ("  dont quote-part capital", _fmt_eur(capital_pur)),
            ("─" * 35, ""),
            (f"Abattement IR appliqué", _fmt_eur(abat)),
            ("Base imposable IR", _fmt_eur(result["base_ir"])),
            (f"IR dû ({result['taux_ir']*100:.1f}%)", _fmt_eur(result["montant_ir"])),
            ("Prélèvements sociaux (17,2%)", _fmt_eur(result["montant_ps"])),
        ]
        if ps_deja_preleves > 0:
            rows_display.append((f"PS déjà prélevés (fonds €)", f"− {_fmt_eur(ps_deja_preleves)}"))
        rows_display += [
            ("─" * 35, ""),
            ("TOTAL prélèvements", _fmt_eur(total_impots_net)),
            ("MONTANT NET PERÇU", _fmt_eur(net_percu_net)),
        ]
        df_recap = pd.DataFrame(rows_display, columns=["", "Montant"])
        st.dataframe(df_recap, hide_index=True, use_container_width=True)

        st.info(
            f"**Régime appliqué :** {result['regime']}\n\n"
            f"**Taux :** {result['tmi_applicable']}"
        )

        # Conseil stratégique pour contrats ≥8 ans
        if anciennete_annees >= 8 and gains > 0:
            abat_total = result["abattement_applique"]
            rachat_capital_pur = versements_nets / (valeur_contrat / montant_brut) if valeur_contrat > 0 else 0
            brut_sans_ir, _ = calc_rachat_depuis_net(
                capital_pur, valeur_contrat, versements_nets,
                anciennete_annees, situation_familiale, versements_nets_total,
            ) if abat_total > 0 else (capital_pur, {})
            g_limite = calc_quote_part_gains(valeur_contrat, versements_nets, brut_sans_ir)
            ps_limite = g_limite * 0.172
            st.caption(
                f"💡 **Stratégie :** un rachat limité à {_fmt_eur(brut_sans_ir)} "
                f"(capital pur + abattement) ne génèrerait aucun IR, "
                f"uniquement {_fmt_eur(ps_limite)} de PS."
            )


def _tab_optimisation_abattement() -> None:
    st.markdown("#### Optimisation de l'abattement annuel")
    st.info(
        "L'abattement de 4 600 € (célibataire) ou 9 200 € (couple) est global "
        "tous contrats d'assurance-vie. Cette simulation calcule le rachat optimal "
        "pour 'purger' les plus-values progressivement sans payer d'IR."
    )

    c1, c2 = st.columns(2)
    with c1:
        date_ouv2 = st.date_input(
            "Date d'ouverture du contrat",
            value=st.session_state.get("tax_date_ouverture", date(2016, 1, 2)),
            key="tax_opt_date_ouverture",
        )
        anc2 = (date.today() - date_ouv2).days / 365.25
        st.caption(f"Ancienneté : **{anc2:.1f} ans**")

        valeur2 = st.number_input(
            "Valeur actuelle du contrat (€)",
            min_value=0.0, max_value=10_000_000.0,
            value=float(st.session_state.get("tax_valeur_contrat", 100_000.0)),
            step=1_000.0, key="tax_opt_valeur",
        )
        versements2 = st.number_input(
            "Total versements nets (€)",
            min_value=0.0, max_value=float(max(valeur2, 0.01)),
            value=float(min(st.session_state.get("tax_versements_nets", 80_000.0), max(valeur2, 0.0))),
            step=1_000.0, key="tax_opt_versements",
        )
    with c2:
        sit2 = st.radio(
            "Situation familiale",
            ["Célibataire / veuf / divorcé", "Couple (imposition commune)"],
            key="tax_opt_situation",
        )
        versements_total2 = st.number_input(
            "Total versements nets tous contrats (€)",
            min_value=0.0, max_value=10_000_000.0,
            value=float(st.session_state.get("tax_versements_nets_total", 80_000.0)),
            step=1_000.0, key="tax_opt_versements_total",
        )
        abat_deja = st.number_input(
            "Abattement déjà utilisé cette année sur d'autres contrats (€)",
            min_value=0.0, max_value=9_200.0,
            value=float(st.session_state.get("tax_abattement_deja_utilise", 0.0)),
            step=100.0, key="tax_abattement_deja_utilise",
            help="L'abattement de 4 600€/9 200€ est global tous contrats AV. Indiquez ce qui a déjà été consommé.",
        )

    st.markdown("---")

    if anc2 < 8:
        st.warning(
            f"Ce contrat n'a pas encore atteint 8 ans ({anc2:.1f} ans). "
            "L'optimisation par abattement ne s'applique pas."
        )
        return

    if valeur2 <= 0 or versements2 >= valeur2:
        st.info("Pas de plus-value latente — aucun calcul nécessaire.")
        return

    res = calc_optimisation_abattement(
        valeur2, versements2, anc2, sit2, versements_total2, abat_deja,
    )

    abat_total = 9_200.0 if "Couple" in sit2 else 4_600.0
    abat_restant = max(0.0, abat_total - abat_deja)

    mc1, mc2, mc3 = st.columns(3)
    mc1.metric("Abattement disponible restant", _fmt_eur(abat_restant))
    mc2.metric("Rachat recommandé", _fmt_eur(res["rachat_optimal"]))
    mc3.metric("IR dû", _fmt_eur(res["ir_du"]))

    mc4, mc5 = st.columns(2)
    mc4.metric("PS dus (inévitables)", _fmt_eur(res["ps_du"]))
    mc5.metric("Économie IR vs rachat non optimisé", _fmt_eur(res["economies_ir"]))

    # Graphique
    chart_data = pd.DataFrame({
        "Catégorie": ["Capital récupéré", "PS dus", "IR économisé"],
        "Montant (€)": [
            res["rachat_optimal"] - res["ps_du"],
            res["ps_du"],
            res["economies_ir"],
        ],
    })
    if MATPLOTLIB_AVAILABLE:
        fig, ax = plt.subplots(figsize=(5, 2.8))
        colors_bar = ["#2ecc71", "#e67e22", "#3498db"]
        bars = ax.bar(chart_data["Catégorie"], chart_data["Montant (€)"], color=colors_bar)
        ax.bar_label(bars, fmt="%.0f €", padding=3, fontsize=9)
        ax.set_ylabel("€")
        ax.set_title("Optimisation abattement — répartition")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        st.pyplot(fig)
        plt.close(fig)
    else:
        st.bar_chart(chart_data.set_index("Catégorie"))

    # Projection pluriannuelle
    nb_annees = st.slider(
        "Nombre d'années de la stratégie",
        min_value=1, max_value=30,
        value=int(st.session_state.get("tax_nb_annees", 10)),
        key="tax_nb_annees",
    )
    economie_cumulee = res["economies_ir"] * nb_annees
    st.info(
        f"📅 Cette optimisation se renouvelle chaque année civile. "
        f"Sur **{nb_annees} années**, l'économie d'IR cumulée estimée serait de "
        f"**{_fmt_eur(economie_cumulee)}** (hypothèse : situation stable)."
    )


def _tab_transmission() -> None:
    st.markdown("#### Fiscalité au décès — Clause bénéficiaire")

    regime_versements = st.radio(
        "Les versements ont été effectués :",
        [
            "Principalement avant 70 ans (Art. 990I)",
            "Principalement après 70 ans (Art. 757B)",
            "Mix avant et après 70 ans",
        ],
        key="tax_regime_versements",
    )

    c1, c2 = st.columns(2)
    with c1:
        capital_deces = st.number_input(
            "Valeur du contrat au décès (€)",
            min_value=0.0, max_value=10_000_000.0,
            value=float(st.session_state.get("tax_capital_deces", 200_000.0)),
            step=5_000.0, key="tax_capital_deces",
        )
    with c2:
        primes_apres_70 = 0.0
        if "757B" in regime_versements or "Mix" in regime_versements:
            primes_apres_70 = st.number_input(
                "Primes versées après 70 ans (€)",
                min_value=0.0, max_value=capital_deces,
                value=0.0, step=5_000.0,
                help="Seules les primes (pas les gains) au-delà de 30 500€ sont soumises aux droits de succession.",
            )

    st.markdown("---")
    nb_benef = int(st.number_input(
        "Nombre de bénéficiaires",
        min_value=1, max_value=5,
        value=int(st.session_state.get("tax_nb_beneficiaires", 2)),
        step=1, key="tax_nb_beneficiaires",
    ))

    noms, types_benef, parts_pct, liens_succ = [], [], [], []
    TYPES_OPTS = [
        "Conjoint/PACS", "Enfant", "Frère/Sœur", "Neveu/Nièce",
        "Tiers", "Démembré - Usufruitier", "Démembré - Nu-propriétaire",
    ]
    LIENS_OPTS = ["Conjoint/PACS", "Enfant", "Frère/Sœur", "Neveu/Nièce", "Tiers"]

    for i in range(nb_benef):
        with st.expander(f"Bénéficiaire {i + 1}", expanded=(i == 0)):
            bc1, bc2, bc3 = st.columns([2, 2, 1])
            with bc1:
                nom = st.text_input("Nom / prénom (facultatif)", key=f"tax_benef_nom_{i}")
            with bc2:
                t = st.selectbox("Lien avec l'assuré", TYPES_OPTS, key=f"tax_benef_type_{i}")
            with bc3:
                p = st.number_input(
                    "Quote-part (%)",
                    min_value=0.0, max_value=100.0,
                    value=round(100.0 / nb_benef, 0),
                    step=5.0, key=f"tax_benef_part_{i}",
                )
            lien_succ = st.selectbox(
                "Lien de parenté (pour droits de succession 757B)",
                LIENS_OPTS,
                index=LIENS_OPTS.index("Enfant") if "Enfant" in t else 0,
                key=f"tax_benef_lien_{i}",
            )
            noms.append(nom)
            types_benef.append(t)
            parts_pct.append(float(p))
            liens_succ.append(lien_succ)

    total_parts = sum(parts_pct)
    if abs(total_parts - 100.0) > 0.5:
        st.error(f"⚠️ La somme des quotes-parts est {total_parts:.1f}% — elle doit être égale à 100%.")
        return

    st.markdown("---")

    if capital_deces <= 0:
        st.info("Saisissez la valeur du contrat au décès.")
        return

    # Calcul selon régime
    show_990i = "990I" in regime_versements or "Mix" in regime_versements
    show_757b = "757B" in regime_versements or "Mix" in regime_versements

    if show_990i:
        st.markdown("##### Art. 990I — Versements avant 70 ans")
        types_990i = [
            "Conjoint/PACS" if ("Conjoint" in t or "PACS" in t) else t
            for t in types_benef
        ]
        results_990i = calc_transmission_990I(capital_deces, nb_benef, parts_pct, types_990i)
        cap_990i = capital_deces
        if "Mix" in regime_versements:
            cap_990i_pct = 1.0 - (primes_apres_70 / capital_deces) if capital_deces > 0 else 1.0
            results_990i = calc_transmission_990I(
                capital_deces * cap_990i_pct, nb_benef, parts_pct, types_990i
            )
        _render_table_transmission(results_990i, noms, "990I")
        # Note démembrement
        for r in results_990i:
            if r.get("note"):
                st.info(r["note"])

    if show_757b:
        st.markdown("##### Art. 757B — Versements après 70 ans")
        st.info(
            "Les **gains** générés par les primes versées après 70 ans sont exonérés. "
            "Seules les **primes** (hors gains) au-delà de 30 500 € (abattement global) "
            "sont soumises aux droits de succession selon le lien de parenté."
        )
        st.caption(
            "Les droits de succession réels dépendent du patrimoine global, des donations "
            "antérieures et des abattements de droit commun (ex. 100 000€ par enfant). "
            "Cette simulation utilise des taux simplifiés."
        )
        results_757b = calc_transmission_757B(
            primes_apres_70, nb_benef, parts_pct, types_benef, capital_deces, liens_succ,
        )
        _render_table_transmission(results_757b, noms, "757B")


def _render_table_transmission(results: List[Dict[str, Any]], noms: List[str], regime: str) -> None:
    """Affiche le tableau de résultats par bénéficiaire."""
    rows = []
    for r in results:
        nom = noms[r["index"] - 1] if r["index"] - 1 < len(noms) and noms[r["index"] - 1] else f"Bénéficiaire {r['index']}"
        if regime == "990I":
            rows.append({
                "Bénéficiaire": nom,
                "Lien": r["type"],
                "Part (€)": _fmt_eur(r["part_eur"]),
                "Abattement": _fmt_eur(r.get("abattement", 0.0)),
                "Taxable": _fmt_eur(r["taxable"]),
                "Taxe due": _fmt_eur(r["taxe"]),
                "Net reçu": _fmt_eur(r["net_recu"]),
            })
        else:
            rows.append({
                "Bénéficiaire": nom,
                "Lien": r["type"],
                "Part (€)": _fmt_eur(r["part_eur"]),
                "Primes part": _fmt_eur(r.get("primes_part", 0.0)),
                "Abattement": _fmt_eur(r.get("abattement_part", 0.0)),
                "Taxable": _fmt_eur(r["taxable"]),
                "Taxe due": _fmt_eur(r["taxe"]),
                "Net reçu": _fmt_eur(r["net_recu"]),
            })
    total_taxe = sum(r["taxe"] for r in results)
    total_net = sum(r["net_recu"] for r in results)
    total_part = sum(r["part_eur"] for r in results)

    if rows:
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

    tc1, tc2, tc3 = st.columns(3)
    tc1.metric("Capital total transmis", _fmt_eur(total_part))
    tc2.metric("Total taxes", _fmt_eur(total_taxe))
    taux_eff = (total_taxe / total_part * 100) if total_part > 0 else 0.0
    tc3.metric("Taux effectif global", f"{taux_eff:.1f}%")


def _tab_exoneration() -> None:
    st.info(
        "Dans certaines situations, les gains issus d'un rachat sont totalement exonérés "
        "d'IR **ET** de prélèvements sociaux. Ces exonérations s'appliquent sur justificatif."
    )

    st.markdown("#### Cochez votre situation")
    cas_licenciement = st.checkbox(
        "Licenciement (inscription à Pôle Emploi / France Travail)"
    )
    cas_retraite_anticipee = st.checkbox(
        "Mise à la retraite anticipée par l'employeur"
    )
    cas_invalidite = st.checkbox(
        "Invalidité 2e ou 3e catégorie (assuré, conjoint ou enfant à charge)"
    )
    cas_liquidation = st.checkbox(
        "Liquidation judiciaire (non-salariés : cessation d'activité suite à jugement de liquidation)"
    )
    cas_cessation_ns = st.checkbox(
        "Cessation d'activité non salariée suite à jugement de liquidation"
    )

    cas_coches = [
        (cas_licenciement, "Licenciement",
         "Justificatif : lettre de licenciement + attestation Pôle Emploi / France Travail"),
        (cas_retraite_anticipee, "Retraite anticipée",
         "Justificatif : notification de mise à la retraite par l'employeur"),
        (cas_invalidite, "Invalidité",
         "Justificatif : notification CPAM de classement en invalidité 2e ou 3e catégorie"),
        (cas_liquidation, "Liquidation judiciaire",
         "Justificatif : jugement du tribunal de commerce"),
        (cas_cessation_ns, "Cessation NS",
         "Justificatif : jugement du tribunal de commerce"),
    ]

    au_moins_un = any(c[0] for c in cas_coches)

    if au_moins_un:
        st.success(
            "✅ **Exonération totale applicable** : ni IR ni prélèvements sociaux "
            "ne sont dus sur les gains du rachat."
        )
        # Tableau comparatif
        rachat_result: Optional[Dict[str, Any]] = st.session_state.get("tax_rachat_result")
        if rachat_result:
            ir_normal = rachat_result.get("montant_ir", 0.0)
            ps_normal = rachat_result.get("montant_ps", 0.0)
            total_normal = rachat_result.get("total_impots", 0.0)
            df_comp = pd.DataFrame([
                {"": "IR dû",
                 "Sans exonération": _fmt_eur(ir_normal),
                 "Avec exonération": _fmt_eur(0.0),
                 "Économie": _fmt_eur(ir_normal)},
                {"": "PS dûs",
                 "Sans exonération": _fmt_eur(ps_normal),
                 "Avec exonération": _fmt_eur(0.0),
                 "Économie": _fmt_eur(ps_normal)},
                {"": "TOTAL",
                 "Sans exonération": _fmt_eur(total_normal),
                 "Avec exonération": _fmt_eur(0.0),
                 "Économie": _fmt_eur(total_normal)},
            ])
            st.dataframe(df_comp, hide_index=True, use_container_width=True)
        else:
            st.caption("Renseignez d'abord l'onglet **Rachat** pour voir l'économie chiffrée.")

        st.markdown("#### Justificatifs requis")
        for coched, label, justif in cas_coches:
            if coched:
                st.caption(f"**{label}** — {justif}")

    st.markdown("---")
    st.warning(
        "⚠️ **Procédure :** l'exonération n'est pas appliquée automatiquement par l'assureur. "
        "Le bénéficiaire doit fournir les justificatifs à l'assureur au moment du rachat "
        "**ET** mentionner l'exonération dans sa déclaration de revenus (case spécifique)."
    )


def render_tax_module() -> None:
    """Module Fiscalité assurance-vie — entièrement autonome."""
    # Constantes fiscales (locales, non globales)
    PS_RATE            = 0.172
    PFU_RATE           = 0.128
    PFL_LOW_RATE       = 0.075
    PFL_HIGH_RATE      = 0.128
    ABATTEMENT_SEUL    = 4_600.0
    ABATTEMENT_COUPLE  = 9_200.0
    SEUIL_150K         = 150_000.0
    ABATTEMENT_990I    = 152_500.0
    TAUX_990I_TRANCHE1 = 0.20
    TAUX_990I_TRANCHE2 = 0.3125
    SEUIL_990I_T2      = 700_000.0
    ABATTEMENT_757B    = 30_500.0
    AGE_SEUIL          = 70

    # Initialisation session_state
    st.session_state.setdefault("tax_date_ouverture", date(2016, 1, 2))
    st.session_state.setdefault("tax_valeur_contrat", 100_000.0)
    st.session_state.setdefault("tax_versements_nets", 80_000.0)
    st.session_state.setdefault("tax_versements_nets_total", 80_000.0)
    st.session_state.setdefault("tax_situation_familiale", "Célibataire / veuf / divorcé")
    st.session_state.setdefault("tax_montant_brut", 10_000.0)
    st.session_state.setdefault("tax_montant_net", 9_000.0)
    st.session_state.setdefault("tax_last_edited", "brut")
    st.session_state.setdefault("tax_rachat_result", None)
    st.session_state.setdefault("tax_nb_beneficiaires", 2)
    st.session_state.setdefault("tax_regime_versements", "Principalement avant 70 ans (Art. 990I)")
    st.session_state.setdefault("tax_capital_deces", 200_000.0)
    st.session_state.setdefault("tax_abattement_deja_utilise", 0.0)
    st.session_state.setdefault("tax_nb_annees", 10)

    st.title("Fiscalité assurance-vie")
    st.warning(
        "⚠️ Simulation indicative à titre pédagogique. Ne constitue pas un conseil "
        "fiscal ou juridique. Les résultats dépendent de la situation personnelle "
        "de chaque client. Consultez un professionnel fiscaliste pour toute décision."
    )

    tab1, tab2, tab3, tab4 = st.tabs([
        "Rachat",
        "Optimisation abattement",
        "Transmission / Décès",
        "Cas d'exonération",
    ])
    with tab1:
        _tab_rachat()
    with tab2:
        _tab_optimisation_abattement()
    with tab3:
        _tab_transmission()
    with tab4:
        _tab_exoneration()


def run_perfect_portfolio():
    render_portfolio_builder()


# ============================================================
# MODULE FICHE FONDS — Morningstar via mstarpy
# ============================================================

_FS_SECTOR_LABELS: Dict[str, str] = {
    "basicMaterials":        "Matériaux de base",
    "consumerCyclical":      "Conso. cyclique",
    "financialServices":     "Services financiers",
    "realEstate":            "Immobilier",
    "communicationServices": "Communication",
    "energy":                "Énergie",
    "industrials":           "Industrie",
    "technology":            "Technologie",
    "consumerDefensive":     "Conso. défensive",
    "healthcare":            "Santé",
    "utilities":             "Services publics",
}

_FS_FI_LABELS: Dict[str, str] = {
    "government":         "Gouvernemental",
    "municipal":          "Municipal",
    "corporate":          "Corporate",
    "securitized":        "Titrisé",
    "cashAndEquivalents": "Cash & équivalents",
    "derivative":         "Dérivés",
}

_FS_TRAILING_PERIODS: Dict[str, str] = {
    "M1":  "1 mois",
    "M3":  "3 mois",
    "M6":  "6 mois",
    "M12": "1 an",
    "M36": "3 ans",
    "M60": "5 ans",
    "M120": "10 ans",
}


@st.cache_data(ttl=604800, show_spinner="Chargement des données Morningstar...")
def _load_fund_data(isin: str) -> Optional[Dict[str, Any]]:
    """Charge les données d'un fonds via mstarpy — TTL 7 jours."""
    if not MSTARPY_AVAILABLE:
        return None
    try:
        fund = mstarpy.Funds(term=isin, pageSize=1)
        if not fund.isin:
            return None

        result: Dict[str, Any] = {
            "name":            getattr(fund, "name", isin),
            "isin":            fund.isin,
            "holdings":        None,
            "sector_equity":   None,
            "sector_fi":       None,
            "trailing_returns": None,
            "risk":            None,
            "esg":             None,
        }

        # Holdings
        try:
            df = fund.holdings(holdingType="all")
            if df is not None:
                if not isinstance(df, pd.DataFrame):
                    df = pd.DataFrame(df)
                if not df.empty:
                    wanted = ["securityName", "weighting", "country", "sector",
                              "holdingType", "isin", "currency", "morningstarRating"]
                    present = [c for c in wanted if c in df.columns]
                    result["holdings"] = df[present].head(15).to_dict("records")
        except Exception:
            pass

        # Secteurs
        try:
            sectors = fund.sector()
            if isinstance(sectors, dict):
                eq_raw = (sectors.get("EQUITY") or {}).get("fundPortfolio") or {}
                fi_raw = (sectors.get("FIXEDINCOME") or {}).get("fundPortfolio") or {}
                result["sector_equity"] = {
                    k: float(v) for k, v in eq_raw.items()
                    if k != "portfolioDate" and isinstance(v, (int, float)) and float(v) > 0
                }
                result["sector_fi"] = {
                    k: float(v) for k, v in fi_raw.items()
                    if k != "portfolioDate" and isinstance(v, (int, float)) and float(v) > 0
                }
        except Exception:
            pass

        # Performances glissantes
        try:
            result["trailing_returns"] = fund.trailingReturn()
        except Exception:
            pass

        # Risque / volatilité
        try:
            result["risk"] = fund.riskVolatility()
        except Exception:
            pass

        # ESG
        try:
            result["esg"] = fund.esgRisk()
        except Exception:
            pass

        return result
    except Exception:
        return None


def _fs_extract_trailing(raw: Any) -> Optional[pd.DataFrame]:
    """Extrait les performances glissantes depuis la structure réelle de trailingReturn()."""
    if not isinstance(raw, dict):
        return None
    try:
        col_defs = raw.get("columnDefs") or []
        nav      = raw.get("totalReturnNAV") or []
        cat      = raw.get("totalReturnCategory") or []
        idx      = raw.get("totalReturnIndex") or []

        PERIODS = {
            "1Month":     "1 mois",
            "3Month":     "3 mois",
            "YearToDate": "Depuis le 1er janv.",
            "1Year":      "1 an",
            "3Year":      "3 ans",
            "5Year":      "5 ans",
            "10Year":     "10 ans",
        }

        def _val(lst: Any, i: int) -> Optional[float]:
            try:
                v = lst[i]
                return float(v) if v is not None else None
            except Exception:
                return None

        rows = []
        for period_key, period_label in PERIODS.items():
            if period_key not in col_defs:
                continue
            i  = col_defs.index(period_key)
            fv = _val(nav, i)
            cv = _val(cat, i)
            iv = _val(idx, i)
            if fv is not None:
                rows.append({
                    "Période":       period_label,
                    "Fonds (%)":     fv,
                    "Catégorie (%)": cv,
                    "Indice (%)":    iv,
                })
        return pd.DataFrame(rows) if rows else None
    except Exception:
        return None


def _fs_extract_risk(raw: Any) -> Dict[str, Optional[float]]:
    """Extrait vol 3 ans, max drawdown, sharpe depuis la structure réelle de riskVolatility()."""
    out: Dict[str, Optional[float]] = {"vol": None, "mdd": None, "sharpe": None}
    if not isinstance(raw, dict):
        return out
    try:
        frv = raw.get("fundRiskVolatility") or {}
        for period in ("for3Year", "for5Year", "for1Year"):
            period_data = frv.get(period) or {}
            if not isinstance(period_data, dict):
                continue
            vol    = period_data.get("standardDeviation")
            sharpe = period_data.get("sharpeRatio")
            if vol is not None or sharpe is not None:
                try:
                    out["vol"]    = float(vol)    if vol    is not None else None
                    out["sharpe"] = float(sharpe) if sharpe is not None else None
                except (TypeError, ValueError):
                    pass
                break
    except Exception:
        pass
    return out


def _fs_extract_esg(raw: Any) -> Dict[str, Any]:
    """Extrait score, catégorie calculée, date depuis la structure réelle de esgRisk()."""
    out: Dict[str, Any] = {"score": None, "category": None, "date": None}
    if not isinstance(raw, dict):
        return out
    try:
        score = raw.get("fundSustainabilityScore")
        if score is not None:
            try:
                score = float(score)
                out["score"] = score
                if score <= 10:
                    out["category"] = "Négligeable"
                elif score <= 20:
                    out["category"] = "Faible"
                elif score <= 30:
                    out["category"] = "Moyen"
                elif score <= 40:
                    out["category"] = "Élevé"
                else:
                    out["category"] = "Sévère"
            except (TypeError, ValueError):
                pass
        date_raw = (
            raw.get("portfolioDateSustainabilityRating")
            or raw.get("portfolioDate")
            or raw.get("categoryRankDate")
        )
        if date_raw:
            try:
                out["date"] = pd.Timestamp(date_raw).strftime("%d/%m/%Y")
            except Exception:
                out["date"] = str(date_raw)[:10]
    except Exception:
        pass
    return out


def _fs_bar_chart(labels: List[str], values: List[float], color: str, title: str) -> None:
    """Graphique barres horizontal — Plotly si disponible, Altair sinon."""
    # Filtrer valeurs négligeables
    pairs = [(l, v) for l, v in zip(labels, values) if v >= 0.5]
    if not pairs:
        return
    pairs_sorted = sorted(pairs, key=lambda x: x[1], reverse=True)
    labs_s = [p[0] for p in pairs_sorted]
    vals_s = [p[1] for p in pairs_sorted]
    max_val = max(vals_s) if vals_s else 1.0
    # Espace pour les labels : 50% si valeurs < 15%, sinon 30%
    x_max = max_val * (1.5 if max_val < 15.0 else 1.3)
    text_pos = "inside" if max_val >= 15.0 else "outside"

    if PLOTLY_AVAILABLE and go is not None:
        fig = go.Figure(go.Bar(
            x=vals_s,
            y=labs_s,
            orientation="h",
            marker_color=color,
            marker_line_width=0,
            text=[f"{v:.1f}%" for v in vals_s],
            textposition=text_pos,
            textfont=dict(size=11),
            hovertemplate="%{y}: %{x:.1f}%<extra></extra>",
        ))
        fig.update_layout(
            height=max(220, 32 * len(labs_s) + 60),
            margin=dict(l=10, r=60, t=10, b=30),
            xaxis=dict(
                range=[0, x_max],
                ticksuffix="%",
                showgrid=True,
                gridcolor="rgba(200,200,200,0.3)",
                zeroline=False,
                title=None,
            ),
            yaxis=dict(
                autorange="reversed",
                tickfont=dict(size=11),
            ),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        # Fallback Altair inchangé
        df_chart = pd.DataFrame({"Catégorie": labs_s, "Valeur (%)": vals_s})
        chart = (
            alt.Chart(df_chart)
            .mark_bar(color=color)
            .encode(
                x=alt.X(
                    "Valeur (%):Q",
                    title=None,
                    scale=alt.Scale(domain=[0, x_max]),
                ),
                y=alt.Y("Catégorie:N", sort="-x"),
                tooltip=[
                    "Catégorie:N",
                    alt.Tooltip("Valeur (%):Q", format=".1f"),
                ],
            )
            .properties(height=max(200, 28 * len(labs_s) + 50))
        )
        st.altair_chart(chart, use_container_width=True)


def _render_fund_sheet_content(data: Dict[str, Any]) -> None:
    """Affiche le contenu de la fiche fonds à partir du dict chargé."""
    st.header(data.get("name", "—"))
    st.caption(f"ISIN : {data.get('isin', '—')}")
    st.markdown("---")

    # ── Section 1 : Holdings ─────────────────────────────────
    try:
        holdings = data.get("holdings")
        if holdings:
            st.subheader("🏦 Top 15 positions")
            df_h = pd.DataFrame(holdings)
            rename_map = {
                "securityName":    "Titre",
                "weighting":       "Poids (%)",
                "country":         "Pays",
                "sector":          "Secteur",
                "holdingType":     "Type",
                "currency":        "Devise",
                "morningstarRating": "Note MS",
            }
            df_h = df_h.rename(columns={k: v for k, v in rename_map.items() if k in df_h.columns})
            if "Poids (%)" in df_h.columns:
                df_h["Poids (%)"] = pd.to_numeric(df_h["Poids (%)"], errors="coerce").round(2)

            type_col = "Type"
            df_eq = df_h[df_h[type_col].str.lower().str.contains("equity", na=False)] if type_col in df_h.columns else pd.DataFrame()
            df_bond = df_h[df_h[type_col].str.lower().str.contains("bond|fi|fixed", na=False)] if type_col in df_h.columns else pd.DataFrame()
            df_other = df_h[~df_h.index.isin(df_eq.index.tolist() + df_bond.index.tolist())] if not df_eq.empty or not df_bond.empty else df_h

            with st.expander("📈 Actions", expanded=True):
                if not df_eq.empty:
                    st.dataframe(df_eq.drop(columns=[type_col], errors="ignore"), use_container_width=True, hide_index=True)
                else:
                    st.info("Aucune position actions dans le top 15.")
            with st.expander("📉 Obligations", expanded=True):
                if not df_bond.empty:
                    st.dataframe(df_bond.drop(columns=[type_col], errors="ignore"), use_container_width=True, hide_index=True)
                else:
                    st.info("Aucune position obligataire dans le top 15.")
            if not df_other.empty and type_col in df_h.columns:
                with st.expander("🔹 Autres", expanded=False):
                    st.dataframe(df_other, use_container_width=True, hide_index=True)
        else:
            st.subheader("🏦 Top 15 positions")
            st.info("Holdings non disponibles pour ce fonds.")
    except Exception:
        st.subheader("🏦 Top 15 positions")
        st.info("Holdings non disponibles pour ce fonds.")

    st.markdown("---")

    # ── Section 2 : Secteurs ─────────────────────────────────
    col_eq, col_fi = st.columns(2)
    with col_eq:
        try:
            sec_eq = data.get("sector_equity") or {}
            if sec_eq:
                st.subheader("📊 Secteurs (Actions)")
                labels = [_FS_SECTOR_LABELS.get(k, k) for k in sec_eq]
                values = list(sec_eq.values())
                _fs_bar_chart(labels, values, "#1f77b4", "Secteurs actions")
            else:
                st.subheader("📊 Secteurs (Actions)")
                st.info("Données sectorielles actions non disponibles.")
        except Exception:
            st.info("Données sectorielles actions non disponibles.")

    with col_fi:
        try:
            sec_fi = data.get("sector_fi") or {}
            if sec_fi:
                st.subheader("📊 Obligations")
                labels = [_FS_FI_LABELS.get(k, k) for k in sec_fi]
                values = list(sec_fi.values())
                _fs_bar_chart(labels, values, "#ff7f0e", "Répartition obligataire")
            else:
                st.subheader("📊 Obligations")
                st.info("Données de répartition obligataire non disponibles.")
        except Exception:
            st.info("Données de répartition obligataire non disponibles.")

    st.markdown("---")

    # ── Section 3 : Performances glissantes ──────────────────
    try:
        st.subheader("📈 Performances glissantes")
        df_perf = _fs_extract_trailing(data.get("trailing_returns"))
        if df_perf is not None and not df_perf.empty:
            num_cols = [c for c in df_perf.columns if c != "Période"]
            col_cfg: Dict[str, Any] = {}
            for c in num_cols:
                col_cfg[c] = st.column_config.NumberColumn(c, format="%.2f%%")
            st.dataframe(
                df_perf,
                use_container_width=True,
                hide_index=True,
                column_config=col_cfg,
            )
        else:
            st.info("Performances non disponibles.")
    except Exception:
        st.info("Performances non disponibles.")

    st.markdown("---")

    # ── Section 4 : Risque ───────────────────────────────────
    try:
        st.subheader("⚠️ Indicateurs de risque")
        risk = _fs_extract_risk(data.get("risk"))
        if any(v is not None for v in risk.values()):
            rc1, rc2, rc3 = st.columns(3)
            rc1.metric(
                "Volatilité 3 ans (ann.)",
                f"{risk['vol']:.1f}%" if risk["vol"] is not None else "—",
            )
            rc2.metric(
                "Max Drawdown",
                f"{risk['mdd']:.1f}%" if risk["mdd"] is not None else "—",
            )
            rc3.metric(
                "Sharpe 3 ans",
                f"{risk['sharpe']:.2f}" if risk["sharpe"] is not None else "—",
            )
        else:
            st.info("Données de risque non disponibles.")
    except Exception:
        st.info("Données de risque non disponibles.")

    st.markdown("---")

    # ── Section 5 : ESG ──────────────────────────────────────
    try:
        st.subheader("🌱 Score ESG (Morningstar Sustainalytics)")
        esg = _fs_extract_esg(data.get("esg"))
        if esg["score"] is not None or esg["category"] is not None:
            ec1, ec2 = st.columns([1, 2])
            with ec1:
                st.metric("Score ESG", f"{esg['score']:.1f}" if esg["score"] is not None else "—")
            with ec2:
                cat_str  = esg["category"] or "—"
                date_str = esg["date"] or "—"
                st.caption(f"Catégorie : **{cat_str}** — Mis à jour : {date_str}")
        else:
            st.info("Données ESG non disponibles.")
    except Exception:
        st.info("Données ESG non disponibles.")


def render_fund_sheet() -> None:
    """Mode Fiche fonds — chargement Morningstar via mstarpy."""
    if not MSTARPY_AVAILABLE:
        st.error(
            "La librairie mstarpy n'est pas installée. "
            "Ajoutez 'mstarpy' dans requirements.txt."
        )
        return

    st.title("📊 Fiche fonds")
    st.info(
        "ℹ️ Données issues de Morningstar (données publiques). "
        "Holdings mis à jour mensuellement. À titre indicatif."
    )

    # Saisie ISIN
    isin_raw = st.text_input(
        "ISIN du fonds",
        placeholder="Ex : FR0010135103",
        max_chars=12,
        key="fs_isin",
    ).strip().upper()

    # Validation format : 12 chars, 2 premières lettres, reste alphanumérique
    valid_isin = (
        len(isin_raw) == 12
        and isin_raw[:2].isalpha()
        and isin_raw[2:].isalnum()
    )

    # Détecter changement d'ISIN → réinitialiser le trigger
    last_isin = st.session_state.get("fs_last_isin", "")
    if isin_raw != last_isin:
        st.session_state["fs_triggered"] = False

    if isin_raw and not valid_isin:
        st.warning("Format ISIN invalide. Un ISIN fait 12 caractères : 2 lettres (pays) + 10 alphanumériques. Ex : FR0010135103")
        return

    col_btn, _ = st.columns([1, 4])
    with col_btn:
        if st.button("🔍 Rechercher", key="fs_search_btn", type="primary"):
            st.session_state["fs_triggered"] = True
            st.session_state["fs_last_isin"] = isin_raw

    if not st.session_state.get("fs_triggered"):
        return

    if not isin_raw:
        st.warning("Saisissez un ISIN avant de lancer la recherche.")
        return

    data = _load_fund_data(isin_raw)
    if data is None:
        st.error(
            f"Fonds non trouvé pour l'ISIN **{isin_raw}**. "
            "Vérifiez l'ISIN ou essayez une autre part du même fonds."
        )
        return

    _render_fund_sheet_content(data)


def render_mode_router():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    mode = st.radio(
        "Mode",
        [
            "Comparer des portefeuilles",
            "Construction de portefeuille optimisé",
            "Fiscalité assurance-vie",
            "📊 Fiche fonds",
        ],
        horizontal=True,
    )
    if mode == "Comparer des portefeuilles":
        run_comparator()
    elif mode == "Construction de portefeuille optimisé":
        run_perfect_portfolio()
    elif mode == "Fiscalité assurance-vie":
        render_tax_module()
    else:
        render_fund_sheet()


def _render_with_crash_shield():
    try:
        render_mode_router()
        st.session_state["APP_STATUS"] = "OK"
    except Exception as e:
        st.session_state["APP_STATUS"] = "KO"
        st.session_state["LAST_EXCEPTION"] = str(e)
        st.title(APP_TITLE)
        st.info("App chargée, statut KO")
        st.error("Une erreur est survenue pendant le rendu.")
        st.exception(e)
        st.markdown("""
Conseils :
- Vérifiez vos dépendances (reportlab/matplotlib).
- Vérifiez la clé EODHD dans les secrets.
- Réessayez après avoir vidé le cache Streamlit.
""")


_render_with_crash_shield()

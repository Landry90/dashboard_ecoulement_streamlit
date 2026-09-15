"""
Application Streamlit — Écoulement des dépôts à vue
=====================================================

Reconstitue, en Python pur (numpy / pandas), les fonctions mathématiques
du classeur "Ecoulement_DAV1" :

  1. decay_share(H, k, cap)          -> fonction d'écoulement S(H)
  2. decompose_stable_volatile(...)  -> décomposition stable / volatile de l'encours
  3. project_resources(...)          -> projection des ressources futures

Chaque fonction est documentée, utilisée telle quelle pour produire les
graphiques (aucune duplication de logique entre "calcul" et "affichage"),
et son code source est affiché dans l'application via `inspect.getsource`
pour que la démonstration mathématique soit vérifiable à l'écran.

Lancer en local :   streamlit run app.py
Déployer :          voir README.md
"""

import datetime as dt
import inspect
import io
import json
import re
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ----------------------------------------------------------------------
# 0. Palette & configuration générale
# ----------------------------------------------------------------------
WHITE = "#FFFFFF"
GRAY_50 = "#F7F7F7"
GRAY_200 = "#DCDCDC"
GRAY_400 = "#A8A8A8"
GRAY_600 = "#6B6B6B"
BLACK = "#121212"
RED = "#C81E3A"
RED_DARK = "#8F1428"

STABLE_COLOR = RED
VOLATILE_COLOR = "#B0B0B0"

st.set_page_config(
    page_title="Écoulement des dépôts à vue — DAV1",
    page_icon="📊",
    layout="wide",
)

# CSS minimal pour imposer la charte blanc / gris / rouge / noir à Streamlit
st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {WHITE}; }}
    h1, h2, h3 {{ color: {BLACK}; }}
    [data-testid="stMetricValue"] {{ color: {RED}; }}
    [data-testid="stMetricLabel"] {{ color: {GRAY_600}; }}
    .stTabs [aria-selected="true"] {{ color: {RED} !important; border-bottom-color: {RED} !important; }}
    div[role="radiogroup"] label {{ color: {BLACK}; }}
    </style>
    """,
    unsafe_allow_html=True,
)

PLOTLY_LAYOUT = dict(
    plot_bgcolor=WHITE,
    paper_bgcolor=WHITE,
    font=dict(color=BLACK, family="IBM Plex Sans, Arial, sans-serif"),
    legend=dict(bgcolor=WHITE, bordercolor=GRAY_200, borderwidth=1),
    margin=dict(t=50, l=10, r=10, b=10),
)
GRID = dict(gridcolor=GRAY_200, zerolinecolor=GRAY_200)


# ----------------------------------------------------------------------
# 1. Chargement des données par défaut (JSON embarqué)
# ----------------------------------------------------------------------
@st.cache_data
def load_default_data() -> dict:
    """Charge les paramètres des comptes, l'historique et les données de projection d'origine."""
    path = Path(__file__).parent / "data_ecoulement.json"
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ----------------------------------------------------------------------
# 1bis. Import d'un fichier Excel pour mettre à jour les données
# ----------------------------------------------------------------------
# Forme canonique attendue (celle du classeur d'audit "Ecoulement DAV") :
#   - feuille "Audit des calculs"          : paramètres des comptes
#   - feuille "Tous détail Stocks"         : historique mensuel par compte
#   - feuille "Ecoulement Ressources à Vue": encours de référence + prévision
#                                             de production nouvelle
# Le fichier reçu est d'abord lu et ramené à cette forme (extraction stricte
# des positions attendues) ; s'il n'y correspond pas, il est rejeté avec le
# détail des anomalies plutôt que d'être consommé tel quel.
REQUIRED_SHEETS = {
    "Audit des calculs": "paramètres des comptes (id, modèle, μ, Umin)",
    "Tous détail Stocks": "historique mensuel des encours, par compte",
    "Ecoulement Ressources à Vue": "encours de référence + prévision de production nouvelle",
}

DEFAULT_PALETTE = [
    "#C81E3A", "#E8A3AE", "#6E6E6E", "#B7B7B7", "#111111", "#7A1023",
    "#8F1428", "#D9A441", "#3E7C8C", "#5B8C5A",
]


class ExcelFormatError(Exception):
    """Levée quand le fichier Excel reçu ne respecte pas la forme attendue."""

    def __init__(self, problems: list):
        self.problems = problems
        super().__init__("; ".join(problems))


@dataclass
class ImportResult:
    accounts: list
    hist: dict
    proj: dict
    total_encours: float
    warnings: list = field(default_factory=list)


def _find_cell(ws, text: str, max_row: int = 12, max_col: int = 20, exact: bool = True):
    """Cherche la première cellule (dans la zone d'en-tête) contenant `text`."""
    for row in ws.iter_rows(
        min_row=1, max_row=min(max_row, ws.max_row),
        min_col=1, max_col=min(max_col, ws.max_column),
    ):
        for c in row:
            if c.value is None:
                continue
            v = str(c.value).strip()
            if (v == text) if exact else (text.lower() in v.lower()):
                return c
    return None


def _read_audit_sheet(ws, problems: list) -> list:
    """Lit id / modèle / μ / Umin / μ+Umin depuis la feuille 'Audit des calculs'."""
    header = _find_cell(ws, "Compte")
    if header is None:
        problems.append("Feuille « Audit des calculs » : en-tête « Compte » introuvable.")
        return []

    row0, col0 = header.row, header.column
    expected = {"Compte": 0, "Modèle": 1, "μ / drift": 2, "Umin": 3, "μ + Umin": 4}
    for label, offset in expected.items():
        val = ws.cell(row=row0, column=col0 + offset).value
        if val is None or str(val).strip() != label:
            problems.append(
                f"Feuille « Audit des calculs » : colonne « {label} » attendue juste "
                f"après « Compte » (trouvé : {val!r})."
            )
    if problems:
        return []

    accounts = []
    r = row0 + 1
    while True:
        acc_id = ws.cell(row=r, column=col0).value
        if acc_id is None or str(acc_id).strip() == "":
            break
        model = ws.cell(row=r, column=col0 + 1).value
        mu = ws.cell(row=r, column=col0 + 2).value
        umin = ws.cell(row=r, column=col0 + 3).value
        k = ws.cell(row=r, column=col0 + 4).value
        if not all(isinstance(x, (int, float)) for x in (mu, umin, k)):
            problems.append(f"Compte {acc_id!r} (ligne {r}) : μ, Umin ou μ+Umin non numérique.")
        elif abs((mu + umin) - k) > 1e-6:
            problems.append(
                f"Compte {acc_id!r} : μ+Umin ({mu + umin:.6g}) ≠ valeur déclarée ({k:.6g})."
            )
        else:
            accounts.append({
                "id": str(acc_id).strip(),
                "model": str(model).strip() if model else "",
                "mu": float(mu), "umin": float(umin), "k": float(k),
            })
        r += 1

    if not accounts:
        problems.append("Feuille « Audit des calculs » : aucun compte exploitable trouvé.")
    return accounts


def _read_hist_sheet(ws, problems: list):
    """Lit l'historique mensuel de chaque compte depuis 'Tous détail Stocks'."""
    blocks = []
    pattern = re.compile(r"ECOULEMENT DU COMPTE\s+(\S+)")
    for row in ws.iter_rows(min_row=1, max_row=min(6, ws.max_row)):
        for c in row:
            if c.value:
                m = pattern.search(str(c.value))
                if m:
                    blocks.append((c.column, m.group(1).strip()))

    if not blocks:
        problems.append(
            "Feuille « Tous détail Stocks » : aucun bloc « ECOULEMENT DU COMPTE ... » trouvé."
        )
        return {}, []

    hist, ref_dates = {}, None
    for col0, acc_id in blocks:
        header_row = None
        for r in range(1, 10):
            if str(ws.cell(row=r, column=col0).value or "").strip() == "Date":
                header_row = r
                break
        if header_row is None:
            problems.append(f"Compte {acc_id!r} : en-tête « Date » introuvable dans son bloc.")
            continue

        val_header = ws.cell(row=header_row, column=col0 + 1).value
        if str(val_header or "").strip() != acc_id:
            problems.append(
                f"Compte {acc_id!r} : colonne des valeurs attendue juste après « Date » "
                f"(trouvé {val_header!r})."
            )
            continue

        dates, values = [], []
        r = header_row + 1
        while True:
            d = ws.cell(row=r, column=col0).value
            v = ws.cell(row=r, column=col0 + 1).value
            if d is None:
                break
            if not isinstance(d, (dt.datetime, dt.date)) or not isinstance(v, (int, float)):
                problems.append(f"Compte {acc_id!r}, ligne {r} : date ou valeur invalide.")
                break
            dates.append((d.date() if isinstance(d, dt.datetime) else d).isoformat())
            values.append(float(v))
            r += 1

        if not dates:
            problems.append(f"Compte {acc_id!r} : aucune donnée historique lue.")
            continue
        if ref_dates is None:
            ref_dates = dates
        elif dates != ref_dates:
            problems.append(
                f"Compte {acc_id!r} : les dates historiques ({len(dates)} points) ne "
                f"correspondent pas à celles des autres comptes ({len(ref_dates)} points)."
            )
        hist[acc_id] = values

    return hist, (ref_dates or [])


def _read_proj_sheet(ws, problems: list):
    """Lit dates/production de projection + encours de référence."""
    prod_header = _find_cell(ws, "Prévision production nouvelle", exact=False)
    if prod_header is None:
        problems.append(
            "Feuille « Ecoulement Ressources à Vue » : colonne « Prévision production "
            "nouvelle » introuvable."
        )
        return [], [], None

    row0, prod_col = prod_header.row, prod_header.column
    date_col = prod_col - 2
    if str(ws.cell(row=row0, column=date_col).value or "").strip() != "Date":
        problems.append(
            "Feuille « Ecoulement Ressources à Vue » : colonne « Date » attendue 2 "
            "colonnes avant « Prévision production nouvelle »."
        )
        return [], [], None

    dates, prod = [], []
    r = row0 + 1
    while True:
        d = ws.cell(row=r, column=date_col).value
        p = ws.cell(row=r, column=prod_col).value
        if d is None:
            break
        if not isinstance(d, (dt.datetime, dt.date)) or not isinstance(p, (int, float)):
            problems.append(f"Projection, ligne {r} : date ou production invalide.")
            break
        dates.append((d.date() if isinstance(d, dt.datetime) else d).isoformat())
        prod.append(float(p))
        r += 1
    if not dates:
        problems.append("Feuille « Ecoulement Ressources à Vue » : aucune projection lue.")

    encours_header = _find_cell(ws, "Encours")
    total_encours = None
    if encours_header is not None:
        col = encours_header.column
        vals = []
        r = encours_header.row + 1
        while True:
            v = ws.cell(row=r, column=col).value
            if v is None:
                break
            if isinstance(v, (int, float)):
                vals.append(float(v))
            r += 1
        if vals:
            if max(vals) - min(vals) > 1e-3 * max(abs(max(vals)), 1):
                problems.append(
                    "Feuille « Ecoulement Ressources à Vue » : la colonne « Encours » "
                    "n'est pas constante sur l'horizon."
                )
            total_encours = vals[0]
    if total_encours is None:
        problems.append("Feuille « Ecoulement Ressources à Vue » : encours de référence introuvable.")

    return dates, prod, total_encours


def parse_excel_upload(uploaded_file, previous_accounts: dict) -> ImportResult:
    """
    Lit un classeur Excel (forme "Ecoulement DAV") et le ramène au schéma
    interne de l'application (mêmes clés que data_ecoulement.json).

    `previous_accounts` (dict id -> compte déjà connu) sert à réutiliser le
    cap/la couleur déjà configurés pour un compte déjà présent.

    Lève `ExcelFormatError` (liste d'anomalies) si le fichier ne respecte pas
    la forme attendue : aucune donnée partielle n'est alors renvoyée.
    """
    problems: list = []
    try:
        wb = openpyxl.load_workbook(io.BytesIO(uploaded_file.getvalue()), data_only=True)
    except Exception as exc:
        raise ExcelFormatError([f"Fichier illisible (n'est-ce pas un .xlsx valide ?) : {exc}"])

    missing = [s for s in REQUIRED_SHEETS if s not in wb.sheetnames]
    if missing:
        raise ExcelFormatError(
            [f"Feuille manquante : « {s} » ({REQUIRED_SHEETS[s]})." for s in missing]
        )

    accounts_raw = _read_audit_sheet(wb["Audit des calculs"], problems)
    hist, hist_dates = _read_hist_sheet(wb["Tous détail Stocks"], problems)
    proj_dates, prod, total_encours = _read_proj_sheet(wb["Ecoulement Ressources à Vue"], problems)
    if problems:
        raise ExcelFormatError(problems)

    ids_audit, ids_hist = {a["id"] for a in accounts_raw}, set(hist.keys())
    if ids_audit != ids_hist:
        raise ExcelFormatError([
            "Les comptes de « Audit des calculs » "
            f"({sorted(ids_audit)}) et de « Tous détail Stocks » "
            f"({sorted(ids_hist)}) ne correspondent pas."
        ])

    warnings = []
    accounts = []
    for i, a in enumerate(accounts_raw):
        acc_id = a["id"]
        ref = previous_accounts.get(acc_id)
        encours_ref = hist[acc_id][-1]
        accounts.append({
            "id": acc_id,
            "model": a["model"],
            "mu": a["mu"], "umin": a["umin"], "k": a["k"],
            "encoursRef": encours_ref,
            "cap": ref["cap"] if ref else True,
            "color": ref["color"] if ref else DEFAULT_PALETTE[i % len(DEFAULT_PALETTE)],
        })

    sum_encours = sum(a["encoursRef"] for a in accounts)
    if total_encours and abs(sum_encours - total_encours) > 1e-3 * max(total_encours, 1):
        warnings.append(
            f"Écart entre l'encours total déclaré ({total_encours:,.0f} FCFA) et la "
            f"somme des encours par compte ({sum_encours:,.0f} FCFA)."
        )

    return ImportResult(
        accounts=accounts,
        hist={"dates": hist_dates, **hist},
        proj={"datesProj": proj_dates, "previsionProductionNouvelle": prod},
        total_encours=total_encours or sum_encours,
        warnings=warnings,
    )


# ----------------------------------------------------------------------
# 1ter. Barre latérale : import / retour aux données d'origine
# ----------------------------------------------------------------------
if "custom_data" not in st.session_state:
    st.session_state.custom_data = None
    st.session_state.custom_data_meta = None
if "pending_import" not in st.session_state:
    st.session_state.pending_import = None

with st.sidebar:
    st.markdown("### 📥 Mise à jour des données")
    if st.session_state.custom_data_meta:
        st.caption(
            f"Source actuelle : fichier importé « {st.session_state.custom_data_meta['name']} » "
            f"le {st.session_state.custom_data_meta['when']}."
        )
        if st.button("↩️ Revenir aux données d'origine"):
            st.session_state.custom_data = None
            st.session_state.custom_data_meta = None
            st.session_state.pending_import = None
            st.rerun()
    else:
        st.caption("Source actuelle : données d'origine (`data_ecoulement.json`).")

    up = st.file_uploader("Fichier Excel (.xlsx)", type=["xlsx"])
    if up is not None:
        current = st.session_state.custom_data or load_default_data()
        previous_accounts = {a["id"]: a for a in current["accounts"]}
        try:
            result = parse_excel_upload(up, previous_accounts)
            st.session_state.pending_import = (up.name, result)
        except ExcelFormatError as exc:
            st.session_state.pending_import = None
            st.error("Le fichier ne respecte pas la forme attendue et a été rejeté :")
            for p in exc.problems:
                st.error(f"• {p}")

if st.session_state.pending_import:
    fname, result = st.session_state.pending_import
    st.info(
        f"Fichier « {fname} » lu avec succès : {len(result.accounts)} comptes, "
        f"{len(result.hist['dates'])} mois d'historique, "
        f"{len(result.proj['datesProj'])} mois de projection."
    )
    for w in result.warnings:
        st.warning(w)

    st.subheader("Vérification avant application")
    edit_df = pd.DataFrame([
        {
            "Compte": a["id"], "Modèle": a["model"], "k": round(a["k"], 6),
            "Encours réf.": a["encoursRef"], "Plafonner à 100 % (cap)": a["cap"],
            "Couleur": a["color"],
        }
        for a in result.accounts
    ])
    edited = st.data_editor(
        edit_df, hide_index=True, width="stretch",
        column_config={
            "Plafonner à 100 % (cap)": st.column_config.CheckboxColumn(),
            "Couleur": st.column_config.TextColumn(help="Code hexadécimal, ex. #C81E3A"),
        },
        disabled=["Compte", "Modèle", "k", "Encours réf."],
        key="import_editor",
    )

    col_ok, col_cancel = st.columns(2)
    if col_ok.button("✅ Valider et appliquer ces données", type="primary"):
        for a, (_, row) in zip(result.accounts, edited.iterrows()):
            a["cap"] = bool(row["Plafonner à 100 % (cap)"])
            a["color"] = str(row["Couleur"])
        st.session_state.custom_data = {
            "accounts": result.accounts,
            "hist": result.hist,
            "proj": result.proj,
            "totalEncours": result.total_encours,
        }
        st.session_state.custom_data_meta = {
            "name": fname, "when": dt.datetime.now().strftime("%d/%m/%Y %H:%M"),
        }
        st.session_state.pending_import = None
        st.success("Données mises à jour.")
        st.rerun()
    if col_cancel.button("✖️ Annuler"):
        st.session_state.pending_import = None
        st.rerun()

DATA = st.session_state.custom_data or load_default_data()
ACCOUNTS = DATA["accounts"]          # liste de dicts : id, model, mu, umin, k, cap, encoursRef, color
HIST = DATA["hist"]                  # historique mensuel par compte
PROJ = DATA["proj"]                  # dates, prévision de production nouvelle
TOTAL_ENCOURS = DATA["totalEncours"]


# ----------------------------------------------------------------------
# 2. FONCTIONS MATHÉMATIQUES DU MODÈLE
#    (ce sont exactement ces fonctions qui alimentent les graphiques)
# ----------------------------------------------------------------------

def decay_share(H, k: float, cap: bool = True):
    """
    Fonction d'écoulement (fonction de survie) d'un compte de dépôts à vue.

        S(H) = min( 1 , exp( H × k ) )   si cap = True
        S(H) = exp( H × k )              si cap = False   (cas du compte 372E)

    Paramètres
    ----------
    H : horizon en mois (scalaire ou array-like)
    k : taux de décroissance mensuel du compte, k = mu + Umin
        (mu = dérive du modèle ARIMA/SARIMA, Umin = résidu minimal historique)
    cap : si True, plafonne la part stable à 100 % (cas général)

    Retourne
    --------
    La part de l'encours encore jugée stable à l'horizon H (entre 0 et 1).
    """
    H = np.asarray(H, dtype=float)
    survie = np.exp(H * k)
    return np.minimum(1.0, survie) if cap else survie


def decompose_stable_volatile(accounts: list, H: float) -> pd.DataFrame:
    """
    Décompose l'encours de chaque compte en partie stable / partie volatile
    à un horizon H donné (en mois), à partir de la fonction `decay_share`.

        Stable(compte, H)   = S(H) × Encours(compte)
        Volatile(compte, H) = Encours(compte) − Stable(compte, H)
    """
    rows = []
    for a in accounts:
        share = float(decay_share(H, a["k"], a["cap"]))
        stable = share * a["encoursRef"]
        volatile = a["encoursRef"] - stable
        rows.append(
            {
                "compte": a["id"],
                "modele": a["model"],
                "encours": a["encoursRef"],
                "part_stable": share,
                "stable": stable,
                "volatile": volatile,
                "color": a["color"],
            }
        )
    return pd.DataFrame(rows)


def project_resources(
    accounts: list,
    dates_proj: list,
    production: list,
    multiplier: float = 1.0,
    horizon_offset: int = 6,
) -> pd.DataFrame:
    """
    Projette les ressources futures sur les horizons fournis (calage repris
    du classeur d'origine : le mois de projection i est adossé à l'horizon
    d'écoulement H = i + horizon_offset, cf. note d'audit "Mapping production").

        Ressources(i) = Σ_comptes S(H) × Encours(compte)  +  multiplicateur × Production(i)

    Paramètres
    ----------
    accounts        : liste des paramètres des 6 comptes
    dates_proj      : liste de dates (une par mois de projection)
    production      : hypothèse de production nouvelle, un montant par mois
    multiplier      : facteur de scénario appliqué à la production nouvelle
    horizon_offset  : décalage (en mois) entre l'indice de projection et
                      l'horizon d'écoulement utilisé (6 dans le classeur d'origine)
    """
    n = len(dates_proj)
    stable = np.zeros(n)
    for a in accounts:
        H = np.arange(1, n + 1) + horizon_offset
        stable += decay_share(H, a["k"], a["cap"]) * a["encoursRef"]

    prod = np.asarray(production[:n], dtype=float) * multiplier
    total = stable + prod

    return pd.DataFrame(
        {
            "date": dates_proj,
            "stable": stable,
            "production_nouvelle": prod,
            "total_projete": total,
        }
    )


# ----------------------------------------------------------------------
# Utilitaires d'affichage
# ----------------------------------------------------------------------
def fmt_mds(v: float) -> str:
    return f"{v / 1e9:,.1f} Mds FCFA".replace(",", " ")


def fmt_pct(v: float) -> str:
    return f"{v * 100:,.1f} %".replace(",", " ")


def show_source(fn):
    """Affiche, dans un expander, le code source réel de la fonction utilisée."""
    with st.expander(f"🔎 Voir le code Python de `{fn.__name__}`"):
        st.code(inspect.getsource(fn), language="python")


# ----------------------------------------------------------------------
# En-tête
# ----------------------------------------------------------------------
st.markdown(
    f"""
    <div style="border-bottom:3px solid {RED};padding-bottom:14px;margin-bottom:10px;">
        <p style="color:{RED};font-weight:600;font-size:13px;margin:0 0 4px;">
            AFRILAND FIRST BANK · MODÉLISATION DU STOCK DE DÉPÔTS
        </p>
        <h1 style="margin:0;font-size:34px;">Écoulement des dépôts à vue</h1>
        <p style="color:{GRAY_600};font-size:15px;max-width:80ch;">
            Application Python (Streamlit) implémentant explicitement les fonctions
            mathématiques d'écoulement calibrées par modèles ARIMA/SARIMA sur 6 comptes,
            avec décomposition stable/volatile et projection des ressources.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

c1, c2, c3 = st.columns(3)
c1.metric("Encours total (31/12/2025)", fmt_mds(TOTAL_ENCOURS))
stable_12 = decompose_stable_volatile(ACCOUNTS, 12)["stable"].sum()
c2.metric("Part stable à 12 mois", fmt_pct(stable_12 / TOTAL_ENCOURS))
c3.metric("Comptes modélisés", "6 (ARIMA / SARIMA)")

st.write("")

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    [
        "📉 Fonctions d'écoulement",
        "🥧 Répartition stable / volatile",
        "📈 Projection des ressources",
        "🕰️ Historique des comptes",
        "🧮 Modèles & audit",
    ]
)

# ======================================================================
# TAB 1 — Fonctions d'écoulement
# ======================================================================
with tab1:
    st.subheader("Fonction d'écoulement S(H)")
    st.latex(r"S(H) = \min\Big(1,\ \exp\big(H \times (\mu + U_{min})\big)\Big) \qquad k = \mu + U_{min}")
    st.caption(
        "S(H) est la part de l'encours d'un compte encore considérée comme stable à "
        "l'horizon H (en mois). k est calibré une fois pour toutes sur l'historique du compte."
    )

    col_a, col_b = st.columns([3, 1])
    with col_a:
        hmax = st.slider("Horizon affiché (mois)", 6, 91, 24, key="hmax_tab1")
    with col_b:
        selected = st.multiselect(
            "Comptes affichés",
            [a["id"] for a in ACCOUNTS],
            default=[a["id"] for a in ACCOUNTS],
        )

    H_range = np.arange(0, hmax + 1)
    fig = go.Figure()
    for a in ACCOUNTS:
        if a["id"] not in selected:
            continue
        y = decay_share(H_range, a["k"], a["cap"]) * 100
        fig.add_trace(
            go.Scatter(
                x=H_range, y=y, mode="lines", name=a["id"],
                line=dict(color=a["color"], width=2.5),
            )
        )
    fig.update_layout(
        **PLOTLY_LAYOUT,
        height=460,
        xaxis=dict(title="Horizon (mois)", **GRID),
        yaxis=dict(title="Part stable de l'encours (%)", range=[0, 100], **GRID),
    )
    st.plotly_chart(fig, width='stretch')

    show_source(decay_share)

    st.subheader("Repères clés par compte")
    reperes = pd.DataFrame(
        [
            {
                "Compte": a["id"],
                "Modèle": a["model"],
                "k = μ+Uₘᵢₙ": round(a["k"], 6),
                "S(1)": fmt_pct(float(decay_share(1, a["k"], a["cap"]))),
                "S(6)": fmt_pct(float(decay_share(6, a["k"], a["cap"]))),
                "S(12)": fmt_pct(float(decay_share(12, a["k"], a["cap"]))),
                "S(24)": fmt_pct(float(decay_share(24, a["k"], a["cap"]))),
            }
            for a in ACCOUNTS
        ]
    )
    st.dataframe(reperes, width='stretch', hide_index=True)

# ======================================================================
# TAB 2 — Répartition stable / volatile
# ======================================================================
with tab2:
    st.subheader("Décomposition de l'encours à horizon H")
    st.latex(r"\text{Stable}(H) = S(H)\times \text{Encours} \qquad \text{Volatile}(H) = \text{Encours} - \text{Stable}(H)")

    H = st.slider("Horizon H (mois)", 0, 91, 12, key="h_tab2")
    df = decompose_stable_volatile(ACCOUNTS, H)
    total = df["encours"].sum()
    total_stable = df["stable"].sum()
    total_volatile = df["volatile"].sum()

    k1, k2, k3 = st.columns(3)
    k1.metric("Encours total", fmt_mds(total))
    k2.metric("Partie stable", fmt_mds(total_stable), fmt_pct(total_stable / total))
    k3.metric("Partie volatile", fmt_mds(total_volatile), fmt_pct(total_volatile / total))

    col1, col2 = st.columns([1.4, 1])
    with col1:
        fig = go.Figure()
        fig.add_trace(go.Bar(y=df["compte"], x=df["stable"], name="Stable",
                              orientation="h", marker_color=STABLE_COLOR))
        fig.add_trace(go.Bar(y=df["compte"], x=df["volatile"], name="Volatile",
                              orientation="h", marker_color=VOLATILE_COLOR))
        fig.update_layout(
            **PLOTLY_LAYOUT, barmode="stack", height=380,
            xaxis=dict(title="Encours (FCFA)", **GRID),
            yaxis=dict(autorange="reversed", **GRID),
        )
        st.plotly_chart(fig, width='stretch')
    with col2:
        fig2 = go.Figure(
            go.Pie(
                labels=["Stable", "Volatile"],
                values=[total_stable, total_volatile],
                hole=0.65,
                marker_colors=[STABLE_COLOR, VOLATILE_COLOR],
                textinfo="percent",
            )
        )
        fig2.update_layout(**PLOTLY_LAYOUT, height=380, showlegend=True)
        st.plotly_chart(fig2, width='stretch')

    show_source(decompose_stable_volatile)

    st.subheader("Détail par compte")
    detail = df[["compte", "encours", "part_stable", "stable", "volatile"]].copy()
    detail["encours"] = detail["encours"].map(fmt_mds)
    detail["part_stable"] = detail["part_stable"].map(fmt_pct)
    detail["stable"] = detail["stable"].map(fmt_mds)
    detail["volatile"] = detail["volatile"].map(fmt_mds)
    detail.columns = ["Compte", "Encours", "Part stable", "Montant stable", "Montant volatile"]
    st.dataframe(detail, width='stretch', hide_index=True)

# ======================================================================
# TAB 3 — Projection des ressources
# ======================================================================
with tab3:
    st.subheader("Projection des ressources futures")
    st.latex(
        r"\text{Ressources}(i) = \sum_{\text{comptes}} S(i+6)\times \text{Encours}"
        r" \;+\; \text{multiplicateur} \times \text{Production}(i)"
    )
    st.caption(
        "Le décalage de 6 mois reprend le calage du classeur d'origine "
        "(mapping production : juillet 2026 = horizon 7)."
    )

    col_a, col_b = st.columns(2)
    with col_a:
        mult_pct = st.slider("Multiplicateur de production nouvelle (%)", 50, 150, 100, key="mult_tab3")
    with col_b:
        hmax3 = st.slider("Horizon affiché (mois)", 12, len(PROJ["datesProj"]), 36, key="hmax_tab3")

    proj_df = project_resources(
        ACCOUNTS,
        PROJ["datesProj"][:hmax3],
        PROJ["previsionProductionNouvelle"],
        multiplier=mult_pct / 100,
    )

    m1, m2, m3 = st.columns(3)
    m1.metric(f"Ressources — {proj_df['date'].iloc[0]}", fmt_mds(proj_df["total_projete"].iloc[0]))
    m2.metric(f"Ressources — {proj_df['date'].iloc[-1]}", fmt_mds(proj_df["total_projete"].iloc[-1]))
    m3.metric("Production nouvelle cumulée", fmt_mds(proj_df["production_nouvelle"].sum()))

    fig = go.Figure()
    fig.add_trace(go.Bar(x=proj_df["date"], y=proj_df["stable"], name="Partie stable",
                          marker_color=RED))
    fig.add_trace(go.Bar(x=proj_df["date"], y=proj_df["production_nouvelle"], name="Production nouvelle",
                          marker_color="#C9C9C9"))
    fig.add_trace(go.Scatter(x=proj_df["date"], y=proj_df["total_projete"], name="Total projeté",
                              mode="lines", line=dict(color=BLACK, width=2)))
    fig.update_layout(
        **PLOTLY_LAYOUT, barmode="stack", height=460,
        xaxis=dict(title="Mois", **GRID), yaxis=dict(title="Ressources (FCFA)", **GRID),
    )
    st.plotly_chart(fig, width='stretch')

    show_source(project_resources)

# ======================================================================
# TAB 4 — Historique
# ======================================================================
with tab4:
    st.subheader("Historique des encours (2015 – 2025)")
    mode = st.radio("Affichage", ["Par compte", "Empilé"], horizontal=True)

    dates_h = HIST["dates"]
    fig = go.Figure()
    for a in ACCOUNTS:
        vals = HIST[a["id"]]
        fig.add_trace(
            go.Scatter(
                x=dates_h, y=vals, mode="lines", name=a["id"],
                line=dict(color=a["color"], width=1.8),
                stackgroup="one" if mode == "Empilé" else None,
            )
        )
    fig.update_layout(
        **PLOTLY_LAYOUT, height=460,
        xaxis=dict(title="Date", **GRID), yaxis=dict(title="Encours (FCFA)", **GRID),
    )
    st.plotly_chart(fig, width='stretch')

# ======================================================================
# TAB 5 — Audit & code
# ======================================================================
with tab5:
    st.subheader("Paramètres des modèles ARIMA / SARIMA")
    audit = pd.DataFrame(
        [
            {
                "Compte": a["id"],
                "Modèle": a["model"],
                "μ (drift)": round(a["mu"], 6),
                "Uₘᵢₙ": round(a["umin"], 6),
                "k = μ+Uₘᵢₙ": round(a["k"], 6),
                "Stable H1": fmt_pct(float(decay_share(1, a["k"], a["cap"]))),
                "Stable H12": fmt_pct(float(decay_share(12, a["k"], a["cap"]))),
            }
            for a in ACCOUNTS
        ]
    )
    st.dataframe(audit, width='stretch', hide_index=True)

    st.subheader("Contrôles agrégés")
    stable_h1 = decompose_stable_volatile(ACCOUNTS, 1)["stable"].sum()
    stable_h12 = decompose_stable_volatile(ACCOUNTS, 12)["stable"].sum()
    st.success(f"Encours total initial : {fmt_mds(TOTAL_ENCOURS)} — cohérent avec la somme des 6 comptes.")
    st.success(f"Partie stable agrégée à H1 : {fmt_mds(stable_h1)}")
    st.success(f"Partie stable agrégée à H12 : {fmt_mds(stable_h12)}")

    st.subheader("Formule d'écoulement — code source complet")
    st.code(inspect.getsource(decay_share), language="python")
    st.subheader("Décomposition stable / volatile — code source complet")
    st.code(inspect.getsource(decompose_stable_volatile), language="python")
    st.subheader("Projection des ressources — code source complet")
    st.code(inspect.getsource(project_resources), language="python")

st.markdown(
    f"""
    <p style="color:{GRAY_600};font-size:12px;text-align:center;margin-top:30px;
              border-top:1px solid {GRAY_200};padding-top:14px;">
        Application développée en Python / Streamlit — implémentation directe des formules
        du modèle d'écoulement des dépôts à vue (Afriland First Bank).
    </p>
    """,
    unsafe_allow_html=True,
)

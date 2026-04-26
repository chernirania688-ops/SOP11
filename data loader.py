"""
data_loader.py — Chargeur générique de données S&OP
=====================================================
Détecte automatiquement la structure de chaque fichier Excel
et retourne des structures Python propres utilisables par tous les agents.
Fonctionne pour TOUT fichier ayant la même structure que les fichiers d'exemple.
"""

import re
import warnings
from pathlib import Path

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")


# ═══════════════════════════════════════════════════════════════════════════
# UTILITAIRES
# ═══════════════════════════════════════════════════════════════════════════

def clean_numeric(val) -> float:
    """Convertit n'importe quelle valeur en float (gère '1 400,00', '3 840.00', etc.)."""
    if val is None:
        return np.nan
    if isinstance(val, (int, float)):
        return float(val) if not (isinstance(val, float) and np.isnan(val)) else np.nan
    if isinstance(val, str):
        val = val.strip().replace("\u202f", "").replace("\xa0", "").replace(" ", "")
        val = val.replace(",", ".")
        try:
            return float(val)
        except ValueError:
            return np.nan
    return np.nan


def is_date_column(col_name: str) -> bool:
    """Détecte si un nom de colonne est une période temporelle."""
    col = str(col_name).strip()
    patterns = [
        r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{4}$",  # Jan-2020
        r"^M\d{2}\s+Y\d{4}$",                                           # M01 Y2020
        r"^W\d{2}\s+Y\d{2,4}$",                                         # W34 Y23
        r"^\d{4}-\d{2}$",                                                # 2020-01
        r"^Q[1-4]\s+\d{4}$",                                             # Q1 2020
    ]
    return any(re.match(p, col, re.IGNORECASE) for p in patterns)


def extract_time_columns(df: pd.DataFrame) -> list:
    """Retourne la liste ordonnée des colonnes temporelles détectées."""
    return [c for c in df.columns if is_date_column(str(c))]


# ═══════════════════════════════════════════════════════════════════════════
# CHARGEUR FICHIER DEMANDE (article_report_merged structure)
# ═══════════════════════════════════════════════════════════════════════════

def load_demand_file(filepath: str) -> dict:
    """
    Charge un fichier de demande ayant la structure article_report_merged.xlsx.
    Retourne:
        {
          "articles": { article_id: { "history": pd.Series, "forecast": pd.Series } },
          "kpis": pd.DataFrame,  # MAE/MAPE/RMSE si disponible
          "methods": { article_id: { method_name: pd.Series } },
          "time_cols": list,
          "meta": dict
        }
    """
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"Fichier introuvable : {filepath}")

    xl = pd.ExcelFile(filepath)
    sheets = xl.sheet_names
    result = {"articles": {}, "kpis": None, "methods": {}, "time_cols": [], "meta": {"filepath": str(filepath), "sheets": sheets}}

    # ── Feuille Series (historique + forecast statistique) ─────────────────
    series_sheet = next((s for s in sheets if "series" in s.lower() or "serie" in s.lower()), sheets[0])
    df_series = pd.read_excel(filepath, sheet_name=series_sheet, header=0)

    # Détecte la colonne article et la colonne data field (dynamiquement)
    art_col   = df_series.columns[0]
    field_col = df_series.columns[1]
    time_cols = extract_time_columns(df_series)
    result["time_cols"] = time_cols

    current_article = None
    for _, row in df_series.iterrows():
        art_val   = row[art_col]
        field_val = str(row[field_col]).strip() if pd.notna(row[field_col]) else ""

        if pd.notna(art_val) and str(art_val).strip() and str(art_val).strip().lower() != "nan":
            current_article = str(art_val).strip()

        if not current_article or not field_val or field_val.lower() == "nan":
            continue

        values = pd.Series(
            [clean_numeric(row[c]) for c in time_cols],
            index=time_cols,
            dtype=float,
        )

        if current_article not in result["articles"]:
            result["articles"][current_article] = {"history": None, "forecast": None, "all_fields": {}}

        result["articles"][current_article]["all_fields"][field_val] = values

        field_lower = field_val.lower()
        if "calculation history" in field_lower:
            result["articles"][current_article]["history"] = values
        elif "statistical history and forecast" in field_lower:
            result["articles"][current_article]["forecast"] = values

    # ── Feuille KPIs ────────────────────────────────────────────────────────
    kpi_sheet = next((s for s in sheets if "kpi" in s.lower() or "compare" in s.lower()), None)
    if kpi_sheet:
        df_kpi = pd.read_excel(filepath, sheet_name=kpi_sheet, header=0)
        df_kpi.columns = [str(c).strip() for c in df_kpi.columns]
        result["kpis"] = df_kpi

    # ── Feuille Methods ─────────────────────────────────────────────────────
    meth_sheet = next((s for s in sheets if "method" in s.lower()), None)
    if meth_sheet:
        df_meth = pd.read_excel(filepath, sheet_name=meth_sheet, header=0)
        art_col_m   = df_meth.columns[0]
        field_col_m = df_meth.columns[1]
        time_cols_m = extract_time_columns(df_meth)

        current_article = None
        for _, row in df_meth.iterrows():
            art_val   = row[art_col_m]
            field_val = str(row[field_col_m]).strip() if pd.notna(row[field_col_m]) else ""

            if pd.notna(art_val) and str(art_val).strip() and str(art_val).strip().lower() != "nan":
                current_article = str(art_val).strip()

            if not current_article or not field_val or field_val.lower() == "nan":
                continue

            values = pd.Series(
                [clean_numeric(row[c]) for c in time_cols_m],
                index=time_cols_m,
                dtype=float,
            )

            if current_article not in result["methods"]:
                result["methods"][current_article] = {}
            result["methods"][current_article][field_val] = values

    n_articles = len(result["articles"])
    n_periods  = len(time_cols)
    print(f"  ✅ Fichier demande : {n_articles} articles × {n_periods} périodes")
    return result


# ═══════════════════════════════════════════════════════════════════════════
# CHARGEUR FICHIER PRODUCTION (exemple_prod structure)
# ═══════════════════════════════════════════════════════════════════════════

def load_production_file(filepath: str) -> dict:
    """
    Charge un fichier MPS ayant la structure exemple_prod.xlsx.
    Retourne:
        {
          "articles": { article_id: { indicator_name: pd.Series } },
          "time_cols": list,
          "resources": list,
          "meta": dict
        }
    """
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"Fichier introuvable : {filepath}")

    xl = pd.ExcelFile(filepath)
    sheets = xl.sheet_names

    # Essaie chaque feuille jusqu'à en trouver une avec des données temporelles
    df = None
    for sh in sheets:
        candidate = pd.read_excel(filepath, sheet_name=sh, header=0)
        time_cols  = extract_time_columns(candidate)
        if len(time_cols) >= 3:
            df = candidate
            break

    if df is None:
        raise ValueError(f"Aucune feuille avec colonnes temporelles trouvée dans {filepath}")

    time_cols = extract_time_columns(df)

    # Colonnes article, indicateur, ressource (positionnelles + heuristiques)
    art_col      = df.columns[0]
    donnees_col  = df.columns[1]
    resource_col = df.columns[2] if len(df.columns) > len(time_cols) + 2 else None

    result = {
        "articles": {},
        "time_cols": time_cols,
        "resources": [],
        "meta": {"filepath": str(filepath), "sheets": sheets},
    }

    current_article = None
    for _, row in df.iterrows():
        art_val  = row[art_col]
        ind_val  = str(row[donnees_col]).strip() if pd.notna(row[donnees_col]) else ""
        res_val  = str(row[resource_col]).strip() if resource_col and pd.notna(row[resource_col]) else None

        if pd.notna(art_val) and str(art_val).strip() and str(art_val).strip().lower() != "nan":
            current_article = str(art_val).strip()

        if not current_article or not ind_val or ind_val.lower() == "nan":
            continue

        values = pd.Series(
            [clean_numeric(row[c]) for c in time_cols],
            index=time_cols,
            dtype=float,
        )

        if current_article not in result["articles"]:
            result["articles"][current_article] = {}

        result["articles"][current_article][ind_val] = values

        if res_val and res_val.lower() != "nan" and res_val not in result["resources"]:
            result["resources"].append(res_val)

    n_articles = len(result["articles"])
    n_periods  = len(time_cols)
    print(f"  ✅ Fichier production : {n_articles} article(s) × {n_periods} périodes | ressources: {result['resources']}")
    return result


# ═══════════════════════════════════════════════════════════════════════════
# DÉTECTEUR AUTOMATIQUE DE TYPE DE FICHIER
# ═══════════════════════════════════════════════════════════════════════════

def auto_load(filepath: str) -> dict:
    """
    Charge automatiquement un fichier Excel en détectant son type.
    Retourne un dict standardisé avec la clé 'type': 'demand' | 'production' | 'unknown'
    """
    path = Path(filepath)
    xl   = pd.ExcelFile(filepath)
    sheets = xl.sheet_names

    # Lit la première feuille pour heuristique
    df_peek = pd.read_excel(filepath, sheet_name=sheets[0], header=0, nrows=5)
    cols    = [str(c).lower() for c in df_peek.columns]

    # Heuristique : production si "donnees" ou "gross requirements" visible
    has_prod_keywords = any(
        kw in " ".join(cols) for kw in ["donnees", "gross", "production plan", "capacity", "safety stock"]
    )
    # Heuristique : demande si "data field" ou "article" + "calculation"
    has_demand_keywords = any(
        kw in " ".join(cols) for kw in ["data field", "data fields", "calculation"]
    ) or any("series" in s.lower() or "kpi" in s.lower() for s in sheets)

    if has_prod_keywords and not has_demand_keywords:
        data = load_production_file(filepath)
        data["type"] = "production"
    elif has_demand_keywords or "series" in " ".join(s.lower() for s in sheets):
        data = load_demand_file(filepath)
        data["type"] = "demand"
    else:
        # Fallback : essaie demande d'abord
        try:
            data = load_demand_file(filepath)
            data["type"] = "demand"
        except Exception:
            data = load_production_file(filepath)
            data["type"] = "production"

    print(f"  🔍 Type détecté : {data['type']} — {path.name}")
    return data


# ═══════════════════════════════════════════════════════════════════════════
# EXTRACTION DE L'HISTORIQUE RÉEL (série temporelle propre)
# ═══════════════════════════════════════════════════════════════════════════

def get_clean_history(article_data: dict) -> pd.Series:
    """
    Extrait la série historique réelle d'un article (valeurs > 0, sans NaN).
    Fonctionne quelle que soit la taille ou la période des données.
    """
    history = article_data.get("history")
    if history is None:
        # Cherche dans all_fields
        all_f = article_data.get("all_fields", {})
        for k, v in all_f.items():
            if "history" in k.lower() and "statistical" not in k.lower():
                history = v
                break
    if history is None:
        return pd.Series(dtype=float)

    series = history.dropna()
    series = series[series > 0]
    return series


def get_existing_forecast(article_data: dict) -> pd.Series:
    """Extrait la prévision statistique existante (baseline ERP)."""
    forecast = article_data.get("forecast")
    if forecast is None:
        all_f = article_data.get("all_fields", {})
        for k, v in all_f.items():
            if "statistical" in k.lower() and "forecast" in k.lower():
                forecast = v
                break
    if forecast is None:
        return pd.Series(dtype=float)
    return forecast.dropna()


def get_production_indicator(prod_article: dict, indicator_keywords: list) -> pd.Series:
    """
    Cherche un indicateur MPS par mots-clés (insensible à la casse).
    Exemple: get_production_indicator(data, ["gross", "requirement"])
    """
    for key, series in prod_article.items():
        key_lower = key.lower()
        if all(kw.lower() in key_lower for kw in indicator_keywords):
            return series
    return pd.Series(dtype=float)
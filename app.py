"""
app.py — S&OP Agentique · Version corrigée
==========================================
Corrections :
- sys.path configuré AVANT tout import de module
- Boutons gérés via st.session_state (pas de rerun prématuré)
- Chaque agent a sa logique isolée
- Messages d'erreur visibles si un module manque
"""

import sys
import os
import re
import warnings
import tempfile
import importlib.util
from pathlib import Path
from datetime import datetime

import streamlit as st
import pandas as pd
import numpy as np

warnings.filterwarnings("ignore")

# ── PATH ROBUSTE : fonctionne même si __file__ n'est pas défini ───────────────
# Streamlit peut changer le cwd — on détecte le dossier de app.py de 3 façons
def _get_app_dir() -> Path:
    # Méthode 1 : __file__ standard
    try:
        return Path(__file__).parent.resolve()
    except NameError:
        pass
    # Méthode 2 : depuis st.session_state s'il a été sauvegardé
    if hasattr(st, "session_state") and "_app_dir" in st.session_state:
        return Path(st.session_state._app_dir)
    # Méthode 3 : cherche data_loader.py dans les dossiers probables
    for candidate in [Path.cwd(), Path.cwd().parent, Path("/app"), Path("/mount/src")]:
        for sub in [candidate] + list(candidate.glob("*/")):
            if (sub / "data_loader.py").exists():
                return sub
    return Path.cwd()

APP_DIR = _get_app_dir()
# Sauvegarde pour les reruns
try:
    if "_app_dir" not in st.session_state:
        st.session_state["_app_dir"] = str(APP_DIR)
except: pass

# Ajoute au sys.path de toutes les façons possibles
for p in [str(APP_DIR), str(APP_DIR.parent)]:
    if p not in sys.path:
        sys.path.insert(0, p)
os.chdir(str(APP_DIR))

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="S&OP Agentique",
    page_icon="⬡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ══════════════════════════════════════════════════════════════════════════════
# CSS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Syne:wght@400;600;700&display=swap');
:root{
  --bg:#0a0a0b;--bg1:#111113;--bg2:#18181c;--bg3:#222228;
  --border:#2a2a32;--border2:#3a3a45;
  --text:#e8e8f0;--muted:#6b6b80;
  --amber:#f59e0b;--green:#10b981;--red:#ef4444;--blue:#3b82f6;
}
html,body,[class*="css"]{font-family:'Syne',sans-serif !important;background:var(--bg) !important;color:var(--text) !important}
#MainMenu,footer,header{visibility:hidden}
.block-container{padding:1rem !important;max-width:100% !important}

.topbar{background:var(--bg1);border:1px solid var(--border);border-radius:10px;padding:.7rem 1.2rem;display:flex;align-items:center;gap:1rem;margin-bottom:1rem}
.topbar-logo{font-family:'Space Mono',monospace;font-size:.9rem;font-weight:700;color:var(--amber)}
.topbar-sep{width:1px;height:16px;background:var(--border2)}
.topbar-sub{font-size:.72rem;color:var(--muted);font-family:'Space Mono',monospace}

.stTabs [data-baseweb="tab-list"]{background:var(--bg1) !important;border-radius:8px 8px 0 0;gap:0 !important}
.stTabs [data-baseweb="tab"]{background:transparent !important;color:var(--muted) !important;font-family:'Syne',sans-serif !important;font-weight:600 !important;padding:.65rem 1.2rem !important;border-bottom:2px solid transparent !important;border-radius:0 !important;font-size:.82rem !important}
.stTabs [aria-selected="true"]{color:var(--amber) !important;border-bottom-color:var(--amber) !important;background:var(--bg2) !important}
.stTabs [data-baseweb="tab-panel"]{background:var(--bg) !important;padding:0 !important}

.agent-panel{background:var(--bg1);border:1px solid var(--border);border-radius:8px;padding:.75rem;height:100%}
.sec-label{font-size:.67rem;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);font-family:'Space Mono',monospace;padding:.4rem 0 .6rem;border-bottom:1px solid var(--border);margin-bottom:.6rem}

.chat-area{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:.75rem;min-height:320px;max-height:400px;overflow-y:auto}
.msg-user{display:flex;justify-content:flex-end;margin:.4rem 0}
.msg-agent{display:flex;justify-content:flex-start;margin:.4rem 0}
.bubble{padding:.55rem .85rem;border-radius:8px;font-size:.83rem;line-height:1.6;max-width:88%;border:1px solid var(--border)}
.bubble-u{background:var(--bg3);color:var(--text)}
.bubble-a{background:var(--bg1);color:var(--text)}
.avatar{width:24px;height:24px;border-radius:5px;display:flex;align-items:center;justify-content:center;font-size:.75rem;font-weight:700;flex-shrink:0;margin:.1rem .4rem 0}

.stTextInput>div>div>input{background:var(--bg2) !important;border:1px solid var(--border2) !important;color:var(--text) !important;border-radius:6px !important;font-family:'Syne',sans-serif !important}
.stTextInput>div>div>input:focus{border-color:var(--amber) !important;box-shadow:0 0 0 1px var(--amber) !important}
.stButton>button{background:var(--bg2) !important;color:var(--text) !important;border:1px solid var(--border2) !important;border-radius:6px !important;font-weight:600 !important;font-family:'Syne',sans-serif !important;transition:all .15s !important}
.stButton>button:hover{border-color:var(--amber) !important;color:var(--amber) !important}
.btn-send>button{background:var(--amber) !important;color:#000 !important;border:none !important;font-weight:700 !important}
.stFileUploader>div{background:var(--bg2) !important;border:1px solid var(--border2) !important;border-radius:8px !important}

.trace-item{padding:.35rem .6rem;border-bottom:1px solid var(--border);font-size:.76rem;display:flex;gap:.5rem;align-items:flex-start}
.trace-dot{width:6px;height:6px;border-radius:50%;flex-shrink:0;margin-top:.3rem}
.trace-agent{font-family:'Space Mono',monospace;font-size:.72rem;font-weight:700;min-width:75px;flex-shrink:0}
.trace-msg{color:var(--muted);flex:1;line-height:1.4}
.trace-ts{color:var(--border2);font-family:'Space Mono',monospace;font-size:.67rem;flex-shrink:0}

.fbadge{background:var(--bg2);border:1px solid var(--green);border-radius:5px;padding:.3rem .6rem;font-size:.75rem;font-family:'Space Mono',monospace;color:var(--green);margin:.4rem 0}
.cap-item{padding:.25rem 0;font-size:.78rem;color:var(--muted);display:flex;gap:.4rem;align-items:center}
.cap-dot{color:var(--amber);font-size:.6rem}

.warn-box{background:#2d1f00;border:1px solid var(--amber);border-radius:6px;padding:.5rem .75rem;font-size:.8rem;color:var(--amber);margin:.4rem 0}
.err-box{background:#2d0a0a;border:1px solid var(--red);border-radius:6px;padding:.5rem .75rem;font-size:.8rem;color:var(--red);margin:.4rem 0}

::-webkit-scrollbar{width:4px}
::-webkit-scrollbar-track{background:var(--bg1)}
::-webkit-scrollbar-thumb{background:var(--border2);border-radius:4px}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# IMPORT MODULES AGENTS (robuste)
# ══════════════════════════════════════════════════════════════════════════════
def import_modules():
    """
    Importe les modules agents depuis leur chemin absolu.
    Utilise spec_from_file_location pour éviter les problèmes de sys.path avec Streamlit.
    """
    import importlib.util as ilu

    result = {}
    modules_cfg = {
        "data_loader":      ["auto_load", "load_demand_file", "load_production_file", "get_clean_history"],
        "agent_demande":    ["run", "analyse_article"],
        "agent_production": ["run", "analyse_capacity", "generate_adjusted_plan"],
        "agent_marketing":  ["run", "analyse_article_promo"],
        "agent_finance":    ["run", "compute_pl_scenario", "estimate_financials_from_history"],
    }
    aliases = {
        "data_loader":      {"auto_load":"auto_load","load_demand_file":"load_dem","load_production_file":"load_prod","get_clean_history":"get_hist"},
        "agent_demande":    {"run":"rd","analyse_article":"analyse_art"},
        "agent_production": {"run":"rp","analyse_capacity":"analyse_cap","generate_adjusted_plan":"gen_adj"},
        "agent_marketing":  {"run":"rm","analyse_article_promo":"analyse_promo"},
        "agent_finance":    {"run":"rf","compute_pl_scenario":"compute_pl","estimate_financials_from_history":"est_fin"},
    }
    errors = []

    def load_module(name, path):
        """Import un module depuis son chemin absolu."""
        if name in sys.modules:
            return sys.modules[name]
        spec = ilu.spec_from_file_location(name, str(path))
        if spec is None:
            raise ImportError(f"Impossible de charger {path}")
        mod = ilu.module_from_spec(spec)
        sys.modules[name] = mod  # enregistre avant exec pour éviter imports circulaires
        spec.loader.exec_module(mod)
        return mod

    # Ordre important : data_loader doit être chargé en premier
    load_order = ["data_loader", "excel_writer", "agent_demande", "agent_production", "agent_marketing", "agent_finance"]

    for mod_name in load_order:
        py_file = APP_DIR / f"{mod_name}.py"
        if not py_file.exists():
            errors.append(f"❌ {mod_name}.py introuvable dans {APP_DIR}")
            continue
        try:
            mod = load_module(mod_name, py_file)
            attrs = modules_cfg.get(mod_name, [])
            for attr in attrs:
                alias = aliases.get(mod_name, {}).get(attr, attr)
                result[alias] = getattr(mod, attr)
        except Exception as e:
            errors.append(f"❌ {mod_name} : {e}")

    result["_errors"] = errors
    result["_app_dir"] = str(APP_DIR)
    return result

# Import une seule fois
if "BE" not in st.session_state:
    st.session_state.BE = import_modules()
BE = st.session_state.BE

# ══════════════════════════════════════════════════════════════════════════════
# STATE
# ══════════════════════════════════════════════════════════════════════════════
AGENTS = ["orchestrateur", "demande", "production", "marketing", "finance"]
AGENT_CFG = {
    "orchestrateur": {"icon":"⬡",  "label":"Orchestrateur","color":"#f59e0b","av":"background:#2d1f00;color:#f59e0b"},
    "demande":       {"icon":"📈", "label":"Demande",       "color":"#10b981","av":"background:#001a0f;color:#10b981"},
    "production":    {"icon":"🏭", "label":"Production",    "color":"#f97316","av":"background:#1f0d00;color:#f97316"},
    "marketing":     {"icon":"📣", "label":"Marketing",     "color":"#ec4899","av":"background:#1f0011;color:#ec4899"},
    "finance":       {"icon":"💰", "label":"Finance",       "color":"#3b82f6","av":"background:#00103f;color:#3b82f6"},
}
QUICK = {
    "orchestrateur": ["Lance une analyse S&OP complète","Simule une promotion -20%","Vérifie la capacité de production","La promo -20% est-elle rentable ?"],
    "demande":       ["Calcule le forecast 6 mois","Quelle est la meilleure méthode ?","Analyse la saisonnalité","Quel article a le MAPE le plus élevé ?"],
    "production":    ["Analyse la capacité Fill-L1","Y a-t-il des surcharges ?","Calcule le MRP complet","Simulation demande +30%"],
    "marketing":     ["Simule une promo -20%","Quel est le meilleur mois ?","Compare promo -10% vs -30%","Calcule l'uplift attendu"],
    "finance":       ["Calcule le ROI promotion","Le budget est-il respecté ?","Quel est le seuil de rentabilité ?","Alerte si marge < 10%"],
}
CAPS = {
    "orchestrateur": ["Coordination de tous les agents","Décision GO / NO-GO","Trace live des appels","Synthèse multi-agents"],
    "demande":       ["Prévision LES / Holt-Winters / ARIMA","Classification auto des séries","Sélection par MAE/MAPE/RMSE","Détection anomalies & saisonnalité"],
    "production":    ["Calcul MRP complet (8 indicateurs)","Détection surcharges & alertes","Ajustement plan de production","Simulation scenarios demande"],
    "marketing":     ["Simulation promotions élasticité prix","Calcul uplift par article","Analyse saisonnalité réelle","Comparaison multi-scénarios"],
    "finance":       ["Calcul P&L & ROI par scénario","Seuil de rentabilité automatique","Vérification budget promotion","Alertes marge & déficit"],
}

def init_state():
    if "files"  not in st.session_state: st.session_state.files  = {}
    if "chats"  not in st.session_state: st.session_state.chats  = {a: [] for a in AGENTS}
    if "traces" not in st.session_state: st.session_state.traces = []
    for a in AGENTS:
        if a not in st.session_state.chats:
            st.session_state.chats[a] = []

init_state()

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def save_upload(agent, uploaded):
    tmp  = Path(tempfile.mkdtemp())
    dest = tmp / uploaded.name
    dest.write_bytes(uploaded.getvalue())
    try:    df_p = pd.read_excel(str(dest), nrows=6)
    except: df_p = pd.DataFrame()
    st.session_state.files[agent] = {"name": uploaded.name, "path": str(dest), "df": df_p}
    return str(dest)

def add_msg(agent, role, content, excel=None):
    st.session_state.chats[agent].append({
        "role": role, "content": content,
        "ts": datetime.now().isoformat(), "excel": excel
    })

def add_trace(agent, msg, status="info"):
    st.session_state.traces.append({
        "agent": agent, "msg": msg, "status": status,
        "ts": datetime.now().strftime("%H:%M:%S")
    })

def get_file(agent):
    """Retourne le chemin du fichier pour un agent, avec auto-détection."""
    p = st.session_state.files.get(agent, {}).get("path")
    if p and Path(p).exists():
        return p
    # Auto-détection depuis les fichiers disponibles
    if "auto_load" in BE:
        for fp in [v["path"] for v in st.session_state.files.values() if "path" in v]:
            try:
                d = BE["auto_load"](fp)
                if agent == "production" and d["type"] == "production": return fp
                if agent in ("demande","marketing","finance") and d["type"] == "demand": return fp
            except: pass
    # Dernier recours : premier fichier disponible
    paths = [v["path"] for v in st.session_state.files.values() if "path" in v]
    return paths[0] if paths else None

def fmt_html(text):
    """Convertit le markdown basique en HTML pour les bulles de chat."""
    import html
    text = html.escape(text)
    text = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', text)
    text = re.sub(r'`(.*?)`', r'<code style="background:#222;padding:.1rem .3rem;border-radius:3px;font-family:monospace;color:#fbbf24">\1</code>', text)
    text = text.replace("\n", "<br>")
    return text

# ══════════════════════════════════════════════════════════════════════════════
# CERVEAU DE CHAQUE AGENT
# ══════════════════════════════════════════════════════════════════════════════
def think(agent: str, question: str) -> tuple:
    """
    Cerveau de l'agent : analyse la question, exécute, retourne (réponse, excel_path).
    Toujours retourne quelque chose — jamais silencieux.
    """
    q = question.lower().strip()
    parts = []
    excel_out = None
    fp = get_file(agent)

    # Vérifie les imports
    if BE.get("_errors"):
        errs = [e for e in BE["_errors"] if agent.replace("_","") in e.lower() or "loader" in e.lower()]
        if errs:
            return f"⚠️ Module non chargé : {errs[0]}\n\nVérifie que tous les fichiers .py sont dans le même dossier.", None

    try:
        # ══ AGENT DEMANDE ══════════════════════════════════════════════════════
        if agent == "demande":
            if not fp:
                return "📁 **Importe un fichier Excel** de demande (ex: article_report_merged.xlsx) pour que je puisse analyser.", None

            data = BE["load_dem"](fp)
            arts = data["articles"]
            tc   = data["time_cols"]
            n    = len(arts)

            if not arts:
                return "⚠️ Aucun article détecté dans le fichier. Vérifie la structure (colonnes Article + Data field).", None

            pct = next((int(m) for m in re.findall(r"(\d+)\s*%", q)), 0)

            if any(w in q for w in ["forecast","prévision","prévoir","futur","mois","calcule"]):
                parts.append(f"✅ **Calcul du forecast — {n} article(s) :**\n")
                tmp = str(Path(fp).parent / "output_dem.xlsx")
                BE["rd"](context={"demand_file": fp, "output_path": tmp})
                excel_out = tmp
                # Résumé rapide
                for aid, adata in list(arts.items())[:4]:
                    hist = BE["get_hist"](adata)
                    if len(hist) < 4: continue
                    try:
                        sel = BE["analyse_art"](aid, adata, {})
                        if not sel: continue
                        row = sel["all_results"][sel["all_results"]["Méthode"]==sel["best_method"]].iloc[0]
                        parts.append(f"**{aid}** → `{sel['best_method']}` | MAPE=`{row['MAPE(%)']:.1f}%` | F6M=**{np.sum(sel['forecast']):,.0f} U**")
                    except: pass
                parts.append(f"\n📥 Fichier Excel généré avec le plan de demande complet.")

            elif any(w in q for w in ["mape","kpi","qualité","erreur","précision","audit","meilleur"]):
                parts.append("**📊 Audit qualité des prévisions :**\n")
                kdf = data.get("kpis")
                if kdf is not None and not kdf.empty:
                    for _, r in kdf.iterrows():
                        try:
                            mape = float(str(r.iloc[2]).replace(" ","").replace(",","."))
                            st_txt = "🔴 Mauvais" if mape>50 else ("🟡 Moyen" if mape>25 else "🟢 Bon")
                            parts.append(f"**{r.iloc[0]}** — MAPE=`{mape:.1f}%` {st_txt}")
                        except: pass
                else:
                    for aid, adata in arts.items():
                        hist = BE["get_hist"](adata)
                        if len(hist)>0:
                            cv = float(np.std(hist.values)/np.mean(hist.values)) if np.mean(hist.values)>0 else 0
                            parts.append(f"**{aid}** — N=`{len(hist)}` mois | Moy=`{hist.mean():,.0f}` U | CV=`{cv:.2f}`")

            elif any(w in q for w in ["saisonnalité","tendance","type","classif","saison"]):
                parts.append("**🔍 Classification des séries temporelles :**\n")
                MONTHS=["Jan","Fév","Mar","Avr","Mai","Jun","Jul","Aoû","Sep","Oct","Nov","Déc"]
                for aid, adata in arts.items():
                    hist = BE["get_hist"](adata)
                    if len(hist)<4: continue
                    vals = hist.values.astype(float)
                    z  = np.sum(vals==0)/len(vals)
                    cv = np.std(vals)/(np.mean(vals)+1e-9)
                    typ = "intermittente" if z>0.5 else ("saisonnière" if len(vals)>=24 else ("stable" if cv<0.2 else "variable"))
                    parts.append(f"**{aid}** → Type: `{typ}` | CV=`{cv:.2f}` | Moy=`{hist.mean():,.0f}` U | N=`{len(hist)}`")

            elif any(w in q for w in ["compare","les","holt","arima","méthode","modèle"]):
                parts.append("**⚖️ Comparaison des méthodes de prévision :**\n")
                for aid, adata in list(arts.items())[:2]:
                    hist = BE["get_hist"](adata)
                    if len(hist)<4: continue
                    try:
                        sel = BE["analyse_art"](aid, adata, {})
                        if not sel: continue
                        parts.append(f"**{aid}** :")
                        for _, row in sel["all_results"].iterrows():
                            best = " ✅ MEILLEUR" if row["Méthode"]==sel["best_method"] else ""
                            parts.append(f"  `{row['Méthode']:30s}` MAE=`{row['MAE']:8.0f}` MAPE=`{row['MAPE(%)']:5.1f}%`{best}")
                    except: pass

            else:
                parts.append(f"📂 Fichier `{Path(fp).name}` chargé — **{n} article(s)** | **{len(tc)} périodes**\n")
                for aid, adata in list(arts.items())[:4]:
                    hist = BE["get_hist"](adata)
                    if len(hist)>0:
                        parts.append(f"- `{aid}` : {len(hist)} mois | moy=`{hist.mean():,.0f}` U")
                parts.append("\n💡 Suggestions : *forecast 6 mois*, *audit MAPE*, *saisonnalité*, *comparer méthodes*")

        # ══ AGENT PRODUCTION ══════════════════════════════════════════════════
        elif agent == "production":
            if not fp:
                return "📁 **Importe un fichier MPS Excel** (ex: exemple_prod.xlsx) pour que je puisse calculer le MRP.", None

            data = BE["load_prod"](fp)
            arts = data["articles"]
            tc   = data["time_cols"]

            if any(w in q for w in ["mrp","complet","calcule","manquant","surcharge","capacité","charge","fill","alerte"]):
                parts.append(f"🏭 **Calcul MRP complet — {len(arts)} article(s) :**\n")
                tmp = str(Path(fp).parent / "output_prod_complet.xlsx")
                result = BE["rp"](context={"prod_file": fp, "output_path": tmp})
                excel_out = tmp

                for art_id in arts.keys():
                    n_sur = result.get("surcharges_detectees", 0)
                    n_ale = result.get("alertes_detectees", 0)
                    parts.append(f"**{art_id}** — {len(tc)} périodes analysées :")
                    if n_sur > 0:
                        periods = result.get("periodes_surcharge", [])
                        parts.append(f"  🔴 **{n_sur} surcharge(s)** : `{'`, `'.join(str(p) for p in periods[:4])}{'...' if len(periods)>4 else ''}`")
                    if n_ale > 0:
                        parts.append(f"  🟡 **{n_ale} alerte(s)** capacité")
                    if n_sur == 0 and n_ale == 0:
                        parts.append("  🟢 Capacité suffisante sur tout l'horizon")
                    parts.append(f"\n✅ Fichier Excel complet généré avec :")
                    for c in result.get("calculs_ajoutes", []):
                        parts.append(f"  · {c}")

            elif any(w in q for w in ["optimise","ajuste","plan","corrige","stock"]):
                parts.append("**🔧 Optimisation du plan de production :**\n")
                tmp = str(Path(fp).parent / "output_prod_complet.xlsx")
                result = BE["rp"](context={"prod_file": fp, "output_path": tmp})
                excel_out = tmp
                n_adj = result.get("surcharges_detectees", 0)
                parts.append(f"**{n_adj} période(s) ajustée(s)** — plan réduit aux limites de capacité.")
                parts.append("📥 Fichier Excel mis à jour avec le plan ajusté.")

            elif any(w in q for w in ["30%","20%","simulation","hausse","augmente","scénario"]):
                pct = next((int(m) for m in re.findall(r"(\d+)\s*%", q)), 30)
                parts.append(f"**📊 Simulation demande +{pct}% :**\n")
                for aid, adict in arts.items():
                    df_cap  = BE["analyse_cap"](aid, adict, tc, {"demand_boost": 1+pct/100})
                    surcharges = df_cap[df_cap["Statut"].str.contains("SURCHARGE", na=False)]
                    verdict = f"⚠️ **{len(surcharges)} surcharge(s)** → capacité insuffisante" if len(surcharges)>0 else "✅ Capacité absorbable"
                    parts.append(f"**{aid}** → +{pct}% : {verdict}")

            else:
                parts.append(f"📂 Fichier `{Path(fp).name}` chargé — **{len(arts)} article(s)** | **{len(tc)} périodes**\n")
                for aid, adict in arts.items():
                    parts.append(f"- Article `{aid}` : {len(adict)} indicateurs MPS détectés")
                parts.append("\n💡 Suggestions : *calcule le MRP complet*, *Y a-t-il des surcharges ?*, *simulation +30%*")

        # ══ AGENT MARKETING ═══════════════════════════════════════════════════
        elif agent == "marketing":
            if not fp:
                return "📁 **Importe un fichier Excel** de demande pour analyser les promotions.", None

            data = BE["load_dem"](fp)
            arts = data["articles"]
            pct  = next((int(m) for m in re.findall(r"(\d+)\s*%", q)), 20)

            if any(w in q for w in ["promo","remise","discount","réduction","simulation","uplift"]):
                parts.append(f"**📣 Simulation promotion -{pct}% — {len(arts)} article(s) :**\n")
                ctx = {"discount_pct": pct, "promo_horizon": 3, "price_elasticity": -1.5}
                tmp = str(Path(fp).parent / "output_mkt.xlsx")
                BE["rm"](context={**ctx, "demand_file": fp, "output_path": tmp})
                excel_out = tmp
                for aid, adata in arts.items():
                    try:
                        a = BE["analyse_promo"](aid, adata, ctx.copy())
                        parts.append(f"**{aid}** → Uplift max=`+{a['uplift_max_pct']:.1f}%` | Pic=`{a['demand_peak']:,.0f}` U | Base=`{a['base_demand']:,.0f}` U/mois")
                    except: pass
                parts.append("\n📥 Fichier Excel généré avec détail mois par mois.")

            elif any(w in q for w in ["saisonnalité","meilleur mois","saison","période","timing"]):
                parts.append("**📅 Meilleurs mois pour une promotion :**\n")
                MONTHS = ["Jan","Fév","Mar","Avr","Mai","Jun","Jul","Aoû","Sep","Oct","Nov","Déc"]
                for aid, adata in arts.items():
                    hist = BE["get_hist"](adata)
                    if len(hist) < 12: continue
                    vals = hist.values.astype(float)
                    monthly = {}
                    for i, v in enumerate(vals):
                        if v > 0: monthly.setdefault(i%12, []).append(v)
                    mean_all = np.mean([v for v in vals if v>0])
                    factors  = {m: round(np.mean(vs)/mean_all, 2) for m,vs in monthly.items()}
                    best  = max(factors, key=factors.get)
                    worst = min(factors, key=factors.get)
                    parts.append(f"**{aid}** → 📈 Meilleur: `{MONTHS[best]}` ({factors[best]:.2f}x) | 📉 Moins bon: `{MONTHS[worst]}` ({factors[worst]:.2f}x)")

            elif any(w in q for w in ["compare","vs","versus","scénarios","plusieurs"]):
                parts.append("**⚖️ Comparaison scénarios promotion :**\n")
                art_id   = list(arts.keys())[0]
                art_data = list(arts.values())[0]
                for pct_t in [10, 15, 20, 25, 30]:
                    ctx = {"discount_pct": pct_t, "promo_horizon": 3, "price_elasticity": -1.5}
                    try:
                        a = BE["analyse_promo"](art_id, art_data, ctx.copy())
                        parts.append(f"  **-{pct_t}%** → Uplift `+{a['uplift_max_pct']:.1f}%` | Pic `{a['demand_peak']:,.0f}` U")
                    except: pass

            else:
                parts.append(f"📂 Fichier chargé — **{len(arts)} article(s)**\n")
                parts.append("💡 Suggestions : *simuler une promo -20%*, *meilleur mois pour une promo*, *comparer scénarios*")

        # ══ AGENT FINANCE ══════════════════════════════════════════════════════
        elif agent == "finance":
            if not fp:
                return "📁 **Importe un fichier Excel** de demande pour l'analyse financière.", None

            data = BE["load_dem"](fp)
            arts = data["articles"]
            pct  = next((int(m) for m in re.findall(r"(\d+)\s*%", q)), 0)
            has_promo = any(w in q for w in ["promo","remise","discount"])
            discount  = pct if has_promo else 0

            if any(w in q for w in ["roi","rentab","p&l","marge","profit","calcule"]):
                parts.append(f"**💰 Calcul P&L & ROI{' promo -'+str(discount)+'%' if discount else ''} :**\n")
                tmp = str(Path(fp).parent / "output_fin.xlsx")
                result = BE["rf"](context={"demand_file": fp, "discount_pct": discount, "output_path": tmp})
                excel_out = tmp
                for aid, adata in arts.items():
                    hist = BE["get_hist"](adata)
                    fin  = BE["est_fin"](hist)
                    base = float(hist.mean()) if len(hist)>0 else 1000
                    ctx  = {"discount_pct": discount, "demand_forecast": [base]*3, "uplift_pcts": [20,14,8] if discount else [0,0,0]}
                    df_pl = BE["compute_pl"](aid, fin, ctx)
                    roi   = df_pl["ROI (%)"].mean()
                    mg    = df_pl["Taux marge (%)"].mean()
                    verdict = "✅ Rentable" if roi>10 else ("⚠️ ROI faible" if roi>0 else "❌ Non rentable")
                    parts.append(f"**{aid}** → ROI=`{roi:.1f}%` | Marge=`{mg:.1f}%` {verdict}")
                parts.append("\n📥 Fichier Excel P&L généré.")

            elif any(w in q for w in ["seuil","rentabilité","break","budget"]):
                parts.append("**📐 Seuil de rentabilité :**\n")
                for aid, adata in arts.items():
                    hist = BE["get_hist"](adata)
                    fin  = BE["est_fin"](hist)
                    prix  = fin["price_per_unit"]; cout = fin["cost_per_unit"]; fixed = fin["fixed_costs_month"]
                    seuil = fixed/(prix-cout) if (prix-cout)>0 else 0
                    base  = float(hist.mean()) if len(hist)>0 else 0
                    ok    = base >= seuil
                    parts.append(f"**{aid}** → Prix=`{prix:.2f}€` | Coût=`{cout:.2f}€` | Seuil=**`{seuil:,.0f}` U/mois** → {'✅ Au-dessus' if ok else '⚠️ En-dessous'} (ventes=`{base:,.0f}`)")

            elif any(w in q for w in ["alerte","surveill","déficit","risque","marge <"]):
                parts.append("**🔔 Surveillance financière :**\n")
                for aid, adata in arts.items():
                    hist = BE["get_hist"](adata)
                    fin  = BE["est_fin"](hist)
                    base = float(hist.mean()) if len(hist)>0 else 1000
                    ctx  = {"discount_pct":0,"demand_forecast":[base]*3,"uplift_pcts":[0,0,0]}
                    df_pl = BE["compute_pl"](aid, fin, ctx)
                    defs  = df_pl[df_pl["Marge promo (€)"]<0]
                    fbles = df_pl[df_pl["Taux marge (%)"]<10]
                    if len(defs)>0:  parts.append(f"  🔴 **{aid}** : {len(defs)} mois déficitaire(s)")
                    elif len(fbles)>0: parts.append(f"  🟡 **{aid}** : marge < 10% sur {len(fbles)} mois")
                    else:              parts.append(f"  🟢 **{aid}** : finances saines")

            else:
                parts.append(f"📂 Fichier chargé — **{len(arts)} article(s)**\n")
                parts.append("💡 Suggestions : *calcule le ROI*, *seuil de rentabilité*, *alerte si marge < 10%*")

        else:
            parts.append(f"Agent `{agent}` prêt. Pose-moi une question sur tes données.")

    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        parts.append(f"❌ **Erreur lors de l'analyse :**\n`{e}`\n\n```\n{tb[-500:]}\n```")

    return "\n\n".join(parts) if parts else "Analyse terminée.", excel_out


def orchestrate(question: str) -> str:
    """Orchestrateur : décide les agents, les appelle, synthétise."""
    q = question.lower()
    add_trace("orchestrateur", f"Demande : « {question[:50]}{'...' if len(question)>50 else ''} »")

    if any(w in q for w in ["promo","remise","discount","simulation"]):
        agents = ["marketing","demande","production","finance"]
        add_trace("orchestrateur", "Scénario PROMOTION → 4 agents", "decision")
    elif any(w in q for w in ["capacité","surcharge","production","mrp","plan"]):
        agents = ["production"]
        add_trace("orchestrateur", "Scénario PRODUCTION → Agent Production", "decision")
    elif any(w in q for w in ["forecast","prévision","demande","mape","méthode"]):
        agents = ["demande"]
        add_trace("orchestrateur", "Scénario DEMANDE → Agent Demande", "decision")
    elif any(w in q for w in ["roi","rentab","marge","finance","coût","budget"]):
        agents = ["finance"]
        add_trace("orchestrateur", "Scénario FINANCE → Agent Finance", "decision")
    else:
        agents = ["demande","production"]
        add_trace("orchestrateur", "Analyse générale → Demande + Production", "decision")

    lines = [f"J'ai mobilisé **{len(agents)} agent(s)** : {', '.join([AGENT_CFG[a]['icon']+' '+a.capitalize() for a in agents])}\n"]
    for ag in agents:
        add_trace(ag, "Analyse en cours...", "start")
        resp, excel = think(ag, question)
        first_line  = resp.split("\n")[0].replace("**","").replace("`","")[:80]
        lines.append(f"{AGENT_CFG[ag]['icon']} **{ag.capitalize()}** → {first_line}")
        # Stocke aussi dans le chat de l'agent
        add_msg(ag, "user", f"[Via Orchestrateur] {question}")
        add_msg(ag, "agent", resp, excel=excel)
        add_trace(ag, first_line[:80], "success")

    add_trace("orchestrateur", "Synthèse terminée", "synthesis")
    # Décision simple
    if "finance" in agents:
        lines.append("\n**Décision :** Consulte l'onglet Finance pour le verdict ROI complet.")
    else:
        lines.append("\n**Résultats disponibles** dans chaque onglet agent.")

    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# COMPOSANTS UI
# ══════════════════════════════════════════════════════════════════════════════
def render_chat(agent):
    msgs = st.session_state.chats.get(agent, [])
    cfg  = AGENT_CFG[agent]
    if not msgs:
        st.markdown('<div style="color:var(--muted);font-size:.82rem;padding:.5rem;text-align:center">💬 Pose une question ou utilise une suggestion ci-dessous</div>', unsafe_allow_html=True)
        return
    for m in msgs:
        if m["role"] == "user":
            st.markdown(
                f'<div class="msg-user"><div class="bubble bubble-u">{fmt_html(m["content"])}</div>'
                f'<div class="avatar" style="background:var(--bg3);color:var(--muted)">U</div></div>',
                unsafe_allow_html=True)
        else:
            st.markdown(
                f'<div class="msg-agent"><div class="avatar" style="{cfg["av"]}">{cfg["icon"]}</div>'
                f'<div class="bubble bubble-a">{fmt_html(m["content"])}</div></div>',
                unsafe_allow_html=True)
            if m.get("excel") and Path(m["excel"]).exists():
                with open(m["excel"],"rb") as f:
                    st.download_button(
                        f"⬇ Télécharger Excel — {Path(m['excel']).name}",
                        data=f.read(),
                        file_name=Path(m["excel"]).name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{agent}_{m['ts']}",
                    )


def render_agent_tab(agent):
    """Rendu complet d'un onglet agent."""
    cfg = AGENT_CFG[agent]
    L, R = st.columns([1, 2], gap="small")

    # ── PANNEAU GAUCHE ──────────────────────────────────────────────────────
    with L:
        # Fichier
        st.markdown(f'<div class="sec-label">📁 FICHIER DE DONNÉES</div>', unsafe_allow_html=True)
        fi = st.session_state.files.get(agent)
        if fi:
            st.markdown(f'<div class="fbadge">✓ {fi["name"]}</div>', unsafe_allow_html=True)
            with st.expander("Aperçu", expanded=False):
                if not fi["df"].empty:
                    st.dataframe(fi["df"], use_container_width=True, hide_index=True, height=120)
        up = st.file_uploader(
            "Importer .xlsx", type=["xlsx"],
            key=f"up_{agent}", label_visibility="collapsed"
        )
        if up is not None:
            path = save_upload(agent, up)
            if not any(m["content"].startswith("✅ Fichier") and up.name in m["content"]
                       for m in st.session_state.chats[agent]):
                add_msg(agent, "agent", f"✅ Fichier **{up.name}** importé ({up.size//1024} KB). Prêt à analyser.")

        # Capacités
        st.markdown(f'<div class="sec-label" style="margin-top:.75rem">⚙ CAPACITÉS</div>', unsafe_allow_html=True)
        for cap in CAPS[agent]:
            st.markdown(f'<div class="cap-item"><span class="cap-dot">✦</span>{cap}</div>', unsafe_allow_html=True)

        # Erreurs de modules
        if BE.get("_errors"):
            with st.expander("⚠ Avertissements modules", expanded=False):
                for e in BE["_errors"]:
                    st.markdown(f'<div class="warn-box">{e}</div>', unsafe_allow_html=True)

    # ── PANNEAU DROIT ───────────────────────────────────────────────────────
    with R:
        st.markdown(f'<div class="sec-label">{cfg["icon"]} AGENT {agent.upper()} — CHAT</div>', unsafe_allow_html=True)

        # Zone chat
        with st.container(height=360):
            render_chat(agent)

        # Suggestions rapides
        st.markdown('<div style="padding:.3rem 0 .2rem;font-size:.67rem;color:var(--muted);font-family:Space Mono,monospace;letter-spacing:.08em">SUGGESTIONS RAPIDES</div>', unsafe_allow_html=True)
        q_cols = st.columns(2)
        prompts = QUICK[agent]
        for i, prompt in enumerate(prompts):
            with q_cols[i % 2]:
                if st.button(prompt, key=f"q_{agent}_{i}", use_container_width=True):
                    add_msg(agent, "user", prompt)
                    with st.spinner(f"Agent {agent} analyse..."):
                        resp, excel = think(agent, prompt)
                    add_msg(agent, "agent", resp, excel=excel)
                    st.rerun()

        # Saisie libre
        c1, c2 = st.columns([5, 1])
        with c1:
            user_input = st.text_input(
                "", key=f"input_{agent}", label_visibility="collapsed",
                placeholder=f"Pose une question à l'Agent {agent.capitalize()}..."
            )
        with c2:
            st.markdown('<div class="btn-send">', unsafe_allow_html=True)
            send = st.button("→", key=f"send_{agent}")
            st.markdown('</div>', unsafe_allow_html=True)

        if send and user_input and user_input.strip():
            add_msg(agent, "user", user_input.strip())
            with st.spinner(f"Agent {agent} réfléchit..."):
                resp, excel = think(agent, user_input.strip())
            add_msg(agent, "agent", resp, excel=excel)
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TOPBAR
# ══════════════════════════════════════════════════════════════════════════════
n_files = len(st.session_state.files)
file_status = f"✓ {n_files} fichier(s)" if n_files else "○ Aucun fichier"
file_color  = "#10b981" if n_files else "#6b6b80"

st.markdown(f"""
<div class="topbar">
  <div class="topbar-logo">⬡ S&OP AGENTIQUE</div>
  <div class="topbar-sep"></div>
  <div class="topbar-sub">5 AGENTS AUTONOMES · MRP COMPLET · DÉCISION EN TEMPS RÉEL</div>
  <div style="margin-left:auto;font-size:.72rem;font-family:'Space Mono',monospace;color:{file_color}">{file_status}</div>
</div>
""", unsafe_allow_html=True)

# Affiche les erreurs critiques si modules manquants
if BE.get("_errors") and len(BE["_errors"]) > 2:
    st.warning(f"⚠️ {len(BE['_errors'])} module(s) non chargé(s) — vérifie que tous les fichiers .py sont dans le même dossier.\n" + "\n".join(BE["_errors"][:3]))

# ══════════════════════════════════════════════════════════════════════════════
# ONGLETS
# ══════════════════════════════════════════════════════════════════════════════
tabs = st.tabs(["⬡ Orchestrateur", "📈 Demande", "🏭 Production", "📣 Marketing", "💰 Finance"])

# ── TAB ORCHESTRATEUR ────────────────────────────────────────────────────────
with tabs[0]:
    L, R = st.columns([1, 2], gap="small")
    with L:
        st.markdown('<div class="sec-label">📁 FICHIERS CHARGÉS</div>', unsafe_allow_html=True)
        if st.session_state.files:
            for ag, fi in st.session_state.files.items():
                st.markdown(f'<div class="fbadge">{AGENT_CFG.get(ag,{}).get("icon","📄")} {fi["name"]}</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div style="color:var(--muted);font-size:.8rem">Importe des fichiers dans les onglets agents</div>', unsafe_allow_html=True)

        up_orch = st.file_uploader("Importer (auto-détecté)", type=["xlsx"], key="up_orch", label_visibility="collapsed")
        if up_orch is not None:
            path = save_upload("orchestrateur", up_orch)
            if "auto_load" in BE:
                try:
                    d = BE["auto_load"](path)
                    if d["type"] == "production":
                        st.session_state.files["production"] = st.session_state.files["orchestrateur"]
                    else:
                        for ag in ["demande","marketing","finance"]:
                            st.session_state.files[ag] = st.session_state.files["orchestrateur"]
                except: pass

        st.markdown('<div class="sec-label" style="margin-top:.75rem">⬡ TRACE DES AGENTS</div>', unsafe_allow_html=True)
        STATUS_C = {"info":"#6b6b80","decision":"#f59e0b","start":"#3b82f6","success":"#10b981","error":"#ef4444","synthesis":"#8b5cf6"}
        with st.container(height=280):
            traces = st.session_state.traces[-30:]
            if not traces:
                st.markdown('<div style="color:var(--muted);font-size:.8rem;padding:.5rem">Lance un scénario pour voir la trace...</div>', unsafe_allow_html=True)
            for tr in traces:
                c = STATUS_C.get(tr["status"],"#6b6b80")
                ac = AGENT_CFG.get(tr["agent"], AGENT_CFG["orchestrateur"])
                st.markdown(
                    f'<div class="trace-item">'
                    f'<div class="trace-dot" style="background:{c}"></div>'
                    f'<div class="trace-agent" style="color:{ac["color"]}">{ac["icon"]} {tr["agent"][:6].upper()}</div>'
                    f'<div class="trace-msg">{tr["msg"]}</div>'
                    f'<div class="trace-ts">{tr["ts"]}</div>'
                    f'</div>', unsafe_allow_html=True)

        if st.button("🗑 Effacer trace", key="clear_traces"):
            st.session_state.traces = []
            st.rerun()

    with R:
        st.markdown('<div class="sec-label">⬡ ORCHESTRATEUR — COORDINATION</div>', unsafe_allow_html=True)
        with st.container(height=360):
            render_chat("orchestrateur")

        st.markdown('<div style="padding:.3rem 0 .2rem;font-size:.67rem;color:var(--muted);font-family:Space Mono,monospace;letter-spacing:.08em">SUGGESTIONS RAPIDES</div>', unsafe_allow_html=True)
        q_cols = st.columns(2)
        for i, prompt in enumerate(QUICK["orchestrateur"]):
            with q_cols[i%2]:
                if st.button(prompt, key=f"q_orch_{i}", use_container_width=True):
                    add_msg("orchestrateur", "user", prompt)
                    with st.spinner("Orchestrateur coordonne les agents..."):
                        resp = orchestrate(prompt)
                    add_msg("orchestrateur", "agent", resp)
                    st.rerun()

        c1, c2 = st.columns([5, 1])
        with c1:
            orch_input = st.text_input("", key="orch_input", label_visibility="collapsed",
                                       placeholder="Donne un ordre à l'orchestrateur...")
        with c2:
            st.markdown('<div class="btn-send">', unsafe_allow_html=True)
            orch_send = st.button("→", key="orch_send")
            st.markdown('</div>', unsafe_allow_html=True)

        if orch_send and orch_input and orch_input.strip():
            add_msg("orchestrateur", "user", orch_input.strip())
            with st.spinner("Orchestrateur coordonne les agents..."):
                resp = orchestrate(orch_input.strip())
            add_msg("orchestrateur", "agent", resp)
            st.rerun()

# ── TABS AGENTS ───────────────────────────────────────────────────────────────
with tabs[1]: render_agent_tab("demande")
with tabs[2]: render_agent_tab("production")
with tabs[3]: render_agent_tab("marketing")
with tabs[4]: render_agent_tab("finance")

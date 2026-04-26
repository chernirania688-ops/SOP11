"""
app.py — S&OP Agentique · Interface complète multi-onglets
==========================================================
5 onglets : Orchestrateur · Demande · Production · Marketing · Finance
Chaque onglet : import fichier + chat IA + actions directes Excel
"""

import sys, os, json, re, warnings, tempfile
from pathlib import Path
from datetime import datetime

import streamlit as st
import pandas as pd
import numpy as np

warnings.filterwarnings("ignore")
sys.path.append(str(Path(__file__).parent))

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="S&OP Agentique", page_icon="⬡", layout="wide", initial_sidebar_state="collapsed")

# ── Design system ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Syne:wght@400;600;700;800&display=swap');
:root{
  --bg:#0a0a0b;--bg1:#111113;--bg2:#18181c;--bg3:#222228;
  --border:#2a2a32;--border2:#3a3a45;
  --text:#e8e8f0;--muted:#6b6b80;
  --amber:#f59e0b;--amber2:#fbbf24;
  --green:#10b981;--red:#ef4444;--blue:#3b82f6;--purple:#8b5cf6;
  --pink:#ec4899;--orange:#f97316;
}
*{box-sizing:border-box}
html,body,[class*="css"]{font-family:'Syne',sans-serif;background:var(--bg)!important;color:var(--text)}
#MainMenu,footer,header{visibility:hidden}
.block-container{padding:0!important;max-width:100%!important}
section[data-testid="stSidebar"]{display:none!important}

/* Topbar */
.topbar{background:var(--bg1);border-bottom:1px solid var(--border);padding:.7rem 1.5rem;display:flex;align-items:center;gap:1rem}
.topbar-logo{font-family:'Space Mono',monospace;font-size:.95rem;font-weight:700;color:var(--amber);letter-spacing:.06em}
.topbar-sep{width:1px;height:18px;background:var(--border2)}
.topbar-sub{font-size:.75rem;color:var(--muted);font-family:'Space Mono',monospace}

/* Streamlit tab overrides */
.stTabs [data-baseweb="tab-list"]{background:var(--bg1)!important;border-bottom:1px solid var(--border)!important;gap:0!important}
.stTabs [data-baseweb="tab"]{background:transparent!important;color:var(--muted)!important;font-family:'Syne',sans-serif!important;font-size:.82rem!important;font-weight:600!important;padding:.7rem 1.3rem!important;border-bottom:2px solid transparent!important;border-radius:0!important}
.stTabs [aria-selected="true"]{color:var(--amber)!important;border-bottom-color:var(--amber)!important;background:var(--bg2)!important}
.stTabs [data-baseweb="tab-panel"]{background:var(--bg)!important;padding:0!important}

/* Inputs */
.stTextInput>div>div>input{background:var(--bg2)!important;border:1px solid var(--border2)!important;color:var(--text)!important;border-radius:6px!important;font-family:'Syne',sans-serif!important;font-size:.88rem!important}
.stTextInput>div>div>input:focus{border-color:var(--amber)!important;box-shadow:0 0 0 1px var(--amber)!important}
.stButton>button{background:var(--amber)!important;color:#0a0a0b!important;border:none!important;border-radius:6px!important;font-weight:700!important;font-family:'Syne',sans-serif!important;font-size:.82rem!important;padding:.4rem 1rem!important}
.stButton>button:hover{opacity:.85!important}
.stFileUploader>div{background:var(--bg2)!important;border:1px solid var(--border2)!important;border-radius:8px!important}
.stFileUploader label{color:var(--muted)!important;font-size:.8rem!important}

/* Section headers */
.sec-hdr{padding:.55rem 1rem;font-size:.67rem;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:var(--muted);background:var(--bg1);border-bottom:1px solid var(--border);font-family:'Space Mono',monospace}

/* File badge */
.fbadge{margin:.5rem 0;background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:.45rem .75rem;font-size:.76rem;font-family:'Space Mono',monospace;color:var(--green);display:flex;align-items:center;gap:.5rem}

/* Chat messages */
.chat-wrap{display:flex;flex-direction:column;gap:.6rem;padding:.75rem}
.msg{display:flex;gap:.65rem;max-width:90%}
.msg.u{align-self:flex-end;flex-direction:row-reverse}
.msg.a{align-self:flex-start}
.avatar{width:26px;height:26px;border-radius:5px;display:flex;align-items:center;justify-content:center;font-size:.72rem;font-weight:700;flex-shrink:0;font-family:'Space Mono',monospace}
.bubble{padding:.6rem .85rem;border-radius:8px;font-size:.83rem;line-height:1.6;border:1px solid var(--border)}
.bubble.u{background:var(--bg2);color:var(--text);border-color:var(--border2)}
.bubble.a{background:var(--bg1);color:var(--text)}
.bubble.a strong{color:var(--amber)}
.bubble.a code{background:var(--bg3);padding:.1rem .3rem;border-radius:3px;font-family:'Space Mono',monospace;font-size:.78em;color:var(--amber2)}

/* Quick prompts */
.qp-wrap{padding:.5rem .75rem;border-top:1px solid var(--border);display:flex;flex-wrap:wrap;gap:.35rem}
.qp-lbl{font-size:.67rem;color:var(--muted);font-family:'Space Mono',monospace;letter-spacing:.08em;padding:.4rem 0 .2rem;width:100%}

/* Trace log */
.trace-row{display:flex;align-items:flex-start;gap:.65rem;padding:.5rem 1rem;border-bottom:1px solid var(--border);font-size:.78rem}
.trace-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0;margin-top:.3rem}
.trace-agent{font-weight:700;font-family:'Space Mono',monospace;min-width:85px;font-size:.75rem}
.trace-msg{color:var(--muted);flex:1;line-height:1.45}
.trace-time{color:var(--border2);font-family:'Space Mono',monospace;font-size:.68rem;flex-shrink:0}

/* Capabilities list */
.cap-item{padding:.35rem 1rem;font-size:.78rem;color:var(--muted);border-bottom:1px solid var(--border);display:flex;align-items:center;gap:.5rem}
.cap-item span{color:var(--amber)}

/* Scrollbars */
::-webkit-scrollbar{width:4px;height:4px}
::-webkit-scrollbar-track{background:var(--bg1)}
::-webkit-scrollbar-thumb{background:var(--border2);border-radius:4px}

/* Dataframe */
.stDataFrame{background:var(--bg2)!important}
.element-container .stDataFrame{font-size:.78rem}
</style>
""", unsafe_allow_html=True)

# ── State ─────────────────────────────────────────────────────────────────────
AGENT_KEYS = ["orchestrateur", "demande", "production", "marketing", "finance"]
for k in ["files", "chats", "traces"]:
    if k not in st.session_state:
        st.session_state[k] = {} if k != "traces" else []
for a in AGENT_KEYS:
    if a not in st.session_state.chats:
        st.session_state.chats[a] = []

AGENTS = {
    "orchestrateur": {"icon":"⬡", "label":"Orchestrateur","color":"#f59e0b","av_css":"background:#2d1f00;color:#f59e0b"},
    "demande":       {"icon":"📈","label":"Demande",      "color":"#10b981","av_css":"background:#001a0f;color:#10b981"},
    "production":    {"icon":"🏭","label":"Production",   "color":"#f97316","av_css":"background:#1f0d00;color:#f97316"},
    "marketing":     {"icon":"📣","label":"Marketing",    "color":"#ec4899","av_css":"background:#1f0011;color:#ec4899"},
    "finance":       {"icon":"💰","label":"Finance",      "color":"#3b82f6","av_css":"background:#00103f;color:#3b82f6"},
}

QUICK = {
    "orchestrateur":["Lance une analyse S&OP complète","Simule une promotion -20%","Vérifie la capacité de production","La promo est-elle rentable ?","Analyse globale et recommandation"],
    "demande":      ["Calcule le forecast 6 mois","Quelle est la meilleure méthode ?","Analyse la saisonnalité","Détecte les anomalies","Compare LES vs Holt-Winters"],
    "production":   ["Analyse la capacité Fill-L1","Y a-t-il des surcharges ?","Optimise le plan de production","Simulation demande +30%","Quel est le stock de sécurité optimal ?"],
    "marketing":    ["Simule une promo -20%","Quel est le meilleur mois ?","Compare promo -10% vs -30%","Calcule l'uplift attendu","Analyse la saisonnalité des ventes"],
    "finance":      ["Calcule le ROI promotion","Le budget est-il respecté ?","Quel est le seuil de rentabilité ?","Alerte si marge < 10%","Simule le P&L ventes +20%"],
}

# ── Backends ──────────────────────────────────────────────────────────────────
@st.cache_resource
def load_be():
    """Charge tous les modules agents de manière sécurisée."""
    import importlib
    base_dir = str(Path(__file__).parent)
    if base_dir not in sys.path:
        sys.path.insert(0, base_dir)

    result = {}
    modules_attrs = {
        "data_loader":     [("auto_load","auto_load"),("load_demand_file","load_dem"),("load_production_file","load_prod"),("get_clean_history","get_hist")],
        "agent_demande":   [("run","rd"),("analyse_article","analyse_art")],
        "agent_production":[("run","rp"),("analyse_capacity","analyse_cap"),("generate_adjusted_plan","gen_adj")],
        "agent_marketing": [("run","rm"),("analyse_article_promo","analyse_promo")],
        "agent_finance":   [("run","rf"),("compute_pl_scenario","compute_pl"),("estimate_financials_from_history","est_fin")],
        "orchestrateur":   [("run","ro"),("detect_files","detect"),("synthesize","synth"),("SCENARIO_AGENTS","sc_agents")],
    }
    warnings_list = []
    for mod_name, attrs in modules_attrs.items():
        try:
            mod = importlib.import_module(mod_name)
            for orig, alias in attrs:
                result[alias] = getattr(mod, orig)
        except Exception as e:
            warnings_list.append(f"{mod_name}: {e}")
    if warnings_list:
        result["WARNINGS"] = warnings_list
    return result

BE = load_be()

# ── Helpers ───────────────────────────────────────────────────────────────────
def save_upload(agent, f):
    tmp = Path(tempfile.mkdtemp()); dest = tmp / f.name
    dest.write_bytes(f.read())
    try: df_prev = pd.read_excel(str(dest), nrows=6)
    except: df_prev = pd.DataFrame()
    st.session_state.files[agent] = {"name": f.name, "path": str(dest), "df": df_prev}
    return str(dest)

def gpath(agent):
    return st.session_state.files.get(agent, {}).get("path")

def all_paths():
    return [v["path"] for v in st.session_state.files.values() if "path" in v]

def add_msg(agent, role, content, excel=None):
    st.session_state.chats[agent].append({"role":role,"content":content,"ts":datetime.now().isoformat(),"excel":excel})

def trace(agent, msg, status="info"):
    ts = datetime.now().strftime("%H:%M:%S")
    st.session_state.traces.append({"agent":agent,"msg":msg,"status":status,"ts":ts})

def auto_file(agent):
    """Auto-détecte un fichier compatible pour cet agent parmi ceux uploadés."""
    p = gpath(agent)
    if p: return p
    if "ERR" in BE: return None
    for path in all_paths():
        try:
            d = BE["auto_load"](path)
            if agent == "production" and d["type"] == "production": return path
            if agent in ("demande","marketing","finance") and d["type"] == "demand": return path
        except: pass
    return all_paths()[0] if all_paths() else None

# ── Agent brain ───────────────────────────────────────────────────────────────
def think(agent, question, file_path):
    q = question.lower()
    parts = []; excel_out = None
    pct = next((int(m) for m in re.findall(r"(\d+)\s*%", q)), 20)

    try:
        # ── DEMANDE ──────────────────────────────────────────────────────────
        if agent == "demande":
            if not file_path:
                return "📁 Importe un fichier de demande pour que je puisse agir.", None
            data = BE["load_dem"](file_path)
            arts = data["articles"]

            if any(w in q for w in ["forecast","prévision","prévoir","6 mois","futur"]):
                parts.append(f"✅ **Calcul du forecast** — {len(arts)} article(s) :\n")
                for aid, adata in arts.items():
                    hist = BE["get_hist"](adata)
                    if len(hist) < 4: continue
                    sel = BE["analyse_art"](aid, adata, {})
                    if not sel: continue
                    kpi_row = sel["all_results"][sel["all_results"]["Méthode"]==sel["best_method"]].iloc[0]
                    fc_total = round(float(np.sum(sel["forecast"])),0)
                    parts.append(f"**{aid}** → `{sel['best_method']}` | MAPE=`{kpi_row['MAPE(%)']:.1f}%` | Forecast 6M = **{fc_total:,.0f} U**")
                tmp = str(Path(file_path).parent/"output_dem.xlsx")
                BE["rd"](context={"demand_file":file_path,"output_path":tmp})
                excel_out = tmp

            elif any(w in q for w in ["mape","kpi","qualité","erreur","anomalie","précision"]):
                parts.append("**Audit qualité des prévisions :**\n")
                kdf = data.get("kpis")
                if kdf is not None:
                    for _, r in kdf.iterrows():
                        try:
                            mape = float(str(r.iloc[2]).replace(" ","").replace(",","."))
                            st = "🔴 Mauvais (>50%)" if mape>50 else "🟡 Moyen" if mape>25 else "🟢 Bon"
                            parts.append(f"**{r.iloc[0]}** — MAPE=`{mape:.1f}%` {st}")
                        except: pass
                else:
                    for aid, adata in arts.items():
                        hist = BE["get_hist"](adata)
                        if len(hist)>0:
                            cv = float(np.std(hist.values)/np.mean(hist.values)) if np.mean(hist.values)>0 else 0
                            parts.append(f"**{aid}** — CV=`{cv:.2f}` | Moy=`{hist.mean():,.0f}` U | N=`{len(hist)}` mois")

            elif any(w in q for w in ["saisonnalité","tendance","type","classif","pattern"]):
                parts.append("**Classification des séries temporelles :**\n")
                for aid, adata in arts.items():
                    hist = BE["get_hist"](adata)
                    if len(hist)<4: continue
                    vals = hist.values.astype(float)
                    z = np.sum(vals==0)/len(vals)
                    cv = np.std(vals)/(np.mean(vals)+1e-9)
                    typ = "intermittente" if z>0.5 else "saisonnière" if len(vals)>=24 else "stable" if cv<0.2 else "variable"
                    parts.append(f"**{aid}** → Type: `{typ}` | CV=`{cv:.2f}` | N=`{len(hist)}` mois | Moy=`{hist.mean():,.0f}` U")

            elif any(w in q for w in ["compare","les","holt","hw","arima","ma","méthode"]):
                parts.append("**Comparaison des méthodes de prévision :**\n")
                for aid, adata in list(arts.items())[:2]:
                    hist = BE["get_hist"](adata)
                    if len(hist)<4: continue
                    sel = BE["analyse_art"](aid, adata, {})
                    if not sel: continue
                    parts.append(f"**{aid}** :")
                    for _, row in sel["all_results"].iterrows():
                        best_marker = " ← MEILLEUR" if row["Méthode"]==sel["best_method"] else ""
                        parts.append(f"  `{row['Méthode']:30s}` MAE=`{row['MAE']:7.0f}` MAPE=`{row['MAPE(%)']:5.1f}%` Score=`{row['Score']:.4f}`{best_marker}")

            else:
                parts.append(f"Fichier `{Path(file_path).name}` — **{len(arts)} article(s)** | **{len(data['time_cols'])} périodes**")
                for aid, adata in arts.items():
                    hist = BE["get_hist"](adata)
                    if len(hist)>0:
                        parts.append(f"- `{aid}` : {len(hist)} mois | moy = `{hist.mean():,.0f}` U")
                parts.append("\nSuggestions : forecast, qualité MAPE, saisonnalité, comparer les méthodes.")

        # ── PRODUCTION ───────────────────────────────────────────────────────
        elif agent == "production":
            if not file_path:
                return "📁 Importe un fichier MPS pour que je puisse agir.", None
            data = BE["load_prod"](file_path)
            arts = data["articles"]; tc = data["time_cols"]

            if any(w in q for w in ["capacité","surcharge","charge","ressource","fill"]):
                parts.append("**Analyse capacité & surcharges :**\n")
                for aid, adict in arts.items():
                    dfc = BE["analyse_cap"](aid, adict, tc, {})
                    dfa = BE["gen_adj"](dfc, adict)
                    sur = dfa[dfa["Statut"].str.contains("SURCHARGE", na=False)]
                    ale = dfa[dfa["Statut"].str.contains("ALERTE", na=False)]
                    parts.append(f"**Article {aid}** — {len(tc)} périodes :")
                    if len(sur)>0: parts.append(f"  🔴 **{len(sur)} surcharge(s)** : `{'`, `'.join(sur['Période'].head(4).tolist())}`")
                    if len(ale)>0: parts.append(f"  🟡 **{len(ale)} alerte(s)** de capacité")
                    if len(sur)==0 and len(ale)==0: parts.append("  🟢 Capacité suffisante sur tout l'horizon")
                tmp = str(Path(file_path).parent/"output_prod.xlsx")
                BE["rp"](context={"prod_file":file_path,"output_path":tmp})
                excel_out = tmp

            elif any(w in q for w in ["optimise","ajuste","plan","corrige"]):
                parts.append("**Optimisation du plan de production :**\n")
                for aid, adict in arts.items():
                    dfc = BE["analyse_cap"](aid, adict, tc, {})
                    dfa = BE["gen_adj"](dfc, adict)
                    if "Action" in dfa.columns:
                        adj = dfa[dfa["Action"].str.contains("Réduit", na=False)]
                        parts.append(f"**{aid}** — {len(adj)} ajustement(s) effectué(s) :")
                        for _, r in adj.head(4).iterrows():
                            parts.append(f"  ↳ `{r['Période']}` : {r['Action']}")
                        if len(adj)==0: parts.append(f"  ✅ Aucun ajustement nécessaire")
                tmp = str(Path(file_path).parent/"output_prod.xlsx")
                BE["rp"](context={"prod_file":file_path,"output_path":tmp})
                excel_out = tmp

            elif any(w in q for w in ["30%","20%","hausse","augmente","scénario","simulation"]):
                parts.append(f"**Simulation demande +{pct}% :**\n")
                for aid, adict in arts.items():
                    dfc = BE["analyse_cap"](aid, adict, tc, {})
                    sur_avant = len(dfc[dfc["Statut"].str.contains("SURCHARGE", na=False)])
                    # Simulation boost
                    adict_boost = {}
                    for k, s in adict.items():
                        if "gross" in k.lower():
                            adict_boost[k] = s * (1+pct/100)
                        else:
                            adict_boost[k] = s
                    dfc2 = BE["analyse_cap"](aid, adict_boost, tc, {})
                    sur_apres = len(dfc2[dfc2["Statut"].str.contains("SURCHARGE", na=False)])
                    verdict = f"⚠️ {sur_apres} surcharge(s) → action requise" if sur_apres>0 else "✅ Capacité absorbable"
                    parts.append(f"**{aid}** : avant={sur_avant} surcharge(s) → après +{pct}%: **{sur_apres} surcharge(s)** — {verdict}")

            else:
                parts.append(f"Fichier `{Path(file_path).name}` — **{len(arts)} article(s)** | **{len(tc)} périodes**")
                for aid, adict in arts.items():
                    parts.append(f"- `{aid}` : {len(list(adict.keys()))} indicateurs MPS")
                parts.append("\nSuggestions : capacité, surcharges, optimiser le plan, simulation +30%.")

        # ── MARKETING ────────────────────────────────────────────────────────
        elif agent == "marketing":
            if not file_path:
                return "📁 Importe un fichier de données pour que je puisse agir.", None
            data = BE["load_dem"](file_path); arts = data["articles"]

            if any(w in q for w in ["promo","remise","discount","réduction","simulation","-20","-10","-30"]):
                parts.append(f"**Simulation promotion -{pct}% :**\n")
                ctx = {"discount_pct":pct,"promo_horizon":3,"price_elasticity":-1.5}
                for aid, adata in arts.items():
                    a = BE["analyse_promo"](aid, adata, ctx.copy())
                    parts.append(f"**{aid}** → Uplift max=`+{a['uplift_max_pct']:.1f}%` | Pic=`{a['demand_peak']:,.0f} U` | Base=`{a['base_demand']:,.0f} U/mois`")
                tmp = str(Path(file_path).parent/"output_mkt.xlsx")
                BE["rm"](context={**ctx,"demand_file":file_path,"output_path":tmp})
                excel_out = tmp

            elif any(w in q for w in ["saisonnalité","meilleur mois","saison","période"]):
                parts.append("**Analyse saisonnalité — meilleurs mois pour une promotion :**\n")
                MONTHS=["Jan","Fév","Mar","Avr","Mai","Jun","Jul","Aoû","Sep","Oct","Nov","Déc"]
                for aid, adata in arts.items():
                    hist = BE["get_hist"](adata)
                    if len(hist)<12: continue
                    vals = hist.values.astype(float)
                    monthly = {}
                    for i,v in enumerate(vals):
                        if v>0: monthly.setdefault(i%12,[]).append(v)
                    mean_all = np.mean([v for v in vals if v>0])
                    factors = {m:round(np.mean(vs)/mean_all,2) for m,vs in monthly.items()}
                    best=max(factors,key=factors.get); worst=min(factors,key=factors.get)
                    parts.append(f"**{aid}** → Meilleur: `{MONTHS[best]}` ({factors[best]:.2f}x) | Moins bon: `{MONTHS[worst]}` ({factors[worst]:.2f}x)")

            elif any(w in q for w in ["compare","vs","versus","scénarios","plusieurs"]):
                parts.append("**Comparaison scénarios promotion :**\n")
                art_id = list(arts.keys())[0]; art_data = list(arts.values())[0]
                for pct_t in [10,15,20,25,30]:
                    ctx = {"discount_pct":pct_t,"promo_horizon":3,"price_elasticity":-1.5}
                    a = BE["analyse_promo"](art_id, art_data, ctx.copy())
                    parts.append(f"  **-{pct_t}%** → Uplift `+{a['uplift_max_pct']:.1f}%` | Pic `{a['demand_peak']:,.0f} U`")

            else:
                parts.append(f"Fichier `{Path(file_path).name}` — **{len(arts)} article(s)**")
                parts.append("Suggestions : simuler une promo, analyser la saisonnalité, comparer des scénarios.")

        # ── FINANCE ──────────────────────────────────────────────────────────
        elif agent == "finance":
            if not file_path:
                return "📁 Importe un fichier de données pour que je puisse agir.", None
            data = BE["load_dem"](file_path); arts = data["articles"]

            if any(w in q for w in ["roi","rentab","p&l","marge","profit","ventes +","ventes+"]):
                discount = pct if any(w in q for w in ["promo","remise","discount"]) else 0
                parts.append(f"**Calcul P&L & ROI{' promo -'+str(discount)+'%' if discount else ''} :**\n")
                for aid, adata in arts.items():
                    hist = BE["get_hist"](adata)
                    fin = BE["est_fin"](hist)
                    base = float(hist.mean()) if len(hist)>0 else 1000
                    ctx = {"discount_pct":discount,"demand_forecast":[base]*3,
                           "uplift_pcts":[20,14,8] if discount else [0,0,0]}
                    df_pl = BE["compute_pl"](aid, fin, ctx)
                    roi = df_pl["ROI (%)"].mean(); mg = df_pl["Taux marge (%)"].mean()
                    verdict = "✅ Rentable" if roi>10 else "⚠️ ROI faible" if roi>0 else "❌ Non rentable"
                    parts.append(f"**{aid}** → ROI moy=`{roi:.1f}%` | Marge=`{mg:.1f}%` {verdict}")
                tmp = str(Path(file_path).parent/"output_fin.xlsx")
                BE["rf"](context={"demand_file":file_path,"discount_pct":discount,"output_path":tmp})
                excel_out = tmp

            elif any(w in q for w in ["seuil","rentabilité","break even","budget"]):
                parts.append("**Analyse seuil de rentabilité :**\n")
                for aid, adata in arts.items():
                    hist = BE["get_hist"](adata)
                    fin = BE["est_fin"](hist)
                    prix=fin["price_per_unit"]; cout=fin["cost_per_unit"]; fixed=fin["fixed_costs_month"]
                    seuil = fixed/(prix-cout) if (prix-cout)>0 else 0
                    base = float(hist.mean()) if len(hist)>0 else 0
                    ok = base>=seuil
                    parts.append(f"**{aid}** → Prix=`{prix:.2f}€` | Coût=`{cout:.2f}€` | Charges fixes=`{fixed:,.0f}€/mois`")
                    parts.append(f"  Seuil de rentabilité = **`{seuil:,.0f} U/mois`** | Ventes moy=`{base:,.0f}` → {'✅ AU-DESSUS' if ok else '⚠️ EN-DESSOUS'}")

            elif any(w in q for w in ["alerte","surveill","déficit","risque"]):
                parts.append("**Surveillance financière automatique :**\n")
                for aid, adata in arts.items():
                    hist = BE["get_hist"](adata)
                    fin = BE["est_fin"](hist)
                    base = float(hist.mean()) if len(hist)>0 else 1000
                    ctx = {"discount_pct":0,"demand_forecast":[base]*3,"uplift_pcts":[0,0,0]}
                    df_pl = BE["compute_pl"](aid, fin, ctx)
                    deficits = df_pl[df_pl["Marge promo (€)"]<0]
                    faibles = df_pl[df_pl["Taux marge (%)"]<10]
                    if len(deficits)>0: parts.append(f"  🔴 **{aid}** : {len(deficits)} mois déficitaire(s)")
                    elif len(faibles)>0: parts.append(f"  🟡 **{aid}** : marge < 10% sur {len(faibles)} mois")
                    else: parts.append(f"  🟢 **{aid}** : finances saines")

            else:
                parts.append(f"Fichier `{Path(file_path).name}` — **{len(arts)} article(s)**")
                parts.append("Suggestions : ROI, seuil de rentabilité, alertes marges, P&L promotion.")

        else:
            parts.append(f"Agent `{agent}` prêt. Pose-moi une question sur tes données.")

    except Exception as e:
        parts.append(f"⚠️ Erreur lors de l'analyse : `{e}`")
        import traceback; parts.append(f"```\n{traceback.format_exc()[-400:]}\n```")

    return "\n\n".join(parts), excel_out


def do_orchestrate(question, ctx={}):
    q = question.lower()
    trace("orchestrateur", f"Demande reçue : « {question[:55]}{'...' if len(question)>55 else ''} »")

    if any(w in q for w in ["promo","remise","discount","simulation"]):
        agents = ["marketing","demande","production","finance"]
        trace("orchestrateur","Scénario PROMOTION → 4 agents mobilisés","decision")
    elif any(w in q for w in ["capacité","surcharge","prod","plan"]):
        agents = ["production"]
        trace("orchestrateur","Scénario CAPACITÉ → Agent Production","decision")
    elif any(w in q for w in ["forecast","prévision","demande","méthode","mape"]):
        agents = ["demande"]
        trace("orchestrateur","Scénario DEMANDE → Agent Demande","decision")
    elif any(w in q for w in ["roi","rentab","marge","finance","coût","budget"]):
        agents = ["finance"]
        trace("orchestrateur","Scénario FINANCE → Agent Finance","decision")
    else:
        agents = ["demande","production"]
        trace("orchestrateur","Question générale → Demande + Production","decision")

    results = {}
    for ag in agents:
        trace(ag, "Démarrage de l'analyse...", "start")
        fp = auto_file(ag)
        resp, excel = think(ag, question, fp)
        results[ag] = {"resp": resp, "excel": excel}
        trace(ag, resp[:120].replace("\n"," ").replace("**","")+"...", "success")

    # Synthèse
    trace("orchestrateur", "Synthèse des résultats en cours...", "synthesis")
    pct = next((int(m) for m in re.findall(r"(\d+)\s*%", q)), 20)

    lines = [f"J'ai mobilisé **{len(agents)} agent(s)** : {', '.join([AGENTS[a]['icon']+' '+a.capitalize() for a in agents])}\n"]
    for ag in agents:
        r = results[ag]
        cfg = AGENTS[ag]
        first_line = r["resp"].split("\n")[0].replace("**","").replace("`","")[:80]
        lines.append(f"{cfg['icon']} **{ag.capitalize()}** → {first_line}")

    # Décision finale
    has_fin = "finance" in results and results["finance"]["resp"]
    if has_fin and "Rentable" in results["finance"]["resp"]:
        lines.append("\n**Décision finale : ✅ GO — Conditions favorables**")
    elif has_fin and "Non rentable" in results["finance"]["resp"]:
        lines.append("\n**Décision finale : ❌ NO-GO — Vérifier la rentabilité**")
    else:
        lines.append("\n**Analyse complète disponible — consulte chaque onglet agent.**")

    trace("orchestrateur", lines[-1].replace("**","").replace("\n",""), "decision")
    return "\n".join(lines), agents, results


# ── Render helpers ────────────────────────────────────────────────────────────
def render_msgs(agent):
    for m in st.session_state.chats[agent]:
        cfg = AGENTS[agent]
        if m["role"] == "user":
            st.markdown(f'<div class="chat-wrap"><div class="msg u"><div class="avatar" style="background:var(--bg3);color:var(--muted)">U</div><div class="bubble u">{m["content"]}</div></div></div>', unsafe_allow_html=True)
        else:
            content_html = m["content"].replace("\n","<br>").replace("**","<strong>",1)
            # Fix unclosed strong tags
            import re as _re
            content_html = _re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', m["content"].replace("\n","<br>"))
            content_html = content_html.replace("`","<code>",1).replace("`","</code>",1)
            content_html = _re.sub(r'`(.*?)`', r'<code>\1</code>', m["content"].replace("\n","<br>"))
            st.markdown(f'<div class="chat-wrap"><div class="msg a"><div class="avatar" style="{cfg["av_css"]}">{cfg["icon"]}</div><div class="bubble a">{content_html}</div></div></div>', unsafe_allow_html=True)
            if m.get("excel") and Path(m["excel"]).exists():
                with open(m["excel"],"rb") as f:
                    st.download_button(f"⬇ Excel modifié — {Path(m['excel']).name}", data=f.read(),
                                       file_name=Path(m["excel"]).name,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       key=f"dl_{agent}_{m['ts']}")

def render_file_panel(agent):
    st.markdown('<div class="sec-hdr">📁 SOURCE DE DONNÉES</div>', unsafe_allow_html=True)
    fi = st.session_state.files.get(agent)
    if fi:
        st.markdown(f'<div class="fbadge">✓ {fi["name"]}</div>', unsafe_allow_html=True)
        with st.expander("Aperçu données", expanded=False):
            if not fi["df"].empty:
                st.dataframe(fi["df"], use_container_width=True, hide_index=True, height=140)
    up = st.file_uploader("Importer fichier Excel", type=["xlsx"], key=f"up_{agent}", label_visibility="collapsed")
    if up:
        save_upload(agent, up)
        add_msg(agent, "agent", f"✅ Fichier **{up.name}** importé. Je suis prêt à analyser tes données.")
        st.rerun()

def render_caps(agent, caps):
    st.markdown('<div class="sec-hdr">⚙ CAPACITÉS</div>', unsafe_allow_html=True)
    for c in caps:
        st.markdown(f'<div class="cap-item"><span>✦</span> {c}</div>', unsafe_allow_html=True)

def render_quick(agent):
    st.markdown('<div style="padding:.4rem .75rem 0;font-size:.67rem;color:var(--muted);font-family:Space Mono,monospace;letter-spacing:.08em">SUGGESTIONS RAPIDES</div>', unsafe_allow_html=True)
    prompts = QUICK[agent]
    cols = st.columns(min(len(prompts),3))
    for i, p in enumerate(prompts):
        with cols[i%3]:
            if st.button(p, key=f"q_{agent}_{i}", use_container_width=True):
                fp = auto_file(agent)
                add_msg(agent, "user", p)
                with st.spinner(f"Agent {agent} réfléchit..."):
                    resp, excel = think(agent, p, fp)
                add_msg(agent, "agent", resp, excel=excel)
                st.rerun()

def send_panel(agent, placeholder):
    c1,c2 = st.columns([6,1])
    with c1:
        inp = st.text_input("", key=f"inp_{agent}", label_visibility="collapsed", placeholder=placeholder)
    with c2:
        btn = st.button("→", key=f"btn_{agent}")
    if btn and inp:
        fp = auto_file(agent)
        add_msg(agent, "user", inp)
        with st.spinner(f"Agent {agent} réfléchit..."):
            resp, excel = think(agent, inp, fp)
        add_msg(agent, "agent", resp, excel=excel)
        st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# TOPBAR
# ══════════════════════════════════════════════════════════════════════════════
files_loaded = len(all_paths())
st.markdown(f"""
<div class="topbar">
  <div class="topbar-logo">⬡ S&OP AGENTIQUE</div>
  <div class="topbar-sep"></div>
  <div class="topbar-sub">5 AGENTS AUTONOMES · ANALYSE · ACTION · DÉCISION</div>
  <div style="margin-left:auto;font-size:.72rem;font-family:'Space Mono',monospace;color:{'#10b981' if files_loaded else '#6b6b80'}">
    {'✓ ' + str(files_loaded) + ' fichier(s) chargé(s)' if files_loaded else '○ Aucun fichier'}
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tabs = st.tabs(["⬡ Orchestrateur","📈 Demande","🏭 Production","📣 Marketing","💰 Finance"])

# ── TAB 0 : ORCHESTRATEUR ────────────────────────────────────────────────────
with tabs[0]:
    L, R = st.columns([1, 1.9], gap="small")
    with L:
        render_file_panel("orchestrateur")
        st.markdown('<div class="sec-hdr">⬡ TRACE DES AGENTS</div>', unsafe_allow_html=True)
        STATUS_C = {"info":"#6b6b80","decision":"#f59e0b","start":"#3b82f6","success":"#10b981","error":"#ef4444","data":"#8b5cf6","synthesis":"#f59e0b"}
        trace_c = st.container(height=280)
        with trace_c:
            for tr in st.session_state.traces[-25:]:
                c = STATUS_C.get(tr["status"],"#6b6b80")
                ac = AGENTS.get(tr["agent"], AGENTS["orchestrateur"])
                st.markdown(
                    f'<div class="trace-row"><div class="trace-dot" style="background:{c}"></div>'
                    f'<div class="trace-agent" style="color:{ac["color"]}">{ac["icon"]} {tr["agent"][:6].upper()}</div>'
                    f'<div class="trace-msg">{tr["msg"]}</div>'
                    f'<div class="trace-time">{tr["ts"]}</div></div>', unsafe_allow_html=True)
        if st.button("🗑 Effacer", key="clear_t"):
            st.session_state.traces = []; st.rerun()

    with R:
        st.markdown('<div class="sec-hdr">⬡ ORCHESTRATEUR — COORDINATION DES AGENTS</div>', unsafe_allow_html=True)
        chat_c = st.container(height=390)
        with chat_c:
            render_msgs("orchestrateur")
        render_quick("orchestrateur")
        c1,c2 = st.columns([6,1])
        with c1:
            oi = st.text_input("", key="oi", label_visibility="collapsed",
                               placeholder="Ex: Simule une promo -20% et dis-moi si c'est rentable")
        with c2:
            ob = st.button("→", key="ob")
        if ob and oi:
            add_msg("orchestrateur","user", oi)
            with st.spinner("Orchestrateur en action..."):
                resp, agents_called, _ = do_orchestrate(oi)
            add_msg("orchestrateur","agent", resp)
            st.rerun()

# ── TAB 1 : DEMANDE ──────────────────────────────────────────────────────────
with tabs[1]:
    L, R = st.columns([1, 1.9], gap="small")
    with L:
        render_file_panel("demande")
        render_caps("demande",["Prévision LES / Holt-Winters / ARIMA / MA","Classification auto des séries","Sélection par MAE / MAPE / RMSE","Détection anomalies & saisonnalité","Modification directe fichier Excel"])
    with R:
        st.markdown('<div class="sec-hdr">📈 AGENT DEMANDE — PRÉVISIONS & ANALYSE</div>', unsafe_allow_html=True)
        with st.container(height=380):
            render_msgs("demande")
        render_quick("demande")
        send_panel("demande","Ex: Calcule le forecast 6 mois · Analyse la saisonnalité · Compare les méthodes")

# ── TAB 2 : PRODUCTION ───────────────────────────────────────────────────────
with tabs[2]:
    L, R = st.columns([1, 1.9], gap="small")
    with L:
        render_file_panel("production")
        render_caps("production",["Analyse capacité par ressource (PHR)","Détection surcharges & alertes","Ajustement automatique plan MPS","Simulation hausse/baisse demande","Modification directe fichier Excel"])
    with R:
        st.markdown('<div class="sec-hdr">🏭 AGENT PRODUCTION — MPS & CAPACITÉ</div>', unsafe_allow_html=True)
        with st.container(height=380):
            render_msgs("production")
        render_quick("production")
        send_panel("production","Ex: Y a-t-il des surcharges ? · Optimise le plan · Simulation +30%")

# ── TAB 3 : MARKETING ────────────────────────────────────────────────────────
with tabs[3]:
    L, R = st.columns([1, 1.9], gap="small")
    with L:
        render_file_panel("marketing")
        render_caps("marketing",["Simulation promotions (élasticité prix)","Calcul d'uplift par article","Analyse saisonnalité réelle","Comparaison multi-scénarios","Modification directe fichier Excel"])
    with R:
        st.markdown('<div class="sec-hdr">📣 AGENT MARKETING — PROMOTIONS & UPLIFT</div>', unsafe_allow_html=True)
        with st.container(height=380):
            render_msgs("marketing")
        render_quick("marketing")
        send_panel("marketing","Ex: Simule une promo -20% · Quel est le meilleur mois ? · Compare -10% vs -30%")

# ── TAB 4 : FINANCE ──────────────────────────────────────────────────────────
with tabs[4]:
    L, R = st.columns([1, 1.9], gap="small")
    with L:
        render_file_panel("finance")
        render_caps("finance",["Calcul P&L & ROI par scénario","Seuil de rentabilité automatique","Vérification budget promotion","Alertes marge & déficit","Modification directe fichier Excel"])
    with R:
        st.markdown('<div class="sec-hdr">💰 AGENT FINANCE — P&L & ROI</div>', unsafe_allow_html=True)
        with st.container(height=380):
            render_msgs("finance")
        render_quick("finance")
        send_panel("finance","Ex: Calcule le ROI promotion · Seuil de rentabilité · Alerte si marge < 10%")
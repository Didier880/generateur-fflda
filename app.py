import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
import time as pytime
import json
import io
import os
import pickle
import zipfile
import urllib.request
import re
import collections
import streamlit.components.v1 as components
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Générateur Officiel FFLDA", page_icon="🤼", layout="wide")

# --- OPTIMISATION DE PERFORMANCE & FRAGMENTS STREAMLIT ---
def fragment_compat(run_every=None):
    """
    Décorateur garantissant la réactivité ultra-rapide des composants (Scoreboard, Déroulé)
    sans recharger l'intégralité de l'application Streamlit.
    """
    def decorator(fn):
        if hasattr(st, "fragment"):
            if run_every:
                return st.fragment(run_every=run_every)(fn)
            return st.fragment(fn)
        elif hasattr(st, "experimental_fragment"):
            if run_every:
                return st.experimental_fragment(run_every=run_every)(fn)
            return st.experimental_fragment(fn)
        return fn
    return decorator

# --- GESTION DE SESSION & SÉCURITÉ ---
def get_cle_secrete_table(code_session):
    """
    Génère un jeton cryptographique déterministe pour sécuriser l'accès aux tables de marque.
    Seuls les QR codes imprimés sur les tables contiennent ce jeton.
    """
    import hashlib
    sel = "FFLDA_SECURE_TABLE_2026_SALT"
    code_str = str(code_session or "").strip().upper()
    return hashlib.sha256(f"{code_str}_{sel}".encode()).hexdigest()[:8]

# --- CONNEXION AUTOMATIQUE VIA PARAMÈTRES D'URL (QR CODE / LIEN DIRECT TAPIS) ---
try:
    qp = getattr(st, "query_params", {})
    qp_code = qp.get("code")
    qp_tapis = qp.get("tapis")
    qp_cle = qp.get("cle")
    qp_mode = qp.get("mode")
    qp_vue = qp.get("vue") or qp.get("ecran")
    qp_sk = qp.get("sk")
    qp_su = qp.get("su")
    if qp_sk:
        st.session_state["supabase_key"] = str(qp_sk).strip()
    if qp_su:
        st.session_state["supabase_url"] = str(qp_su).strip()
    if qp_code:
        st.session_state["code_session"] = str(qp_code).strip().upper()
        st.session_state["authentifie"] = True
    if qp_cle:
        st.session_state["table_autorisee"] = True
    if qp_vue and str(qp_vue).lower() in ["scoreboard", "score", "tv", "ecran", "public"]:
        st.session_state["vue_scoreboard_active"] = True
        st.session_state["mode_kiosque_qr"] = True
    if qp_tapis:
        if str(qp_tapis).lower() in ["all", "tous", "global", "deroule", "0"]:
            st.session_state["radio_tapis_direct"] = "📋 Déroulé Général (Tous Tapis)"
        else:
            st.session_state["radio_tapis_direct"] = f"Tapis {qp_tapis}"
        st.session_state["mode_app_index"] = 1
        if not st.session_state.get("force_full_menu", False):
            st.session_state["mode_kiosque_qr"] = True
    elif qp_mode in ["direct", "2"]:
        st.session_state["mode_app_index"] = 1
except Exception:
    pass

if st.session_state.get("vue_scoreboard_active"):
    st.markdown('''
    <style>
    header[data-testid="stHeader"] { display: none !important; }
    footer { display: none !important; }
    #MainMenu { visibility: hidden !important; }
    section[data-testid="stSidebar"] { display: none !important; }
    div[data-testid="collapsedControl"] { display: none !important; }
    .block-container {
        padding-top: 0.2rem !important;
        padding-bottom: 0.2rem !important;
        padding-left: 0.6rem !important;
        padding-right: 0.6rem !important;
        max-width: 100% !important;
    }
    div[data-testid="stVerticalBlock"] {
        gap: 0.3rem !important;
    }
    </style>
    ''', unsafe_allow_html=True)

# --- ÉCRAN DE CONNEXION OBLIGATOIRE ---
if not st.session_state.get("authentifie"):
    st.markdown("<br><br>", unsafe_allow_html=True)
    col_acc1, col_acc2, col_acc3 = st.columns([1, 2, 1])
    with col_acc2:
        import os
        if os.path.exists("logo_fflda.png"):
            st.image("logo_fflda.png", use_container_width=True)
        else:
            st.image("https://www.fflutte.com/content/uploads/2021/10/fflutte-bleu-1024x842.png", use_container_width=True)
        
        st.markdown("<h2 style='text-align: center; color: #0055A4;'>🔒 Espace Sécurisé Organisateur FFLDA</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #555;'>Veuillez saisir votre code d'accès organisateur pour déverrouiller l'application.</p>", unsafe_allow_html=True)
        st.markdown("---")
        
        code_saisi = st.text_input("🔑 Code d'accès / Organisateur", placeholder="Ex: TEST2", key="input_code_login")
        
        if st.button("Se Connecter", use_container_width=True, type="primary"):
            c_clean = str(code_saisi).strip().upper()
            if c_clean:
                st.session_state["authentifie"] = True
                st.session_state["code_session"] = c_clean
                st.rerun()
            else:
                st.error("Veuillez renseigner un code d'accès valide.")
                
        st.markdown("---")
        st.caption("Fédération Française de Lutte et Disciplines Associées — Plateforme Officielle de Gestion de Tournois")
    st.stop()

URL_BILLING_WEBHOOK_DEFAUT = "https://script.google.com/macros/s/AKfycbxboXVY0FbLYQX6ZeBgDt4lg2fJ6eDIxfW-1BswZRoz5sLqAqBLCUGk7sLHoqpMW_C0/exec"

def enregistrer_log_facturation(code_organisateur, nom_tournoi, nb_inscrits, nb_peses, nb_matchs, details_resultats=None):
    """
    Transmet silencieusement en arrière-plan les métriques anonymisées du tournoi sans bloquer l'interface.
    """
    try:
        import json
        url_webhook = None
        try:
            url_webhook = st.secrets.get("BILLING_WEBHOOK_URL", URL_BILLING_WEBHOOK_DEFAUT)
        except Exception:
            pass
        if not url_webhook:
            url_webhook = st.session_state.get("url_billing_webhook", URL_BILLING_WEBHOOK_DEFAUT)
            
        if not url_webhook or "http" not in url_webhook:
            return
            
        try:
            from zoneinfo import ZoneInfo
            now_str = datetime.now(ZoneInfo("Europe/Paris")).strftime("%d/%m/%Y %H:%M:%S")
        except Exception:
            now_str = (datetime.utcnow() + timedelta(hours=2)).strftime("%d/%m/%Y %H:%M:%S")
            
        payload = {
            "code_organisateur": str(code_organisateur or "ANONYME"),
            "nom_tournoi": str(nom_tournoi or "Tournoi sans nom"),
            "date": now_str,
            "total_inscrits": int(nb_inscrits or 0),
            "total_peses": int(nb_peses or 0),
            "total_matchs": int(nb_matchs or 0),
            "details_resultats": json.dumps(details_resultats or [], ensure_ascii=False)
        }
        
        data = urllib.parse.urlencode(payload).encode('utf-8')
        req = urllib.request.Request(url_webhook, data=data, headers={'User-Agent': 'FFLDA-Billing/1.0'})
        with urllib.request.urlopen(req, timeout=2.0):
            pass
    except Exception:
        pass

# --- OPTION GOOGLE SHEETS DÉSACTIVÉE (APPLICATION 100% AUTONOME) ---
def upload_xlsx_to_google_sheets(excel_bytes, title):
    return None, "Option Google Sheets désactivée."

def download_google_sheet_as_xlsx(url_or_id):
    return None, "Option Google Sheets désactivée."

# --- MODULE SUPABASE & SAISIE TAPIS EN DIRECT (OPTION 2) ---
URL_SUPABASE_DEFAUT = "https://jjakkhuqgdetbkkhepnv.supabase.co"
KEY_SUPABASE_DEFAUT = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpqYWtraHVxZ2RldGJra2hlcG52Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTA3Nzg5NDUsImV4cCI6MjEwNjM1NDk0NX0.welFMmzlA-7J2qur0-7AMSy_tvx2bcpb31uYtWr3Jqo"

def nettoyer_url_supabase(url_brute):
    if not url_brute:
        return ""
    u = str(url_brute).strip().rstrip('/')
    while u.endswith('/rest/v1'):
        u = u[:-8].rstrip('/')
    return u

def get_supabase_config():
    # 1. Valeurs par défaut intégrées
    url = nettoyer_url_supabase(URL_SUPABASE_DEFAUT)
    key = KEY_SUPABASE_DEFAUT.strip()
    
    # 2. Si Secrets Streamlit définis et valides (pas un texte d'exemple)
    if hasattr(st, "secrets"):
        try:
            s_u = nettoyer_url_supabase(st.secrets.get("SUPABASE_URL", ""))
            s_k = str(st.secrets.get("SUPABASE_KEY", "")).strip()
            if s_u and "http" in s_u:
                url = s_u
            if s_k and s_k.startswith("eyJ") and "votre_cle" not in s_k and len(s_k) > 100:
                key = s_k
        except Exception:
            pass
            
    # 3. Session state (seulement si valide et non exemple)
    if st.session_state.get("supabase_url"):
        sess_u = nettoyer_url_supabase(st.session_state.get("supabase_url"))
        if sess_u and "http" in sess_u:
            url = sess_u
    if st.session_state.get("supabase_key"):
        sess_k = str(st.session_state.get("supabase_key")).strip()
        if sess_k and sess_k.startswith("eyJ") and "votre_cle" not in sess_k and len(sess_k) > 100:
            key = sess_k
        
    url = nettoyer_url_supabase(url)
    key = str(key).strip()
    is_ok = bool(url and key and url.startswith("http"))
    return url, key, is_ok

COLONNES_AUTORISEES_MATCHS_LUTTE = {
    "id", "tournoi_id", "code_organisateur", "tapis", "match_num",
    "heure", "duree", "categorie", "tour",
    "lutteur_rouge", "club_rouge", "comite_rouge",
    "lutteur_bleu", "club_bleu", "comite_bleu",
    "statut", "vainqueur", "type_victoire",
    "score_rouge", "score_bleu", "pt_clt_rouge", "pt_clt_bleu",
    "sheet", "typer_cell", "tot_r_cell", "typeb_cell", "tot_b_cell"
}

def supabase_request(endpoint, method="GET", data=None, params=None):
    url, key, is_ok = get_supabase_config()
    if not is_ok:
        return None, "NO_CREDENTIALS"
    
    clean_url = nettoyer_url_supabase(url)
    full_url = clean_url + "/rest/v1/" + endpoint.lstrip('/')
    if params:
        import urllib.parse
        full_url += "?" + urllib.parse.urlencode(params)
        
    import urllib.request, urllib.error, json
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json"
    }
    if method == "POST":
        headers["Prefer"] = "resolution=merge-duplicates"
        
    # Filtrer automatiquement les colonnes pour correspondre exactement à la table SQL
    if "matchs_lutte" in endpoint and data is not None:
        if isinstance(data, list):
            data = [{k: v for k, v in item.items() if k in COLONNES_AUTORISEES_MATCHS_LUTTE} for item in data]
        elif isinstance(data, dict):
            data = {k: v for k, v in data.items() if k in COLONNES_AUTORISEES_MATCHS_LUTTE}
            
    req_body = json.dumps(data).encode("utf-8") if data is not None else None
    req = urllib.request.Request(full_url, data=req_body, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            resp_body = resp.read().decode("utf-8")
            return json.loads(resp_body) if resp_body else {}, None
    except urllib.error.HTTPError as e:
        err_detail = ""
        try:
            err_detail = e.read().decode("utf-8")
        except Exception:
            pass
        return None, f"HTTP {e.code} ({e.reason}) : {err_detail} [Cible: {full_url}]" if err_detail else f"HTTP Error {e.code}: {e.reason} [Cible: {full_url}]"
    except Exception as e:
        return None, str(e)

def calculer_pts_fflda_match(categorie, vainqueur, type_victoire, score_rouge, score_bleu):
    """
    Calcule les points de classement officiel FFLDA (Pt Clt) pour le vainqueur et le perdant
    selon la catégorie d'âge (U13 vs U9/U11).
    """
    age_cat = extraire_age_de_texte(str(categorie)).upper()
    is_u13 = "U13" in age_cat
    
    score_vainq = score_rouge if vainqueur == "Rouge" else score_bleu
    score_perd = score_bleu if vainqueur == "Rouge" else score_rouge
    
    if is_u13:
        if type_victoire.startswith("VT"):
            pt_v, pt_p = 5, 0
        elif type_victoire.startswith("VST"):
            pt_v = 4
            pt_p = 1 if score_perd >= 1 else 0
        elif type_victoire.startswith("VP"):
            pt_v = 3
            pt_p = 1 if score_perd >= 1 else 0
        else:
            pt_v, pt_p = 5, 0
    else:
        # U9 / U11
        pt_v = 2
        pt_p = 1 if score_perd >= 1 else 0
        
    pt_rouge = pt_v if vainqueur == "Rouge" else pt_p
    pt_bleu = pt_p if vainqueur == "Rouge" else pt_v
    return pt_rouge, pt_bleu

def formater_badge_victoire_html(type_victoire, compact=False):
    """
    Génère un badge / pastille HTML stylé, moderne et arrondi pour le type de victoire :
    - VT : Violet ⚡ Tombé (VT)
    - VST : Orange 💥 Supériorité (VST)
    - VP : Bleu 🎯 Aux Points (VP)
    - Abandon / Forfait / Disqualification : Ardoise 🛑 Forfait / Abandon
    """
    if not type_victoire:
        return ""
    tv = str(type_victoire).strip().upper()
    
    if "VT" in tv or "TOMB" in tv:
        bg = "#7e22ce"
        border = "#a855f7"
        label = "⚡ VT" if compact else "⚡ Tombé (VT)"
    elif "VST" in tv or "SUP" in tv:
        bg = "#ea580c"
        border = "#f97316"
        label = "💥 VST" if compact else "💥 Supériorité (VST)"
    elif "VP" in tv or "POINT" in tv:
        bg = "#2563eb"
        border = "#60a5fa"
        label = "🎯 VP" if compact else "🎯 Aux Points (VP)"
    elif any(k in tv for k in ["AB", "FOR", "DSQ", "DISQ"]):
        bg = "#475569"
        border = "#94a3b8"
        label = "🛑 FOR" if compact else "🛑 Forfait / Abandon"
    else:
        bg = "#334155"
        border = "#64748b"
        label = f"⚖️ {tv}"
        
    pad = "2px 7px" if compact else "3px 10px"
    fsize = "11px" if compact else "12px"
    return f"<span style='background-color:{bg}; color:white; border:1px solid {border}; padding:{pad}; border-radius:12px; font-weight:800; font-size:{fsize}; display:inline-block; vertical-align:middle; line-height:1.2; box-shadow:0 1px 3px rgba(0,0,0,0.25); margin: 0 4px;'>{label}</span>"

def formater_label_court_victoire(type_victoire):
    """Génère un libellé emoji élégant pour les tableaux de données Streamlit."""
    if not type_victoire or str(type_victoire).strip() in ["-", ""]:
        return "-"
    tv = str(type_victoire).strip().upper()
    if "VT" in tv or "TOMB" in tv:
        return "⚡ VT (Tombé)"
    elif "VST" in tv or "SUP" in tv:
        return "💥 VST (Supériorité)"
    elif "VP" in tv or "POINT" in tv:
        return "🎯 VP (Points)"
    elif any(k in tv for k in ["AB", "FOR", "DSQ", "DISQ"]):
        return "🛑 Forfait / Abandon"
    return str(type_victoire)

def evaluer_departage_uww(actions_r, actions_b, cautions_r=0, cautions_b=0):
    """
    Évalue le meneur et le motif du départage selon les règles officielles UWW :
    1. Score total
    2. Si égalité au score :
       - Valeur maximale des prises techniques (5, 4, 2, 1)
       - Nombre de prises de cette valeur max (ex: 3x2 bat 2x2)
       - Nombre d'avertissements (le moins de cartons gagne)
       - Dernier point technique marqué (chronologie)
    """
    score_r = sum(a.get('val', 0) for a in actions_r)
    score_b = sum(a.get('val', 0) for a in actions_b)

    if score_r > score_b:
        return 'Rouge', f"Score technique ({score_r} - {score_b})", 0
    if score_b > score_r:
        return 'Bleu', f"Score technique ({score_r} - {score_b})", 0
    if score_r == 0 and score_b == 0:
        return None, "Score vierge (0 - 0)", 0

    # Prises techniques de valeur (5, 4, 2 points uniquement). 
    # Les points à 1 pt (sorties de zone, passivités, pénalités) ne sont pas des prises de grande valeur.
    tech_vals = [5, 4, 2]
    counts_r = {v: sum(1 for a in actions_r if a.get('val') == v and a.get('type') == 'tech') for v in tech_vals}
    counts_b = {v: sum(1 for a in actions_b if a.get('val') == v and a.get('type') == 'tech') for v in tech_vals}

    # 1. Plus haute valeur technique (5, 4, 2)
    max_r = max([v for v in tech_vals if counts_r[v] > 0], default=0)
    max_b = max([v for v in tech_vals if counts_b[v] > 0], default=0)
    if max_r > max_b:
        return 'Rouge', f"Plus haute valeur technique (prise à {max_r} pts)", 1
    elif max_b > max_r:
        return 'Bleu', f"Plus haute valeur technique (prise à {max_b} pts)", 1

    # 2. Plus grand nombre de prises de cette valeur max (5, 4, 2)
    if max_r > 0:
        if counts_r[max_r] > counts_b[max_r]:
            return 'Rouge', f"Plus grand nombre de {max_r} pts ({counts_r[max_r]} contre {counts_b[max_r]})", 2
        elif counts_b[max_r] > counts_r[max_r]:
            return 'Bleu', f"Plus grand nombre de {max_r} pts ({counts_b[max_r]} contre {counts_r[max_r]})", 2

        # Comparer successivement les valeurs inférieures (seulement parmi 5, 4, 2, jamais les 1 pt)
        for v in [x for x in tech_vals if x < max_r]:
            if counts_r[v] > counts_b[v]:
                return 'Rouge', f"Plus grand nombre de {v} pts ({counts_r[v]} contre {counts_b[v]})", 2
            elif counts_b[v] > counts_r[v]:
                return 'Bleu', f"Plus grand nombre de {v} pts ({counts_b[v]} contre {counts_r[v]})", 2

    # 3. Moins d'avertissements (cautions)
    if cautions_r < cautions_b:
        return 'Rouge', f"Moins d'avertissements ({cautions_r} contre {cautions_b})", 3
    elif cautions_b < cautions_r:
        return 'Bleu', f"Moins d'avertissements ({cautions_b} contre {cautions_r})", 3

    # 4. Dernier point technique marqué (chronologie seq)
    last_seq_r = max([a.get('seq', 0) for a in actions_r], default=-1)
    last_seq_b = max([a.get('seq', 0) for a in actions_b], default=-1)
    if last_seq_r > last_seq_b:
        return 'Rouge', "Dernier point marqué", 4
    elif last_seq_b > last_seq_r:
        return 'Bleu', "Dernier point marqué", 4

    return 'Rouge', "Priorité Coin Rouge", 5

# --- SYNCHRONISATION MULTI-ÉCRANS TEMPS RÉEL (TABLE DE MARQUE & SCOREBOARD) ---
_SHARED_ACTIONS_SYNC = {}
_SHARED_TAPIS_STATE = {}
_SHARED_MATCHS_DATA = {}
_SHARED_TOURNAMENT_MATCHS = {}

def detecter_ip_reseau_local():
    """Détecte l'adresse IP locale de la machine sur le réseau local Wi-Fi / Ethernet."""
    for iface in ["en0", "en1", "en2", "wlan0", "eth0"]:
        try:
            import subprocess
            out = subprocess.check_output(["ipconfig", "getifaddr", iface], stderr=subprocess.DEVNULL, timeout=0.5).decode().strip()
            if out and not out.startswith("127."):
                return out
        except Exception:
            pass
    try:
        import subprocess
        out = subprocess.check_output(["hostname", "-I"], stderr=subprocess.DEVNULL, timeout=0.5).decode().split()
        if out and out[0] and not out[0].startswith("127."):
            return out[0].strip()
    except Exception:
        pass
    try:
        import socket
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(0.2)
        s.connect(('10.255.255.255', 1))
        ip = s.getsockname()[0]
        s.close()
        if ip and not ip.startswith("127."):
            return ip
    except Exception:
        pass
    return "localhost"

def definir_combat_actif_tapis(tapis_num, idx_match, match_id=None, dernier_termine=None):
    """Enregistre le combat actuellement ouvert sur la table de marque d'un tapis."""
    try:
        t_key = int(tapis_num)
        prev_state = dict(_SHARED_TAPIS_STATE.get(t_key, {}))
        p = f".cache_tournois/sync_tapis_{t_key}.json"
        if not prev_state and os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    prev_state = json.load(f)
            except Exception:
                pass
        new_state = {
            "active_idx": int(idx_match),
            "match_id": str(match_id or ""),
            "ts": pytime.time()
        }
        if dernier_termine:
            new_state["dernier_termine"] = dernier_termine
        elif "dernier_termine" in prev_state:
            new_state["dernier_termine"] = prev_state["dernier_termine"]

        _SHARED_TAPIS_STATE[t_key] = new_state
        os.makedirs(".cache_tournois", exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(new_state, f)
    except Exception:
        pass

def recuperer_etat_tapis(tapis_num):
    """Récupère l'état actif d'un tapis (combat en cours sélectionné par l'arbitre)."""
    try:
        t_key = int(tapis_num)
        p = f".cache_tournois/sync_tapis_{t_key}.json"
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                d = json.load(f)
                _SHARED_TAPIS_STATE[t_key] = d
                return dict(d)
        if t_key in _SHARED_TAPIS_STATE:
            return dict(_SHARED_TAPIS_STATE[t_key])
    except Exception:
        pass
    return {}

def publier_actions_match_sync(match_id, actions_r=None, actions_b=None, cautions_r=None, cautions_b=None, extra=None):
    """Publie en temps réel les actions, scores et avertissements d'un combat."""
    try:
        cur = _SHARED_ACTIONS_SYNC.setdefault(match_id, {})
        if actions_r is not None:
            cur["actions_r"] = list(actions_r)
        if actions_b is not None:
            cur["actions_b"] = list(actions_b)
        if cautions_r is not None:
            cur["cautions_r"] = int(cautions_r)
        if cautions_b is not None:
            cur["cautions_b"] = int(cautions_b)
        if extra and isinstance(extra, dict):
            cur.update(extra)
        cur["ts"] = pytime.time()

        safe_id = re.sub(r'[^a-zA-Z0-9_-]', '_', str(match_id))
        os.makedirs(".cache_tournois", exist_ok=True)
        with open(f".cache_tournois/live_act_{safe_id}.json", "w", encoding="utf-8") as f:
            json.dump(cur, f)
    except Exception:
        pass

def recuperer_actions_match_sync(match_id):
    """Récupère les actions et points synchronisés en direct pour un combat."""
    try:
        safe_id = re.sub(r'[^a-zA-Z0-9_-]', '_', str(match_id))
        p = f".cache_tournois/live_act_{safe_id}.json"
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                d = json.load(f)
                _SHARED_ACTIONS_SYNC[match_id] = d
                return dict(d)
        if match_id in _SHARED_ACTIONS_SYNC:
            return dict(_SHARED_ACTIONS_SYNC[match_id])
    except Exception:
        pass
    return {}

def publier_mise_a_jour_match(code_sess, match_obj):
    """Publie la mise à jour d'un combat (sauvegarde, validation, correction) pour tous les écrans."""
    try:
        m_id = match_obj.get("id")
        if not m_id:
            return
        _SHARED_MATCHS_DATA[m_id] = dict(match_obj)
        
        # Mettre à jour les listes partagées
        for c_k in [code_sess, "DEFAULT"]:
            if c_k in _SHARED_TOURNAMENT_MATCHS:
                for idx_m, m_sh in enumerate(_SHARED_TOURNAMENT_MATCHS[c_k]):
                    if m_sh.get("id") == m_id:
                        _SHARED_TOURNAMENT_MATCHS[c_k][idx_m] = dict(match_obj)
                        break
                        
        # Mettre à jour le cache de tournoi sur disque si existant
        if code_sess:
            safe_c = re.sub(r'[^a-zA-Z0-9_-]', '_', str(code_sess))
            p_cache = f".cache_tournois/cache_{safe_c}.pkl"
            if os.path.exists(p_cache):
                with open(p_cache, "rb") as f_r:
                    bndl = pickle.load(f_r)
                if bndl and "matchs_direct" in bndl:
                    for idx_m, m_sh in enumerate(bndl["matchs_direct"]):
                        if m_sh.get("id") == m_id:
                            bndl["matchs_direct"][idx_m] = dict(match_obj)
                            with open(p_cache, "wb") as f_w:
                                pickle.dump(bndl, f_w)
                            break
    except Exception:
        pass

def ajouter_action_lutte(match_id, coin, val, type_action="tech", label=""):
    import streamlit as st
    cle_act = f"actions_{'r' if coin == 'Rouge' else 'b'}_{match_id}"
    if cle_act not in st.session_state:
        st.session_state[cle_act] = []
    
    st.session_state["uww_action_counter"] = st.session_state.get("uww_action_counter", 0) + 1
    seq = st.session_state["uww_action_counter"]
    
    st.session_state[cle_act].append({
        "val": val,
        "type": type_action,
        "label": label or str(val),
        "seq": seq
    })
    _SHARED_ACTIONS_SYNC.setdefault(match_id, {})["actions_" + ('r' if coin == 'Rouge' else 'b')] = list(st.session_state[cle_act])
    publier_actions_match_sync(
        match_id,
        actions_r=st.session_state.get(f"actions_r_{match_id}", []),
        actions_b=st.session_state.get(f"actions_b_{match_id}", []),
        cautions_r=st.session_state.get(f"cautions_r_{match_id}", 0),
        cautions_b=st.session_state.get(f"cautions_b_{match_id}", 0)
    )

def ajouter_avertissement_lutte(match_id, coin_qui_faut):
    import streamlit as st
    cle_caut = f"cautions_{'r' if coin_qui_faut == 'Rouge' else 'b'}_{match_id}"
    st.session_state[cle_caut] = st.session_state.get(cle_caut, 0) + 1
    
    # L'avertissement est enregistré directement chez le lutteur fautif, sans donner de point automatique
    cle_act = f"actions_{'r' if coin_qui_faut == 'Rouge' else 'b'}_{match_id}"
    if cle_act not in st.session_state:
        st.session_state[cle_act] = []
    
    st.session_state["uww_action_counter"] = st.session_state.get("uww_action_counter", 0) + 1
    seq = st.session_state["uww_action_counter"]
    
    st.session_state[cle_act].append({
        "val": 0,
        "type": "avert",
        "label": "⚠️",
        "seq": seq
    })
    _SHARED_ACTIONS_SYNC.setdefault(match_id, {})["cautions_" + ('r' if coin_qui_faut == 'Rouge' else 'b')] = st.session_state[cle_caut]
    _SHARED_ACTIONS_SYNC.setdefault(match_id, {})["actions_" + ('r' if coin_qui_faut == 'Rouge' else 'b')] = list(st.session_state[cle_act])
    publier_actions_match_sync(
        match_id,
        actions_r=st.session_state.get(f"actions_r_{match_id}", []),
        actions_b=st.session_state.get(f"actions_b_{match_id}", []),
        cautions_r=st.session_state.get(f"cautions_r_{match_id}", 0),
        cautions_b=st.session_state.get(f"cautions_b_{match_id}", 0)
    )

def annuler_action_lutte(match_id, coin):
    import streamlit as st
    cle_act = f"actions_{'r' if coin == 'Rouge' else 'b'}_{match_id}"
    if cle_act in st.session_state and len(st.session_state[cle_act]) > 0:
        derniere = st.session_state[cle_act].pop()
        # Si c'était un avertissement, on décrémente le compteur de ce lutteur
        if derniere.get("type") == "avert":
            cle_caut = f"cautions_{'r' if coin == 'Rouge' else 'b'}_{match_id}"
            if cle_caut in st.session_state and st.session_state[cle_caut] > 0:
                st.session_state[cle_caut] -= 1
            _SHARED_ACTIONS_SYNC.setdefault(match_id, {})["cautions_" + ('r' if coin == 'Rouge' else 'b')] = st.session_state.get(cle_caut, 0)
        _SHARED_ACTIONS_SYNC.setdefault(match_id, {})["actions_" + ('r' if coin == 'Rouge' else 'b')] = list(st.session_state[cle_act])
        publier_actions_match_sync(
            match_id,
            actions_r=st.session_state.get(f"actions_r_{match_id}", []),
            actions_b=st.session_state.get(f"actions_b_{match_id}", []),
            cautions_r=st.session_state.get(f"cautions_r_{match_id}", 0),
            cautions_b=st.session_state.get(f"cautions_b_{match_id}", 0)
        )

def reinitialiser_score_lutteur(match_id, coin):
    import streamlit as st
    cle_act = f"actions_{'r' if coin == 'Rouge' else 'b'}_{match_id}"
    st.session_state[cle_act] = []
    _SHARED_ACTIONS_SYNC.setdefault(match_id, {})["actions_" + ('r' if coin == 'Rouge' else 'b')] = []
    publier_actions_match_sync(
        match_id,
        actions_r=st.session_state.get(f"actions_r_{match_id}", []),
        actions_b=st.session_state.get(f"actions_b_{match_id}", []),
        cautions_r=st.session_state.get(f"cautions_r_{match_id}", 0),
        cautions_b=st.session_state.get(f"cautions_b_{match_id}", 0)
    )

def annuler_derniere_action_globale(match_id):
    """
    Annule la toute dernière action enregistrée sur le match (qu'elle vienne de Rouge ou de Bleu),
    en comparant le numéro de séquence (seq) des dernières actions.
    """
    import streamlit as st
    cle_act_r = f"actions_r_{match_id}"
    cle_act_b = f"actions_b_{match_id}"
    acts_r = st.session_state.get(cle_act_r, [])
    acts_b = st.session_state.get(cle_act_b, [])

    if not acts_r and not acts_b:
        return

    seq_r = acts_r[-1].get("seq", 0) if acts_r else -1
    seq_b = acts_b[-1].get("seq", 0) if acts_b else -1

    if seq_r >= seq_b and acts_r:
        annuler_action_lutte(match_id, "Rouge")
    elif acts_b:
        annuler_action_lutte(match_id, "Bleu")

def decomposer_score_lutte(score):
    """
    Décompose un score total en une suite d'actions de lutte réalistes (+4, +2, +1)
    correspondant exactement aux touches du clavier (+5, +4, +2, +1).
    Permet de retrouver l'addition de points plutôt qu'un chiffre brut monolithique.
    """
    if score <= 0:
        return []
    if score == 8:
        return [4, 2, 2]
    res = []
    reste = score
    while reste >= 4 and reste not in [6, 2, 3]:
        res.append(4)
        reste -= 4
    while reste >= 2:
        res.append(2)
        reste -= 2
    while reste >= 1:
        res.append(1)
        reste -= 1
    return res

def sauvegarder_cache_actions_match(code_org, match_id, actions_r, actions_b, cautions_r=0, cautions_b=0):
    try:
        import os, json, re
        os.makedirs(".cache_tournois", exist_ok=True)
        safe_code = re.sub(r'[^a-zA-Z0-9_-]', '_', str(code_org or "ORG"))
        p = f".cache_tournois/actions_matchs_{safe_code}.json"
        data = {}
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = {}
        data[str(match_id)] = {
            "actions_r": actions_r,
            "actions_b": actions_b,
            "cautions_r": cautions_r,
            "cautions_b": cautions_b
        }
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    except Exception:
        pass

def charger_cache_actions_match(code_org, match_id, coin):
    try:
        import os, json, re
        safe_code = re.sub(r'[^a-zA-Z0-9_-]', '_', str(code_org or "ORG"))
        p = f".cache_tournois/actions_matchs_{safe_code}.json"
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            if str(match_id) in data:
                m_data = data[str(match_id)]
                key_c = "actions_r" if coin == "Rouge" else "actions_b"
                return m_data.get(key_c, [])
    except Exception:
        pass
    return None

def charger_actions_match(match_id, coin, score_total, m_info=None, code_org=""):
    """
    Récupère ou reconstitue fidèlement la décomposition des points (l'addition réelle)
    pour un lutteur (Rouge ou Bleu) lors du retour sur un combat déjà enregistré.
    """
    import streamlit as st
    cle_act = f"actions_{'r' if coin == 'Rouge' else 'b'}_{match_id}"
    
    # 1. Si déjà présent en session_state et non vide
    if cle_act in st.session_state and len(st.session_state[cle_act]) > 0:
        current_acts = st.session_state[cle_act]
        # Si c'était un bloc monolithique d'ancienne version non standard (ex: [8] ou [7]), on le décompose
        if len(current_acts) == 1 and current_acts[0].get("val") not in [1, 2, 4, 5] and score_total > 0:
            pass
        else:
            return st.session_state[cle_act]

    # 2. Vérifier si encodé dans le match (m_info["actions_..."] ou tot_r_cell/tot_b_cell)
    if m_info and isinstance(m_info, dict):
        act_key = "actions_rouge" if coin == "Rouge" else "actions_bleu"
        if m_info.get(act_key) and isinstance(m_info[act_key], list) and len(m_info[act_key]) > 0:
            return list(m_info[act_key])
        
        cell_field = "tot_r_cell" if coin == "Rouge" else "tot_b_cell"
        cell_val = str(m_info.get(cell_field) or "")
        if "#" in cell_val:
            acts_part = cell_val.split("#")[1].split(";")[0]
            tokens = [t.strip() for t in acts_part.split(",") if t.strip()]
            parsed_acts = []
            for idx_t, tok in enumerate(tokens, 1):
                try:
                    v_tok = int(tok)
                    if v_tok > 0:
                        parsed_acts.append({"val": v_tok, "type": "tech", "label": str(v_tok), "seq": idx_t})
                except Exception:
                    pass
            if parsed_acts:
                return parsed_acts

    # 3. Vérifier le cache disque local
    cached_acts = charger_cache_actions_match(code_org, match_id, coin)
    if cached_acts:
        return cached_acts

    # 4. Reconstitution intelligente de l'addition de lutte (+4, +2, +1)
    if score_total > 0:
        points_decomposes = decomposer_score_lutte(score_total)
        reconstructed = []
        for idx_d, v_pts in enumerate(points_decomposes, 1):
            reconstructed.append({"val": v_pts, "type": "tech", "label": str(v_pts), "seq": idx_d})
        return reconstructed

    return []

def injecter_scores_matchs_dans_classeur(excel_bytes, liste_matchs):
    """
    Injecte les résultats et scores saisis en direct sur smartphone/tablette
    dans les cellules exactes de la Grille Tapis du classeur Excel officiel.
    """
    import openpyxl, io
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=False)
    for m in liste_matchs:
        if m.get("statut") == "Terminé":
            sheet_name = m.get("sheet")
            if sheet_name and sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                typer_c = m.get("typer_cell")
                tot_r_c = str(m.get("tot_r_cell") or "").split("#")[0].strip()
                typeb_c = m.get("typeb_cell")
                tot_b_c = str(m.get("tot_b_cell") or "").split("#")[0].strip()
                
                v = m.get("vainqueur")
                t_vic = m.get("type_victoire", "VT")
                
                if v == "Rouge":
                    tr = t_vic
                    tb = "DT" if t_vic == "VT" else ("DST" if t_vic == "VST" else "DP")
                else:
                    tb = t_vic
                    tr = "DT" if t_vic == "VT" else ("DST" if t_vic == "VST" else "DP")
                    
                if typer_c: ws[typer_c] = tr
                if tot_r_c: ws[tot_r_c] = int(m.get("score_rouge", 0))
                if typeb_c: ws[typeb_c] = tb
                if tot_b_c: ws[tot_b_c] = int(m.get("score_bleu", 0))

                # Injecter aussi les Points de Classement (Pt Clt) dans la ligne d'en-tête du match sur Grille Tapis
                # row_h est 2 lignes au-dessus de tot_r_c (ex: E8 -> E6 et I6)
                pt_r = int(m.get("pt_clt_rouge", 0) or 0)
                pt_b = int(m.get("pt_clt_bleu", 0) or 0)
                if pt_r == 0 and pt_b == 0 and v:
                    pt_r, pt_b = calculer_pts_fflda_match(m.get("categorie", ""), v, t_vic, int(m.get("score_rouge", 0)), int(m.get("score_bleu", 0)))
                
                if tot_r_c:
                    digits = "".join(filter(str.isdigit, tot_r_c))
                    if digits:
                        row_h = int(digits) - 2
                        ws[f"E{row_h}"] = pt_r
                        ws[f"I{row_h}"] = pt_b
                
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def generer_archive_excel_combats(liste_matchs, excel_base=None):
    """
    Génère un classeur Excel d'archivage sécurisé :
    - Si le classeur officiel initial (excel_base) est fourni, injecte les scores dans les grilles officielles.
    - Sinon, crée un classeur Excel récapitulatif détaillé avec l'ensemble des combats, scores, vainqueurs et points.
    """
    import io, pandas as pd
    if excel_base:
        try:
            return injecter_scores_matchs_dans_classeur(excel_base, liste_matchs).getvalue()
        except Exception:
            pass

    rows = []
    for m in (liste_matchs or []):
        rows.append({
            "Tapis": m.get("tapis", 1),
            "N° Match": m.get("match_num", ""),
            "Catégorie": m.get("categorie", ""),
            "Tour": m.get("tour", ""),
            "Statut": m.get("statut", ""),
            "Lutteur Rouge": m.get("lutteur_rouge", ""),
            "Club Rouge": m.get("club_rouge", ""),
            "Score Rouge": m.get("score_rouge", 0),
            "Score Bleu": m.get("score_bleu", 0),
            "Lutteur Bleu": m.get("lutteur_bleu", ""),
            "Club Bleu": m.get("club_bleu", ""),
            "Vainqueur": m.get("vainqueur", ""),
            "Type Victoire": m.get("type_victoire", ""),
            "Pts Clt Rouge": m.get("pt_clt_rouge", 0),
            "Pts Clt Bleu": m.get("pt_clt_bleu", 0)
        })
    df_exp = pd.DataFrame(rows) if rows else pd.DataFrame(columns=[
        "Tapis", "N° Match", "Catégorie", "Tour", "Statut",
        "Lutteur Rouge", "Club Rouge", "Score Rouge", "Score Bleu",
        "Lutteur Bleu", "Club Bleu", "Vainqueur", "Type Victoire",
        "Pts Clt Rouge", "Pts Clt Bleu"
    ])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_exp.to_excel(writer, index=False, sheet_name="Résultats Combats")
    buf.seek(0)
    return buf.getvalue()

# --- MENU LATÉRAL (PARAMÈTRES INTERACTIFS) ---
with st.sidebar:
    import os
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
    with col_l2:
        if os.path.exists("logo_fflda.png"):
            st.image("logo_fflda.png", use_container_width=True)
        else:
            st.image("https://www.fflutte.com/content/uploads/2021/10/fflutte-bleu-1024x842.png", use_container_width=True)
    if st.session_state.get("authentifie"):
        col_s1, col_s2 = st.columns([3, 1])
        with col_s1:
            code_sess_aff = st.session_state.get('code_session', 'DIRECT')
            st.caption(f"**Session** : `{code_sess_aff}`")
        with col_s2:
            if st.button("🔒", help="Se déconnecter de l'application"):
                st.session_state["authentifie"] = False
                st.session_state.pop("code_session", None)
                if hasattr(st, "query_params"):
                    st.query_params.clear()
                st.rerun()

    st.markdown("### Configuration du Tournoi")
    st.markdown("---")
    
    # --- NOM DE LA COMPÉTITION & NOMBRE DE TAPIS (LE STRICT MINIMUM EN AMONT) ---
    nom_competition = st.text_input("🏆 Nom de la compétition", value="Tournoi Officiel FFLDA - U7/U9/U11/U13", key="cfg_nom_competition")
    nb_tapis = st.number_input("Nombre de tapis", min_value=1, max_value=10, value=3, key="cfg_nb_tapis")

    # Valeurs officielles par défaut pour la génération initiale (ajustables en aval une fois le tournoi affiché)
    type_pesee = "2 Pesées (U7/U9 puis U11/U13)"
    heure_pesee_u9 = time(9, 0)
    duree_pesee = 45
    activer_pause = True
    duree_pause = 45

    with st.expander("⚙️ Paramètres avancés (Règles & Temps)", expanded=False):
        mixte_active = st.checkbox("Catégories Mixtes (U7, U9, U11 uniquement)", value=True, help="Conformément à la réglementation officielle FFLDA, la mixité n'existe plus en U13 : les catégories Féminine (LF), Libre (LL) et Gréco (LG) sont toujours strictement séparées.", key="cfg_mixte_active")
        poules_par_niveau = st.checkbox("Créer des poules par niveau (débutants/confirmés)", value=True, key="cfg_poules_par_niveau")
        separer_clubs = st.checkbox("Éviter les lutteurs d'un même club dans la même poule (dans la mesure du possible)", value=True, key="cfg_separer_clubs")
        eviter_arbitre_meme_club = st.checkbox("Éviter les matchs entre arbitres et lutteurs du même club", value=True, key="cfg_eviter_arbitre_meme_club")
        meme_tapis_poule = st.checkbox("Maintenir chaque poule / lutteur sur un même tapis", value=True, help="Chaque catégorie de même style et de même poids (ex: U13 Gréco 30 kg) est affectée intégralement à un tapis fixe unique pour tous ses combats. Les catégories de styles ou poids différents sont réparties de manière à équilibrer au mieux le nombre total de matchs par tapis.", key="cfg_meme_tapis_poule")
        tolerance_poids = st.number_input("Tolérance d'écart de poids (%) [U7, U9, U11]", min_value=10, max_value=15, value=10, step=1, key="cfg_tolerance_poids")
        repos_matchs = st.number_input("Matchs de repos minimum", min_value=1, max_value=10, value=3, key="cfg_repos_matchs")
        duree_plateau_u7 = st.number_input("Temps de chaque plateau U7 (min)", min_value=3, max_value=30, value=10, step=1, help="L'animation U7 est sous forme de 3 plateaux d'activités avec rotation. Par exemple, 10 min par plateau = 3x10 min = 30 min consacrées aux U7 au total.", key="cfg_duree_plateau_u7")
        duree_u7 = duree_plateau_u7 * 3
        duree_u9 = st.number_input("Temps total U9 (min)", value=3, key="cfg_duree_u9")
        duree_u11 = st.number_input("Temps total U11 (min)", value=4, key="cfg_duree_u11")
        duree_u13 = st.number_input("Temps total U13 (min)", value=5, key="cfg_duree_u13")

    # JavaScript pour le comportement d'accordéon à ouverture unique (ferme les autres au clic)
    components.html("""
    <script>
    (function() {
        const parentDoc = window.parent.document;
        function initAccordion() {
            const sidebar = parentDoc.querySelector('section[data-testid="stSidebar"]');
            if (!sidebar) return;
            const expanders = sidebar.querySelectorAll('div[data-testid="stExpander"]');
            expanders.forEach(exp => {
                if (exp.dataset.accordionAttached) return;
                exp.dataset.accordionAttached = "true";
                exp.addEventListener('click', function() {
                    setTimeout(() => {
                        const details = exp.querySelector('details');
                        if (details && details.open) {
                            expanders.forEach(otherExp => {
                                if (otherExp !== exp) {
                                    const otherDetails = otherExp.querySelector('details');
                                    if (otherDetails) {
                                        otherDetails.open = false;
                                    }
                                }
                            });
                        }
                    }, 50);
                }, false);
            });
        }
        initAccordion();
        const observer = new MutationObserver(initAccordion);
        observer.observe(parentDoc.body, { childList: true, subtree: true });
    })();
    </script>
    """, height=0)

def formater_poids(val):
    if val is None or str(val).strip() in ['', 'None', 'nan']:
        return ''
    val_str = str(val).lower().replace('kg', '').replace(',', '.').strip()
    try:
        val_float = round(float(val_str), 1)
        if val_float == int(val_float):
            num_str = str(int(val_float))
        else:
            num_str = str(val_float)
        return f"{num_str} kg"
    except (ValueError, TypeError):
        s = str(val).strip()
        return f"{s} kg" if (s and not s.lower().endswith('kg')) else s

def formater_poids_court(val):
    if val is None or str(val).strip() in ['', 'None', 'nan']:
        return ''
    val_str = str(val).lower().replace('kg', '').replace(',', '.').strip()
    try:
        val_float = round(float(val_str), 1)
        if val_float == int(val_float):
            num_str = str(int(val_float))
        else:
            num_str = str(val_float)
        return f"{num_str}kg"
    except (ValueError, TypeError):
        s = str(val).strip()
        return f"{s}kg" if (s and not s.lower().endswith('kg')) else s

def abreger_nom_onglet(nom_poule):
    txt = str(nom_poule)
    txt = txt.replace("Mixte (LL/LF)", "Mxt").replace("LG (Gréco)", "LG").replace("LL (Libre)", "LL").replace("LF (Féminine)", "LF")
    txt = txt.replace(" | ", " ").replace(" (", " ").replace(")", "").replace(" - ", "-")
    txt = txt.replace("/", "-").replace("\\", "-").replace(":", "-").replace("?", "").replace("*", "")
    txt = re.sub(r'\s+', ' ', txt)
    return txt[:31].strip()

def nettoyer_nom_tour(val):
    if not val:
        return ""
    val_str = str(val).strip()
    if val_str.lower().startswith("tour "):
        return val_str[5:].strip()
    return val_str

def charger_liste_arbitres(fichier_arbitres_in=None):
    """
    Charge la liste des arbitres inscrits depuis le fichier téléversé ou le fichier par défaut FFLDA - Inscription arbitres.xlsx.
    """
    arbitres = []
    filepath = None
    
    if fichier_arbitres_in is not None:
        filepath = fichier_arbitres_in
    else:
        import os
        default_name = 'FFLDA - Inscription arbitres.xlsx'
        if os.path.exists(default_name):
            filepath = default_name
            
    if filepath is None:
        return arbitres
        
    try:
        if hasattr(filepath, 'name') and filepath.name.endswith('.csv'):
            df_arb = pd.read_csv(filepath, sep=';', encoding='utf-8')
            if len(df_arb.columns) == 1:
                filepath.seek(0)
                df_arb = pd.read_csv(filepath, sep=',', encoding='utf-8')
        elif isinstance(filepath, str) and filepath.endswith('.csv'):
            df_arb = pd.read_csv(filepath, sep=';', encoding='utf-8')
            if len(df_arb.columns) == 1:
                df_arb = pd.read_csv(filepath, sep=',', encoding='utf-8')
        else:
            df_temp = pd.read_excel(filepath, nrows=5)
            header_row = 0
            for i, row in df_temp.iterrows():
                if 'Licence' in str(row.values) or 'Nom' in str(row.values) or "Inscrit Par" in str(row.values):
                    header_row = i + 1
                    break
            if hasattr(filepath, 'seek'):
                filepath.seek(0)
            df_arb = pd.read_excel(filepath, header=header_row)

        col_nom = next((c for c in df_arb.columns if str(c).strip().lower() == 'nom'), None)
        col_prenom = next((c for c in df_arb.columns if 'prénom' in str(c).strip().lower() or 'prenom' in str(c).strip().lower()), None)
        col_club = next((c for c in df_arb.columns if any(k in str(c).strip().lower() for k in ['sigle du club', 'club'])), None)
        col_comite = next((c for c in df_arb.columns if any(k in str(c).strip().lower() for k in ['comité', 'comite', 'ligue', 'région', 'region'])), None)
        col_licence = next((c for c in df_arb.columns if 'licence' in str(c).strip().lower()), None)

        for _, row in df_arb.iterrows():
            nom_val = str(row[col_nom]).strip() if (col_nom and pd.notna(row[col_nom])) else ''
            prenom_val = str(row[col_prenom]).strip() if (col_prenom and pd.notna(row[col_prenom])) else ''
            if not nom_val or nom_val.lower() in ['nan', 'none', 'photo']:
                continue
            
            club_val = str(row[col_club]).strip() if (col_club and pd.notna(row[col_club])) else 'Indépendant'
            comite_val = str(row[col_comite]).strip() if (col_comite and pd.notna(row[col_comite])) else 'Comité Non Renseigné'
            licence_val = str(row[col_licence]).strip() if (col_licence and pd.notna(row[col_licence])) else ''

            arbitres.append({
                'Nom_Complet': f"{nom_val} {prenom_val}".strip(),
                'Nom': nom_val,
                'Prenom': prenom_val,
                'Licence': licence_val,
                'Club': club_val if club_val not in ['', 'None', 'nan', '-'] else 'Indépendant',
                'Comite': comite_val if comite_val not in ['', 'None', 'nan', '-'] else 'Comité Non Renseigné'
            })
    except Exception:
        pass
        
    return arbitres

# --- CORPS PRINCIPAL ---
is_kiosque = st.session_state.get("mode_kiosque_qr", False) or st.session_state.get("vue_scoreboard_active", False)

if is_kiosque:
    st.markdown("""
    <style>
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    #MainMenu { visibility: hidden !important; }
    header { visibility: hidden !important; }
    footer { visibility: hidden !important; }
    .block-container { padding-top: 1.2rem !important; padding-bottom: 2rem !important; max-width: 98% !important; }
    </style>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <style>
    /* ========================================================================= */
    /* --- DESIGN VENDEUR & ARENA PRO FFLDA (SYSTEM THEME & ERGONOMICS) -------- */
    /* ========================================================================= */

    :root {
        --text-color: #0f172a !important;
        --primary-color: #0055A4 !important;
    }

    html, body, [data-testid="stAppViewContainer"], .stApp {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif !important;
        letter-spacing: -0.01em;
        background-color: #f8fafc !important;
        background: #f8fafc !important;
        color: #0f172a !important;
    }

    /* Lisibilité sombre par défaut sur tous les éléments textuels */
    .stApp p, .stApp span, .stApp label, .stApp div[data-testid="stMarkdownContainer"] p,
    .stApp div[data-testid="stMarkdownContainer"] span, .stApp li, .stApp td, .stApp th, .stApp h1, .stApp h2, .stApp h3, .stApp h4 {
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
    }

    /* Éléments avec fond sombre explicite gardent leur texte blanc net */
    .stApp [style*="background: linear-gradient(135deg, #0b1528"],
    .stApp [style*="background: linear-gradient(135deg, #0b1528"] *,
    .stApp [style*="background: #2e7d32"],
    .stApp [style*="background: #2e7d32"] *,
    .stApp [style*="background: #0055A4"],
    .stApp [style*="background: #0055A4"] *,
    .stApp [style*="background:#0055A4"],
    .stApp [style*="background:#0055A4"] *,
    .stApp [style*="background:#334155"],
    .stApp [style*="background:#334155"] *,
    .stApp [style*="background: #ea580c"],
    .stApp [style*="background: #ea580c"] * {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }

    /* --- ACCORDÉONS (st.expander) SURÉLEVÉS & DESIGN VENDEUR --- */
    div[data-testid="stExpander"] {
        border-radius: 14px !important;
        border: 1.5px solid #cbd5e1 !important;
        background: #ffffff !important;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.05) !important;
        margin-bottom: 14px !important;
        overflow: hidden !important;
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }

    div[data-testid="stExpander"]:hover {
        border-color: #94a3b8 !important;
        box-shadow: 0 6px 20px rgba(0, 85, 164, 0.09) !important;
    }

    div[data-testid="stExpander"] details {
        border-radius: 14px !important;
    }

    /* En-tête de l'accordéon : Plus haut, plus confortable au toucher (52px+) */
    div[data-testid="stExpander"] details summary {
        min-height: 52px !important;
        padding: 12px 18px !important;
        background: linear-gradient(90deg, rgba(0, 85, 164, 0.04) 0%, rgba(248, 250, 252, 0.95) 100%) !important;
        border-radius: 12px !important;
        font-weight: 800 !important;
        font-size: 15px !important;
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
        cursor: pointer !important;
        display: flex !important;
        align-items: center !important;
        transition: background 0.2s ease, color 0.2s ease !important;
    }

    div[data-testid="stExpander"] details summary:hover {
        background: linear-gradient(90deg, rgba(0, 85, 164, 0.09) 0%, rgba(241, 245, 249, 1) 100%) !important;
        color: #0055A4 !important;
        -webkit-text-fill-color: #0055A4 !important;
    }

    div[data-testid="stExpander"] details summary,
    div[data-testid="stExpander"] details summary *,
    div[data-testid="stExpander"] details summary p,
    div[data-testid="stExpander"] details summary span,
    div[data-testid="stExpander"] details summary div {
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
        font-size: 15px !important;
        font-weight: 800 !important;
        margin: 0 !important;
    }

    div[data-testid="stExpander"] details summary:hover *,
    div[data-testid="stExpander"] details summary:hover p,
    div[data-testid="stExpander"] details summary:hover span,
    div[data-testid="stExpander"] details summary:hover div {
        color: #0055A4 !important;
        -webkit-text-fill-color: #0055A4 !important;
    }

    div[data-testid="stExpander"] details summary svg {
        color: #0055A4 !important;
        fill: #0055A4 !important;
        width: 20px !important;
        height: 20px !important;
        transition: transform 0.25s ease !important;
    }

    div[data-testid="stExpander"] details[open] {
        border-color: #0055A4 !important;
        box-shadow: 0 6px 22px rgba(0, 85, 164, 0.12) !important;
    }

    div[data-testid="stExpander"] details[open] summary {
        border-bottom: 1.5px solid rgba(0, 85, 164, 0.15) !important;
        border-bottom-left-radius: 0 !important;
        border-bottom-right-radius: 0 !important;
        background: linear-gradient(90deg, rgba(0, 85, 164, 0.08) 0%, rgba(241, 245, 249, 0.95) 100%) !important;
    }

    div[data-testid="stExpander"] details > div:last-child {
        padding: 16px 20px !important;
        background: #ffffff !important;
    }

    /* Contraste garanti de tous les textes à l'intérieur des accordéons */
    div[data-testid="stExpander"] [data-testid="stExpanderDetails"],
    div[data-testid="stExpander"] [data-testid="stExpanderDetails"] *,
    div[data-testid="stExpander"] [data-testid="stExpanderDetails"] p,
    div[data-testid="stExpander"] [data-testid="stExpanderDetails"] label,
    div[data-testid="stExpander"] [data-testid="stExpanderDetails"] [data-testid="stWidgetLabel"] p,
    div[data-testid="stExpander"] [data-testid="stExpanderDetails"] [data-testid="stMarkdownContainer"] > p,
    div[data-testid="stExpander"] [data-testid="stExpanderDetails"] caption,
    div[data-testid="stExpander"] [data-testid="stExpanderDetails"] small {
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
    }

    /* --- BOUTONS PRO (EFFET RELIEF, TACTILE & HOVER LIFT) --- */
    div[data-testid="stButton"] button {
        min-height: 44px !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        letter-spacing: 0.2px !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
        cursor: pointer !important;
    }

    /* Boutons Primaires : Bleu FFLDA avec texte blanc net */
    button[kind="primary"],
    button[data-testid="baseButton-primary"],
    div[data-testid="stButton"] button[kind="primary"],
    div[data-testid="stButton"] button[data-testid="baseButton-primary"] {
        background: linear-gradient(135deg, #0055A4 0%, #1d4ed8 100%) !important;
        background-color: #0055A4 !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(0, 85, 164, 0.28) !important;
    }

    button[kind="primary"] *,
    button[data-testid="baseButton-primary"] *,
    div[data-testid="stButton"] button[kind="primary"] *,
    div[data-testid="stButton"] button[data-testid="baseButton-primary"] *,
    div[data-testid="stButton"] button[kind="primary"] p,
    div[data-testid="stButton"] button[data-testid="baseButton-primary"] p {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }

    button[kind="primary"]:hover,
    button[data-testid="baseButton-primary"]:hover,
    div[data-testid="stButton"] button[kind="primary"]:hover,
    div[data-testid="stButton"] button[data-testid="baseButton-primary"]:hover {
        background: linear-gradient(135deg, #004488 0%, #1e40af 100%) !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 18px rgba(0, 85, 164, 0.42) !important;
    }

    button[kind="primary"]:active,
    button[data-testid="baseButton-primary"]:active,
    div[data-testid="stButton"] button[kind="primary"]:active,
    div[data-testid="stButton"] button[data-testid="baseButton-primary"]:active {
        transform: scale(0.98) !important;
    }

    /* Boutons Secondaires : Fond blanc avec texte sombre ardoise garanti */
    button[kind="secondary"],
    button[data-testid="baseButton-secondary"],
    div[data-testid="stButton"] button[kind="secondary"],
    div[data-testid="stButton"] button[data-testid="baseButton-secondary"],
    div[data-testid="stDownloadButton"] button,
    div[data-testid="stLinkButton"] a {
        background-color: #ffffff !important;
        background: #ffffff !important;
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
        border: 1.5px solid #cbd5e1 !important;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04) !important;
    }

    button[kind="secondary"] *,
    button[data-testid="baseButton-secondary"] *,
    div[data-testid="stButton"] button[kind="secondary"] *,
    div[data-testid="stButton"] button[data-testid="baseButton-secondary"] *,
    div[data-testid="stButton"] button[kind="secondary"] p,
    div[data-testid="stButton"] button[data-testid="baseButton-secondary"] p,
    div[data-testid="stButton"] button[kind="secondary"] span,
    div[data-testid="stButton"] button[data-testid="baseButton-secondary"] span,
    div[data-testid="stButton"] button[kind="secondary"] div,
    div[data-testid="stButton"] button[data-testid="baseButton-secondary"] div,
    div[data-testid="stDownloadButton"] button *,
    div[data-testid="stLinkButton"] a * {
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
    }

    button[kind="secondary"]:hover,
    button[data-testid="baseButton-secondary"]:hover,
    div[data-testid="stButton"] button[kind="secondary"]:hover,
    div[data-testid="stButton"] button[data-testid="baseButton-secondary"]:hover,
    div[data-testid="stDownloadButton"] button:hover,
    div[data-testid="stLinkButton"] a:hover {
        border-color: #0055A4 !important;
        color: #0055A4 !important;
        -webkit-text-fill-color: #0055A4 !important;
        background-color: #f1f5f9 !important;
        background: #f1f5f9 !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08) !important;
    }

    button[kind="secondary"]:hover *,
    button[data-testid="baseButton-secondary"]:hover *,
    div[data-testid="stButton"] button[kind="secondary"]:hover *,
    div[data-testid="stButton"] button[data-testid="baseButton-secondary"]:hover *,
    div[data-testid="stDownloadButton"] button:hover *,
    div[data-testid="stLinkButton"] a:hover * {
        color: #0055A4 !important;
        -webkit-text-fill-color: #0055A4 !important;
    }

    button[kind="secondary"]:active,
    button[data-testid="baseButton-secondary"]:active,
    div[data-testid="stButton"] button[kind="secondary"]:active,
    div[data-testid="stButton"] button[data-testid="baseButton-secondary"]:active {
        transform: scale(0.98) !important;
    }

    /* --- TÉLÉCHARGEMENT & LIENS (st.download_button, st.link_button) --- */
    div[data-testid="stDownloadButton"] button,
    div[data-testid="stLinkButton"] a {
        min-height: 44px !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        background-color: #ffffff !important;
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
        border: 1.5px solid #cbd5e1 !important;
        transition: all 0.2s ease !important;
    }

    /* --- CHAMPS DE FORMULAIRE & MENUS DÉROULANTS --- */
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
    div[data-testid="stTextInput"] div[data-baseweb="input"] > div,
    div[data-testid="stNumberInput"] div[data-baseweb="input"] > div {
        border-radius: 10px !important;
        border: 1.5px solid #cbd5e1 !important;
        background-color: #ffffff !important;
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
        transition: border-color 0.2s ease, box-shadow 0.2s ease !important;
    }

    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div:focus-within,
    div[data-testid="stTextInput"] div[data-baseweb="input"] > div:focus-within,
    div[data-testid="stNumberInput"] div[data-baseweb="input"] > div:focus-within {
        border-color: #0055A4 !important;
        box-shadow: 0 0 0 3px rgba(0, 85, 164, 0.18) !important;
    }

    div[data-testid="stSelectbox"] label p,
    div[data-testid="stSelectbox"] [data-testid="stWidgetLabel"] p,
    div[data-testid="stTextInput"] label p,
    div[data-testid="stTextInput"] [data-testid="stWidgetLabel"] p,
    div[data-testid="stNumberInput"] label p,
    div[data-testid="stNumberInput"] [data-testid="stWidgetLabel"] p,
    div[data-testid="stCheckbox"] label p,
    div[data-testid="stCheckbox"] label span,
    div[data-testid="stRadio"] label p,
    div[data-testid="stRadio"] label span {
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
        font-weight: 700 !important;
    }

    div[data-testid="stTextInput"] input,
    div[data-testid="stNumberInput"] input,
    div[data-testid="stSelectbox"] div[data-baseweb="select"] * {
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
    }

    /* --- ONGLETS (st.tabs) --- */
    div[data-testid="stTabs"] button[role="tab"] {
        min-height: 44px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        color: #64748b !important;
        -webkit-text-fill-color: #64748b !important;
        border-radius: 8px 8px 0 0 !important;
        transition: all 0.2s ease !important;
    }

    div[data-testid="stTabs"] button[role="tab"] p,
    div[data-testid="stTabs"] button[role="tab"] span {
        color: #64748b !important;
        -webkit-text-fill-color: #64748b !important;
    }

    div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
        color: #0055A4 !important;
        -webkit-text-fill-color: #0055A4 !important;
        border-bottom: 3px solid #0055A4 !important;
    }

    div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] p,
    div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] span {
        color: #0055A4 !important;
        -webkit-text-fill-color: #0055A4 !important;
    }

    /* --- ALERTES & NOTIFICATIONS (st.info, st.success, st.warning) --- */
    div[data-testid="stAlert"] {
        border-radius: 12px !important;
        border: 1px solid rgba(0,0,0,0.08) !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04) !important;
    }

    div[data-testid="stAlert"] p,
    div[data-testid="stAlert"] span {
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # Bannière Header Arena Pro
    st.markdown(
        f"<div style='background: linear-gradient(135deg, #0b1528 0%, #0055A4 100%); color: white; padding: 20px 26px; border-radius: 16px; margin-bottom: 18px; box-shadow: 0 8px 24px rgba(0, 85, 164, 0.25); display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 14px;'>"
        f"<div>"
        f"<div style='font-size: 11px; font-weight: 800; letter-spacing: 1.5px; text-transform: uppercase; color: #facc15; margin-bottom: 4px;'>FFLDA ARENA TOURNAMENT SUITE</div>"
        f"<div style='font-size: clamp(20px, 3vw, 32px); font-weight: 900; line-height: 1.15; margin: 0;'>{nom_competition}</div>"
        f"<div style='font-size: 13px; opacity: 0.9; margin-top: 4px;'>Plateforme officielle d'optimisation des tournois de lutte et d'édition des bilans fédéraux.</div>"
        f"</div>"
        f"<div style='background: rgba(255,255,255,0.12); backdrop-filter: blur(8px); border: 1px solid rgba(255,255,255,0.25); padding: 7px 15px; border-radius: 10px; font-size: 12px; font-weight: 700; white-space: nowrap;'>"
        f"FFLDA 2026 OFFICIAL"
        f"</div>"
        f"</div>",
        unsafe_allow_html=True
    )

def generer_rondes_fflda(participants_in):
    participants = list(participants_in)
    n = len(participants)
    
    if n <= 1:
        return []
    elif n == 2:
        return [[(participants[0], participants[1])]]
    elif n == 3:
        return [[(participants[0], participants[1])],
                [(participants[2], participants[0])],
                [(participants[1], participants[2])]]
    elif n == 4:
        return [[(participants[0], participants[1]), (participants[2], participants[3])],
                [(participants[0], participants[2]), (participants[1], participants[3])],
                [(participants[0], participants[3]), (participants[1], participants[2])]]
    elif n == 5:
        return [[(participants[0], participants[1]), (participants[2], participants[3])],
                [(participants[4], participants[0]), (participants[1], participants[2])],
                [(participants[3], participants[4]), (participants[0], participants[2])],
                [(participants[1], participants[3]), (participants[2], participants[4])],
                [(participants[0], participants[3]), (participants[1], participants[4])]]
    else:
        if len(participants) % 2 != 0:
            participants.append({"Nom": "BYE", "Club": "-"})
        num_p = len(participants)
        rondes = []
        for i in range(num_p - 1):
            matchs_ronde = []
            for j in range(num_p // 2):
                p1 = participants[j]
                p2 = participants[num_p - 1 - j]
                if p1["Nom"] != "BYE" and p2["Nom"] != "BYE":
                    matchs_ronde.append((p1, p2))
            rondes.append(matchs_ronde)
            participants.insert(1, participants.pop())
        return rondes

def attribuer_categorie_poids_u13(poids_val):
    try:
        p = float(poids_val) if (poids_val is not None and str(poids_val).strip() != '') else 0.0
    except (ValueError, TypeError):
        p = 0.0
    if p <= 30.0: return "30 kg"
    elif p <= 33.0: return "33 kg"
    elif p <= 36.0: return "36 kg"
    elif p <= 39.0: return "39 kg"
    elif p <= 42.0: return "42 kg"
    elif p <= 46.0: return "46 kg"
    elif p <= 50.0: return "50 kg"
    elif p <= 55.0: return "55 kg"
    elif p <= 60.0: return "60 kg"
    else: return "+60 kg"

ORDRE_POIDS_U13 = ["30 kg", "33 kg", "36 kg", "39 kg", "42 kg", "46 kg", "50 kg", "55 kg", "60 kg", "+60 kg"]

def interleave_bracket_slots(list_a, list_b):
    res = []
    i, j = 0, 0
    while i < len(list_a) or j < len(list_b):
        if i < len(list_a):
            res.append(list_a[i])
            i += 1
        if j < len(list_b):
            res.append(list_b[j])
            j += 1
    return res

def repartir_tableau_protection_clubs(participants, nb_byes, nb_prelim):
    """
    Répartit les participants entre exempts (byes) et tour préliminaire (prelim_pts)
    selon la Règle FFLDA de Protection des clubs.
    
    1. Attribue les exemptions prioritairement de manière à ce qu'aucun club n'ait
       plus de 'nb_prelim' lutteurs dans le tour préliminaire (évite obligatoirement
       les affrontements fratricides au tour préliminaire).
    2. Répartit équitablement les exempts entre les clubs qui ont plusieurs inscrits.
    3. Garantit que dans tous les matchs préliminaires (p1, p2), p1['Club'] != p2['Club'].
    """
    clubs = collections.defaultdict(list)
    for p in participants:
        c = str(p.get('Club', '') or '').strip()
        if not c or c in ['-', 'Comité Non Renseigné', 'Sans club']:
            clubs[f"_indiv_{id(p)}"].append(p)
        else:
            clubs[c].append(p)
            
    sorted_clubs = sorted(clubs.values(), key=len, reverse=True)
    
    byes_list = []
    prelim_list = []
    
    if nb_byes == 0:
        prelim_list = list(participants)
    else:
        club_byes_count = {i: 0 for i in range(len(sorted_clubs))}
        total_byes_assigned = 0
        
        # Pass 1 : Surplus > nb_prelim vers les byes
        for i, c_members in enumerate(sorted_clubs):
            surplus = len(c_members) - nb_prelim
            if surplus > 0:
                take = min(surplus, nb_byes - total_byes_assigned)
                club_byes_count[i] += take
                total_byes_assigned += take
                
        # Pass 2 : Byes attribués en priorité aux clubs multi-membres (>= 2)
        while total_byes_assigned < nb_byes:
            progress = False
            for i, c_members in enumerate(sorted_clubs):
                if total_byes_assigned >= nb_byes:
                    break
                if len(c_members) >= 2 and club_byes_count[i] < len(c_members):
                    club_byes_count[i] += 1
                    total_byes_assigned += 1
                    progress = True
            if not progress:
                # Clubs à 1 membre
                for i, c_members in enumerate(sorted_clubs):
                    if total_byes_assigned >= nb_byes:
                        break
                    if club_byes_count[i] < len(c_members):
                        club_byes_count[i] += 1
                        total_byes_assigned += 1
                        progress = True
            if not progress:
                break
                
        for i, c_members in enumerate(sorted_clubs):
            n_b = club_byes_count[i]
            byes_list.extend(c_members[:n_b])
            prelim_list.extend(c_members[n_b:])
            
    prelim_matches = []
    if nb_prelim > 0:
        club_groups = collections.defaultdict(list)
        for p in prelim_list:
            c = str(p.get('Club', '') or '').strip()
            club_groups[c].append(p)
        sorted_prelim_clubs = sorted(club_groups.values(), key=len, reverse=True)
        flattened = [p for grp in sorted_prelim_clubs for p in grp]
        
        M = nb_prelim
        half1 = flattened[:M]
        half2 = flattened[M:]
        
        pairs = []
        for i in range(M):
            p1 = half1[i]
            p2 = half2[i]
            if p1.get('Club') and p1.get('Club') == p2.get('Club') and p1.get('Club') not in ['-', '']:
                for j in range(M):
                    if j != i and half2[j].get('Club') != p1.get('Club') and half2[i].get('Club') != half1[j].get('Club'):
                        half2[i], half2[j] = half2[j], half2[i]
                        p2 = half2[i]
                        break
            pairs.append((p1, p2))
        prelim_matches = pairs

    return byes_list, prelim_matches

def generer_competition_u13(age, style_grp, suffixe_niveau, cat_poids, participants, separer_clubs=True):
    n = len(participants)
    if n == 0:
        return None
    elif n == 1:
        nom = f"{age} | {style_grp}{suffixe_niveau} | {cat_poids} (1 seul inscrit)"
        return {
            'nom': nom, 
            'participants': list(participants), 
            'rondes': [], 
            'type_formule': 'seul',
            'cat_poids': cat_poids,
            'style_grp': style_grp
        }
    elif n < 6:
        # Cas 1 : Poule nordique unique (2 à 5 lutteurs)
        nom = f"{age} | {style_grp}{suffixe_niveau} | {cat_poids} (Poule unique)"
        return {
            'nom': nom, 
            'participants': list(participants), 
            'rondes': generer_rondes_fflda(participants), 
            'type_formule': 'poule',
            'cat_poids': cat_poids,
            'style_grp': style_grp
        }
    elif n == 6:
        # Cas 2 : Exactement 6 lutteurs (2 poules de 3 + Phase finale croisée)
        if separer_clubs:
            clubs_vu = {}
            poule_a, poule_b = [], []
            for p in sorted(participants, key=lambda x: x.get('Club', '')):
                c = p.get('Club', '')
                if clubs_vu.get(c, 0) % 2 == 0:
                    if len(poule_a) < 3: poule_a.append(p)
                    else: poule_b.append(p)
                else:
                    if len(poule_b) < 3: poule_b.append(p)
                    else: poule_a.append(p)
                clubs_vu[c] = clubs_vu.get(c, 0) + 1
        else:
            poule_a = [participants[0], participants[2], participants[4]]
            poule_b = [participants[1], participants[3], participants[5]]
        
        while len(poule_a) < 3 and len(poule_b) > 3:
            poule_a.append(poule_b.pop())
        while len(poule_b) < 3 and len(poule_a) > 3:
            poule_b.append(poule_a.pop())

        rondes_a = generer_rondes_fflda(poule_a)
        rondes_b = generer_rondes_fflda(poule_b)
        
        r1 = (rondes_a[0] if len(rondes_a) > 0 else []) + (rondes_b[0] if len(rondes_b) > 0 else [])
        r2 = (rondes_a[1] if len(rondes_a) > 1 else []) + (rondes_b[1] if len(rondes_b) > 1 else [])
        r3 = (rondes_a[2] if len(rondes_a) > 2 else []) + (rondes_b[2] if len(rondes_b) > 2 else [])
        
        # Tour 4 : Demi-finales croisées
        sf1 = (
            {"Nom": f"1er Poule A ({cat_poids})", "Club": "Qualifié A", "Comité": "-"},
            {"Nom": f"2ème Poule B ({cat_poids})", "Club": "Qualifié B", "Comité": "-"}
        )
        sf2 = (
            {"Nom": f"1er Poule B ({cat_poids})", "Club": "Qualifié B", "Comité": "-"},
            {"Nom": f"2ème Poule A ({cat_poids})", "Club": "Qualifié A", "Comité": "-"}
        )
        r4 = [sf1, sf2]
        
        # Tour 5 : Finales (Or/Argent et Bronze unique)
        f_or = (
            {"Nom": f"Vainqueur 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"},
            {"Nom": f"Vainqueur 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
        )
        f_bronze = (
            {"Nom": f"Perdant 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"},
            {"Nom": f"Perdant 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
        )
        r5 = [f_or, f_bronze]
        
        nom = f"{age} | {style_grp}{suffixe_niveau} | {cat_poids} (2 Poules + Finales)"
        return {
            'nom': nom, 
            'participants': list(participants), 
            'rondes': [r1, r2, r3, r4, r5], 
            'type_formule': 'poules_croisees',
            'poule_a': poule_a,
            'poule_b': poule_b,
            'cat_poids': cat_poids,
            'style_grp': style_grp
        }
    else:
        # Cas 3 : Plus de 6 lutteurs (Tableau avec repêchage des 1/4 - 2 médailles de bronze)
        rondes = []
        if n == 7:
            # 7 lutteurs : 3 quarts de finale, 1 exempt direct en demi-finale
            if separer_clubs:
                clubs = collections.defaultdict(list)
                for p in participants:
                    c = str(p.get('Club', '') or '').strip()
                    if not c or c in ['-', 'Comité Non Renseigné', 'Sans club']:
                        clubs[f"_indiv_{id(p)}"].append(p)
                    else:
                        clubs[c].append(p)
                sorted_clubs = sorted(clubs.values(), key=len, reverse=True)
                # Choix de l'exempt : le 1er membre du club le plus représenté (pour protéger ses coéquipiers)
                p_exempt = sorted_clubs[0][0]
                reste = []
                first_skipped = False
                for grp in sorted_clubs:
                    for p in grp:
                        if not first_skipped and p is p_exempt:
                            first_skipped = True
                        else:
                            reste.append(p)
                # Appariement des 6 autres sans fratricide
                club_groups = collections.defaultdict(list)
                for p in reste:
                    c = str(p.get('Club', '') or '').strip()
                    club_groups[c].append(p)
                sorted_reste_clubs = sorted(club_groups.values(), key=len, reverse=True)
                flattened = [p for grp in sorted_reste_clubs for p in grp]
                h1 = flattened[:3]
                h2 = flattened[3:]
                pairs = []
                for i in range(3):
                    p1 = h1[i]
                    p2 = h2[i]
                    if p1.get('Club') and p1.get('Club') == p2.get('Club') and p1.get('Club') not in ['-', '']:
                        for j in range(3):
                            if j != i and h2[j].get('Club') != p1.get('Club') and h2[i].get('Club') != h1[j].get('Club'):
                                h2[i], h2[j] = h2[j], h2[i]
                                p2 = h2[i]
                                break
                    pairs.append((p1, p2))
                # Séparation Haut / Bas : q1 et q2 sont en Haut, q3 et p_exempt sont en Bas
                # Si un match contient un coéquipier de p_exempt, il doit être en q1 ou q2 (Haut)
                c_ex = str(p_exempt.get('Club', '') or '').strip()
                if c_ex and c_ex not in ['-', 'Comité Non Renseigné', 'Sans club']:
                    for i in range(3):
                        m = pairs[i]
                        has_teammate = any(p.get('Club') == c_ex for p in m)
                        if has_teammate and i == 2:
                            for target_i in [0, 1]:
                                if not any(p.get('Club') == c_ex for p in pairs[target_i]):
                                    pairs[2], pairs[target_i] = pairs[target_i], pairs[2]
                                    break
                q1, q2, q3 = pairs[0], pairs[1], pairs[2]
            else:
                q1 = (participants[0], participants[1])
                q2 = (participants[2], participants[3])
                q3 = (participants[4], participants[5])
                p_exempt = participants[6]

            rondes.append([q1, q2, q3])
            
            sf1 = (
                {"Nom": f"Vainqueur 1/4 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"},
                {"Nom": f"Vainqueur 1/4 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            sf2 = (
                {"Nom": f"Vainqueur 1/4 (3) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"},
                {"Nom": p_exempt['Nom'], "Club": p_exempt.get('Club', ''), "Comité": p_exempt.get('Comité', '-')}
            )
            rep1 = (
                {"Nom": f"Perdant 1/4 (1) [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/4 (2) [{cat_poids}]", "Club": "Repêché", "Comité": "-"}
            )
            rondes.append([sf1, sf2, rep1])
            
            f_or = (
                {"Nom": f"Vainqueur 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"},
                {"Nom": f"Vainqueur 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            f_b1 = (
                {"Nom": f"Vainqueur Repêchage 1 [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            f_b2 = (
                {"Nom": f"Perdant 1/4 (3) [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            rondes.append([f_or, f_b1, f_b2])

        elif n <= 16:
            # 8 <= n <= 16 : Qualification pour 8 quarts-de-finalistes
            nb_prelim = n - 8
            nb_byes = 16 - n
            
            if separer_clubs:
                byes, prelim_matches = repartir_tableau_protection_clubs(participants, nb_byes, nb_prelim)
                prelim_winners = []
                for i in range(nb_prelim):
                    prelim_winners.append({
                        "Nom": f"Vainqueur Prél. {i+1} [{cat_poids}]",
                        "Club": "Qualifié",
                        "Comité": "-"
                    })
            else:
                byes = participants[:nb_byes]
                prelim_pts = participants[nb_byes:]
                prelim_matches = []
                prelim_winners = []
                for i in range(nb_prelim):
                    p1 = prelim_pts[i * 2]
                    p2 = prelim_pts[i * 2 + 1]
                    prelim_matches.append((p1, p2))
                    prelim_winners.append({
                        "Nom": f"Vainqueur Prél. {i+1} [{cat_poids}]",
                        "Club": "Qualifié",
                        "Comité": "-"
                    })
                
            if prelim_matches:
                rondes.append(prelim_matches)
                
            slots_qf = [None] * 8
            order_prelim_slots = [7, 3, 5, 1, 6, 2, 4, 0]
            for w_idx, slot_idx in enumerate(order_prelim_slots[:nb_prelim]):
                slots_qf[slot_idx] = prelim_winners[w_idx]
                
            if separer_clubs:
                order_upper = [s for s in [0, 2, 1, 3] if slots_qf[s] is None]
                order_lower = [s for s in [4, 6, 5, 7] if slots_qf[s] is None]
                interleaved_bye_slots = interleave_bracket_slots(order_upper, order_lower)
                for b_idx, s_idx in enumerate(interleaved_bye_slots):
                    if b_idx < len(byes):
                        slots_qf[s_idx] = byes[b_idx]
            else:
                bye_idx = 0
                for s_idx in range(8):
                    if slots_qf[s_idx] is None:
                        slots_qf[s_idx] = byes[bye_idx]
                        bye_idx += 1
                        
            qf_matches = [
                (slots_qf[0], slots_qf[1]),
                (slots_qf[2], slots_qf[3]),
                (slots_qf[4], slots_qf[5]),
                (slots_qf[6], slots_qf[7])
            ]
            rondes.append(qf_matches)
            
            sf1 = (
                {"Nom": f"Vainqueur 1/4 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"},
                {"Nom": f"Vainqueur 1/4 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            sf2 = (
                {"Nom": f"Vainqueur 1/4 (3) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"},
                {"Nom": f"Vainqueur 1/4 (4) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            rep1 = (
                {"Nom": f"Perdant 1/4 (1) [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/4 (2) [{cat_poids}]", "Club": "Repêché", "Comité": "-"}
            )
            rep2 = (
                {"Nom": f"Perdant 1/4 (3) [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/4 (4) [{cat_poids}]", "Club": "Repêché", "Comité": "-"}
            )
            rondes.append([sf1, sf2, rep1, rep2])
            
            f_or = (
                {"Nom": f"Vainqueur 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"},
                {"Nom": f"Vainqueur 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            f_b1 = (
                {"Nom": f"Vainqueur Repêchage 1 [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            f_b2 = (
                {"Nom": f"Vainqueur Repêchage 2 [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            rondes.append([f_or, f_b1, f_b2])

        elif n <= 32:
            # 17 <= n <= 32 (Tableau de 32 : 1/16, 1/8, 1/4, 1/2, Finales & Repêchages)
            nb_prelim = n - 16
            nb_byes = max(0, 32 - n)
            
            if separer_clubs:
                byes_16, prelim_matches = repartir_tableau_protection_clubs(participants, nb_byes, nb_prelim)
                prelim_winners = []
                for i in range(nb_prelim):
                    prelim_winners.append({
                        "Nom": f"Vainqueur 1/16 ({i+1}) [{cat_poids}]",
                        "Club": "Qualifié",
                        "Comité": "-"
                    })
                rondes.append(prelim_matches)
                
                slots_16 = [None] * 16
                order_prelim_slots_16 = [15, 7, 11, 3, 13, 5, 9, 1, 14, 6, 10, 2, 12, 4, 8, 0]
                for w_idx, slot_idx in enumerate(order_prelim_slots_16[:nb_prelim]):
                    slots_16[slot_idx] = prelim_winners[w_idx]
                    
                order_upper = [s for s in [0, 4, 2, 6, 1, 5, 3, 7] if slots_16[s] is None]
                order_lower = [s for s in [8, 12, 10, 14, 9, 13, 11, 15] if slots_16[s] is None]
                interleaved_bye_slots = interleave_bracket_slots(order_upper, order_lower)
                for b_idx, s_idx in enumerate(interleaved_bye_slots):
                    if b_idx < len(byes_16):
                        slots_16[s_idx] = byes_16[b_idx]
                        
                for s_idx in range(16):
                    if slots_16[s_idx] is None:
                        slots_16[s_idx] = {"Nom": f"Qualifié 1/8 ({s_idx+1})", "Club": "Qualifié", "Comité": "-"}
            else:
                byes_16 = participants[:nb_byes]
                prelim_pts = participants[nb_byes:]
                prelim_matches = []
                prelim_winners = []
                for i in range(nb_prelim):
                    p1 = prelim_pts[i * 2] if i * 2 < len(prelim_pts) else {"Nom": f"Lutteur {i*2+1}", "Club": "-"}
                    p2 = prelim_pts[i * 2 + 1] if i * 2 + 1 < len(prelim_pts) else {"Nom": f"Lutteur {i*2+2}", "Club": "-"}
                    prelim_matches.append((p1, p2))
                    prelim_winners.append({
                        "Nom": f"Vainqueur 1/16 ({i+1}) [{cat_poids}]",
                        "Club": "Qualifié",
                        "Comité": "-"
                    })
                rondes.append(prelim_matches)
                slots_16 = (byes_16 + prelim_winners)[:16]
                while len(slots_16) < 16:
                    slots_16.append({"Nom": f"Qualifié 1/8 ({len(slots_16)+1})", "Club": "Qualifié", "Comité": "-"})
            
            matches_18 = []
            winners_18 = []
            for i in range(8):
                p1 = slots_16[i * 2]
                p2 = slots_16[i * 2 + 1]
                matches_18.append((p1, p2))
                winners_18.append({
                    "Nom": f"Vainqueur 1/8 ({i+1}) [{cat_poids}]",
                    "Club": "Qualifié",
                    "Comité": "-"
                })
            rondes.append(matches_18)
            
            qf_matches = [
                (winners_18[0], winners_18[1]),
                (winners_18[2], winners_18[3]),
                (winners_18[4], winners_18[5]),
                (winners_18[6], winners_18[7])
            ]
            rondes.append(qf_matches)
            
            sf1 = (
                {"Nom": f"Vainqueur 1/4 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"},
                {"Nom": f"Vainqueur 1/4 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            sf2 = (
                {"Nom": f"Vainqueur 1/4 (3) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"},
                {"Nom": f"Vainqueur 1/4 (4) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            rep1 = (
                {"Nom": f"Perdant 1/4 (1) [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/4 (2) [{cat_poids}]", "Club": "Repêché", "Comité": "-"}
            )
            rep2 = (
                {"Nom": f"Perdant 1/4 (3) [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/4 (4) [{cat_poids}]", "Club": "Repêché", "Comité": "-"}
            )
            rondes.append([sf1, sf2, rep1, rep2])
            
            f_or = (
                {"Nom": f"Vainqueur 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"},
                {"Nom": f"Vainqueur 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            f_b1 = (
                {"Nom": f"Vainqueur Repêchage 1 [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            f_b2 = (
                {"Nom": f"Vainqueur Repêchage 2 [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            rondes.append([f_or, f_b1, f_b2])

        else:
            # 33 <= n <= 64 (Tableau de 64 : 1/32, 1/16, 1/8, 1/4, 1/2, Finales & Repêchages)
            nb_prelim = n - 32
            nb_byes = max(0, 64 - n)
            
            if separer_clubs:
                byes_32, prelim_matches = repartir_tableau_protection_clubs(participants, nb_byes, nb_prelim)
                prelim_winners = []
                for i in range(nb_prelim):
                    prelim_winners.append({
                        "Nom": f"Vainqueur 1/32 ({i+1}) [{cat_poids}]",
                        "Club": "Qualifié",
                        "Comité": "-"
                    })
                rondes.append(prelim_matches)
                
                slots_32 = [None] * 32
                order_prelim_slots_32 = [31, 15, 23, 7, 27, 11, 19, 3, 29, 13, 21, 5, 25, 9, 17, 1, 30, 14, 22, 6, 26, 10, 18, 2, 28, 12, 20, 4, 24, 8, 16, 0]
                for w_idx, slot_idx in enumerate(order_prelim_slots_32[:nb_prelim]):
                    slots_32[slot_idx] = prelim_winners[w_idx]
                    
                order_upper = [s for s in [0, 8, 4, 12, 2, 10, 6, 14, 1, 9, 5, 13, 3, 11, 7, 15] if slots_32[s] is None]
                order_lower = [s for s in [16, 24, 20, 28, 18, 26, 22, 30, 17, 25, 21, 29, 19, 27, 23, 31] if slots_32[s] is None]
                interleaved_bye_slots = interleave_bracket_slots(order_upper, order_lower)
                for b_idx, s_idx in enumerate(interleaved_bye_slots):
                    if b_idx < len(byes_32):
                        slots_32[s_idx] = byes_32[b_idx]
                        
                for s_idx in range(32):
                    if slots_32[s_idx] is None:
                        slots_32[s_idx] = {"Nom": f"Qualifié 1/16 ({s_idx+1})", "Club": "Qualifié", "Comité": "-"}
            else:
                byes_32 = participants[:nb_byes]
                prelim_pts = participants[nb_byes:]
                prelim_matches = []
                prelim_winners = []
                for i in range(nb_prelim):
                    p1 = prelim_pts[i * 2] if i * 2 < len(prelim_pts) else {"Nom": f"Lutteur {i*2+1}", "Club": "-"}
                    p2 = prelim_pts[i * 2 + 1] if i * 2 + 1 < len(prelim_pts) else {"Nom": f"Lutteur {i*2+2}", "Club": "-"}
                    prelim_matches.append((p1, p2))
                    prelim_winners.append({
                        "Nom": f"Vainqueur 1/32 ({i+1}) [{cat_poids}]",
                        "Club": "Qualifié",
                        "Comité": "-"
                    })
                rondes.append(prelim_matches)
                slots_32 = (byes_32 + prelim_winners)[:32]
                while len(slots_32) < 32:
                    slots_32.append({"Nom": f"Qualifié 1/16 ({len(slots_32)+1})", "Club": "Qualifié", "Comité": "-"})
            
            matches_116 = []
            winners_116 = []
            for i in range(16):
                p1 = slots_32[i * 2]
                p2 = slots_32[i * 2 + 1]
                matches_116.append((p1, p2))
                winners_116.append({
                    "Nom": f"Vainqueur 1/16 ({i+1}) [{cat_poids}]",
                    "Club": "Qualifié",
                    "Comité": "-"
                })
            rondes.append(matches_116)
            
            matches_18 = []
            winners_18 = []
            for i in range(8):
                p1 = winners_116[i * 2]
                p2 = winners_116[i * 2 + 1]
                matches_18.append((p1, p2))
                winners_18.append({
                    "Nom": f"Vainqueur 1/8 ({i+1}) [{cat_poids}]",
                    "Club": "Qualifié",
                    "Comité": "-"
                })
            rondes.append(matches_18)
            
            qf_matches = [
                (winners_18[0], winners_18[1]),
                (winners_18[2], winners_18[3]),
                (winners_18[4], winners_18[5]),
                (winners_18[6], winners_18[7])
            ]
            rondes.append(qf_matches)
            
            sf1 = (
                {"Nom": f"Vainqueur 1/4 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"},
                {"Nom": f"Vainqueur 1/4 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            sf2 = (
                {"Nom": f"Vainqueur 1/4 (3) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"},
                {"Nom": f"Vainqueur 1/4 (4) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            rep1 = (
                {"Nom": f"Perdant 1/4 (1) [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/4 (2) [{cat_poids}]", "Club": "Repêché", "Comité": "-"}
            )
            rep2 = (
                {"Nom": f"Perdant 1/4 (3) [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/4 (4) [{cat_poids}]", "Club": "Repêché", "Comité": "-"}
            )
            rondes.append([sf1, sf2, rep1, rep2])
            
            f_or = (
                {"Nom": f"Vainqueur 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"},
                {"Nom": f"Vainqueur 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            f_b1 = (
                {"Nom": f"Vainqueur Repêchage 1 [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            f_b2 = (
                {"Nom": f"Vainqueur Repêchage 2 [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            rondes.append([f_or, f_b1, f_b2])

        nom = f"{age} | {style_grp}{suffixe_niveau} | {cat_poids} (Tableau élimination & repêchages)"
        res = {
            'nom': nom,
            'participants': list(participants),
            'rondes': rondes,
            'type_formule': 'tableau',
            'cat_poids': cat_poids,
            'style_grp': style_grp
        }
        if n == 7:
            res['p_exempt'] = p_exempt
        return res

# --- GÉNÉRATEUR VISUEL DE TABLEAU À ÉLIMINATION DIRECTE & REPÊCHAGES U13 (DE GAUCHE À DROITE) ---
def make_bracket_card(p1, p2, title='', badge=''):
    nom1 = p1.get('Nom', 'Lutteur 1')
    c1 = p1.get('Club', '')
    club1_str = f" ({c1})" if c1 and c1 not in ['-', 'Comité Non Renseigné', ''] else ''
    
    nom2 = p2.get('Nom', 'Lutteur 2')
    c2 = p2.get('Club', '')
    club2_str = f" ({c2})" if c2 and c2 not in ['-', 'Comité Non Renseigné', ''] else ''

    badge_html = f'<span style="background: #FEF3C7; color: #92400E; border: 1px solid #FCD34D; padding: 1px 5px; border-radius: 4px; font-size: 9px; font-weight: 700;">{badge}</span>' if badge else ''

    is_bye1 = any(k in str(nom1).lower() for k in ['bye', 'exempt'])
    is_bye2 = any(k in str(nom2).lower() for k in ['bye', 'exempt'])

    dot1 = '#94A3B8' if is_bye1 else '#EF4135'
    dot2 = '#94A3B8' if is_bye2 else '#0055A4'

    return f'''
    <div style="background: #FFFFFF; border: 1.5px solid #CBD5E1; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); width: 230px; margin: 6px 0; overflow: hidden; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; text-align: left;">
        <div style="background: #F8FAFC; border-bottom: 1px solid #E2E8F0; padding: 4px 8px; font-size: 10px; font-weight: 700; color: #475569; display: flex; justify-content: space-between; align-items: center;">
            <span>{title}</span>
            {badge_html}
        </div>
        <div style="padding: 5px 8px; display: flex; align-items: center; border-bottom: 1px solid #F1F5F9; background: {'#F8FAFC' if is_bye1 else '#FFFFFF'};">
            <span style="width: 9px; height: 9px; border-radius: 50%; background: {dot1}; display: inline-block; margin-right: 6px; flex-shrink: 0;"></span>
            <div style="flex: 1; overflow: hidden;">
                <div style="font-size: 11px; font-weight: 700; color: {'#64748B' if is_bye1 else '#1E293B'}; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{nom1}</div>
                <div style="font-size: 9px; color: #94A3B8; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{club1_str}</div>
            </div>
            <div style="font-size: 9px; font-weight: 700; background: #F1F5F9; border: 1px solid #CBD5E1; border-radius: 3px; padding: 1px 4px; color: #475569;">[ &nbsp; ]</div>
        </div>
        <div style="padding: 5px 8px; display: flex; align-items: center; background: {'#F8FAFC' if is_bye2 else '#FFFFFF'};">
            <span style="width: 9px; height: 9px; border-radius: 50%; background: {dot2}; display: inline-block; margin-right: 6px; flex-shrink: 0;"></span>
            <div style="flex: 1; overflow: hidden;">
                <div style="font-size: 11px; font-weight: 700; color: {'#64748B' if is_bye2 else '#1E293B'}; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{nom2}</div>
                <div style="font-size: 9px; color: #94A3B8; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{club2_str}</div>
            </div>
            <div style="font-size: 9px; font-weight: 700; background: #F1F5F9; border: 1px solid #CBD5E1; border-radius: 3px; padding: 1px 4px; color: #475569;">[ &nbsp; ]</div>
        </div>
    </div>
    '''

def render_svg_connectors(count, height, width=40):
    lines = []
    pairs = count // 2
    step = height / count
    for p in range(pairs):
        y_top = (p * 2 + 0.5) * step
        y_bot = (p * 2 + 1.5) * step
        y_mid = (p + 0.5) * (height / pairs)
        path = f"M 0 {y_top:.1f} H {width//2} V {y_bot:.1f} H 0 M {width//2} {y_mid:.1f} H {width}"
        lines.append(f'<path d="{path}" fill="none" stroke="#94A3B8" stroke-width="2" />')
    return f'<svg width="{width}" height="{height}" style="flex-shrink: 0; display: block;">{" ".join(lines)}</svg>'

def generer_arbre_tableau_html(p_obj):
    nom_poule = p_obj.get('nom', 'U13')
    rondes = p_obj.get('rondes', [])
    participants = p_obj.get('participants', [])
    type_formule = p_obj.get('type_formule', '')
    n = len(participants)

    if type_formule == 'poules_croisees':
        sf1, sf2 = rondes[3][0], rondes[3][1]
        f_or, f_b = rondes[4][0], rondes[4][1]
        
        c_sf1 = make_bracket_card(sf1[0], sf1[1], 'DEMI-FINALE 1', '1er A vs 2ème B')
        c_sf2 = make_bracket_card(sf2[0], sf2[1], 'DEMI-FINALE 2', '1er B vs 2ème A')
        c_for = make_bracket_card(f_or[0], f_or[1], 'FINALE OR / ARGENT', '🥇 Or / 🥈 Argent')
        c_fb = make_bracket_card(f_b[0], f_b[1], 'FINALE BRONZE (3-4)', '🥉 Bronze unique')

        h_col = 220
        svg_conn = f'<svg width="40" height="{h_col}" style="display: block; flex-shrink: 0;"><path d="M 0 55 H 20 V 165 H 0 M 20 55 H 40 M 20 165 H 40" fill="none" stroke="#94A3B8" stroke-width="2" /></svg>'

        return f'''
        <div style="background: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 18px; margin: 15px 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;">
            <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px;">
                <span style="background: #0055A4; color: white; padding: 5px 14px; border-radius: 16px; font-size: 12px; font-weight: 800; text-transform: uppercase;">
                    🏆 Tableau Final Croisé U13 (6 Lutteurs)
                </span>
                <span style="font-size: 11px; color: #64748B; font-weight: 600;">Orientation : Gauche ➔ Droite (Demi-Finales ➔ Finales)</span>
            </div>
            
            <div style="overflow-x: auto; padding-bottom: 10px;">
                <div style="display: inline-flex; flex-direction: row; align-items: center; gap: 0;">
                    <!-- Demi-Finales -->
                    <div style="display: flex; flex-direction: column; width: 240px;">
                        <div style="background: #0055A4; color: white; padding: 6px; border-radius: 6px; font-size: 11px; font-weight: 700; text-align: center; margin-bottom: 10px;">
                            DEMI-FINALES CROISÉES
                        </div>
                        <div style="display: flex; flex-direction: column; justify-content: space-around; height: {h_col}px;">
                            {c_sf1}
                            {c_sf2}
                        </div>
                    </div>

                    {svg_conn}

                    <!-- Finales -->
                    <div style="display: flex; flex-direction: column; width: 240px;">
                        <div style="background: #0F172A; color: white; padding: 6px; border-radius: 6px; font-size: 11px; font-weight: 700; text-align: center; margin-bottom: 10px;">
                            FINALES
                        </div>
                        <div style="display: flex; flex-direction: column; justify-content: space-around; height: {h_col}px;">
                            {c_for}
                            {c_fb}
                        </div>
                    </div>

                    <!-- Podiums -->
                    <div style="margin-left: 15px; display: flex; flex-direction: column; justify-content: space-around; height: {h_col}px;">
                        <div style="background: #FEF3C7; border: 1.5px solid #F59E0B; border-radius: 8px; padding: 8px 12px; font-size: 11px;">
                            <div style="font-weight: 800; color: #B45309;">🥇 OR : Vainqueur Finale</div>
                            <div style="font-weight: 800; color: #475569; margin-top: 4px;">🥈 ARGENT : Perdant Finale</div>
                        </div>
                        <div style="background: #EFF6FF; border: 1.5px solid #3B82F6; border-radius: 8px; padding: 8px 12px; font-size: 11px;">
                            <div style="font-weight: 800; color: #1D4ED8;">🥉 BRONZE (Unique) : Vainqueur 3-4</div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        '''

    elif type_formule == 'tableau':
        f_or = rondes[-1][0]
        f_b1 = rondes[-1][1] if len(rondes[-1]) > 1 else None
        f_b2 = rondes[-1][2] if len(rondes[-1]) > 2 else None

        sf1 = rondes[-2][0]
        sf2 = rondes[-2][1]
        rep1 = rondes[-2][2] if len(rondes[-2]) > 2 else None
        rep2 = rondes[-2][3] if len(rondes[-2]) > 3 else None

        main_columns = []
        if len(rondes) == 3:
            if n == 7:
                q1, q2, q3 = rondes[0][0], rondes[0][1], rondes[0][2]
                p_ex = p_obj.get('p_exempt', participants[6])
                q_matches = [
                    (q1[0], q1[1], '1/4 DE FINALE 1'),
                    (q2[0], q2[1], '1/4 DE FINALE 2'),
                    (q3[0], q3[1], '1/4 DE FINALE 3'),
                    (p_ex, {'Nom': 'EXEMPT (BYE)', 'Club': '-'}, '1/4 DE FINALE 4 (Exempt)')
                ]
            else:
                q_matches = [(m[0], m[1], f'1/4 DE FINALE {i+1}') for i, m in enumerate(rondes[0])]
            main_columns.append(('QUARTS DE FINALE', q_matches))

        elif len(rondes) == 4:
            prelim_m = [(m[0], m[1], f'PRÉLIMINAIRE {i+1}') for i, m in enumerate(rondes[0])]
            q_matches = [(m[0], m[1], f'1/4 DE FINALE {i+1}') for i, m in enumerate(rondes[1])]
            main_columns.append(('TOUR PRÉLIMINAIRE', prelim_m))
            main_columns.append(('QUARTS DE FINALE', q_matches))

        elif len(rondes) >= 6:
            p32 = [(m[0], m[1], f'1/32 DE FINALE {i+1}') for i, m in enumerate(rondes[0])]
            p16 = [(m[0], m[1], f'1/16 DE FINALE {i+1}') for i, m in enumerate(rondes[1])]
            p18 = [(m[0], m[1], f'1/8 DE FINALE {i+1}') for i, m in enumerate(rondes[2])]
            q_matches = [(m[0], m[1], f'1/4 DE FINALE {i+1}') for i, m in enumerate(rondes[3])]
            main_columns.append(('1/32 DE FINALE', p32))
            main_columns.append(('1/16 DE FINALE', p16))
            main_columns.append(('1/8 DE FINALE', p18))
            main_columns.append(('QUARTS DE FINALE', q_matches))

        elif len(rondes) == 5:
            p16 = [(m[0], m[1], f'1/16 DE FINALE {i+1}') for i, m in enumerate(rondes[0])]
            p18 = [(m[0], m[1], f'1/8 DE FINALE {i+1}') for i, m in enumerate(rondes[1])]
            q_matches = [(m[0], m[1], f'1/4 DE FINALE {i+1}') for i, m in enumerate(rondes[2])]
            main_columns.append(('1/16 DE FINALE', p16))
            main_columns.append(('1/8 DE FINALE', p18))
            main_columns.append(('QUARTS DE FINALE', q_matches))

        main_columns.append(('DEMI-FINALES', [(sf1[0], sf1[1], 'DEMI-FINALE 1'), (sf2[0], sf2[1], 'DEMI-FINALE 2')]))
        main_columns.append(('FINALE (OR / ARGENT)', [(f_or[0], f_or[1], 'GRANDE FINALE')]))

        max_matches_in_col = max(len(col[1]) for col in main_columns)
        base_h = max(max_matches_in_col * 110, 440)

        main_cols_html = []
        for c_idx, (col_title, matches_list) in enumerate(main_columns):
            cards_html = []
            for m in matches_list:
                cards_html.append(make_bracket_card(m[0], m[1], m[2]))
            
            bg_h = '#0F172A' if 'FINALE' in col_title and 'DEMI' not in col_title else '#0055A4'

            col_div = f'''
            <div style="display: inline-flex; flex-direction: column; width: 240px; margin: 0 5px;">
                <div style="background: {bg_h}; color: white; padding: 6px; border-radius: 6px; font-size: 11px; font-weight: 700; text-align: center; margin-bottom: 10px;">
                    {col_title}
                </div>
                <div style="display: flex; flex-direction: column; justify-content: space-around; height: {base_h}px;">
                    {''.join(cards_html)}
                </div>
            </div>
            '''
            main_cols_html.append(col_div)

            if c_idx < len(main_columns) - 1:
                next_count = len(main_columns[c_idx + 1][1])
                curr_count = len(matches_list)
                if curr_count == next_count * 2:
                    svg_c = render_svg_connectors(curr_count, base_h, width=40)
                else:
                    lines = [f'<line x1="0" y1="{base_h * (i + 0.5) / curr_count:.1f}" x2="40" y2="{base_h * (i + 0.5) / curr_count:.1f}" stroke="#94A3B8" stroke-width="2" />' for i in range(curr_count)]
                    svg_c = f'<svg width="40" height="{base_h}" style="flex-shrink: 0; display: block;">{" ".join(lines)}</svg>'
                main_cols_html.append(svg_c)

        podium_main = f'''
        <div style="display: inline-flex; align-items: center; margin-left: 10px;">
            <div style="background: #FFFDF5; border: 1.5px solid #F59E0B; border-radius: 8px; padding: 10px 14px; box-shadow: 0 2px 6px rgba(245,158,11,0.15); font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;">
                <div style="font-size: 12px; font-weight: 800; color: #B45309; margin-bottom: 6px;">🥇 CHAMPION (OR)</div>
                <div style="font-size: 11px; color: #1E293B;">Vainqueur Grande Finale</div>
                <div style="border-top: 1px solid #FDE68A; margin: 6px 0;"></div>
                <div style="font-size: 12px; font-weight: 800; color: #64748B; margin-bottom: 6px;">🥈 VICE-CHAMPION (ARGENT)</div>
                <div style="font-size: 11px; color: #1E293B;">Finaliste Grande Finale</div>
            </div>
        </div>
        '''
        main_cols_html.append(podium_main)

        rep_cards = []
        if rep1: rep_cards.append(make_bracket_card(rep1[0], rep1[1], 'REPÊCHAGE 1 (1/4)', 'Perdants QF 1-2'))
        if rep2: rep_cards.append(make_bracket_card(rep2[0], rep2[1], 'REPÊCHAGE 2 (1/4)', 'Perdants QF 3-4'))
        elif n == 7: rep_cards.append('<div style="background: #F1F5F9; border: 1.5px dashed #CBD5E1; border-radius: 8px; padding: 12px; font-size: 11px; color: #64748B; text-align: center; width: 230px;">Exempt de 1er repêchage (Avance direct en Finale Bronze 2)</div>')

        bronze_cards = []
        if f_b1: bronze_cards.append(make_bracket_card(f_b1[0], f_b1[1], 'FINALE BRONZE 1', '🥉 Médaille de Bronze 1'))
        if f_b2: bronze_cards.append(make_bracket_card(f_b2[0], f_b2[1], 'FINALE BRONZE 2', '🥉 Médaille de Bronze 2'))

        h_rep = 220
        svg_rep = f'''
        <svg width="40" height="{h_rep}" style="flex-shrink: 0; display: block;">
            <line x1="0" y1="55" x2="40" y2="55" stroke="#94A3B8" stroke-width="2" />
            <line x1="0" y1="165" x2="40" y2="165" stroke="#94A3B8" stroke-width="2" />
        </svg>
        '''

        rep_html = f'''
        <div style="margin-top: 24px; padding-top: 18px; border-top: 2px dashed #CBD5E1;">
            <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 14px;">
                <span style="background: #0284C7; color: white; padding: 5px 14px; border-radius: 16px; font-size: 12px; font-weight: 800; text-transform: uppercase;">
                    🔄 Repêchages & Attribution du Bronze (2 Troisièmes Places)
                </span>
                <span style="font-size: 11px; color: #64748B; font-weight: 600;">Perdants des 1/4 ➔ Repêchages ➔ Finales Bronze contre Perdants des 1/2 de leur propre demi-tableau</span>
            </div>

            <div style="overflow-x: auto; padding-bottom: 10px;">
                <div style="display: inline-flex; flex-direction: row; align-items: center; gap: 0;">
                    <div style="display: flex; flex-direction: column; width: 240px; margin: 0 5px;">
                        <div style="background: #0284C7; color: white; padding: 6px; border-radius: 6px; font-size: 11px; font-weight: 700; text-align: center; margin-bottom: 10px;">
                            REPÊCHAGES (1/4)
                        </div>
                        <div style="display: flex; flex-direction: column; justify-content: space-around; height: {h_rep}px;">
                            {''.join(rep_cards)}
                        </div>
                    </div>

                    {svg_rep}

                    <div style="display: flex; flex-direction: column; width: 240px; margin: 0 5px;">
                        <div style="background: #B45309; color: white; padding: 6px; border-radius: 6px; font-size: 11px; font-weight: 700; text-align: center; margin-bottom: 10px;">
                            MATCHS POUR LE BRONZE
                        </div>
                        <div style="display: flex; flex-direction: column; justify-content: space-around; height: {h_rep}px;">
                            {''.join(bronze_cards)}
                        </div>
                    </div>

                    <div style="margin-left: 15px; display: flex; flex-direction: column; justify-content: space-around; height: {h_rep}px;">
                        <div style="background: #EFF6FF; border: 1.5px solid #3B82F6; border-radius: 8px; padding: 8px 12px; font-size: 11px;">
                            <div style="font-weight: 800; color: #1D4ED8;">🥉 3ème PLACE (Bronze 1)</div>
                            <div style="color: #1E293B; margin-top: 2px;">Vainqueur Finale Bronze 1</div>
                        </div>
                        <div style="background: #EFF6FF; border: 1.5px solid #3B82F6; border-radius: 8px; padding: 8px 12px; font-size: 11px;">
                            <div style="font-weight: 800; color: #1D4ED8;">🥉 3ème PLACE (Bronze 2)</div>
                            <div style="color: #1E293B; margin-top: 2px;">Vainqueur Finale Bronze 2</div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        '''

        return f'''
        <div style="background: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 18px; margin: 15px 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;">
            <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px;">
                <span style="background: #0055A4; color: white; padding: 5px 14px; border-radius: 16px; font-size: 12px; font-weight: 800; text-transform: uppercase;">
                    🏆 Tableau Principal à Élimination Directe — {nom_poule}
                </span>
                <span style="font-size: 11px; color: #64748B; font-weight: 600;">Orientation : Gauche ➔ Droite (Éliminatoires ➔ Quarts ➔ Demi-Finales ➔ Finale)</span>
            </div>

            <div style="overflow-x: auto; padding-bottom: 10px;">
                <div style="display: inline-flex; flex-direction: row; align-items: center; gap: 0;">
                    {''.join(main_cols_html)}
                </div>
            </div>

            {rep_html}
        </div>
        '''
    return ''

def generer_document_bracket_imprimable(nom_poule, nom_comp, bracket_html):
    return f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Tableau U13 - {nom_poule} - {nom_comp}</title>
    <style>
        @page {{
            size: landscape;
            margin: 8mm;
        }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            margin: 0;
            padding: 10px;
            background: white;
            color: #1E293B;
        }}
        .print-btn {{
            background-color: #0055A4;
            color: white;
            border: none;
            padding: 8px 16px;
            font-size: 13px;
            font-weight: bold;
            border-radius: 6px;
            cursor: pointer;
            margin-bottom: 12px;
        }}
        @media print {{
            .print-btn {{
                display: none !important;
            }}
            body {{
                padding: 0;
            }}
        }}
    </style>
</head>
<body>
    <button class="print-btn" onclick="window.print()">🖨️ Imprimer ce Tableau (A4 Paysage)</button>
    <div style="margin-bottom: 10px;">
        <h2 style="color: #0055A4; margin: 0 0 4px 0; font-size: 20px;">🏆 {nom_comp.upper()}</h2>
        <div style="font-size: 12px; color: #64748B; font-weight: bold;">FFLDA — TABLEAU OFFICIEL À ÉLIMINATION DIRECTE & REPÊCHAGES</div>
    </div>
    {bracket_html}
</body>
</html>"""

def generer_document_poule_imprimable(nom_poule, nom_comp, participants, rondes):
    nb_tours = len(rondes) if rondes else 0
    
    lutteurs_par_tour = []
    for ronde in (rondes or []):
        combatants = set()
        for match in ronde:
            p1_nom = match[0].get('Nom', '') if isinstance(match[0], dict) else str(match[0])
            p2_nom = match[1].get('Nom', '') if isinstance(match[1], dict) else str(match[1])
            combatants.add(p1_nom)
            combatants.add(p2_nom)
        lutteurs_par_tour.append(combatants)

    lignes_html = []
    for idx, p in enumerate(participants, 1):
        bg = "#F8FAFC" if idx % 2 == 0 else "#FFFFFF"
        nom = p.get('Nom', '')
        club = p.get('Club', '')
        comite = p.get('Comité', '')
        poids = formater_poids(p.get('Poids', ''))
        
        tours_tds_list = []
        for t_idx in range(nb_tours):
            if t_idx < len(lutteurs_par_tour) and nom in lutteurs_par_tour[t_idx]:
                tours_tds_list.append('<td style="border: 1px solid #CBD5E1; padding: 6px; text-align: center; background: #E2E8F0;"></td>')
            else:
                tours_tds_list.append('<td style="border: 1px solid #CBD5E1; padding: 6px; text-align: center;"></td>')
        tours_tds = ''.join(tours_tds_list)
        
        lignes_html.append(f"""
        <tr style="background: {bg};">
            <td style="border: 1px solid #CBD5E1; padding: 6px; text-align: center; font-weight: bold; color: #0055A4;"></td>
            <td style="border: 1px solid #CBD5E1; padding: 6px; text-align: center; font-weight: bold;">{idx}</td>
            <td style="border: 1px solid #CBD5E1; padding: 6px; font-weight: 600;">{nom}</td>
            <td style="border: 1px solid #CBD5E1; padding: 6px;">{club}</td>
            <td style="border: 1px solid #CBD5E1; padding: 6px;">{comite}</td>
            {tours_tds}
            <td style="border: 1px solid #CBD5E1; padding: 6px; text-align: center; font-weight: bold;"></td>
            <td style="border: 1px solid #CBD5E1; padding: 6px; text-align: center;"></td>
            <td style="border: 1px solid #CBD5E1; padding: 6px; text-align: center;">{poids}</td>
        </tr>
        """)

    headers_tours = ''.join([f'<th style="border: 1px solid #CBD5E1; padding: 8px; font-size: 11px;">Tour {t}</th>' for t in range(1, nb_tours + 1)])
    
    n_p = len(participants)
    podium_html = f"""
    <div style="display: flex; gap: 12px; margin-top: 15px; margin-bottom: 20px;">
        <div style="flex: 1; background: #FEF3C7; border: 1.5px solid #F59E0B; border-radius: 8px; padding: 10px; text-align: center;">
            <div style="font-weight: 800; font-size: 13px; color: #B45309;">🥇 1ère PLACE (OR)</div>
            <div style="font-size: 12px; font-weight: 600; color: #78350F; margin-top: 4px;">En attente</div>
        </div>
        <div style="flex: 1; background: #F1F5F9; border: 1.5px solid #94A3B8; border-radius: 8px; padding: 10px; text-align: center;">
            <div style="font-weight: 800; font-size: 13px; color: #475569;">🥈 2ème PLACE (ARGENT)</div>
            <div style="font-size: 12px; font-weight: 600; color: #1E293B; margin-top: 4px;">En attente</div>
        </div>
    """
    if n_p >= 3:
        podium_html += """
        <div style="flex: 1; background: #FFEDD5; border: 1.5px solid #F97316; border-radius: 8px; padding: 10px; text-align: center;">
            <div style="font-weight: 800; font-size: 13px; color: #9A3412;">🥉 3ème PLACE (BRONZE)</div>
            <div style="font-size: 12px; font-weight: 600; color: #7C2D12; margin-top: 4px;">En attente</div>
        </div>
        """
    podium_html += "</div>"
    
    matchs_html = []
    m_count = 1
    for tour_idx, ronde in enumerate(rondes, 1):
        tour_cards = []
        for m in ronde:
            p1, p2 = m[0], m[1]
            nom1 = p1.get('Nom', '')
            club1 = f" ({p1.get('Club', '')})" if p1.get('Club') and p1.get('Club') != '-' else ''
            nom2 = p2.get('Nom', '')
            club2 = f" ({p2.get('Club', '')})" if p2.get('Club') and p2.get('Club') != '-' else ''
            
            tour_cards.append(f"""
            <div style="background: white; border: 1px solid #CBD5E1; border-radius: 6px; margin-bottom: 8px; overflow: hidden; width: 100%;">
                <div style="background: #334155; color: white; font-size: 10px; font-weight: bold; padding: 3px 8px; text-align: center;">
                    COMBAT N°{m_count}
                </div>
                <div style="display: flex; justify-content: space-between; align-items: center; background: #FEF2F2; padding: 5px 8px; border-bottom: 1px solid #CBD5E1;">
                    <span style="font-size: 11px; font-weight: 700; color: #991B1B;">🔴 {nom1}{club1}</span>
                    <span style="border: 1px solid #CBD5E1; background: white; padding: 2px 8px; font-weight: bold; border-radius: 4px; font-size: 11px;">&nbsp;&nbsp;&nbsp;</span>
                </div>
                <div style="display: flex; justify-content: space-between; align-items: center; background: #EFF6FF; padding: 5px 8px;">
                    <span style="font-size: 11px; font-weight: 700; color: #1E40AF;">🔵 {nom2}{club2}</span>
                    <span style="border: 1px solid #CBD5E1; background: white; padding: 2px 8px; font-weight: bold; border-radius: 4px; font-size: 11px;">&nbsp;&nbsp;&nbsp;</span>
                </div>
            </div>
            """)
            m_count += 1
            
        matchs_html.append(f"""
        <div style="margin-bottom: 14px;">
            <div style="background: #0055A4; color: white; font-size: 11px; font-weight: bold; padding: 4px 10px; border-radius: 4px; margin-bottom: 6px;">
                🎯 TOUR {tour_idx}
            </div>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 10px;">
                {''.join(tour_cards)}
            </div>
        </div>
        """)

    return f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Poule Officielle - {nom_poule} - {nom_comp}</title>
    <style>
        @page {{
            size: landscape;
            margin: 8mm;
        }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            margin: 0;
            padding: 10px;
            background: white;
            color: #1E293B;
        }}
        .print-btn {{
            background-color: #0055A4;
            color: white;
            border: none;
            padding: 8px 16px;
            font-size: 13px;
            font-weight: bold;
            border-radius: 6px;
            cursor: pointer;
            margin-bottom: 12px;
        }}
        @media print {{
            .print-btn {{
                display: none !important;
            }}
            body {{
                padding: 0;
            }}
        }}
    </style>
</head>
<body>
    <button class="print-btn" onclick="window.print()">🖨️ Imprimer cette Poule (A4 Paysage)</button>
    <div style="margin-bottom: 12px;">
        <h2 style="color: #0055A4; margin: 0 0 4px 0; font-size: 20px;">🏆 {nom_comp.upper()}</h2>
        <div style="font-size: 12px; color: #64748B; font-weight: bold;">FFLDA — POULE OFFICIELLE : {nom_poule} (TOUS CONTRE TOUS)</div>
    </div>
    
    <table style="width: 100%; border-collapse: collapse; font-size: 11px; margin-bottom: 10px;">
        <thead>
            <tr style="background: #0055A4; color: white;">
                <th style="border: 1px solid #CBD5E1; padding: 8px; width: 45px;">CLT</th>
                <th style="border: 1px solid #CBD5E1; padding: 8px; width: 35px;">N°</th>
                <th style="border: 1px solid #CBD5E1; padding: 8px; text-align: left;">NOM Prénom</th>
                <th style="border: 1px solid #CBD5E1; padding: 8px; text-align: left;">CLUB</th>
                <th style="border: 1px solid #CBD5E1; padding: 8px; text-align: left;">COMITÉ</th>
                {headers_tours}
                <th style="border: 1px solid #CBD5E1; padding: 8px; width: 65px;">Total Pts</th>
                <th style="border: 1px solid #CBD5E1; padding: 8px; width: 65px;">Total Vict</th>
                <th style="border: 1px solid #CBD5E1; padding: 8px; width: 55px;">Poids</th>
            </tr>
        </thead>
        <tbody>
            {''.join(lignes_html)}
        </tbody>
    </table>
    
    {podium_html}
    
    <div style="margin-top: 15px;">
        {''.join(matchs_html)}
    </div>
</body>
</html>"""

def generer_document_plateau_u7_imprimable(nom_poule, nom_comp, participants):
    lignes_html = []
    for idx, p in enumerate(participants, 1):
        bg = "#FFFBEB" if idx % 2 == 0 else "#FFFFFF"
        nom = p.get('Nom', '')
        club = p.get('Club', '')
        poids = formater_poids(p.get('Poids', ''))
        lignes_html.append(f"""
        <tr style="background: {bg};">
            <td style="border: 1px solid #CBD5E1; padding: 8px; text-align: center; font-weight: bold;">{idx}</td>
            <td style="border: 1px solid #CBD5E1; padding: 8px; font-weight: 600;">{nom}</td>
            <td style="border: 1px solid #CBD5E1; padding: 8px;">{club}</td>
            <td style="border: 1px solid #CBD5E1; padding: 8px; text-align: center;">{poids}</td>
            <td style="border: 1px solid #CBD5E1; padding: 8px; text-align: center; color: #166534; font-weight: bold;">[ ✓ ] Validé</td>
            <td style="border: 1px solid #CBD5E1; padding: 8px; text-align: center; color: #166534; font-weight: bold;">[ ✓ ] Validé</td>
            <td style="border: 1px solid #CBD5E1; padding: 8px; text-align: center; color: #166534; font-weight: bold;">[ ✓ ] Validé</td>
            <td style="border: 1px solid #CBD5E1; padding: 8px; text-align: center; background: #FEF3C7; color: #92400E; font-weight: 800;">🥇 Médaille d'Or</td>
        </tr>
        """)

    return f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Passeport U7 - {nom_poule} - {nom_comp}</title>
    <style>
        @page {{ size: landscape; margin: 8mm; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; margin: 0; padding: 10px; background: white; color: #1E293B; }}
        .print-btn {{ background-color: #D97706; color: white; border: none; padding: 8px 16px; font-size: 13px; font-weight: bold; border-radius: 6px; cursor: pointer; margin-bottom: 12px; }}
        @media print {{ .print-btn {{ display: none !important; }} body {{ padding: 0; }} }}
    </style>
</head>
<body>
    <button class="print-btn" onclick="window.print()">🖨️ Imprimer cette Fiche Plateau U7 (A4 Paysage)</button>
    <div style="margin-bottom: 12px;">
        <h2 style="color: #D97706; margin: 0 0 4px 0; font-size: 20px;">🏆 {nom_comp.upper()}</h2>
        <div style="font-size: 13px; color: #64748B; font-weight: bold;">FFLDA — ANIMATION PLATEAU U7 : {nom_poule}</div>
        <div style="font-size: 11px; color: #92400E; font-style: italic;">Formule officielle FFLDA : Découverte pédagogique sous forme de 3 plateaux d'activités avec rotation. Tous les enfants sont récompensés !</div>
    </div>
    
    <table style="width: 100%; border-collapse: collapse; font-size: 12px; margin-bottom: 15px;">
        <thead>
            <tr style="background: #D97706; color: white;">
                <th style="border: 1px solid #CBD5E1; padding: 8px; width: 40px;">N°</th>
                <th style="border: 1px solid #CBD5E1; padding: 8px; text-align: left;">NOM Prénom</th>
                <th style="border: 1px solid #CBD5E1; padding: 8px; text-align: left;">CLUB</th>
                <th style="border: 1px solid #CBD5E1; padding: 8px; width: 70px;">Poids</th>
                <th style="border: 1px solid #CBD5E1; padding: 8px;">Plateau 1 : Motricité & Agilité</th>
                <th style="border: 1px solid #CBD5E1; padding: 8px;">Plateau 2 : Ateliers techniques</th>
                <th style="border: 1px solid #CBD5E1; padding: 8px;">Plateau 3 : Oppositions</th>
                <th style="border: 1px solid #CBD5E1; padding: 8px; width: 150px;">Validation / Récompense</th>
            </tr>
        </thead>
        <tbody>
            {''.join(lignes_html)}
        </tbody>
    </table>
    
    <div style="background: #FFFBEB; border: 1.5px solid #F59E0B; border-radius: 8px; padding: 12px; font-size: 11px;">
        <div style="font-weight: 800; font-size: 12px; color: #B45309; margin-bottom: 6px;">ℹ️ ORGANISATION DES 3 PLATEAUX D'ACTIVITÉ U7 :</div>
        <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 10px;">
            <div style="background: white; border: 1px solid #FCD34D; border-radius: 6px; padding: 8px;">
                <b style="color: #B45309;">🟡 Plateau 1 : Motricité & Agilité</b>
                <p style="margin: 4px 0 0 0; color: #78350F;">Parcours gymnique, franchissements, équilibre, réactivité motrice.</p>
            </div>
            <div style="background: white; border: 1px solid #FCD34D; border-radius: 6px; padding: 8px;">
                <b style="color: #166534;">🟢 Plateau 2 : Ateliers techniques</b>
                <p style="margin: 4px 0 0 0; color: #14532D;">Ateliers d'apprentissage technique, habiletés motrices et gestes de lutte adaptés.</p>
            </div>
            <div style="background: white; border: 1px solid #FCD34D; border-radius: 6px; padding: 8px;">
                <b style="color: #1E40AF;">🔵 Plateau 3 : Oppositions</b>
                <p style="margin: 4px 0 0 0; color: #1E3A8A;">Jeux de lutte et oppositions adaptées, combats éducatifs aménagés.</p>
            </div>
        </div>
        <div style="margin-top: 8px; text-align: center; font-weight: bold; color: #92400E;">
            🏅 Chaque enfant passe successivement sur les 3 plateaux. Tous reçoivent une médaille d'or FFLDA et un diplôme !
        </div>
    </div>
</body>
</html>"""

def make_winner_formula(c_r, c_b, pt_r, pt_b, placeholder_name, target_corner="🔴"):
    default_text = f"{target_corner} {placeholder_name}"
    r_val = f"IF({pt_r.coordinate}=\"\", 0, {pt_r.coordinate})"
    b_val = f"IF({pt_b.coordinate}=\"\", 0, {pt_b.coordinate})"
    
    if target_corner == "🔴":
        win_r = c_r.coordinate
        win_b = f'SUBSTITUTE({c_b.coordinate}, "🔵 ", "🔴 ")'
    else:
        win_r = f'SUBSTITUTE({c_r.coordinate}, "🔴 ", "🔵 ")'
        win_b = c_b.coordinate
        
    return (
        f'=IF({r_val}+{b_val}=0, "{default_text}", '
        f'IF({r_val}>{b_val}, {win_r}, '
        f'IF({b_val}>{r_val}, {win_b}, '
        f'"{default_text}")))'
    )

def make_loser_formula(c_r, c_b, pt_r, pt_b, placeholder_name, target_corner="🔴"):
    default_text = f"{target_corner} {placeholder_name}"
    r_val = f"IF({pt_r.coordinate}=\"\", 0, {pt_r.coordinate})"
    b_val = f"IF({pt_b.coordinate}=\"\", 0, {pt_b.coordinate})"
    
    if target_corner == "🔴":
        lose_r = f'SUBSTITUTE({c_b.coordinate}, "🔵 ", "🔴 ")'
        lose_b = c_r.coordinate
    else:
        lose_r = c_b.coordinate
        lose_b = f'SUBSTITUTE({c_r.coordinate}, "🔴 ", "🔵 ")'
        
    return (
        f'=IF({r_val}+{b_val}=0, "{default_text}", '
        f'IF({r_val}>{b_val}, {lose_r}, '
        f'IF({b_val}>{r_val}, {lose_b}, '
        f'"{default_text}")))'
    )

def draw_excel_match_card(ws, start_row, start_col, title, p1, p2, cat_poule, coords_map=None, bg_header=None, end_col=None, with_details=False):
    b_thin = Side(style='thin', color='CBD5E1')
    b_style = Border(left=b_thin, right=b_thin, top=b_thin, bottom=b_thin)
    font_match_h = Font(name="Arial", size=9, bold=True, color="FFFFFF")
    font_p_bold = Font(name="Arial", size=9, bold=True)
    font_pts = Font(name="Arial", size=10, bold=True)
    
    fill_header = bg_header if bg_header is not None else PatternFill("solid", fgColor="475569")
    fill_red_card = PatternFill("solid", fgColor="FEF2F2")
    fill_blue_card = PatternFill("solid", fgColor="EFF6FF")
    fill_gray_box = PatternFill("solid", fgColor="F8FAFC")

    c_ty_r, c_ty_b, c_sc_r, c_sc_b = None, None, None, None

    if with_details and end_col is not None and end_col > start_col + 3:
        col1 = start_col
        col2 = end_col
        col_sc = end_col - 2
        col_ty = end_col - 1
        col_pt = end_col
        col_nom_fin = col_sc - 1
        
        # En-tête du match
        ws.merge_cells(start_row=start_row, start_column=col1, end_row=start_row, end_column=col_nom_fin)
        c_h = ws.cell(row=start_row, column=col1, value=title)
        c_h.font, c_h.fill = font_match_h, fill_header
        c_h.alignment = Alignment(horizontal="center", vertical="center")
        
        c_h_sc = ws.cell(row=start_row, column=col_sc, value="Score")
        c_h_sc.font, c_h_sc.fill = font_match_h, fill_header
        c_h_sc.alignment = Alignment(horizontal="center", vertical="center")
        
        c_h_ty = ws.cell(row=start_row, column=col_ty, value="Type")
        c_h_ty.font, c_h_ty.fill = font_match_h, fill_header
        c_h_ty.alignment = Alignment(horizontal="center", vertical="center")
        
        c_h_pt = ws.cell(row=start_row, column=col_pt, value="Pt Clt")
        c_h_pt.font, c_h_pt.fill = font_match_h, fill_header
        c_h_pt.alignment = Alignment(horizontal="center", vertical="center")
        
        # Combattant Rouge (fusionné de col1 à col_nom_fin)
        ws.merge_cells(start_row=start_row+1, start_column=col1, end_row=start_row+1, end_column=col_nom_fin)
        # Combattant Bleu (fusionné de col1 à col_nom_fin)
        ws.merge_cells(start_row=start_row+2, start_column=col1, end_row=start_row+2, end_column=col_nom_fin)
        
        for r_k in range(start_row, start_row + 3):
            for c_k in range(col1, col2 + 1):
                ws.cell(row=r_k, column=c_k).border = b_style
                if r_k == start_row:
                    ws.cell(row=r_k, column=c_k).fill = fill_header
                elif r_k == start_row + 1 and c_k <= col_nom_fin:
                    ws.cell(row=r_k, column=c_k).fill = fill_red_card
                elif r_k == start_row + 2 and c_k <= col_nom_fin:
                    ws.cell(row=r_k, column=c_k).fill = fill_blue_card
                    
        # Combattant Rouge
        nom1 = p1.get('Nom', '') if isinstance(p1, dict) else str(p1)
        club1 = p1.get('Club', '') if isinstance(p1, dict) else ''
        text_r = f"🔴 {nom1}" + (f" ({club1})" if club1 and club1 != '-' else "")
        c_r = ws.cell(row=start_row+1, column=col1)
        if isinstance(p1, dict) and 'formula' in p1:
            c_r.value = p1['formula']
        else:
            c_r.value = text_r
        c_r.font = font_p_bold
        c_r.fill = fill_red_card
        c_r.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        
        c_sc_r = ws.cell(row=start_row+1, column=col_sc)
        c_sc_r.font, c_sc_r.fill = font_pts, fill_gray_box
        c_sc_r.alignment = Alignment(horizontal="center", vertical="center")
        
        c_ty_r = ws.cell(row=start_row+1, column=col_ty)
        c_ty_r.font, c_ty_r.fill = font_pts, fill_gray_box
        c_ty_r.alignment = Alignment(horizontal="center", vertical="center")
        
        c_pt_r = ws.cell(row=start_row+1, column=col_pt)
        c_pt_r.font, c_pt_r.fill = font_pts, fill_gray_box
        c_pt_r.alignment = Alignment(horizontal="center", vertical="center")
        
        # Combattant Bleu
        nom2 = p2.get('Nom', '') if isinstance(p2, dict) else str(p2)
        club2 = p2.get('Club', '') if isinstance(p2, dict) else ''
        text_b = f"🔵 {nom2}" + (f" ({club2})" if club2 and club2 != '-' else "")
        c_b = ws.cell(row=start_row+2, column=col1)
        if isinstance(p2, dict) and 'formula' in p2:
            c_b.value = p2['formula']
        else:
            c_b.value = text_b
        c_b.font = font_p_bold
        c_b.fill = fill_blue_card
        c_b.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        
        c_sc_b = ws.cell(row=start_row+2, column=col_sc)
        c_sc_b.font, c_sc_b.fill = font_pts, fill_gray_box
        c_sc_b.alignment = Alignment(horizontal="center", vertical="center")
        
        c_ty_b = ws.cell(row=start_row+2, column=col_ty)
        c_ty_b.font, c_ty_b.fill = font_pts, fill_gray_box
        c_ty_b.alignment = Alignment(horizontal="center", vertical="center")
        
        c_pt_b = ws.cell(row=start_row+2, column=col_pt)
        c_pt_b.font, c_pt_b.fill = font_pts, fill_gray_box
        c_pt_b.alignment = Alignment(horizontal="center", vertical="center")

    elif end_col is not None and end_col > start_col + 1:
        col1 = start_col
        col2 = end_col
        # En-tête du match sur toute la largeur (col1..col2)
        ws.merge_cells(start_row=start_row, start_column=col1, end_row=start_row, end_column=col2)
        c_h = ws.cell(row=start_row, column=col1, value=title)
        c_h.font = font_match_h
        c_h.fill = fill_header
        c_h.alignment = Alignment(horizontal="center", vertical="center")
        
        # Combattant Rouge (fusionné de col1 à col2-1)
        ws.merge_cells(start_row=start_row+1, start_column=col1, end_row=start_row+1, end_column=col2-1)
        # Combattant Bleu (fusionné de col1 à col2-1)
        ws.merge_cells(start_row=start_row+2, start_column=col1, end_row=start_row+2, end_column=col2-1)
        
        for r_k in range(start_row, start_row + 3):
            for c_k in range(col1, col2 + 1):
                ws.cell(row=r_k, column=c_k).border = b_style
                if r_k == start_row:
                    ws.cell(row=r_k, column=c_k).fill = fill_header
                elif r_k == start_row + 1 and c_k < col2:
                    ws.cell(row=r_k, column=c_k).fill = fill_red_card
                elif r_k == start_row + 2 and c_k < col2:
                    ws.cell(row=r_k, column=c_k).fill = fill_blue_card

        # Combattant Rouge
        nom1 = p1.get('Nom', '') if isinstance(p1, dict) else str(p1)
        club1 = p1.get('Club', '') if isinstance(p1, dict) else ''
        text_r = f"🔴 {nom1}" + (f" ({club1})" if club1 and club1 != '-' else "")
        c_r = ws.cell(row=start_row+1, column=col1)
        if isinstance(p1, dict) and 'formula' in p1:
            c_r.value = p1['formula']
        else:
            c_r.value = text_r
        c_r.font = font_p_bold
        c_r.fill = fill_red_card
        c_r.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c_r.border = b_style
        
        c_pt_r = ws.cell(row=start_row+1, column=col2)
        c_pt_r.font = font_pts
        c_pt_r.fill = fill_gray_box
        c_pt_r.alignment = Alignment(horizontal="center", vertical="center")
        c_pt_r.border = b_style
        
        # Combattant Bleu
        nom2 = p2.get('Nom', '') if isinstance(p2, dict) else str(p2)
        club2 = p2.get('Club', '') if isinstance(p2, dict) else ''
        text_b = f"🔵 {nom2}" + (f" ({club2})" if club2 and club2 != '-' else "")
        c_b = ws.cell(row=start_row+2, column=col1)
        if isinstance(p2, dict) and 'formula' in p2:
            c_b.value = p2['formula']
        else:
            c_b.value = text_b
        c_b.font = font_p_bold
        c_b.fill = fill_blue_card
        c_b.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c_b.border = b_style
        
        c_pt_b = ws.cell(row=start_row+2, column=col2)
        c_pt_b.font = font_pts
        c_pt_b.fill = fill_gray_box
        c_pt_b.alignment = Alignment(horizontal="center", vertical="center")
        c_pt_b.border = b_style
    else:
        col1 = start_col
        col2 = start_col + 1
        
        # En-tête du match
        ws.merge_cells(start_row=start_row, start_column=col1, end_row=start_row, end_column=col2)
        c_h = ws.cell(row=start_row, column=col1, value=title)
        c_h.font = font_match_h
        c_h.fill = fill_header
        c_h.alignment = Alignment(horizontal="center", vertical="center")
        c_h.border = b_style
        ws.cell(row=start_row, column=col2).border = b_style

        # Combattant Rouge
        nom1 = p1.get('Nom', '') if isinstance(p1, dict) else str(p1)
        club1 = p1.get('Club', '') if isinstance(p1, dict) else ''
        text_r = f"🔴 {nom1}" + (f" ({club1})" if club1 and club1 != '-' else "")
        c_r = ws.cell(row=start_row+1, column=col1)
        if isinstance(p1, dict) and 'formula' in p1:
            c_r.value = p1['formula']
        else:
            c_r.value = text_r
        c_r.font = font_p_bold
        c_r.fill = fill_red_card
        c_r.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c_r.border = b_style
        
        c_pt_r = ws.cell(row=start_row+1, column=col2)
        c_pt_r.font = font_pts
        c_pt_r.fill = fill_gray_box
        c_pt_r.alignment = Alignment(horizontal="center", vertical="center")
        c_pt_r.border = b_style
        
        # Combattant Bleu
        nom2 = p2.get('Nom', '') if isinstance(p2, dict) else str(p2)
        club2 = p2.get('Club', '') if isinstance(p2, dict) else ''
        text_b = f"🔵 {nom2}" + (f" ({club2})" if club2 and club2 != '-' else "")
        c_b = ws.cell(row=start_row+2, column=col1)
        if isinstance(p2, dict) and 'formula' in p2:
            c_b.value = p2['formula']
        else:
            c_b.value = text_b
        c_b.font = font_p_bold
        c_b.fill = fill_blue_card
        c_b.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c_b.border = b_style
        
        c_pt_b = ws.cell(row=start_row+2, column=col2)
        c_pt_b.font = font_pts
        c_pt_b.fill = fill_gray_box
        c_pt_b.alignment = Alignment(horizontal="center", vertical="center")
        c_pt_b.border = b_style

    if coords_map:
        m_info = coords_map.get((cat_poule, nom1, nom2))
        if not m_info:
            m_info = coords_map.get((cat_poule, nom2, nom1))
        if not m_info:
            b1 = re.sub(r'\s*\[.*?\]', '', nom1).strip()
            b2 = re.sub(r'\s*\[.*?\]', '', nom2).strip()
            m_info = coords_map.get((cat_poule, b1, b2))
            if not m_info:
                m_info = coords_map.get((cat_poule, b2, b1))
        if not m_info:
            b1 = re.sub(r'\s*\[.*?\]', '', nom1).strip()
            b2 = re.sub(r'\s*\[.*?\]', '', nom2).strip()
            for (c, k1, k2), inf in coords_map.items():
                if cat_poule == c or cat_poule in c or c in cat_poule:
                    kb1 = re.sub(r'\s*\[.*?\]', '', str(k1)).strip()
                    kb2 = re.sub(r'\s*\[.*?\]', '', str(k2)).strip()
                    if (nom1 == k1 and nom2 == k2) or (b1 == kb1 and b2 == kb2) or (nom1 == k2 and nom2 == k1) or (b1 == kb2 and b2 == kb1):
                        m_info = inf
                        break
        if m_info and isinstance(m_info, dict):
            s_name = m_info.get('sheet') or m_info.get('sheet_name', '')
            b1 = re.sub(r'\s*\[.*?\]', '', nom1).strip()
            bp1 = re.sub(r'\s*\[.*?\]', '', str(m_info.get('p1', ''))).strip()
            ptr_c = m_info.get('ptr_cell')
            ptb_c = m_info.get('ptb_cell')
            typer_c = m_info.get('typer_cell')
            typeb_c = m_info.get('typeb_cell')
            totr_c = m_info.get('tot_r_cell')
            totb_c = m_info.get('tot_b_cell')
            if not ptr_c and 'ptr' in m_info and hasattr(m_info['ptr'], 'coordinate'):
                ptr_c = m_info['ptr'].coordinate
            if not ptb_c and 'ptb' in m_info and hasattr(m_info['ptb'], 'coordinate'):
                ptb_c = m_info['ptb'].coordinate
            
            is_p1_red = (nom1 == m_info.get('p1') or (b1 and b1 == bp1) or (b1 and b1 in bp1) or (bp1 and bp1 in b1))
            
            if s_name and ptr_c and ptb_c:
                if is_p1_red:
                    c_pt_r.value = f"=IF('{s_name}'!{ptr_c}<>\"\", '{s_name}'!{ptr_c}, \"\")"
                    c_pt_b.value = f"=IF('{s_name}'!{ptb_c}<>\"\", '{s_name}'!{ptb_c}, \"\")"
                else:
                    c_pt_r.value = f"=IF('{s_name}'!{ptb_c}<>\"\", '{s_name}'!{ptb_c}, \"\")"
                    c_pt_b.value = f"=IF('{s_name}'!{ptr_c}<>\"\", '{s_name}'!{ptr_c}, \"\")"
            
            if with_details and s_name:
                if typer_c and typeb_c and c_ty_r and c_ty_b:
                    deduce_r_from_b = f'IF(\'{s_name}\'!{typeb_c}="VT","DT",IF(\'{s_name}\'!{typeb_c}="VST","DST",IF(\'{s_name}\'!{typeb_c}="VP","DP",IF(\'{s_name}\'!{typeb_c}="DT","VT",IF(\'{s_name}\'!{typeb_c}="DST","VST",IF(\'{s_name}\'!{typeb_c}="DP","VP",\"\"))))))'
                    deduce_b_from_r = f'IF(\'{s_name}\'!{typer_c}="VT","DT",IF(\'{s_name}\'!{typer_c}="VST","DST",IF(\'{s_name}\'!{typer_c}="VP","DP",IF(\'{s_name}\'!{typer_c}="DT","VT",IF(\'{s_name}\'!{typer_c}="DST","VST",IF(\'{s_name}\'!{typer_c}="DP","VP",\"\"))))))'
                    if is_p1_red:
                        c_ty_r.value = f"=IF('{s_name}'!{typer_c}<>\"\", '{s_name}'!{typer_c}, {deduce_r_from_b})"
                        c_ty_b.value = f"=IF('{s_name}'!{typeb_c}<>\"\", '{s_name}'!{typeb_c}, {deduce_b_from_r})"
                    else:
                        c_ty_r.value = f"=IF('{s_name}'!{typeb_c}<>\"\", '{s_name}'!{typeb_c}, {deduce_b_from_r})"
                        c_ty_b.value = f"=IF('{s_name}'!{typer_c}<>\"\", '{s_name}'!{typer_c}, {deduce_r_from_b})"
                if totr_c and totb_c and c_sc_r and c_sc_b:
                    if is_p1_red:
                        c_sc_r.value = f"=IF('{s_name}'!{totr_c}<>\"\", '{s_name}'!{totr_c}, 0)"
                        c_sc_b.value = f"=IF('{s_name}'!{totb_c}<>\"\", '{s_name}'!{totb_c}, 0)"
                    else:
                        c_sc_r.value = f"=IF('{s_name}'!{totb_c}<>\"\", '{s_name}'!{totb_c}, 0)"
                        c_sc_b.value = f"=IF('{s_name}'!{totr_c}<>\"\", '{s_name}'!{totr_c}, 0)"

    if with_details:
        return c_pt_r, c_pt_b, c_r, c_b, c_ty_r, c_ty_b, c_sc_r, c_sc_b
    return c_pt_r, c_pt_b, c_r, c_b


def draw_excel_vertical_connector(ws, start_row, end_row, col):
    for r in range(start_row, end_row + 1):
        cell = ws.cell(row=r, column=col)
        cell.border = Border(right=Side(style='medium', color='94A3B8'))


def draw_excel_podium_card(ws, start_row, start_col, title, subtitle, fill_bg, font_color="000000", formula_val=None):
    b_thin = Side(style='thin', color='CBD5E1')
    b_style = Border(left=b_thin, right=b_thin, top=b_thin, bottom=b_thin)
    ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row+1, end_column=start_col+1)
    c = ws.cell(row=start_row, column=start_col)
    if formula_val:
        c.value = formula_val
    else:
        c.value = f"{title}\n{subtitle}"
    c.font = Font(name="Arial", size=10, bold=True, color=font_color)
    c.fill = fill_bg
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(start_row, start_row+2):
        for col in range(start_col, start_col+2):
            ws.cell(row=r, column=col).border = b_style


def link_tapis_slot(tapis_slots, cat, p_nom, ws_bracket, bracket_cell):
    if not tapis_slots or not p_nom or not bracket_cell:
        return
    slots_to_update = []
    
    def collect_slots(k):
        v = tapis_slots.get(k)
        if v:
            if isinstance(v, list):
                slots_to_update.extend(v)
            elif isinstance(v, tuple):
                slots_to_update.append(v)

    collect_slots((cat, p_nom))
    collect_slots(p_nom)
    
    base_nom = re.sub(r'\s*\[.*?\]', '', str(p_nom)).strip()
    if base_nom and base_nom != p_nom:
        collect_slots((cat, base_nom))
        collect_slots(base_nom)
        
    if not slots_to_update:
        for k, sl in tapis_slots.items():
            if isinstance(k, tuple):
                c, nom = k
                c_clean = str(c)
                nom_clean = str(nom)
                if (p_nom == nom_clean or p_nom in nom_clean or nom_clean in p_nom or (base_nom and base_nom in nom_clean)) and (cat in c_clean or c_clean in cat):
                    if isinstance(sl, list): slots_to_update.extend(sl)
                    elif isinstance(sl, tuple): slots_to_update.append(sl)
            elif isinstance(k, str):
                if p_nom == k or p_nom in k or k in p_nom or (base_nom and base_nom in k):
                    if isinstance(sl, list): slots_to_update.extend(sl)
                    elif isinstance(sl, tuple): slots_to_update.append(sl)
                    
    seen = set()
    formula_str = f'=SUBSTITUTE(SUBSTITUTE(\'{ws_bracket.title}\'!{bracket_cell.coordinate}, "🔴 ", ""), "🔵 ", "")'
    for ws_m, cell_coord in slots_to_update:
        if (ws_m.title, cell_coord) not in seen:
            seen.add((ws_m.title, cell_coord))
            ws_m[cell_coord].value = formula_str



def construire_feuille_tableau_excel(ws, p_obj, nom_poule, liste_p, coords_matchs_tapis, nom_competition, tapis_slots=None, tapis_num=None):
    font_title = Font(name="Arial", size=13, bold=True, color="0055A4")
    font_sub = Font(name="Arial", size=9, italic=True, color="64748B")
    font_hdr = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    
    fill_dark = PatternFill("solid", fgColor="0F172A")
    fill_blue = PatternFill("solid", fgColor="0055A4")
    fill_sky = PatternFill("solid", fgColor="0284C7")
    fill_amber = PatternFill("solid", fgColor="B45309")
    fill_gold = PatternFill("solid", fgColor="FEF3C7")
    fill_silver = PatternFill("solid", fgColor="F1F5F9")
    fill_bronze = PatternFill("solid", fgColor="FFEDD5")
    fill_zebra = PatternFill("solid", fgColor="F8FAFC")
    fill_gray_h = PatternFill("solid", fgColor="475569")
    
    b_thin = Side(style='thin', color='CBD5E1')
    b_style = Border(left=b_thin, right=b_thin, top=b_thin, bottom=b_thin)

    ws.views.sheetView[0].showGridLines = True
    
    ws.cell(row=1, column=1, value=f"COMPÉTITION : {nom_competition.upper()} — TABLEAU OFFICIEL U13 : {nom_poule}").font = font_title
    ws.cell(row=2, column=1, value="Formule officielle FFLDA : Élimination directe avec repêchage des 1/4 de finale (2 Médailles de Bronze) — Orientation : Gauche ➔ Droite").font = font_sub
    if tapis_num:
        c_ret = ws.cell(row=3, column=1, value=f'=HYPERLINK("#\'Grille Tapis {tapis_num}\'!A1", "⬅️ Revenir à la Grille Tapis {tapis_num}")')
        c_ret.font = Font(name="Arial", size=9, bold=True, color="0055A4", underline="single")

    # Table des participants inscrits (Cols A-D)
    ws.cell(row=4, column=1, value="LISTE DES PARTICIPANTS").font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    for c_i, h in enumerate(["N°", "NOM Prénom", "CLUB", "POIDS"], 1):
        c = ws.cell(row=4, column=c_i, value=h)
        c.fill, c.font, c.alignment, c.border = fill_blue, font_hdr, Alignment(horizontal="center", vertical="center"), b_style

    for idx, p in enumerate(liste_p, 1):
        r = 4 + idx
        c_num = ws.cell(row=r, column=1, value=idx)
        c_num.alignment, c_num.border = Alignment(horizontal="center", vertical="center"), b_style
        
        c_nom = ws.cell(row=r, column=2, value=p['Nom'])
        c_nom.border = b_style
        
        c_club = ws.cell(row=r, column=3, value=p.get('Club', ''))
        c_club.border = b_style
        
        c_pds = ws.cell(row=r, column=4, value=formater_poids(p.get('Poids', '')))
        c_pds.alignment, c_pds.border = Alignment(horizontal="center", vertical="center"), b_style
        
        if idx % 2 == 0:
            c_num.fill = fill_zebra
            c_nom.fill = fill_zebra
            c_club.fill = fill_zebra
            c_pds.fill = fill_zebra

    max_len_nom = max([len(str(p.get('Nom', ''))) for p in liste_p] + [12])
    max_len_club = max([len(str(p.get('Club', ''))) for p in liste_p if p.get('Club') and p.get('Club') != '-'] + [8])
    col_nom_w = max(max_len_nom + 4, 24)
    col_club_w = max(max_len_club + 4, 16)

    max_len_match = max([
        len(f"🔴 {p.get('Nom', '')}" + (f" ({p.get('Club', '')})" if p.get('Club') and p.get('Club') != '-' else ""))
        for p in liste_p
    ] + [22])
    col_match_w = max(max_len_match + 3, 24)
    col_pod_w = max(int(col_match_w * 0.55), 12)

    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = col_nom_w
    ws.column_dimensions['C'].width = col_club_w
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 3

    rondes = p_obj.get('rondes', [])
    n = len(liste_p)
    
    m_cp = re.search(r'(\+?\d+\s*kg)', nom_poule)
    cat_poids = m_cp.group(1) if m_cp else ""

    f_or = rondes[-1][0]
    f_b1 = rondes[-1][1] if len(rondes[-1]) > 1 else None
    f_b2 = rondes[-1][2] if len(rondes[-1]) > 2 else None

    sf1 = rondes[-2][0]
    sf2 = rondes[-2][1]
    rep1 = rondes[-2][2] if len(rondes[-2]) > 2 else None
    rep2 = rondes[-2][3] if len(rondes[-2]) > 3 else None

    if len(rondes) >= 3:
        # Configuration dynamique des colonnes selon le nombre de tours (1/32, 1/16, 1/8, 1/4, 1/2, Finales)
        col_cur = 6
        if len(rondes) >= 6:
            # 1/32 de finale + 1/16 + 1/8 + 1/4 + 1/2 + Finales (33 <= n <= 64)
            col_32 = col_cur
            col_c_pre = col_cur + 2
            col_cur += 3

            col_16 = col_cur
            col_c0 = col_cur + 2
            col_cur += 3

            col_18 = col_cur
            col_c1 = col_cur + 2
            col_cur += 3

            col_qf = col_cur
            col_c2 = col_cur + 2
            col_cur += 3

            col_sf = col_cur
            col_c3 = col_cur + 2
            col_cur += 3

            col_fn = col_cur
            col_pod = col_cur + 2

            stage_cols = [col_32, col_16, col_18, col_qf, col_sf]
            conn_col_qf_sf = col_c2
            conn_col_sf_fn = col_c3
        elif len(rondes) == 5:
            # 1/16 de finale + 1/8 de finale + 1/4 + 1/2 + Finales (17 <= n <= 32)
            col_16 = col_cur
            col_c0 = col_cur + 2
            col_cur += 3
            
            col_18 = col_cur
            col_c1 = col_cur + 2
            col_cur += 3
            
            col_qf = col_cur
            col_c2 = col_cur + 2
            col_cur += 3
            
            col_sf = col_cur
            col_c3 = col_cur + 2
            col_cur += 3
            
            col_fn = col_cur
            col_pod = col_cur + 2
            
            stage_cols = [col_16, col_18, col_qf, col_sf]
            conn_col_qf_sf = col_c2
            conn_col_sf_fn = col_c3
        elif len(rondes) == 4:
            # Tour préliminaire (1/8) + 1/4 + 1/2 + Finales (9 <= n <= 16)
            col_prelim = col_cur
            col_c0 = col_cur + 2
            col_cur += 3
            
            col_qf = col_cur
            col_c1 = col_cur + 2
            col_cur += 3
            
            col_sf = col_cur
            col_c2 = col_cur + 2
            col_cur += 3
            
            col_fn = col_cur
            col_pod = col_cur + 2
            
            stage_cols = [col_prelim, col_qf, col_sf]
            conn_col_qf_sf = col_c1
            conn_col_sf_fn = col_c2
        else:
            # 1/4 + 1/2 + Finales (n = 7 ou 8)
            col_qf = col_cur
            col_c1 = col_cur + 2
            col_cur += 3
            
            col_sf = col_cur
            col_c2 = col_cur + 2
            col_cur += 3
            
            col_fn = col_cur
            col_pod = col_cur + 2
            
            stage_cols = [col_qf, col_sf]
            conn_col_qf_sf = col_c1
            conn_col_sf_fn = col_c2

        for sc in stage_cols:
            ws.column_dimensions[get_column_letter(sc)].width = col_match_w
            ws.column_dimensions[get_column_letter(sc+1)].width = 6
            ws.column_dimensions[get_column_letter(sc+2)].width = 3
        ws.column_dimensions[get_column_letter(col_fn)].width = col_match_w
        ws.column_dimensions[get_column_letter(col_fn+1)].width = 6
        ws.column_dimensions[get_column_letter(col_pod)].width = col_pod_w
        ws.column_dimensions[get_column_letter(col_pod+1)].width = col_pod_w

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=col_pod+1)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=col_pod+1)

        # Bannière principale
        ws.merge_cells(start_row=4, start_column=6, end_row=4, end_column=col_pod+1)
        c_bann = ws.cell(row=4, column=6, value="🏆 TABLEAU PRINCIPAL D'ÉLIMINATION DIRECTE (OR / ARGENT)")
        c_bann.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        c_bann.fill = fill_blue
        c_bann.alignment = Alignment(horizontal="center", vertical="center")

        qualif_map = {}
        max_upper_row = 23

        # 1. Traitement des tours préliminaires (1/32, 1/16 et 1/8) si n > 16 ou 9 <= n <= 16
        if len(rondes) >= 6:
            # 1/32 de finale
            ws.merge_cells(start_row=5, start_column=col_32, end_row=5, end_column=col_32+1)
            c_h32 = ws.cell(row=5, column=col_32, value="1/32 DE FINALE")
            c_h32.font, c_h32.fill, c_h32.alignment = Font(name="Arial", size=9, bold=True, color="FFFFFF"), fill_dark, Alignment(horizontal="center", vertical="center")

            r32_matches = rondes[0]
            max_upper_row = max(max_upper_row, 6 + len(r32_matches) * 4)
            for i, m in enumerate(r32_matches):
                r_i = 6 + i * 4
                ptr, ptb, c_r, c_b = draw_excel_match_card(ws, r_i, col_32, f"1/32 DE FINALE {i+1}", m[0], m[1], nom_poule, coords_matchs_tapis, bg_header=fill_gray_h)
                k1 = f"Vainqueur 1/32 ({i+1}) [{cat_poids}]"
                k2 = f"Vainqueur 1/32 ({i+1})"
                w_info = {'c_r': c_r, 'c_b': c_b, 'ptr': ptr, 'ptb': ptb, 'nom': f"Vainqueur 1/32 ({i+1})"}
                qualif_map[k1] = w_info
                qualif_map[k2] = w_info

            # 1/16 de finale
            ws.merge_cells(start_row=5, start_column=col_16, end_row=5, end_column=col_16+1)
            c_h16 = ws.cell(row=5, column=col_16, value="1/16 DE FINALE")
            c_h16.font, c_h16.fill, c_h16.alignment = Font(name="Arial", size=9, bold=True, color="FFFFFF"), fill_blue, Alignment(horizontal="center", vertical="center")

            r16_matches = rondes[1]
            max_upper_row = max(max_upper_row, 6 + len(r16_matches) * 4)
            for i, m in enumerate(r16_matches):
                r_i = 6 + i * 4
                p1_nom = m[0]['Nom'] if isinstance(m[0], dict) else str(m[0])
                p2_nom = m[1]['Nom'] if isinstance(m[1], dict) else str(m[1])
                
                if p1_nom in qualif_map:
                    qi = qualif_map[p1_nom]
                    p1_in = {'Nom': p1_nom, 'formula': make_winner_formula(qi['c_r'], qi['c_b'], qi['ptr'], qi['ptb'], qi['nom'], "🔴")}
                else:
                    p1_in = m[0]
                    
                if p2_nom in qualif_map:
                    qi = qualif_map[p2_nom]
                    p2_in = {'Nom': p2_nom, 'formula': make_winner_formula(qi['c_r'], qi['c_b'], qi['ptr'], qi['ptb'], qi['nom'], "🔵")}
                else:
                    p2_in = m[1]
                    
                ptr, ptb, c_r, c_b = draw_excel_match_card(ws, r_i, col_16, f"1/16 DE FINALE {i+1}", p1_in, p2_in, nom_poule, coords_matchs_tapis, bg_header=fill_gray_h)
                
                if p1_nom in qualif_map:
                    link_tapis_slot(tapis_slots, nom_poule, p1_nom, ws, c_r)
                if p2_nom in qualif_map:
                    link_tapis_slot(tapis_slots, nom_poule, p2_nom, ws, c_b)
                    
                k1 = f"Vainqueur 1/16 ({i+1}) [{cat_poids}]"
                k2 = f"Vainqueur 1/16 ({i+1})"
                w_info = {'c_r': c_r, 'c_b': c_b, 'ptr': ptr, 'ptb': ptb, 'nom': f"Vainqueur 1/16 ({i+1})"}
                qualif_map[k1] = w_info
                qualif_map[k2] = w_info

            # 1/8 de finale
            ws.merge_cells(start_row=5, start_column=col_18, end_row=5, end_column=col_18+1)
            c_h18 = ws.cell(row=5, column=col_18, value="1/8 DE FINALE")
            c_h18.font, c_h18.fill, c_h18.alignment = Font(name="Arial", size=9, bold=True, color="FFFFFF"), fill_blue, Alignment(horizontal="center", vertical="center")

            r18_matches = rondes[2]
            max_upper_row = max(max_upper_row, 6 + len(r18_matches) * 4)
            for i, m in enumerate(r18_matches):
                r_i = 6 + i * 4
                p1_nom = m[0]['Nom'] if isinstance(m[0], dict) else str(m[0])
                p2_nom = m[1]['Nom'] if isinstance(m[1], dict) else str(m[1])
                
                if p1_nom in qualif_map:
                    qi = qualif_map[p1_nom]
                    p1_in = {'Nom': p1_nom, 'formula': make_winner_formula(qi['c_r'], qi['c_b'], qi['ptr'], qi['ptb'], qi['nom'], "🔴")}
                else:
                    p1_in = m[0]
                    
                if p2_nom in qualif_map:
                    qi = qualif_map[p2_nom]
                    p2_in = {'Nom': p2_nom, 'formula': make_winner_formula(qi['c_r'], qi['c_b'], qi['ptr'], qi['ptb'], qi['nom'], "🔵")}
                else:
                    p2_in = m[1]
                    
                ptr, ptb, c_r, c_b = draw_excel_match_card(ws, r_i, col_18, f"1/8 DE FINALE {i+1}", p1_in, p2_in, nom_poule, coords_matchs_tapis, bg_header=fill_sky)
                
                if p1_nom in qualif_map:
                    link_tapis_slot(tapis_slots, nom_poule, p1_nom, ws, c_r)
                if p2_nom in qualif_map:
                    link_tapis_slot(tapis_slots, nom_poule, p2_nom, ws, c_b)
                    
                k1 = f"Vainqueur 1/8 ({i+1}) [{cat_poids}]"
                k2 = f"Vainqueur 1/8 ({i+1})"
                w_info = {'c_r': c_r, 'c_b': c_b, 'ptr': ptr, 'ptb': ptb, 'nom': f"Vainqueur 1/8 ({i+1})"}
                qualif_map[k1] = w_info
                qualif_map[k2] = w_info

        elif len(rondes) == 5:
            ws.merge_cells(start_row=5, start_column=col_16, end_row=5, end_column=col_16+1)
            c_h16 = ws.cell(row=5, column=col_16, value="1/16 DE FINALE")
            c_h16.font, c_h16.fill, c_h16.alignment = Font(name="Arial", size=9, bold=True, color="FFFFFF"), fill_dark, Alignment(horizontal="center", vertical="center")

            r16_matches = rondes[0]
            max_upper_row = max(max_upper_row, 6 + len(r16_matches) * 4)
            for i, m in enumerate(r16_matches):
                r_i = 6 + i * 4
                ptr, ptb, c_r, c_b = draw_excel_match_card(ws, r_i, col_16, f"1/16 DE FINALE {i+1}", m[0], m[1], nom_poule, coords_matchs_tapis, bg_header=fill_gray_h)
                k1 = f"Vainqueur 1/16 ({i+1}) [{cat_poids}]"
                k2 = f"Vainqueur 1/16 ({i+1})"
                w_info = {'c_r': c_r, 'c_b': c_b, 'ptr': ptr, 'ptb': ptb, 'nom': f"Vainqueur 1/16 ({i+1})"}
                qualif_map[k1] = w_info
                qualif_map[k2] = w_info

            # Traitement des 1/8 de finale
            ws.merge_cells(start_row=5, start_column=col_18, end_row=5, end_column=col_18+1)
            c_h18 = ws.cell(row=5, column=col_18, value="1/8 DE FINALE")
            c_h18.font, c_h18.fill, c_h18.alignment = Font(name="Arial", size=9, bold=True, color="FFFFFF"), fill_blue, Alignment(horizontal="center", vertical="center")

            r18_matches = rondes[1]
            max_upper_row = max(max_upper_row, 6 + len(r18_matches) * 4)
            for i, m in enumerate(r18_matches):
                r_i = 6 + i * 4
                p1_nom = m[0]['Nom'] if isinstance(m[0], dict) else str(m[0])
                p2_nom = m[1]['Nom'] if isinstance(m[1], dict) else str(m[1])
                
                if p1_nom in qualif_map:
                    qi = qualif_map[p1_nom]
                    p1_in = {'Nom': p1_nom, 'formula': make_winner_formula(qi['c_r'], qi['c_b'], qi['ptr'], qi['ptb'], qi['nom'], "🔴")}
                else:
                    p1_in = m[0]
                    
                if p2_nom in qualif_map:
                    qi = qualif_map[p2_nom]
                    p2_in = {'Nom': p2_nom, 'formula': make_winner_formula(qi['c_r'], qi['c_b'], qi['ptr'], qi['ptb'], qi['nom'], "🔵")}
                else:
                    p2_in = m[1]
                    
                ptr, ptb, c_r, c_b = draw_excel_match_card(ws, r_i, col_18, f"1/8 DE FINALE {i+1}", p1_in, p2_in, nom_poule, coords_matchs_tapis, bg_header=fill_sky)
                
                if p1_nom in qualif_map:
                    link_tapis_slot(tapis_slots, nom_poule, p1_nom, ws, c_r)
                if p2_nom in qualif_map:
                    link_tapis_slot(tapis_slots, nom_poule, p2_nom, ws, c_b)
                    
                k1 = f"Vainqueur 1/8 ({i+1}) [{cat_poids}]"
                k2 = f"Vainqueur 1/8 ({i+1})"
                w_info = {'c_r': c_r, 'c_b': c_b, 'ptr': ptr, 'ptb': ptb, 'nom': f"Vainqueur 1/8 ({i+1})"}
                qualif_map[k1] = w_info
                qualif_map[k2] = w_info

        elif len(rondes) == 4:
            # Traitement du Tour préliminaire (1/8)
            ws.merge_cells(start_row=5, start_column=col_prelim, end_row=5, end_column=col_prelim+1)
            c_hp = ws.cell(row=5, column=col_prelim, value="TOUR PRÉLIMINAIRE (1/8)")
            c_hp.font, c_hp.fill, c_hp.alignment = Font(name="Arial", size=9, bold=True, color="FFFFFF"), fill_dark, Alignment(horizontal="center", vertical="center")

            prelim_matches = rondes[0]
            max_upper_row = max(max_upper_row, 6 + len(prelim_matches) * 4)
            for i, m in enumerate(prelim_matches):
                r_i = 6 + i * 4
                ptr, ptb, c_r, c_b = draw_excel_match_card(ws, r_i, col_prelim, f"PRÉLIMINAIRE {i+1}", m[0], m[1], nom_poule, coords_matchs_tapis, bg_header=fill_gray_h)
                k1 = f"Vainqueur Prél. {i+1} [{cat_poids}]"
                k2 = f"Vainqueur Prél. {i+1}"
                w_info = {'c_r': c_r, 'c_b': c_b, 'ptr': ptr, 'ptb': ptb, 'nom': f"Vainqueur Prél. {i+1}"}
                qualif_map[k1] = w_info
                qualif_map[k2] = w_info

        # 2. Quarts de finale (Toujours rondes[-3])
        ws.merge_cells(start_row=5, start_column=col_qf, end_row=5, end_column=col_qf+1)
        c_hqf = ws.cell(row=5, column=col_qf, value="QUARTS DE FINALE")
        c_hqf.font, c_hqf.fill, c_hqf.alignment = Font(name="Arial", size=9, bold=True, color="FFFFFF"), fill_blue, Alignment(horizontal="center", vertical="center")

        q_matches = rondes[-3]
        q_list = [
            q_matches[0],
            q_matches[1],
            q_matches[2],
            q_matches[3] if len(q_matches) > 3 else (p_obj.get('p_exempt', liste_p[6]), {'Nom': 'EXEMPT (BYE)', 'Club': '-'})
        ]

        qf_cards = []
        qf_rows = [6, 11, 16, 21]
        for k in range(4):
            m = q_list[k]
            p1_obj = m[0]
            p2_obj = m[1]
            p1_nom = p1_obj.get('Nom', '') if isinstance(p1_obj, dict) else str(p1_obj)
            p2_nom = p2_obj.get('Nom', '') if isinstance(p2_obj, dict) else str(p2_obj)
            
            p1_in = p1_obj
            if p1_nom in qualif_map:
                qi = qualif_map[p1_nom]
                p1_in = {'Nom': p1_nom, 'formula': make_winner_formula(qi['c_r'], qi['c_b'], qi['ptr'], qi['ptb'], qi['nom'], "🔴")}
                
            p2_in = p2_obj
            if p2_nom in qualif_map:
                qi = qualif_map[p2_nom]
                p2_in = {'Nom': p2_nom, 'formula': make_winner_formula(qi['c_r'], qi['c_b'], qi['ptr'], qi['ptb'], qi['nom'], "🔵")}
                
            q_ptr, q_ptb, q_r, q_b = draw_excel_match_card(ws, qf_rows[k], col_qf, f"1/4 DE FINALE {k+1}", p1_in, p2_in, nom_poule, coords_matchs_tapis)
            qf_cards.append((q_ptr, q_ptb, q_r, q_b))
            
            if p1_nom in qualif_map:
                link_tapis_slot(tapis_slots, nom_poule, p1_nom, ws, q_r)
            if p2_nom in qualif_map:
                link_tapis_slot(tapis_slots, nom_poule, p2_nom, ws, q_b)

        qf1_ptr, qf1_ptb, qf1_r, qf1_b = qf_cards[0]
        qf2_ptr, qf2_ptb, qf2_r, qf2_b = qf_cards[1]
        qf3_ptr, qf3_ptb, qf3_r, qf3_b = qf_cards[2]
        qf4_ptr, qf4_ptb, qf4_r, qf4_b = qf_cards[3]

        draw_excel_vertical_connector(ws, 7, 12, conn_col_qf_sf)
        ws.cell(row=10, column=conn_col_qf_sf).border = Border(right=Side(style='medium', color='94A3B8'), bottom=Side(style='medium', color='94A3B8'))

        draw_excel_vertical_connector(ws, 17, 22, conn_col_qf_sf)
        ws.cell(row=20, column=conn_col_qf_sf).border = Border(right=Side(style='medium', color='94A3B8'), bottom=Side(style='medium', color='94A3B8'))

        # 3. Demi-finales dynamiques
        ws.merge_cells(start_row=5, start_column=col_sf, end_row=5, end_column=col_sf+1)
        c_hsf = ws.cell(row=5, column=col_sf, value="DEMI-FINALES")
        c_hsf.font, c_hsf.fill, c_hsf.alignment = Font(name="Arial", size=9, bold=True, color="FFFFFF"), fill_sky, Alignment(horizontal="center", vertical="center")

        sf1_p1 = {'Nom': sf1[0]['Nom'], 'formula': make_winner_formula(qf1_r, qf1_b, qf1_ptr, qf1_ptb, "Vainqueur 1/4 (1)", "🔴")}
        sf1_p2 = {'Nom': sf1[1]['Nom'], 'formula': make_winner_formula(qf2_r, qf2_b, qf2_ptr, qf2_ptb, "Vainqueur 1/4 (2)", "🔵")}
        sf1_ptr, sf1_ptb, sf1_r, sf1_b = draw_excel_match_card(ws, 8, col_sf, "DEMI-FINALE 1", sf1_p1, sf1_p2, nom_poule, coords_matchs_tapis, bg_header=fill_sky)

        sf2_p1 = {'Nom': sf2[0]['Nom'], 'formula': make_winner_formula(qf3_r, qf3_b, qf3_ptr, qf3_ptb, "Vainqueur 1/4 (3)", "🔴")}
        if n == 7:
            sf2_p2 = sf2[1]
        else:
            sf2_p2 = {'Nom': sf2[1]['Nom'], 'formula': make_winner_formula(qf4_r, qf4_b, qf4_ptr, qf4_ptb, "Vainqueur 1/4 (4)", "🔵")}
        sf2_ptr, sf2_ptb, sf2_r, sf2_b = draw_excel_match_card(ws, 18, col_sf, "DEMI-FINALE 2", sf2_p1, sf2_p2, nom_poule, coords_matchs_tapis, bg_header=fill_sky)

        draw_excel_vertical_connector(ws, 10, 19, conn_col_sf_fn)
        ws.cell(row=14, column=conn_col_sf_fn).border = Border(right=Side(style='medium', color='94A3B8'), bottom=Side(style='medium', color='94A3B8'))

        # 4. Grande Finale dynamique
        ws.merge_cells(start_row=5, start_column=col_fn, end_row=5, end_column=col_fn+1)
        c_hfn = ws.cell(row=5, column=col_fn, value="GRANDE FINALE")
        c_hfn.font, c_hfn.fill, c_hfn.alignment = Font(name="Arial", size=9, bold=True, color="FFFFFF"), fill_dark, Alignment(horizontal="center", vertical="center")

        fn_p1 = {'Nom': f_or[0]['Nom'], 'formula': make_winner_formula(sf1_r, sf1_b, sf1_ptr, sf1_ptb, "Vainqueur 1/2 (1)", "🔴")}
        fn_p2 = {'Nom': f_or[1]['Nom'], 'formula': make_winner_formula(sf2_r, sf2_b, sf2_ptr, sf2_ptb, "Vainqueur 1/2 (2)", "🔵")}
        fn_ptr, fn_ptb, fn_r, fn_b = draw_excel_match_card(ws, 13, col_fn, "GRANDE FINALE (OR)", fn_p1, fn_p2, nom_poule, coords_matchs_tapis, bg_header=fill_dark)

        # 5. Podiums Or & Argent
        r_fn = f"IF({fn_ptr.coordinate}=\"\",0,{fn_ptr.coordinate})"
        b_fn = f"IF({fn_ptb.coordinate}=\"\",0,{fn_ptb.coordinate})"
        form_gold = f'=IFERROR(IF({r_fn}+{b_fn}=0, "🥇 CHAMPION (OR)" & CHAR(10) & "En attente", "🥇 CHAMPION (OR)" & CHAR(10) & SUBSTITUTE(SUBSTITUTE(IF({r_fn}>{b_fn}, {fn_r.coordinate}, IF({b_fn}>{r_fn}, {fn_b.coordinate}, "En attente")), "🔴 ", ""), "🔵 ", "")), "🥇 CHAMPION (OR)" & CHAR(10) & "En attente")'
        form_silver = f'=IFERROR(IF({r_fn}+{b_fn}=0, "🥈 VICE-CHAMPION (ARGENT)" & CHAR(10) & "En attente", "🥈 VICE-CHAMPION (ARGENT)" & CHAR(10) & SUBSTITUTE(SUBSTITUTE(IF({r_fn}>{b_fn}, {fn_b.coordinate}, IF({b_fn}>{r_fn}, {fn_r.coordinate}, "En attente")), "🔴 ", ""), "🔵 ", "")), "🥈 VICE-CHAMPION (ARGENT)" & CHAR(10) & "En attente")'
        draw_excel_podium_card(ws, 12, col_pod, "🥇 CHAMPION (OR)", "Vainqueur Grande Finale", fill_gold, font_color="B45309", formula_val=form_gold)
        draw_excel_podium_card(ws, 15, col_pod, "🥈 VICE-CHAMPION (ARGENT)", "Perdant Grande Finale", fill_silver, font_color="475569", formula_val=form_silver)

        # 6. Repêchages & Bronze
        row_rep = max(26, max_upper_row + 3)
        ws.merge_cells(start_row=row_rep, start_column=col_qf, end_row=row_rep, end_column=col_pod+1)
        c_rep_h = ws.cell(row=row_rep, column=col_qf, value="🔄 TABLEAU DE REPÊCHAGE & MATCHS POUR LE BRONZE (2 Troisièmes Places)")
        c_rep_h.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        c_rep_h.fill = fill_amber
        c_rep_h.alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells(start_row=row_rep+1, start_column=col_qf, end_row=row_rep+1, end_column=col_pod+1)
        c_rep_sub = ws.cell(row=row_rep+1, column=col_qf, value="Perdants des 1/4 ➔ Repêchages ➔ Finales Bronze contre les perdants des Demi-Finales")
        c_rep_sub.font = font_sub
        c_rep_sub.alignment = Alignment(horizontal="left", vertical="center")

        # Repêchage 1 (Perdant QF 1 vs Perdant QF 2)
        rep1_p1 = {'Nom': rep1[0]['Nom'], 'formula': make_loser_formula(qf1_r, qf1_b, qf1_ptr, qf1_ptb, "Perdant 1/4 (1)", "🔴")}
        rep1_p2 = {'Nom': rep1[1]['Nom'], 'formula': make_loser_formula(qf2_r, qf2_b, qf2_ptr, qf2_ptb, "Perdant 1/4 (2)", "🔵")}
        rep1_ptr, rep1_ptb, rep1_r, rep1_b = draw_excel_match_card(ws, row_rep+3, col_qf, "REPÊCHAGE 1/4 (1)", rep1_p1, rep1_p2, nom_poule, coords_matchs_tapis, bg_header=fill_sky)

        if n >= 8:
            rep2_p1 = {'Nom': rep2[0]['Nom'], 'formula': make_loser_formula(qf3_r, qf3_b, qf3_ptr, qf3_ptb, "Perdant 1/4 (3)", "🔴")}
            rep2_p2 = {'Nom': rep2[1]['Nom'], 'formula': make_loser_formula(qf4_r, qf4_b, qf4_ptr, qf4_ptb, "Perdant 1/4 (4)", "🔵")}
            rep2_ptr, rep2_ptb, rep2_r, rep2_b = draw_excel_match_card(ws, row_rep+8, col_qf, "REPÊCHAGE 1/4 (2)", rep2_p1, rep2_p2, nom_poule, coords_matchs_tapis, bg_header=fill_sky)
        elif n == 7:
            ws.merge_cells(start_row=row_rep+8, start_column=col_qf, end_row=row_rep+10, end_column=col_qf+1)
            c_ex = ws.cell(row=row_rep+8, column=col_qf, value="Exempt de repêchage 1\n(Avance direct en Finale Bronze 2)")
            c_ex.font = Font(name="Arial", size=9, italic=True, color="64748B")
            c_ex.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for r_k in range(row_rep+8, row_rep+11):
                for c_k in range(col_qf, col_qf+2):
                    ws.cell(row=r_k, column=c_k).border = b_style

        ws.cell(row=row_rep+4, column=conn_col_qf_sf).border = Border(bottom=Side(style='medium', color='94A3B8'))
        ws.cell(row=row_rep+9, column=conn_col_qf_sf).border = Border(bottom=Side(style='medium', color='94A3B8'))

        # Finale Bronze 1 (Vainqueur Repêchage 1 vs Perdant Demi-Finale 1)
        fb1_p1 = {'Nom': f_b1[0]['Nom'], 'formula': make_winner_formula(rep1_r, rep1_b, rep1_ptr, rep1_ptb, "Vainqueur Repêchage 1", "🔴")}
        fb1_p2 = {'Nom': f_b1[1]['Nom'], 'formula': make_loser_formula(sf1_r, sf1_b, sf1_ptr, sf1_ptb, "Perdant 1/2 (1)", "🔵")}
        b1_ptr, b1_ptb, b1_r, b1_b = draw_excel_match_card(ws, row_rep+3, col_sf, "FINALE BRONZE 1", fb1_p1, fb1_p2, nom_poule, coords_matchs_tapis, bg_header=fill_amber)

        r_b1 = f"IF({b1_ptr.coordinate}=\"\",0,{b1_ptr.coordinate})"
        b_b1 = f"IF({b1_ptb.coordinate}=\"\",0,{b1_ptb.coordinate})"
        form_b1 = f'=IFERROR(IF({r_b1}+{b_b1}=0, "🥉 3ème PLACE (Bronze 1)" & CHAR(10) & "En attente", "🥉 3ème PLACE (Bronze 1)" & CHAR(10) & SUBSTITUTE(SUBSTITUTE(IF({r_b1}>{b_b1}, {b1_r.coordinate}, IF({b_b1}>{r_b1}, {b1_b.coordinate}, "En attente")), "🔴 ", ""), "🔵 ", "")), "🥉 3ème PLACE (Bronze 1)" & CHAR(10) & "En attente")'
        draw_excel_podium_card(ws, row_rep+3, col_pod, "🥉 3ème PLACE (Bronze 1)", "Vainqueur Finale Bronze 1", fill_bronze, font_color="9A3412", formula_val=form_b1)

        # Finale Bronze 2 (Vainqueur Repêchage 2 ou Perdant QF 3 vs Perdant Demi-Finale 2)
        if n == 7:
            fb2_p1 = {'Nom': f_b2[0]['Nom'], 'formula': make_loser_formula(qf3_r, qf3_b, qf3_ptr, qf3_ptb, "Perdant 1/4 (3)", "🔴")}
        else:
            fb2_p1 = {'Nom': f_b2[0]['Nom'], 'formula': make_winner_formula(rep2_r, rep2_b, rep2_ptr, rep2_ptb, "Vainqueur Repêchage 2", "🔴")}
        fb2_p2 = {'Nom': f_b2[1]['Nom'], 'formula': make_loser_formula(sf2_r, sf2_b, sf2_ptr, sf2_ptb, "Perdant 1/2 (2)", "🔵")}
        b2_ptr, b2_ptb, b2_r, b2_b = draw_excel_match_card(ws, row_rep+8, col_sf, "FINALE BRONZE 2", fb2_p1, fb2_p2, nom_poule, coords_matchs_tapis, bg_header=fill_amber)

        r_b2 = f"IF({b2_ptr.coordinate}=\"\",0,{b2_ptr.coordinate})"
        b_b2 = f"IF({b2_ptb.coordinate}=\"\",0,{b2_ptb.coordinate})"
        form_b2 = f'=IFERROR(IF({r_b2}+{b_b2}=0, "🥉 3ème PLACE (Bronze 2)" & CHAR(10) & "En attente", "🥉 3ème PLACE (Bronze 2)" & CHAR(10) & SUBSTITUTE(SUBSTITUTE(IF({r_b2}>{b_b2}, {b2_r.coordinate}, IF({b_b2}>{r_b2}, {b2_b.coordinate}, "En attente")), "🔴 ", ""), "🔵 ", "")), "🥉 3ème PLACE (Bronze 2)" & CHAR(10) & "En attente")'
        draw_excel_podium_card(ws, row_rep+8, col_pod, "🥉 3ème PLACE (Bronze 2)", "Vainqueur Finale Bronze 2", fill_bronze, font_color="9A3412", formula_val=form_b2)

        # Liaison dynamique vers les cartes de match sur les Grilles Tapis
        link_tapis_slot(tapis_slots, nom_poule, sf1[0]['Nom'], ws, sf1_r)
        link_tapis_slot(tapis_slots, nom_poule, sf1[1]['Nom'], ws, sf1_b)
        link_tapis_slot(tapis_slots, nom_poule, sf2[0]['Nom'], ws, sf2_r)
        if n >= 8 and isinstance(sf2[1], dict):
            link_tapis_slot(tapis_slots, nom_poule, sf2[1]['Nom'], ws, sf2_b)

        if rep1:
            link_tapis_slot(tapis_slots, nom_poule, rep1[0]['Nom'], ws, rep1_r)
            link_tapis_slot(tapis_slots, nom_poule, rep1[1]['Nom'], ws, rep1_b)
        if n >= 8 and rep2:
            link_tapis_slot(tapis_slots, nom_poule, rep2[0]['Nom'], ws, rep2_r)
            link_tapis_slot(tapis_slots, nom_poule, rep2[1]['Nom'], ws, rep2_b)

        link_tapis_slot(tapis_slots, nom_poule, f_or[0]['Nom'], ws, fn_r)
        link_tapis_slot(tapis_slots, nom_poule, f_or[1]['Nom'], ws, fn_b)

        if f_b1:
            link_tapis_slot(tapis_slots, nom_poule, f_b1[0]['Nom'], ws, b1_r)
            link_tapis_slot(tapis_slots, nom_poule, f_b1[1]['Nom'], ws, b1_b)
        if f_b2:
            link_tapis_slot(tapis_slots, nom_poule, f_b2[0]['Nom'], ws, b2_r)
            link_tapis_slot(tapis_slots, nom_poule, f_b2[1]['Nom'], ws, b2_b)

    else:
        col_cur = 6
        for r_idx, ronde in enumerate(rondes):
            col_match = col_cur
            col_conn = col_cur + 2
            
            ws.column_dimensions[get_column_letter(col_match)].width = col_match_w
            ws.column_dimensions[get_column_letter(col_match+1)].width = 6
            ws.column_dimensions[get_column_letter(col_conn)].width = 3

            r_title = f"RONDE {r_idx+1}"
            if r_idx == 0: r_title = "TOUR PRÉLIMINAIRE"
            elif r_idx == len(rondes)-3: r_title = "QUARTS DE FINALE"
            elif r_idx == len(rondes)-2: r_title = "DEMI-FINALES & REP."
            elif r_idx == len(rondes)-1: r_title = "FINALES"

            ws.merge_cells(start_row=4, start_column=col_match, end_row=4, end_column=col_match+1)
            c_rt = ws.cell(row=4, column=col_match, value=r_title)
            c_rt.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            c_rt.fill = fill_blue if r_idx < len(rondes)-1 else fill_dark
            c_rt.alignment = Alignment(horizontal="center", vertical="center")

            row_cur = 6
            for m_i, m in enumerate(ronde):
                m_label = f"Match {m_i+1}"
                draw_excel_match_card(ws, row_cur, col_match, m_label, m[0], m[1], nom_poule, coords_matchs_tapis)
                row_cur += 4
            
            col_cur += 3

        col_pod = col_cur
        ws.column_dimensions[get_column_letter(col_pod)].width = col_pod_w
        ws.column_dimensions[get_column_letter(col_pod+1)].width = col_pod_w
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=col_pod+1)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=col_pod+1)
        draw_excel_podium_card(ws, 6, col_pod, "🥇 CHAMPION (OR)", "Vainqueur Finale", fill_gold, font_color="B45309")
        draw_excel_podium_card(ws, 9, col_pod, "🥈 VICE-CHAMPION", "Finaliste", fill_silver, font_color="475569")
        draw_excel_podium_card(ws, 12, col_pod, "🥉 3ème PLACE (1)", "Bronze 1", fill_bronze, font_color="9A3412")
        draw_excel_podium_card(ws, 15, col_pod, "🥉 3ème PLACE (2)", "Bronze 2", fill_bronze, font_color="9A3412")

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def construire_feuille_poules_croisees_excel(ws, p_obj, nom_poule, liste_p, coords_matchs_tapis, nom_competition, tapis_slots=None, tapis_num=None):
    font_title = Font(name="Arial", size=13, bold=True, color="0055A4")
    font_sub = Font(name="Arial", size=9, italic=True, color="64748B")
    font_hdr = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_match_h = Font(name="Arial", size=9, bold=True, color="FFFFFF")
    font_pts = Font(name="Arial", size=10, bold=True)
    
    fill_dark = PatternFill("solid", fgColor="0F172A")
    fill_blue = PatternFill("solid", fgColor="0055A4")
    fill_sky = PatternFill("solid", fgColor="0284C7")
    fill_amber = PatternFill("solid", fgColor="B45309")
    fill_gray_h = PatternFill("solid", fgColor="475569")
    fill_gold = PatternFill("solid", fgColor="FEF3C7")
    fill_silver = PatternFill("solid", fgColor="F1F5F9")
    fill_bronze = PatternFill("solid", fgColor="FFEDD5")
    fill_active_match = PatternFill("solid", fgColor="E2E8F0")
    
    b_thin = Side(style='thin', color='CBD5E1')
    b_style = Border(left=b_thin, right=b_thin, top=b_thin, bottom=b_thin)

    ws.views.sheetView[0].showGridLines = True
    
    ws.cell(row=1, column=1, value=f"COMPÉTITION : {nom_competition.upper()} — POULES CROISÉES U13 (6 LUTTEURS) : {nom_poule}").font = font_title
    ws.cell(row=2, column=1, value="Formule officielle FFLDA : Phase 1 (2 Poules de 3 Nordiques) ➔ Phase 2 (Demi-Finales Croisées & Finales Or/Argent et Bronze unique) — Départage : Victoires ➔ Rencontre directe (si 2) ➔ Pt Clt ➔ VT ➔ VST ➔ Pts marqués ➔ Pts concédés ➔ N° Tirage").font = font_sub
    if tapis_num:
        c_ret = ws.cell(row=3, column=1, value=f'=HYPERLINK("#\'Grille Tapis {tapis_num}\'!A1", "⬅️ Revenir à la Grille Tapis {tapis_num}")')
        c_ret.font = Font(name="Arial", size=9, bold=True, color="0055A4", underline="single")

    poule_a = p_obj.get('poule_a', liste_p[:3])
    poule_b = p_obj.get('poule_b', liste_p[3:])
    rondes = p_obj.get('rondes', [])

    dv_type = DataValidation(type="list", formula1='"VT,VST,VP,DT,DST,DP"', allow_blank=True)
    dv_pts = DataValidation(type="list", formula1='"0,1,3,4,5"', allow_blank=True)
    has_dv_type = False
    has_dv_pts = False

    lignes_lutteurs = {}
    lutteur_cells = {p['Nom']: {'type_cells': [], 'marq_cells': [], 'conc_cells': []} for p in liste_p}
    match_col_map = {}

    for idx, p in enumerate(poule_a, 1):
        lignes_lutteurs[p['Nom']] = 5 + idx
    for idx, p in enumerate(poule_b, 1):
        lignes_lutteurs[p['Nom']] = 11 + idx

    col_fin_table = 13
    r_matches = 16
    ws.merge_cells(start_row=r_matches, start_column=1, end_row=r_matches, end_column=col_fin_table)
    c_m_hdr = ws.cell(row=r_matches, column=1, value="🤼 RENCONTRES DES POULES A & B (TOURS 1 À 3)")
    c_m_hdr.font, c_m_hdr.fill, c_m_hdr.alignment = font_hdr, fill_gray_h, Alignment(horizontal="center", vertical="center")
    
    r_matches += 1
    for tour_idx in range(3):
        if tour_idx < len(rondes):
            ronde = rondes[tour_idx]
            col_t_lettre = get_column_letter(5 + tour_idx)
            for m_idx, m in enumerate(ronde):
                grp_tag = "Poule A" if m_idx == 0 else "Poule B"
                title_m = f"T{tour_idx+1} ({grp_tag})"
                c_pt_r, c_pt_b, c_r, c_b, c_ty_r, c_ty_b, c_sc_r, c_sc_b = draw_excel_match_card(
                    ws, r_matches, 1, title_m, m[0], m[1], nom_poule, coords_matchs_tapis, 
                    bg_header=fill_gray_h, end_col=col_fin_table, with_details=True
                )
                
                if hasattr(c_pt_r, 'coordinate'):
                    dv_pts.add(c_pt_r.coordinate)
                    has_dv_pts = True
                if hasattr(c_pt_b, 'coordinate'):
                    dv_pts.add(c_pt_b.coordinate)
                    has_dv_pts = True
                if c_ty_r and hasattr(c_ty_r, 'coordinate'):
                    dv_type.add(c_ty_r.coordinate)
                    has_dv_type = True
                if c_ty_b and hasattr(c_ty_b, 'coordinate'):
                    dv_type.add(c_ty_b.coordinate)
                    has_dv_type = True

                p1_nom = m[0]['Nom']
                p2_nom = m[1]['Nom']
                match_col_map[(p1_nom, p2_nom)] = col_t_lettre
                match_col_map[(p2_nom, p1_nom)] = col_t_lettre

                if p1_nom in lignes_lutteurs:
                    c_cell = ws.cell(row=lignes_lutteurs[p1_nom], column=5+tour_idx, value=f"={c_pt_r.coordinate}")
                    c_cell.alignment, c_cell.border, c_cell.fill = Alignment(horizontal="center", vertical="center"), b_style, fill_active_match
                if p2_nom in lignes_lutteurs:
                    c_cell = ws.cell(row=lignes_lutteurs[p2_nom], column=5+tour_idx, value=f"={c_pt_b.coordinate}")
                    c_cell.alignment, c_cell.border, c_cell.fill = Alignment(horizontal="center", vertical="center"), b_style, fill_active_match
                
                if p1_nom in lutteur_cells:
                    if c_ty_r: lutteur_cells[p1_nom]['type_cells'].append(c_ty_r.coordinate)
                    if c_sc_r: lutteur_cells[p1_nom]['marq_cells'].append(c_sc_r.coordinate)
                    if c_sc_b: lutteur_cells[p1_nom]['conc_cells'].append(c_sc_b.coordinate)
                if p2_nom in lutteur_cells:
                    if c_ty_b: lutteur_cells[p2_nom]['type_cells'].append(c_ty_b.coordinate)
                    if c_sc_b: lutteur_cells[p2_nom]['marq_cells'].append(c_sc_b.coordinate)
                    if c_sc_r: lutteur_cells[p2_nom]['conc_cells'].append(c_sc_r.coordinate)

                r_matches += 4

    def render_sub_poule_table(ws, start_row, title, participants, bg_color):
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=col_fin_table)
        c_title = ws.cell(row=start_row, column=1, value=title)
        c_title.font, c_title.fill, c_title.alignment = font_hdr, bg_color, Alignment(horizontal="center", vertical="center")
        
        headers = ["CLT", "N°", "NOM Prénom", "CLUB", "Tour 1", "Tour 2", "Tour 3", "Total Vict", "Total Pts", "VT", "VST", "Pts Marq.", "Pts Conc."]
        for c_i, h in enumerate(headers, 1):
            c = ws.cell(row=start_row+1, column=c_i, value=h)
            c.font, c.fill, c.alignment, c.border = font_match_h, fill_gray_h, Alignment(horizontal="center", vertical="center"), b_style
            
        row_cur = start_row + 2
        plage_vict = f"H${start_row+2}:H${start_row+1+len(participants)}"
        plage_tot = f"I${start_row+2}:I${start_row+1+len(participants)}"

        for idx, p in enumerate(participants, 1):
            p_nom = p['Nom']
            l_info = lutteur_cells.get(p_nom, {'type_cells': [], 'marq_cells': [], 'conc_cells': []})
            
            c_num = ws.cell(row=row_cur, column=2, value=idx)
            c_num.alignment, c_num.border = Alignment(horizontal="center", vertical="center"), b_style
            
            c_nom = ws.cell(row=row_cur, column=3, value=p_nom)
            c_nom.border = b_style
            
            c_club = ws.cell(row=row_cur, column=4, value=p.get('Club', ''))
            c_club.border = b_style
            
            for t_i in range(3):
                c_t = ws.cell(row=row_cur, column=5+t_i)
                if not c_t.value:
                    c_t.value = "-"
                    c_t.alignment, c_t.border = Alignment(horizontal="center", vertical="center"), b_style
                else:
                    c_t.alignment, c_t.border, c_t.fill = Alignment(horizontal="center", vertical="center"), b_style, fill_active_match
            
            # Total Vict (basé sur le type : VT, VST, VP, avec repli sur les points >= 3 si le type n'est pas renseigné)
            c_vict = ws.cell(row=row_cur, column=8)
            if l_info['type_cells']:
                sum_vict = " + ".join([f'COUNTIF({coord}, "VT") + COUNTIF({coord}, "VST") + COUNTIF({coord}, "VP")' for coord in l_info['type_cells']])
                has_type = " + ".join([f'COUNTIF({coord}, "VT") + COUNTIF({coord}, "VST") + COUNTIF({coord}, "VP") + COUNTIF({coord}, "DT") + COUNTIF({coord}, "DST") + COUNTIF({coord}, "DP")' for coord in l_info['type_cells']])
                c_vict.value = f'=IF(({has_type})>0, {sum_vict}, COUNTIF(E{row_cur}:G{row_cur}, ">=3"))'
            else:
                c_vict.value = f'=COUNTIF(E{row_cur}:G{row_cur}, ">=3")'
            c_vict.font, c_vict.alignment, c_vict.border = font_pts, Alignment(horizontal="center", vertical="center"), b_style

            # Total Pts
            c_tot = ws.cell(row=row_cur, column=9, value=f"=SUM(E{row_cur}:G{row_cur})")
            c_tot.font, c_tot.alignment, c_tot.border = font_pts, Alignment(horizontal="center", vertical="center"), b_style

            # VT
            c_vt = ws.cell(row=row_cur, column=10)
            if l_info['type_cells']:
                c_vt.value = "=" + " + ".join([f'COUNTIF({coord}, "VT")' for coord in l_info['type_cells']])
            else:
                c_vt.value = 0
            c_vt.alignment, c_vt.border = Alignment(horizontal="center", vertical="center"), b_style

            # VST
            c_vst = ws.cell(row=row_cur, column=11)
            if l_info['type_cells']:
                c_vst.value = "=" + " + ".join([f'COUNTIF({coord}, "VST")' for coord in l_info['type_cells']])
            else:
                c_vst.value = 0
            c_vst.alignment, c_vst.border = Alignment(horizontal="center", vertical="center"), b_style

            # Pts Marq.
            c_marq = ws.cell(row=row_cur, column=12)
            if l_info['marq_cells']:
                c_marq.value = f"=SUM({', '.join(l_info['marq_cells'])})"
            else:
                c_marq.value = 0
            c_marq.alignment, c_marq.border = Alignment(horizontal="center", vertical="center"), b_style

            # Pts Conc.
            c_conc = ws.cell(row=row_cur, column=13)
            if l_info['conc_cells']:
                c_conc.value = f"=SUM({', '.join(l_info['conc_cells'])})"
            else:
                c_conc.value = 0
            c_conc.alignment, c_conc.border = Alignment(horizontal="center", vertical="center"), b_style

            # Formule CLT officielle FFLDA complète
            comparisons = []
            for idx_j, p_j in enumerate(participants, 1):
                if idx_j == idx:
                    continue
                r_j = start_row + 1 + idx_j
                col_t_lettre = match_col_map.get((p['Nom'], p_j['Nom']))
                if col_t_lettre:
                    comp = (
                        f"IF(H{r_j}>H{row_cur}, 1, "
                        f"IF(H{r_j}<H{row_cur}, 0, "
                        f"IF(AND(COUNTIF({plage_vict}, H{row_cur})=2, {col_t_lettre}{r_j}>{col_t_lettre}{row_cur}), 1, "
                        f"IF(AND(COUNTIF({plage_vict}, H{row_cur})=2, {col_t_lettre}{r_j}<{col_t_lettre}{row_cur}), 0, "
                        f"IF(I{r_j}>I{row_cur}, 1, "
                        f"IF(I{r_j}<I{row_cur}, 0, "
                        f"IF(J{r_j}>J{row_cur}, 1, "
                        f"IF(J{r_j}<J{row_cur}, 0, "
                        f"IF(K{r_j}>K{row_cur}, 1, "
                        f"IF(K{r_j}<K{row_cur}, 0, "
                        f"IF(L{r_j}>L{row_cur}, 1, "
                        f"IF(L{r_j}<L{row_cur}, 0, "
                        f"IF(M{r_j}<M{row_cur}, 1, "
                        f"IF(M{r_j}>M{row_cur}, 0, "
                        f"IF(B{r_j}<B{row_cur}, 1, 0)))))))))))))))"
                    )
                else:
                    comp = (
                        f"IF(H{r_j}>H{row_cur}, 1, "
                        f"IF(H{r_j}<H{row_cur}, 0, "
                        f"IF(I{r_j}>I{row_cur}, 1, "
                        f"IF(I{r_j}<I{row_cur}, 0, "
                        f"IF(J{r_j}>J{row_cur}, 1, "
                        f"IF(J{r_j}<J{row_cur}, 0, "
                        f"IF(K{r_j}>K{row_cur}, 1, "
                        f"IF(K{r_j}<K{row_cur}, 0, "
                        f"IF(L{r_j}>L{row_cur}, 1, "
                        f"IF(L{r_j}<L{row_cur}, 0, "
                        f"IF(M{r_j}<M{row_cur}, 1, "
                        f"IF(M{r_j}>M{row_cur}, 0, "
                        f"IF(B{r_j}<B{row_cur}, 1, 0)))))))))))))"
                    )
                comparisons.append(comp)

            somme_comp = " + ".join(comparisons) if comparisons else "0"
            c_clt = ws.cell(row=row_cur, column=1, value=f'=IF(SUM({plage_tot})=0, "", 1 + {somme_comp})')
            c_clt.alignment, c_clt.border = Alignment(horizontal="center", vertical="center"), b_style
            c_clt.font = Font(name="Arial", size=10, bold=True, color="0055A4")

            row_cur += 1

    render_sub_poule_table(ws, 4, "🥋 PHASE 1 : POULE A (3 Lutteurs)", poule_a, fill_blue)
    render_sub_poule_table(ws, 10, "🥋 PHASE 1 : POULE B (3 Lutteurs)", poule_b, fill_dark)

    max_len_nom = max([len(str(p.get('Nom', ''))) for p in liste_p] + [12])
    max_len_club = max([len(str(p.get('Club', ''))) for p in liste_p if p.get('Club') and p.get('Club') != '-'] + [8])
    col_nom_w = max(max_len_nom + 4, 22)
    col_club_w = max(max_len_club + 4, 16)

    max_len_match = max([
        len(f"🔴 {p.get('Nom', '')}" + (f" ({p.get('Club', '')})" if p.get('Club') and p.get('Club') != '-' else ""))
        for p in liste_p
    ] + [22])
    col_match_w = max(max_len_match + 3, 23)
    col_pod_w = max(int(col_match_w * 0.55), 12)

    col_sf = 15
    col_c = 17
    col_fn = 18
    col_pod = 20

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 5
    ws.column_dimensions['C'].width = col_nom_w
    ws.column_dimensions['D'].width = col_club_w
    ws.column_dimensions['E'].width = 8
    ws.column_dimensions['F'].width = 8
    ws.column_dimensions['G'].width = 8
    ws.column_dimensions['H'].width = 10
    ws.column_dimensions['I'].width = 10
    ws.column_dimensions['J'].width = 7
    ws.column_dimensions['K'].width = 7
    ws.column_dimensions['L'].width = 10
    ws.column_dimensions['M'].width = 10
    ws.column_dimensions['N'].width = 3

    ws.column_dimensions['O'].width = col_match_w
    ws.column_dimensions['P'].width = 6
    ws.column_dimensions['Q'].width = 3
    ws.column_dimensions['R'].width = col_match_w
    ws.column_dimensions['S'].width = 6
    ws.column_dimensions['T'].width = col_pod_w
    ws.column_dimensions['U'].width = col_pod_w

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=col_pod+1)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=col_pod+1)

    ws.merge_cells(start_row=4, start_column=col_sf, end_row=4, end_column=col_pod+1)
    c_fin_h = ws.cell(row=4, column=col_sf, value="🏆 PHASE 2 : PHASE FINALE CROISÉE (Gauche ➔ Droite)")
    c_fin_h.font, c_fin_h.fill, c_fin_h.alignment = font_hdr, fill_sky, Alignment(horizontal="center", vertical="center")

    sf1 = rondes[3][0]
    sf2 = rondes[3][1]
    f_or = rondes[4][0]
    f_b = rondes[4][1]

    form_1er_a = '=IF(OR(SUM(I$6:I$8)=0, ISNA(MATCH(1, A$6:A$8, 0))), "🔴 1er Poule A", "🔴 " & INDEX(C$6:C$8, MATCH(1, A$6:A$8, 0)) & IF(INDEX(D$6:D$8, MATCH(1, A$6:A$8, 0))<>"", " (" & INDEX(D$6:D$8, MATCH(1, A$6:A$8, 0)) & ")", ""))'
    form_2e_b  = '=IF(OR(SUM(I$12:I$14)=0, ISNA(MATCH(2, A$12:A$14, 0))), "🔵 2ème Poule B", "🔵 " & INDEX(C$12:C$14, MATCH(2, A$12:A$14, 0)) & IF(INDEX(D$12:D$14, MATCH(2, A$12:A$14, 0))<>"", " (" & INDEX(D$12:D$14, MATCH(2, A$12:A$14, 0)) & ")", ""))'
    form_1er_b = '=IF(OR(SUM(I$12:I$14)=0, ISNA(MATCH(1, A$12:A$14, 0))), "🔴 1er Poule B", "🔴 " & INDEX(C$12:C$14, MATCH(1, A$12:A$14, 0)) & IF(INDEX(D$12:D$14, MATCH(1, A$12:A$14, 0))<>"", " (" & INDEX(D$12:D$14, MATCH(1, A$12:A$14, 0)) & ")", ""))'
    form_2e_a  = '=IF(OR(SUM(I$6:I$8)=0, ISNA(MATCH(2, A$6:A$8, 0))), "🔵 2ème Poule A", "🔵 " & INDEX(C$6:C$8, MATCH(2, A$6:A$8, 0)) & IF(INDEX(D$6:D$8, MATCH(2, A$6:A$8, 0))<>"", " (" & INDEX(D$6:D$8, MATCH(2, A$6:A$8, 0)) & ")", ""))'

    p_sf1_1 = {'Nom': sf1[0]['Nom'], 'formula': form_1er_a}
    p_sf1_2 = {'Nom': sf1[1]['Nom'], 'formula': form_2e_b}
    sf1_ptr, sf1_ptb, sf1_r, sf1_b = draw_excel_match_card(ws, 6, col_sf, "DEMI-FINALE 1 (1er A vs 2ème B)", p_sf1_1, p_sf1_2, nom_poule, coords_matchs_tapis, bg_header=fill_blue)

    p_sf2_1 = {'Nom': sf2[0]['Nom'], 'formula': form_1er_b}
    p_sf2_2 = {'Nom': sf2[1]['Nom'], 'formula': form_2e_a}
    sf2_ptr, sf2_ptb, sf2_r, sf2_b = draw_excel_match_card(ws, 12, col_sf, "DEMI-FINALE 2 (1er B vs 2ème A)", p_sf2_1, p_sf2_2, nom_poule, coords_matchs_tapis, bg_header=fill_blue)

    draw_excel_vertical_connector(ws, 7, 13, col_c)
    ws.cell(row=8, column=col_c).border = Border(right=Side(style='medium', color='94A3B8'), bottom=Side(style='medium', color='94A3B8'))
    ws.cell(row=14, column=col_c).border = Border(right=Side(style='medium', color='94A3B8'), bottom=Side(style='medium', color='94A3B8'))

    # Finales dynamiques
    f_or_p1 = {'Nom': f_or[0]['Nom'], 'formula': make_winner_formula(sf1_r, sf1_b, sf1_ptr, sf1_ptb, "Vainqueur SF1", "🔴")}
    f_or_p2 = {'Nom': f_or[1]['Nom'], 'formula': make_winner_formula(sf2_r, sf2_b, sf2_ptr, sf2_ptb, "Vainqueur SF2", "🔵")}
    for_ptr, for_ptb, for_r, for_b = draw_excel_match_card(ws, 7, col_fn, "FINALE 1-2 (OR / ARGENT)", f_or_p1, f_or_p2, nom_poule, coords_matchs_tapis, bg_header=fill_dark)

    f_b_p1 = {'Nom': f_b[0]['Nom'], 'formula': make_loser_formula(sf1_r, sf1_b, sf1_ptr, sf1_ptb, "Perdant SF1", "🔴")}
    f_b_p2 = {'Nom': f_b[1]['Nom'], 'formula': make_loser_formula(sf2_r, sf2_b, sf2_ptr, sf2_ptb, "Perdant SF2", "🔵")}
    fb_ptr, fb_ptb, fb_r, fb_b = draw_excel_match_card(ws, 13, col_fn, "FINALE 3-4 (BRONZE UNIQUE)", f_b_p1, f_b_p2, nom_poule, coords_matchs_tapis, bg_header=fill_amber)

    # Podiums dynamiques
    r_for = f"IF({for_ptr.coordinate}=\"\",0,{for_ptr.coordinate})"
    b_for = f"IF({for_ptb.coordinate}=\"\",0,{for_ptb.coordinate})"
    r_fb = f"IF({fb_ptr.coordinate}=\"\",0,{fb_ptr.coordinate})"
    b_fb = f"IF({fb_ptb.coordinate}=\"\",0,{fb_ptb.coordinate})"
    form_or = f'=IF({r_for}+{b_for}=0, "🥇 OR" & CHAR(10) & "Vainqueur Finale 1-2", "🥇 OR" & CHAR(10) & SUBSTITUTE(SUBSTITUTE(IF({r_for}>{b_for}, {for_r.coordinate}, IF({b_for}>{r_for}, {for_b.coordinate}, "En attente")), "🔴 ", ""), "🔵 ", ""))'
    form_arg = f'=IF({r_for}+{b_for}=0, "🥈 ARGENT" & CHAR(10) & "Perdant Finale 1-2", "🥈 ARGENT" & CHAR(10) & SUBSTITUTE(SUBSTITUTE(IF({r_for}>{b_for}, {for_b.coordinate}, IF({b_for}>{r_for}, {for_r.coordinate}, "En attente")), "🔴 ", ""), "🔵 ", ""))'
    form_brz = f'=IF({r_fb}+{b_fb}=0, "🥉 BRONZE (Unique)" & CHAR(10) & "Vainqueur Finale 3-4", "🥉 BRONZE" & CHAR(10) & SUBSTITUTE(SUBSTITUTE(IF({r_fb}>{b_fb}, {fb_r.coordinate}, IF({b_fb}>{r_fb}, {fb_b.coordinate}, "En attente")), "🔴 ", ""), "🔵 ", ""))'

    draw_excel_podium_card(ws, 6, col_pod, "🥇 OR", "Vainqueur Finale 1-2", fill_gold, font_color="B45309", formula_val=form_or)
    draw_excel_podium_card(ws, 9, col_pod, "🥈 ARGENT", "Perdant Finale 1-2", fill_silver, font_color="475569", formula_val=form_arg)
    draw_excel_podium_card(ws, 13, col_pod, "🥉 BRONZE (Unique)", "Vainqueur Finale 3-4", fill_bronze, font_color="9A3412", formula_val=form_brz)

    # Liaison dynamique vers les cartes de match sur les Grilles Tapis
    link_tapis_slot(tapis_slots, nom_poule, sf1[0]['Nom'], ws, sf1_r)
    link_tapis_slot(tapis_slots, nom_poule, sf1[1]['Nom'], ws, sf1_b)
    link_tapis_slot(tapis_slots, nom_poule, sf2[0]['Nom'], ws, sf2_r)
    link_tapis_slot(tapis_slots, nom_poule, sf2[1]['Nom'], ws, sf2_b)

    link_tapis_slot(tapis_slots, nom_poule, f_or[0]['Nom'], ws, for_r)
    link_tapis_slot(tapis_slots, nom_poule, f_or[1]['Nom'], ws, for_b)

    link_tapis_slot(tapis_slots, nom_poule, f_b[0]['Nom'], ws, fb_r)
    link_tapis_slot(tapis_slots, nom_poule, f_b[1]['Nom'], ws, fb_b)

    if has_dv_pts:
        ws.add_data_validation(dv_pts)
    if has_dv_type:
        ws.add_data_validation(dv_type)

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def construire_feuille_plateau_u7_excel(ws, p_obj, nom_poule, liste_p, coords_matchs_tapis, nom_competition, tapis_num=None):
    font_title = Font(name="Arial", size=13, bold=True, color="B45309")
    font_sub = Font(name="Arial", size=9, italic=True, color="64748B")
    font_hdr = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_pts = Font(name="Arial", size=10, bold=True)
    
    fill_amber = PatternFill("solid", fgColor="D97706")  # Ambre U7
    fill_pastel = PatternFill("solid", fgColor="FEF3C7") # Pastel U7
    fill_zebra = PatternFill("solid", fgColor="FFFBEB")
    
    b_thin = Side(style='thin', color='CBD5E1')
    b_style = Border(left=b_thin, right=b_thin, top=b_thin, bottom=b_thin)
    
    ws.views.sheetView[0].showGridLines = True
    
    ws.cell(row=1, column=1, value=f"COMPÉTITION : {nom_competition.upper()} — ANIMATION PLATEAU U7 : {nom_poule}").font = font_title
    ws.cell(row=2, column=1, value="Formule officielle FFLDA : Découverte pédagogique sous forme de 3 plateaux d'activités avec rotation (Tous les enfants sont récompensés)").font = font_sub
    if tapis_num:
        c_ret = ws.cell(row=3, column=1, value=f'=HYPERLINK("#\'Grille Tapis {tapis_num}\'!A1", "⬅️ Revenir à la Grille Tapis {tapis_num}")')
        c_ret.font = Font(name="Arial", size=9, bold=True, color="B45309", underline="single")
    
    headers = [
        "N°", "NOM Prénom", "CLUB", "POIDS", 
        "Plateau 1 : Motricité & Agilité", "Plateau 2 : Ateliers techniques", "Plateau 3 : Oppositions", 
        "Validation / Récompense"
    ]
    
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    
    for c_i, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=c_i, value=h)
        c.fill, c.font, c.alignment, c.border = fill_amber, font_hdr, Alignment(horizontal="center", vertical="center", wrap_text=True), b_style
    ws.row_dimensions[4].height = 28
    
    for idx, p in enumerate(liste_p, 1):
        r = 4 + idx
        ws.row_dimensions[r].height = 22
        is_even = (idx % 2 == 0)
        row_fill = fill_zebra if is_even else PatternFill(fill_type=None)
        
        c_num = ws.cell(row=r, column=1, value=idx)
        c_num.alignment, c_num.border, c_num.fill = Alignment(horizontal="center", vertical="center"), b_style, row_fill
        
        c_nom = ws.cell(row=r, column=2, value=p.get('Nom', ''))
        c_nom.border, c_nom.fill = b_style, row_fill
        c_nom.alignment = Alignment(vertical="center", indent=1)
        
        c_club = ws.cell(row=r, column=3, value=p.get('Club', ''))
        c_club.border, c_club.fill = b_style, row_fill
        c_club.alignment = Alignment(vertical="center", indent=1)
        
        c_pds = ws.cell(row=r, column=4, value=formater_poids(p.get('Poids', '')))
        c_pds.alignment, c_pds.border, c_pds.fill = Alignment(horizontal="center", vertical="center"), b_style, row_fill
        
        for p_col in [5, 6, 7]:
            c_plat = ws.cell(row=r, column=p_col, value="[ ✓ ] Validé")
            c_plat.alignment, c_plat.border, c_plat.fill = Alignment(horizontal="center", vertical="center"), b_style, row_fill
            c_plat.font = Font(name="Arial", size=9, bold=True, color="166534")
            
        c_rec = ws.cell(row=r, column=8, value="🥇 Médaille d'Or / Diplôme")
        c_rec.alignment, c_rec.border, c_rec.fill = Alignment(horizontal="center", vertical="center"), b_style, fill_pastel
        c_rec.font = Font(name="Arial", size=9, bold=True, color="92400E")
        
    max_len_nom = max([len(str(p.get('Nom', ''))) for p in liste_p] + [12])
    max_len_club = max([len(str(p.get('Club', ''))) for p in liste_p if p.get('Club') and p.get('Club') != '-'] + [10])
    
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = max(max_len_nom + 4, 22)
    ws.column_dimensions['C'].width = max(max_len_club + 4, 16)
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 24
    ws.column_dimensions['F'].width = 24
    ws.column_dimensions['G'].width = 24
    ws.column_dimensions['H'].width = 24
    
    r_info = 4 + len(liste_p) + 2
    ws.merge_cells(start_row=r_info, start_column=1, end_row=r_info, end_column=8)
    c_info_h = ws.cell(row=r_info, column=1, value="ℹ️ DÉROULEMENT DES 3 PLATEAUX D'ACTIVITÉ U7")
    c_info_h.fill, c_info_h.font, c_info_h.alignment = fill_amber, font_hdr, Alignment(horizontal="center", vertical="center")
    
    explications = [
        ("🟡 Plateau 1 : Motricité & Agilité", "Parcours gymnique, franchissements, agilité, équilibre, réactivité motrice."),
        ("🟢 Plateau 2 : Ateliers techniques", "Ateliers d'apprentissage technique, habiletés motrices et gestes de lutte adaptés."),
        ("🔵 Plateau 3 : Oppositions", "Jeux de lutte et oppositions adaptées, combats éducatifs aménagés."),
        ("🏅 Esprit FFLDA", "Chaque enfant passe successivement sur les 3 plateaux. Tous reçoivent un diplôme et une médaille !")
    ]
    for idx_e, (tit, desc) in enumerate(explications, start=1):
        r_e = r_info + idx_e
        ws.cell(row=r_e, column=1, value=tit).font = Font(name="Arial", size=9, bold=True, color="92400E")
        ws.merge_cells(start_row=r_e, start_column=2, end_row=r_e, end_column=8)
        ws.cell(row=r_e, column=2, value=desc).font = Font(name="Arial", size=9, italic=True)
        
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def construire_feuille_poule_nordique_excel(ws, nom_poule, liste_p, rondes, coords_matchs_tapis, nom_competition, tapis_num=None):
    font_title = Font(name="Arial", size=13, bold=True, color="0055A4")
    font_sub = Font(name="Arial", size=9, italic=True, color="64748B")
    font_hdr = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_pts = Font(name="Arial", size=10, bold=True)
    
    fill_blue = PatternFill("solid", fgColor="0055A4")
    fill_gold = PatternFill("solid", fgColor="FEF3C7")
    fill_silver = PatternFill("solid", fgColor="F1F5F9")
    fill_bronze = PatternFill("solid", fgColor="FFEDD5")
    fill_zebra = PatternFill("solid", fgColor="F8FAFC")
    fill_gray_h = PatternFill("solid", fgColor="475569")
    
    b_thin = Side(style='thin', color='CBD5E1')
    b_style = Border(left=b_thin, right=b_thin, top=b_thin, bottom=b_thin)

    ws.views.sheetView[0].showGridLines = True
    
    # Titre officiel FFLDA & sous-titre
    ws.cell(row=1, column=1, value=f"COMPÉTITION : {nom_competition.upper()} — POULE OFFICIELLE : {nom_poule}").font = font_title
    ws.cell(row=2, column=1, value="Formule officielle FFLDA : Tournoi nordique (Tous contre tous) — Départage : Victoires ➔ Rencontre directe (si 2) ➔ Pt Clt ➔ VT ➔ VST ➔ Pts marqués ➔ Pts concédés ➔ N° Tirage").font = font_sub
    if tapis_num:
        c_ret = ws.cell(row=3, column=1, value=f'=HYPERLINK("#\'Grille Tapis {tapis_num}\'!A1", "⬅️ Revenir à la Grille Tapis {tapis_num}")')
        c_ret.font = Font(name="Arial", size=9, bold=True, color="0055A4", underline="single")

    is_u13_nordique = "u13" in nom_poule.lower()
    seuil_vict = 3 if is_u13_nordique else 2

    dv_nordic_type = DataValidation(type="list", formula1='"VT,VST,VP,DT,DST,DP"', allow_blank=True)
    dv_nordic_pts = DataValidation(type="list", formula1='"0,1,3,4,5"' if is_u13_nordique else '"0,1,2"', allow_blank=True)
    has_nordic_type = False
    has_nordic_pts = False

    nb_tours = len(rondes) if rondes else 0
    headers = ["CLT", "N°", "NOM Prénom", "CLUB", "COMITÉ"]
    for t in range(1, nb_tours + 1):
        headers.append(f"Tour {t}")
    headers.extend(["Total Vict", "Total Pts", "VT", "VST", "Pts Marq.", "Pts Conc.", "Poids"])
    
    col_vict_idx = 6 + nb_tours
    col_pts_idx = 7 + nb_tours
    col_vt_idx = 8 + nb_tours
    col_vst_idx = 9 + nb_tours
    col_marq_idx = 10 + nb_tours
    col_conc_idx = 11 + nb_tours
    col_poids_idx = 12 + nb_tours
    col_fin_table = col_poids_idx
    
    col_vict_lettre = get_column_letter(col_vict_idx)
    col_pts_lettre = get_column_letter(col_pts_idx)
    col_vt_lettre = get_column_letter(col_vt_idx)
    col_vst_lettre = get_column_letter(col_vst_idx)
    col_marq_lettre = get_column_letter(col_marq_idx)
    col_conc_lettre = get_column_letter(col_conc_idx)
    col_poids_lettre = get_column_letter(col_poids_idx)
    col_num_lettre = "B"
    col_debut_tours_lettre = get_column_letter(6)
    col_fin_tours_lettre = get_column_letter(5 + nb_tours) if nb_tours > 0 else col_pts_lettre

    row_cursor = 4
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=row_cursor, column=col_idx, value=h)
        c.font, c.alignment, c.border = font_hdr, Alignment(horizontal="center", vertical="center"), b_style
        c.fill = fill_blue
        
    lignes_lutteurs = {}
    ligne_debut_poule = row_cursor + 1
    ligne_fin_poule = ligne_debut_poule + len(liste_p) - 1
    
    plage_totaux = f"${col_pts_lettre}${ligne_debut_poule}:${col_pts_lettre}${ligne_fin_poule}"
    plage_vict = f"${col_vict_lettre}${ligne_debut_poule}:${col_vict_lettre}${ligne_fin_poule}"
    plage_clt = f"$A${ligne_debut_poule}:$A${ligne_fin_poule}"
    plage_nom = f"$C${ligne_debut_poule}:$C${ligne_fin_poule}"
    plage_club = f"$D${ligne_debut_poule}:$D${ligne_fin_poule}"

    # Correspondance des combats directs entre lutteurs pour le départage FFLDA
    match_col_map = {}
    lutteurs_par_tour = []
    for tour_idx, ronde in enumerate(rondes or [], 1):
        col_t_lettre = get_column_letter(5 + tour_idx)
        combatants = set()
        for match in ronde:
            p1_nom = match[0].get('Nom', '') if isinstance(match[0], dict) else str(match[0])
            p2_nom = match[1].get('Nom', '') if isinstance(match[1], dict) else str(match[1])
            match_col_map[(p1_nom, p2_nom)] = col_t_lettre
            match_col_map[(p2_nom, p1_nom)] = col_t_lettre
            combatants.add(p1_nom)
            combatants.add(p2_nom)
        lutteurs_par_tour.append(combatants)

    fill_active_match = PatternFill("solid", fgColor="E2E8F0")

    for idx, p in enumerate(liste_p, 1):
        r = ligne_debut_poule + idx - 1
        lignes_lutteurs[p['Nom']] = r
        
        c_num = ws.cell(row=r, column=2, value=idx)
        c_num.alignment, c_num.border = Alignment(horizontal="center", vertical="center"), b_style
        
        c_nom = ws.cell(row=r, column=3, value=p['Nom'])
        c_nom.border = b_style
        
        c_club = ws.cell(row=r, column=4, value=p.get('Club', ''))
        c_club.border = b_style
        
        c_comite = ws.cell(row=r, column=5, value=p.get('Comité', ''))
        c_comite.border = b_style
        
        # Total Vict
        c_tot_vict = ws.cell(row=r, column=col_vict_idx)
        if nb_tours > 0:
            c_tot_vict.value = f'=COUNTIF({col_debut_tours_lettre}{r}:{col_fin_tours_lettre}{r}, ">={seuil_vict}")'
        else:
            c_tot_vict.value = 0
        c_tot_vict.font, c_tot_vict.alignment, c_tot_vict.border = font_pts, Alignment(horizontal="center", vertical="center"), b_style
        
        # Total Pts
        c_tot_pts = ws.cell(row=r, column=col_pts_idx)
        if nb_tours > 0:
            c_tot_pts.value = f"=SUM({col_debut_tours_lettre}{r}:{col_fin_tours_lettre}{r})"
        else:
            c_tot_pts.value = 0
        c_tot_pts.font, c_tot_pts.alignment, c_tot_pts.border = font_pts, Alignment(horizontal="center", vertical="center"), b_style
        
        # Poids
        c_poids = ws.cell(row=r, column=col_poids_idx, value=formater_poids(p.get('Poids', '')))
        c_poids.alignment, c_poids.border = Alignment(horizontal="center", vertical="center"), b_style
        
        if idx % 2 == 0:
            for col_k in range(1, col_fin_table + 1):
                ws.cell(row=r, column=col_k).fill = fill_zebra

        for t in range(nb_tours):
            c_tour = ws.cell(row=r, column=6 + t)
            c_tour.alignment, c_tour.border = Alignment(horizontal="center", vertical="center"), b_style
            if t < len(lutteurs_par_tour) and p['Nom'] in lutteurs_par_tour[t]:
                c_tour.fill = fill_active_match

    # Fiches de combat modernes en bas de feuille
    lutteur_cells = {
        p['Nom']: {
            'type_cells': [],
            'marq_cells': [],
            'conc_cells': []
        }
        for p in liste_p
    }

    row_cursor = max(ligne_debut_poule + len(liste_p) + 2, 13)
    match_cpt = 1
    
    for tour_idx, ronde in enumerate(rondes, 1):
        ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=col_fin_table)
        c_tour = ws.cell(row=row_cursor, column=1, value=f"🎯 TOUR {tour_idx}")
        c_tour.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c_tour.fill = fill_blue
        c_tour.alignment = Alignment(horizontal="center", vertical="center")
        for c_k in range(1, col_fin_table + 1):
            ws.cell(row=row_cursor, column=c_k).border = b_style
        row_cursor += 2
        
        for match in ronde:
            p1, p2 = match[0], match[1]
            m_title = f"COMBAT N°{match_cpt} (TOUR {tour_idx})"
            
            c_pt_r, c_pt_b, c_r, c_b, c_ty_r, c_ty_b, c_sc_r, c_sc_b = draw_excel_match_card(
                ws=ws,
                start_row=row_cursor,
                start_col=1,
                title=m_title,
                p1=p1,
                p2=p2,
                cat_poule=nom_poule,
                coords_map=coords_matchs_tapis,
                bg_header=fill_gray_h,
                end_col=col_fin_table,
                with_details=True
            )
            
            if hasattr(c_pt_r, 'coordinate'):
                dv_nordic_pts.add(c_pt_r.coordinate)
                has_nordic_pts = True
            if hasattr(c_pt_b, 'coordinate'):
                dv_nordic_pts.add(c_pt_b.coordinate)
                has_nordic_pts = True
            if c_ty_r and hasattr(c_ty_r, 'coordinate'):
                dv_nordic_type.add(c_ty_r.coordinate)
                has_nordic_type = True
            if c_ty_b and hasattr(c_ty_b, 'coordinate'):
                dv_nordic_type.add(c_ty_b.coordinate)
                has_nordic_type = True
            
            p1_nom = p1.get('Nom', '') if isinstance(p1, dict) else str(p1)
            p2_nom = p2.get('Nom', '') if isinstance(p2, dict) else str(p2)
            
            if p1_nom in lignes_lutteurs:
                lig1 = lignes_lutteurs[p1_nom]
                ws.cell(row=lig1, column=6 + (tour_idx - 1), value=f"={c_pt_r.coordinate}").alignment = Alignment(horizontal="center", vertical="center")
            if p2_nom in lignes_lutteurs:
                lig2 = lignes_lutteurs[p2_nom]
                ws.cell(row=lig2, column=6 + (tour_idx - 1), value=f"={c_pt_b.coordinate}").alignment = Alignment(horizontal="center", vertical="center")
                
            if p1_nom in lutteur_cells:
                if c_ty_r: lutteur_cells[p1_nom]['type_cells'].append(c_ty_r.coordinate)
                if c_sc_r: lutteur_cells[p1_nom]['marq_cells'].append(c_sc_r.coordinate)
                if c_sc_b: lutteur_cells[p1_nom]['conc_cells'].append(c_sc_b.coordinate)
            if p2_nom in lutteur_cells:
                if c_ty_b: lutteur_cells[p2_nom]['type_cells'].append(c_ty_b.coordinate)
                if c_sc_b: lutteur_cells[p2_nom]['marq_cells'].append(c_sc_b.coordinate)
                if c_sc_r: lutteur_cells[p2_nom]['conc_cells'].append(c_sc_r.coordinate)
                
            match_cpt += 1
            row_cursor += 4
        row_cursor += 1

    # Remplissage des statistiques FFLDA (VT, VST, Pts Marq., Pts Conc.) et calcul du CLT
    for idx, p in enumerate(liste_p, 1):
        r = ligne_debut_poule + idx - 1
        p_nom = p['Nom']
        l_info = lutteur_cells.get(p_nom, {'type_cells': [], 'marq_cells': [], 'conc_cells': []})
        
        # Total Vict (basé sur le type : VT, VST, VP, avec repli sur les points >= seuil_vict si le type n'est pas renseigné)
        c_tot_vict = ws.cell(row=r, column=col_vict_idx)
        if l_info['type_cells']:
            sum_vict = " + ".join([f'COUNTIF({coord}, "VT") + COUNTIF({coord}, "VST") + COUNTIF({coord}, "VP")' for coord in l_info['type_cells']])
            has_type = " + ".join([f'COUNTIF({coord}, "VT") + COUNTIF({coord}, "VST") + COUNTIF({coord}, "VP") + COUNTIF({coord}, "DT") + COUNTIF({coord}, "DST") + COUNTIF({coord}, "DP")' for coord in l_info['type_cells']])
            if nb_tours > 0:
                c_tot_vict.value = f'=IF(({has_type})>0, {sum_vict}, COUNTIF({col_debut_tours_lettre}{r}:{col_fin_tours_lettre}{r}, ">={seuil_vict}"))'
            else:
                c_tot_vict.value = f'={sum_vict}'
        elif nb_tours > 0:
            c_tot_vict.value = f'=COUNTIF({col_debut_tours_lettre}{r}:{col_fin_tours_lettre}{r}, ">={seuil_vict}")'
        else:
            c_tot_vict.value = 0
        c_tot_vict.font, c_tot_vict.alignment, c_tot_vict.border = font_pts, Alignment(horizontal="center", vertical="center"), b_style

        # VT (Victoires par tombé)
        c_vt = ws.cell(row=r, column=col_vt_idx)
        if l_info['type_cells']:
            c_vt.value = "=" + " + ".join([f'COUNTIF({coord}, "VT")' for coord in l_info['type_cells']])
        else:
            c_vt.value = 0
        c_vt.alignment, c_vt.border = Alignment(horizontal="center", vertical="center"), b_style
        
        # VST (Victoires par supériorité technique)
        c_vst = ws.cell(row=r, column=col_vst_idx)
        if l_info['type_cells']:
            c_vst.value = "=" + " + ".join([f'COUNTIF({coord}, "VST")' for coord in l_info['type_cells']])
        else:
            c_vst.value = 0
        c_vst.alignment, c_vst.border = Alignment(horizontal="center", vertical="center"), b_style
        
        # Pts Marq. (Points techniques marqués - Total Score)
        c_marq = ws.cell(row=r, column=col_marq_idx)
        if l_info['marq_cells']:
            c_marq.value = f"=SUM({', '.join(l_info['marq_cells'])})"
        else:
            c_marq.value = 0
        c_marq.alignment, c_marq.border = Alignment(horizontal="center", vertical="center"), b_style
        
        # Pts Conc. (Points techniques concédés - Total Score adverse)
        c_conc = ws.cell(row=r, column=col_conc_idx)
        if l_info['conc_cells']:
            c_conc.value = f"=SUM({', '.join(l_info['conc_cells'])})"
        else:
            c_conc.value = 0
        c_conc.alignment, c_conc.border = Alignment(horizontal="center", vertical="center"), b_style
        
        # Formule CLT officielle FFLDA complète :
        # 1. Total Victoires
        # 2. Rencontre directe si 2 lutteurs à égalité de victoires
        # 3. Si 3+ lutteurs : Total Pts (Pt Clt)
        # 4. VT (Victoires par tombé)
        # 5. VST (Victoires par supériorité technique)
        # 6. Pts Marq. (Points techniques marqués - Total Score)
        # 7. Pts Conc. (Points techniques concédés - Le plus petit nombre)
        # 8. N° Tirage au sort le plus bas
        comparisons = []
        for idx_j, p_j in enumerate(liste_p, 1):
            if idx_j == idx:
                continue
            r_j = ligne_debut_poule + idx_j - 1
            col_t_lettre = match_col_map.get((p['Nom'], p_j['Nom']))
            
            if col_t_lettre:
                comp = (
                    f"IF({col_vict_lettre}{r_j}>{col_vict_lettre}{r}, 1, "
                    f"IF({col_vict_lettre}{r_j}<{col_vict_lettre}{r}, 0, "
                    f"IF(AND(COUNTIF({plage_vict}, {col_vict_lettre}{r})=2, {col_t_lettre}{r_j}>{col_t_lettre}{r}), 1, "
                    f"IF(AND(COUNTIF({plage_vict}, {col_vict_lettre}{r})=2, {col_t_lettre}{r_j}<{col_t_lettre}{r}), 0, "
                    f"IF({col_pts_lettre}{r_j}>{col_pts_lettre}{r}, 1, "
                    f"IF({col_pts_lettre}{r_j}<{col_pts_lettre}{r}, 0, "
                    f"IF({col_vt_lettre}{r_j}>{col_vt_lettre}{r}, 1, "
                    f"IF({col_vt_lettre}{r_j}<{col_vt_lettre}{r}, 0, "
                    f"IF({col_vst_lettre}{r_j}>{col_vst_lettre}{r}, 1, "
                    f"IF({col_vst_lettre}{r_j}<{col_vst_lettre}{r}, 0, "
                    f"IF({col_marq_lettre}{r_j}>{col_marq_lettre}{r}, 1, "
                    f"IF({col_marq_lettre}{r_j}<{col_marq_lettre}{r}, 0, "
                    f"IF({col_conc_lettre}{r_j}<{col_conc_lettre}{r}, 1, "
                    f"IF({col_conc_lettre}{r_j}>{col_conc_lettre}{r}, 0, "
                    f"IF({col_num_lettre}{r_j}<{col_num_lettre}{r}, 1, 0)))))))))))))))"
                )
            else:
                comp = (
                    f"IF({col_vict_lettre}{r_j}>{col_vict_lettre}{r}, 1, "
                    f"IF({col_vict_lettre}{r_j}<{col_vict_lettre}{r}, 0, "
                    f"IF({col_pts_lettre}{r_j}>{col_pts_lettre}{r}, 1, "
                    f"IF({col_pts_lettre}{r_j}<{col_pts_lettre}{r}, 0, "
                    f"IF({col_vt_lettre}{r_j}>{col_vt_lettre}{r}, 1, "
                    f"IF({col_vt_lettre}{r_j}<{col_vt_lettre}{r}, 0, "
                    f"IF({col_vst_lettre}{r_j}>{col_vst_lettre}{r}, 1, "
                    f"IF({col_vst_lettre}{r_j}<{col_vst_lettre}{r}, 0, "
                    f"IF({col_marq_lettre}{r_j}>{col_marq_lettre}{r}, 1, "
                    f"IF({col_marq_lettre}{r_j}<{col_marq_lettre}{r}, 0, "
                    f"IF({col_conc_lettre}{r_j}<{col_conc_lettre}{r}, 1, "
                    f"IF({col_conc_lettre}{r_j}>{col_conc_lettre}{r}, 0, "
                    f"IF({col_num_lettre}{r_j}<{col_num_lettre}{r}, 1, 0)))))))))))))"
                )
            comparisons.append(comp)

        if comparisons:
            somme_comp = " + ".join(comparisons)
            form_clt = f'=IF(SUM({plage_totaux})=0, "", 1 + {somme_comp})'
        else:
            form_clt = f'=IF(SUM({plage_totaux})=0, "", 1)'

        c_clt = ws.cell(row=r, column=1, value=form_clt)
        c_clt.alignment, c_clt.border = Alignment(horizontal="center", vertical="center"), b_style
        c_clt.font = Font(name="Arial", size=10, bold=True, color="0055A4")

    # Ajustement dynamique des largeurs de colonnes
    max_len_nom = max([len(str(p.get('Nom', ''))) for p in liste_p] + [12])
    max_len_club = max([len(str(p.get('Club', ''))) for p in liste_p if p.get('Club') and p.get('Club') != '-'] + [10])
    max_len_comite = max([len(str(p.get('Comité', ''))) for p in liste_p if p.get('Comité') and p.get('Comité') != '-'] + [10])

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 5
    ws.column_dimensions['C'].width = max(max_len_nom + 4, 24)
    ws.column_dimensions['D'].width = max(max_len_club + 4, 16)
    ws.column_dimensions['E'].width = max(max_len_comite + 4, 16)
    for t in range(nb_tours):
        col_t_lettre = get_column_letter(6 + t)
        ws.column_dimensions[col_t_lettre].width = 9
    ws.column_dimensions[col_vict_lettre].width = 11
    ws.column_dimensions[col_pts_lettre].width = 11
    ws.column_dimensions[col_vt_lettre].width = 7
    ws.column_dimensions[col_vst_lettre].width = 7
    ws.column_dimensions[col_marq_lettre].width = 11
    ws.column_dimensions[col_conc_lettre].width = 11
    ws.column_dimensions[col_poids_lettre].width = 10

    # Cartes Podium Dynamiques (reliées au classement)
    col_pod = col_poids_idx + 2
    col_pod_lettre1 = get_column_letter(col_pod)
    col_pod_lettre2 = get_column_letter(col_pod + 1)
    ws.column_dimensions[get_column_letter(col_pod - 1)].width = 3
    ws.column_dimensions[col_pod_lettre1].width = 16
    ws.column_dimensions[col_pod_lettre2].width = 16
    
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=col_pod+1)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=col_pod+1)
    
    form_gold = (
        f'=IFERROR(IF(SUM({plage_totaux})=0, "🥇 CHAMPION (OR)" & CHAR(10) & "En attente", '
        f'"🥇 CHAMPION (OR)" & CHAR(10) & INDEX({plage_nom}, MATCH(1, {plage_clt}, 0)) & '
        f'IF(INDEX({plage_club}, MATCH(1, {plage_clt}, 0))<>"", " (" & INDEX({plage_club}, MATCH(1, {plage_clt}, 0)) & ")", "")), "🥇 CHAMPION (OR)" & CHAR(10) & "En attente")'
    )
    form_silver = (
        f'=IFERROR(IF(SUM({plage_totaux})=0, "🥈 VICE-CHAMPION (ARGENT)" & CHAR(10) & "En attente", '
        f'"🥈 VICE-CHAMPION (ARGENT)" & CHAR(10) & INDEX({plage_nom}, MATCH(2, {plage_clt}, 0)) & '
        f'IF(INDEX({plage_club}, MATCH(2, {plage_clt}, 0))<>"", " (" & INDEX({plage_club}, MATCH(2, {plage_clt}, 0)) & ")", "")), "🥈 VICE-CHAMPION (ARGENT)" & CHAR(10) & "En attente")'
    )
    
    draw_excel_podium_card(ws, 4, col_pod, "🥇 CHAMPION (OR)", "En attente", fill_gold, font_color="B45309", formula_val=form_gold)
    draw_excel_podium_card(ws, 7, col_pod, "🥈 VICE-CHAMPION (ARGENT)", "En attente", fill_silver, font_color="475569", formula_val=form_silver)
    
    if len(liste_p) >= 3:
        form_bronze = (
            f'=IFERROR(IF(SUM({plage_totaux})=0, "🥉 3ème PLACE (BRONZE)" & CHAR(10) & "En attente", '
            f'"🥉 3ème PLACE (BRONZE)" & CHAR(10) & INDEX({plage_nom}, MATCH(3, {plage_clt}, 0)) & '
            f'IF(INDEX({plage_club}, MATCH(3, {plage_clt}, 0))<>"", " (" & INDEX({plage_club}, MATCH(3, {plage_clt}, 0)) & ")", "")), "🥉 3ème PLACE (BRONZE)" & CHAR(10) & "En attente")'
        )
        draw_excel_podium_card(ws, 10, col_pod, "🥉 3ème PLACE (BRONZE)", "En attente", fill_bronze, font_color="9A3412", formula_val=form_bronze)

    if has_nordic_pts:
        ws.add_data_validation(dv_nordic_pts)
    if has_nordic_type:
        ws.add_data_validation(dv_nordic_type)

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def fusionner_poules_isolees(poules, multiplicateur_poids, max_size):
    """
    Évite d'avoir un lutteur seul dans une poule de 1 tout en respectant 
    STRICTEMENT la tolérance d'écart de poids et la taille maximale de poule.
    """
    if len(poules) <= 1:
        return poules
    
    ameliore = True
    iterations = 0
    max_iterations = 20
    
    while ameliore and iterations < max_iterations:
        ameliore = False
        iterations += 1
        
        i = 0
        while i < len(poules):
            if len(poules[i]['participants']) == 1:
                p_iso = poules[i]['participants'][0]
                fusionne = False
                
                # Option 1: Essayer de fusionner directement avec la poule précédente (i-1)
                if i > 0:
                    cand = sorted(poules[i-1]['participants'] + [p_iso], key=lambda x: float(x.get('Poids_Num') or 0.0))
                    if len(cand) <= max_size and poule_poids_valide(cand, multiplicateur_poids):
                        poules[i-1]['participants'] = cand
                        poules[i-1]['rondes'] = generer_rondes_fflda(cand)
                        p_min = cand[0].get('Poids_Num') or 0.0
                        p_max = cand[-1].get('Poids_Num') or 0.0
                        prefix = poules[i-1]['nom'].split(' (')[0]
                        poules[i-1]['nom'] = f"{prefix} ({formater_poids_court(p_min)} - {formater_poids_court(p_max)})"
                        poules.pop(i)
                        fusionne = True
                        ameliore = True
                        break
                
                # Option 2: Essayer de fusionner directement avec la poule suivante (i+1)
                if not fusionne and i < len(poules) - 1:
                    cand = sorted([p_iso] + poules[i+1]['participants'], key=lambda x: float(x.get('Poids_Num') or 0.0))
                    if len(cand) <= max_size and poule_poids_valide(cand, multiplicateur_poids):
                        poules[i+1]['participants'] = cand
                        poules[i+1]['rondes'] = generer_rondes_fflda(cand)
                        p_min = cand[0].get('Poids_Num') or 0.0
                        p_max = cand[-1].get('Poids_Num') or 0.0
                        prefix = poules[i+1]['nom'].split(' (')[0]
                        poules[i+1]['nom'] = f"{prefix} ({formater_poids_court(p_min)} - {formater_poids_court(p_max)})"
                        poules.pop(i)
                        fusionne = True
                        ameliore = True
                        break
                
                # Option 3: Rééquilibrer avec la poule précédente (i-1)
                if not fusionne and i > 0 and len(poules[i-1]['participants']) >= 3:
                    prev_parts = poules[i-1]['participants']
                    for k in range(1, len(prev_parts) - 1):
                        new_prev = sorted(prev_parts[:-k], key=lambda x: float(x.get('Poids_Num') or 0.0))
                        new_iso = sorted(prev_parts[-k:] + [p_iso], key=lambda x: float(x.get('Poids_Num') or 0.0))
                        
                        if (len(new_prev) >= 2 and len(new_iso) >= 2 and len(new_iso) <= max_size and
                            poule_poids_valide(new_prev, multiplicateur_poids) and
                            poule_poids_valide(new_iso, multiplicateur_poids)):
                            
                            poules[i-1]['participants'] = new_prev
                            poules[i-1]['rondes'] = generer_rondes_fflda(new_prev)
                            p_min1 = new_prev[0].get('Poids_Num') or 0.0
                            p_max1 = new_prev[-1].get('Poids_Num') or 0.0
                            prefix1 = poules[i-1]['nom'].split(' (')[0]
                            poules[i-1]['nom'] = f"{prefix1} ({formater_poids_court(p_min1)} - {formater_poids_court(p_max1)})"
                            
                            poules[i]['participants'] = new_iso
                            poules[i]['rondes'] = generer_rondes_fflda(new_iso)
                            p_min2 = new_iso[0].get('Poids_Num') or 0.0
                            p_max2 = new_iso[-1].get('Poids_Num') or 0.0
                            prefix2 = poules[i]['nom'].split(' (')[0]
                            poules[i]['nom'] = f"{prefix2} ({formater_poids_court(p_min2)} - {formater_poids_court(p_max2)})"
                            
                            fusionne = True
                            ameliore = True
                            break
                    if fusionne:
                        break
                
                # Option 4: Rééquilibrer avec la poule suivante (i+1)
                if not fusionne and i < len(poules) - 1 and len(poules[i+1]['participants']) >= 3:
                    next_parts = poules[i+1]['participants']
                    for k in range(1, len(next_parts) - 1):
                        new_iso = sorted([p_iso] + next_parts[:k], key=lambda x: float(x.get('Poids_Num') or 0.0))
                        new_next = sorted(next_parts[k:], key=lambda x: float(x.get('Poids_Num') or 0.0))
                        
                        if (len(new_iso) >= 2 and len(new_next) >= 2 and len(new_iso) <= max_size and
                            poule_poids_valide(new_iso, multiplicateur_poids) and
                            poule_poids_valide(new_next, multiplicateur_poids)):
                            
                            poules[i]['participants'] = new_iso
                            poules[i]['rondes'] = generer_rondes_fflda(new_iso)
                            p_min1 = new_iso[0].get('Poids_Num') or 0.0
                            p_max1 = new_iso[-1].get('Poids_Num') or 0.0
                            prefix1 = poules[i]['nom'].split(' (')[0]
                            poules[i]['nom'] = f"{prefix1} ({formater_poids_court(p_min1)} - {formater_poids_court(p_max1)})"
                            
                            poules[i+1]['participants'] = new_next
                            poules[i+1]['rondes'] = generer_rondes_fflda(new_next)
                            p_min2 = new_next[0].get('Poids_Num') or 0.0
                            p_max2 = new_next[-1].get('Poids_Num') or 0.0
                            prefix2 = poules[i+1]['nom'].split(' (')[0]
                            poules[i+1]['nom'] = f"{prefix2} ({formater_poids_court(p_min2)} - {formater_poids_court(p_max2)})"
                            
                            fusionne = True
                            ameliore = True
                            break
                    if fusionne:
                        break
            i += 1
            
    return poules

def compter_collisions_club(participants_poule):
    clubs = [str(p.get('Club', '')).strip().lower() for p in participants_poule if str(p.get('Club', '')).strip() not in ['', '-', 'indépendant', 'independant', 'none', 'nan']]
    if not clubs:
        return 0
    from collections import Counter
    counts = Counter(clubs)
    return sum(c - 1 for c in counts.values() if c > 1)

def poule_poids_valide(participants_poule, multiplicateur_poids):
    if not participants_poule:
        return True
    poids_list = []
    for p in participants_poule:
        try:
            val = p.get('Poids_Num')
            if val is not None and str(val).strip() != '':
                f_val = float(val)
                if f_val > 0:
                    poids_list.append(f_val)
        except (ValueError, TypeError):
            continue
    if not poids_list:
        return True
    p_min = min(poids_list)
    p_max = max(poids_list)
    try:
        mult = float(multiplicateur_poids)
    except (ValueError, TypeError):
        mult = 1.15
    return p_max <= round(p_min * mult, 4)

def optimiser_poules_clubs(poules_groupe, multiplicateur_poids):
    """
    Permute les lutteurs entre poules d'un même groupe (âge/sexe/niveau) pour réduire 
    au maximum les affrontements entre lutteurs d'un même club, tout en respectant 
    strictement la tolérance d'écart de poids.
    """
    if len(poules_groupe) <= 1:
        return poules_groupe
    
    ameliore = True
    iterations = 0
    max_iterations = 50
    
    while ameliore and iterations < max_iterations:
        ameliore = False
        iterations += 1
        
        for i in range(len(poules_groupe)):
            for j in range(i + 1, len(poules_groupe)):
                p1 = poules_groupe[i]['participants']
                p2 = poules_groupe[j]['participants']
                
                cost_before = compter_collisions_club(p1) + compter_collisions_club(p2)
                if cost_before == 0:
                    continue
                
                best_swap = None
                best_cost = cost_before
                
                for idx1, w1 in enumerate(p1):
                    for idx2, w2 in enumerate(p2):
                        p1_test = p1[:idx1] + [w2] + p1[idx1+1:]
                        p2_test = p2[:idx2] + [w1] + p2[idx2+1:]
                        
                        if poule_poids_valide(p1_test, multiplicateur_poids) and poule_poids_valide(p2_test, multiplicateur_poids):
                            cost_after = compter_collisions_club(p1_test) + compter_collisions_club(p2_test)
                            if cost_after < best_cost:
                                best_cost = cost_after
                                best_swap = (idx1, idx2, p1_test, p2_test)
                
                if best_swap:
                    idx1, idx2, p1_test, p2_test = best_swap
                    poules_groupe[i]['participants'] = sorted(p1_test, key=lambda x: float(x.get('Poids_Num') or 0.0))
                    poules_groupe[j]['participants'] = sorted(p2_test, key=lambda x: float(x.get('Poids_Num') or 0.0))
                    
                    for idx_p in [i, j]:
                        parts = poules_groupe[idx_p]['participants']
                        prefix = poules_groupe[idx_p]['nom'].split(' (')[0]
                        p_min = parts[0].get('Poids_Num') or 0.0
                        p_max = parts[-1].get('Poids_Num') or 0.0
                        poules_groupe[idx_p]['nom'] = f"{prefix} ({formater_poids_court(p_min)} - {formater_poids_court(p_max)})"
                        poules_groupe[idx_p]['rondes'] = generer_rondes_fflda(parts)
                    
                    ameliore = True
                    break
            if ameliore:
                break
                
    return poules_groupe

# --- CONFIGURATION HARMONISÉE DES COULEURS PAR CATÉGORIE D'ÂGE (GRILLES DE PASSAGE & TAPIS) ---
COULEURS_AGE_GRILLE = {
    'U7': {
        'bg_excel_pastel': PatternFill("solid", fgColor="FEF3C7"),  # Ambre / Jaune doré doux
        'bg_excel_header': PatternFill("solid", fgColor="D97706"),  # Ambre soutenu
        'hex_pastel': '#FEF3C7',
        'hex_header': '#D97706',
        'hex_text': '#92400E',
        'badge': '🟡 U7',
        'nom': 'U7 (Jaune ambré)'
    },
    'U9': {
        'bg_excel_pastel': PatternFill("solid", fgColor="DCFCE7"),  # Vert menthe doux
        'bg_excel_header': PatternFill("solid", fgColor="059669"),  # Vert émeraude soutenu
        'hex_pastel': '#DCFCE7',
        'hex_header': '#059669',
        'hex_text': '#166534',
        'badge': '🟢 U9',
        'nom': 'U9 (Vert menthe)'
    },
    'U11': {
        'bg_excel_pastel': PatternFill("solid", fgColor="E0F2FE"),  # Bleu ciel doux
        'bg_excel_header': PatternFill("solid", fgColor="0284C7"),  # Bleu ciel soutenu
        'hex_pastel': '#E0F2FE',
        'hex_header': '#0284C7',
        'hex_text': '#075985',
        'badge': '🔵 U11',
        'nom': 'U11 (Bleu ciel)'
    },
    'U13': {
        'bg_excel_pastel': PatternFill("solid", fgColor="F3E8FF"),  # Violet / Lavande doux
        'bg_excel_header': PatternFill("solid", fgColor="7C3AED"),  # Violet / Indigo soutenu
        'hex_pastel': '#F3E8FF',
        'hex_header': '#7C3AED',
        'hex_text': '#6B21A8',
        'badge': '🟣 U13',
        'nom': 'U13 (Violet)'
    },
    'U15': {
        'bg_excel_pastel': PatternFill("solid", fgColor="FCE7F3"),  # Rose doux
        'bg_excel_header': PatternFill("solid", fgColor="DB2777"),  # Rose soutenu
        'hex_pastel': '#FCE7F3',
        'hex_header': '#DB2777',
        'hex_text': '#9D174D',
        'badge': '🌸 U15',
        'nom': 'U15 (Rose)'
    },
    'U17': {
        'bg_excel_pastel': PatternFill("solid", fgColor="FFEDD5"),  # Orange pêche doux
        'bg_excel_header': PatternFill("solid", fgColor="EA580C"),  # Orange soutenu
        'hex_pastel': '#FFEDD5',
        'hex_header': '#EA580C',
        'hex_text': '#9A3412',
        'badge': '🟠 U17',
        'nom': 'U17 (Orange)'
    },
    'U20': {
        'bg_excel_pastel': PatternFill("solid", fgColor="E2E8F0"),  # Ardoise argenté doux
        'bg_excel_header': PatternFill("solid", fgColor="475569"),  # Ardoise soutenu
        'hex_pastel': '#E2E8F0',
        'hex_header': '#475569',
        'hex_text': '#1E293B',
        'badge': '⚪ U20',
        'nom': 'U20 (Ardoise)'
    },
    'SENIOR': {
        'bg_excel_pastel': PatternFill("solid", fgColor="CCFBF1"),  # Turquoise doux
        'bg_excel_header': PatternFill("solid", fgColor="0D9488"),  # Sarcelle soutenu
        'hex_pastel': '#CCFBF1',
        'hex_header': '#0D9488',
        'hex_text': '#115E59',
        'badge': '🔘 Senior',
        'nom': 'Senior (Turquoise)'
    },
    'AUTRE': {
        'bg_excel_pastel': PatternFill("solid", fgColor="F1F5F9"),  # Gris clair neutre
        'bg_excel_header': PatternFill("solid", fgColor="0055A4"),  # Bleu officiel FFLDA
        'hex_pastel': '#F1F5F9',
        'hex_header': '#0055A4',
        'hex_text': '#1E293B',
        'badge': '🥋 Autre',
        'nom': 'Général (Bleu)'
    }
}

def extraire_age_de_texte(texte):
    """
    Extrait la catégorie d'âge (U7, U9, U11, U13, etc.) depuis le libellé d'un match ou d'une poule.
    """
    if not texte:
        return 'AUTRE'
    txt = str(texte).upper()
    for age in ['U7', 'U9', 'U11', 'U13', 'U15', 'U17', 'U20', 'SENIOR']:
        if re.search(rf'\[\s*{age}\b|\b{age}\b', txt):
            return age
    return 'AUTRE'


# --- GESTION DU CHRONOMÈTRE DE COMBAT (MINUTERIE & PÉRIODES) ---
_CHRONO_GLOBAL_SYNC = {}

def standardiser_id_match(m_id, tapis_num=None):
    """Garantit une clé d'identification unique et stable pour un combat ou un tapis."""
    if m_id is not None and str(m_id).strip() not in ("", "None", "0"):
        return str(m_id).strip()
    if tapis_num is not None and str(tapis_num).strip() not in ("", "None"):
        return f"tapis_{tapis_num}"
    return "default_match"

def formater_chrono_mm_ss(secondes):
    """Formate les secondes restantes en format MM:SS."""
    m = int(max(0, secondes)) // 60
    s = int(max(0, secondes)) % 60
    return f"{m:02d}:{s:02d}"

def publier_chrono_sync(m_id, tapis_num=None, sec=120, run=False, top=None, per=1, cat="", pause30=False, buzzer_event=None):
    """
    Publie l'état du chronomètre en temps réel sur disque et en mémoire partagée.
    Garantit une synchronisation parfaite à la seconde entre la Table de Marque et le Scoreboard TV,
    y compris à travers différents terminaux (smartphones, tablettes, TV connectées, QR codes).
    """
    try:
        import time as pytime, json, os, re
        
        m_id_str = standardiser_id_match(m_id, tapis_num)
        safe_m_id = re.sub(r'[^a-zA-Z0-9_-]', '_', m_id_str)
        t_key = int(tapis_num) if (tapis_num is not None and str(tapis_num).isdigit()) else None
        
        # Récupérer l'état existant pour préserver le timestamp buzzer_event si non fourni
        ancien_etat = _CHRONO_GLOBAL_SYNC.get(m_id_str, {})
        evt_b = float(buzzer_event) if buzzer_event is not None else float(ancien_etat.get("buzzer_event", 0.0))
        
        etat = {
            "m_id": m_id_str,
            "tapis": t_key,
            "sec": float(sec),
            "run": bool(run),
            "top": float(top) if top is not None else None,
            "per": int(per or 1),
            "cat": str(cat or ""),
            "pause30": bool(pause30),
            "buzzer_event": evt_b,
            "ts": pytime.time()
        }
        
        # 1. Enregistrement en mémoire partagée multi-sessions
        _CHRONO_GLOBAL_SYNC[m_id_str] = dict(etat)
        if m_id is not None:
            _CHRONO_GLOBAL_SYNC[str(m_id)] = dict(etat)
            _CHRONO_GLOBAL_SYNC[m_id] = dict(etat)
        if t_key is not None:
            _CHRONO_GLOBAL_SYNC[f"tapis_{t_key}"] = dict(etat)
            
        # 2. Persistance sur disque (IPC multi-processus / multi-écrans)
        os.makedirs(".cache_tournois", exist_ok=True)
        
        # Fichier lié au combat
        p_match = f".cache_tournois/chrono_match_{safe_m_id}.json"
        with open(p_match, "w", encoding="utf-8") as f:
            json.dump(etat, f)
            
        # Fichier lié au tapis (pour Scoreboard TV dédié à ce tapis)
        if t_key is not None:
            p_tapis = f".cache_tournois/chrono_tapis_{t_key}.json"
            with open(p_tapis, "w", encoding="utf-8") as f:
                json.dump(etat, f)

        # 3. Synchronisation immédiate de la session locale Streamlit
        try:
            import streamlit as st
            st.session_state[f"chrono_sec_{m_id_str}"] = int(sec)
            st.session_state[f"chrono_run_{m_id_str}"] = bool(run)
            st.session_state[f"chrono_top_{m_id_str}"] = top
            st.session_state[f"chrono_per_{m_id_str}"] = int(per or 1)
            st.session_state[f"chrono_pause30_{m_id_str}"] = bool(pause30)
            if buzzer_event is not None:
                st.session_state[f"buzzer_event_{m_id_str}"] = float(buzzer_event)
            if m_id is not None:
                st.session_state[f"chrono_sec_{m_id}"] = int(sec)
                st.session_state[f"chrono_run_{m_id}"] = bool(run)
                st.session_state[f"chrono_top_{m_id}"] = top
                st.session_state[f"chrono_per_{m_id}"] = int(per or 1)
                st.session_state[f"chrono_pause30_{m_id}"] = bool(pause30)
                if buzzer_event is not None:
                    st.session_state[f"buzzer_event_{m_id}"] = float(buzzer_event)
            if t_key is not None:
                st.session_state[f"chrono_per_tapis_{t_key}"] = int(per or 1)
                st.session_state[f"chrono_pause30_tapis_{t_key}"] = bool(pause30)
        except Exception:
            pass
    except Exception:
        pass

def recuperer_chrono_sync(m_id=None, tapis_num=None, categorie=""):
    """
    Récupère l'état le plus récent du chronomètre (disque ou mémoire partagée).
    """
    import time as pytime, json, os, re
    
    m_id_str = standardiser_id_match(m_id, tapis_num)
    safe_m_id = re.sub(r'[^a-zA-Z0-9_-]', '_', m_id_str)
    t_key = int(tapis_num) if (tapis_num is not None and str(tapis_num).isdigit()) else None
    
    etat = None
    
    # 1. Vérifier le fichier disque du tapis si t_key est fourni
    if t_key is not None:
        p_tapis = f".cache_tournois/chrono_tapis_{t_key}.json"
        if os.path.exists(p_tapis):
            try:
                with open(p_tapis, "r", encoding="utf-8") as f:
                    d = json.load(f)
                    if m_id is None or d.get("m_id") == m_id_str or str(d.get("m_id")) == str(m_id):
                        etat = d
                    elif not m_id_str or m_id_str == f"tapis_{t_key}":
                        etat = d
            except Exception:
                pass
                
    # 2. Vérifier le fichier disque du match si pas encore trouvé
    if etat is None and safe_m_id:
        p_match = f".cache_tournois/chrono_match_{safe_m_id}.json"
        if os.path.exists(p_match):
            try:
                with open(p_match, "r", encoding="utf-8") as f:
                    etat = json.load(f)
            except Exception:
                pass
                
    # 3. Vérifier en mémoire si pas trouvé sur disque ou si mémoire plus récente
    mem_etat = None
    if m_id_str in _CHRONO_GLOBAL_SYNC:
        mem_etat = _CHRONO_GLOBAL_SYNC[m_id_str]
    elif m_id in _CHRONO_GLOBAL_SYNC:
        mem_etat = _CHRONO_GLOBAL_SYNC[m_id]
    elif t_key is not None and f"tapis_{t_key}" in _CHRONO_GLOBAL_SYNC:
        mem_etat = _CHRONO_GLOBAL_SYNC[f"tapis_{t_key}"]
        
    if mem_etat:
        if etat is None or mem_etat.get("ts", 0) > etat.get("ts", 0):
            etat = mem_etat

    # Si aucun état n'existe encore, initialiser par défaut
    if etat is None:
        age = extraire_age_de_texte(categorie)
        duree = 180 if age in ["U13", "U15", "U17", "SENIOR"] else 120
        etat = {
            "m_id": m_id_str,
            "tapis": t_key,
            "sec": float(duree),
            "run": False,
            "top": None,
            "per": 1,
            "cat": str(categorie or ""),
            "pause30": False,
            "buzzer_event": 0.0,
            "ts": pytime.time()
        }
        publier_chrono_sync(m_id, tapis_num, duree, False, None, 1, categorie, pause30=False)
        
    return etat

def initialiser_chrono_match(m_id, categorie="", tapis_num=None):
    """Initialise l'état du chronomètre pour un combat donné."""
    return calculer_temps_restant_chrono(m_id, categorie, tapis_num)

def calculer_temps_restant_chrono(m_id, categorie="", tapis_num=None):
    """
    Calcule le temps restant réel en secondes avec synchronisation disk/IPC.
    Gère la transition automatique vers la pause de 30 secondes à la fin de la 1ère période,
    et la transition vers la 2ème période prête à la fin des 30 secondes.
    Met à jour session_state et retourne (sec_restante, est_en_marche).
    """
    import streamlit as st, time as pytime
    
    etat = recuperer_chrono_sync(m_id, tapis_num, categorie)
    
    sec_base = etat.get("sec", 120.0)
    en_marche = bool(etat.get("run", False))
    top = etat.get("top", None)
    per = int(etat.get("per", 1))
    pause30 = bool(etat.get("pause30", False))
    
    m_id_str = standardiser_id_match(m_id, tapis_num)
    t_key = int(tapis_num) if (tapis_num is not None and str(tapis_num).isdigit()) else None
    
    cle_sec = f"chrono_sec_{m_id_str}"
    cle_run = f"chrono_run_{m_id_str}"
    cle_top = f"chrono_top_{m_id_str}"
    cle_per = f"chrono_per_{m_id_str}"
    cle_p30 = f"chrono_pause30_{m_id_str}"
    
    if en_marche and top is not None:
        maintenant = pytime.time()
        ecoule = maintenant - top
        sec_restante = max(0.0, sec_base - ecoule)
        if sec_restante <= 0.0:
            if per == 1 and not pause30:
                # 🔔 Fin P1 -> Démarrage automatique de la pause réglementaire de 30 secondes
                pause30 = True
                sec_restante = 30.0
                en_marche = True
                top = maintenant
                publier_chrono_sync(m_id, tapis_num, 30.0, True, top, 1, categorie, pause30=True, buzzer_event=maintenant)
            elif pause30:
                # 🔔 Fin pause 30s -> Passage à la Période 2 prête (arrêtée, attend le coup de sifflet)
                age = extraire_age_de_texte(categorie)
                duree_p2 = 180.0 if age in ["U13", "U15", "U17", "SENIOR"] else 120.0
                pause30 = False
                per = 2
                sec_restante = duree_p2
                en_marche = False
                top = None
                publier_chrono_sync(m_id, tapis_num, duree_p2, False, None, 2, categorie, pause30=False, buzzer_event=maintenant)
            else:
                # 🔔 Fin de combat (Période 2 expirée)
                sec_restante = 0.0
                en_marche = False
                top = None
                publier_chrono_sync(m_id, tapis_num, 0, False, None, per, categorie, pause30=False, buzzer_event=maintenant)
    else:
        sec_restante = sec_base
        
    sec_int = int(sec_restante)
    
    # Mettre à jour la session locale Streamlit
    st.session_state[cle_sec] = sec_int
    st.session_state[cle_run] = en_marche
    st.session_state[cle_top] = top
    st.session_state[cle_per] = per
    st.session_state[cle_p30] = pause30
    if m_id is not None:
        st.session_state[f"chrono_sec_{m_id}"] = sec_int
        st.session_state[f"chrono_run_{m_id}"] = en_marche
        st.session_state[f"chrono_top_{m_id}"] = top
        st.session_state[f"chrono_per_{m_id}"] = per
        st.session_state[f"chrono_pause30_{m_id}"] = pause30
    if t_key is not None:
        st.session_state[f"chrono_per_tapis_{t_key}"] = per
        st.session_state[f"chrono_pause30_tapis_{t_key}"] = pause30
        
    return sec_int, en_marche

def toggle_chrono_match(m_id, categorie="", tapis_num=None):
    """Bascule entre Démarrer et Pause pour le chronomètre."""
    import time as pytime
    
    sec_actuelle, en_cours = calculer_temps_restant_chrono(m_id, categorie, tapis_num)
    etat = recuperer_chrono_sync(m_id, tapis_num, categorie)
    per = etat.get("per", 1)
    pause30 = bool(etat.get("pause30", False))
    
    if en_cours:
        # Pause : on fige le temps restant
        nouv_sec = sec_actuelle
        nouv_run = False
        nouv_top = None
    else:
        # Démarrage
        if sec_actuelle <= 0:
            if pause30:
                # Si la pause 30s était terminée, on passe en P2
                pause30 = False
                per = 2
            age = extraire_age_de_texte(categorie)
            sec_actuelle = 180 if age in ["U13", "U15", "U17", "SENIOR"] else 120
        nouv_sec = sec_actuelle
        nouv_run = True
        nouv_top = pytime.time()
        
    publier_chrono_sync(m_id, tapis_num, nouv_sec, nouv_run, nouv_top, per, categorie, pause30=pause30)

def reset_chrono_match(m_id, categorie="", tapis_num=None):
    """Remet à zéro le chronomètre selon la durée officielle de la catégorie (Période 1)."""
    age = extraire_age_de_texte(categorie)
    duree_defaut = 180 if age in ["U13", "U15", "U17", "SENIOR"] else 120
    publier_chrono_sync(m_id, tapis_num, duree_defaut, False, None, 1, categorie, pause30=False)

def ajuster_chrono_match(m_id, delta_sec, categorie="", tapis_num=None):
    """Ajuste finement le temps restant (+/- 10s)."""
    import time as pytime
    sec_actuelle, en_cours = calculer_temps_restant_chrono(m_id, categorie, tapis_num)
    nouvelle_val = max(0, sec_actuelle + delta_sec)
    etat = recuperer_chrono_sync(m_id, tapis_num, categorie)
    per = etat.get("per", 1)
    pause30 = bool(etat.get("pause30", False))
    nouv_top = pytime.time() if en_cours else None
    publier_chrono_sync(m_id, tapis_num, nouvelle_val, en_cours, nouv_top, per, categorie, pause30=pause30)

def changer_periode_chrono(m_id, tapis_num=None, categorie=""):
    """Change la période de combat (1 ou 2), ou quitte la pause 30s pour lancer P2 immédiatement."""
    import time as pytime
    sec_actuelle, en_cours = calculer_temps_restant_chrono(m_id, categorie, tapis_num)
    etat = recuperer_chrono_sync(m_id, tapis_num, categorie)
    p_cur = etat.get("per", 1)
    pause30 = bool(etat.get("pause30", False))
    
    age = extraire_age_de_texte(categorie)
    duree_regl = 180 if age in ["U13", "U15", "U17", "SENIOR"] else 120
    
    if pause30:
        # Lancer ou préparer directement la Période 2
        publier_chrono_sync(m_id, tapis_num, duree_regl, False, None, 2, categorie, pause30=False)
    else:
        nouv_per = 2 if p_cur == 1 else 1
        nouv_sec = duree_regl if sec_actuelle <= 0 else sec_actuelle
        nouv_top = pytime.time() if en_cours else None
        publier_chrono_sync(m_id, tapis_num, nouv_sec, en_cours, nouv_top, nouv_per, categorie, pause30=False)

def declencher_buzzer_manuel(m_id, categorie="", tapis_num=None):
    """Déclenche manuellement le signal sonore (buzzer / klaxon) sur la table et le scoreboard."""
    import time as pytime
    now = pytime.time()
    sec_actuelle, en_cours = calculer_temps_restant_chrono(m_id, categorie, tapis_num)
    etat = recuperer_chrono_sync(m_id, tapis_num, categorie)
    per = etat.get("per", 1)
    pause30 = bool(etat.get("pause30", False))
    top = pytime.time() if en_cours else None
    publier_chrono_sync(m_id, tapis_num, sec_actuelle, en_cours, top, per, categorie, pause30=pause30, buzzer_event=now)
    st.session_state[f"play_local_buzzer_{m_id}"] = now


def generer_html_grille_coloree(grille_lignes, nb_tapis):
    """
    Génère un tableau HTML stylé avec les couleurs de fond spécifiques à chaque catégorie d'âge.
    Parfait pour l'affichage interactif Streamlit et l'export HTML d'impression.
    """
    html = []
    html.append('''
    <div style="display:flex; flex-wrap:wrap; gap:10px; align-items:center; margin-bottom:14px; padding:10px 14px; background:#F8FAFC; border:1px solid #E2E8F0; border-radius:8px;">
        <span style="font-weight:bold; color:#1E293B; font-size:13px;">🎨 Légende des Catégories d'Âge :</span>
        <span style="background:#FEF3C7; color:#92400E; padding:3px 10px; border-radius:12px; font-weight:bold; font-size:12px; border:1px solid #FCD34D;">🟡 U7</span>
        <span style="background:#DCFCE7; color:#166534; padding:3px 10px; border-radius:12px; font-weight:bold; font-size:12px; border:1px solid #86EFAC;">🟢 U9</span>
        <span style="background:#E0F2FE; color:#075985; padding:3px 10px; border-radius:12px; font-weight:bold; font-size:12px; border:1px solid #7DD3FC;">🔵 U11</span>
        <span style="background:#F3E8FF; color:#6B21A8; padding:3px 10px; border-radius:12px; font-weight:bold; font-size:12px; border:1px solid #D8B4FE;">🟣 U13</span>
        <span style="background:#EF4135; color:#FFFFFF; padding:3px 10px; border-radius:12px; font-weight:bold; font-size:12px;">⏸️ Pause</span>
        <span style="background:#EFEFEF; color:#666666; padding:3px 10px; border-radius:12px; font-style:italic; font-size:12px;">⏳ Repos / Pesée</span>
    </div>
    ''')
    
    html.append('<div style="overflow-x:auto;"><table style="width:100%; border-collapse:collapse; font-family:Arial, sans-serif;">')
    html.append('<thead><tr>')
    for t in range(nb_tapis):
        html.append(f'<th style="background:#0055A4; color:#ffffff; padding:10px 8px; border:1px solid #CBD5E1; text-align:center; font-size:14px; font-weight:bold;">🥋 Tapis {t + 1}</th>')
    html.append('</tr></thead><tbody>')
    
    for row in grille_lignes:
        html.append('<tr>')
        for t in range(nb_tapis):
            col_name = f"Tapis {t + 1}"
            val = row.get(col_name, "")
            if not val:
                html.append('<td style="background:#FFFFFF; border:1px solid #E2E8F0; padding:8px;"></td>')
                continue
            val_s = str(val)
            if "PAUSE" in val_s:
                bg = "#EF4135"
                color = "#FFFFFF"
                border_c = "#DC2626"
                weight = "bold"
            elif any(k in val_s for k in ["Attente", "Pesée", "échauffement", "Repos"]):
                bg = "#F1F5F9"
                color = "#64748B"
                border_c = "#CBD5E1"
                weight = "normal"
            else:
                age = extraire_age_de_texte(val_s)
                cfg = COULEURS_AGE_GRILLE.get(age, COULEURS_AGE_GRILLE['AUTRE'])
                bg = cfg['hex_pastel']
                color = "#0F172A"
                border_c = cfg['hex_header']
                weight = "normal"
            
            cell_content = val_s.replace("\n", "<br>")
            html.append(f'<td style="background:{bg}; color:{color}; border:1px solid #CBD5E1; border-top:3px solid {border_c}; padding:8px 6px; text-align:center; font-size:12px; font-weight:{weight}; vertical-align:middle; line-height:1.4;">{cell_content}</td>')
        html.append('</tr>')
        
    html.append('</tbody></table></div>')
    return "\n".join(html)

# --- GÉNÉRATEUR DE DOCUMENTS HTML AUTONOMES POUR IMPRESSION PAYSAGE A4 ---
def generer_document_html_imprimable(titre, nom_comp, sections):
    """
    Génère un document HTML 100% autonome prêt pour l'impression A4 Paysage.
    Chaque section / onglet occupe sa propre page (page-break-after: always) et les tables ne sont jamais coupées.
    """
    html_sections = []
    for section_title, content in sections:
        if isinstance(content, pd.DataFrame):
            if "Grille Globale" in section_title or "Grille de Passage" in section_title:
                table_html = generer_html_grille_coloree(content.to_dict('records'), len(content.columns))
            else:
                table_html = content.to_html(index=False, classes="print-table")
        else:
            table_html = str(content)
        
        html_sections.append(f"""
        <div class="block-table">
            <h3 class="block-title">{section_title}</h3>
            {table_html}
        </div>
        """)
    
    sections_str = "\n".join(html_sections)
    
    html_doc = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>{titre} - {nom_comp}</title>
    <style>
        @page {{
            size: landscape;
            margin: 10mm;
        }}
        body {{
            font-family: Arial, Helvetica, sans-serif;
            margin: 0;
            padding: 15px;
            background: #ffffff;
            color: #111;
        }}
        .header-print {{
            text-align: center;
            border-bottom: 3px solid #0055A4;
            padding-bottom: 12px;
            margin-bottom: 25px;
        }}
        .header-print h1 {{
            color: #0055A4;
            margin: 0 0 6px 0;
            font-size: 24px;
            text-transform: uppercase;
        }}
        .header-print p {{
            margin: 0;
            color: #555;
            font-size: 13px;
        }}
        .block-table {{
            page-break-after: always !important;
            break-after: page !important;
            page-break-inside: avoid !important;
            break-inside: avoid-page !important;
            margin-bottom: 30px;
            width: 100%;
            clear: both;
        }}
        .block-table:last-child {{
            page-break-after: auto !important;
            break-after: auto !important;
        }}
        .block-title {{
            background-color: #0055A4;
            color: #ffffff;
            padding: 8px 14px;
            font-size: 15px;
            font-weight: bold;
            border-radius: 4px 4px 0 0;
            margin: 0 0 5px 0;
        }}
        table, .print-table {{
            width: 100% !important;
            border-collapse: collapse;
            margin-top: 0;
            margin-bottom: 10px;
            page-break-inside: avoid !important;
            break-inside: avoid-page !important;
        }}
        th {{
            background-color: #EF4135;
            color: #ffffff;
            padding: 8px 12px;
            font-size: 13px;
            border: 1px solid #000;
            text-align: center;
        }}
        td {{
            padding: 7px 12px;
            font-size: 12px;
            border: 1px solid #ccc;
            text-align: center;
        }}
        tr:nth-child(even) {{
            background-color: #f8f9fa;
        }}
        .no-print-bar {{
            text-align: center;
            padding: 12px;
            background-color: #f0f4f8;
            border: 1px solid #d0d7de;
            border-radius: 8px;
            margin-bottom: 20px;
        }}
        .btn-imprimer {{
            background-color: #0055A4;
            color: white;
            font-size: 15px;
            font-weight: bold;
            padding: 10px 24px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            box-shadow: 0 3px 6px rgba(0,0,0,0.15);
        }}
        .btn-imprimer:hover {{
            background-color: #003f7d;
        }}
        @media print {{
            .no-print-bar {{
                display: none !important;
            }}
        }}
    </style>
</head>
<body onload="window.print()">
    <div class="no-print-bar">
        <button class="btn-imprimer" onclick="window.print()">🖨️ Imprimer le Document (Format Paysage A4)</button>
    </div>
    <div class="header-print">
        <h1>🏆 {nom_comp}</h1>
        <p><strong>{titre}</strong> — Document Officiel FFLDA — Édité le {datetime.now().strftime('%d/%m/%Y à %H:%M')}</p>
    </div>
    {sections_str}
</body>
</html>"""
    return html_doc

# --- GÉNÉRATEUR DE DOCUMENTS PDF VECTORIELS DEPUIS LE CLASSEUR EXCEL OFFICIEL (A4 PORTRAIT) ---
def generer_pdf_depuis_classeur_excel(workbook_or_sheets, nom_competition="Tournoi FFLDA", wb=None):
    """
    Génère un fichier PDF vectoriel A4 Portrait à partir des feuilles du classeur Excel officiel FFLDA (openpyxl).
    Chaque feuille Excel (Poule nordique, Tableau éliminatoire, Poules croisées, Plateaux U7,
    Grille de Tapis, Planning) est fidèlement convertie en page(s) PDF en mode Portrait avec :
    - La disposition exacte des cellules et fusions (SPAN)
    - Les largeurs de colonnes proportionnelles adaptées au mode Portrait A4
    - Les couleurs de fond officielles FFLDA
    - Les bordures de cellules
    - Les styles de police (gras, tailles, alignements)
    - La résolution intelligente des formules et liens inter-onglets
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab import rl_config
    rl_config.allowTableBoundsErrors = 1

    wb_ref = wb
    if wb_ref is None:
        if hasattr(workbook_or_sheets, 'worksheets'):
            sheets = list(workbook_or_sheets.worksheets)
            wb_ref = workbook_or_sheets
        elif hasattr(workbook_or_sheets, 'sheets'):
            sheets = list(workbook_or_sheets.sheets.values()) if isinstance(workbook_or_sheets.sheets, dict) else list(workbook_or_sheets.sheets)
            wb_ref = getattr(workbook_or_sheets, 'book', None)
        elif isinstance(workbook_or_sheets, (list, tuple)):
            sheets = list(workbook_or_sheets)
            if sheets and hasattr(sheets[0], 'parent'):
                wb_ref = sheets[0].parent
        else:
            sheets = [workbook_or_sheets]
            if hasattr(workbook_or_sheets, 'parent'):
                wb_ref = workbook_or_sheets.parent
    else:
        if hasattr(workbook_or_sheets, 'worksheets'):
            sheets = list(workbook_or_sheets.worksheets)
        elif hasattr(workbook_or_sheets, 'sheets'):
            sheets = list(workbook_or_sheets.sheets.values()) if isinstance(workbook_or_sheets.sheets, dict) else list(workbook_or_sheets.sheets)
        elif isinstance(workbook_or_sheets, (list, tuple)):
            sheets = list(workbook_or_sheets)
        else:
            sheets = [workbook_or_sheets]

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=A4, # Format A4 Portrait
        rightMargin=12, 
        leftMargin=12, 
        topMargin=12, 
        bottomMargin=12
    )
    styles = getSampleStyleSheet()
    page_width = A4[0] - 24   # 571.27 pt
    page_height = A4[1] - 24  # 817.89 pt

    def to_col_letter(col_idx):
        res = ""
        while col_idx > 0:
            col_idx, rem = divmod(col_idx - 1, 26)
            res = chr(65 + rem) + res
        return res

    def get_hex(openpyxl_color, default=None):
        if not openpyxl_color:
            return default
        rgb = getattr(openpyxl_color, 'rgb', None)
        if not rgb:
            return default
        rgb_s = str(rgb).strip()
        if rgb_s in ['00000000', '0', 'None', '', 'none']:
            return default
        if re.match(r'^[0-9a-fA-F]{6}$', rgb_s):
            return '#' + rgb_s
        elif re.match(r'^[0-9a-fA-F]{8}$', rgb_s):
            return '#' + rgb_s[2:]
        return default

    def safe_hex_color(hex_str, default='#CBD5E1'):
        try:
            if hex_str and re.match(r'^#[0-9a-fA-F]{6}$', str(hex_str)):
                return colors.HexColor(hex_str)
        except Exception:
            pass
        return colors.HexColor(default)

    def resolve_formula_cell(formula_str, current_ws, wb_reference, depth=0):
        if depth > 4:
            return ""
        s = str(formula_str).strip()
        if not s.startswith('='):
            return s
            
        # Détection podiums et classements
        if '🥇' in s or 'CHAMPION' in s or 'OR' in s:
            quoted = re.findall(r'"([^"]*)"', s)
            p_parts = [q for q in quoted if any(k in q for k in ['🥇', 'OR', 'CHAMPION', 'En attente', 'Vainqueur'])]
            if p_parts:
                return "\n".join(p_parts[:2])
        if '🥈' in s or 'ARGENT' in s or 'VICE' in s:
            quoted = re.findall(r'"([^"]*)"', s)
            p_parts = [q for q in quoted if any(k in q for k in ['🥈', 'ARGENT', 'VICE', 'En attente', 'Perdant'])]
            if p_parts:
                return "\n".join(p_parts[:2])
        if '🥉' in s or 'BRONZE' in s or 'Bronze' in s:
            quoted = re.findall(r'"([^"]*)"', s)
            p_parts = [q for q in quoted if any(k in q for k in ['🥉', 'BRONZE', 'Bronze', 'En attente', 'Vainqueur'])]
            if p_parts:
                return "\n".join(p_parts[:2])

        quoted = re.findall(r'"([^"]*)"', s)
        priority_keywords = ['🔴', '🔵', 'Vainqueur', 'Perdant', 'Qualifié', 'Repêché', 'TOUR', 'COMBAT', 'Plateau']
        for q in quoted:
            if any(k in q for k in priority_keywords):
                return q
                
        try:
            m_ref = re.search(r"(?:'([^']+)'|([A-Za-z0-9_]+))!([A-Z]+[0-9]+)", s)
            if m_ref and wb_reference:
                target_sheet_name = m_ref.group(1) or m_ref.group(2)
                coord = m_ref.group(3)
                if hasattr(wb_reference, 'sheetnames') and target_sheet_name in wb_reference.sheetnames:
                    cell_obj = wb_reference[target_sheet_name][coord]
                    t_val = getattr(cell_obj, 'value', None)
                    res = resolve_formula_cell(t_val, wb_reference[target_sheet_name], wb_reference, depth + 1)
                    if res:
                        return res.replace('🔴 ', '').replace('🔵 ', '').strip()
        except Exception:
            pass
                    
        try:
            m_local = re.match(r"^=([A-Z]+[0-9]+)$", s)
            if m_local and current_ws:
                coord = m_local.group(1)
                cell_obj = current_ws[coord]
                t_val = getattr(cell_obj, 'value', None)
                return resolve_formula_cell(t_val, current_ws, wb_reference, depth + 1)
        except Exception:
            pass
            
        if quoted:
            for q in quoted:
                if q.strip() and q not in [" ", "🔴 ", "🔵 "]:
                    return q
        return ""

    def clean_val(val, current_ws, row=None, col=None):
        if val is None or val == "":
            if row is not None and col is not None and wb_ref is not None:
                try:
                    s_name = current_ws.title
                    if hasattr(wb_ref, 'sheetnames') and s_name in wb_ref.sheetnames:
                        ref_cell = wb_ref[s_name].cell(row=row, column=col)
                        f_val = getattr(ref_cell, 'value', None)
                        if f_val and str(f_val).strip().startswith('='):
                            return resolve_formula_cell(f_val, current_ws, wb_ref)
                except Exception:
                    pass
            return ""
        s = str(val).strip()
        if s.startswith('='):
            return resolve_formula_cell(s, current_ws, wb_ref)
        return s

    def cell_has_content(c_obj):
        if c_obj.value not in [None, '']:
            return True
        f = getattr(c_obj, 'fill', None)
        if getattr(f, 'fill_type', None) in ['solid']:
            c_hex = get_hex(getattr(f, 'fgColor', None))
            if c_hex and c_hex.upper() not in ['#FFFFFF', '#00000000']:
                return True
        b = getattr(c_obj, 'border', None)
        if b and (getattr(b.left, 'style', None) or getattr(b.top, 'style', None) or getattr(b.right, 'style', None) or getattr(b.bottom, 'style', None)):
            return True
        return False

    story = []
    sheet_has_pages = False

    for sheet_idx, ws in enumerate(sheets):
        if not hasattr(ws, 'cell'):
            continue
        max_r = ws.max_row or 1
        max_c = ws.max_column or 1
        
        min_allowed_r = 1
        min_allowed_c = 1
        for mr in list(ws.merged_cells.ranges):
            if mr.max_row > min_allowed_r:
                min_allowed_r = min(mr.max_row, max_r)
            if mr.max_col > min_allowed_c:
                min_allowed_c = min(mr.max_col, max_c)
                
        while max_r > min_allowed_r and all(not cell_has_content(ws.cell(row=max_r, column=c)) for c in range(1, max_c + 1)):
            max_r -= 1
        while max_c > min_allowed_c and all(not cell_has_content(ws.cell(row=r, column=max_c)) for r in range(1, max_r + 1)):
            max_c -= 1
            
        if max_r == 1 and max_c == 1 and not cell_has_content(ws.cell(1, 1)):
            continue

        col_widths = []
        for c in range(1, max_c + 1):
            col_letter = to_col_letter(c)
            w = ws.column_dimensions[col_letter].width if col_letter in ws.column_dimensions else None
            try:
                w_val = float(w) if (w is not None and str(w).strip() != '') else 10.0
                col_widths.append(w_val if w_val > 0 else 10.0)
            except (ValueError, TypeError):
                col_widths.append(10.0)
            
        total_w = sum(col_widths)
        scale = page_width / total_w if total_w > 0 else 1.0
        scaled_widths = [max(cw * scale, 5.0) for cw in col_widths]
        sum_sw = sum(scaled_widths)
        if sum_sw > page_width:
            scaled_widths = [sw * (page_width / sum_sw) for sw in scaled_widths]

        data = []
        t_styles = [
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 1.5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 1.5),
            ('TOPPADDING', (0, 0), (-1, -1), 1.0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1.0),
        ]

        base_scale = 0.50 if max_c > 14 else (0.60 if max_c > 9 else 0.70)

        for r in range(1, max_r + 1):
            row_cells = []
            for c in range(1, max_c + 1):
                cell = ws.cell(row=r, column=c)
                raw = cell.value
                txt = clean_val(raw, ws, row=r, col=c)
                
                font = cell.font
                is_bold = bool(font.bold) if font else False
                orig_f_size = 10
                if font and getattr(font, 'size', None) is not None:
                    try:
                        orig_f_size = float(font.size)
                    except (ValueError, TypeError):
                        orig_f_size = 10
                
                if r == 1:
                    f_size = max(9.5, min(int(orig_f_size * 0.9), 13))
                elif r == 2:
                    f_size = max(7.0, min(int(orig_f_size * 0.85), 9))
                else:
                    f_size = max(4.5, min(int(orig_f_size * base_scale), 11))
                
                f_color = get_hex(getattr(font, 'color', None), default='#1E293B')
                if not f_color or f_color in ['#00000000', '#000000']:
                    f_color = '#1E293B'
                    
                align_obj = cell.alignment
                h_align = getattr(align_obj, 'horizontal', 'center') or 'center'
                align_code = 1 # Center
                if h_align == 'left': align_code = 0
                elif h_align == 'right': align_code = 2
                
                fill_obj = getattr(cell, 'fill', None)
                if getattr(fill_obj, 'fill_type', None) in ['solid', 'lightGrid', 'darkGrid']:
                    bg_hex = get_hex(getattr(fill_obj, 'fgColor', None))
                    if bg_hex and bg_hex.upper() not in ['#FFFFFF', '#00000000']:
                        t_styles.append(('BACKGROUND', (c - 1, r - 1), (c - 1, r - 1), safe_hex_color(bg_hex)))
                        if bg_hex.upper() in ['#0055A4', '#EF4135', '#E53935', '#000000', '#334155', '#475569', '#1E88E5']:
                            f_color = '#FFFFFF'
                
                if (not fill_obj or not getattr(fill_obj, 'fill_type', None)) and f_color == '#FFFFFF':
                    f_color = '#1E293B'
                    
                b = cell.border
                if b and (getattr(b.left, 'style', None) or getattr(b.top, 'style', None) or getattr(b.right, 'style', None) or getattr(b.bottom, 'style', None)):
                    b_col = get_hex(getattr(b.top, 'color', None), default='#CBD5E1')
                    t_styles.append(('BOX', (c - 1, r - 1), (c - 1, r - 1), 0.5, safe_hex_color(b_col)))
                    
                p_style = ParagraphStyle(
                    f'S_{sheet_idx}_{r}_{c}',
                    parent=styles['Normal'],
                    fontName='Helvetica-Bold' if is_bold else 'Helvetica',
                    fontSize=f_size,
                    leading=f_size + 1.2,
                    textColor=safe_hex_color(f_color, default='#1E293B'),
                    alignment=align_code
                )
                
                txt_clean = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', txt)
                txt_safe = txt_clean.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('\n', '<br/>')
                try:
                    p = Paragraph(txt_safe, p_style) if txt_safe else Paragraph("&nbsp;", p_style)
                except Exception:
                    p = Paragraph("&nbsp;", p_style)
                row_cells.append(p)
            data.append(row_cells)

        spanned_cells = set()
        for m_range in list(ws.merged_cells.ranges):
            min_col, min_row, max_col_r, max_row_r = m_range.min_col, m_range.min_row, m_range.max_col, m_range.max_row
            if min_row <= max_r and min_col <= max_c:
                end_c = min(max_col_r, max_c)
                end_r = min(max_row_r, max_r)
                if (end_c > min_col or end_r > min_row) and min_col >= 1 and min_row >= 1:
                    c1 = min_col - 1
                    r1 = min_row - 1
                    c2 = min(end_c - 1, len(data[0]) - 1)
                    r2 = min(end_r - 1, len(data) - 1)
                    if c2 >= c1 and r2 >= r1 and (c2 > c1 or r2 > r1):
                        span_coords = [(col_k, row_k) for col_k in range(c1, c2 + 1) for row_k in range(r1, r2 + 1)]
                        if not any(coord in spanned_cells for coord in span_coords):
                            t_styles.append(('SPAN', (c1, r1), (c2, r2)))
                            for coord in span_coords:
                                spanned_cells.add(coord)

        # Auto-fusion horizontale pour lignes d'en-tête / bannières non fusionnées
        for r_chk in range(1, min(max_r + 1, 5)):
            r_idx = r_chk - 1
            if (0, r_idx) not in spanned_cells:
                v1 = ws.cell(row=r_chk, column=1).value
                if v1 and str(v1).strip():
                    last_empty_c = 1
                    for c_chk in range(2, max_c + 1):
                        if (c_chk - 1, r_idx) in spanned_cells or cell_has_content(ws.cell(row=r_chk, column=c_chk)):
                            break
                        last_empty_c = c_chk
                    if last_empty_c > 1:
                        c2_span = last_empty_c - 1
                        span_coords = [(col_k, r_idx) for col_k in range(0, c2_span + 1)]
                        if not any(coord in spanned_cells for coord in span_coords):
                            t_styles.append(('SPAN', (0, r_idx), (c2_span, r_idx)))
                            for coord in span_coords:
                                spanned_cells.add(coord)
                
        safe_styles = []
        for cmd in t_styles:
            try:
                op = cmd[0]
                if op in ('BACKGROUND', 'BOX', 'SPAN'):
                    sc1, sr1 = cmd[1]
                    sc2, sr2 = cmd[2]
                    if 0 <= sc1 < max_c and 0 <= sc2 < max_c and 0 <= sr1 < max_r and 0 <= sr2 < max_r:
                        if sc2 >= sc1 and sr2 >= sr1:
                            safe_styles.append(cmd)
                else:
                    safe_styles.append(cmd)
            except Exception:
                pass

        rep_rows = 2 if ('tapis' in ws.title.lower() or 'passage' in ws.title.lower()) and max_r > 4 else 0
        try:
            table = Table(data, colWidths=scaled_widths, repeatRows=rep_rows, splitByRow=1)
            table.setStyle(TableStyle(safe_styles))
        except Exception:
            table = Table(data, colWidths=scaled_widths, splitByRow=1)
            table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
            ]))

        if sheet_has_pages:
            story.append(PageBreak())
        story.append(table)
        sheet_has_pages = True

    try:
        doc.build(story)
        return buffer.getvalue()
    except Exception:
        # Fallback d'urgence
        try:
            buf_fb = io.BytesIO()
            doc_fb = SimpleDocTemplate(buf_fb, pagesize=A4, rightMargin=12, leftMargin=12, topMargin=12, bottomMargin=12)
            fb_story = []
            for item in story:
                if isinstance(item, Table):
                    fb_t = Table(item._cellvalues, colWidths=item._colWidths, splitByRow=1)
                    fb_t.setStyle(TableStyle([
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
                    ]))
                    fb_story.append(fb_t)
                else:
                    fb_story.append(item)
            doc_fb.build(fb_story)
            return buf_fb.getvalue()
        except Exception:
            return None

# --- GÉNÉRATEUR DE DOCUMENTS PDF VECTORIELS (REPORTLAB - A4 PORTRAIT - 1 PAGE PAR ONGLET) ---
def generer_pdf_tournoi_complet(titre, nom_comp, sections):
    """
    Génère un fichier PDF vectoriel (A4 Portrait) prêt à imprimer et télécharger.
    Chaque onglet / section commence sur une nouvelle page (PageBreak) et aucun tableau n'est coupé.
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=A4, 
        rightMargin=20, 
        leftMargin=20, 
        topMargin=20, 
        bottomMargin=20
    )
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'PDFTitle', 
        parent=styles['Heading1'], 
        fontName='Helvetica-Bold', 
        fontSize=16, 
        textColor=colors.HexColor('#0055A4'), 
        alignment=1,
        spaceAfter=6
    )
    
    subtitle_style = ParagraphStyle(
        'PDFSubtitle', 
        parent=styles['Normal'], 
        fontName='Helvetica', 
        fontSize=10, 
        textColor=colors.HexColor('#555555'), 
        alignment=1,
        spaceAfter=12
    )

    sec_banner_style = ParagraphStyle(
        'SecBanner', 
        parent=styles['Heading2'], 
        fontName='Helvetica-Bold', 
        fontSize=12, 
        textColor=colors.white, 
        backColor=colors.HexColor('#0055A4'), 
        borderPadding=6,
        spaceAfter=10,
        alignment=0
    )
    
    cell_head_style = ParagraphStyle(
        'CellHead', 
        parent=styles['Normal'], 
        fontName='Helvetica-Bold', 
        fontSize=9, 
        textColor=colors.white, 
        alignment=1
    )
    
    cell_body_style = ParagraphStyle(
        'CellBody', 
        parent=styles['Normal'], 
        fontName='Helvetica', 
        fontSize=8, 
        textColor=colors.black, 
        alignment=1
    )

    story = []
    page_width = A4[0] - 40  # 555.27 pt

    for idx, (sec_title, content) in enumerate(sections):
        if idx > 0:
            story.append(PageBreak())
        
        story.append(Paragraph(f"🏆 {nom_comp.upper()}", title_style))
        story.append(Paragraph(f"<b>{titre}</b> — Édité le {datetime.now().strftime('%d/%m/%Y à %H:%M')}", subtitle_style))
        story.append(Paragraph(f"<b>{sec_title}</b>", sec_banner_style))
        story.append(Spacer(1, 8))
        
        if isinstance(content, pd.DataFrame):
            df = content
            if df.empty:
                continue
            
            headers = [Paragraph(str(col), cell_head_style) for col in df.columns]
            data = [headers]
            
            for _, row in df.iterrows():
                r_cells = []
                for val in row:
                    txt = str(val).replace('\n', '<br/>') if pd.notna(val) else ''
                    r_cells.append(Paragraph(txt, cell_body_style))
                data.append(r_cells)
            
            nb_cols = len(df.columns)
            col_w = page_width / nb_cols if nb_cols > 0 else page_width
            col_widths = [col_w] * nb_cols
            
            table_obj = Table(data, colWidths=col_widths, repeatRows=1)
            t_styles = [
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0055A4')),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ]
            if "Grille Globale" in sec_title or "Grille de Passage" in sec_title:
                for r_idx, (_, row_data) in enumerate(df.iterrows(), 1):
                    for c_idx, val in enumerate(row_data):
                        if pd.notna(val) and val:
                            val_str = str(val)
                            if "PAUSE" in val_str:
                                t_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#EF4135')))
                            elif any(k in val_str for k in ["Attente", "Pesée", "échauffement", "Repos"]):
                                t_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#EFEFEF')))
                            else:
                                age_k = extraire_age_de_texte(val_str)
                                hex_col = COULEURS_AGE_GRILLE.get(age_k, COULEURS_AGE_GRILLE['AUTRE'])['hex_pastel']
                                t_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor(hex_col)))
            else:
                t_styles.append(('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8F9FA')]))
            table_obj.setStyle(TableStyle(t_styles))
            story.append(KeepTogether(table_obj))
        else:
            story.append(Paragraph(str(content), cell_body_style))

    doc.build(story)
    return buffer.getvalue()

# --- STYLES D'IMPRESSION DIRECTE (CSS @media print) ---
st.markdown("""
    <style>
    @media print {
        @page {
            size: landscape;
            margin: 8mm;
        }

        html, body, .stApp, .main, .block-container, div[data-testid="stMain"], div[data-testid="stBlock"] {
            width: 100% !important;
            max-width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
            background: white !important;
            color: black !important;
            overflow: visible !important;
        }

        table, .stTable, div[data-testid="stTable"], .stDataFrame, div[data-testid="stDataFrame"] {
            width: 100% !important;
            max-width: 100% !important;
            table-layout: auto !important;
            page-break-inside: avoid !important;
            break-inside: avoid-page !important;
            break-inside: avoid !important;
        }

        tr, tbody, thead {
            page-break-inside: avoid !important;
            break-inside: avoid-page !important;
            break-inside: avoid !important;
        }

        section[data-testid="stSidebar"], 
        header, 
        footer, 
        button, 
        .stButton,
        iframe,
        div[data-baseweb="tab-list"] {
            display: none !important;
        }
    }
    </style>
""", unsafe_allow_html=True)

def bouton_imprimer(html_data=None, filename="Fiche_Impression_Paysage.html", label="🖨️ Imprimer / Télécharger Fiche Paysage (HTML)", key=None):
    if html_data:
        st.download_button(
            label=label,
            data=html_data,
            file_name=filename,
            mime="text/html",
            key=key
        )
    else:
        print_code = f"""
            <button onclick="window.parent.print()" style="
                background-color: #0055A4; 
                color: white; 
                border: none; 
                padding: 10px 18px; 
                font-size: 14px; 
                font-weight: bold; 
                border-radius: 6px; 
                cursor: pointer;
                box-shadow: 0px 2px 5px rgba(0,0,0,0.2);
            ">
                {label}
            </button>
        """
        components.html(print_code, height=50)

def nettoyer_rang(val):
    """Extrait le rang numérique sous forme d'entier (1, 2, 3...) ou 'NR'."""
    if val is None:
        return "NR"
    s = str(val).strip()
    if not s or s in ["NR", "None", "nan", "<NA>", "NoneType", "-", "0"]:
        return "NR"
    try:
        v_num = int(float(s))
        return v_num if v_num > 0 else "NR"
    except (ValueError, TypeError):
        pass
    m = re.search(r'(\d+)', s)
    if m:
        v_num = int(m.group(1))
        return v_num if v_num > 0 else "NR"
    return "NR"

def extraire_resultats_classeur_excel(wb_data, wb_formula=None):
    """
    Extrait l'intégralité des résultats et classements officiels directement depuis le classeur Excel
    fourni par l'utilisateur (format officiel FFLDA).
    Gère les Poules Nordiques, Tableaux à élimination directe U13, Poules Croisées et Plateaux U7.
    Extrait et calcule fidèlement les points de victoires et classements y compris en cas de formules non recalculées.
    """
    tous_les_resultats = []
    
    # Filtrer les onglets de compétition
    mots_exclus = [
        "résumé", "resume", "grille", "passage", "planning", "programme", 
        "tapis", "déroulement", "deroulement", "schedule", "recap", "récap", 
        "classement", "bilan", "general", "général", "liste", "inscrits", "club", "comité", "comite"
    ]
    onglets_poules = [
        f for f in wb_data.sheetnames 
        if not any(x in f.lower() for x in mots_exclus)
    ]
    
    # Tri officiel des onglets : U7 d'abord, puis U9, U11, U13
    onglets_poules.sort(key=lambda x: (
        0 if "u7" in x.lower() else (1 if "u9" in x.lower() else (2 if "u11" in x.lower() else 3)),
        x
    ))
    
    def formater_poids_local(val):
        if val is None or str(val).strip() in ['', 'None', 'nan']:
            return ''
        try:
            val_float = float(str(val).replace(',', '.'))
            if val_float == int(val_float):
                return int(val_float)
            else:
                return round(val_float, 1)
        except (ValueError, TypeError):
            return val

    def est_ligne_lutteur_valide(nom_val, club_val):
        if nom_val is None:
            return False
        nom_s = str(nom_val).strip().lower()
        club_s = str(club_val or "").strip().lower()
        if not nom_s or nom_s in ["none", "nan", "-", "", "0"]:
            return False
        invalides = [
            "nom", "nom prénom", "nom prenom", "nom & prénom", "nom/prénom", 
            "participant", "lutteur", "liste des participants", "n°", "clt", "rang",
            "club", "poids", "comité", "comite", "points", "total pts", "total vict",
            "indépendant", "independant", "horaire", "heure", "tapis"
        ]
        if nom_s in invalides:
            return False
        mots_interdits = [
            "catégorie", "poule", "compétition", "tournoi", "repos", "pause", 
            "pesée", "pesee", "échauffement", "echauffement", "matinée", "apres-midi", 
            "après-midi", "remise", "récompense", "recompense", "grille", "passage", 
            "tapis", "min", "h00", "h15", "h30", "h45", "00]", "15]", "30]", "45]"
        ]
        if any(m in nom_s for m in mots_interdits) or any(m in club_s for m in ["repos", "pause", "pesée", "échauffement"]):
            return False
        if club_s in ["club", "poids", "comité", "comite", "points", "total pts", "clt", "n°"]:
            return False
        if "kg" in nom_s or "kg" in club_s:
            return False
        return True

    def extraire_texte_podium(v_data, v_form, ws_target):
        txt_data = str(v_data or "").strip()
        if txt_data and not txt_data.startswith("=") and txt_data not in ["None", "nan"]:
            return txt_data
        txt_form = str(v_form or "").strip()
        if not txt_form or txt_form in ["None", "nan"]:
            return ""
        m_refs = re.findall(r'([A-Z]+[0-9]+)', txt_form)
        parts = []
        literals = re.findall(r'"([^"]*)"', txt_form)
        if literals:
            parts.extend(literals)
        for cell_ref in m_refs:
            try:
                val_ref = str(ws_target[cell_ref].value or "").strip()
                if val_ref and not val_ref.startswith("="):
                    parts.append(val_ref)
            except Exception:
                pass
        return " ".join(parts) if parts else txt_form

    def normaliser_nom_comparaison(txt):
        if not txt:
            return ""
        import unicodedata
        txt = re.sub(r'[🔴🔵🥇🥈🥉🏆🛡️]|\([^\)]*\)', ' ', str(txt))
        txt_nfkd = unicodedata.normalize('NFD', txt)
        txt_sans_accents = "".join([c for c in txt_nfkd if unicodedata.category(c) != 'Mn'])
        return re.sub(r'[^a-z0-9]', '', txt_sans_accents.lower())

    def comparer_nom_lutteur_podium(nom_lutteur, podium_txt):
        if not nom_lutteur or not podium_txt:
            return False
        norm_lutteur = normaliser_nom_comparaison(nom_lutteur)
        norm_podium = normaliser_nom_comparaison(podium_txt)
        if not norm_lutteur or not norm_podium:
            return False
        if norm_lutteur in norm_podium:
            return True
        mots_lutteur = [m for m in re.findall(r'[a-z0-9]+', str(nom_lutteur).lower()) if len(m) >= 2]
        if mots_lutteur and all(m in norm_podium for m in mots_lutteur):
            return True
        return False

    def extraire_valeur_numerique_cellule(cell_data_val, cell_form_val, ws_target=None, depth=0):
        """Extrait une valeur numérique d'une cellule de score (nombre, '4 pts', formule un-évaluée)."""
        if depth > 5:
            return 0.0

        if cell_data_val is not None:
            s_val = str(cell_data_val).strip()
            if s_val not in ['', 'None', 'nan', '-', 'NR']:
                s_num = s_val.replace(',', '.')
                try:
                    return float(s_num)
                except ValueError:
                    pass
                m_pts = re.search(r'(-?\d+(?:\.\d+)?)\s*pts?', s_num, re.IGNORECASE)
                if m_pts:
                    return float(m_pts.group(1))
                m_start = re.search(r'^(-?\d+(?:\.\d+)?)', s_num)
                if m_start:
                    return float(m_start.group(1))

        if cell_form_val is not None:
            s_form = str(cell_form_val).strip()
            if s_form.startswith('='):
                expr = s_form[1:].strip()
                
                # Formule de somme: =SUM(...) ou =SOMME(...)
                m_sum = re.match(r'^(?:SUM|SOMME)\s*\(\s*([A-Z]+[0-9]+)\s*[:;]\s*([A-Z]+[0-9]+)\s*\)$', expr, re.IGNORECASE)
                if m_sum and ws_target is not None:
                    try:
                        c_start, c_end = m_sum.group(1).upper(), m_sum.group(2).upper()
                        cells = ws_target[c_start:c_end]
                        sum_val = 0.0
                        for row_cells in cells:
                            for cell_item in row_cells:
                                v_sub = extraire_valeur_numerique_cellule(cell_item.value, None, ws_target, depth + 1)
                                sum_val += v_sub
                        return sum_val
                    except Exception:
                        pass

                # Addition simple de cellules: =E5+F5+G5
                if '+' in expr and ws_target is not None and not re.search(r'[()/*]', expr):
                    parts = expr.split('+')
                    sum_parts = 0.0
                    all_valid = True
                    for p_item in parts:
                        p_clean = p_item.strip().upper()
                        if re.match(r'^[A-Z]+[0-9]+$', p_clean):
                            v_sub = extraire_valeur_numerique_cellule(ws_target[p_clean].value, None, ws_target, depth + 1)
                            sum_parts += v_sub
                        elif re.match(r'^-?\d+(\.\d+)?$', p_clean):
                            sum_parts += float(p_clean)
                        else:
                            all_valid = False
                            break
                    if all_valid and sum_parts > 0:
                        return sum_parts

                # Référence directe de cellule: =G25 ou =H5
                m_ref = re.match(r'^([A-Z]+[0-9]+)$', expr, re.IGNORECASE)
                if m_ref and ws_target is not None:
                    try:
                        ref_cell_val = ws_target[m_ref.group(1).upper()].value
                        if ref_cell_val is not None:
                            return extraire_valeur_numerique_cellule(ref_cell_val, None, ws_target, depth + 1)
                    except Exception:
                        pass

                # Formule avec valeur numérique directe
                m_num = re.search(r'(-?\d+(?:\.\d+)?)', expr)
                if m_num and not re.search(r'[A-Z]', expr):
                    try:
                        return float(m_num.group(1))
                    except ValueError:
                        pass

        return 0.0

    def compter_victoires_tableau(ws_target, ws_form, nom_lutteur):
        """Compte le nombre de matchs remportés par un lutteur dans la grille de tableau à élimination directe."""
        if not nom_lutteur:
            return 0
        victoires = 0
        for r_v in range(1, ws_target.max_row + 1):
            for c_v in range(1, ws_target.max_column + 1):
                val_cell = str(ws_target.cell(row=r_v, column=c_v).value or "")
                if val_cell and (val_cell.startswith("🔴") or val_cell.startswith("🔵")):
                    if comparer_nom_lutteur_podium(nom_lutteur, val_cell):
                        c_score = c_v + 1
                        if c_score <= ws_target.max_column:
                            v_s_d = ws_target.cell(row=r_v, column=c_score).value
                            v_s_f = ws_form.cell(row=r_v, column=c_score).value if ws_form else None
                            pts_m = extraire_valeur_numerique_cellule(v_s_d, v_s_f, ws_target)
                            if pts_m > 0:
                                victoires += 1
        return victoires

    def resoudre_classement_deux_poules(ws_target, ws_form, lutteurs_poule):
        """
        Résout automatiquement les rangs officiels (1er à 6ème) d'une catégorie à 2 Poules (6 lutteurs)
        en analysant la Phase Finale Croisée (Demi-Finales 1 & 2, Finale 1-2 et Finale 3-4).
        """
        if len(lutteurs_poule) != 6:
            return

        has_phase2 = False
        for r in range(1, min(20, ws_target.max_row + 1)):
            for c in range(1, min(25, ws_target.max_column + 1)):
                v = str(ws_target.cell(row=r, column=c).value or "").upper()
                if "DEMI-FINALE" in v or "FINALE 1-2" in v or "PHASE FINALE" in v:
                    has_phase2 = True
                    break
            if has_phase2:
                break

        if not has_phase2:
            return

        poule_a = lutteurs_poule[:3]
        poule_b = lutteurs_poule[3:]

        deuxieme_a, deuxieme_b = None, None
        premier_a, premier_b = None, None

        for r in range(1, min(25, ws_target.max_row + 1)):
            for c in range(9, min(20, ws_target.max_column + 1)):
                val_cell = str(ws_target.cell(row=r, column=c).value or "").upper()
                if "DEMI-FINALE 1" in val_cell:
                    for r_sf in [r+1, r+2]:
                        txt_sf = str(ws_target.cell(row=r_sf, column=c).value or "")
                        if not txt_sf:
                            txt_sf = str(ws_target.cell(row=r_sf, column=c+1).value or "")
                        for p in poule_b:
                            if comparer_nom_lutteur_podium(p["Nom"], txt_sf):
                                deuxieme_b = p
                                break
                elif "DEMI-FINALE 2" in val_cell:
                    for r_sf in [r+1, r+2]:
                        txt_sf = str(ws_target.cell(row=r_sf, column=c).value or "")
                        if not txt_sf:
                            txt_sf = str(ws_target.cell(row=r_sf, column=c+1).value or "")
                        for p in poule_a:
                            if comparer_nom_lutteur_podium(p["Nom"], txt_sf):
                                deuxieme_a = p
                                break

        reste_a = [p for p in poule_a if p != deuxieme_a]
        if len(reste_a) == 2:
            v0 = compter_victoires_tableau(ws_target, ws_form, reste_a[0]["Nom"])
            v1 = compter_victoires_tableau(ws_target, ws_form, reste_a[1]["Nom"])
            if v0 >= v1:
                premier_a, troisieme_a = reste_a[0], reste_a[1]
            else:
                premier_a, troisieme_a = reste_a[1], reste_a[0]
        else:
            troisieme_a = None

        reste_b = [p for p in poule_b if p != deuxieme_b]
        if len(reste_b) == 2:
            v0 = compter_victoires_tableau(ws_target, ws_form, reste_b[0]["Nom"])
            v1 = compter_victoires_tableau(ws_target, ws_form, reste_b[1]["Nom"])
            if v0 >= v1:
                premier_b, troisieme_b = reste_b[0], reste_b[1]
            else:
                premier_b, troisieme_b = reste_b[1], reste_b[0]
        else:
            troisieme_b = None

        gold_winner = None
        bronze_winner = None

        for r in range(1, min(25, ws_target.max_row + 1)):
            for c in range(9, min(25, ws_target.max_column + 1)):
                v_cell = str(ws_target.cell(row=r, column=c).value or "").upper()
                if "FINALE 1-2" in v_cell:
                    for r_m in [r+1, r+2]:
                        score_cell = ws_target.cell(row=r_m, column=c+1).value if ws_target.cell(row=r_m, column=c).value else ws_target.cell(row=r_m, column=c+2).value
                        score_form = ws_form.cell(row=r_m, column=c+1).value if ws_form else None
                        score_val = extraire_valeur_numerique_cellule(score_cell, score_form, ws_target)
                        if score_val > 0:
                            txt_m = str(ws_target.cell(row=r_m, column=c).value or ws_target.cell(row=r_m, column=c+1).value or "")
                            if premier_a and ("POULE A" in txt_m.upper() or comparer_nom_lutteur_podium(premier_a["Nom"], txt_m)):
                                gold_winner = premier_a
                            elif premier_b and ("POULE B" in txt_m.upper() or comparer_nom_lutteur_podium(premier_b["Nom"], txt_m)):
                                gold_winner = premier_b
                elif "FINALE 3-4" in v_cell:
                    for r_m in [r+1, r+2]:
                        score_cell = ws_target.cell(row=r_m, column=c+1).value if ws_target.cell(row=r_m, column=c).value else ws_target.cell(row=r_m, column=c+2).value
                        score_form = ws_form.cell(row=r_m, column=c+1).value if ws_form else None
                        score_val = extraire_valeur_numerique_cellule(score_cell, score_form, ws_target)
                        if score_val > 0:
                            txt_m = str(ws_target.cell(row=r_m, column=c).value or ws_target.cell(row=r_m, column=c+1).value or "")
                            if deuxieme_a and comparer_nom_lutteur_podium(deuxieme_a["Nom"], txt_m):
                                bronze_winner = deuxieme_a
                            elif deuxieme_b and comparer_nom_lutteur_podium(deuxieme_b["Nom"], txt_m):
                                bronze_winner = deuxieme_b

        if gold_winner == premier_a:
            if premier_a: premier_a["Clt_Excel"] = 1
            if premier_b: premier_b["Clt_Excel"] = 2
        elif gold_winner == premier_b:
            if premier_b: premier_b["Clt_Excel"] = 1
            if premier_a: premier_a["Clt_Excel"] = 2
        else:
            if premier_a: premier_a["Clt_Excel"] = 1
            if premier_b: premier_b["Clt_Excel"] = 2

        if bronze_winner == deuxieme_a:
            if deuxieme_a: deuxieme_a["Clt_Excel"] = 3
            if deuxieme_b: deuxieme_b["Clt_Excel"] = 4
        elif bronze_winner == deuxieme_b:
            if deuxieme_b: deuxieme_b["Clt_Excel"] = 3
            if deuxieme_a: deuxieme_a["Clt_Excel"] = 4
        else:
            if deuxieme_a: deuxieme_a["Clt_Excel"] = 3
            if deuxieme_b: deuxieme_b["Clt_Excel"] = 4

        if troisieme_a: troisieme_a["Clt_Excel"] = 5
        if troisieme_b: troisieme_b["Clt_Excel"] = 6

    def determiner_points_lutteur(ws_target, ws_form, r_row, col_tot, col_start_t=5, h_row=4, nom_lutteur=""):
        if nom_lutteur:
            vics_tab = compter_victoires_tableau(ws_target, ws_form, nom_lutteur)
            if vics_tab > 0:
                return vics_tab

        pts = 0.0
        if col_tot:
            val_d = ws_target.cell(row=r_row, column=col_tot).value
            val_f = ws_form.cell(row=r_row, column=col_tot).value if ws_form else None
            pts = extraire_valeur_numerique_cellule(val_d, val_f, ws_target)

        if pts == 0.0:
            sum_tours = 0.0
            found_tours = False
            max_col_tours = col_tot if (col_tot and col_tot > col_start_t) else (ws_target.max_column + 1)
            for c_t in range(col_start_t, max_col_tours):
                val_header = str(ws_target.cell(row=h_row, column=c_t).value or "").strip().lower()
                if any(x in val_header for x in ["total", "clt", "rang", "place", "club", "comité", "poids", "vict"]):
                    continue
                v_d = ws_target.cell(row=r_row, column=c_t).value
                v_f = ws_form.cell(row=r_row, column=c_t).value if ws_form else None
                val_num = extraire_valeur_numerique_cellule(v_d, v_f, ws_target)
                if val_num > 0:
                    sum_tours += val_num
                    found_tours = True
            if found_tours:
                pts = sum_tours

        return int(round(pts))

    def collecter_stats_lutteur_nordique(wb, nom_lutteur, nom_poule):
        stats = {
            'vict': 0,
            'pt_clt': 0.0,
            'vt': 0,
            'vst': 0,
            'pts_marq': 0.0,
            'pts_conc': 0.0,
            'face_a_face': {}
        }
        for s_name in wb.sheetnames:
            if not s_name.lower().startswith("grille tapis"):
                continue
            ws_t = wb[s_name]
            for r_m in range(4, ws_t.max_row, 4):
                val_r = str(ws_t.cell(row=r_m, column=3).value or "")
                val_b = str(ws_t.cell(row=r_m, column=8).value or "")
                if not val_r and not val_b:
                    continue
                is_r = comparer_nom_lutteur_podium(nom_lutteur, val_r)
                is_b = comparer_nom_lutteur_podium(nom_lutteur, val_b)
                if not is_r and not is_b:
                    continue
                
                ptr_val = extraire_valeur_numerique_cellule(ws_t.cell(row=r_m, column=5).value, None, ws_t)
                typer_val = str(ws_t.cell(row=r_m, column=6).value or "").strip().upper()
                ptb_val = extraire_valeur_numerique_cellule(ws_t.cell(row=r_m, column=10).value, None, ws_t)
                typeb_val = str(ws_t.cell(row=r_m, column=11).value or "").strip().upper()
                
                tot_r = extraire_valeur_numerique_cellule(ws_t.cell(row=r_m+2, column=5).value, None, ws_t)
                tot_b = extraire_valeur_numerique_cellule(ws_t.cell(row=r_m+2, column=10).value, None, ws_t)
                
                is_r_win = typer_val in ['VT', 'VST', 'VP'] or typeb_val in ['DT', 'DST', 'DP']
                is_b_win = typeb_val in ['VT', 'VST', 'VP'] or typer_val in ['DT', 'DST', 'DP']
                
                if is_r:
                    stats['pt_clt'] += ptr_val
                    stats['pts_marq'] += tot_r
                    stats['pts_conc'] += tot_b
                    if "VT" in typer_val: stats['vt'] += 1
                    if "VST" in typer_val: stats['vst'] += 1
                    if is_r_win:
                        stats['vict'] += 1
                        stats['face_a_face'][val_b] = 1
                    elif is_b_win:
                        stats['face_a_face'][val_b] = -1
                    elif ptr_val > ptb_val or (ptr_val == ptb_val and tot_r > tot_b and tot_r > 0):
                        stats['vict'] += 1
                        stats['face_a_face'][val_b] = 1
                    elif ptb_val > ptr_val or (ptr_val == ptb_val and tot_b > tot_r and tot_b > 0):
                        stats['face_a_face'][val_b] = -1
                else:
                    stats['pt_clt'] += ptb_val
                    stats['pts_marq'] += tot_b
                    stats['pts_conc'] += tot_r
                    if "VT" in typeb_val: stats['vt'] += 1
                    if "VST" in typeb_val: stats['vst'] += 1
                    if is_b_win:
                        stats['vict'] += 1
                        stats['face_a_face'][val_r] = 1
                    elif is_r_win:
                        stats['face_a_face'][val_r] = -1
                    elif ptb_val > ptr_val or (ptr_val == ptb_val and tot_b > tot_r and tot_b > 0):
                        stats['vict'] += 1
                        stats['face_a_face'][val_r] = 1
                    elif ptr_val > ptb_val or (ptr_val == ptb_val and tot_r > tot_b and tot_r > 0):
                        stats['face_a_face'][val_r] = -1
        return stats

    for nom_feuille in onglets_poules:
        ws = wb_data[nom_feuille]
        ws_f = wb_formula[nom_feuille] if (wb_formula and nom_feuille in wb_formula.sheetnames) else None
        
        titre_feuille = ""
        for r_t in range(1, 4):
            val_t = str(ws.cell(row=r_t, column=1).value or "").strip()
            if val_t:
                titre_feuille = val_t
                break
        
        titre_lower = (titre_feuille + " " + nom_feuille).lower()
        
        # --- CAS A : PLATEAU U7 ---
        if ("u7" in nom_feuille.lower() or "plateau" in nom_feuille.lower()) or ("u7" in titre_lower and not any(k in titre_lower for k in ["u9", "u11", "u13"])):
            h_row = 4
            col_nom, col_club, col_poids = 2, 3, 4
            for r_search in [4, 5, 3, 2]:
                for c_idx in range(1, ws.max_column + 1):
                    val_h = str(ws.cell(row=r_search, column=c_idx).value or "").strip().lower()
                    if "nom" in val_h:
                        col_nom = c_idx
                        h_row = r_search
                    elif "club" in val_h:
                        col_club = c_idx
                    elif "poids" in val_h:
                        col_poids = c_idx
                if col_nom:
                    break
            
            r = h_row + 1
            while r <= ws.max_row:
                nom_raw = ws.cell(row=r, column=col_nom).value
                if nom_raw is None and r > h_row + 20:
                    break
                if nom_raw is not None:
                    nom = str(nom_raw).strip()
                    club = str(ws.cell(row=r, column=col_club).value or "Indépendant").strip() if col_club else "Indépendant"
                    poids = formater_poids_local(ws.cell(row=r, column=col_poids).value) if col_poids else ""
                    
                    if est_ligne_lutteur_valide(nom, club):
                        tous_les_resultats.append({
                            "Poule": nom_feuille,
                            "Nom": nom,
                            "Club": club if club not in ["", "-", "None"] else "Indépendant",
                            "Comité": "Comité Non Renseigné",
                            "Poids": poids,
                            "Points": 0,
                            "Clt": 1  # Tous récompensés en U7 FFLDA
                        })
                r += 1
            continue

        # --- CAS UNIVERSEL : SCANNER DYNAMIQUE (Poules Nordiques, Poules Croisées, Tableaux U13) ---
        h_row = None
        col_clt, col_nom, col_club, col_comite, col_total_pts, col_poids = None, None, None, None, None, None
        
        for r_search in range(1, min(15, ws.max_row + 1)):
            for c_idx in range(1, min(25, ws.max_column + 1)):
                val_h = str(ws.cell(row=r_search, column=c_idx).value or "").strip().lower()
                if any(k in val_h for k in ["compétition", "competition", "tournoi", "formule", "phase", "rencontres", "planning", "grille", "tapis", "déroulement", "deroulement", "programme"]) or " — " in str(ws.cell(row=r_search, column=c_idx).value or ""):
                    continue
                if any(k in val_h for k in ["club", "équipe", "equipe"]):
                    col_club = c_idx
                elif any(k in val_h for k in ["comité", "comite", "ligue", "région"]):
                    col_comite = c_idx
                elif any(k in val_h for k in ["poids", "kg"]):
                    col_poids = c_idx
                elif any(k in val_h for k in ["total pts", "total points", "pts total", "points total", "pts clt", "tot pts", "pts", "points"]):
                    if not any(k in val_h for k in ["tour", "t1", "t2", "t3", "t4", "t5"]):
                        col_total_pts = c_idx
                elif any(k in val_h for k in ["nom", "prénom", "prenom", "lutteur", "athlete"]):
                    if not col_nom:
                        col_nom = c_idx
                        h_row = r_search
            if col_nom and h_row:
                break

        if h_row:
            for c_idx in range(1, min(25, ws.max_column + 1)):
                val_h = str(ws.cell(row=h_row, column=c_idx).value or "").strip().lower()
                if any(k in val_h for k in ["clt", "rang", "classt", "classement", "place"]) and not any(k in val_h for k in ["n°", "no "]):
                    col_clt = c_idx
                    break

        if not h_row:
            h_row = 4

        # Repérage des cartes de podium si présentes (pour Tableaux et Poules Croisées)
        podium_cards = {}
        for r_c in range(1, min(ws.max_row + 1, 60)):
            for c_c in range(1, min(ws.max_column + 1, 40)):
                v_cell_data = ws.cell(row=r_c, column=c_c).value
                v_cell_form = ws_f.cell(row=r_c, column=c_c).value if ws_f else None
                v_cell = extraire_texte_podium(v_cell_data, v_cell_form, ws)
                if not v_cell or "en attente" in v_cell.lower():
                    continue
                if "🥇" in v_cell or ("CHAMPION" in v_cell.upper() and "VICE" not in v_cell.upper()):
                    podium_cards[1] = v_cell
                elif "🥈" in v_cell or "VICE-CHAMPION" in v_cell.upper() or "ARGENT" in v_cell.upper():
                    podium_cards[2] = v_cell
                elif "🥉" in v_cell or "BRONZE" in v_cell.upper():
                    if 3 not in podium_cards:
                        podium_cards[3] = [v_cell]
                    else:
                        if isinstance(podium_cards[3], list):
                            podium_cards[3].append(v_cell)
                        else:
                            podium_cards[3] = [podium_cards[3], v_cell]

        r = h_row + 1
        lutteurs_poule = []
        while r <= ws.max_row:
            nom_raw = ws.cell(row=r, column=col_nom).value
            if nom_raw is None and r > h_row + 25:
                break
            if nom_raw is not None:
                nom = str(nom_raw).strip()
                club = str(ws.cell(row=r, column=col_club).value or "Indépendant").strip() if col_club else "Indépendant"
                
                if est_ligne_lutteur_valide(nom, club):
                    comite_val = ws.cell(row=r, column=col_comite).value if col_comite else None
                    comite = str(comite_val).strip() if (comite_val and str(comite_val).strip() not in ["", "None", "nan", "-"]) else "Comité Non Renseigné"
                    
                    poids_raw = ws.cell(row=r, column=col_poids).value if col_poids else 0
                    poids_val = formater_poids_local(poids_raw)
                    if not poids_val or str(poids_val).strip() in ["0", "0 kg", "0kg", "-", "None", "nan"]:
                        m_p = re.search(r'(\+?\d+(?:[\.,]\d+)?\s*kg)', nom_feuille, re.IGNORECASE)
                        if m_p:
                            poids_val = m_p.group(1).strip()
                            if not poids_val.lower().endswith("kg"):
                                poids_val += " kg"
                    
                    pts_val = determiner_points_lutteur(ws, ws_f, r, col_total_pts, col_start_t=5, h_row=h_row, nom_lutteur=nom)
                    
                    clt_raw = ws.cell(row=r, column=col_clt).value if col_clt else None
                    clt_cleaned = nettoyer_rang(clt_raw)
                    clt_val = clt_cleaned if isinstance(clt_cleaned, int) else None

                    # Priorité absolue aux cartes podium / médailles si présentes (Tableaux U13 et 2 Poules)
                    if podium_cards:
                        podium_match = None
                        if 1 in podium_cards:
                            if comparer_nom_lutteur_podium(nom, podium_cards[1]):
                                podium_match = 1
                        if podium_match is None and 2 in podium_cards:
                            if comparer_nom_lutteur_podium(nom, podium_cards[2]):
                                podium_match = 2
                        if podium_match is None and 3 in podium_cards:
                            b_list = podium_cards[3] if isinstance(podium_cards[3], list) else [podium_cards[3]]
                            for b_txt in b_list:
                                if comparer_nom_lutteur_podium(nom, b_txt):
                                    podium_match = 3
                                    break

                        if podium_match is not None:
                            clt_val = podium_match
                        else:
                            clt_val = None

                    tours_dict = {}
                    col_vict_found = None
                    for c_t in range(1, min(ws.max_column + 1, 35)):
                        v_h = str(ws.cell(row=h_row, column=c_t).value or "").strip()
                        if v_h.lower().startswith("tour ") or v_h.lower() == "tour":
                            v_td = ws.cell(row=r, column=c_t).value
                            v_tf = ws_f.cell(row=r, column=c_t).value if ws_f else None
                            val_t_num = extraire_valeur_numerique_cellule(v_td, v_tf, ws)
                            tours_dict[v_h] = int(val_t_num) if val_t_num is not None and str(val_t_num) != "" else "-"
                        elif "vict" in v_h.lower():
                            col_vict_found = c_t

                    vict_num = 0
                    if col_vict_found:
                        v_vd = ws.cell(row=r, column=col_vict_found).value
                        v_vf = ws_f.cell(row=r, column=col_vict_found).value if ws_f else None
                        val_v = extraire_valeur_numerique_cellule(v_vd, v_vf, ws)
                        if val_v is not None and str(val_v) != "":
                            vict_num = int(val_v)

                    dict_lut = {
                        "Poule": nom_feuille,
                        "Nom": nom,
                        "Club": club if club not in ["", "-", "None"] else "Indépendant",
                        "Comité": comite,
                        "Poids": poids_val,
                        "Points": pts_val,
                        "Total Vict": vict_num,
                        "Clt_Excel": clt_val
                    }
                    dict_lut.update(tours_dict)
                    lutteurs_poule.append(dict_lut)
            r += 1

        if lutteurs_poule:
            resoudre_classement_deux_poules(ws, ws_f, lutteurs_poule)
            has_explicit_clt = any(p.get("Clt_Excel") is not None for p in lutteurs_poule)
            if not has_explicit_clt and len(lutteurs_poule) > 1:
                import functools
                for p_idx, p in enumerate(lutteurs_poule, 1):
                    p["_tirage"] = p_idx
                    st_l = collecter_stats_lutteur_nordique(wb_data, p["Nom"], nom_feuille)
                    p["_vict"] = st_l["vict"]
                    p["_pt_clt"] = max(float(p.get("Points", 0)), float(st_l["pt_clt"]))
                    p["_vt"] = st_l["vt"]
                    p["_vst"] = st_l["vst"]
                    p["_pts_marq"] = st_l["pts_marq"]
                    p["_pts_conc"] = st_l["pts_conc"]
                    p["_h2h"] = st_l["face_a_face"]

                def f_cmp_fflda(p_a, p_b):
                    # 1. Nombre de victoires
                    if p_a["_vict"] != p_b["_vict"]:
                        return -1 if p_a["_vict"] > p_b["_vict"] else 1
                    
                    # 2. Si 2 lutteurs à égalité : départage direct par le résultat de leur rencontre
                    for nom_adv, res_h in p_a.get("_h2h", {}).items():
                        if comparer_nom_lutteur_podium(p_b["Nom"], nom_adv):
                            if res_h == 1: return -1
                            elif res_h == -1: return 1
                    
                    # 3. Plus grand nombre de points de classement obtenus
                    if p_a["_pt_clt"] != p_b["_pt_clt"]:
                        return -1 if p_a["_pt_clt"] > p_b["_pt_clt"] else 1
                    # 4. Plus grand nombre de victoires par tombé (VT)
                    if p_a["_vt"] != p_b["_vt"]:
                        return -1 if p_a["_vt"] > p_b["_vt"] else 1
                    # 5. Plus grand nombre de victoires par supériorité technique (VST)
                    if p_a["_vst"] != p_b["_vst"]:
                        return -1 if p_a["_vst"] > p_b["_vst"] else 1
                    # 6. Plus grand nombre de points techniques marqués (Total Score)
                    if p_a["_pts_marq"] != p_b["_pts_marq"]:
                        return -1 if p_a["_pts_marq"] > p_b["_pts_marq"] else 1
                    # 7. Plus petit nombre de points techniques concédés
                    if p_a["_pts_conc"] != p_b["_pts_conc"]:
                        return -1 if p_a["_pts_conc"] < p_b["_pts_conc"] else 1
                    # 8. Numéro de tirage au sort le plus bas
                    return -1 if p_a["_tirage"] < p_b["_tirage"] else 1

                lutteurs_poule.sort(key=functools.cmp_to_key(f_cmp_fflda))
                for rank_idx, p in enumerate(lutteurs_poule, 1):
                    p["Clt_Excel"] = rank_idx

            sorted_p = sorted(
                lutteurs_poule, 
                key=lambda x: (
                    x["Clt_Excel"] if x.get("Clt_Excel") is not None else 999, 
                    -x.get("Points", 0)
                )
            )
            
            cur_rank = 1
            for idx, p in enumerate(sorted_p):
                if p.get("Clt_Excel") is not None:
                    p["Clt"] = p["Clt_Excel"]
                    try:
                        cur_rank = max(cur_rank, int(p["Clt_Excel"]) + 1)
                    except (ValueError, TypeError):
                        pass
                else:
                    if idx > 0 and sorted_p[idx-1].get("Clt") is not None and str(sorted_p[idx-1].get("Clt")) != "NR" and p.get("Points", 0) == sorted_p[idx-1].get("Points", 0) and p.get("Points", 0) > 0:
                        p["Clt"] = sorted_p[idx-1]["Clt"]
                    else:
                        p["Clt"] = cur_rank
                        cur_rank += 1
                    
            tous_les_resultats.extend(lutteurs_poule)

    # Toujours rechercher si un onglet récapitulatif ("Classements Individuels", "Bilan", "Classement Général") existe
    def extraire_depuis_feuille_classement_individuel(ws_target, ws_form=None):
        results = []
        h_row = None
        
        for r_s in range(1, min(15, ws_target.max_row + 1)):
            c_nom_found, c_clt_found, c_club_found = None, None, None
            for c_i in range(1, min(20, ws_target.max_column + 1)):
                v_h = str(ws_target.cell(row=r_s, column=c_i).value or "").strip().lower()
                if any(k in v_h for k in ["club", "équipe", "equipe"]):
                    c_club_found = c_i
                elif any(k in v_h for k in ["clt", "rang", "place"]):
                    c_clt_found = c_i
                elif any(k in v_h for k in ["nom", "prénom", "prenom", "lutteur"]) and "catégorie" not in v_h and "poule" not in v_h:
                    if not c_nom_found:
                        c_nom_found = c_i
                    
            if c_nom_found and (c_clt_found or c_club_found):
                h_row = r_s
                break

        if not h_row:
            return results

        col_clt, col_nom, col_club, col_comite, col_poids, col_pts = None, None, None, None, None, None
        for c_i in range(1, min(20, ws_target.max_column + 1)):
            v_h = str(ws_target.cell(row=h_row, column=c_i).value or "").strip().lower()
            if any(k in v_h for k in ["club", "équipe", "equipe"]):
                col_club = c_i
            elif any(k in v_h for k in ["clt", "rang", "place"]):
                col_clt = c_i
            elif any(k in v_h for k in ["comité", "comite", "ligue"]):
                col_comite = c_i
            elif "poids" in v_h:
                col_poids = c_i
            elif any(k in v_h for k in ["pts", "points", "total"]):
                col_pts = c_i
            elif any(k in v_h for k in ["nom", "prénom", "prenom", "lutteur"]):
                if not col_nom:
                    col_nom = c_i

        current_poule = ws_target.title
        for r_pre in range(1, h_row):
            v_pre = str(ws_target.cell(row=r_pre, column=1).value or "").strip()
            if any(k in v_pre.lower() for k in ["poule", "catégorie", "categorie"]):
                current_poule = re.sub(r'^\s*catégorie\s*/\s*poule\s*:\s*', '', v_pre, flags=re.IGNORECASE).strip()

        r = h_row + 1
        while r <= ws_target.max_row:
            v_cell1 = str(ws_target.cell(row=r, column=1).value or "").strip()
            if any(k in v_cell1.lower() for k in ["poule", "catégorie", "categorie"]):
                if not est_ligne_lutteur_valide(v_cell1, ""):
                    current_poule = re.sub(r'^\s*catégorie\s*/\s*poule\s*:\s*', '', v_cell1, flags=re.IGNORECASE).strip()
                    r += 1
                    continue

            nom_raw = ws_target.cell(row=r, column=col_nom).value if col_nom else None
            if nom_raw is not None:
                nom = str(nom_raw).strip()
                club = str(ws_target.cell(row=r, column=col_club).value or "Indépendant").strip() if col_club else "Indépendant"
                if est_ligne_lutteur_valide(nom, club):
                    clt_raw = ws_target.cell(row=r, column=col_clt).value if col_clt else None
                    clt_val = nettoyer_rang(clt_raw)
                    
                    comite_val = str(ws_target.cell(row=r, column=col_comite).value or "").strip() if col_comite else "Comité Non Renseigné"
                    if not comite_val or comite_val in ["None", "nan", "-"]:
                        comite_val = "Comité Non Renseigné"

                    poids_val = formater_poids_local(ws_target.cell(row=r, column=col_poids).value) if col_poids else ""
                    
                    pts_val = 0
                    if col_pts:
                        pts_val = determiner_points_lutteur(ws_target, ws_form, r, col_pts, col_start_t=5, h_row=h_row)

                    results.append({
                        "Poule": current_poule,
                        "Nom": nom,
                        "Club": club if club not in ["", "-", "None"] else "Indépendant",
                        "Comité": comite_val,
                        "Poids": poids_val,
                        "Points": pts_val,
                        "Clt": clt_val
                    })
            r += 1

        return results

    summary_results = []
    for s_name in wb_data.sheetnames:
        s_lower = s_name.lower()
        if any(k in s_lower for k in ["classements individuels", "classement individuel", "classement général", "bilan"]):
            if not any(k in s_lower for k in ["club", "comité", "comite", "résumé", "resume"]):
                res_sum = extraire_depuis_feuille_classement_individuel(wb_data[s_name], wb_formula[s_name] if (wb_formula and s_name in wb_formula.sheetnames) else None)
                if res_sum:
                    summary_results = res_sum
                    break

    if not tous_les_resultats:
        tous_les_resultats = summary_results
    elif summary_results:
        # Fusionner / enrichir avec les rangs explicites de la feuille récapitulative
        map_summary = {}
        for r_s in summary_results:
            k_name = normaliser_nom_comparaison(r_s.get("Nom", ""))
            if k_name:
                map_summary[k_name] = r_s

        for r_p in tous_les_resultats:
            k_p = normaliser_nom_comparaison(r_p.get("Nom", ""))
            if k_p in map_summary:
                s_item = map_summary[k_p]
                s_clt = nettoyer_rang(s_item.get("Clt"))
                if s_clt != "NR":
                    r_p["Clt"] = s_clt
                if s_item.get("Points", 0) > 0 and r_p.get("Points", 0) == 0:
                    r_p["Points"] = s_item["Points"]

    return tous_les_resultats


def trier_dataframe_bilan(df):
    """Trie numériquement le dataframe de bilan par Poule, puis par Rang (1er, 2ème, 3ème...), puis Points."""
    if df is None or df.empty:
        return df
    
    def key_clt(v):
        r_clean = nettoyer_rang(v)
        if isinstance(r_clean, int):
            return (0, r_clean)
        else:
            return (2, 999)

    df_copy = df.copy()
    df_copy["_sort_clt"] = df_copy["Clt"].apply(key_clt)
    df_copy["_sort_pts"] = pd.to_numeric(df_copy["Points"], errors="coerce").fillna(0)
    
    df_sorted = df_copy.sort_values(
        by=["Poule", "_sort_clt", "_sort_pts"], 
        ascending=[True, True, False]
    ).drop(columns=["_sort_clt", "_sort_pts"]).reset_index(drop=True)
    
    return df_sorted


def generer_pdf_fiche_qr_codes(nom_tournoi, code_acces, url_base, tapis_dispos, type_fiche="table"):
    import io
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib import colors
    from reportlab.graphics.barcode import qr
    from reportlab.graphics.shapes import Drawing
    from reportlab.graphics import renderPDF

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    # Header Bandeau FFLDA
    col_bandeau = "#0055A4" if type_fiche == "table" else "#0f172a"
    c.setFillColor(colors.HexColor(col_bandeau))
    c.rect(0, h - 70, w, 70, fill=True, stroke=False)
    
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(w / 2, h - 32, "FÉDÉRATION FRANÇAISE DE LUTTE (FFLDA)")
    c.setFont("Helvetica", 11)
    sous_titre_h = f"Accès Tables de Marque & Déroulé Général — Code : {code_acces}" if type_fiche == "table" else f"Scoreboards Officiels Grand Écran (TV / Tablette Tapis) — Code : {code_acces}"
    c.drawCentredString(w / 2, h - 52, sous_titre_h)

    # Titre compétition
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w / 2, h - 90, f"Compétition : {nom_tournoi}")
    c.setFont("Helvetica-Oblique", 9)
    c.setFillColor(colors.HexColor("#555555"))
    consigne = "Scannez avec l'appareil photo d'un smartphone pour ouvrir directement le Tapis" if type_fiche == "table" else "Scannez avec une Smart TV, tablette ou smartphone pour afficher le tableau en direct"
    c.drawCentredString(w / 2, h - 104, consigne)

    cle_sec_table = get_cle_secrete_table(code_acces)
    cartes = []
    if type_fiche == "table":
        for t_id in tapis_dispos:
            lien_t = f"{url_base}/?code={code_acces}&mode=direct&tapis={t_id}&cle={cle_sec_table}"
            cartes.append((f"TAPIS {t_id}", lien_t, colors.HexColor("#0055A4"), f"Table de Marque Tapis {t_id} (Sécurisé)"))
    else:
        for t_id in tapis_dispos:
            lien_scb = f"{url_base}/?code={code_acces}&mode=direct&tapis={t_id}&vue=scoreboard"
            cartes.append((f"📺 SCOREBOARD TAPIS {t_id}", lien_scb, colors.HexColor("#1e3a8a"), f"Scoreboard TV Tapis {t_id} (Haute Visibilité)"))

    lien_all = f"{url_base}/?code={code_acces}&mode=direct&tapis=all"
    cartes.append(("📋 DÉROULÉ (3 TAPIS)", lien_all, colors.HexColor("#B71C1C"), "Accès Public (Lecture Seule)"))

    margin_x = 40
    margin_y = 50
    card_w = (w - (2 * margin_x) - 20) / 2
    card_h = (h - 130 - margin_y - 20) / 2

    for idx, (titre, lien, col_theme, sous_titre) in enumerate(cartes[:4]):
        col_idx = idx % 2
        row_idx = idx // 2
        x = margin_x + col_idx * (card_w + 20)
        y = h - 130 - (row_idx + 1) * card_h

        # Boîte de la carte
        c.setStrokeColor(colors.HexColor("#DDDDDD"))
        c.setLineWidth(1)
        c.setFillColor(colors.HexColor("#FAFAFA"))
        c.roundRect(x, y, card_w, card_h, 8, fill=True, stroke=True)

        # Bandeau de titre
        c.setFillColor(col_theme)
        c.roundRect(x, y + card_h - 32, card_w, 32, 6, fill=True, stroke=False)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 12)
        c.drawCentredString(x + card_w / 2, y + card_h - 22, titre)

        # Sous-titre
        c.setFillColor(colors.HexColor("#333333"))
        c.setFont("Helvetica", 9)
        c.drawCentredString(x + card_w / 2, y + card_h - 45, sous_titre)

        # QR Code natif vectoriel
        qr_size = 130
        try:
            qr_widget = qr.QrCodeWidget(lien)
            bounds = qr_widget.getBounds()
            qw = bounds[2] - bounds[0]
            qh = bounds[3] - bounds[1]
            d = Drawing(qr_size, qr_size, transform=[qr_size/qw, 0, 0, qr_size/qh, 0, 0])
            d.add(qr_widget)
            renderPDF.draw(d, c, x + (card_w - qr_size) / 2, y + 42)
        except Exception:
            pass

        # Légende
        c.setFont("Courier", 7)
        c.setFillColor(colors.HexColor("#777777"))
        url_court = lien if len(lien) < 45 else lien[:42] + "..."
        c.drawCentredString(x + card_w / 2, y + 26, url_court)
        c.setFont("Helvetica-Bold", 8)
        c.setFillColor(col_theme)
        c.drawCentredString(x + card_w / 2, y + 12, "👉 Connexion automatique sans mot de passe")

    # Pied de page
    c.setFont("Helvetica", 8)
    c.setFillColor(colors.HexColor("#888888"))
    c.drawCentredString(w / 2, 25, "Document officiel FFLDA — Développé pour la gestion des compétitions fédérales")

    c.save()
    buf.seek(0)
    return buf.getvalue()


# --- GESTION DU CACHE PERSISTANT DE TOURNOI ---
def sauvegarder_cache_tournoi(code_org, bundle):
    try:
        os.makedirs(".cache_tournois", exist_ok=True)
        safe_code = re.sub(r'[^a-zA-Z0-9_-]', '_', str(code_org))
        with open(f".cache_tournois/cache_{safe_code}.pkl", "wb") as f:
            pickle.dump(bundle, f)
        if bundle.get("excel_bytes"):
            with open(f".cache_tournois/excel_base_{safe_code}.xlsx", "wb") as f_xl:
                f_xl.write(bundle["excel_bytes"])
    except Exception:
        pass

def charger_cache_tournoi(code_org):
    try:
        safe_code = re.sub(r'[^a-zA-Z0-9_-]', '_', str(code_org))
        p = f".cache_tournois/cache_{safe_code}.pkl"
        if os.path.exists(p):
            with open(p, "rb") as f:
                return pickle.load(f)
    except Exception:
        pass
    return None

def supprimer_cache_tournoi(code_org):
    try:
        safe_code = re.sub(r'[^a-zA-Z0-9_-]', '_', str(code_org))
        for p in [f".cache_tournois/cache_{safe_code}.pkl", f".cache_tournois/excel_base_{safe_code}.xlsx"]:
            if os.path.exists(p):
                os.remove(p)
    except Exception:
        pass


def calculer_resultats_depuis_matchs_direct(liste_matchs):
    """
    Calcule fidèlement l'intégralité des résultats et classements officiels FFLDA
    directement à partir des matchs saisis en direct (Tablettes / Supabase / Tapis).
    Garantit qu'aucune catégorie ni aucun lutteur n'est omis, applique le barème officiel FFLDA
    et résout les départages (victoires, confrontations directes, points de classement, points techniques).
    """
    if not liste_matchs:
        return []

    def extraire_poids_categorie_local(cat_str):
        m = re.search(r'\(([^)]*kg[^)]*)\)', cat_str, re.IGNORECASE)
        if m: return m.group(1).strip()
        m2 = re.search(r'(\d+(?:[\.,]\d+)?\s*kg)', cat_str, re.IGNORECASE)
        if m2: return m2.group(1).strip()
        return ""

    def tri_cat_officiel(c_name):
        c_up = c_name.upper()
        age_order = 0 if "U7" in c_up else (1 if "U9" in c_up else (2 if "U11" in c_up else 3))
        m_p = re.search(r'poule\s*(\d+)', c_name, re.IGNORECASE)
        p_num = int(m_p.group(1)) if m_p else 99
        return (age_order, p_num, c_name)

    cats = {}
    for m in liste_matchs:
        cat = str(m.get("categorie") or "Générale").strip()
        if not cat:
            cat = "Générale"
        if cat not in cats:
            cats[cat] = []
        cats[cat].append(m)

    tous_les_resultats = []

    for cat_nom in sorted(cats.keys(), key=tri_cat_officiel):
        matchs_cat = cats[cat_nom]
        poids_cat = extraire_poids_categorie_local(cat_nom)
        is_u7 = "U7" in cat_nom.upper()
        is_u13 = "U13" in cat_nom.upper()

        # Identifier tous les tours présents dans la catégorie
        tours_cat_set = set()
        for m in matchs_cat:
            t_raw = str(m.get("tour") or "").strip()
            m_num = re.search(r'\d+', t_raw)
            if m_num:
                tours_cat_set.add(int(m_num.group()))
        tours_poule_tries = sorted(list(tours_cat_set)) if tours_cat_set else [1]

        participants = {}
        for m in matchs_cat:
            for color in ["rouge", "bleu"]:
                nom = str(m.get(f"lutteur_{color}") or "").strip()
                club = str(m.get(f"club_{color}") or "").strip()
                comite = str(m.get(f"comite_{color}") or "").strip()
                if nom and nom not in ["-", "None", "nan", ""]:
                    if nom not in participants:
                        participants[nom] = {
                            "nom": nom,
                            "club": club if club and club not in ["-", "None", "nan"] else "Indépendant",
                            "comite": comite if comite and comite not in ["-", "None", "nan"] else "Comité Non Renseigné",
                            "poids": poids_cat,
                            "victoires": 0,
                            "pts_clt": 0,
                            "pts_tech_pour": 0,
                            "pts_tech_contre": 0,
                            "vt": 0,
                            "vst": 0,
                            "vp": 0,
                            "head_to_head": {},
                            "matchs_joues": 0,
                            "pts_tours": {t: "-" for t in tours_poule_tries}
                        }

        # Détection des finales éventuelles (Tableaux U13 ou poules croisées)
        finales_or = []
        finales_bronze = []

        for m in matchs_cat:
            t_raw = str(m.get("tour") or "").strip()
            m_num = re.search(r'\d+', t_raw)
            t_idx = int(m_num.group()) if m_num else 1

            if m.get("statut") == "Terminé":
                nr = str(m.get("lutteur_rouge") or "").strip()
                nb = str(m.get("lutteur_bleu") or "").strip()
                v = m.get("vainqueur")
                tv = str(m.get("type_victoire") or "VT")
                tour_m = str(m.get("tour") or "").upper()
                try: sr = int(float(str(m.get("score_rouge", 0) or 0)))
                except Exception: sr = 0
                try: sb = int(float(str(m.get("score_bleu", 0) or 0)))
                except Exception: sb = 0
                try: pr = int(float(str(m.get("pt_clt_rouge", 0) or 0)))
                except Exception: pr = 0
                try: pb = int(float(str(m.get("pt_clt_bleu", 0) or 0)))
                except Exception: pb = 0

                if pr == 0 and pb == 0 and v in ["Rouge", "Bleu"]:
                    pr, pb = calculer_pts_fflda_match(cat_nom, v, tv, sr, sb)

                if "FINALE 1" in tour_m or "FINALE D'OR" in tour_m:
                    finales_or.append((nr if v == "Rouge" else nb, nb if v == "Rouge" else nr))
                elif "FINALE 3" in tour_m or "BRONZE" in tour_m:
                    finales_bronze.append((nr if v == "Rouge" else nb, nb if v == "Rouge" else nr))

                if nr in participants:
                    participants[nr]["matchs_joues"] += 1
                    participants[nr]["pts_clt"] += pr
                    participants[nr]["pts_tech_pour"] += sr
                    participants[nr]["pts_tech_contre"] += sb
                    participants[nr]["pts_tours"][t_idx] = pr
                    if v == "Rouge":
                        participants[nr]["victoires"] += 1
                        participants[nr]["head_to_head"][nb] = "win"
                        if "VT" in tv: participants[nr]["vt"] += 1
                        elif "VST" in tv: participants[nr]["vst"] += 1
                        elif "VP" in tv: participants[nr]["vp"] += 1
                    elif v == "Bleu":
                        participants[nr]["head_to_head"][nb] = "loss"

                if nb in participants:
                    participants[nb]["matchs_joues"] += 1
                    participants[nb]["pts_clt"] += pb
                    participants[nb]["pts_tech_pour"] += sb
                    participants[nb]["pts_tech_contre"] += sr
                    participants[nb]["pts_tours"][t_idx] = pb
                    if v == "Bleu":
                        participants[nb]["victoires"] += 1
                        participants[nb]["head_to_head"][nr] = "win"
                        if "VT" in tv: participants[nb]["vt"] += 1
                        elif "VST" in tv: participants[nb]["vst"] += 1
                        elif "VP" in tv: participants[nb]["vp"] += 1
                    elif v == "Rouge":
                        participants[nb]["head_to_head"][nr] = "loss"

        liste_p = list(participants.values())
        if is_u7:
            for p in liste_p:
                dict_u7 = {
                    "Poule": cat_nom,
                    "Nom": p["nom"],
                    "Club": p["club"],
                    "Comité": p["comite"],
                    "Poids": p["poids"],
                    "Points": p["pts_clt"],
                    "Total Vict": p["victoires"],
                    "Clt": 1
                }
                for t in tours_poule_tries:
                    dict_u7[f"Tour {t}"] = p["pts_tours"].get(t, "-")
                tous_les_resultats.append(dict_u7)
        else:
            # Tri officiel FFLDA : Victoires, Pts Clt, VT, VST, Diff Tech, Pts Tech Pour
            def tri_cle(p):
                return (
                    p["victoires"],
                    p["pts_clt"],
                    p["vt"],
                    p["vst"],
                    p["pts_tech_pour"] - p["pts_tech_contre"],
                    p["pts_tech_pour"]
                )
            liste_p.sort(key=tri_cle, reverse=True)

            # Départage direct en cas d'égalité à 2 lutteurs
            for i in range(len(liste_p) - 1):
                p1 = liste_p[i]
                p2 = liste_p[i+1]
                if p1["victoires"] == p2["victoires"] and p1["pts_clt"] == p2["pts_clt"]:
                    if p2["head_to_head"].get(p1["nom"]) == "win":
                        liste_p[i], liste_p[i+1] = p2, p1

            # Remplacement éventuel par les finales 1-2 et 3-4
            rangs_attribues = {}
            if finales_or:
                w_or, l_or = finales_or[-1]
                rangs_attribues[w_or] = 1
                rangs_attribues[l_or] = 2
            if finales_bronze:
                w_br, l_br = finales_bronze[-1]
                rangs_attribues[w_br] = 3
                rangs_attribues[l_br] = 4

            for rank_num, p in enumerate(liste_p, 1):
                clt_final = rangs_attribues.get(p["nom"], rank_num)
                dict_p = {
                    "Poule": cat_nom,
                    "Nom": p["nom"],
                    "Club": p["club"],
                    "Comité": p["comite"],
                    "Poids": p["poids"],
                    "Points": p["pts_clt"],
                    "Total Vict": p["victoires"],
                    "Clt": clt_final
                }
                for t in tours_poule_tries:
                    dict_p[f"Tour {t}"] = p["pts_tours"].get(t, "-")
                tous_les_resultats.append(dict_p)

    return tous_les_resultats


def generer_classeur_complet_officiel(nom_competition, tous_les_resultats, liste_matchs):
    """
    Génère un classeur Excel officiel FFLDA complet et autonome
    contenant la grille des combats enregistrés et les feuilles de chaque poule/catégorie.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    bleu_fflda = PatternFill("solid", fgColor="0055A4")
    rouge_fflda = PatternFill("solid", fgColor="EF4135")
    gris_zebrage = PatternFill("solid", fgColor="F2F5F8")
    fond_blanc = PatternFill("solid", fgColor="FFFFFF")
    or_fill = PatternFill("solid", fgColor="FFF2CC")
    argent_fill = PatternFill("solid", fgColor="EFEFEF")
    bronze_fill = PatternFill("solid", fgColor="F8CBAD")

    font_titre = Font(name="Arial", size=14, bold=True, color="0055A4")
    font_entete = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Arial", size=10, color="000000")
    font_data_bold = Font(name="Arial", size=10, bold=True, color="000000")
    b_fin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                   top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    # 1. Feuille Grille des Combats
    ws_combats = wb.create_sheet("Résultats Combats")
    ws_combats.views.sheetView[0].showGridLines = True
    ws_combats.cell(row=1, column=1, value=f"COMPÉTITION : {nom_competition.upper()} — GRILLE OFFICIELLE DES COMBATS").font = font_titre
    headers_c = ["Tapis", "N° Match", "Catégorie", "Tour", "Statut", "Lutteur Rouge", "Club Rouge", "Score R", "Score B", "Lutteur Bleu", "Club Bleu", "Vainqueur", "Type Victoire", "Pts R", "Pts B"]
    for c_idx, h in enumerate(headers_c, 1):
        cell = ws_combats.cell(row=3, column=c_idx, value=h)
        cell.fill, cell.font, cell.alignment = bleu_fflda, font_entete, Alignment(horizontal="center", vertical="center")
    ws_combats.row_dimensions[3].height = 24

    for r_idx, m in enumerate(liste_matchs or [], 4):
        ws_combats.row_dimensions[r_idx].height = 20
        is_even = (r_idx % 2 == 0)
        vals = [
            m.get("tapis", 1),
            m.get("match_num", ""),
            m.get("categorie", ""),
            m.get("tour", ""),
            m.get("statut", ""),
            m.get("lutteur_rouge", ""),
            m.get("club_rouge", ""),
            m.get("score_rouge", 0),
            m.get("score_bleu", 0),
            m.get("lutteur_bleu", ""),
            m.get("club_bleu", ""),
            m.get("vainqueur", ""),
            m.get("type_victoire", ""),
            m.get("pt_clt_rouge", 0),
            m.get("pt_clt_bleu", 0)
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws_combats.cell(row=r_idx, column=c_idx, value=val)
            cell.border, cell.font = b_fin, font_data
            cell.fill = gris_zebrage if is_even else fond_blanc
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if c_idx in [6, 10]:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    for col in ws_combats.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_combats.column_dimensions[col_letter].width = max(max_len + 3, 10)

    # 2. Feuilles par catégorie (poules)
    cats_vus = {}
    for r in (tous_les_resultats or []):
        p = r.get("Poule", "Générale")
        if p not in cats_vus:
            cats_vus[p] = []
        cats_vus[p].append(r)

    feuilles_existantes = set(wb.sheetnames)
    for cat_nom, luts in cats_vus.items():
        nom_base = abreger_nom_onglet(cat_nom)
        nom_s = nom_base
        suf = 1
        while nom_s in feuilles_existantes:
            nom_s = f"{nom_base[:28]}_{suf}"
            suf += 1
        feuilles_existantes.add(nom_s)

        ws_cat = wb.create_sheet(nom_s)
        ws_cat.views.sheetView[0].showGridLines = True
        ws_cat.cell(row=1, column=1, value=f"{nom_competition.upper()} — {cat_nom}").font = font_titre
        
        # Collecter les colonnes de tours présentes dans les résultats de cette poule
        tours_in_luts = []
        for lut in luts:
            for k in lut.keys():
                if str(k).startswith("Tour ") and k not in tours_in_luts:
                    tours_in_luts.append(k)
        def _sort_t_fn(t_nom):
            try:
                return int(str(t_nom).split()[1])
            except Exception:
                return 999
        tours_in_luts.sort(key=_sort_t_fn)

        headers_p = ["Rang", "Nom Prénom", "Club", "Comité"] + tours_in_luts + ["Total Pts", "Total Vict", "Poids"]
        for c_idx, h in enumerate(headers_p, 1):
            cell = ws_cat.cell(row=3, column=c_idx, value=h)
            cell.fill, cell.font, cell.alignment = rouge_fflda, font_entete, Alignment(horizontal="center", vertical="center")
        ws_cat.row_dimensions[3].height = 22

        r_last = 3
        for r_idx, lut in enumerate(luts, 4):
            r_last = r_idx
            ws_cat.row_dimensions[r_idx].height = 20
            row_vals = [
                lut.get("Clt"),
                lut.get("Nom"),
                lut.get("Club"),
                lut.get("Comité", "")
            ]
            for ct in tours_in_luts:
                row_vals.append(lut.get(ct, "-"))
            row_vals.append(lut.get("Points", 0))
            row_vals.append(lut.get("Total Vict", 0))
            row_vals.append(lut.get("Poids", ""))

            for c_idx, val in enumerate(row_vals, 1):
                c_cell = ws_cat.cell(row=r_idx, column=c_idx, value=val)
                c_cell.border, c_cell.font = b_fin, font_data
                c_cell.alignment = Alignment(horizontal="center", vertical="center")
                if c_idx == 2:
                    c_cell.alignment = Alignment(horizontal="left", vertical="center")
                if c_idx == 1 or c_idx == (4 + len(tours_in_luts) + 1):
                    c_cell.font = font_data_bold
                clt_val = lut.get("Clt")
                if c_idx == 1:
                    if str(clt_val) == "1": c_cell.fill = or_fill
                    elif str(clt_val) == "2": c_cell.fill = argent_fill
                    elif str(clt_val) == "3": c_cell.fill = bronze_fill

        # Détail des combats par tour de la poule
        def _meme_cat_fn(c1, c2):
            s1 = re.sub(r'[^a-zA-Z0-9]', '', str(c1 or '').lower())
            s2 = re.sub(r'[^a-zA-Z0-9]', '', str(c2 or '').lower())
            return s1 == s2 or (len(s1) > 4 and (s1 in s2 or s2 in s1))

        matchs_cat = [
            m for m in (liste_matchs or [])
            if _meme_cat_fn(m.get("categorie"), cat_nom)
        ]
        if matchs_cat:
            r_m = r_last + 2
            ws_cat.cell(row=r_m, column=1, value="DÉTAIL DES COMBATS DE LA POULE").font = Font(name="Arial", size=11, bold=True, color="0055A4")
            r_m += 1
            headers_mc = ["N° Match", "Tour", "Lutteur Rouge", "Score R", "Score B", "Lutteur Bleu", "Vainqueur", "Type Victoire", "Pts Clt R", "Pts Clt B"]
            for c_idx, h in enumerate(headers_mc, 1):
                cell = ws_cat.cell(row=r_m, column=c_idx, value=h)
                cell.fill, cell.font, cell.alignment = bleu_fflda, font_entete, Alignment(horizontal="center", vertical="center")
            ws_cat.row_dimensions[r_m].height = 20
            
            matchs_cat_tries = sorted(matchs_cat, key=lambda x: (int(x.get("tour") or 1), int(x.get("match_num") or 0)))
            for m in matchs_cat_tries:
                r_m += 1
                ws_cat.row_dimensions[r_m].height = 18
                est_term = (m.get("statut") == "Terminé")
                v_nom = m.get("lutteur_rouge") if m.get("vainqueur") == "Rouge" else (m.get("lutteur_bleu") if m.get("vainqueur") == "Bleu" else m.get("vainqueur", ""))
                m_vals = [
                    m.get("match_num", ""),
                    f"Tour {m.get('tour', 1)}",
                    m.get("lutteur_rouge", ""),
                    m.get("score_rouge", 0) if est_term else "",
                    m.get("score_bleu", 0) if est_term else "",
                    m.get("lutteur_bleu", ""),
                    v_nom if est_term else "À venir",
                    (m.get("victoire_type") or m.get("type_victoire", "")) if est_term else "",
                    m.get("pt_clt_rouge", "") if est_term else "",
                    m.get("pt_clt_bleu", "") if est_term else ""
                ]
                for c_idx, val in enumerate(m_vals, 1):
                    c_cell = ws_cat.cell(row=r_m, column=c_idx, value=val)
                    c_cell.border, c_cell.font = b_fin, font_data
                    c_cell.alignment = Alignment(horizontal="center", vertical="center")
                    if c_idx in [3, 6, 7]:
                        c_cell.alignment = Alignment(horizontal="left", vertical="center")

        for col in ws_cat.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws_cat.column_dimensions[col_letter].width = max(max_len + 3, 10)

    return wb


def recuperer_excel_base(code_org, nom_tournoi="", categories_attendues=None):
    # 1. Directement dans session_state
    if st.session_state.get("excel_tournoi_base"):
        return st.session_state["excel_tournoi_base"]

    # 2. Dans le tournoi actif en cache mémoire
    if st.session_state.get("tournoi_actif_cache", {}).get("excel_bytes"):
        b = st.session_state["tournoi_actif_cache"]["excel_bytes"]
        st.session_state["excel_tournoi_base"] = b
        return b

    # 3. Dans le fichier dédié sur disque .cache_tournois/excel_base_{safe_code}.xlsx
    safe_code = re.sub(r'[^a-zA-Z0-9_-]', '_', str(code_org or "ORG"))
    p_xl = f".cache_tournois/excel_base_{safe_code}.xlsx"
    if os.path.exists(p_xl):
        try:
            with open(p_xl, "rb") as f:
                b = f.read()
                st.session_state["excel_tournoi_base"] = b
                return b
        except Exception:
            pass

    # 4. Dans le cache pickle
    cached = charger_cache_tournoi(code_org)
    if cached and cached.get("excel_bytes"):
        b = cached["excel_bytes"]
        st.session_state["excel_tournoi_base"] = b
        return b

    # 5. Recherche automatique intelligente dans les fichiers Tournoi_*.xlsx locaux récents
    try:
        import glob
        import zipfile
        import io
        import xml.etree.ElementTree as ET
        candidats = sorted(glob.glob("Tournoi_*.xlsx"), key=os.path.getmtime, reverse=True)
        for cand in candidats:
            try:
                with open(cand, "rb") as f:
                    content = f.read()
                with zipfile.ZipFile(io.BytesIO(content)) as z:
                    if 'xl/workbook.xml' in z.namelist():
                        if categories_attendues:
                            tree = ET.fromstring(z.read('xl/workbook.xml'))
                            sheets = [node.attrib['name'].lower() for node in tree.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet')]
                            match_count = 0
                            for cat in categories_attendues:
                                cat_l = str(cat).lower()
                                if any(s in cat_l or abreger_nom_onglet(cat).lower() == s for s in sheets):
                                    match_count += 1
                            if match_count == 0:
                                continue
                        st.session_state["excel_tournoi_base"] = content
                        try:
                            os.makedirs(".cache_tournois", exist_ok=True)
                            with open(p_xl, "wb") as f_save:
                                f_save.write(content)
                        except Exception:
                            pass
                        return content
            except Exception:
                continue
    except Exception:
        pass

    return None

def creer_bundle_depuis_matchs(nom_tournoi, matchs):
    if not matchs:
        return None
    try:
        tapis_set = sorted(list({int(m.get("tapis", 1)) for m in matchs}))
        nb_tapis = max(tapis_set) if tapis_set else 1
        
        planning_tapis = {t: [] for t in range(nb_tapis)}
        categories = set()
        lutteurs_vus = set()
        
        for m in matchs:
            t_idx = max(0, min(nb_tapis - 1, int(m.get("tapis", 1)) - 1))
            cat = m.get("categorie", "Générale")
            categories.add(cat)
            planning_tapis[t_idx].append({
                "Type": "MATCH",
                "Heure": m.get("heure", "09:00"),
                "Duree": m.get("duree", 2),
                "Cat": cat,
                "Nom_Tour": m.get("tour", ""),
                "Combattant 1": m.get("lutteur_rouge", ""),
                "Club 1": m.get("club_rouge", ""),
                "Comité 1": m.get("comite_rouge", ""),
                "Combattant 2": m.get("lutteur_bleu", ""),
                "Club 2": m.get("club_bleu", ""),
                "Comité 2": m.get("comite_bleu", ""),
                "Arbitre": m.get("arbitre", "Non attribué")
            })
            if m.get("lutteur_rouge"):
                lutteurs_vus.add((m.get("lutteur_rouge"), m.get("club_rouge", ""), cat))
            if m.get("lutteur_bleu"):
                lutteurs_vus.add((m.get("lutteur_bleu"), m.get("club_bleu", ""), cat))

        participants_par_poule = {}
        for lutteur, club, cat in lutteurs_vus:
            if cat not in participants_par_poule:
                participants_par_poule[cat] = []
            participants_par_poule[cat].append({"Nom": lutteur, "Club": club, "Poids": "-"})

        heures = [m.get("heure") for m in matchs if m.get("heure")]
        h_deb = min(heures) if heures else "09:00"
        h_fin = max(heures) if heures else "14:00"

        # Calcul d'horaires réalistes pour la pesée (45 min avant) et la pause
        try:
            h_dt = datetime.strptime(h_deb, "%H:%M")
            h_pesee_calc = (h_dt - timedelta(minutes=45)).strftime("%H:%M")
        except Exception:
            h_pesee_calc = "08:30"

        # Tentative de récupération des lignes personnalisées enregistrées sur disque
        lignes_accueil = None
        try:
            c_sess_rec = st.session_state.get("code_session", "ORG")
            safe_c_rec = re.sub(r'[^a-zA-Z0-9_-]', '_', str(c_sess_rec))
            p_res_json = f".cache_tournois/resume_{safe_c_rec}.json"
            if os.path.exists(p_res_json):
                with open(p_res_json, "r", encoding="utf-8") as f_rj:
                    lignes_accueil = json.load(f_rj)
        except Exception:
            lignes_accueil = None

        if not lignes_accueil:
            lignes_accueil = [
                {"Étape de la journée": "Pesée de la compétition", "Horaire / Valeur": f"{h_pesee_calc} - {h_deb}"},
                {"Étape de la journée": "Début de la compétition", "Horaire / Valeur": h_deb},
                {"Étape de la journée": "Compétition en cours", "Horaire / Valeur": f"{h_deb} - {h_fin}"},
                {"Étape de la journée": "Pause de la compétition", "Horaire / Valeur": "12:00 - 12:45"},
                {"Étape de la journée": "Total Catégories", "Horaire / Valeur": f"{len(categories)} catégories"},
                {"Étape de la journée": "Total Combats programmés", "Horaire / Valeur": f"{len(matchs)} combats"},
                {"Étape de la journée": "Fin estimée", "Horaire / Valeur": h_fin}
            ]

        return {
            "nom_competition": nom_tournoi,
            "lignes_accueil": lignes_accueil,
            "total_participants_peses": len(lutteurs_vus),
            "total_non_peses": 0,
            "poules_u7": [],
            "poules_u9": [c for c in categories if "U9" in c.upper()],
            "poules_u11": [c for c in categories if "U11" in c.upper()],
            "poules_u13": [c for c in categories if "U13" in c.upper()],
            "total_matchs_calcules": len(matchs),
            "liste_arbitres": [],
            "tapis_arbitres": {t: [] for t in range(nb_tapis)},
            "nb_tapis": nb_tapis,
            "grille_ui": [],
            "planning_tapis": planning_tapis,
            "participants_par_poule": participants_par_poule,
            "poule_obj_map": {},
            "excel_bytes": None,
            "pdf_bytes": None,
            "matchs_direct": matchs
        }
    except Exception:
        return None

# ==============================================================================
# MOTEUR UNIFIÉ DE CHRONOGRAMME DYNAMIQUE & CASCADE HORAIRE FFLDA
# Permet de recalculer automatiquement l'ensemble de la compétition en aval
# dès qu'un organisateur ajuste l'heure de pesée, les durées, les pauses ou l'ordre.
# ==============================================================================

def nommer_tour_match_global(poule_obj, r_idx, m_idx, m=None):
    if not poule_obj:
        return f"Tour {r_idx + 1}"
    type_f = poule_obj.get('type_formule', 'poule')
    rondes = poule_obj.get('rondes', [])
    total_r = len(rondes)
    
    if type_f == 'tableau':
        if total_r >= 6:
            if r_idx == 0: return f"1/32 de Finale ({m_idx + 1})"
            elif r_idx == 1: return f"1/16 de Finale ({m_idx + 1})"
            elif r_idx == 2: return f"1/8 de Finale ({m_idx + 1})"
            elif r_idx == 3: return f"1/4 de Finale ({m_idx + 1})"
            elif r_idx == total_r - 2:
                return f"Demi-Finale {m_idx + 1}" if m_idx in (0, 1) else f"Repêchage 1/4 ({m_idx - 1})"
            elif r_idx == total_r - 1:
                if m_idx == 0: return "Grande Finale (Or / Argent)"
                elif m_idx == 1: return "Finale Bronze 1"
                elif m_idx == 2: return "Finale Bronze 2"
                return f"Finale {m_idx + 1}"
            return f"Tour {r_idx + 1}"
        elif total_r == 5:
            if r_idx == 0: return f"1/16 de Finale ({m_idx + 1})"
            elif r_idx == 1: return f"1/8 de Finale ({m_idx + 1})"
            elif r_idx == 2: return f"1/4 de Finale ({m_idx + 1})"
            elif r_idx == total_r - 2:
                return f"Demi-Finale {m_idx + 1}" if m_idx in (0, 1) else f"Repêchage 1/4 ({m_idx - 1})"
            elif r_idx == total_r - 1:
                if m_idx == 0: return "Grande Finale (Or / Argent)"
                elif m_idx == 1: return "Finale Bronze 1"
                elif m_idx == 2: return "Finale Bronze 2"
                return f"Finale {m_idx + 1}"
            return f"Tour {r_idx + 1}"
        elif total_r == 4:
            if r_idx == 0: return f"Tour Préliminaire 1/8 ({m_idx + 1})"
            elif r_idx == 1: return f"1/4 de Finale ({m_idx + 1})"
            elif r_idx == 2:
                return f"Demi-Finale {m_idx + 1}" if m_idx in (0, 1) else f"Repêchage 1/4 ({m_idx - 1})"
            elif r_idx == 3:
                if m_idx == 0: return "Grande Finale (Or / Argent)"
                elif m_idx == 1: return "Finale Bronze 1"
                elif m_idx == 2: return "Finale Bronze 2"
                return f"Finale {m_idx + 1}"
            return f"Tour {r_idx + 1}"
        elif total_r == 3:
            if r_idx == 0: return f"1/4 de Finale ({m_idx + 1})"
            elif r_idx == 1:
                return f"Demi-Finale {m_idx + 1}" if m_idx in (0, 1) else f"Repêchage 1/4 ({m_idx - 1})"
            elif r_idx == 2:
                if m_idx == 0: return "Grande Finale (Or / Argent)"
                elif m_idx == 1: return "Finale Bronze 1"
                elif m_idx == 2: return "Finale Bronze 2"
                return f"Finale {m_idx + 1}"
            return f"Tour {r_idx + 1}"
        return f"Tour {r_idx + 1}"
            
    elif type_f == 'poules_croisees':
        if r_idx == 4:
            return "Grande Finale (Or / Argent)" if m_idx == 0 else "Finale 3-4 (Bronze unique)"
        elif r_idx == 3:
            return f"Demi-Finale Croisée {m_idx + 1}"
        return f"Poule - Tour {r_idx + 1}"
            
    return f"Tour {r_idx + 1}"

def choisir_arbitre_match_global(c1_club, c2_club, t_idx, tapis_arbitres=None, liste_arbitres=None, ref_usage_count=None, eviter_meme_club=True):
    cand_list = (tapis_arbitres or {}).get(t_idx, [])
    if not cand_list and liste_arbitres:
        cand_list = liste_arbitres
    if not cand_list:
        return "Non attribué"
    
    c1_clean = str(c1_club).strip().lower()
    c2_clean = str(c2_club).strip().lower()
    ignored_clubs = ['', '-', 'indépendant', 'independant', 'none', 'nan']
    
    if eviter_meme_club:
        sans_conflit = []
        for arb in cand_list:
            arb_club_clean = str(arb.get('Club', '')).strip().lower()
            if arb_club_clean in ignored_clubs or (arb_club_clean != c1_clean and arb_club_clean != c2_clean):
                sans_conflit.append(arb)
        pool_choix = sans_conflit if sans_conflit else cand_list
    else:
        pool_choix = cand_list
        
    if ref_usage_count is None:
        ref_usage_count = {}
    arb_choisi = min(pool_choix, key=lambda a: ref_usage_count.get(a.get('Nom_Complet', ''), 0))
    nom_arb = arb_choisi.get('Nom_Complet', 'Arbitre')
    ref_usage_count[nom_arb] = ref_usage_count.get(nom_arb, 0) + 1
    return nom_arb

def extraire_age_de_texte_chronogramme(nom_p):
    parts = str(nom_p).split('|')
    cat_prefix = parts[0].strip().upper() if parts else ''
    for a in ['U7', 'U9', 'U11', 'U13', 'U15', 'U17', 'U20', 'SENIOR']:
        if a in cat_prefix:
            return a
    for a in ['U7', 'U9', 'U11', 'U13', 'U15', 'U17', 'U20', 'SENIOR']:
        if a in str(nom_p).upper():
            return a
    return 'AUTRE'

def obtenir_duree_combat_par_categorie_global(nom_cat, durees_config=None, defaut=3):
    if not durees_config:
        durees_config = {"U7": 10, "U9": 3, "U11": 4, "U13": 5}
    age = extraire_age_de_texte_chronogramme(nom_cat)
    return durees_config.get(age, defaut)

def construire_grille_ui_depuis_planning(planning_tapis, nb_tapis):
    max_lignes = max((len(planning_tapis.get(t, [])) for t in range(nb_tapis)), default=0)
    grille_ui = []
    for row_idx in range(max_lignes):
        ligne = {}
        for t in range(nb_tapis):
            col = f"Tapis {t + 1}"
            plan_t = planning_tapis.get(t, [])
            if row_idx < len(plan_t):
                m = plan_t[row_idx]
                m_type = m.get("Type", "MATCH")
                if m_type == "PAUSE":
                    ligne[col] = f"[{m.get('Heure', '')}] ⏸️ PAUSE"
                elif m_type == "ATTENTE":
                    ligne[col] = f"[{m.get('Heure', '')}] {m.get('Texte', '')}"
                elif m_type == "VIDE":
                    ligne[col] = ""
                else:
                    arb_str = f" (Arbitre : {m['Arbitre']})" if m.get('Arbitre') and m['Arbitre'] != "Non attribué" else ""
                    t_nom = nettoyer_nom_tour(m.get('Nom_Tour') or m.get('Tour') or "")
                    tour_str = f" [🎯 Tour : {t_nom}]" if t_nom else ""
                    ligne[col] = f"[{m.get('Heure', '')}] ({m.get('Duree', 2)}m) [{m.get('Cat', '')}]{tour_str} - {m.get('Combattant 1', '')} vs {m.get('Combattant 2', '')}{arb_str}"
            else:
                ligne[col] = ""
        grille_ui.append(ligne)
    return grille_ui

def repartir_tapis_categories(categories, poules_dict, nb_tapis):
    """
    Répartit les nb_tapis (0 à nb_tapis-1) entre les différentes catégories d'une session simultanée (+).
    Chaque catégorie reçoit une liste d'indices de tapis [t_start, ..., t_end].
    Garantit que deux catégories d'âge distinctes luttent simultanément sur des tapis distincts.
    """
    if not categories:
        return {}
    if len(categories) == 1 or nb_tapis == 1:
        return {c: list(range(nb_tapis)) for c in categories}
    
    # Calcul de la charge (nombre de matchs ou de rotations) par catégorie
    poids_cat = {}
    for c in categories:
        poules = poules_dict.get(c, [])
        if c == "U7":
            nb_m = 3 * max(1, len(poules))
        else:
            nb_m = sum(sum(len(r) for r in p.get('rondes', [])) for p in poules if isinstance(p, dict))
            if nb_m == 0:
                nb_m = max(1, len(poules))
        poids_cat[c] = max(1, nb_m)
    
    total_poids = sum(poids_cat.values())
    
    if nb_tapis <= len(categories):
        alloc = {}
        for idx, c in enumerate(categories):
            t_idx = idx % nb_tapis
            alloc.setdefault(c, []).append(t_idx)
        return alloc

    tapis_par_cat = {c: 1 for c in categories}
    restants = nb_tapis - len(categories)
    
    fractions = []
    for c in categories:
        q = (poids_cat[c] / total_poids) * restants
        part_entiere = int(q)
        tapis_par_cat[c] += part_entiere
        fractions.append((q - part_entiere, c))
    
    restants_apres_entiere = nb_tapis - sum(tapis_par_cat.values())
    fractions.sort(key=lambda x: x[0], reverse=True)
    for i in range(restants_apres_entiere):
        tapis_par_cat[fractions[i][1]] += 1
        
    res = {}
    curr_t = 0
    for c in categories:
        nb_t = tapis_par_cat[c]
        res[c] = list(range(curr_t, curr_t + nb_t))
        curr_t += nb_t
        
    return res

def ordonnancer_phase_multi(poules_phase, heure_debut_phase, durees_config=None, nb_tapis=3, repos_matchs=3, meme_tapis=True, tapis_arbitres=None, liste_arbitres=None, ref_usage_count=None, eviter_meme_club=True, tapis_cibles=None):
    if not poules_phase:
        return [heure_debut_phase for _ in range(nb_tapis)], {t: [] for t in range(nb_tapis)}
    
    if tapis_cibles is None:
        tapis_cibles = list(range(nb_tapis))

    if durees_config is None:
        durees_config = {"U7": 10, "U9": 3, "U11": 4, "U13": 5}
    if ref_usage_count is None:
        ref_usage_count = {}

    planning = {t: [] for t in range(nb_tapis)}
    tapis_heure = [heure_debut_phase for _ in range(nb_tapis)]
    last_match_time = {}

    def obtenir_cle_poids_lot(p_obj):
        nom_p = p_obj.get('nom', '')
        parts_nom = [x.strip() for x in nom_p.split('|')]
        age_p = parts_nom[0] if parts_nom else 'U13'
        style_p = p_obj.get('style_grp') or (parts_nom[1] if len(parts_nom) >= 2 else '')
        cat_p = p_obj.get('cat_poids')
        if cat_p:
            return f"{age_p} | {style_p} | {str(cat_p).strip()}"
        m_poids = re.search(r'\b(\+?\d+\s*kg)\b', nom_p, re.IGNORECASE)
        if m_poids:
            return f"{age_p} | {style_p} | {m_poids.group(1).strip()}"
        return nom_p

    # 1. Regroupement par lot (poids / style / poule)
    lots_poids = {}
    for p in poules_phase:
        cle = obtenir_cle_poids_lot(p) if meme_tapis else p.get('nom', '')
        if cle not in lots_poids:
            lots_poids[cle] = []
        lots_poids[cle].append(p)

    # 2. Nombre de matchs par lot
    lots_avec_poids = []
    for cle, liste_p in lots_poids.items():
        nb_m = sum(sum(len(r) for r in p.get('rondes', [])) for p in liste_p)
        lots_avec_poids.append((cle, liste_p, nb_m))

    # 3. Tri LPT décroissant pour équilibrer parfaitement les tapis
    lots_avec_poids.sort(key=lambda x: x[2], reverse=True)

    # 4. Affectation équilibrée sur les tapis cibles désignés
    tapis_poules = {t: [] for t in range(nb_tapis)}
    charge_matchs_tapis = [0] * nb_tapis

    for cle, liste_p, nb_m in lots_avec_poids:
        t_min = min(tapis_cibles, key=lambda t: (charge_matchs_tapis[t], t))
        tapis_poules[t_min].extend(liste_p)
        charge_matchs_tapis[t_min] += nb_m

    # 5. Ordonnancement chronologique ronde par ronde
    for t in tapis_cibles:
        poules_t = tapis_poules[t]
        if not poules_t:
            continue
        
        matchs_tapis = []
        max_rondes = max((len(p.get('rondes', [])) for p in poules_t), default=0)
        for r in range(max_rondes):
            for p in poules_t:
                rondes_p = p.get('rondes', [])
                if r < len(rondes_p):
                    for m_idx, m in enumerate(rondes_p[r]):
                        c1 = m[0] if len(m) > 0 else {}
                        c2 = m[1] if len(m) > 1 else {}
                        label_tour = nommer_tour_match_global(p, r, m_idx, m)
                        matchs_tapis.append({
                            'poule': p.get('nom', ''),
                            'poule_obj': p,
                            'p1': c1.get('Nom', ''),
                            'p1_club': c1.get('Club', ''),
                            'p1_comite': c1.get('Comité', 'IDF'),
                            'p2': c2.get('Nom', ''),
                            'p2_club': c2.get('Club', ''),
                            'p2_comite': c2.get('Comité', 'IDF'),
                            'tour': r + 1,
                            'nom_tour': label_tour
                        })

        while matchs_tapis:
            t_curr_time = tapis_heure[t]
            match_choisi_idx = None
            for idx, m in enumerate(matchs_tapis):
                p1, p2 = m['p1'], m['p2']
                t_pret_p1 = last_match_time.get(p1, heure_debut_phase)
                t_pret_p2 = last_match_time.get(p2, heure_debut_phase)
                if t_pret_p1 <= t_curr_time and t_pret_p2 <= t_curr_time:
                    match_choisi_idx = idx
                    break
            
            if match_choisi_idx is not None:
                m = matchs_tapis.pop(match_choisi_idx)
                heure_debut_match = t_curr_time
            else:
                m = matchs_tapis.pop(0)
                p1, p2 = m['p1'], m['p2']
                t_pret = max(last_match_time.get(p1, heure_debut_phase), last_match_time.get(p2, heure_debut_phase), t_curr_time)
                heure_debut_match = t_pret

            duree_combat = obtenir_duree_combat_par_categorie_global(m['poule'], durees_config, defaut=3)
            nom_arb = choisir_arbitre_match_global(m['p1_club'], m['p2_club'], t, tapis_arbitres, liste_arbitres, ref_usage_count, eviter_meme_club)

            planning[t].append({
                "Type": "MATCH",
                "Heure": heure_debut_match.strftime("%H:%M"),
                "Duree": duree_combat,
                "Cat": m['poule'],
                "Combattant 1": m['p1'],
                "Club 1": m['p1_club'],
                "Comité 1": m['p1_comite'],
                "Combattant 2": m['p2'],
                "Club 2": m['p2_club'],
                "Comité 2": m['p2_comite'],
                "Tour": m['tour'],
                "Nom_Tour": m.get('nom_tour', f"Tour {m['tour']}"),
                "Arbitre": nom_arb
            })
            
            fin_match = heure_debut_match + timedelta(minutes=duree_combat)
            delai_repos = timedelta(minutes=(repos_matchs * duree_combat))
            last_match_time[m['p1']] = fin_match + delai_repos
            last_match_time[m['p2']] = fin_match + delai_repos
            tapis_heure[t] = fin_match

    return tapis_heure, planning

def analyser_formule_deroule(formule_str, cats_disponibles, heure_pesee_1=time(9, 0), duree_pesee_1=45,
                             heure_pesee_2=time(11, 30), duree_pesee_2=45, duree_pause=45):
    """
    Parse une formule textuelle comme 'PESEE 1 / U7 / U9 / PAUSE + PESEE 2 / U11 + U13'
    ou 'U7 / U9 + U11 / U13'.
    Retourne la liste des sessions séquentielles pour ordonnancer_chronogramme_complet.
    """
    f = str(formule_str or "").strip().upper()
    f = f.replace("PESÉE", "PESEE")
    f = re.sub(r'PESEE\s*1\b', 'PESEE_1', f)
    f = re.sub(r'PESEE\s*2\b', 'PESEE_2', f)
    
    segments = [s.strip() for s in f.split('/') if s.strip()]
    cats_upper_map = {c.upper(): c for c in cats_disponibles}
    sessions = []
    has_pesee1 = False
    cats_incluses = set()

    for seg in segments:
        parts = [p.strip() for p in seg.split('+') if p.strip()]
        seg_has_pause = any(p == "PAUSE" for p in parts)
        seg_has_pesee2 = any(p == "PESEE_2" for p in parts)
        seg_has_pesee1 = any(p in ["PESEE_1", "PESEE"] for p in parts)
        
        # 1. Pesée 1
        if seg_has_pesee1:
            has_pesee1 = True
            sessions.append({
                "type": "PESEE",
                "nom": "1ère Pesée",
                "heure_debut": heure_pesee_1,
                "duree": duree_pesee_1
            })
            
        # 2. Pause et/ou Pesée 2
        if seg_has_pause and seg_has_pesee2:
            sessions.append({
                "type": "PAUSE",
                "nom": "Pause & 2ème Pesée",
                "duree": duree_pause,
                "pesee2": True,
                "duree_pesee2": duree_pesee_2,
                "heure_pesee2": heure_pesee_2
            })
        elif seg_has_pause:
            sessions.append({
                "type": "PAUSE",
                "nom": "Pause de la compétition",
                "duree": duree_pause
            })
        elif seg_has_pesee2:
            sessions.append({
                "type": "PESEE_2",
                "nom": "2ème Pesée (Intermédiaire)",
                "duree": 0,
                "pesee2": True,
                "duree_pesee2": duree_pesee_2,
                "heure_pesee2": heure_pesee_2
            })
            
        # 3. Catégories de combat dans ce segment (simultanées avec '+')
        seg_cats = []
        for p in parts:
            if p in cats_upper_map:
                real_cat = cats_upper_map[p]
                if real_cat not in seg_cats:
                    seg_cats.append(real_cat)
                    cats_incluses.add(real_cat)
                    
        if seg_cats:
            label = " + ".join(seg_cats)
            sessions.append({
                "type": "COMBATS",
                "nom": f"Combats ({label})",
                "categories": seg_cats
            })

    if not has_pesee1:
        sessions.insert(0, {
            "type": "PESEE",
            "nom": "1ère Pesée",
            "heure_debut": heure_pesee_1,
            "duree": duree_pesee_1
        })

    manquantes = [c for c in cats_disponibles if c not in cats_incluses]
    if manquantes:
        sessions.append({
            "type": "COMBATS",
            "nom": f"Combats ({' + '.join(manquantes)})",
            "categories": manquantes
        })

    return sessions

def ordonnancer_chronogramme_complet(sessions_chrono, poules_dict, durees_config=None, nb_tapis=3, repos_matchs=3, meme_tapis=True, tapis_arbitres=None, liste_arbitres=None, ref_usage_count=None, eviter_meme_club=True):
    if durees_config is None:
        durees_config = {"U7": 10, "U9": 3, "U11": 4, "U13": 5}
    if ref_usage_count is None:
        ref_usage_count = {}

    planning_tapis = {t: [] for t in range(nb_tapis)}
    tapis_heure = [None for _ in range(nb_tapis)]
    lignes_accueil = []
    date_ref = datetime.today().date()
    heure_courante = None

    for step in sessions_chrono:
        s_type = step.get("type")
        
        if s_type == "PESEE":
            h_conf = step.get("heure_debut")
            duree = step.get("duree", 45)
            if h_conf is not None:
                if isinstance(h_conf, str):
                    try:
                        h_conf = datetime.strptime(h_conf, "%H:%M").time()
                    except Exception:
                        h_conf = time(9, 0)
                h_deb = datetime.combine(date_ref, h_conf)
            elif heure_courante is not None:
                h_deb = heure_courante
            else:
                h_deb = datetime.combine(date_ref, time(9, 0))
            
            h_fin = h_deb + timedelta(minutes=duree)
            heure_courante = h_fin
            for t in range(nb_tapis):
                if tapis_heure[t] is None or tapis_heure[t] < h_fin:
                    tapis_heure[t] = h_fin
            
            lignes_accueil.append({
                "Étape de la journée": step.get("nom", "Pesée"),
                "Horaire / Valeur": f"{h_deb.strftime('%H:%M')} (Échauffement jusqu'à {h_fin.strftime('%H:%M')})"
            })
            
        elif s_type == "COMBATS":
            cats = step.get("categories", [])
            valid_t = [t for t in tapis_heure if t is not None]
            debut_session = max(valid_t) if valid_t else datetime.combine(date_ref, time(9, 45))
            
            max_lignes = max((len(planning_tapis[t]) for t in range(nb_tapis)), default=0)
            for t in range(nb_tapis):
                while len(planning_tapis[t]) < max_lignes:
                    planning_tapis[t].append({"Type": "VIDE"})
            
            # Répartition dynamique des tapis pour les catégories simultanées (+)
            alloc_tapis = repartir_tapis_categories(cats, poules_dict, nb_tapis)
            
            for c in cats:
                tapis_c = alloc_tapis.get(c, list(range(nb_tapis)))
                if not tapis_c:
                    tapis_c = list(range(nb_tapis))

                if c == "U7" and poules_dict.get("U7"):
                    d_plat = durees_config.get("U7", 10)
                    noms_plateaux = [
                        "Plateau 1 : Motricité & Agilité",
                        "Plateau 2 : Ateliers techniques",
                        "Plateau 3 : Oppositions"
                    ]
                    rot_affectations = [
                        {0: 0, 1: 1, 2: 2},
                        {0: 2, 1: 0, 2: 1},
                        {0: 1, 1: 2, 2: 0}
                    ]
                    for r in range(3):
                        h_rot = debut_session + timedelta(minutes=r * d_plat)
                        for idx_sub, t in enumerate(tapis_c):
                            plat_idx = idx_sub % 3
                            grp_idx = rot_affectations[r][plat_idx]
                            grp_poule = poules_dict["U7"][grp_idx] if grp_idx < len(poules_dict["U7"]) else poules_dict["U7"][0]
                            nb_parts_u7 = len(grp_poule.get('participants', [])) if isinstance(grp_poule, dict) else 0
                            nom_arb = choisir_arbitre_match_global("", "", t, tapis_arbitres, liste_arbitres, ref_usage_count, eviter_meme_club)
                            arb_label = nom_arb if (nom_arb and nom_arb != "Non attribué") else "Animateur FFLDA"
                            planning_tapis[t].append({
                                "Type": "MATCH",
                                "Heure": h_rot.strftime("%H:%M"),
                                "Duree": d_plat,
                                "Cat": f"U7 | Plateau {plat_idx + 1}",
                                "Combattant 1": f"Groupe {grp_idx + 1} ({nb_parts_u7} lutteurs)",
                                "Club 1": "Tous clubs",
                                "Comité 1": "",
                                "Combattant 2": noms_plateaux[plat_idx],
                                "Club 2": "",
                                "Comité 2": "",
                                "Tour": f"Rotation {r + 1}/3",
                                "Nom_Tour": f"Rotation {r + 1} / 3",
                                "Arbitre": arb_label
                            })
                    fin_u7 = debut_session + timedelta(minutes=3 * d_plat)
                    for t in tapis_c:
                        tapis_heure[t] = fin_u7
                        
                elif c != "U7":
                    poules_c = [p for p in poules_dict.get(c, []) if isinstance(p, dict) and p.get('rondes')]
                    if poules_c:
                        t_h, pl_res = ordonnancer_phase_multi(
                            poules_phase=poules_c,
                            heure_debut_phase=debut_session,
                            durees_config=durees_config,
                            nb_tapis=nb_tapis,
                            repos_matchs=repos_matchs,
                            meme_tapis=meme_tapis,
                            tapis_arbitres=tapis_arbitres,
                            liste_arbitres=liste_arbitres,
                            ref_usage_count=ref_usage_count,
                            eviter_meme_club=eviter_meme_club,
                            tapis_cibles=tapis_c
                        )
                        for t in tapis_c:
                            planning_tapis[t].extend(pl_res[t])
                            tapis_heure[t] = t_h[t]

            fin_session = max((t for t in tapis_heure if t is not None), default=debut_session)
            if len(cats) > 1:
                desc_sim = []
                for c in cats:
                    t_list = alloc_tapis.get(c, [])
                    t_str = f"Tapis {', '.join(str(t+1) for t in t_list)}" if t_list else "Tous tapis"
                    desc_sim.append(f"{c} ({t_str})")
                titre_etape = f"Combats simultanés ({' + '.join(desc_sim)})"
            else:
                titre_etape = f"Combats ({cats[0]})"
                
            lignes_accueil.append({
                "Étape de la journée": titre_etape,
                "Horaire / Valeur": f"{debut_session.strftime('%H:%M')} - {fin_session.strftime('%H:%M')}"
            })
            for c in cats:
                nb_p = len(poules_dict.get(c, []))
                if nb_p > 0:
                    lignes_accueil.append({
                        "Étape de la journée": f"Poules {c}",
                        "Horaire / Valeur": f"{nb_p} poules"
                    })
            heure_courante = fin_session

        elif s_type in ["PAUSE", "PESEE_2"]:
            valid_t = [t for t in tapis_heure if t is not None]
            debut_pause = max(valid_t) if valid_t else datetime.combine(date_ref, time(12, 0))
            duree = step.get("duree", 45 if s_type == "PAUSE" else 0)
            fin_pause = debut_pause + timedelta(minutes=duree)
            
            if duree > 0:
                max_lignes = max((len(planning_tapis[t]) for t in range(nb_tapis)), default=0)
                for t in range(nb_tapis):
                    while len(planning_tapis[t]) < max_lignes:
                        planning_tapis[t].append({"Type": "VIDE"})
                    planning_tapis[t].append({"Type": "PAUSE", "Heure": debut_pause.strftime("%H:%M")})
                    tapis_heure[t] = fin_pause
                    
                lignes_accueil.append({
                    "Étape de la journée": step.get("nom", "Pause de la compétition"),
                    "Horaire / Valeur": f"{debut_pause.strftime('%H:%M')} - {fin_pause.strftime('%H:%M')} ({duree} min)"
                })
            
            if step.get("pesee2"):
                h_p2_conf = step.get("heure_pesee2")
                d_p2 = step.get("duree_pesee2", 45)
                if h_p2_conf is not None:
                    if isinstance(h_p2_conf, str):
                        try:
                            h_p2_conf = datetime.strptime(h_p2_conf, "%H:%M").time()
                        except Exception:
                            h_p2_conf = debut_pause.time()
                    dt_p2_deb = datetime.combine(date_ref, h_p2_conf)
                else:
                    dt_p2_deb = debut_pause

                fin_p2 = dt_p2_deb + timedelta(minutes=d_p2)
                lignes_accueil.append({
                    "Étape de la journée": "2ème pesée (intermédiaire)",
                    "Horaire / Valeur": f"{dt_p2_deb.strftime('%H:%M')} (Échauffement jusqu'à {fin_p2.strftime('%H:%M')})"
                })
                debut_aprem = max(fin_pause, fin_p2)
                for t in range(nb_tapis):
                    tapis_heure[t] = debut_aprem
                    planning_tapis[t].append({
                        "Type": "ATTENTE",
                        "Heure": dt_p2_deb.strftime("%H:%M"),
                        "Texte": f"⚖️ 2ème Pesée ({dt_p2_deb.strftime('%H:%M')} - {fin_p2.strftime('%H:%M')})"
                    })
                fin_pause = debut_aprem
            
            heure_courante = fin_pause

    valid_fin = [t for t in tapis_heure if t is not None]
    fin_estimee = max(valid_fin) if valid_fin else datetime.combine(date_ref, time(17, 0))
    total_poules = sum(len(p) for p in poules_dict.values())
    lignes_accueil.append({
        "Étape de la journée": "Total Groupes & Poules",
        "Horaire / Valeur": f"{total_poules} groupes"
    })
    lignes_accueil.append({
        "Étape de la journée": "Fin de la compétition estimée",
        "Horaire / Valeur": fin_estimee.strftime('%H:%M')
    })
    
    return planning_tapis, lignes_accueil, fin_estimee

def recalculer_chronogramme_tournoi_bundle(bundle, heure_pesee_1, duree_pesee_1=45, activer_pause=True, duree_pause=45, activer_pesee_2=True, duree_pesee_2=45, format_ordre="standard", custom_cats_matin=None, custom_cats_aprem=None, heure_pesee_2=None, formule_custom=None):
    if not bundle:
        return bundle
    
    nb_tapis = bundle.get("nb_tapis", 1)
    poules_dict = {
        "U7": bundle.get("poules_u7", []),
        "U9": bundle.get("poules_u9", []),
        "U11": bundle.get("poules_u11", []),
        "U13": bundle.get("poules_u13", [])
    }
    
    # Vérifier si les poules contiennent des objets complets avec rondes
    has_full_poules = any(
        any(isinstance(p, dict) and (p.get('rondes') or p.get('type_formule') == 'plateau_u7') for p in p_list)
        for p_list in poules_dict.values()
    )
    
    if has_full_poules:
        cats_disponibles = [k for k, v in poules_dict.items() if len(v) > 0]
        if not cats_disponibles:
            cats_disponibles = ["U7", "U9", "U11", "U13"]

        sessions = []
        fmt_low = str(format_ordre).lower()
        
        if formule_custom and str(formule_custom).strip():
            sessions = analyser_formule_deroule(
                formule_str=formule_custom,
                cats_disponibles=cats_disponibles,
                heure_pesee_1=heure_pesee_1,
                duree_pesee_1=duree_pesee_1,
                heure_pesee_2=heure_pesee_2 or time(11, 30),
                duree_pesee_2=duree_pesee_2,
                duree_pause=duree_pause if activer_pause else 0
            )
        elif "simultané" in fmt_low or "simultane" in fmt_low:
            sessions.append({"type": "PESEE", "nom": "1ère Pesée", "heure_debut": heure_pesee_1, "duree": duree_pesee_1})
            sessions.append({"type": "COMBATS", "nom": "Session Matin (U7 + U9)", "categories": ["U7", "U9"]})
            if (activer_pause and duree_pause > 0) or activer_pesee_2:
                sessions.append({"type": "PAUSE", "nom": "Pause & 2ème Pesée" if activer_pesee_2 else "Pause de la compétition", "duree": duree_pause if activer_pause else 0, "pesee2": activer_pesee_2, "duree_pesee2": duree_pesee_2, "heure_pesee2": heure_pesee_2})
            sessions.append({"type": "COMBATS", "nom": "Session Après-midi (U11 + U13)", "categories": ["U11", "U13"]})
            
        elif "enchaîné" in fmt_low or "enchaine" in fmt_low or "1 pesée" in fmt_low:
            sessions.append({"type": "PESEE", "nom": "Pesée Générale Unique", "heure_debut": heure_pesee_1, "duree": duree_pesee_1})
            sessions.append({"type": "COMBATS", "nom": "Combats U7", "categories": ["U7"]})
            sessions.append({"type": "COMBATS", "nom": "Combats U9", "categories": ["U9"]})
            sessions.append({"type": "COMBATS", "nom": "Combats U11", "categories": ["U11"]})
            sessions.append({"type": "COMBATS", "nom": "Combats U13", "categories": ["U13"]})
            
        elif "grands" in fmt_low or "inversé" in fmt_low or "inverse" in fmt_low:
            sessions.append({"type": "PESEE", "nom": "1ère Pesée (Grands : U11 / U13)", "heure_debut": heure_pesee_1, "duree": duree_pesee_1})
            sessions.append({"type": "COMBATS", "nom": "Session Matin (U11 + U13)", "categories": ["U11", "U13"]})
            if (activer_pause and duree_pause > 0) or activer_pesee_2:
                sessions.append({"type": "PAUSE", "nom": "Pause & 2ème Pesée (Petits : U7 / U9)" if activer_pesee_2 else "Pause de la compétition", "duree": duree_pause if activer_pause else 0, "pesee2": activer_pesee_2, "duree_pesee2": duree_pesee_2, "heure_pesee2": heure_pesee_2})
            sessions.append({"type": "COMBATS", "nom": "Session Après-midi (U7 + U9)", "categories": ["U7", "U9"]})
            
        elif "personnalisé" in fmt_low or "personnalise" in fmt_low:
            matin_cats = custom_cats_matin or ["U7", "U9"]
            aprem_cats = custom_cats_aprem or ["U11", "U13"]
            sessions.append({"type": "PESEE", "nom": "1ère Pesée", "heure_debut": heure_pesee_1, "duree": duree_pesee_1})
            sessions.append({"type": "COMBATS", "nom": f"Session Matin ({', '.join(matin_cats)})", "categories": matin_cats})
            if (activer_pause and duree_pause > 0) or activer_pesee_2:
                sessions.append({"type": "PAUSE", "nom": "Pause & 2ème Pesée" if activer_pesee_2 else "Pause", "duree": duree_pause if activer_pause else 0, "pesee2": activer_pesee_2, "duree_pesee2": duree_pesee_2, "heure_pesee2": heure_pesee_2})
            sessions.append({"type": "COMBATS", "nom": f"Session Après-midi ({', '.join(aprem_cats)})", "categories": aprem_cats})
            
        else: # Standard FFLDA : U7 -> U9 -> Pause -> U11 -> U13
            sessions.append({"type": "PESEE", "nom": "1ère Pesée (U7 / U9)", "heure_debut": heure_pesee_1, "duree": duree_pesee_1})
            sessions.append({"type": "COMBATS", "nom": "Combats U7", "categories": ["U7"]})
            sessions.append({"type": "COMBATS", "nom": "Combats U9", "categories": ["U9"]})
            if (activer_pause and duree_pause > 0) or activer_pesee_2:
                sessions.append({"type": "PAUSE", "nom": "Pause & 2ème Pesée (U11 / U13)" if activer_pesee_2 else "Pause de la compétition", "duree": duree_pause if activer_pause else 0, "pesee2": activer_pesee_2, "duree_pesee2": duree_pesee_2, "heure_pesee2": heure_pesee_2})
            sessions.append({"type": "COMBATS", "nom": "Combats U11", "categories": ["U11"]})
            sessions.append({"type": "COMBATS", "nom": "Combats U13", "categories": ["U13"]})

        tapis_arbitres = bundle.get("tapis_arbitres", {})
        liste_arbitres = bundle.get("liste_arbitres", [])
        
        nouveau_planning, nouvelles_lignes, fin_estimee = ordonnancer_chronogramme_complet(
            sessions_chrono=sessions,
            poules_dict=poules_dict,
            durees_config={"U7": 10, "U9": 3, "U11": 4, "U13": 5},
            nb_tapis=nb_tapis,
            tapis_arbitres=tapis_arbitres,
            liste_arbitres=liste_arbitres
        )
        
        bundle["planning_tapis"] = nouveau_planning
        bundle["lignes_accueil"] = nouvelles_lignes
        bundle["grille_ui"] = construire_grille_ui_depuis_planning(nouveau_planning, nb_tapis)
        
        # Synchronisation complète de la liste des matchs (Tablette / Scoreboard TV / Direct)
        anc_matchs = bundle.get("matchs_direct", [])
        anciens_scores_map = {}
        for m in anc_matchs:
            k_m = (int(m.get("tapis", 1)), str(m.get("categorie", "")), str(m.get("lutteur_rouge", "")), str(m.get("lutteur_bleu", "")))
            anciens_scores_map[k_m] = m
            anciens_scores_map[m.get("id")] = m

        m_prefix_code = "ORG"
        try:
            if 'st' in globals() and hasattr(st, 'session_state'):
                m_prefix_code = st.session_state.get('code_session', 'ORG')
        except Exception:
            pass
        nom_comp = bundle.get("nom_competition", "tournoi")
        m_nom_tourn = re.sub(r'[^a-zA-Z0-9]', '_', nom_comp)[:15]

        nouveaux_matchs_direct = []
        for t in range(nb_tapis):
            m_count_t = 0
            for item in nouveau_planning.get(t, []):
                if item.get("Type") == "MATCH":
                    m_count_t += 1
                    cat_item = item.get("Cat", "")
                    c1_m = item.get("Combattant 1", "")
                    c2_m = item.get("Combattant 2", "")
                    k_m = (t + 1, cat_item, c1_m, c2_m)
                    anc_data = anciens_scores_map.get(k_m, {})
                    
                    m_dict = {
                        "id": f"{m_prefix_code}_{m_nom_tourn}_T{t + 1}_M{m_count_t:03d}",
                        "tournoi_id": nom_comp,
                        "code_organisateur": m_prefix_code,
                        "tapis": t + 1,
                        "match_num": m_count_t,
                        "heure": item.get("Heure", ""),
                        "duree": item.get("Duree", 2),
                        "categorie": cat_item,
                        "tour": nettoyer_nom_tour(item.get("Nom_Tour") or item.get("Tour") or ""),
                        "lutteur_rouge": c1_m,
                        "club_rouge": item.get("Club 1", ""),
                        "comite_rouge": item.get("Comité 1", ""),
                        "lutteur_bleu": c2_m,
                        "club_bleu": item.get("Club 2", ""),
                        "comite_bleu": item.get("Comité 2", ""),
                        "sheet": f"Grille Tapis {t + 1}",
                        "statut": anc_data.get("statut", "À venir"),
                        "vainqueur": anc_data.get("vainqueur", ""),
                        "type_victoire": anc_data.get("type_victoire", ""),
                        "score_rouge": anc_data.get("score_rouge", 0),
                        "score_bleu": anc_data.get("score_bleu", 0),
                        "pt_clt_rouge": anc_data.get("pt_clt_rouge", 0),
                        "pt_clt_bleu": anc_data.get("pt_clt_bleu", 0)
                    }
                    nouveaux_matchs_direct.append(m_dict)
                    
        bundle["matchs_direct"] = nouveaux_matchs_direct

        # Mise à jour du classeur Excel et du PDF officiel si existants
        if bundle.get("excel_bytes"):
            try:
                import openpyxl
                from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
                wb = openpyxl.load_workbook(io.BytesIO(bundle["excel_bytes"]), data_only=False)
                
                # 1. Mise à jour de l'onglet Résumé
                if "Résumé" in wb.sheetnames:
                    ws_res = wb["Résumé"]
                    for r in range(2, ws_res.max_row + 1):
                        for c in range(1, 3):
                            ws_res.cell(row=r, column=c, value="")
                    for idx_lig, lig in enumerate(bundle.get("lignes_accueil", []), start=2):
                        ws_res.cell(row=idx_lig, column=1, value=lig.get("Étape de la journée", ""))
                        ws_res.cell(row=idx_lig, column=2, value=lig.get("Horaire / Valeur", ""))
                        ws_res.cell(row=idx_lig, column=1).alignment = Alignment(horizontal="center", vertical="center")
                        ws_res.cell(row=idx_lig, column=2).alignment = Alignment(horizontal="center", vertical="center")
                
                # 2. Mise à jour de l'onglet Grille de Passage
                if "Grille de Passage" in wb.sheetnames:
                    ws_grille = wb["Grille de Passage"]
                    b_style = Border(left=Side(style='thin', color='B0BEC5'), right=Side(style='thin', color='B0BEC5'),
                                     top=Side(style='thin', color='B0BEC5'), bottom=Side(style='thin', color='B0BEC5'))
                    rouge_fill = PatternFill("solid", fgColor="E53935")
                    for r in range(3, ws_grille.max_row + 1):
                        for c in range(1, nb_tapis + 1):
                            cell = ws_grille.cell(row=r, column=c)
                            cell.value = ""
                            cell.fill = PatternFill(fill_type=None)
                            
                    for row_idx, r_dict in enumerate(bundle.get("grille_ui", []), start=3):
                        ws_grille.row_dimensions[row_idx].height = 90
                        for col_t in range(1, nb_tapis + 1):
                            cell = ws_grille.cell(row=row_idx, column=col_t)
                            val_t = r_dict.get(f"Tapis {col_t}", "")
                            cell.value = val_t
                            cell.border = b_style
                            cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
                            if val_t:
                                if "PAUSE" in val_t:
                                    cell.fill = rouge_fill
                                    cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=12)
                                elif any(k in val_t for k in ["Attente", "Pesée", "échauffement", "Repos"]):
                                    cell.fill = PatternFill("solid", fgColor="EFEFEF")
                                    cell.font = Font(name="Arial", italic=True, color="666666", size=11)
                                else:
                                    age_k = extraire_age_de_texte(val_t)
                                    cfg_c = COULEURS_AGE_GRILLE.get(age_k, COULEURS_AGE_GRILLE['AUTRE'])
                                    cell.fill = cfg_c['bg_excel_pastel']
                                    cell.font = Font(name="Arial", size=10, color="0F172A")
                                    
                out_mem = io.BytesIO()
                wb.save(out_mem)
                nouv_bytes = out_mem.getvalue()
                bundle["excel_bytes"] = nouv_bytes
                
                try:
                    pdf_b = generer_pdf_depuis_classeur_excel(wb, nom_comp, wb=wb)
                    if pdf_b:
                        bundle["pdf_bytes"] = pdf_b
                except Exception:
                    pass
            except Exception:
                pass
        
    else:
        # Repli : Décalage en cascade complet sur les matchs existants
        try:
            date_ref = datetime.today().date()
            h_ref = heure_pesee_1 if isinstance(heure_pesee_1, time) else datetime.strptime(str(heure_pesee_1)[:5], "%H:%M").time()
            nouveau_debut_matin = datetime.combine(date_ref, h_ref) + timedelta(minutes=duree_pesee_1)
            
            pl = bundle.get("planning_tapis", {})
            nb_tapis = bundle.get("nb_tapis", 1)
            
            def est_match_matin(m_item):
                cat = str(m_item.get("Cat", "")).upper()
                return "U7" in cat or "U9" in cat
            
            heures_matin = [it.get("Heure") for t in pl for it in pl[t] if it.get("Type") == "MATCH" and est_match_matin(it) and it.get("Heure")]
            delta_matin = timedelta(0)
            if heures_matin:
                h_min_m = min(datetime.strptime(h, "%H:%M") for h in heures_matin)
                dt_min_m = datetime.combine(date_ref, h_min_m.time())
                delta_matin = nouveau_debut_matin - dt_min_m
            
            fin_matin = nouveau_debut_matin
            for t in pl:
                for it in pl[t]:
                    if it.get("Type") == "MATCH" and est_match_matin(it) and it.get("Heure"):
                        try:
                            h_dt = datetime.strptime(it["Heure"], "%H:%M")
                            it_debut = datetime.combine(date_ref, h_dt.time()) + delta_matin
                            it["Heure"] = it_debut.strftime("%H:%M")
                            it_fin = it_debut + timedelta(minutes=int(it.get("Duree", 3)))
                            if it_fin > fin_matin:
                                fin_matin = it_fin
                        except Exception:
                            pass

            fin_pause = fin_matin + timedelta(minutes=duree_pause) if (activer_pause and duree_pause > 0) else fin_matin

            if activer_pesee_2:
                if heure_pesee_2 is not None:
                    h2_ref = heure_pesee_2 if isinstance(heure_pesee_2, time) else datetime.strptime(str(heure_pesee_2)[:5], "%H:%M").time()
                else:
                    h2_ref = fin_matin.time()
                dt_p2_deb = datetime.combine(date_ref, h2_ref)
                dt_p2_fin = dt_p2_deb + timedelta(minutes=duree_pesee_2)
            else:
                dt_p2_deb = None
                dt_p2_fin = None

            debut_aprem_voulu = max(fin_pause, dt_p2_fin) if dt_p2_fin else fin_pause

            heures_aprem = [it.get("Heure") for t in pl for it in pl[t] if it.get("Type") == "MATCH" and not est_match_matin(it) and it.get("Heure")]
            fin_aprem = debut_aprem_voulu
            delta_aprem = timedelta(0)
            if heures_aprem:
                h_min_ap = min(datetime.strptime(h, "%H:%M") for h in heures_aprem)
                dt_min_ap = datetime.combine(date_ref, h_min_ap.time())
                delta_aprem = debut_aprem_voulu - dt_min_ap
                
                for t in pl:
                    for it in pl[t]:
                        if it.get("Type") == "MATCH" and not est_match_matin(it) and it.get("Heure"):
                            try:
                                h_dt = datetime.strptime(it["Heure"], "%H:%M")
                                it_debut = datetime.combine(date_ref, h_dt.time()) + delta_aprem
                                it["Heure"] = it_debut.strftime("%H:%M")
                                it_fin = it_debut + timedelta(minutes=int(it.get("Duree", 4)))
                                if it_fin > fin_aprem:
                                    fin_aprem = it_fin
                            except Exception:
                                pass
                        elif it.get("Type") == "PAUSE":
                            it["Heure"] = fin_matin.strftime("%H:%M")

            for m in bundle.get("matchs_direct", []):
                cat_m = str(m.get("categorie", "")).upper()
                delta_use = delta_matin if ("U7" in cat_m or "U9" in cat_m) else (delta_aprem if heures_aprem else delta_matin)
                if m.get("heure"):
                    try:
                        h_dt = datetime.strptime(m["heure"], "%H:%M")
                        m["heure"] = (datetime.combine(date_ref, h_dt.time()) + delta_use).strftime("%H:%M")
                    except Exception:
                        pass

            nouvelles_lignes = [
                {"Étape de la journée": "1ère Pesée", "Horaire / Valeur": f"{h_ref.strftime('%H:%M')} (Échauffement jusqu'à {nouveau_debut_matin.strftime('%H:%M')})"},
                {"Étape de la journée": "Combats Matin (U7 / U9)", "Horaire / Valeur": f"{nouveau_debut_matin.strftime('%H:%M')} - {fin_matin.strftime('%H:%M')}"}
            ]
            if activer_pause and duree_pause > 0:
                nouvelles_lignes.append({"Étape de la journée": "Pause de la compétition", "Horaire / Valeur": f"{fin_matin.strftime('%H:%M')} - {fin_pause.strftime('%H:%M')} ({duree_pause} min)"})
            if dt_p2_deb and dt_p2_fin:
                nouvelles_lignes.append({"Étape de la journée": "2ème pesée (intermédiaire)", "Horaire / Valeur": f"{dt_p2_deb.strftime('%H:%M')} (Échauffement jusqu'à {dt_p2_fin.strftime('%H:%M')})"})
            if heures_aprem:
                nouvelles_lignes.append({"Étape de la journée": "Combats Après-midi (U11 / U13)", "Horaire / Valeur": f"{debut_aprem_voulu.strftime('%H:%M')} - {fin_aprem.strftime('%H:%M')}"})
            nouvelles_lignes.append({"Étape de la journée": "Fin de la compétition estimée", "Horaire / Valeur": fin_aprem.strftime('%H:%M')})
            
            bundle["lignes_accueil"] = nouvelles_lignes
            bundle["grille_ui"] = construire_grille_ui_depuis_planning(pl, nb_tapis)
        except Exception:
            pass

    return bundle

def afficher_onglets_tournoi(bundle):
    if not bundle:
        return
    lignes_accueil = bundle.get("lignes_accueil", [])
    total_participants_peses = bundle.get("total_participants_peses", 0)
    total_non_peses = bundle.get("total_non_peses", 0)
    poules_u7 = bundle.get("poules_u7", [])
    poules_u9 = bundle.get("poules_u9", [])
    poules_u11 = bundle.get("poules_u11", [])
    poules_u13 = bundle.get("poules_u13", [])
    total_matchs_calcules = bundle.get("total_matchs_calcules", 0)
    liste_arbitres = bundle.get("liste_arbitres", [])
    tapis_arbitres = bundle.get("tapis_arbitres", {})
    nb_tapis = bundle.get("nb_tapis", 1)
    grille_ui = bundle.get("grille_ui", [])
    planning_tapis = bundle.get("planning_tapis", {})
    participants_par_poule = bundle.get("participants_par_poule", {})
    poule_obj_map = bundle.get("poule_obj_map", {})

    # Fonction de tri naturel par ordre de poule (Poule 1, Poule 2... ou poids croissant)
    def cle_tri_nom_poule(nom_p):
        s = str(nom_p or "").strip()
        # 1. Numéro explicite : Poule 1, Poule 2, Groupe 1, Plateau 1...
        m_num = re.search(r'\b(?:poule|groupe|plateau)\s*([0-9]+)\b', s, re.IGNORECASE)
        if m_num:
            return (0, int(m_num.group(1)), s)
        # 2. Lettre de poule : Poule A, Poule B...
        m_lettre = re.search(r'\b(?:poule|groupe)\s*([a-zA-Z])\b', s, re.IGNORECASE)
        if m_lettre:
            return (1, ord(m_lettre.group(1).upper()), s)
        # 3. Premier poids numérique trouvé (ex: 20 / 22 kg -> 20.0)
        m_poids = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:kg|/|-)', s, re.IGNORECASE)
        if not m_poids:
            m_poids = re.search(r'(\d+(?:[.,]\d+)?)\b', s)
        if m_poids:
            try:
                return (2, float(m_poids.group(1).replace(',', '.')), s)
            except Exception:
                pass
        # 4. Repli alphabétique
        return (3, 0, s)

    # Regroupement des poules par catégorie d'âge pour une navigation fluide et aérée
    ordre_ages_ref = ['U7', 'U9', 'U11', 'U13', 'U15', 'U17', 'U20', 'SENIOR']
    poules_par_age = {}
    for nom_poule, liste_p in participants_par_poule.items():
        age_c = extraire_age_de_texte(nom_poule)
        poules_par_age.setdefault(age_c, []).append((nom_poule, liste_p))

    ages_presents = sorted(
        poules_par_age.keys(),
        key=lambda a: ordre_ages_ref.index(a) if a in ordre_ages_ref else 99
    )

    def _label_onglet_age(a):
        if a == 'U7':
            return "Groupes U7"
        elif a in ['U9', 'U11', 'U13']:
            return f"Poules {a}"
        else:
            return f"Inscrits {a}"

    onglets_ages_labels = [_label_onglet_age(a) for a in ages_presents]

    # --- ONGLETS INTERACTIFS DE L'APPLICATION ---
    noms_onglets = (
        ["Résumé & Stats", "Grille Globale", "Équipes d'Arbitrage"] +
        [f"Grille Tapis {t + 1}" for t in range(nb_tapis)] +
        onglets_ages_labels
    )
    onglets_ui = st.tabs(noms_onglets)
    
    with onglets_ui[0]:
        if st.session_state.get("chrono_recalc_notif"):
            st.success(st.session_state.pop("chrono_recalc_notif"))
        st.subheader("Résumé prévisionnel de la journée")

        # Recherche des valeurs existantes de Pesée et Pause
        val_pesee_exist = ""
        val_pause_exist = ""
        val_pesee2_exist = ""
        for lig in (lignes_accueil or []):
            et = str(lig.get("Étape de la journée", "")).lower()
            val = str(lig.get("Horaire / Valeur", ""))
            if "2ème pesée" in et or "deuxième pesée" in et:
                val_pesee2_exist = val
            elif "pesée" in et and not val_pesee_exist:
                val_pesee_exist = val
            elif "pause" in et and not val_pause_exist:
                val_pause_exist = val

        # Valeurs par défaut si non renseignées
        if not val_pesee_exist:
            val_pesee_exist = "08:30"
        if not val_pause_exist:
            val_pause_exist = "12:00 - 12:45"

        with st.expander("⏱️ ⚖️ Ajuster le Déroulé & l'Ordre de la Compétition (Cascade Automatique)", expanded=st.session_state.get("expander_regles_open", False)):
            st.caption("Modifiez l'horaire de pesée, les durées, les pauses ou l'ordre des catégories. L'ensemble des tapis, des combats et des horaires est automatiquement recalculé et synchronisé en cascade !")
            
            # Extraction heure existante pour valeur par défaut
            h_init = time(9, 0)
            if val_pesee_exist:
                m_h = re.search(r'(\d{1,2})[:h](\d{2})', val_pesee_exist)
                if m_h:
                    try:
                        h_init = time(int(m_h.group(1)), int(m_h.group(2)))
                    except Exception:
                        pass

            h2_init = time(11, 30)
            if val_pesee2_exist:
                m_h2 = re.search(r'(\d{1,2})[:h](\d{2})', val_pesee2_exist)
                if m_h2:
                    try:
                        h2_init = time(int(m_h2.group(1)), int(m_h2.group(2)))
                    except Exception:
                        pass
            
            # Détermination préalable du format et de la formule pour conditionner l'état activé/grisé des options
            cats_dispos = []
            for c_k in ["U7", "U9", "U11", "U13"]:
                if bundle.get(f"poules_{c_k.lower()}"):
                    cats_dispos.append(c_k)
            if not cats_dispos:
                cats_dispos = ["U7", "U9", "U11", "U13"]

            # Formule par défaut selon l'état actuel
            options_ordre = [
                "Standard FFLDA (U7 ➔ U9 ➔ Pause ➔ U11 ➔ U13)",
                "Simultané par demi-journée (U7 + U9 Matin ➔ Pause ➔ U11 + U13 Aprem)",
                "Tout enchaîné / 1 pesée générale (U7 ➔ U9 ➔ U11 ➔ U13 sans pause)",
                "Grands le matin (U11 + U13 Matin ➔ Pause ➔ U7 + U9 Aprem)",
                "✍️ Personnalisé (Formule libre avec +, /)"
            ]
            format_ordre_courant = st.session_state.get("downstream_format_ordre", options_ordre[0])
            if "Personnalisé" in str(format_ordre_courant) or "formule" in str(format_ordre_courant).lower():
                def_formule = f"PESEE 1 / {' / '.join(cats_dispos[:2])} / PAUSE + PESEE 2 / {' + '.join(cats_dispos[2:])}" if len(cats_dispos) > 2 else f"PESEE 1 / {' / '.join(cats_dispos)}"
                formule_active = st.session_state.get("downstream_formule_custom", def_formule)
            elif "simultané" in str(format_ordre_courant).lower():
                formule_active = "PESEE 1 / U7 + U9 / PAUSE + PESEE 2 / U11 + U13"
            elif "enchaîné" in str(format_ordre_courant).lower() or "1 pesée" in str(format_ordre_courant).lower():
                formule_active = f"PESEE 1 / {' / '.join(cats_dispos)}"
            elif "grands" in str(format_ordre_courant).lower():
                formule_active = "PESEE 1 / U11 + U13 / PAUSE + PESEE 2 / U7 + U9"
            else:
                formule_active = "PESEE 1 / U7 / U9 / PAUSE + PESEE 2 / U11 + U13"

            f_upper_eval = str(formule_active or "").upper().replace("PESÉE", "PESEE")
            has_pesee2_in_formula = ("PESEE 2" in f_upper_eval or "PESEE_2" in f_upper_eval or "PESEE2" in f_upper_eval)
            has_pause_in_formula = ("PAUSE" in f_upper_eval)
            
            col_p1, col_p2, col_p3 = st.columns([1, 1, 1])
            with col_p1:
                st.markdown("##### ⚖️ Pesées")
                nouveau_h_pesee = st.time_input("Heure 1ère pesée", value=h_init, key="downstream_h_pesee")
                duree_pesee_val = st.selectbox("Durée 1ère pesée + échauffement", [30, 45, 60, 75, 90], index=1, format_func=lambda x: f"{x} min", key="downstream_duree_pesee")
                st.write("")
                pesee2_coche = bool(val_pesee2_exist) if has_pesee2_in_formula else False
                nouveau_pesee2_act = st.checkbox(
                    "Activer une 2ème pesée (intermédiaire)",
                    value=pesee2_coche if has_pesee2_in_formula else False,
                    disabled=(not has_pesee2_in_formula),
                    key="downstream_pesee2_act",
                    help="Grisé si PESEE 2 n'apparaît pas dans le déroulé de compétition choisi ou dans la formule."
                )
                if nouveau_pesee2_act and has_pesee2_in_formula:
                    nouveau_h_pesee2 = st.time_input("Heure 2ème pesée", value=h2_init, key="downstream_h_pesee2")
                    duree_pesee2_val = st.selectbox("Durée 2ème pesée", [30, 45, 60], index=1, format_func=lambda x: f"{x} min", key="downstream_pesee2_dur")
                else:
                    nouveau_h_pesee2 = None
                    duree_pesee2_val = 45
            
            with col_p2:
                st.markdown("##### ⏸️ Pause de la Compétition")
                pause_coche = bool(val_pause_exist and "0" not in val_pause_exist and val_pause_exist != "-") if has_pause_in_formula else False
                nouveau_pause_act = st.checkbox(
                    "Activer une pause déjeuner",
                    value=pause_coche if has_pause_in_formula else False,
                    disabled=(not has_pause_in_formula),
                    key="downstream_pause_act",
                    help="Grisé si PAUSE n'apparaît pas dans le déroulé de compétition choisi ou dans la formule."
                )
                duree_pause_val = st.selectbox(
                    "Durée de la pause",
                    [15, 30, 45, 60, 75],
                    index=2,
                    format_func=lambda x: f"{x} min",
                    disabled=(not has_pause_in_formula),
                    key="downstream_pause_dur"
                )

            with col_p3:
                st.markdown("##### 🥋 Déroulé & Ordre des Catégories")
                format_ordre_sel = st.selectbox("Format d'enchaînement", options_ordre, index=0, key="downstream_format_ordre")

                # Callbacks pour la composition par boutons cliquables
                def _cb_ajouter_item_formule(item):
                    st.session_state["expander_regles_open"] = True
                    anc_format = str(st.session_state.get("downstream_format_ordre", ""))
                    etait_personnalise = ("Personnalisé" in anc_format)
                    st.session_state["downstream_format_ordre"] = "✍️ Personnalisé (Formule libre avec +, /)"
                    
                    if not etait_personnalise:
                        cur = ""
                    else:
                        cur = str(st.session_state.get("downstream_formule_custom", "") or "").strip()
                        
                    if item in ["+", "/"]:
                        if cur.endswith("+") or cur.endswith("/"):
                            cur = cur[:-1].strip()
                        nouv = f"{cur} {item}" if cur else item
                    else:
                        if not cur:
                            nouv = item
                        elif cur.endswith("+") or cur.endswith("/"):
                            nouv = f"{cur} {item}"
                        else:
                            nouv = f"{cur} / {item}"
                    st.session_state["downstream_formule_custom"] = nouv.strip()

                def _cb_effacer_dernier_item():
                    st.session_state["expander_regles_open"] = True
                    st.session_state["downstream_format_ordre"] = "✍️ Personnalisé (Formule libre avec +, /)"
                    cur = str(st.session_state.get("downstream_formule_custom", "") or "").strip()
                    if not cur:
                        return
                    tokens = cur.split()
                    if tokens:
                        tokens.pop()
                    st.session_state["downstream_formule_custom"] = " ".join(tokens)

                def _cb_reinitialiser_formule():
                    st.session_state["expander_regles_open"] = True
                    st.session_state["downstream_format_ordre"] = "✍️ Personnalisé (Formule libre avec +, /)"
                    st.session_state["downstream_formule_custom"] = ""

                st.caption("🖱️ **Touches rapides d'enchaînement :**")
                
                # Rangée 1 : Boutons des catégories (U7, U9, U11, U13...)
                cats_boutons = ["U7", "U9", "U11", "U13"]
                for c in cats_dispos:
                    if c not in cats_boutons:
                        cats_boutons.append(c)
                
                cols_cats = st.columns(len(cats_boutons))
                for idx_c, cat_btn in enumerate(cats_boutons):
                    with cols_cats[idx_c]:
                        st.button(cat_btn, key=f"btn_item_{cat_btn}", on_click=_cb_ajouter_item_formule, args=(cat_btn,), use_container_width=True)
                
                # Rangée 2 : Boutons d'opérateurs (+, /) et de contrôle (Pause, ⌫)
                c_op1, c_op2, c_op3, c_op4 = st.columns(4)
                with c_op1:
                    st.button("➕ (+)", key="btn_item_plus", on_click=_cb_ajouter_item_formule, args=("+",), use_container_width=True, help="Simultané (+) : les catégories luttent en même temps sur des tapis distincts")
                with c_op2:
                    st.button("➗ (/)", key="btn_item_slash", on_click=_cb_ajouter_item_formule, args=("/",), use_container_width=True, help="Enchaîné (/) : attendre la fin d'une catégorie avant de lancer la suivante")
                with c_op3:
                    st.button("⏸️ Pause", key="btn_item_pause", on_click=_cb_ajouter_item_formule, args=("PAUSE",), use_container_width=True, help="Insérer une pause déjeuner dans le déroulé")
                with c_op4:
                    st.button("⌫ Effacer", key="btn_item_del", on_click=_cb_effacer_dernier_item, use_container_width=True, help="Effacer le dernier item")
                
                cats_detectees_str = " / ".join(cats_dispos)
                formule_custom_val = None
                if "Personnalisé" in format_ordre_sel:
                    val_init_formule = st.session_state.get("downstream_formule_custom", f"PESEE 1 / {' / '.join(cats_dispos[:2])} / PAUSE + PESEE 2 / {' + '.join(cats_dispos[2:])}" if len(cats_dispos) > 2 else f"PESEE 1 / {' / '.join(cats_dispos)}")
                    if "downstream_formule_custom" not in st.session_state:
                        st.session_state["downstream_formule_custom"] = val_init_formule
                        
                    formule_custom_val = st.text_input(
                        "Formule du déroulé (+ = simultané, / = enchaîné) :",
                        key="downstream_formule_custom",
                        help=f"Exemples : 'U7/U9+U11/U13' ou 'PESEE 1 / U7 / U9 / PAUSE + PESEE 2 / U11 + U13'. Catégories détectées : {cats_detectees_str}"
                    )
                    
                    c_rst1, c_rst2 = st.columns([1, 1])
                    with c_rst1:
                        st.button("⚖️ 2ème Pesée", key="btn_item_pesee2", on_click=_cb_ajouter_item_formule, args=("PESEE 2",), use_container_width=True, help="Insérer une 2ème pesée intermédiaire")
                    with c_rst2:
                        st.button("🔄 Réinitialiser", key="btn_item_rst", on_click=_cb_reinitialiser_formule, use_container_width=True, help="Effacer toute la formule pour recommencer")
                        
                    st.caption(f"💡 **Déroulé actif pris en compte :** `{formule_custom_val or '(vide)'}`", unsafe_allow_html=True)

            # Détection automatique de changement de paramètres pour recalcul réactif immédiat
            c_sess_courant = st.session_state.get("code_session", "ORG")
            nom_comp_active = bundle.get("nom_competition", "tournoi")
            sig_key = f"chrono_last_applied_sig_{c_sess_courant}_{nom_comp_active}"

            cur_sig = (
                str(nouveau_h_pesee),
                int(duree_pesee_val),
                bool(nouveau_pause_act and has_pause_in_formula),
                int(duree_pause_val),
                bool(nouveau_pesee2_act and has_pesee2_in_formula),
                int(duree_pesee2_val),
                str(nouveau_h_pesee2) if (nouveau_pesee2_act and has_pesee2_in_formula) else "NONE",
                str(format_ordre_sel),
                str(formule_custom_val or "")
            )

            param_a_change = False
            if sig_key not in st.session_state:
                st.session_state[sig_key] = cur_sig
            elif st.session_state[sig_key] != cur_sig:
                param_a_change = True

            st.write("")
            col_b1, col_b2 = st.columns([1, 2])
            with col_b1:
                btn_recalc_downstream = st.button("⚡ Recalculer le Déroulé en Cascade", type="primary", use_container_width=True, key="btn_recalc_downstream")

            if btn_recalc_downstream or param_a_change:
                st.session_state[sig_key] = cur_sig
                bundle = recalculer_chronogramme_tournoi_bundle(
                    bundle=bundle,
                    heure_pesee_1=nouveau_h_pesee,
                    duree_pesee_1=duree_pesee_val,
                    activer_pause=nouveau_pause_act and has_pause_in_formula,
                    duree_pause=duree_pause_val,
                    activer_pesee_2=nouveau_pesee2_act and has_pesee2_in_formula,
                    duree_pesee_2=duree_pesee2_val,
                    heure_pesee_2=nouveau_h_pesee2 if (nouveau_pesee2_act and has_pesee2_in_formula) else None,
                    format_ordre=format_ordre_sel,
                    formule_custom=formule_custom_val if "Personnalisé" in format_ordre_sel else None
                )
                st.session_state["tournoi_actif_cache"] = bundle
                st.session_state["matchs_direct"] = bundle.get("matchs_direct", [])
                if bundle.get("excel_bytes"):
                    st.session_state["excel_tournoi_base"] = bundle["excel_bytes"]
                if bundle.get("pdf_bytes"):
                    st.session_state["pdf_tournoi_base"] = bundle["pdf_bytes"]
                _SHARED_TOURNAMENT_MATCHS[c_sess_courant] = bundle.get("matchs_direct", [])
                _SHARED_TOURNAMENT_MATCHS["DEFAULT"] = bundle.get("matchs_direct", [])
                sauvegarder_cache_tournoi(c_sess_courant, bundle)
                try:
                    safe_c = re.sub(r'[^a-zA-Z0-9_-]', '_', str(c_sess_courant))
                    os.makedirs(".cache_tournois", exist_ok=True)
                    with open(f".cache_tournois/resume_{safe_c}.json", "w", encoding="utf-8") as f_res:
                        json.dump(bundle.get("lignes_accueil", []), f_res, ensure_ascii=False, indent=2)
                except Exception:
                    pass
                st.session_state["chrono_recalc_notif"] = "⚡ Programme et chronogramme recalculés et synchronisés en cascade !"
                st.rerun()

        if lignes_accueil:
            st.table(pd.DataFrame(lignes_accueil))
        
        col_m1, col_m2, col_m3, col_m4, col_m5, col_m6, col_m7, col_m8 = st.columns(8)
        col_m1.metric("Participants (pesés)", total_participants_peses)
        col_m2.metric("Absents / Non pesés", total_non_peses)
        col_m3.metric("Groupes U7 (Plateaux)", len(poules_u7))
        col_m4.metric("Poules U9", len(poules_u9))
        col_m5.metric("Poules U11", len(poules_u11))
        col_m6.metric("Groupes U13", len(poules_u13))
        col_m7.metric("Matchs générés", total_matchs_calcules)
        col_m8.metric("Arbitres engagés", len(liste_arbitres))

        if liste_arbitres:
            st.markdown("#### Désignation des Équipes d'Arbitrage par Tapis")
            lignes_arb_sum = []
            for t in range(nb_tapis):
                noms_arb = ", ".join([a['Nom_Complet'] for a in tapis_arbitres.get(t, [])]) if tapis_arbitres.get(t) else "Aucun arbitre affecté"
                lignes_arb_sum.append({"Tapis": f"Tapis {t + 1}", "Effectif": f"{len(tapis_arbitres.get(t, []))} arbitres", "Équipe d'Arbitrage Désignée": noms_arb})
            st.table(pd.DataFrame(lignes_arb_sum))

    with onglets_ui[1]:
        st.subheader("📅 Grille Globale de Passage - Tous les Tapis")
        if liste_arbitres:
            st.markdown("**Équipes d'arbitrage affectées aux tapis :**")
            cols_arb_disp = st.columns(nb_tapis)
            for t in range(nb_tapis):
                with cols_arb_disp[t]:
                    arb_list_txt = "\n".join([f"• **{a['Nom_Complet']}** ({a['Club']})" for a in tapis_arbitres.get(t, [])]) if tapis_arbitres.get(t) else "• Aucun"
                    st.info(f"**Tapis {t+1}** ({len(tapis_arbitres.get(t, []))} arbitres) :\n\n{arb_list_txt}")
        
        if grille_ui:
            html_grille_ui = generer_html_grille_coloree(grille_ui, nb_tapis)
            st.markdown(html_grille_ui, unsafe_allow_html=True)
            st.write("")
        else:
            st.info("Grille globale générée dans le classeur Excel officiel.")

    with onglets_ui[2]:
        st.subheader("Désignation et Affectation des Arbitres par Tapis")
        if liste_arbitres:
            for t in range(nb_tapis):
                st.markdown(f"#### 🥋 Tapis {t + 1} ({len(tapis_arbitres.get(t, []))} arbitres)")
                if tapis_arbitres.get(t):
                    df_arb_tapis = pd.DataFrame(tapis_arbitres[t])[['Nom', 'Prenom', 'Licence', 'Club', 'Comite']]
                    df_arb_tapis.columns = ['Nom', 'Prénom', 'N° Licence', 'Club', 'Comité Régional']
                    st.table(df_arb_tapis)
                else:
                    st.info("Aucun arbitre affecté à ce tapis.")
        else:
            st.info("Aucun fichier d'arbitres n'a été chargé.")

    # Onglets Grille Tapis individuel
    for t in range(nb_tapis):
        with onglets_ui[3 + t]:
            st.subheader(f"🥋 Grille de Passage & Feuille de Marque — Tapis {t + 1}")
            if liste_arbitres and tapis_arbitres.get(t):
                arb_names_st = ", ".join([f"**{a['Nom_Complet']}** ({a['Club']})" for a in tapis_arbitres[t]])
                st.info(f"**Équipe d'arbitrage désignée (Tapis {t + 1})** : {arb_names_st}")
            else:
                st.caption("Aucun arbitre désigné spécifiquement sur ce tapis.")
            
            m_count_st = 0
            for m in planning_tapis.get(t, []):
                if m["Type"] == "PAUSE":
                    st.warning(f"[{m['Heure']}] ⏸️ PAUSE DE LA COMPÉTITION")
                elif m["Type"] == "ATTENTE":
                    st.info(f"[{m['Heure']}] {m['Texte']}")
                elif m["Type"] == "MATCH":
                    m_count_st += 1
                    age_m = extraire_age_de_texte(m['Cat'])
                    badge_age = COULEURS_AGE_GRILLE.get(age_m, {}).get('badge', '')
                    badge_str = f"{badge_age} " if badge_age else ""
                    arb_info_st = f" | Arbitre : {m['Arbitre']}" if m.get('Arbitre') and m['Arbitre'] != "Non attribué" else ""
                    t_nom = nettoyer_nom_tour(m.get('Nom_Tour') or m.get('Tour'))
                    tour_st = f" | 🎯 Tour : **{t_nom}**" if t_nom else ""
                    st.markdown(f"#### 🤼 MATCH N° {m_count_st} — 🕘 {m['Heure']} ({m['Duree']} min) | Catégorie : {badge_str}`{m['Cat']}`{tour_st}{arb_info_st}")
                    
                    c1_cl = f" ({m.get('Club 1', '')})" if m.get('Club 1') else ""
                    c1_co = f" - {m.get('Comité 1', '')}" if (m.get('Comité 1') and m.get('Comité 1') != 'Comité Non Renseigné') else ""
                    c2_cl = f" ({m.get('Club 2', '')})" if m.get('Club 2') else ""
                    c2_co = f" - {m.get('Comité 2', '')}" if (m.get('Comité 2') and m.get('Comité 2') != 'Comité Non Renseigné') else ""
                    
                    df_m_ui = pd.DataFrame([
                        {
                            "N°": m_count_st,
                            "LUTTEUR ROUGE": f"🔴 {m['Combattant 1']}{c1_cl}{c1_co}",
                            "Pt Clt (Rouge)": "[   ]",
                            "Type (Rouge)": "[   ]",
                            "Points Techniques (Actions Rouge)": "[                                 ]",
                            "Total Score (Rouge)": "[   ]",
                            "VS": "VS",
                            "LUTTEUR BLEU": f"🔵 {m['Combattant 2']}{c2_cl}{c2_co}",
                            "Pt Clt (Bleu)": "[   ]",
                            "Type (Bleu)": "[   ]",
                            "Points Techniques (Actions Bleu)": "[                                 ]",
                            "Total Score (Bleu)": "[   ]"
                        }
                    ])
                    st.table(df_m_ui)

    start_idx_ages = 3 + nb_tapis
    for idx_age, age_cle in enumerate(ages_presents):
        with onglets_ui[start_idx_ages + idx_age]:
            liste_poules_age = sorted(
                poules_par_age.get(age_cle, []),
                key=lambda x: cle_tri_nom_poule(x[0])
            )
            nb_tot_lutteurs_age = sum(len(lp) for _, lp in liste_poules_age)
            titre_sec = "🏅 Animation U7 (Plateaux Pédagogiques)" if age_cle == 'U7' else f"🤼 Poules et Inscrits {age_cle}"
            st.markdown(f"### {titre_sec}")
            st.caption(f"Total : **{len(liste_poules_age)}** poule(s) · **{nb_tot_lutteurs_age}** lutteur(s) engagé(s)")
            
            # Barre de recherche rapide pour la catégorie d'âge
            c_rech, c_vue = st.columns([3, 1])
            with c_rech:
                rech_txt = st.text_input(
                    f"🔍 Rechercher un lutteur, un club ou un poids ({age_cle}) :",
                    key=f"rech_poule_{age_cle}",
                    placeholder="Tapez un nom, prénom, club..."
                ).strip().lower()
            with c_vue:
                tout_deplier = st.checkbox("Tout déplier", key=f"deplier_{age_cle}", value=bool(rech_txt))
            
            poules_affichees = 0
            for nom_poule, liste_p in liste_poules_age:
                if rech_txt:
                    match_poule = rech_txt in nom_poule.lower()
                    match_lutteur = any(
                        rech_txt in str(p.get("Nom", "")).lower() or
                        rech_txt in str(p.get("Club", "")).lower() or
                        rech_txt in str(p.get("Poids", "")).lower()
                        for p in liste_p
                    )
                    if not (match_poule or match_lutteur):
                        continue
                
                poules_affichees += 1
                p_obj = poule_obj_map.get(nom_poule) if poule_obj_map else None
                
                exp_label = f"🤼 {nom_poule}  ({len(liste_p)} lutteurs)"
                with st.expander(exp_label, expanded=(tout_deplier or bool(rech_txt))):
                    if p_obj and p_obj.get('type_formule') == 'plateau_u7':
                        st.info("Formule officielle FFLDA : Découverte pédagogique sous forme de 3 plateaux d'activités avec rotation (Motricité, Opposition au sol, Lutte debout). Tous les enfants sont récompensés !")
                        df_u7_tab = pd.DataFrame([
                            {
                                "N°": idx_p,
                                "Nom Prénom": p.get('Nom', ''),
                                "Club": p.get('Club', ''),
                                "Poids": formater_poids(p.get('Poids', '')),
                                "Plateau 1 : Motricité & Agilité": "✓ Validé",
                                "Plateau 2 : Ateliers techniques": "✓ Validé",
                                "Plateau 3 : Oppositions": "✓ Validé",
                                "Récompense": "🥇 Médaille d'Or"
                            }
                            for idx_p, p in enumerate(liste_p, 1)
                        ])
                        st.table(df_u7_tab)
                    else:
                        df_poule_vue = pd.DataFrame(liste_p)[['Nom', 'Club', 'Poids']]
                        st.table(df_poule_vue)

                        if p_obj and p_obj.get('type_formule') == 'poules_croisees':
                            st.markdown("##### 🥋 Répartition en 2 Poules de 3 :")
                            c_pa, c_pb = st.columns(2)
                            with c_pa:
                                st.markdown("**Poule A**")
                                st.table(pd.DataFrame(p_obj['poule_a'])[['Nom', 'Club', 'Poids']])
                            with c_pb:
                                st.markdown("**Poule B**")
                                st.table(pd.DataFrame(p_obj['poule_b'])[['Nom', 'Club', 'Poids']])
            
            if poules_affichees == 0 and rech_txt:
                st.warning(f"Aucune poule ou lutteur ne correspond à la recherche « {rech_txt} » en {age_cle}.")


# Détection proactive d'un tournoi déjà actif dans Supabase
_, _, is_supa_detect = get_supabase_config()
c_sess_detect = st.session_state.get("code_session", "")
tournoi_actif_detecte = None
nb_combats_detectes = 0
nb_termines_detectes = 0
matchs_detectes = []

if is_supa_detect and c_sess_detect:
    params_chk = {"order": "tapis.asc,match_num.asc"}
    if c_sess_detect not in ["FFLDA-ADMIN", "FFLDA2026"]:
        params_chk["code_organisateur"] = f"eq.{c_sess_detect}"
    res_chk, _ = supabase_request("matchs_lutte", params=params_chk)
    if res_chk and isinstance(res_chk, list) and len(res_chk) > 0:
        matchs_detectes = res_chk
        tournoi_actif_detecte = res_chk[0].get("tournoi_id", "Tournoi en cours")
        nb_combats_detectes = len(res_chk)
        nb_termines_detectes = sum(1 for m in res_chk if m.get("statut") == "Terminé")

if not matchs_detectes and st.session_state.get("matchs_direct"):
    m_dir = st.session_state["matchs_direct"]
    if len(m_dir) > 0:
        matchs_detectes = m_dir
        tournoi_actif_detecte = st.session_state.get("nom_competition_active") or m_dir[0].get("tournoi_id", "Tournoi en cours")
        nb_combats_detectes = len(m_dir)
        nb_termines_detectes = sum(1 for m in m_dir if m.get("statut") == "Terminé")

# Restaurer le bundle depuis le cache disque si nécessaire
if c_sess_detect and not st.session_state.get("tournoi_actif_cache"):
    cached_bundle = charger_cache_tournoi(c_sess_detect)
    if cached_bundle:
        if tournoi_actif_detecte:
            if cached_bundle.get("nom_competition") == tournoi_actif_detecte:
                st.session_state["tournoi_actif_cache"] = cached_bundle
                if not st.session_state.get("excel_tournoi_base") and cached_bundle.get("excel_bytes"):
                    st.session_state["excel_tournoi_base"] = cached_bundle["excel_bytes"]
                if not st.session_state.get("pdf_tournoi_base") and cached_bundle.get("pdf_bytes"):
                    st.session_state["pdf_tournoi_base"] = cached_bundle["pdf_bytes"]
                if not st.session_state.get("matchs_direct") and cached_bundle.get("matchs_direct"):
                    st.session_state["matchs_direct"] = cached_bundle["matchs_direct"]
                if not st.session_state.get("nom_competition_active"):
                    st.session_state["nom_competition_active"] = tournoi_actif_detecte
        elif not is_supa_detect:
            st.session_state["tournoi_actif_cache"] = cached_bundle
            tournoi_actif_detecte = cached_bundle.get("nom_competition", "Tournoi Local")
            nb_combats_detectes = cached_bundle.get("total_matchs_calcules", 0)
            if not st.session_state.get("excel_tournoi_base") and cached_bundle.get("excel_bytes"):
                st.session_state["excel_tournoi_base"] = cached_bundle["excel_bytes"]
            if not st.session_state.get("pdf_tournoi_base") and cached_bundle.get("pdf_bytes"):
                st.session_state["pdf_tournoi_base"] = cached_bundle["pdf_bytes"]
            if not st.session_state.get("matchs_direct") and cached_bundle.get("matchs_direct"):
                st.session_state["matchs_direct"] = cached_bundle["matchs_direct"]
            if not st.session_state.get("nom_competition_active"):
                st.session_state["nom_competition_active"] = tournoi_actif_detecte

# Si l'utilisateur relance ou arrive sur l'application, on démarre TOUJOURS sur le Mode 1 (Générer / Résumé prévisionnel)
if "mode_app_index" not in st.session_state:
    st.session_state["mode_app_index"] = 0

# Sélection du mode de travail : 3 boutons sur la même ligne
if not is_kiosque:
    cur_mode_idx = st.session_state.get("mode_app_index", 0)
    cur_saisie = st.session_state.get("mode_saisie_scores", "ligne")

    st.markdown("""
    <style>
    div[data-testid="stVerticalBlock"]:has(#nav_modes_marker) {
        margin-bottom: 8px;
    }
    div[data-testid="stVerticalBlock"]:has(#nav_modes_marker) div[data-testid="stHorizontalBlock"] {
        background: #f1f5f9 !important;
        padding: 6px !important;
        border-radius: 14px !important;
        border: 1.5px solid #cbd5e1 !important;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.04) !important;
    }
    div[data-testid="stVerticalBlock"]:has(#nav_modes_marker) button,
    button[key^="nav_btn_mode_"],
    button[key^="btn_tab_saisie_"] {
        min-height: 48px !important;
        font-size: 14px !important;
        font-weight: 800 !important;
        border-radius: 10px !important;
    }
    div[data-testid="stVerticalBlock"]:has(#nav_modes_marker) button[kind="secondary"],
    div[data-testid="stVerticalBlock"]:has(#nav_modes_marker) button[data-testid="baseButton-secondary"],
    button[key^="nav_btn_mode_"][kind="secondary"],
    button[key^="btn_tab_saisie_"][kind="secondary"] {
        background-color: #ffffff !important;
        background: #ffffff !important;
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
        border: 1.5px solid #cbd5e1 !important;
    }
    div[data-testid="stVerticalBlock"]:has(#nav_modes_marker) button[kind="secondary"] *,
    div[data-testid="stVerticalBlock"]:has(#nav_modes_marker) button[data-testid="baseButton-secondary"] *,
    button[key^="nav_btn_mode_"][kind="secondary"] *,
    button[key^="btn_tab_saisie_"][kind="secondary"] * {
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
    }
    div[data-testid="stVerticalBlock"]:has(#nav_modes_marker) button[kind="secondary"]:hover,
    div[data-testid="stVerticalBlock"]:has(#nav_modes_marker) button[data-testid="baseButton-secondary"]:hover,
    button[key^="nav_btn_mode_"][kind="secondary"]:hover,
    button[key^="btn_tab_saisie_"][kind="secondary"]:hover {
        background-color: #f1f5f9 !important;
        background: #f1f5f9 !important;
        color: #0055A4 !important;
        -webkit-text-fill-color: #0055A4 !important;
        border-color: #0055A4 !important;
    }
    div[data-testid="stVerticalBlock"]:has(#nav_modes_marker) button[kind="secondary"]:hover *,
    div[data-testid="stVerticalBlock"]:has(#nav_modes_marker) button[data-testid="baseButton-secondary"]:hover *,
    button[key^="nav_btn_mode_"][kind="secondary"]:hover *,
    button[key^="btn_tab_saisie_"][kind="secondary"]:hover * {
        color: #0055A4 !important;
        -webkit-text-fill-color: #0055A4 !important;
    }
    div[data-testid="stVerticalBlock"]:has(#nav_modes_marker) button[kind="primary"],
    div[data-testid="stVerticalBlock"]:has(#nav_modes_marker) button[data-testid="baseButton-primary"],
    button[key^="nav_btn_mode_"][kind="primary"],
    button[key^="btn_tab_saisie_"][kind="primary"] {
        background: linear-gradient(135deg, #0055A4 0%, #1d4ed8 100%) !important;
        background-color: #0055A4 !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        border: none !important;
    }
    div[data-testid="stVerticalBlock"]:has(#nav_modes_marker) button[kind="primary"] *,
    button[key^="nav_btn_mode_"][kind="primary"] *,
    button[key^="btn_tab_saisie_"][kind="primary"] * {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div id="nav_modes_marker"></div>', unsafe_allow_html=True)
    col_nav1, col_nav2, col_nav3 = st.columns([1, 1, 1], gap="small")

    with col_nav1:
        if st.button(
            "1. Générer Tournoi (Planning & Poules)", 
            type="primary" if cur_mode_idx == 0 else "secondary", 
            use_container_width=True, 
            key="nav_btn_mode_1"
        ):
            st.session_state["mode_app_index"] = 0
            st.rerun()

    with col_nav2:
        if st.button(
            "2. Rentrer les Scores (Table de Marque)", 
            type="primary" if cur_mode_idx == 1 else "secondary", 
            use_container_width=True, 
            key="nav_btn_mode_2"
        ):
            st.session_state["mode_app_index"] = 1
            st.rerun()

    with col_nav3:
        if st.button(
            "3. Résultats & Bilans Fédéraux", 
            type="primary" if cur_mode_idx == 2 else "secondary", 
            use_container_width=True, 
            key="nav_btn_mode_3"
        ):
            st.session_state["mode_app_index"] = 2
            st.rerun()

    st.markdown("<div style='margin-bottom: 14px;'></div>", unsafe_allow_html=True)

# Détermination du mode actif
cur_mode_idx = st.session_state.get("mode_app_index", 0)
if is_kiosque:
    mode_app = "2. Saisie Tapis en Direct (Table de Marque)"
elif cur_mode_idx == 1:
    mode_app = "2. Saisie Tapis en Direct (Table de Marque)"
elif cur_mode_idx == 2:
    mode_app = "3. Calcul des Résultats & Bilan Fédéral"
else:
    mode_app = "1. Générer un Tournoi (Planning & Poules)"

if mode_app.startswith("2"):
    if not is_kiosque:
        st.markdown("## Saisie des Scores — Table de Marque")
        st.caption("Sélectionnez votre mode de saisie : manuellement sur le classeur Excel ou en ligne sur smartphones/tablettes.")

        col_sub_m1, col_sub_m2 = st.columns([1, 1])
        with col_sub_m1:
            if st.button("EXCEL", type="primary" if st.session_state.get("mode_saisie_scores") == "excel" else "secondary", use_container_width=True, key="btn_tab_saisie_excel"):
                st.session_state["mode_saisie_scores"] = "excel"
                st.session_state["source_scores_choix"] = "Fichier Excel complété (.xlsx)"
                st.rerun()
        with col_sub_m2:
            if st.button("EN LIGNE", type="primary" if st.session_state.get("mode_saisie_scores") != "excel" else "secondary", use_container_width=True, key="btn_tab_saisie_ligne"):
                st.session_state["mode_saisie_scores"] = "ligne"
                st.session_state["source_scores_choix"] = "Scores saisis en direct (Tablettes / Tapis)"
                st.rerun()
        st.markdown("---")

    if not is_kiosque and st.session_state.get("mode_saisie_scores") == "excel":
        st.markdown("### Option A : Remplir les scores sur le Classeur Excel Officiel")
        st.info("Vous avez choisi de gérer la saisie des combats directement dans votre fichier Excel officiel FFLDA.")

        excel_base = st.session_state.get("excel_tournoi_base")
        nom_comp = st.session_state.get("nom_competition_active") or nom_competition

        col_ex1, col_ex2 = st.columns([1, 1])
        with col_ex1:
            st.markdown("#### 1. Fichiers Officiels du Tournoi")
            if excel_base:
                st.download_button(
                    label="Télécharger le Classeur Officiel Excel (.xlsx)",
                    data=excel_base,
                    file_name=f"Tournoi_{re.sub(r'[^a-zA-Z0-9]', '_', nom_comp)}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_down_excel_mode_saisie",
                    use_container_width=True,
                    type="primary"
                )
                st.caption("Classeur complet généré lors de la création du tournoi, prêt avec toutes les grilles de tapis.")
            if st.session_state.get("pdf_tournoi_base"):
                st.download_button(
                    label="Télécharger le Dossier Officiel en PDF",
                    data=st.session_state["pdf_tournoi_base"],
                    file_name=f"Dossier_Officiel_{re.sub(r'[^a-zA-Z0-9]', '_', nom_comp)}.pdf",
                    mime="application/pdf",
                    key="btn_down_pdf_mode_saisie",
                    use_container_width=True
                )
            if not excel_base and not st.session_state.get("pdf_tournoi_base"):
                st.warning("⚠️ Aucun classeur officiel en mémoire pour cette session. Téléversez le fichier initial ci-dessous :")
                f_init = st.file_uploader("Classeur initial du tournoi (.xlsx)", type=["xlsx"], key="up_init_saisie_excel")
                if f_init:
                    st.session_state["excel_tournoi_base"] = f_init.read()
                    st.rerun()

        with col_ex2:
            st.markdown("#### 2. Instructions de Saisie")
            st.markdown(
                """
                - Ouvrez le fichier dans **Microsoft Excel** ou **LibreOffice**.
                - Rendez-vous sur les onglets **Grille Tapis 1**, **Grille Tapis 2**, etc.
                - Pour chaque combat :
                  - Indiquez les points des lutteurs (cellules score).
                  - Indiquez le type de victoire : **VT** (Tombé), **VST** (Grande supériorité), **VP** (Points).
                - Enregistrez votre fichier Excel une fois les combats terminés.
                """
            )

        st.stop()
    
    # 1. Vérification de la configuration Supabase
    _, _, is_supa = get_supabase_config()

    # 2. Récupération des combats du tournoi
    c_sess = st.session_state.get("code_session", "")
    params_q = {"order": "tapis.asc,match_num.asc"}
    if c_sess and c_sess not in ["FFLDA-ADMIN", "FFLDA2026"]:
        params_q["code_organisateur"] = f"eq.{c_sess}"

    matchs_direct = []
    if is_supa:
        res_supa, err_s = supabase_request("matchs_lutte", params=params_q)
        if res_supa and isinstance(res_supa, list) and len(res_supa) > 0:
            tournois_ids = list(dict.fromkeys([m.get("tournoi_id", "") for m in res_supa if m.get("tournoi_id")]))
            if len(tournois_ids) > 1:
                t_nom_choisi = st.selectbox("🏆 Plusieurs tournois détectés. Sélectionnez le tournoi :", tournois_ids, index=len(tournois_ids)-1)
                matchs_direct = [m for m in res_supa if m.get("tournoi_id") == t_nom_choisi]
            else:
                matchs_direct = res_supa
        elif err_s:
            st.error(f"⚠️ Erreur de connexion Supabase : {err_s}")
            matchs_direct = st.session_state.get("matchs_direct", [])
        else:
            matchs_direct = st.session_state.get("matchs_direct", [])
    else:
        matchs_direct = st.session_state.get("matchs_direct", [])

    if not matchs_direct:
        if c_sess and c_sess in _SHARED_TOURNAMENT_MATCHS:
            matchs_direct = _SHARED_TOURNAMENT_MATCHS[c_sess]
            st.session_state["matchs_direct"] = matchs_direct
        elif "DEFAULT" in _SHARED_TOURNAMENT_MATCHS:
            matchs_direct = _SHARED_TOURNAMENT_MATCHS["DEFAULT"]
            st.session_state["matchs_direct"] = matchs_direct
        elif c_sess:
            cb = charger_cache_tournoi(c_sess)
            if cb and cb.get("matchs_direct"):
                matchs_direct = cb["matchs_direct"]
                st.session_state["matchs_direct"] = matchs_direct
                _SHARED_TOURNAMENT_MATCHS[c_sess] = matchs_direct
        
        # Secours fichier disque pour nouvelles sessions / tablettes / téléphones scannant le QR code
        if not matchs_direct and os.path.exists(".cache_tournois/matchs_direct_backup.json"):
            try:
                with open(".cache_tournois/matchs_direct_backup.json", "r", encoding="utf-8") as f_bk:
                    matchs_direct = json.load(f_bk)
                    st.session_state["matchs_direct"] = matchs_direct
                    _SHARED_TOURNAMENT_MATCHS["DEFAULT"] = matchs_direct
                    if c_sess:
                        _SHARED_TOURNAMENT_MATCHS[c_sess] = matchs_direct
            except Exception:
                pass

    if matchs_direct:
        _SHARED_TOURNAMENT_MATCHS["DEFAULT"] = matchs_direct
        if c_sess:
            _SHARED_TOURNAMENT_MATCHS[c_sess] = matchs_direct
        try:
            os.makedirs(".cache_tournois", exist_ok=True)
            with open(".cache_tournois/matchs_direct_backup.json", "w", encoding="utf-8") as f_bk:
                json.dump(matchs_direct, f_bk)
        except Exception:
            pass

        for m_chk in matchs_direct:
            m_id_chk = m_chk.get("id")
            if m_id_chk in _SHARED_MATCHS_DATA:
                m_chk.update(_SHARED_MATCHS_DATA[m_id_chk])

    if not matchs_direct:
        st.warning("⚠️ **Aucun combat n'est actuellement initialisé pour la saisie en direct.**")
        st.info("👉 Allez d'abord dans le **Mode 1 (Générer un Tournoi)** et importez votre fichier d'inscriptions. Vos grilles de combats seront automatiquement créées et prêtes à être saisies ici !")
    else:
        # QR Codes et liens directs pour les arbitres
        tapis_dispos = sorted(list(set(int(m.get("tapis", 1)) for m in matchs_direct)))
        if not tapis_dispos:
            tapis_dispos = [1]

        if not is_kiosque:
            with st.expander("📲 Partager l'accès aux Tables de Marque (QR Codes & Liens Directs)", expanded=True):
                url_defaut_detectee = st.session_state.get("url_app_base", "")
                if not url_defaut_detectee:
                    h = ""
                    try:
                        if hasattr(st, "context") and hasattr(st.context, "headers"):
                            h = st.context.headers.get("host") or ""
                    except Exception:
                        pass
                    if h:
                        if "localhost" in h or "127.0.0.1" in h:
                            port = h.split(":")[-1] if ":" in h else "8501"
                            ip_lan = detecter_ip_reseau_local()
                            if ip_lan and ip_lan != "localhost":
                                url_defaut_detectee = f"http://{ip_lan}:{port}"
                            else:
                                url_defaut_detectee = f"http://{h}"
                        elif h.startswith("192.168.") or h.startswith("10.") or h.startswith("172."):
                            url_defaut_detectee = f"http://{h}"
                        else:
                            url_defaut_detectee = f"https://{h}"
                    else:
                        ip_lan = detecter_ip_reseau_local()
                        if ip_lan and ip_lan != "localhost":
                            url_defaut_detectee = f"http://{ip_lan}:8501"
                        else:
                            url_defaut_detectee = "http://localhost:8501"

                url_prefix = url_defaut_detectee if url_defaut_detectee else "https://votre-app.streamlit.app"
                
                # Onglets pour basculer entre QR Codes Arbitres et QR Codes Scoreboards TV
                tab_qr_tables, tab_qr_scb = st.tabs([
                    "📱 QR Codes Tables de Marque (Arbitres & Saisie)", 
                    "📺 QR Codes Scoreboards (TV & Écran Public Tapis)"
                ])

                supa_u_c, supa_k_c, _ = get_supabase_config()
                has_secret_key = hasattr(st, "secrets") and bool(st.secrets.get("SUPABASE_KEY"))
                sk_param = {}
                if supa_k_c and not has_secret_key and not KEY_SUPABASE_DEFAUT:
                    sk_param["sk"] = supa_k_c

                # --- ONGLET 1 : QR CODES TABLES DE MARQUE (ARBITRAGE SÉCURISÉ) ---
                with tab_qr_tables:
                    st.caption("Ces QR codes permettent aux arbitres et aides de table de saisir les scores avec mot de passe automatique.")
                    nb_cartes_qr = len(tapis_dispos) + 1
                    cols_qr = st.columns(min(nb_cartes_qr, 4))
                    
                    for i_tap, t_id in enumerate(tapis_dispos):
                        with cols_qr[i_tap % 4]:
                            params_lien = {"code": c_sess, "mode": "direct", "tapis": t_id, "cle": get_cle_secrete_table(c_sess), **sk_param}
                            lien_direct = f"{url_prefix}/?{urllib.parse.urlencode(params_lien)}"
                            qr_url = f"https://api.qrserver.com/v1/create-qr-code/?size=220x220&data={urllib.parse.quote(lien_direct)}"
                            st.markdown(f"<div style='text-align:center;'><b>TAPIS {t_id}</b><br><img src='{qr_url}' width='140'></div>", unsafe_allow_html=True)
                            st.link_button(f"👉 Ouvrir Tapis {t_id}", lien_direct, use_container_width=True)
                            st.caption(f"`{lien_direct}`")
                    
                    # 4ème Carte : QR Code Déroulé Général (3 Tapis)
                    with cols_qr[len(tapis_dispos) % 4]:
                        params_lien_all = {"code": c_sess, "mode": "direct", "tapis": "all", **sk_param}
                        lien_all = f"{url_prefix}/?{urllib.parse.urlencode(params_lien_all)}"
                        qr_url_all = f"https://api.qrserver.com/v1/create-qr-code/?size=220x220&data={urllib.parse.quote(lien_all)}"
                        st.markdown(f"<div style='text-align:center;'><b style='color:#B71C1C;'>📋 DÉROULÉ (3 TAPIS)</b><br><img src='{qr_url_all}' width='140'></div>", unsafe_allow_html=True)
                        st.link_button("👉 Ouvrir Déroulé", lien_all, use_container_width=True)
                        st.caption(f"`{lien_all}`")

                    st.markdown("<br>", unsafe_allow_html=True)
                    nom_tourn_pdf = st.session_state.get("nom_competition_active") or "Tournoi_FFLDA"
                    try:
                        pdf_fiche_qr = generer_pdf_fiche_qr_codes(nom_tourn_pdf, c_sess, url_prefix, tapis_dispos, type_fiche="table")
                        st.download_button(
                            label="📄 Télécharger la Fiche A4 des QR Codes Tables de Marque (Arbitres)",
                            data=pdf_fiche_qr,
                            file_name=f"Fiche_QR_Codes_Tables_{c_sess}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                            key="btn_dl_fiche_qr_a4"
                        )
                    except Exception as e_pdf_qr:
                        st.caption(f"Note génération PDF QR : {e_pdf_qr}")

                # --- ONGLET 2 : QR CODES SCOREBOARDS GRAND ÉCRAN (TV / TABLETTES TAPIS) ---
                with tab_qr_scb:
                    st.caption("Scannez ces QR codes avec une Smart TV, tablette ou vidéoprojecteur au bord de chaque tapis pour afficher le Scoreboard officiel en direct.")
                    cols_scb = st.columns(min(nb_cartes_qr, 4))
                    for i_tap, t_id in enumerate(tapis_dispos):
                        with cols_scb[i_tap % 4]:
                            params_lien_scb = {"code": c_sess, "mode": "direct", "tapis": t_id, "vue": "scoreboard", **sk_param}
                            lien_scb = f"{url_prefix}/?{urllib.parse.urlencode(params_lien_scb)}"
                            qr_url_scb = f"https://api.qrserver.com/v1/create-qr-code/?size=220x220&data={urllib.parse.quote(lien_scb)}"
                            st.markdown(f"<div style='text-align:center;'><b style='color:#1e3a8a;'>📺 SCOREBOARD TAPIS {t_id}</b><br><img src='{qr_url_scb}' width='140'></div>", unsafe_allow_html=True)
                            st.link_button(f"📺 Ouvrir Scoreboard {t_id}", lien_scb, use_container_width=True)
                            st.caption(f"`{lien_scb}`")

                    # 4ème Carte : QR Code Déroulé Général (3 Tapis)
                    with cols_scb[len(tapis_dispos) % 4]:
                        params_lien_all = {"code": c_sess, "mode": "direct", "tapis": "all", **sk_param}
                        lien_all = f"{url_prefix}/?{urllib.parse.urlencode(params_lien_all)}"
                        qr_url_all = f"https://api.qrserver.com/v1/create-qr-code/?size=220x220&data={urllib.parse.quote(lien_all)}"
                        st.markdown(f"<div style='text-align:center;'><b style='color:#B71C1C;'>📋 DÉROULÉ (3 TAPIS)</b><br><img src='{qr_url_all}' width='140'></div>", unsafe_allow_html=True)
                        st.link_button("👉 Ouvrir Déroulé", lien_all, use_container_width=True)
                        st.caption(f"`{lien_all}`")

                    st.markdown("<br>", unsafe_allow_html=True)
                    try:
                        pdf_scb_qr = generer_pdf_fiche_qr_codes(nom_tourn_pdf, c_sess, url_prefix, tapis_dispos, type_fiche="scoreboard")
                        st.download_button(
                            label="📺 Télécharger la Fiche A4 des QR Codes Scoreboards TV (Grand Écran)",
                            data=pdf_scb_qr,
                            file_name=f"Fiche_QR_Codes_Scoreboards_TV_{c_sess}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                            key="btn_dl_fiche_qr_scb_a4"
                        )
                    except Exception as e_pdf_scb:
                        st.caption(f"Note génération PDF Scoreboard QR : {e_pdf_scb}")
        
        options_tapis = [f"Tapis {t}" for t in tapis_dispos] + ["📋 Déroulé Général (Tous Tapis)", "🎛️ Tour de Contrôle (Multi-Tapis)"]
        default_idx = 0
        val_actuelle = st.session_state.get("radio_tapis_direct")
        if val_actuelle in options_tapis:
            default_idx = options_tapis.index(val_actuelle)

        if not is_kiosque:
            col_tap_sel, col_stat_glob = st.columns([1, 1])
            with col_tap_sel:
                t_choisi = st.radio(
                    "🥋 **Sélectionnez le Tapis ou la Vue Globale :**", 
                    options_tapis, 
                    index=default_idx,
                    horizontal=True,
                    key="radio_tapis_direct"
                )
        else:
            t_choisi = st.session_state.get("radio_tapis_direct", options_tapis[0])
            if t_choisi not in options_tapis:
                t_choisi = options_tapis[0]
            col_stat_glob = None

        def fragment_compat_deroule(run_every="5s"):
            def decorator(fn):
                if hasattr(st, "fragment"):
                    return st.fragment(run_every=run_every)(fn)
                elif hasattr(st, "experimental_fragment"):
                    return st.experimental_fragment(run_every=5)(fn)
                return fn
            return decorator

        if t_choisi == "📋 Déroulé Général (Tous Tapis)":
            col_titre_d, col_btn_act = st.columns([3, 1])
            with col_titre_d:
                if is_kiosque:
                    st.markdown("<h2 style='margin: 0; color: #0055A4;'>📋 DÉROULÉ DES COMBATS EN DIRECT</h2>", unsafe_allow_html=True)
                    st.markdown(
                        f"<div style='display: flex; align-items: center; gap: 8px; margin-top: 4px;'>"
                        f"<span style='height: 10px; width: 10px; background-color: #00c853; border-radius: 50%; display: inline-block; box-shadow: 0 0 8px #00c853;'></span>"
                        f"<span style='font-size: 13px; font-weight: 600; color: #2e7d32;'>DIRECT ACTIF · Synchronisation auto toutes les 5s</span>"
                        f"<span style='color: #888; font-size: 12px;'>| 🏆 {nom_competition}</span>"
                        f"</div>",
                        unsafe_allow_html=True
                    )
                else:
                    st.markdown("### 📋 Déroulé des Combats en Direct (Écran d'Appel)")
                    st.markdown(
                        f"<div style='display: flex; align-items: center; gap: 8px; margin-top: 2px;'>"
                        f"<span style='height: 9px; width: 9px; background-color: #00c853; border-radius: 50%; display: inline-block; box-shadow: 0 0 6px #00c853;'></span>"
                        f"<span style='font-size: 12px; font-weight: 600; color: #2e7d32;'>Direct actif (Actualisation automatique toutes les 5s)</span>"
                        f"</div>",
                        unsafe_allow_html=True
                    )
            with col_btn_act:
                if st.button("🔄 Actualiser maintenant", use_container_width=True, type="primary" if is_kiosque else "secondary"):
                    st.rerun()

            # --- FILTRE PAR CLUB & ALERTES PRÉDICTIVES (COACHS & PARENTS) ---
            clubs_presents = sorted(list(set(
                str(m.get("club_rouge", "")).strip() for m in matchs_direct if str(m.get("club_rouge", "")).strip()
            ).union(
                set(str(m.get("club_bleu", "")).strip() for m in matchs_direct if str(m.get("club_bleu", "")).strip())
            )))
            options_clubs_filtre = ["🏟️ Tous les clubs (Vue Complète)"] + clubs_presents
            col_f_c1, col_f_c2 = st.columns([2.8, 1.2])
            with col_f_c1:
                club_choisi = st.selectbox(
                    "🔍 Filtrer par club (Coachs & Parents) :",
                    options_clubs_filtre,
                    index=0,
                    key="sb_filtre_club_deroule",
                    help="Isole les lutteurs de votre club et affiche les alertes d'échauffement et d'appel au tapis"
                )

            # 1. Grille des 3 Tapis : File d'appel des 8 prochains combats par tapis (Actualisation automatique)
            @fragment_compat_deroule(run_every="5s")
            def _afficher_grille_direct_fragment():
                # Re-fetch Supabase en direct pour récupérer immédiatement les victoires enregistrées
                params_q_live = {"order": "tapis.asc,match_num.asc"}
                if c_sess and c_sess not in ["FFLDA-ADMIN", "FFLDA2026"]:
                    params_q_live["code_organisateur"] = f"eq.{c_sess}"
                
                matchs_live = list(matchs_direct)
                if is_supa:
                    res_live, _ = supabase_request("matchs_lutte", params=params_q_live)
                    if res_live and isinstance(res_live, list) and len(res_live) > 0:
                        tournois_ids = list(dict.fromkeys([m.get("tournoi_id", "") for m in res_live if m.get("tournoi_id")]))
                        if len(tournois_ids) > 1:
                            t_nom_choisi = st.session_state.get("nom_competition_active") or tournois_ids[-1]
                            matchs_live = [m for m in res_live if m.get("tournoi_id") == t_nom_choisi]
                        else:
                            matchs_live = res_live

                # Bandeau d'alerte spécifique Club sélectionné
                if club_choisi != "🏟️ Tous les clubs (Vue Complète)":
                    alertes_club = []
                    for t_chk in tapis_dispos:
                        a_venir_t = [m for m in matchs_live if int(m.get("tapis", 1)) == t_chk and m.get("statut") != "Terminé"]
                        for idx_pos, m_c in enumerate(a_venir_t[:6]):
                            cr_c = str(m_c.get("club_rouge", "")).strip()
                            cb_c = str(m_c.get("club_bleu", "")).strip()
                            if club_choisi in [cr_c, cb_c]:
                                lutteur_nom = m_c.get("lutteur_rouge") if club_choisi == cr_c else m_c.get("lutteur_bleu")
                                couleur_coin = "🔴 Rouge" if club_choisi == cr_c else "🔵 Bleu"
                                if idx_pos == 0:
                                    alertes_club.append(f"<span style='background:#b91c1c; color:white; padding:4px 10px; border-radius:8px; font-weight:800; font-size:12px;'>🔥 EN COMBAT — Tapis {t_chk} : {lutteur_nom} ({couleur_coin}, N°{m_c.get('match_num')})</span>")
                                elif idx_pos == 1:
                                    alertes_club.append(f"<span style='background:#ea580c; color:white; padding:4px 10px; border-radius:8px; font-weight:800; font-size:12px;'>⚡ COMBAT SUIVANT — Tapis {t_chk} : {lutteur_nom} ({couleur_coin}, N°{m_c.get('match_num')})</span>")
                                elif idx_pos == 2:
                                    alertes_club.append(f"<span style='background:#ca8a04; color:white; padding:4px 10px; border-radius:8px; font-weight:800; font-size:12px;'>⏳ DANS 2 COMBATS (Échauffement) — Tapis {t_chk} : {lutteur_nom} ({couleur_coin}, N°{m_c.get('match_num')})</span>")
                                else:
                                    alertes_club.append(f"<span style='background:#334155; color:white; padding:4px 10px; border-radius:8px; font-weight:800; font-size:12px;'>📋 Dans {idx_pos} combats — Tapis {t_chk} : {lutteur_nom} ({couleur_coin}, N°{m_c.get('match_num')})</span>")

                    if alertes_club:
                        st.markdown(
                            f"<div style='background:#fefce8; border:2px solid #eab308; border-radius:12px; padding:10px 14px; margin-bottom:12px; display:flex; flex-direction:column; gap:6px; box-shadow:0 4px 12px rgba(234,179,8,0.22);'>"
                            f"<div style='font-size:13px; font-weight:900; color:#854d0e;'>📢 ALERTES PASSAGE POUR {club_choisi.upper()} :</div>"
                            f"<div style='display:flex; flex-wrap:wrap; gap:8px;'>{' '.join(alertes_club)}</div>"
                            f"</div>",
                            unsafe_allow_html=True
                        )

                cols_tapis = st.columns(len(tapis_dispos))
                
                for i_ts, t_num in enumerate(tapis_dispos):
                    # Récupérer les combats NON terminés de ce tapis (file glissante)
                    matchs_t_a_venir = [m for m in matchs_live if int(m.get("tapis", 1)) == t_num and m.get("statut") != "Terminé"]
                    matchs_t_termines = [m for m in matchs_live if int(m.get("tapis", 1)) == t_num and m.get("statut") == "Terminé"]
                    tot_t = len(matchs_t_a_venir) + len(matchs_t_termines)
                    term_t = len(matchs_t_termines)
                    pct_t = int(term_t / tot_t * 100) if tot_t > 0 else 0
                    
                    prochains_8 = matchs_t_a_venir[:8]
                    
                    with cols_tapis[i_ts]:
                        st.markdown(
                            f"<div style='background: #0055A4; color: white; padding: 10px; border-radius: 8px 8px 0 0; text-align: center; font-weight: bold; font-size: 18px;'>"
                            f"TAPIS {t_num}"
                            f"</div>",
                            unsafe_allow_html=True
                        )
                        st.caption(f"Avancement : {term_t}/{tot_t} terminés ({pct_t}%) | {len(matchs_t_a_venir)} en attente")
                        st.progress(pct_t / 100)
                        
                        if not prochains_8:
                            st.success("🎉 Tous les combats de ce tapis sont terminés !")
                        else:
                            for idx_p, m_f in enumerate(prochains_8):
                                cr_f = str(m_f.get("club_rouge", "")).strip()
                                cb_f = str(m_f.get("club_bleu", "")).strip()
                                est_club_sel = (club_choisi != "🏟️ Tous les clubs (Vue Complète)") and (club_choisi in [cr_f, cb_f])
                                badge_mon_club = "<span style='background:#fef08a; color:#854d0e; font-size:10px; font-weight:900; padding:2px 7px; border-radius:6px; border:1px solid #eab308; margin-left:6px;'>⭐ MON CLUB</span>" if est_club_sel else ""

                                # Le 1er match est EN COURS SUR LE TAPIS
                                if idx_p == 0:
                                    border_live = "4px solid #eab308" if est_club_sel else "3px solid #2e7d32"
                                    st.markdown(
                                        f"<div style='border: {border_live}; border-radius: 10px; padding: 12px; margin-bottom: 12px; background: #e8f5e9; box-shadow: 0 4px 8px rgba(0,0,0,0.12);'>"
                                        f"<div style='text-align: center; margin-bottom: 6px;'>"
                                        f"<span style='background: #2e7d32; color: white; padding: 4px 12px; border-radius: 12px; font-weight: bold; font-size: 13px; letter-spacing: 1px;'>🔥 EN COURS SUR LE TAPIS</span>{badge_mon_club}"
                                        f"</div>"
                                        f"<div style='text-align: center; font-size: 15px; font-weight: bold; margin-bottom: 4px;'>Combat n°{m_f.get('match_num')} — 🕘 {m_f.get('heure')}</div>"
                                        f"<div style='text-align: center; font-size: 12px; color: #555; margin-bottom: 8px;'>{m_f.get('categorie')} — {m_f.get('tour')}</div>"
                                        f"<div style='background: #d32f2f; color: white; padding: 6px; border-radius: 6px; font-weight: bold; font-size: 14px; text-align: center; margin-bottom: 4px;'>"
                                        f"🔴 {m_f.get('lutteur_rouge')} <small style='opacity:0.85;'>({m_f.get('club_rouge','')})</small>"
                                        f"</div>"
                                        f"<div style='background: #1976d2; color: white; padding: 6px; border-radius: 6px; font-weight: bold; font-size: 14px; text-align: center;'>"
                                        f"🔵 {m_f.get('lutteur_bleu')} <small style='opacity:0.85;'>({m_f.get('club_bleu','')})</small>"
                                        f"</div>"
                                        f"</div>",
                                        unsafe_allow_html=True
                                    )
                                else:
                                    # Les matchs suivants sont en préparation (avec leur VRAI numéro de combat officiel)
                                    is_next_one = (idx_p == 1)
                                    if est_club_sel:
                                        if is_next_one:
                                            badge_stat = "<span style='font-size: 10px; background: #ea580c; color: white; padding: 2px 8px; border-radius: 8px; font-weight: 800;'>⚡ PRÊT AU BORD DU TAPIS</span>"
                                        elif idx_p == 2:
                                            badge_stat = "<span style='font-size: 10px; background: #ca8a04; color: white; padding: 2px 8px; border-radius: 8px; font-weight: 800;'>⏳ DANS 2 COMBATS</span>"
                                        else:
                                            badge_stat = f"<span style='font-size: 10px; background: #fef08a; color: #854d0e; padding: 2px 6px; border-radius: 8px; font-weight: 800;'>⭐ Dans {idx_p} combats</span>"
                                        border_left_col = "#eab308"
                                    else:
                                        badge_stat = (
                                            "<span style='font-size: 10px; background: #fff3e0; color: #e65100; padding: 2px 8px; border-radius: 8px; font-weight: 700; border: 1px solid #ffe0b2;'>⚡ Combat Suivant</span>"
                                            if is_next_one else
                                            "<span style='font-size: 10px; background: #f5f5f5; color: #616161; padding: 2px 6px; border-radius: 8px; font-weight: 600;'>⏳ En attente</span>"
                                        )
                                        border_left_col = "#f57c00" if is_next_one else "#90caf9"

                                    card_bg = "#fefce8" if est_club_sel else "white"
                                    card_border = "2px solid #eab308" if est_club_sel else "1px solid #e0e0e0"

                                    st.markdown(
                                        f"<div style='border: {card_border}; border-left: 5px solid {border_left_col}; border-radius: 6px; padding: 8px 10px; margin-bottom: 6px; background: {card_bg};'>"
                                        f"<div style='display: flex; justify-content: space-between; align-items: center;'>"
                                        f"<span style='font-size: 13px; font-weight: bold; color: #0d47a1;'>Combat n°{m_f.get('match_num')} <span style='font-size: 11px; font-weight: normal; color: #666;'>— 🕘 {m_f.get('heure')}</span></span>"
                                        f"{badge_stat}"
                                        f"</div>"
                                        f"<div style='font-size: 11px; color: #555; margin: 3px 0;'>{m_f.get('categorie')} · {m_f.get('tour')}{badge_mon_club}</div>"
                                        f"<div style='font-size: 12px; margin-top: 3px;'>"
                                        f"<span style='color: #c62828; font-weight: 600;'>🔴 {m_f.get('lutteur_rouge')}</span> <span style='color: #888;'>vs</span> <span style='color: #1565c0; font-weight: 600;'>🔵 {m_f.get('lutteur_bleu')}</span>"
                                        f"</div>"
                                        f"</div>",
                                        unsafe_allow_html=True
                                    )

            _afficher_grille_direct_fragment()

            # Fallback auto-refresh pour les environnements sans st.fragment
            if not hasattr(st, "fragment") and not hasattr(st, "experimental_fragment"):
                try:
                    from streamlit_autorefresh import st_autorefresh
                    st_autorefresh(interval=5000, key="auto_ref_deroule")
                except Exception:
                    try:
                        import streamlit.components.v1 as _components
                        _components.html(
                            """<script>
                            setTimeout(function() {
                                window.parent.location.reload();
                            }, 5000);
                            </script>""",
                            height=0,
                            width=0
                        )
                    except Exception:
                        pass

            st.markdown("---")

            # 2. Formulaire d'Édition / Correction (Superviseur) dans un expander
            if not is_kiosque:
                with st.expander("✏️ Éditer / Corriger un Combat du Tournoi (Superviseur)", expanded=False):
                    st.write("Sélectionnez n'importe quel combat parmi les 3 tapis pour renseigner le vainqueur, modifier les scores ou le réinitialiser :")
                    
                    options_tous_combats = [
                        f"Combat n°{m.get('match_num')} [Tapis {m.get('tapis')}] ({m.get('heure')}) : {m.get('lutteur_rouge')} vs {m.get('lutteur_bleu')} ({m.get('categorie')}) [{'✅' if m.get('statut')=='Terminé' else '⏳'}]"
                        for m in matchs_direct
                    ]
                    
                    sel_m_edit_str = st.selectbox(
                        "Choisir le combat à éditer :",
                        options_tous_combats,
                        key="sb_edit_combat_global"
                    )
                    
                    idx_edit_sel = options_tous_combats.index(sel_m_edit_str)
                    m_to_edit = matchs_direct[idx_edit_sel]
                    
                    col_e1, col_e2, col_e3, col_e4 = st.columns([2, 2, 2, 2])
                    with col_e1:
                        v_choix = st.radio("Vainqueur :", ["🔴 Rouge", "🔵 Bleu"], index=0 if (m_to_edit.get("vainqueur") or "Rouge") == "Rouge" else 1, horizontal=True, key=f"v_edit_{m_to_edit.get('id')}")
                        v_pur_e = "Rouge" if "Rouge" in v_choix else "Bleu"
                    with col_e2:
                        t_vic_options = ["VT (Tombé / Forfait)", "VST (Grande Supériorité)", "VP (Aux Points)"]
                        def_t_vic_idx = 0
                        if (m_to_edit.get("type_victoire") or "").startswith("VST"): def_t_vic_idx = 1
                        elif (m_to_edit.get("type_victoire") or "").startswith("VP"): def_t_vic_idx = 2
                        t_vic_choix = st.selectbox("Type :", t_vic_options, index=def_t_vic_idx, key=f"tv_edit_{m_to_edit.get('id')}")
                        code_type_e = t_vic_choix.split()[0]
                    with col_e3:
                        sc_r_e = st.number_input("Score Rouge :", min_value=0, max_value=99, value=int(m_to_edit.get("score_rouge", 0)), key=f"sc_r_edit_{m_to_edit.get('id')}")
                    with col_e4:
                        sc_b_e = st.number_input("Score Bleu :", min_value=0, max_value=99, value=int(m_to_edit.get("score_bleu", 0)), key=f"sc_b_edit_{m_to_edit.get('id')}")
                    
                    stat_choix = st.selectbox("Statut du match :", ["Terminé", "À venir"], index=0 if m_to_edit.get("statut") == "Terminé" else 1, key=f"stat_edit_{m_to_edit.get('id')}")
                    
                    pt_r_e, pt_b_e = calculer_pts_fflda_match(m_to_edit.get("categorie"), v_pur_e, code_type_e, sc_r_e, sc_b_e)
                    st.caption(f"Points calculés : Rouge = {pt_r_e} pts | Bleu = {pt_b_e} pts")

                    col_btn_e1, col_btn_e2 = st.columns(2)
                    with col_btn_e1:
                        if st.button("💾 Enregistrer la modification sur ce combat", type="primary", use_container_width=True, key=f"btn_save_edit_{m_to_edit.get('id')}"):
                            acts_r_e = decomposer_score_lutte(sc_r_e)
                            acts_b_e = decomposer_score_lutte(sc_b_e)
                            base_r = str(m_to_edit.get("tot_r_cell") or "").split("#")[0].strip()
                            base_b = str(m_to_edit.get("tot_b_cell") or "").split("#")[0].strip()
                            str_r = ",".join(str(a) for a in acts_r_e)
                            str_b = ",".join(str(a) for a in acts_b_e)
                            full_r = f"{base_r}#{str_r}" if (base_r and str_r) else (base_r or (f"#{str_r}" if str_r else ""))
                            full_b = f"{base_b}#{str_b}" if (base_b and str_b) else (base_b or (f"#{str_b}" if str_b else ""))

                            m_to_edit["statut"] = stat_choix
                            m_to_edit["vainqueur"] = v_pur_e if stat_choix == "Terminé" else ""
                            m_to_edit["type_victoire"] = code_type_e if stat_choix == "Terminé" else ""
                            m_to_edit["score_rouge"] = sc_r_e
                            m_to_edit["score_bleu"] = sc_b_e
                            m_to_edit["pt_clt_rouge"] = pt_r_e if stat_choix == "Terminé" else 0
                            m_to_edit["pt_clt_bleu"] = pt_b_e if stat_choix == "Terminé" else 0
                            m_to_edit["tot_r_cell"] = full_r
                            m_to_edit["tot_b_cell"] = full_b

                            acts_r_dicts = [{"val": v, "type": "tech", "label": str(v), "seq": idx+1} for idx, v in enumerate(acts_r_e)]
                            acts_b_dicts = [{"val": v, "type": "tech", "label": str(v), "seq": idx+1} for idx, v in enumerate(acts_b_e)]
                            m_to_edit["actions_rouge"] = acts_r_dicts
                            m_to_edit["actions_bleu"] = acts_b_dicts
                            sauvegarder_cache_actions_match(c_sess, m_to_edit.get('id'), acts_r_dicts, acts_b_dicts, 0, 0)

                            st.session_state.pop(f"actions_r_{m_to_edit.get('id')}", None)
                            st.session_state.pop(f"actions_b_{m_to_edit.get('id')}", None)
                            st.session_state.pop(f"cautions_r_{m_to_edit.get('id')}", None)
                            st.session_state.pop(f"cautions_b_{m_to_edit.get('id')}", None)
                            
                            publier_mise_a_jour_match(c_sess, m_to_edit)
                            publier_actions_match_sync(
                                m_to_edit.get('id'),
                                actions_r=acts_r_dicts,
                                actions_b=acts_b_dicts,
                                cautions_r=0,
                                cautions_b=0,
                                extra={
                                    "statut": stat_choix,
                                    "vainqueur": v_pur_e if stat_choix == "Terminé" else "",
                                    "type_victoire": code_type_e if stat_choix == "Terminé" else "",
                                    "score_r": sc_r_e,
                                    "score_b": sc_b_e
                                }
                            )
                            if is_supa:
                                match_payload = {
                                    "statut": stat_choix,
                                    "vainqueur": m_to_edit["vainqueur"],
                                    "type_victoire": m_to_edit["type_victoire"],
                                    "score_rouge": sc_r_e,
                                    "score_bleu": sc_b_e,
                                    "pt_clt_rouge": m_to_edit["pt_clt_rouge"],
                                    "pt_clt_bleu": m_to_edit["pt_clt_bleu"],
                                    "tot_r_cell": full_r,
                                    "tot_b_cell": full_b
                                }
                                supabase_request(f"matchs_lutte?id=eq.{m_to_edit.get('id')}", method="PATCH", data=match_payload)
                            
                            st.success(f"✅ Combat n°{m_to_edit.get('match_num')} [Tapis {m_to_edit.get('tapis')}] mis à jour avec succès !")
                            st.rerun()

                    with col_btn_e2:
                        if st.button("🔄 Réinitialiser ce combat à zéro (Remettre en attente)", use_container_width=True, key=f"btn_reset_edit_{m_to_edit.get('id')}"):
                            base_r = str(m_to_edit.get("tot_r_cell") or "").split("#")[0].strip()
                            base_b = str(m_to_edit.get("tot_b_cell") or "").split("#")[0].strip()
                            m_to_edit["statut"] = "À venir"
                            m_to_edit["vainqueur"] = ""
                            m_to_edit["type_victoire"] = ""
                            m_to_edit["score_rouge"] = 0
                            m_to_edit["score_bleu"] = 0
                            m_to_edit["pt_clt_rouge"] = 0
                            m_to_edit["pt_clt_bleu"] = 0
                            m_to_edit["tot_r_cell"] = base_r
                            m_to_edit["tot_b_cell"] = base_b
                            m_to_edit["actions_rouge"] = []
                            m_to_edit["actions_bleu"] = []
                            sauvegarder_cache_actions_match(c_sess, m_to_edit.get('id'), [], [], 0, 0)

                            st.session_state.pop(f"actions_r_{m_to_edit.get('id')}", None)
                            st.session_state.pop(f"actions_b_{m_to_edit.get('id')}", None)
                            st.session_state.pop(f"cautions_r_{m_to_edit.get('id')}", None)
                            st.session_state.pop(f"cautions_b_{m_to_edit.get('id')}", None)

                            if is_supa:
                                match_payload = {
                                    "statut": "À venir",
                                    "vainqueur": "",
                                    "type_victoire": "",
                                    "score_rouge": 0,
                                    "score_bleu": 0,
                                    "pt_clt_rouge": 0,
                                    "pt_clt_bleu": 0,
                                    "tot_r_cell": base_r,
                                    "tot_b_cell": base_b
                                }
                                supabase_request(f"matchs_lutte?id=eq.{m_to_edit.get('id')}", method="PATCH", data=match_payload)
                            st.info(f"Combat n°{m_to_edit.get('match_num')} remis en attente.")
                            st.rerun()

            # 3. Tableau Historique des Combats (Fil de l'eau complet) dans un expander
            with st.expander("📊 Voir le Tableau Complet de tous les combats (Historique & Recherche)", expanded=False):
                col_f1, col_f2, col_f3 = st.columns([1, 1, 2])
                with col_f1:
                    f_tapis = st.selectbox("Filtrer par Tapis :", ["Tous les Tapis"] + [f"Tapis {t}" for t in tapis_dispos], key="f_tapis_deroule")
                with col_f2:
                    f_statut = st.selectbox("Filtrer par Statut :", ["Tous", "⏳ À venir", "🟢 Terminé"], key="f_statut_deroule")
                with col_f3:
                    f_recherche = st.text_input("🔍 Rechercher :", placeholder="Nom, club ou catégorie...", key="f_rech_deroule")

                matchs_filtres = list(matchs_direct)
                if f_tapis != "Tous les Tapis":
                    t_num_f = int(f_tapis.replace("Tapis ", ""))
                    matchs_filtres = [m for m in matchs_filtres if int(m.get("tapis", 1)) == t_num_f]
                if f_statut == "⏳ À venir":
                    matchs_filtres = [m for m in matchs_filtres if m.get("statut") != "Terminé"]
                elif f_statut == "🟢 Terminé":
                    matchs_filtres = [m for m in matchs_filtres if m.get("statut") == "Terminé"]
                if f_recherche:
                    rech_low = f_recherche.lower()
                    matchs_filtres = [
                        m for m in matchs_filtres
                        if rech_low in str(m.get("lutteur_rouge", "")).lower()
                        or rech_low in str(m.get("lutteur_bleu", "")).lower()
                        or rech_low in str(m.get("club_rouge", "")).lower()
                        or rech_low in str(m.get("club_bleu", "")).lower()
                        or rech_low in str(m.get("categorie", "")).lower()
                    ]

                data_tableau = []
                for m in matchs_filtres:
                    stat_icon = "🟢 Terminé" if m.get("statut") == "Terminé" else "⏳ À venir"
                    score_str = f"{m.get('score_rouge', 0)} - {m.get('score_bleu', 0)}" if m.get("statut") == "Terminé" else "-"
                    vainq_str = m.get("vainqueur", "-")
                    if vainq_str == "Rouge":
                        vainq_str = f"🔴 {m.get('lutteur_rouge')}"
                    elif vainq_str == "Bleu":
                        vainq_str = f"🔵 {m.get('lutteur_bleu')}"
                        
                    data_tableau.append({
                        "Tapis": f"Tapis {m.get('tapis')}",
                        "N°": m.get("match_num"),
                        "Heure": m.get("heure"),
                        "Catégorie": m.get("categorie"),
                        "Tour": m.get("tour"),
                        "🔴 Rouge": f"{m.get('lutteur_rouge')} ({m.get('club_rouge','')})",
                        "Score": score_str,
                        "🔵 Bleu": f"{m.get('lutteur_bleu')} ({m.get('club_bleu','')})",
                        "Vainqueur": vainq_str,
                        "Type": formater_label_court_victoire(m.get("type_victoire", "-")),
                        "Statut": stat_icon
                    })

                if data_tableau:
                    df_deroule = pd.DataFrame(data_tableau)
                    st.dataframe(df_deroule, use_container_width=True, hide_index=True)
                else:
                    st.info("Aucun combat ne correspond aux critères.")

        elif t_choisi == "🎛️ Tour de Contrôle (Multi-Tapis)":
            col_titre_tc, col_btn_tc = st.columns([3, 1])
            with col_titre_tc:
                st.markdown("<h2 style='margin: 0; color: #0055A4;'>🎛️ TOUR DE CONTRÔLE — MULTI-TAPIS EN DIRECT</h2>", unsafe_allow_html=True)
                st.caption("Surveillance en temps réel de tous les tapis de lutte du tournoi (scores, chronomètres, statuts et accès direct).")
            with col_btn_tc:
                if st.button("🔄 Actualiser Tour de Contrôle", use_container_width=True, key="btn_ref_tc"):
                    st.rerun()

            @fragment_compat(run_every="2s")
            def _fragment_tour_de_controle():
                cols_tc = st.columns(len(tapis_dispos))
                for i_tc, t_num in enumerate(tapis_dispos):
                    with cols_tc[i_tc]:
                        etat_tapis = recuperer_etat_tapis(t_num)
                        matchs_tapis = [m for m in matchs_direct if int(m.get("tapis", 1)) == t_num]
                        tot_m = len(matchs_tapis)
                        nb_term = sum(1 for m in matchs_tapis if m.get("statut") == "Terminé")
                        cur_idx = etat_tapis.get("active_idx")
                        if cur_idx is None or not (0 <= cur_idx < tot_m):
                            cur_idx = 0
                        
                        m_tc = dict(matchs_tapis[cur_idx]) if (matchs_tapis and 0 <= cur_idx < tot_m) else {}
                        m_tc_id = m_tc.get("id") or f"T{t_num}_M{m_tc.get('match_num', cur_idx)}"
                        if m_tc_id in _SHARED_MATCHS_DATA:
                            m_tc.update(_SHARED_MATCHS_DATA[m_tc_id])

                        sh_tc = recuperer_actions_match_sync(m_tc_id)
                        acts_r = (sh_tc and sh_tc.get("actions_r")) or st.session_state.get(f"actions_r_{m_tc_id}", [])
                        acts_b = (sh_tc and sh_tc.get("actions_b")) or st.session_state.get(f"actions_b_{m_tc_id}", [])
                        if not acts_r and not acts_b and m_tc:
                            acts_r = charger_actions_match(m_tc_id, "Rouge", int(m_tc.get("score_rouge", 0) or 0), m_tc, c_sess)
                            acts_b = charger_actions_match(m_tc_id, "Bleu", int(m_tc.get("score_bleu", 0) or 0), m_tc, c_sess)

                        sc_r = sum(a.get("val", 0) for a in acts_r)
                        sc_b = sum(a.get("val", 0) for a in acts_b)

                        cat_tc = m_tc.get("categorie", "")
                        sec_tc, run_tc = calculer_temps_restant_chrono(m_tc_id, cat_tc, t_num)
                        chr_txt = formater_chrono_mm_ss(sec_tc)
                        p30_tc = bool(recuperer_chrono_sync(m_tc_id, t_num, cat_tc).get("pause30", False))
                        if p30_tc:
                            col_chr = "#f59e0b"
                            lbl_tc = "⏸️ PAUSE (30s)"
                        else:
                            col_chr = "#ef4444" if sec_tc == 0 else ("#22c55e" if run_tc else "#f59e0b")
                            lbl_tc = "● EN COMBAT" if run_tc else ("🔔 FIN DU TEMPS" if sec_tc == 0 else "⏸️ PAUSE")

                        nom_r = m_tc.get('lutteur_rouge', 'Rouge')
                        club_r = m_tc.get('club_rouge', '')
                        nom_b = m_tc.get('lutteur_bleu', 'Bleu')
                        club_b = m_tc.get('club_bleu', '')

                        st.markdown(
                            f"<div style='border: 3px solid #334155; border-radius: 16px; background: #0f172a; color: white; padding: 12px; box-shadow: 0 6px 18px rgba(0,0,0,0.35); margin-bottom: 12px; box-sizing: border-box;'>"
                            f"<div style='display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #1e293b; padding-bottom: 8px; margin-bottom: 10px;'>"
                            f"<span style='background: #0055A4; color: white; font-weight: 900; font-size: 15px; padding: 4px 12px; border-radius: 8px;'>TAPIS {t_num}</span>"
                            f"<span style='color: {col_chr}; font-weight: 800; font-size: 12px;'>{lbl_tc}</span>"
                            f"</div>"
                            f"<div style='text-align: center; margin-bottom: 8px;'>"
                            f"<span style='font-size: 14px; font-weight: 800; color: #cbd5e1;'>Combat n°{m_tc.get('match_num', '-')}</span> "
                            f"<span style='font-size: 12px; color: #94a3b8;'>({cat_tc} · Tour {m_tc.get('tour', '-')})</span>"
                            f"</div>"
                            f"<div style='background: #000; border: 2px solid {col_chr}; border-radius: 10px; padding: 6px; text-align: center; margin-bottom: 12px;'>"
                            f"<span style='font-family: monospace; font-size: 28px; font-weight: 900; color: {col_chr}; letter-spacing: 1px;'>⏱️ {chr_txt}</span>"
                            f"</div>"
                            f"<div style='display: grid; grid-template-columns: 1fr auto 1fr; gap: 8px; align-items: center; margin-bottom: 12px;'>"
                            f"<div style='background: #7f1d1d; border: 1px solid #ef4444; border-radius: 10px; padding: 8px; text-align: center; overflow: hidden;'>"
                            f"<div style='font-size: 10px; font-weight: 800; color: #fca5a5;'>🔴 ROUGE</div>"
                            f"<div style='font-size: 13px; font-weight: 900; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;' title='{nom_r}'>{nom_r}</div>"
                            f"<div style='font-size: 10px; color: #fecaca; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;'>{club_r}</div>"
                            f"<div style='font-family: monospace; font-size: 26px; font-weight: 900; color: #ff8080; margin-top: 4px;'>{sc_r}</div>"
                            f"</div>"
                            f"<div style='font-size: 13px; font-weight: 900; color: #64748b;'>VS</div>"
                            f"<div style='background: #1e3a8a; border: 1px solid #3b82f6; border-radius: 10px; padding: 8px; text-align: center; overflow: hidden;'>"
                            f"<div style='font-size: 10px; font-weight: 800; color: #93c5fd;'>🔵 BLEU</div>"
                            f"<div style='font-size: 13px; font-weight: 900; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;' title='{nom_b}'>{nom_b}</div>"
                            f"<div style='font-size: 10px; color: #bfdbfe; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;'>{club_b}</div>"
                            f"<div style='font-family: monospace; font-size: 26px; font-weight: 900; color: #60a5fa; margin-top: 4px;'>{sc_b}</div>"
                            f"</div>"
                            f"</div>"
                            f"<div style='font-size: 11px; color: #94a3b8; display: flex; justify-content: space-between; margin-bottom: 4px;'>"
                            f"<span>Progression</span><span>{nb_term}/{tot_m} ({int(nb_term/tot_m*100) if tot_m else 0}%)</span>"
                            f"</div>"
                            f"</div>",
                            unsafe_allow_html=True
                        )
                        st.progress(nb_term / tot_m if tot_m else 0.0)

                        c_btn1, c_btn2 = st.columns(2)
                        with c_btn1:
                            if st.button(f"📱 Table Tapis {t_num}", key=f"btn_tc_tbl_{t_num}", use_container_width=True):
                                st.session_state["radio_tapis_direct"] = f"Tapis {t_num}"
                                st.session_state["vue_scoreboard_active"] = False
                                st.rerun()
                        with c_btn2:
                            if st.button(f"📺 Scoreboard {t_num}", key=f"btn_tc_scb_{t_num}", use_container_width=True):
                                st.session_state["radio_tapis_direct"] = f"Tapis {t_num}"
                                st.session_state["vue_scoreboard_active"] = True
                                st.rerun()

            _fragment_tour_de_controle()

        else:
            tapis_num_actif = int(t_choisi.replace("Tapis ", ""))

            def afficher_scoreboard_tapis(t_num, m_cur, liste_m, i_cur, sess_code):
                """
                Affiche le grand Scoreboard officiel plein écran haute visibilité pour le tapis.
                Conçu pour tablette ou TV au bord du tapis, face aux lutteurs, entraîneurs et public.
                """
                # Injection de style plein écran universel (Smart TV, tablettes, vidéoprojecteurs)
                st.markdown('''
                <style>
                header[data-testid="stHeader"] { display: none !important; }
                footer { display: none !important; }
                #MainMenu { visibility: hidden !important; }
                section[data-testid="stSidebar"] { display: none !important; }
                div[data-testid="collapsedControl"] { display: none !important; }
                .block-container {
                    padding-top: 0.2rem !important;
                    padding-bottom: 0.2rem !important;
                    padding-left: 0.6rem !important;
                    padding-right: 0.6rem !important;
                    max-width: 100% !important;
                }
                div[data-testid="stVerticalBlock"] {
                    gap: 0.3rem !important;
                }
                @keyframes scorePulseAnim {
                    0% { transform: scale(1); filter: drop-shadow(0 0 4px rgba(255,255,255,0.2)); }
                    50% { transform: scale(1.08); filter: drop-shadow(0 0 35px rgba(255,255,255,0.95)); }
                    100% { transform: scale(1); filter: drop-shadow(0 0 4px rgba(255,255,255,0.2)); }
                }
                .score-pulse {
                    animation: scorePulseAnim 0.6s ease-out !important;
                }
                </style>
                ''', unsafe_allow_html=True)

                col_head_scb1, col_head_scb2 = st.columns([1, 1])
                with col_head_scb1:
                    st.markdown(
                        f"<div style='margin-bottom: 4px;'>"
                        f"<span style='background:#0055A4; color:white; font-size:clamp(14px, 1.5vw, 18px); font-weight:900; padding:5px 14px; border-radius:8px; display:inline-block;'>TAPIS {t_num}</span>"
                        f"</div>",
                        unsafe_allow_html=True
                    )
                with col_head_scb2:
                    components.html("""
                    <script>
                    (function() {
                        var pDoc = window.parent.document;
                        var pWin = window.parent || window;

                        // Fonction universelle Web Audio API Buzzer (sans MP3 externe)
                        pWin._playArenaBuzzer = function() {
                            try {
                                var AudioCtx = pWin.AudioContext || pWin.webkitAudioContext;
                                if (!AudioCtx) return;
                                var ctx = new AudioCtx();
                                if (ctx.state === 'suspended') { ctx.resume(); }
                                var now = ctx.currentTime;
                                // 3 tonalités d'avertisseur de salle
                                [0, 0.35, 0.70].forEach(function(offset) {
                                    var osc = ctx.createOscillator();
                                    var gain = ctx.createGain();
                                    osc.type = 'sawtooth';
                                    osc.frequency.setValueAtTime(500, now + offset);
                                    osc.frequency.exponentialRampToValueAtTime(240, now + offset + 0.30);
                                    gain.gain.setValueAtTime(0.7, now + offset);
                                    gain.gain.exponentialRampToValueAtTime(0.01, now + offset + 0.30);
                                    osc.connect(gain);
                                    gain.connect(ctx.destination);
                                    osc.start(now + offset);
                                    osc.stop(now + offset + 0.30);
                                });
                            } catch(err) {
                                console.error("Erreur buzzer audio :", err);
                            }
                        };

                        // Activation automatique du Plein Écran
                        function declencherPleinEcranAuto() {
                            try {
                                var el = pDoc.documentElement;
                                var isFS = pDoc.fullscreenElement || pDoc.webkitFullscreenElement || pDoc.mozFullScreenElement || pDoc.msFullscreenElement;
                                if (!isFS) {
                                    if (el.requestFullscreen) {
                                        el.requestFullscreen().catch(function(){});
                                    } else if (el.webkitRequestFullscreen) {
                                        el.webkitRequestFullscreen();
                                    } else if (el.msRequestFullscreen) {
                                        el.msRequestFullscreen();
                                    }
                                }
                            } catch(e) {}
                        }

                        // Tentative d'activation automatique immédiate
                        declencherPleinEcranAuto();

                        // Enclenchement au premier clic ou toucher n'importe où sur l'écran
                        try {
                            pDoc.addEventListener('click', declencherPleinEcranAuto, {once: false});
                            pDoc.addEventListener('touchstart', declencherPleinEcranAuto, {once: false});
                        } catch(e) {}

                        try {
                            var metas = [
                                {name: 'apple-mobile-web-app-capable', content: 'yes'},
                                {name: 'apple-mobile-web-app-status-bar-style', content: 'black-translucent'},
                                {name: 'mobile-web-app-capable', content: 'yes'}
                            ];
                            metas.forEach(function(m) {
                                if (!pDoc.querySelector('meta[name="' + m.name + '"]')) {
                                    var meta = pDoc.createElement('meta');
                                    meta.name = m.name;
                                    meta.content = m.content;
                                    pDoc.head.appendChild(meta);
                                }
                            });
                        } catch(e) {}
                    })();
                    </script>
                    """, height=0)

                @fragment_compat(run_every="1s")
                def _fragment_scoreboard_tv():
                    # 1. Vérifier en temps réel le combat actif sélectionné par la table de marque
                    etat_tapis = recuperer_etat_tapis(t_num)
                    cur_idx = etat_tapis.get("active_idx")
                    if cur_idx is None or not (0 <= cur_idx < len(liste_m)):
                        cur_idx = st.session_state.get(f"idx_actif_tapis_{t_num}", i_cur)
                    if cur_idx is None or not (0 <= cur_idx < len(liste_m)):
                        cur_idx = 0
                    st.session_state[f"idx_actif_tapis_{t_num}"] = cur_idx

                    # 2. Données fraîches du combat en direct
                    m_live = dict(liste_m[cur_idx]) if (liste_m and 0 <= cur_idx < len(liste_m)) else {}
                    cur_m_id = m_live.get("id") or etat_tapis.get("match_id") or f"T{t_num}_M{m_live.get('match_num', cur_idx)}"
                    if cur_m_id in _SHARED_MATCHS_DATA:
                        m_live.update(_SHARED_MATCHS_DATA[cur_m_id])

                    cle_r = f"actions_r_{cur_m_id}"
                    cle_b = f"actions_b_{cur_m_id}"
                    cle_caut_r = f"cautions_r_{cur_m_id}"
                    cle_caut_b = f"cautions_b_{cur_m_id}"

                    # 3. Synchronisation multi-écrans en temps réel (actions de la table de marque)
                    sh_act = recuperer_actions_match_sync(cur_m_id)
                    if sh_act:
                        if "actions_r" in sh_act:
                            st.session_state[cle_r] = sh_act["actions_r"]
                        if "actions_b" in sh_act:
                            st.session_state[cle_b] = sh_act["actions_b"]
                        if "cautions_r" in sh_act:
                            st.session_state[cle_caut_r] = sh_act["cautions_r"]
                        if "cautions_b" in sh_act:
                            st.session_state[cle_caut_b] = sh_act["cautions_b"]
                        if "statut" in sh_act:
                            m_live["statut"] = sh_act["statut"]
                        if "vainqueur" in sh_act:
                            m_live["vainqueur"] = sh_act["vainqueur"]
                        if "type_victoire" in sh_act:
                            m_live["type_victoire"] = sh_act["type_victoire"]

                    actions_r = st.session_state.get(cle_r, [])
                    actions_b = st.session_state.get(cle_b, [])
                    cautions_r = st.session_state.get(cle_caut_r, 0)
                    cautions_b = st.session_state.get(cle_caut_b, 0)

                    if cle_r not in st.session_state and cle_b not in st.session_state and not sh_act and m_live:
                        init_sc_r = int(m_live.get("score_rouge", 0) or 0)
                        init_sc_b = int(m_live.get("score_bleu", 0) or 0)
                        actions_r = charger_actions_match(cur_m_id, "Rouge", init_sc_r, m_live, sess_code)
                        actions_b = charger_actions_match(cur_m_id, "Bleu", init_sc_b, m_live, sess_code)
                        st.session_state[cle_r] = actions_r
                        st.session_state[cle_b] = actions_b

                    sc_r = sum(a.get("val", 0) for a in actions_r)
                    sc_b = sum(a.get("val", 0) for a in actions_b)

                    nom_r = m_live.get('lutteur_rouge', 'Rouge')
                    club_r = m_live.get('club_rouge', '')
                    nom_b = m_live.get('lutteur_bleu', 'Bleu')
                    club_b = m_live.get('club_bleu', '')
                    cat_m = m_live.get('categorie', '')
                    tour_m = m_live.get('tour', '')
                    num_m = m_live.get('match_num', '')
                    heure_m = m_live.get('heure', '')

                    sec_rest, est_run = calculer_temps_restant_chrono(cur_m_id, cat_m, t_num)
                    chrono_txt = formater_chrono_mm_ss(sec_rest)
                    per_actuelle = st.session_state.get(f"chrono_per_{cur_m_id}", st.session_state.get(f"chrono_per_tapis_{t_num}", 1))
                    pause30 = st.session_state.get(f"chrono_pause30_{cur_m_id}", False)

                    # Gestion du Buzzer sonore de fin de combat / round / pause (Web Audio API)
                    doit_sonner_buzzer = False
                    etat_chr = recuperer_chrono_sync(cur_m_id, t_num, cat_m)
                    b_event = etat_chr.get("buzzer_event", 0.0)
                    last_b_played = st.session_state.get(f"last_buzzer_played_{cur_m_id}", 0.0)
                    if b_event > 0 and (pytime.time() - b_event < 3.5) and (b_event != last_b_played):
                        st.session_state[f"last_buzzer_played_{cur_m_id}"] = b_event
                        doit_sonner_buzzer = True
                    elif sec_rest == 0 and not est_run:
                        cle_buzzer = f"buzzer_fired_{cur_m_id}_{per_actuelle}"
                        if not st.session_state.get(cle_buzzer, False):
                            st.session_state[cle_buzzer] = True
                            doit_sonner_buzzer = True
                    elif sec_rest > 0:
                        cle_buzzer = f"buzzer_fired_{cur_m_id}_{per_actuelle}"
                        st.session_state[cle_buzzer] = False

                    # Détection d'attribution de points pour pulsation visuelle lumineuse
                    cle_prev_r = f"prev_sc_r_{cur_m_id}"
                    cle_prev_b = f"prev_sc_b_{cur_m_id}"
                    prev_r = st.session_state.get(cle_prev_r, sc_r)
                    prev_b = st.session_state.get(cle_prev_b, sc_b)
                    pulse_r = "score-pulse" if sc_r != prev_r else ""
                    pulse_b = "score-pulse" if sc_b != prev_b else ""
                    st.session_state[cle_prev_r] = sc_r
                    st.session_state[cle_prev_b] = sc_b

                    if doit_sonner_buzzer:
                        components.html("""
                        <script>
                        try {
                            if (window.parent && window.parent._playArenaBuzzer) {
                                window.parent._playArenaBuzzer();
                            }
                        } catch(e) {}
                        </script>
                        """, height=0)

                    meneur, motif_dep, _ = evaluer_departage_uww(actions_r, actions_b, cautions_r, cautions_b)
                    est_egalite = (sc_r == sc_b and sc_r > 0)
                    underline_r = "text-decoration: underline; text-decoration-color: #facc15; text-decoration-thickness: 6px; text-underline-offset: 8px;" if (est_egalite and meneur == "Rouge") else ""
                    underline_b = "text-decoration: underline; text-decoration-color: #facc15; text-decoration-thickness: 6px; text-underline-offset: 8px;" if (est_egalite and meneur == "Bleu") else ""

                    if pause30:
                        col_chr_color = "#f59e0b"
                        stat_badge = "⏸️ PAUSE RÉGLEMENTAIRE (30s)"
                        lbl_periode_tv = "⏸️ PAUSE (30s)"
                    else:
                        stat_badge = "● EN COMBAT" if est_run else ("🔔 TEMPS ÉCOULÉ" if sec_rest == 0 else "⏸️ TEMPS MORT")
                        col_chr_color = "#ef4444" if sec_rest == 0 else ("#22c55e" if est_run else "#f59e0b")
                        lbl_periode_tv = f"PÉRIODE {per_actuelle}"

                    # Détection de l'affichage de victoire plein écran pendant 10 secondes
                    dernier_termine = etat_tapis.get("dernier_termine")
                    afficher_victoire_10s = False
                    data_victoire = None
                    sec_rest_10s = 0

                    if dernier_termine and isinstance(dernier_termine, dict):
                        ts_term = float(dernier_termine.get("ts_termine", 0))
                        delta_t = pytime.time() - ts_term
                        if 0 <= delta_t < 10.0:
                            afficher_victoire_10s = True
                            data_victoire = dernier_termine
                            sec_rest_10s = max(1, int(10.0 - delta_t))

                    if not afficher_victoire_10s and (m_live.get('statut') == "Terminé"):
                        ts_fin_dir = st.session_state.setdefault(f"ts_fin_dir_{cur_m_id}", pytime.time())
                        delta_dir = pytime.time() - ts_fin_dir
                        if 0 <= delta_dir < 10.0:
                            afficher_victoire_10s = True
                            sec_rest_10s = max(1, int(10.0 - delta_dir))
                            data_victoire = {
                                "idx": cur_idx,
                                "match_id": cur_m_id,
                                "ts_termine": ts_fin_dir,
                                "vainqueur": m_live.get('vainqueur'),
                                "type_victoire": m_live.get('type_victoire'),
                                "score_r": sc_r,
                                "score_b": sc_b,
                                "lutteur_r": nom_r,
                                "club_r": club_r,
                                "lutteur_b": nom_b,
                                "club_b": club_b,
                                "match_num": num_m,
                                "categorie": cat_m,
                                "tour": tour_m
                            }
                        else:
                            # 10 secondes écoulées : passer automatiquement au combat suivant si disponible
                            if cur_idx + 1 < len(liste_m):
                                cur_idx = cur_idx + 1
                                st.session_state[f"idx_actif_tapis_{t_num}"] = cur_idx
                                m_live = dict(liste_m[cur_idx])
                                cur_m_id = m_live.get("id") or f"T{t_num}_M{m_live.get('match_num', cur_idx)}"
                                if cur_m_id in _SHARED_MATCHS_DATA:
                                    m_live.update(_SHARED_MATCHS_DATA[cur_m_id])
                                actions_r = charger_actions_match(cur_m_id, "Rouge", int(m_live.get("score_rouge", 0) or 0), m_live, sess_code)
                                actions_b = charger_actions_match(cur_m_id, "Bleu", int(m_live.get("score_bleu", 0) or 0), m_live, sess_code)
                                sc_r = sum(a.get("val", 0) for a in actions_r)
                                sc_b = sum(a.get("val", 0) for a in actions_b)
                                nom_r = m_live.get('lutteur_rouge', 'Rouge')
                                club_r = m_live.get('club_rouge', '')
                                nom_b = m_live.get('lutteur_bleu', 'Bleu')
                                club_b = m_live.get('club_bleu', '')
                                cat_m = m_live.get('categorie', '')
                                tour_m = m_live.get('tour', '')
                                num_m = m_live.get('match_num', '')
                                sec_rest, est_run = calculer_temps_restant_chrono(cur_m_id, cat_m, t_num)
                                chrono_txt = formater_chrono_mm_ss(sec_rest)
                                per_actuelle = st.session_state.get(f"chrono_per_{cur_m_id}", st.session_state.get(f"chrono_per_tapis_{t_num}", 1))
                                pause30 = st.session_state.get(f"chrono_pause30_{cur_m_id}", False)
                                if pause30:
                                    col_chr_color = "#f59e0b"
                                    stat_badge = "⏸️ PAUSE RÉGLEMENTAIRE (30s)"
                                    lbl_periode_tv = "⏸️ PAUSE (30s)"
                                else:
                                    stat_badge = "● EN COMBAT" if est_run else ("🔔 TEMPS ÉCOULÉ" if sec_rest == 0 else "⏸️ TEMPS MORT")
                                    col_chr_color = "#ef4444" if sec_rest == 0 else ("#22c55e" if est_run else "#f59e0b")
                                    lbl_periode_tv = f"PÉRIODE {per_actuelle}"
                                underline_r = ""
                                underline_b = ""

                    # Ligne info combat actif (titre, catégorie, tour)
                    titre_num = data_victoire.get("match_num", num_m) if (afficher_victoire_10s and data_victoire) else num_m
                    titre_cat = data_victoire.get("categorie", cat_m) if (afficher_victoire_10s and data_victoire) else cat_m
                    titre_tour = data_victoire.get("tour", tour_m) if (afficher_victoire_10s and data_victoire) else tour_m

                    st.markdown(
                        f"<div style='background:#f1f5f9; border: 2px solid #cbd5e1; border-radius: 12px; padding: 10px 18px; margin-bottom: 8px; display:flex; align-items:center; justify-content:space-between; box-shadow: 0 2px 8px rgba(0,0,0,0.06);'>"
                        f"<div><span style='font-size: clamp(18px, 2.2vw, 26px); font-weight:900; color:#0f172a;'>Combat n°{titre_num}</span></div>"
                        f"<div><span style='background:#0055A4; color:white; font-size: clamp(13px, 1.4vw, 16px); font-weight:800; padding:6px 14px; border-radius:8px; margin-right:8px;'>{titre_cat}</span><span style='background:#334155; color:white; font-size: clamp(13px, 1.4vw, 16px); font-weight:800; padding:6px 14px; border-radius:8px;'>Tour {titre_tour}</span></div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )

                    if afficher_victoire_10s and data_victoire:
                        # 🏆 ANNONCE DU VAINQUEUR PLEIN ÉCRAN (10 SECONDES)
                        v_nom_r = data_victoire.get("lutteur_r", nom_r)
                        v_club_r = data_victoire.get("club_r", club_r)
                        v_nom_b = data_victoire.get("lutteur_b", nom_b)
                        v_club_b = data_victoire.get("club_b", club_b)
                        v_sc_r = data_victoire.get("score_r", sc_r)
                        v_sc_b = data_victoire.get("score_b", sc_b)

                        v_raw = str(data_victoire.get('vainqueur') or "").strip()
                        if "rouge" in v_raw.lower() or v_raw == v_nom_r:
                            v_cote = "Rouge"
                        elif "bleu" in v_raw.lower() or v_raw == v_nom_b:
                            v_cote = "Bleu"
                        else:
                            if v_sc_r > v_sc_b:
                                v_cote = "Rouge"
                            elif v_sc_b > v_sc_r:
                                v_cote = "Bleu"
                            else:
                                v_cote = "Rouge"

                        if v_cote == "Rouge":
                            nom_vainq = v_nom_r
                            club_vainq = v_club_r
                            bg_grad = "linear-gradient(145deg, #7f1d1d 0%, #b91c1c 50%, #991b1b 100%)"
                            col_shadow = "rgba(220, 38, 38, 0.45)"
                        else:
                            nom_vainq = v_nom_b
                            club_vainq = v_club_b
                            bg_grad = "linear-gradient(145deg, #1e3a8a 0%, #1d4ed8 50%, #172554 100%)"
                            col_shadow = "rgba(37, 99, 235, 0.45)"

                        tv_code = str(data_victoire.get('type_victoire', '') or '').strip().upper()
                        if "VT" in tv_code or "TOMB" in tv_code:
                            vic_label = "⚡ VICTOIRE PAR TOMBÉ (VT)"
                            vic_bg = "#7e22ce"
                            vic_border = "#c084fc"
                        elif "VST" in tv_code or "VSU" in tv_code or "SUP" in tv_code:
                            vic_label = "💥 VICTOIRE PAR SUPÉRIORITÉ TECHNIQUE (VST)"
                            vic_bg = "#ea580c"
                            vic_border = "#fb923c"
                        elif "VP" in tv_code or "POINT" in tv_code:
                            vic_label = "🎯 VICTOIRE AUX POINTS (VP)"
                            vic_bg = "#2563eb"
                            vic_border = "#60a5fa"
                        elif any(k in tv_code for k in ["AB", "FOR", "DSQ", "DISQ"]):
                            vic_label = "🛑 VICTOIRE PAR FORFAIT / ABANDON"
                            vic_bg = "#475569"
                            vic_border = "#94a3b8"
                        elif tv_code:
                            vic_label = f"🏆 VICTOIRE ({tv_code})"
                            vic_bg = "#334155"
                            vic_border = "#64748b"
                        else:
                            vic_label = "🏆 VICTOIRE OFFICIELLE"
                            vic_bg = "#334155"
                            vic_border = "#64748b"

                        st.markdown(
                            f"<div style='background: {bg_grad}; border: 4px solid #facc15; border-radius: 24px; padding: clamp(16px, 3vw, 40px); box-shadow: 0 15px 50px {col_shadow}; text-align: center; color: white; min-height: 55vh; display: flex; flex-direction: column; justify-content: center; align-items: center; box-sizing: border-box; overflow: hidden; margin-bottom: 8px;'>"
                            f"<div style='color: #fde047; font-size: clamp(16px, 2.4vw, 34px); font-weight: 900; letter-spacing: 2px; text-transform: uppercase; margin-bottom: 6px;'>🏆 VAINQUEUR DU COMBAT 🏆</div>"
                            f"<div style='font-size: clamp(32px, 6vw, 92px); font-weight: 900; line-height: 1.1; margin: 10px 0 6px 0; text-shadow: 0 4px 18px rgba(0,0,0,0.7); word-break: break-word;'>{nom_vainq}</div>"
                            f"<div style='font-size: clamp(16px, 2.4vw, 32px); font-weight: 700; color: #f1f5f9; opacity: 0.95; margin-bottom: 18px;'>{club_vainq}</div>"
                            f"<div style='background: {vic_bg}; border: 2px solid {vic_border}; color: white; font-size: clamp(14px, 2vw, 26px); font-weight: 900; padding: clamp(8px, 1.2vw, 14px) clamp(16px, 2.5vw, 36px); border-radius: 9999px; margin-bottom: 22px; box-shadow: 0 4px 16px rgba(0,0,0,0.3); display: inline-block;'>{vic_label}</div>"
                            f"<div style='background: rgba(0,0,0,0.65); border: 2px solid rgba(255,255,255,0.2); border-radius: 16px; padding: clamp(8px, 1.2vw, 14px) clamp(14px, 2vw, 28px); display: inline-flex; align-items: center; justify-content: center; gap: clamp(8px, 1.5vw, 20px); flex-wrap: wrap;'>"
                            f"<span style='font-size: clamp(14px, 1.8vw, 22px); font-weight: 800; color: #ef4444;'>🔴 {v_nom_r} ({v_sc_r})</span>"
                            f"<span style='font-size: clamp(16px, 2vw, 24px); font-weight: 900; color: #94a3b8;'>—</span>"
                            f"<span style='font-size: clamp(14px, 1.8vw, 22px); font-weight: 800; color: #38bdf8;'>({v_sc_b}) {v_nom_b} 🔵</span>"
                            f"</div>"
                            f"<div style='margin-top: 14px; font-size: clamp(12px, 1.3vw, 16px); font-weight: 700; color: #fef08a; opacity: 0.9;'>⏳ Prochain combat dans {sec_rest_10s}s...</div>"
                            f"</div>",
                            unsafe_allow_html=True
                        )
                    else:
                        # GRANDE ARENA HAUTE VISIBILITÉ (3 Colonnes plein écran)
                        col_r, col_c, col_b = st.columns([2.3, 1.4, 2.3], gap="small")

                        # 🔴 COIN ROUGE
                        with col_r:
                            caut_r_html = " ".join(["<span style='background:#facc15; color:#78350f; font-weight:900; font-size:clamp(10px, 1.1vw, 13px); padding:3px 8px; border-radius:6px; margin-right:4px;'>⚠️ AVERT</span>" for _ in range(cautions_r)]) if cautions_r > 0 else ""
                            badge_av_r = "<div style='color: #facc15; font-size: clamp(10px, 1.1vw, 13px); font-weight: 900; letter-spacing: 1px; margin-top: 4px;'>👑 AVANTAGE UWW (DÉPARTAGE)</div>" if (est_egalite and meneur == "Rouge") else ""
                            st.markdown(
                                f"<div style='background: linear-gradient(145deg, #b91c1c 0%, #7f1d1d 100%); border: 3px solid #ef4444; border-radius: 20px; padding: clamp(8px, 1.5vw, 18px); box-shadow: 0 10px 30px rgba(220,38,38,0.3); text-align: center; color: white; min-height: 55vh; display: flex; flex-direction: column; justify-content: space-between; box-sizing: border-box; overflow: hidden;'>"
                                f"<div>"
                                f"<div style='font-size: clamp(18px, 2.8vw, 42px); font-weight: 900; line-height: 1.15; margin: 4px 0 2px 0; text-shadow: 0 2px 6px rgba(0,0,0,0.5); word-break: break-word; {underline_r}'>{nom_r}</div>"
                                f"<div style='font-size: clamp(11px, 1.4vw, 18px); font-weight: 700; opacity: 0.9; margin-bottom: 4px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;'>{club_r}</div>"
                                f"{badge_av_r}"
                                f"</div>"
                                f"<div class='{pulse_r}' style='background: #090d16; border: 3px solid #ef4444; border-radius: 16px; padding: clamp(6px, 1.2vw, 12px) 2px; margin: 8px 0; box-shadow: inset 0 0 25px rgba(239,68,68,0.35); flex-grow: 1; display: flex; align-items: center; justify-content: center; box-sizing: border-box; overflow: hidden; width: 100%;'>"
                                f"<span style='font-family: monospace; font-size: clamp(38px, 8.5vw, 150px); font-weight: 900; color: #ff4444; text-shadow: 0 0 30px rgba(255,68,68,0.85); line-height: 1; white-space: nowrap;'>{sc_r}</span>"
                                f"</div>"
                                f"<div style='min-height: 24px;'>{caut_r_html}</div>"
                                f"</div>",
                                unsafe_allow_html=True
                            )

                        # ⏱️ CHRONOMÈTRE CENTRAL (PUR AFFICHAGE, PAS DE BOUTONS)
                        with col_c:
                            centre_info_html = f"<div style='font-size: clamp(10px, 1.1vw, 12px); color: #cbd5e1; margin-top: 6px; font-weight: 600;'>{('⚖️ <b>Égalité</b> — Avantage <b style=\"color:#facc15;\">' + str(meneur) + '</b><br><span style=\"font-size:10px; color:#94a3b8;\">(' + motif_dep + ')</span>') if est_egalite else ''}</div>"

                            badge_per_style = "background: #78350f; color: #fde68a; border: 1px solid #f59e0b;" if pause30 else "background: #1e293b; color: #94a3b8;"
                            st.markdown(
                                f"<div style='background: #0f172a; border: 3px solid #334155; border-radius: 20px; padding: clamp(8px, 1.5vw, 18px); box-shadow: 0 10px 30px rgba(0,0,0,0.4); text-align: center; color: white; display: flex; flex-direction: column; justify-content: space-between; min-height: 55vh; box-sizing: border-box; overflow: hidden;'>"
                                f"<div>"
                                f"<div style='{badge_per_style} font-size: clamp(10px, 1.2vw, 14px); font-weight: 900; padding: 4px 10px; border-radius: 8px; display: inline-block; margin-bottom: 6px;'>{lbl_periode_tv}</div>"
                                f"<div style='background: #000000; border: 3px solid {col_chr_color}; border-radius: 16px; padding: clamp(6px, 1.2vw, 14px) 2px; margin: 6px 0 10px 0; box-shadow: inset 0 0 20px rgba(0,0,0,0.8); flex-grow: 1; display: flex; align-items: center; justify-content: center; box-sizing: border-box; overflow: hidden; width: 100%;'>"
                                f"<span style='font-family: monospace; font-size: clamp(18px, 4.2vw, 85px); font-weight: 900; color: {col_chr_color}; letter-spacing: clamp(0px, 0.2vw, 2px); text-shadow: 0 0 15px {col_chr_color}; line-height: 1; white-space: nowrap; max-width: 100%; display: inline-block;'>{chrono_txt}</span>"
                                f"</div>"
                                f"<div style='color: {col_chr_color}; font-size: clamp(10px, 1.1vw, 13px); font-weight: 900; letter-spacing: 0.5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;'>{stat_badge}</div>"
                                f"</div>"
                                f"<div style='margin-top: 6px;'>"
                                f"{centre_info_html}"
                                f"</div>"
                                f"</div>",
                                unsafe_allow_html=True
                            )

                        # 🔵 COIN BLEU
                        with col_b:
                            caut_b_html = " ".join(["<span style='background:#facc15; color:#78350f; font-weight:900; font-size:clamp(10px, 1.1vw, 13px); padding:3px 8px; border-radius:6px; margin-right:4px;'>⚠️ AVERT</span>" for _ in range(cautions_b)]) if cautions_b > 0 else ""
                            badge_av_b = "<div style='color: #facc15; font-size: clamp(10px, 1.1vw, 13px); font-weight: 900; letter-spacing: 1px; margin-top: 4px;'>👑 AVANTAGE UWW (DÉPARTAGE)</div>" if (est_egalite and meneur == "Bleu") else ""
                            st.markdown(
                                f"<div style='background: linear-gradient(145deg, #1d4ed8 0%, #1e3a8a 100%); border: 3px solid #3b82f6; border-radius: 20px; padding: clamp(8px, 1.5vw, 18px); box-shadow: 0 10px 30px rgba(37,99,235,0.3); text-align: center; color: white; min-height: 55vh; display: flex; flex-direction: column; justify-content: space-between; box-sizing: border-box; overflow: hidden;'>"
                                f"<div>"
                                f"<div style='font-size: clamp(18px, 2.8vw, 42px); font-weight: 900; line-height: 1.15; margin: 4px 0 2px 0; text-shadow: 0 2px 6px rgba(0,0,0,0.5); word-break: break-word; {underline_b}'>{nom_b}</div>"
                                f"<div style='font-size: clamp(11px, 1.4vw, 18px); font-weight: 700; opacity: 0.9; margin-bottom: 4px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;'>{club_b}</div>"
                                f"{badge_av_b}"
                                f"</div>"
                                f"<div class='{pulse_b}' style='background: #090d16; border: 3px solid #3b82f6; border-radius: 16px; padding: clamp(6px, 1.2vw, 12px) 2px; margin: 8px 0; box-shadow: inset 0 0 25px rgba(59,130,246,0.35); flex-grow: 1; display: flex; align-items: center; justify-content: center; box-sizing: border-box; overflow: hidden; width: 100%;'>"
                                f"<span style='font-family: monospace; font-size: clamp(38px, 8.5vw, 150px); font-weight: 900; color: #38bdf8; text-shadow: 0 0 30px rgba(56,189,248,0.85); line-height: 1; white-space: nowrap;'>{sc_b}</span>"
                                f"</div>"
                                f"<div style='min-height: 24px;'>{caut_b_html}</div>"
                                f"</div>",
                                unsafe_allow_html=True
                            )

                    # Bandeau Combat suivant en préparation (compact)
                    nxt_idx = cur_idx if (afficher_victoire_10s and data_victoire and cur_idx != data_victoire.get("idx")) else (cur_idx + 1)
                    if liste_m and 0 <= nxt_idx < len(liste_m):
                        nxt = liste_m[nxt_idx]
                        st.markdown(
                            f"<div style='margin-top: 10px; background: #f8fafc; border: 2px solid #e2e8f0; border-radius: 12px; padding: 8px 16px; display: flex; align-items: center; justify-content: space-between;'>"
                            f"<span style='font-size: 11px; font-weight: 900; color: #d97706; text-transform: uppercase; letter-spacing: 1px;'>⚡ COMBAT SUIVANT (EN PRÉPARATION) :</span>"
                            f"<span style='font-size: 14px; font-weight: 700; color: #1e293b;'>"
                            f"<span style='color: #dc2626;'>🔴 {nxt.get('lutteur_rouge','')}</span> ({nxt.get('club_rouge','')}) vs <span style='color: #2563eb;'>🔵 {nxt.get('lutteur_bleu','')}</span> ({nxt.get('club_bleu','')})"
                            f"</span>"
                            f"<span style='font-size: 12px; color: #64748b; font-weight: 600;'>Combat n°{nxt.get('match_num')}</span>"
                            f"</div>",
                            unsafe_allow_html=True
                        )

                _fragment_scoreboard_tv()
            
            # --- VÉRIFICATION SÉCURITÉ ACCÈS TABLE DE MARQUE (OPTION A) ---
            cle_attendue = get_cle_secrete_table(c_sess)
            qp_cur = getattr(st, "query_params", {})
            cle_url = qp_cur.get("cle", "")
            est_autorise = st.session_state.get("table_autorisee", False) or (cle_url == cle_attendue) or (not is_kiosque) or st.session_state.get("vue_scoreboard_active", False)

            if is_kiosque and not est_autorise:
                st.markdown(
                    f"<div style='background: #ffebee; border: 2px solid #c62828; border-radius: 12px; padding: 25px; text-align: center; margin: 30px auto; max-width: 600px; box-shadow: 0 4px 10px rgba(0,0,0,0.1);'>"
                    f"<div style='font-size: 45px; margin-bottom: 12px;'>🔒</div>"
                    f"<h2 style='color: #c62828; margin: 0 0 10px 0;'>Accès Table de Marque Réservé</h2>"
                    f"<p style='color: #222; font-size: 15px; margin: 0 0 12px 0;'>"
                    f"La saisie des scores et des résultats sur le <b>Tapis {tapis_num_actif}</b> est strictement réservée aux <b>arbitres officiels</b>."
                    f"</p>"
                    f"<p style='color: #666; font-size: 13px; margin: 0;'>"
                    f"👉 Pour arbitrer ce tapis, vous devez scanner le <b>QR Code officiel physique</b> présent sur la table de marque."
                    f"</p>"
                    f"</div>",
                    unsafe_allow_html=True
                )
                col_b_pub1, col_b_pub2, col_b_pub3 = st.columns([1, 2, 1])
                with col_b_pub2:
                    url_deroule_public = f"{url_prefix}/?code={c_sess}&mode=direct&tapis=all"
                    st.link_button("📋 Accéder au Déroulé Public (Tous Tapis)", url_deroule_public, use_container_width=True, type="primary")
                st.stop()
            
            matchs_du_tapis = [m for m in matchs_direct if int(m.get("tapis", 1)) == tapis_num_actif]
            nb_total_tapis = len(matchs_du_tapis)
            nb_termines_tapis = sum(1 for m in matchs_du_tapis if m.get("statut") == "Terminé")
            pct_tapis = int((nb_termines_tapis / nb_total_tapis * 100)) if nb_total_tapis > 0 else 0
        
            if not is_kiosque and col_stat_glob:
                with col_stat_glob:
                    st.markdown(f"**Avancement Tapis {tapis_num_actif} :** {nb_termines_tapis} / {nb_total_tapis} combats terminés ({pct_tapis}%)")
                    st.progress(pct_tapis / 100)
            elif not st.session_state.get("vue_scoreboard_active"):
                col_k_h1, col_k_h2 = st.columns([3, 1])
                with col_k_h1:
                    st.markdown(f"<div style='background: #0055A4; color: white; padding: 6px 14px; border-radius: 8px; font-weight: bold; font-size: 19px; display: inline-block;'>TABLE DE MARQUE — TAPIS {tapis_num_actif}</div>", unsafe_allow_html=True)
                    st.caption(f"🏆 {nom_competition}")
                with col_k_h2:
                    st.caption(f"Combats : **{nb_termines_tapis} / {nb_total_tapis}** ({pct_tapis}%)")
                    st.progress(pct_tapis / 100)

            if not st.session_state.get("vue_scoreboard_active"):
                st.markdown("---")

            # Trouver l'index du match actif
            idx_match_actif = 0
            cle_sess_idx = f"idx_actif_tapis_{tapis_num_actif}"
            combo_key = f"combo_combats_{tapis_num_actif}"
        
            # Déterminer le premier combat non terminé (le combat en cours "live")
            idx_en_cours = -1
            for i_m, m_chk in enumerate(matchs_du_tapis):
                if m_chk.get("statut") != "Terminé":
                    idx_en_cours = i_m
                    break

            etat_sync_tapis = recuperer_etat_tapis(tapis_num_actif)
            idx_sync_tapis = etat_sync_tapis.get("active_idx")

            if st.session_state.get("vue_scoreboard_active"):
                # Scoreboard TV : suit en priorité l'arbitre sur la table de marque (pur affichage, sans navigation ni boutons)
                if idx_sync_tapis is not None and 0 <= idx_sync_tapis < nb_total_tapis:
                    idx_match_actif = idx_sync_tapis
                elif cle_sess_idx in st.session_state:
                    idx_match_actif = min(max(st.session_state[cle_sess_idx], 0), max(nb_total_tapis - 1, 0))
                else:
                    idx_match_actif = idx_en_cours if idx_en_cours != -1 else 0
                st.session_state[cle_sess_idx] = idx_match_actif
                m_actuel = matchs_du_tapis[idx_match_actif] if (matchs_du_tapis and 0 <= idx_match_actif < len(matchs_du_tapis)) else None
                afficher_scoreboard_tapis(tapis_num_actif, m_actuel, matchs_du_tapis, idx_match_actif, c_sess)
                st.stop()
            else:
                # Table de marque de l'arbitre
                if cle_sess_idx in st.session_state:
                    idx_match_actif = min(max(st.session_state[cle_sess_idx], 0), max(nb_total_tapis - 1, 0))
                elif idx_sync_tapis is not None and 0 <= idx_sync_tapis < nb_total_tapis:
                    idx_match_actif = idx_sync_tapis
                    st.session_state[cle_sess_idx] = idx_match_actif
                else:
                    idx_match_actif = idx_en_cours if idx_en_cours != -1 else 0
                    st.session_state[cle_sess_idx] = idx_match_actif
                
                # Publier le combat actif pour synchroniser le Scoreboard TV
                if matchs_du_tapis and 0 <= idx_match_actif < len(matchs_du_tapis):
                    m_cur_chk = matchs_du_tapis[idx_match_actif]
                    definir_combat_actif_tapis(tapis_num_actif, idx_match_actif, m_cur_chk.get("id") or f"T{tapis_num_actif}_M{m_cur_chk.get('match_num', idx_match_actif)}")

            options_combats = [
                f"Combat n°{m.get('match_num')} ({m.get('heure')}) : {m.get('lutteur_rouge')} vs {m.get('lutteur_bleu')} [{'✅' if m.get('statut')=='Terminé' else '⏳'}]"
                for m in matchs_du_tapis
            ]

            # Callback déclenché lors du changement dans le selectbox
            def _on_change_combo_combat():
                sel_val = st.session_state.get(combo_key)
                if sel_val in options_combats:
                    nouv_idx = options_combats.index(sel_val)
                    st.session_state[cle_sess_idx] = nouv_idx
                    if matchs_du_tapis and 0 <= nouv_idx < len(matchs_du_tapis):
                        m_tgt = matchs_du_tapis[nouv_idx]
                        definir_combat_actif_tapis(tapis_num_actif, nouv_idx, m_tgt.get("id") or f"T{tapis_num_actif}_M{m_tgt.get('match_num', nouv_idx)}")

            col_nav1, col_nav2, col_nav3 = st.columns([1.5, 4, 1.5])
            with col_nav1:
                if st.button("⏮️ Précédent", disabled=(idx_match_actif <= 0), use_container_width=True, key=f"btn_prev_{tapis_num_actif}"):
                    nouv_idx = max(0, idx_match_actif - 1)
                    st.session_state[cle_sess_idx] = nouv_idx
                    if matchs_du_tapis and 0 <= nouv_idx < len(matchs_du_tapis):
                        m_tgt = matchs_du_tapis[nouv_idx]
                        definir_combat_actif_tapis(tapis_num_actif, nouv_idx, m_tgt.get("id") or f"T{tapis_num_actif}_M{m_tgt.get('match_num', nouv_idx)}")
                    st.session_state.pop(combo_key, None)
                    st.rerun()
            with col_nav2:
                st.selectbox(
                    "Choisir un combat à afficher :", 
                    options_combats, 
                    index=idx_match_actif, 
                    label_visibility="collapsed",
                    key=combo_key,
                    on_change=_on_change_combo_combat
                )
            with col_nav3:
                if st.button("Suivant ⏭️", disabled=(idx_match_actif >= nb_total_tapis - 1), use_container_width=True, key=f"btn_next_{tapis_num_actif}"):
                    nouv_idx = min(nb_total_tapis - 1, idx_match_actif + 1)
                    st.session_state[cle_sess_idx] = nouv_idx
                    if matchs_du_tapis and 0 <= nouv_idx < len(matchs_du_tapis):
                        m_tgt = matchs_du_tapis[nouv_idx]
                        definir_combat_actif_tapis(tapis_num_actif, nouv_idx, m_tgt.get("id") or f"T{tapis_num_actif}_M{m_tgt.get('match_num', nouv_idx)}")
                    st.session_state.pop(combo_key, None)
                    st.rerun()

            # LE COMBAT ACTIF (TABLE DE MARQUE)
            m_actuel = matchs_du_tapis[idx_match_actif]
            cat_m = m_actuel.get("categorie", "")
            age_m = extraire_age_de_texte(cat_m)
            is_u13 = "U13" in age_m.upper()
            est_deja_termine = (m_actuel.get("statut") == "Terminé")

            # Alerte Mode Retour en arrière / Correction si on consulte un combat déjà terminé
            if est_deja_termine:
                col_back_info, col_back_act = st.columns([3, 2])
                with col_back_info:
                    b_vic_act = formater_badge_victoire_html(m_actuel.get('type_victoire', ''), compact=False)
                    v_nom_act = m_actuel.get('lutteur_rouge') if m_actuel.get('vainqueur') == "Rouge" else (m_actuel.get('lutteur_bleu') if m_actuel.get('vainqueur') == "Bleu" else m_actuel.get('vainqueur', ''))
                    v_col_act = "#dc2626" if m_actuel.get('vainqueur') == "Rouge" else "#2563eb"
                    st.markdown(f"<div style='background:#fef3c7; border:2px solid #f59e0b; border-radius:10px; padding:10px 14px; color:#92400e; font-size:14px; font-weight:700;'>⏮️ <b>Mode Consultation / Correction :</b> Combat n°{m_actuel.get('match_num')} — Vainqueur : <span style='color:{v_col_act}; font-weight:900;'>{v_nom_act}</span> {b_vic_act}</div>", unsafe_allow_html=True)
                with col_back_act:
                    if idx_en_cours != -1 and idx_en_cours != idx_match_actif:
                        num_live = matchs_du_tapis[idx_en_cours].get('match_num')
                        if st.button(f"▶️ Revenir au combat en cours (N°{num_live})", key=f"btn_go_live_{tapis_num_actif}", use_container_width=True, type="secondary"):
                            st.session_state[cle_sess_idx] = idx_en_cours
                            if matchs_du_tapis and 0 <= idx_en_cours < len(matchs_du_tapis):
                                m_tgt = matchs_du_tapis[idx_en_cours]
                                definir_combat_actif_tapis(tapis_num_actif, idx_en_cours, m_tgt.get("id") or f"T{tapis_num_actif}_M{m_tgt.get('match_num', idx_en_cours)}")
                            st.session_state.pop(combo_key, None)
                            st.rerun()
        
            statut_badge = "🟢 COMBAT TERMINÉ (MODIFIABLE)" if est_deja_termine else "🟡 COMBAT EN COURS"
            st.markdown(f"<div style='text-align: center; margin-bottom: 8px;'><span style='background-color: {'#2e7d32' if est_deja_termine else '#f57c00'}; color: white; padding: 4px 14px; border-radius: 12px; font-weight: bold; font-size: 13px;'>{statut_badge}</span></div>", unsafe_allow_html=True)
        
            st.markdown(
                f"<h3 style='text-align: center; margin-top: 0;'>🤼 MATCH N° {m_actuel.get('match_num')} — 🕘 {m_actuel.get('heure')}<br>"
                f"<span style='font-size: 18px; color: var(--text-color, #64748b); opacity: 0.85;'>Catégorie : <b>{cat_m}</b> | Tour : <b>{m_actuel.get('tour')}</b></span></h3>",
                unsafe_allow_html=True
            )

            # Injection CSS pour Option 2 : Tableau d'affichage / Scoreboard électronique officiel de lutte
            st.markdown('''
            <style>
            /* ========================================================================= */
            /* --- OPTION 2 : TABLEAU D'AFFICHAGE & BOÎTIER DE SCORE OFFICIEL FFLDA ---- */
            /* ========================================================================= */

            /* --- BOUTONS VAINQUEUR STRICTEMENT SCOPÉS AU FORMULAIRE DE MATCH --- */
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] {
                display: flex !important;
                flex-direction: row !important;
                gap: 14px !important;
                width: 100% !important;
            }
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label {
                flex: 1 1 0 !important;
                min-height: 56px !important;
                padding: 10px 14px !important;
                border-radius: 12px !important;
                border: 2px solid #cbd5e1 !important;
                display: flex !important;
                align-items: center !important;
                justify-content: center !important;
                cursor: pointer !important;
                transition: all 0.2s ease-in-out !important;
                box-shadow: 0 2px 6px rgba(0,0,0,0.04) !important;
                text-align: center !important;
            }

            /* Bouton Vainqueur Rouge Non sélectionné */
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:first-child {
                background-color: #fff1f2 !important;
                border-color: #f87171 !important;
            }
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:first-child p,
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:first-child span,
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:first-child div {
                color: #991b1b !important;
                font-weight: 800 !important;
                font-size: 16px !important;
                letter-spacing: 0.5px !important;
            }
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:first-child:hover {
                border-color: #dc2626 !important;
                background-color: #fee2e2 !important;
            }

            /* Bouton Vainqueur Rouge SÉLECTIONNÉ */
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:first-child:has(input:checked) {
                background-color: #dc2626 !important;
                border-color: #991b1b !important;
                box-shadow: 0 4px 14px rgba(220,38,38,0.35) !important;
                transform: translateY(-1px) !important;
            }
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:first-child:has(input:checked) p,
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:first-child:has(input:checked) span,
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:first-child:has(input:checked) div {
                color: #ffffff !important;
                font-weight: 800 !important;
                font-size: 16px !important;
            }

            /* Bouton Vainqueur Bleu Non sélectionné */
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:last-child {
                background-color: #eff6ff !important;
                border-color: #60a5fa !important;
            }
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:last-child p,
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:last-child span,
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:last-child div {
                color: #1e40af !important;
                font-weight: 800 !important;
                font-size: 16px !important;
                letter-spacing: 0.5px !important;
            }
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:last-child:hover {
                border-color: #2563eb !important;
                background-color: #dbeafe !important;
            }

            /* Bouton Vainqueur Bleu SÉLECTIONNÉ */
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:last-child:has(input:checked) {
                background-color: #0055A4 !important;
                border-color: #003d7a !important;
                box-shadow: 0 4px 14px rgba(0,85,164,0.35) !important;
                transform: translateY(-1px) !important;
            }
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:last-child:has(input:checked) p,
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:last-child:has(input:checked) span,
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label:last-child:has(input:checked) div {
                color: #ffffff !important;
                font-weight: 800 !important;
                font-size: 16px !important;
            }

            /* Masquer la puce radio circulaire uniquement sur le choix du vainqueur */
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label input,
            div[data-testid="stVerticalBlock"]:has(#radio_vainqueur_marker) div[data-testid="stRadio"] div[role="radiogroup"] label > div:first-child {
                display: none !important;
            }

            /* ========================================================================= */
            /* --- TOUCHES DE POINTS DU BOÎTIER (SCOPÉES AUX TOUCHES BTN_R ET BTN_B) --- */
            /* ========================================================================= */
            button[key*="btn_r1_"], button[key*="btn_r2_"], button[key*="btn_r4_"],
            button[key*="btn_r5_"], button[key*="btn_rav_"], button[key*="btn_rdel_"],
            button[key*="btn_rrst_"],
            button[key*="btn_b1_"], button[key*="btn_b2_"], button[key*="btn_b4_"],
            button[key*="btn_b5_"], button[key*="btn_bav_"], button[key*="btn_bdel_"],
            button[key*="btn_brst_"] {
                min-height: 44px !important;
                height: 44px !important;
                font-size: 14px !important;
                font-weight: 800 !important;
                border-radius: 10px !important;
                background-color: #ffffff !important;
                padding: 0 4px !important;
                margin: 1px 0 !important;
                transition: all 0.15s ease !important;
                box-shadow: 0 2px 6px rgba(0,0,0,0.04) !important;
                white-space: nowrap !important;
                overflow: visible !important;
            }

            /* Touches Coin Rouge : Bordure rouge et texte rouge */
            button[key*="btn_r1_"], button[key*="btn_r2_"], button[key*="btn_r4_"],
            button[key*="btn_r5_"], button[key*="btn_rav_"], button[key*="btn_rdel_"],
            button[key*="btn_rrst_"] {
                border: 1.5px solid #ef4444 !important;
                color: #b91c1c !important;
                -webkit-text-fill-color: #b91c1c !important;
            }
            button[key*="btn_r1_"] *, button[key*="btn_r2_"] *, button[key*="btn_r4_"] *,
            button[key*="btn_r5_"] *, button[key*="btn_rav_"] *, button[key*="btn_rdel_"] *,
            button[key*="btn_rrst_"] * {
                color: #b91c1c !important;
                -webkit-text-fill-color: #b91c1c !important;
            }
            button[key*="btn_r1_"]:hover, button[key*="btn_r2_"]:hover, button[key*="btn_r4_"]:hover,
            button[key*="btn_r5_"]:hover, button[key*="btn_rav_"]:hover, button[key*="btn_rdel_"]:hover,
            button[key*="btn_rrst_"]:hover {
                background-color: #fee2e2 !important;
                border-color: #dc2626 !important;
                color: #991b1b !important;
                -webkit-text-fill-color: #991b1b !important;
            }
            button[key*="btn_r1_"]:hover *, button[key*="btn_r2_"]:hover *, button[key*="btn_r4_"]:hover *,
            button[key*="btn_r5_"]:hover *, button[key*="btn_rav_"]:hover *, button[key*="btn_rdel_"]:hover *,
            button[key*="btn_rrst_"]:hover * {
                color: #991b1b !important;
                -webkit-text-fill-color: #991b1b !important;
            }

            /* Touches Coin Bleu : Bordure bleu et texte bleu */
            button[key*="btn_b1_"], button[key*="btn_b2_"], button[key*="btn_b4_"],
            button[key*="btn_b5_"], button[key*="btn_bav_"], button[key*="btn_bdel_"],
            button[key*="btn_brst_"] {
                border: 1.5px solid #0055A4 !important;
                color: #0055A4 !important;
                -webkit-text-fill-color: #0055A4 !important;
            }
            button[key*="btn_b1_"] *, button[key*="btn_b2_"] *, button[key*="btn_b4_"] *,
            button[key*="btn_b5_"] *, button[key*="btn_bav_"] *, button[key*="btn_bdel_"] *,
            button[key*="btn_brst_"] * {
                color: #0055A4 !important;
                -webkit-text-fill-color: #0055A4 !important;
            }
            button[key*="btn_b1_"]:hover, button[key*="btn_b2_"]:hover, button[key*="btn_b4_"]:hover,
            button[key*="btn_b5_"]:hover, button[key*="btn_bav_"]:hover, button[key*="btn_bdel_"]:hover,
            button[key*="btn_brst_"]:hover {
                background-color: #eff6ff !important;
                border-color: #1d4ed8 !important;
                color: #1e40af !important;
                -webkit-text-fill-color: #1e40af !important;
            }
            button[key*="btn_b1_"]:hover *, button[key*="btn_b2_"]:hover *, button[key*="btn_b4_"]:hover *,
            button[key*="btn_b5_"]:hover *, button[key*="btn_bav_"]:hover *, button[key*="btn_bdel_"]:hover *,
            button[key*="btn_brst_"]:hover * {
                color: #1e40af !important;
                -webkit-text-fill-color: #1e40af !important;
            }

            /* Chronomètre table de marque : Charte FFLDA cohérente */
            div[data-testid="stVerticalBlock"]:has(#chrono_barre_marker) button {
                min-height: 42px !important;
                font-weight: 800 !important;
                border-radius: 10px !important;
            }
            div[data-testid="stVerticalBlock"]:has(#chrono_barre_marker) button[kind="secondary"],
            div[data-testid="stVerticalBlock"]:has(#chrono_barre_marker) button[data-testid="baseButton-secondary"] {
                background-color: #ffffff !important;
                background: #ffffff !important;
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                border: 1.5px solid #cbd5e1 !important;
            }
            div[data-testid="stVerticalBlock"]:has(#chrono_barre_marker) button[kind="secondary"] *,
            div[data-testid="stVerticalBlock"]:has(#chrono_barre_marker) button[data-testid="baseButton-secondary"] * {
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
            }
            div[data-testid="stVerticalBlock"]:has(#chrono_barre_marker) button[kind="secondary"]:hover,
            div[data-testid="stVerticalBlock"]:has(#chrono_barre_marker) button[data-testid="baseButton-secondary"]:hover {
                background-color: #f1f5f9 !important;
                background: #f1f5f9 !important;
                border-color: #0055A4 !important;
                color: #0055A4 !important;
                -webkit-text-fill-color: #0055A4 !important;
            }
            div[data-testid="stVerticalBlock"]:has(#chrono_barre_marker) button[kind="secondary"]:hover *,
            div[data-testid="stVerticalBlock"]:has(#chrono_barre_marker) button[data-testid="baseButton-secondary"]:hover * {
                color: #0055A4 !important;
                -webkit-text-fill-color: #0055A4 !important;
            }
            div[data-testid="stVerticalBlock"]:has(#chrono_barre_marker) button[kind="primary"],
            div[data-testid="stVerticalBlock"]:has(#chrono_barre_marker) button[data-testid="baseButton-primary"] {
                background: linear-gradient(135deg, #0055A4 0%, #1d4ed8 100%) !important;
                background-color: #0055A4 !important;
                color: #ffffff !important;
                -webkit-text-fill-color: #ffffff !important;
                border: none !important;
            }
            div[data-testid="stVerticalBlock"]:has(#chrono_barre_marker) button[kind="primary"] * {
                color: #ffffff !important;
                -webkit-text-fill-color: #ffffff !important;
            }
            </style>
            ''', unsafe_allow_html=True)

            # GESTION DES SCORES ET DU DÉPARTAGE OFFICIEL UWW (TOUCHES TACTILES & CHRONOLOGIE)
            m_id = m_actuel.get('id') or f"T{tapis_num_actif}_M{m_actuel.get('match_num', idx_match_actif)}"
            nom_lutteur_r = m_actuel.get('lutteur_rouge', 'Rouge')
            club_lutteur_r = m_actuel.get('club_rouge', '')
            nom_lutteur_b = m_actuel.get('lutteur_bleu', 'Bleu')
            club_lutteur_b = m_actuel.get('club_bleu', '')

            cle_act_r = f"actions_r_{m_id}"
            cle_act_b = f"actions_b_{m_id}"
            cle_caut_r = f"cautions_r_{m_id}"
            cle_caut_b = f"cautions_b_{m_id}"

            init_sc_r = int(m_actuel.get("score_rouge", 0) or 0)
            init_sc_b = int(m_actuel.get("score_bleu", 0) or 0)

            sh_act_tb = recuperer_actions_match_sync(m_id)
            if sh_act_tb:
                if "actions_r" in sh_act_tb and cle_act_r not in st.session_state:
                    st.session_state[cle_act_r] = sh_act_tb["actions_r"]
                if "actions_b" in sh_act_tb and cle_act_b not in st.session_state:
                    st.session_state[cle_act_b] = sh_act_tb["actions_b"]
                if "cautions_r" in sh_act_tb and cle_caut_r not in st.session_state:
                    st.session_state[cle_caut_r] = sh_act_tb["cautions_r"]
                if "cautions_b" in sh_act_tb and cle_caut_b not in st.session_state:
                    st.session_state[cle_caut_b] = sh_act_tb["cautions_b"]

            if cle_act_r not in st.session_state:
                st.session_state[cle_act_r] = charger_actions_match(m_id, "Rouge", init_sc_r, m_actuel, c_sess)
            if cle_act_b not in st.session_state:
                st.session_state[cle_act_b] = charger_actions_match(m_id, "Bleu", init_sc_b, m_actuel, c_sess)

            if cle_caut_r not in st.session_state:
                st.session_state[cle_caut_r] = int(m_actuel.get("cautions_rouge", 0) or 0)

            if cle_caut_b not in st.session_state:
                st.session_state[cle_caut_b] = int(m_actuel.get("cautions_bleu", 0) or 0)

            # --- BARRE DE CHRONOMÈTRE TACTILE & CONTRÔLE SCOREBOARD (TABLE DE MARQUE) ---
            @fragment_compat(run_every="1s")
            def _fragment_chrono_barre_arbitre():
                sec_r, en_cours = calculer_temps_restant_chrono(m_id, cat_m, tapis_num_actif)
                txt_chrono = formater_chrono_mm_ss(sec_r)
                per = st.session_state.get(f"chrono_per_{m_id}", st.session_state.get(f"chrono_per_tapis_{tapis_num_actif}", 1))
                pause30 = st.session_state.get(f"chrono_pause30_{m_id}", False)
                
                if pause30:
                    col_c_bg = "#f59e0b"
                    badge_lbl = "⏸️ PAUSE (30s)"
                    per_lbl = "PAUSE 30s"
                else:
                    col_c_bg = "#ef4444" if sec_r == 0 else ("#16a34a" if en_cours else "#d97706")
                    badge_lbl = "● EN COMBAT" if en_cours else ("🔔 FIN DU TEMPS" if sec_r == 0 else "⏸️ PAUSE")
                    per_lbl = f"PÉRIODE {per}"
                
                c_head_chr, c_btn_bz, c_btn_scb = st.columns([2.5, 1.1, 1.1], gap="small")
                with c_head_chr:
                    badge_per_style = "background: #78350f; color: #fde68a; border: 1px solid #f59e0b;" if pause30 else "background: #1e293b; color: #94a3b8;"
                    st.markdown(
                        f"<div style='background: #0f172a; border: 2px solid {col_c_bg}; border-radius: 12px; padding: 6px 12px; margin-bottom: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.25); display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 8px; box-sizing: border-box; overflow: hidden;'>"
                        f"<div style='display: flex; align-items: center; gap: 8px; flex-wrap: wrap;'>"
                        f"<span style='font-family: monospace; font-size: clamp(18px, 4vw, 26px); font-weight: 900; color: {col_c_bg}; letter-spacing: 1px; text-shadow: 0 0 10px {col_c_bg}; white-space: nowrap;'>⏱️ {txt_chrono}</span>"
                        f"<span style='{badge_per_style} font-size: clamp(10px, 1.2vw, 12px); font-weight: 800; padding: 2px 7px; border-radius: 6px; white-space: nowrap;'>{per_lbl}</span>"
                        f"<span style='color: {col_c_bg}; font-size: clamp(10px, 1.1vw, 11px); font-weight: 800; letter-spacing: 0.5px; white-space: nowrap;'>{badge_lbl}</span>"
                        f"</div>"
                        f"<span style='font-size: 11px; color: #94a3b8; white-space: nowrap;'>Tapis {tapis_num_actif}</span>"
                        f"</div>",
                        unsafe_allow_html=True
                    )
                with c_btn_bz:
                    st.button("🔊 Test Buzzer", key=f"btn_tb_bz_{m_id}", on_click=declencher_buzzer_manuel, args=(m_id, cat_m, tapis_num_actif), use_container_width=True, help="Déclencher le signal sonore / klaxon sur la table et le Scoreboard TV")
                with c_btn_scb:
                    if st.button("📺 Scoreboard TV", key=f"btn_tb_scb_{m_id}", use_container_width=True, help="Afficher le grand Scoreboard TV haute visibilité pour ce tapis"):
                        st.session_state["vue_scoreboard_active"] = True
                        st.rerun()

                # Signal sonore local Table de Marque (Web Audio API)
                etat_chr_tb = recuperer_chrono_sync(m_id, tapis_num_actif, cat_m)
                b_evt_tb = etat_chr_tb.get("buzzer_event", 0.0)
                last_b_tb = st.session_state.get(f"last_buzzer_table_{m_id}", 0.0)
                if b_evt_tb > 0 and (pytime.time() - b_evt_tb < 3.0) and (b_evt_tb != last_b_tb):
                    st.session_state[f"last_buzzer_table_{m_id}"] = b_evt_tb
                    st.session_state[f"play_local_buzzer_{m_id}"] = pytime.time()

                if pytime.time() - st.session_state.get(f"play_local_buzzer_{m_id}", 0.0) < 2.5:
                    components.html("""
                    <script>
                    (function() {
                        try {
                            var pWin = window.parent || window;
                            var AudioCtx = pWin.AudioContext || pWin.webkitAudioContext;
                            if (!AudioCtx) return;
                            var ctx = new AudioCtx();
                            if (ctx.state === 'suspended') { ctx.resume(); }
                            var now = ctx.currentTime;
                            [0, 0.35, 0.70].forEach(function(offset) {
                                var osc = ctx.createOscillator();
                                var gain = ctx.createGain();
                                osc.type = 'sawtooth';
                                osc.frequency.setValueAtTime(500, now + offset);
                                osc.frequency.exponentialRampToValueAtTime(240, now + offset + 0.30);
                                gain.gain.setValueAtTime(0.7, now + offset);
                                gain.gain.exponentialRampToValueAtTime(0.01, now + offset + 0.30);
                                osc.connect(gain);
                                gain.connect(ctx.destination);
                                osc.start(now + offset);
                                osc.stop(now + offset + 0.30);
                            });
                        } catch(e) {}
                    })();
                    </script>
                    """, height=0)

                st.markdown('<div id="chrono_barre_marker"></div>', unsafe_allow_html=True)
                c_ch1, c_ch2, c_ch3, c_ch4, c_ch5 = st.columns([1.5, 1.2, 0.9, 0.9, 1.3], gap="small")
                with c_ch1:
                    lbl_tgl = "⏸️ Pause" if en_cours else "▶️ Démarrer"
                    st.button(lbl_tgl, key=f"btn_tb_tgl_{m_id}", on_click=toggle_chrono_match, args=(m_id, cat_m, tapis_num_actif), use_container_width=True, type="primary" if en_cours else "secondary")
                with c_ch2:
                    st.button("🔄 Reset", key=f"btn_tb_rst_{m_id}", on_click=reset_chrono_match, args=(m_id, cat_m, tapis_num_actif), use_container_width=True, help="Remettre à zéro le chrono (durée officielle P1)")
                with c_ch3:
                    st.button("-10s", key=f"btn_tb_m10_{m_id}", on_click=ajuster_chrono_match, args=(m_id, -10, cat_m, tapis_num_actif), use_container_width=True)
                with c_ch4:
                    st.button("+10s", key=f"btn_tb_p10_{m_id}", on_click=ajuster_chrono_match, args=(m_id, 10, cat_m, tapis_num_actif), use_container_width=True)
                with c_ch5:
                    lbl_btn_per = "▶️ Lancer P2" if pause30 else f"P{2 if per == 1 else 1} 🔁"
                    help_btn_per = "Passer directement à la Période 2 sans attendre la fin des 30s" if pause30 else "Basculer vers la période suivante"
                    st.button(lbl_btn_per, key=f"btn_tb_per_{m_id}", on_click=changer_periode_chrono, args=(m_id, tapis_num_actif, cat_m), use_container_width=True, help=help_btn_per)

            _fragment_chrono_barre_arbitre()

            @fragment_compat()
            def _afficher_boitier_score_fragment():
                actions_r = st.session_state[cle_act_r]
                actions_b = st.session_state[cle_act_b]
                cautions_r = st.session_state[cle_caut_r]
                cautions_b = st.session_state[cle_caut_b]

                score_r = sum(a.get("val", 0) for a in actions_r)
                score_b = sum(a.get("val", 0) for a in actions_b)

                # Évaluation officielle du départage UWW
                meneur, motif_departage, crit_num = evaluer_departage_uww(actions_r, actions_b, cautions_r, cautions_b)

                # 5 colonnes : Ratio 2.0 pour la carte lutteur / 3.0 pour le boîtier de score LED + touches
                st.markdown('<div id="boitier_score_marker"></div>', unsafe_allow_html=True)
                col_cr_nom, col_cr_sc, col_vs, col_cb_nom, col_cb_sc = st.columns([2.0, 3.0, 0.35, 2.0, 3.0], gap="small")
        
                with col_cr_nom:
                    st.markdown(
                        f"<div style='background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%); color: white; padding: 8px 12px; border-radius: 14px; height: 100px; display: flex; flex-direction: column; justify-content: center; box-shadow: 0 4px 14px rgba(220,38,38,0.28); border: 2px solid #dc2626; box-sizing: border-box; overflow: hidden;'>"
                        f"<div style='font-size: 10px; text-transform: uppercase; letter-spacing: 1px; font-weight: 900; opacity: 0.95;'>🔴 COIN ROUGE</div>"
                        f"<div style='font-size: clamp(13px, 1.8vw, 19px); font-weight: 900; line-height: 1.15; margin: 2px 0 1px 0; overflow: hidden; text-overflow: ellipsis; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; word-break: break-word;' title='{nom_lutteur_r}'>{nom_lutteur_r}</div>"
                        f"<div style='font-size: 11px; font-weight: 600; opacity: 0.9; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;'>{club_lutteur_r}</div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )

                with col_cr_sc:
                    c_r_num, c_r_pad = st.columns([1, 1.8], gap="small")
                    with c_r_num:
                        st.markdown(
                            f"<div style='height: 100px; background-color: #0b0f19; border: 3px solid #dc2626; border-radius: 14px; display: flex; align-items: center; justify-content: center; box-shadow: 0 4px 14px rgba(220,38,38,0.28); box-sizing: border-box;'>"
                            f"<span style='font-size: clamp(34px, 4.5vw, 50px); font-weight: 900; color: #ff4d4d; text-shadow: 0 0 14px rgba(255,77,77,0.75); font-family: monospace;'>{score_r}</span>"
                            f"</div>",
                            unsafe_allow_html=True
                        )
                    with c_r_pad:
                        # Ligne 1 : Les 4 touches de points
                        br1, br2, br4, br5 = st.columns(4, gap="small")
                        with br1:
                            st.button("+1", key=f"btn_r1_{m_id}_{idx_match_actif}", on_click=ajouter_action_lutte, args=(m_id, "Rouge", 1, "tech", "1"), use_container_width=True)
                        with br2:
                            st.button("+2", key=f"btn_r2_{m_id}_{idx_match_actif}", on_click=ajouter_action_lutte, args=(m_id, "Rouge", 2, "tech", "2"), use_container_width=True)
                        with br4:
                            st.button("+4", key=f"btn_r4_{m_id}_{idx_match_actif}", on_click=ajouter_action_lutte, args=(m_id, "Rouge", 4, "tech", "4"), use_container_width=True)
                        with br5:
                            st.button("+5", key=f"btn_r5_{m_id}_{idx_match_actif}", on_click=ajouter_action_lutte, args=(m_id, "Rouge", 5, "tech", "5"), use_container_width=True)
                        
                        # Ligne 2 : Avertissement, Annulation de la dernière prise & Remise à 0
                        bav, bdel, brst = st.columns([1.2, 1, 1], gap="small")
                        with bav:
                            st.button("⚠️ Avert.", key=f"btn_rav_{m_id}_{idx_match_actif}", on_click=ajouter_avertissement_lutte, args=(m_id, "Rouge"), use_container_width=True, help="Enregistrer un avertissement contre Rouge")
                        with bdel:
                            st.button("⌫", key=f"btn_rdel_{m_id}_{idx_match_actif}", on_click=annuler_action_lutte, args=(m_id, "Rouge"), use_container_width=True, help="Annuler la dernière action Rouge de l'addition")
                        with brst:
                            st.button("🗑️", key=f"btn_rrst_{m_id}_{idx_match_actif}", on_click=reinitialiser_score_lutteur, args=(m_id, "Rouge"), use_container_width=True, help="Effacer tous les points de Rouge (Remise à 0)")
                
                with col_vs:
                    st.markdown(
                        "<div style='display: flex; align-items: center; justify-content: center; height: 100px;'>"
                        "<div style='width: 32px; height: 32px; line-height: 28px; border-radius: 50%; background: #0f172a; border: 2px solid #475569; color: #94a3b8; font-weight: 900; font-size: 11px; text-align: center; box-shadow: 0 4px 10px rgba(0,0,0,0.35);'>VS</div>"
                        "</div>", 
                        unsafe_allow_html=True
                    )
                
                with col_cb_nom:
                    st.markdown(
                        f"<div style='background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%); color: white; padding: 8px 12px; border-radius: 14px; height: 100px; display: flex; flex-direction: column; justify-content: center; box-shadow: 0 4px 14px rgba(37,99,235,0.28); border: 2px solid #2563eb; box-sizing: border-box; overflow: hidden;'>"
                        f"<div style='font-size: 10px; text-transform: uppercase; letter-spacing: 1px; font-weight: 900; opacity: 0.95;'>🔵 COIN BLEU</div>"
                        f"<div style='font-size: clamp(13px, 1.8vw, 19px); font-weight: 900; line-height: 1.15; margin: 2px 0 1px 0; overflow: hidden; text-overflow: ellipsis; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; word-break: break-word;' title='{nom_lutteur_b}'>{nom_lutteur_b}</div>"
                        f"<div style='font-size: 11px; font-weight: 600; opacity: 0.9; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;'>{club_lutteur_b}</div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )

                with col_cb_sc:
                    c_b_num, c_b_pad = st.columns([1, 1.8], gap="small")
                    with c_b_num:
                        st.markdown(
                            f"<div style='height: 100px; background-color: #0b0f19; border: 3px solid #2563eb; border-radius: 14px; display: flex; align-items: center; justify-content: center; box-shadow: 0 4px 14px rgba(37,99,235,0.28); box-sizing: border-box;'>"
                            f"<span style='font-size: clamp(34px, 4.5vw, 50px); font-weight: 900; color: #38bdf8; text-shadow: 0 0 14px rgba(56,189,248,0.75); font-family: monospace;'>{score_b}</span>"
                            f"</div>",
                            unsafe_allow_html=True
                        )
                    with c_b_pad:
                        # Ligne 1 : Les 4 touches de points
                        bb1, bb2, bb4, bb5 = st.columns(4, gap="small")
                        with bb1:
                            st.button("+1", key=f"btn_b1_{m_id}_{idx_match_actif}", on_click=ajouter_action_lutte, args=(m_id, "Bleu", 1, "tech", "1"), use_container_width=True)
                        with bb2:
                            st.button("+2", key=f"btn_b2_{m_id}_{idx_match_actif}", on_click=ajouter_action_lutte, args=(m_id, "Bleu", 2, "tech", "2"), use_container_width=True)
                        with bb4:
                            st.button("+4", key=f"btn_b4_{m_id}_{idx_match_actif}", on_click=ajouter_action_lutte, args=(m_id, "Bleu", 4, "tech", "4"), use_container_width=True)
                        with bb5:
                            st.button("+5", key=f"btn_b5_{m_id}_{idx_match_actif}", on_click=ajouter_action_lutte, args=(m_id, "Bleu", 5, "tech", "5"), use_container_width=True)
                        
                        # Ligne 2 : Avertissement, Annulation de la dernière prise & Remise à 0
                        bav, bdel, brst = st.columns([1.2, 1, 1], gap="small")
                        with bav:
                            st.button("⚠️ Avert.", key=f"btn_bav_{m_id}_{idx_match_actif}", on_click=ajouter_avertissement_lutte, args=(m_id, "Bleu"), use_container_width=True, help="Enregistrer un avertissement contre Bleu")
                        with bdel:
                            st.button("⌫", key=f"btn_bdel_{m_id}_{idx_match_actif}", on_click=annuler_action_lutte, args=(m_id, "Bleu"), use_container_width=True, help="Annuler la dernière action Bleu de l'addition")
                        with brst:
                            st.button("🗑️", key=f"btn_brst_{m_id}_{idx_match_actif}", on_click=reinitialiser_score_lutteur, args=(m_id, "Bleu"), use_container_width=True, help="Effacer tous les points de Bleu (Remise à 0)")

                # Bandeau chronologique des prises sous les scores
                col_rib_r, col_rib_b = st.columns(2, gap="medium")
                with col_rib_r:
                    pastilles_html_r = ""
                    for a in actions_r:
                        if a.get('type') == 'avert':
                            pastilles_html_r += "<span style='display: inline-block; background-color: #854d0e; color: #fef08a; font-weight: 800; font-size: 13px; padding: 2px 7px; border-radius: 6px; margin: 1px 2px; border: 1px solid #eab308;'>⚠️</span>"
                        else:
                            pastilles_html_r += f"<span style='display: inline-block; background-color: #7f1d1d; color: #fecaca; font-weight: 800; font-size: 13px; padding: 2px 7px; border-radius: 6px; margin: 1px 2px; border: 1px solid #ef4444;'>+{a.get('val')}</span>"
                    if not pastilles_html_r:
                        pastilles_html_r = "<span style='color: #64748b; font-style: italic; font-size: 12px;'>Aucun point</span>"
                    
                    cautions_icons_r = (" ⚠️" * cautions_r) if cautions_r > 0 else ""
                    st.markdown(
                        f"<div style='background-color: #0b0f19; border: 1px solid #334155; border-left: 4px solid #dc2626; border-radius: 8px; padding: 6px 12px; margin-top: 8px; min-height: 40px; display: flex; align-items: center; justify-content: space-between;'>"
                        f"<span style='font-size: 12px; font-weight: 800; color: #f87171;'>🔴 PRISES ROUGE{cautions_icons_r} :</span>"
                        f"<div>{pastilles_html_r}</div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )

                with col_rib_b:
                    pastilles_html_b = ""
                    for a in actions_b:
                        if a.get('type') == 'avert':
                            pastilles_html_b += "<span style='display: inline-block; background-color: #854d0e; color: #fef08a; font-weight: 800; font-size: 13px; padding: 2px 7px; border-radius: 6px; margin: 1px 2px; border: 1px solid #eab308;'>⚠️</span>"
                        else:
                            pastilles_html_b += f"<span style='display: inline-block; background-color: #1e3a8a; color: #bfdbfe; font-weight: 800; font-size: 13px; padding: 2px 7px; border-radius: 6px; margin: 1px 2px; border: 1px solid #3b82f6;'>+{a.get('val')}</span>"
                    if not pastilles_html_b:
                        pastilles_html_b = "<span style='color: #64748b; font-style: italic; font-size: 12px;'>Aucun point</span>"

                    cautions_icons_b = (" ⚠️" * cautions_b) if cautions_b > 0 else ""
                    st.markdown(
                        f"<div style='background-color: #0b0f19; border: 1px solid #334155; border-left: 4px solid #2563eb; border-radius: 8px; padding: 6px 12px; margin-top: 8px; min-height: 40px; display: flex; align-items: center; justify-content: space-between;'>"
                        f"<span style='font-size: 12px; font-weight: 800; color: #60a5fa;'>🔵 PRISES BLEU{cautions_icons_b} :</span>"
                        f"<div>{pastilles_html_b}</div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )

                # --- RETOUR HAPTIQUE (VIBRATION TACTILE MOBILE) & UNDO GLOBAL ---
                components.html("""
                <script>
                (function() {
                    try {
                        var pDoc = window.parent.document;
                        if (!pDoc || pDoc._fflda_haptic_init) return;
                        pDoc._fflda_haptic_init = true;
                        pDoc.addEventListener('pointerdown', function(e) {
                            var btn = e.target.closest('button');
                            if (!btn) return;
                            var txt = (btn.innerText || btn.textContent || '').trim();
                            if (/^(\\+[1245]|⌫|⚠️|🗑️|↩️|▶️|⏸️|🔄|\\-10s|\\+10s)/.test(txt)) {
                                if (window.navigator && window.navigator.vibrate) {
                                    try { window.navigator.vibrate(40); } catch(err) {}
                                }
                            }
                        }, {passive: true});
                    } catch(e) {}
                })();
                </script>
                """, height=0)

                # Bouton Annuler la dernière action (Undo rapide)
                has_actions_to_undo = bool(actions_r or actions_b)
                col_und_l, col_und_c, col_und_r = st.columns([1, 2.4, 1])
                with col_und_c:
                    st.button(
                        "↩️ Annuler la dernière action (Undo)",
                        key=f"btn_undo_global_{m_id}_{idx_match_actif}",
                        on_click=annuler_derniere_action_globale,
                        args=(m_id,),
                        disabled=not has_actions_to_undo,
                        use_container_width=True,
                        help="Annule immédiatement la dernière action saisie (Rouge ou Bleu) sans réinitialiser le match"
                    )

                # Bandeau de départage officiel UWW si égalité
                if score_r == score_b and score_r > 0:
                    couleur_meneur = "#dc2626" if meneur == "Rouge" else "#2563eb"
                    badge_meneur = f"<span style='background-color: {couleur_meneur}; color: white; padding: 4px 12px; border-radius: 8px; font-weight: 900; font-size: 13px;'>👑 AVANTAGE {meneur.upper()}</span>"
                    st.markdown(
                        f"<div style='background: linear-gradient(90deg, #1e293b, #0f172a); border: 2px solid #eab308; border-radius: 10px; padding: 9px 14px; margin: 10px 0 6px 0; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 14px rgba(234,179,8,0.22);'>"
                        f"<div><span style='color: #facc15; font-weight: 900; font-size: 14px;'>⚖️ ÉGALITÉ ({score_r} - {score_b}) — DÉPARTAGE OFFICIEL UWW :</span> <span style='color: #ffffff; font-weight: 700; font-size: 14px;'>{motif_departage}</span></div>"
                        f"<div>{badge_meneur}</div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )

                st.markdown("<div style='margin-top: 14px;'></div>", unsafe_allow_html=True)
            
                # Détermination du vainqueur par défaut
                default_vainq = m_actuel.get("vainqueur")
                if not default_vainq or default_vainq not in ["Rouge", "Bleu"]:
                    default_vainq = meneur if meneur else "Rouge"
                elif meneur and not est_deja_termine:
                    default_vainq = meneur

                opt_vainq_r = "🔴 VICTOIRE ROUGE"
                opt_vainq_b = "🔵 VICTOIRE BLEU"

                # Ligne combinée : Vainqueur Rouge / Bleu ET Type de victoire
                col_row_v1, col_row_v2 = st.columns([3, 2])
                with col_row_v1:
                    st.markdown("<div id='radio_vainqueur_marker'></div><div style='font-size: 17px; font-weight: 800; margin-bottom: 6px;'>🏆 Vainqueur désigné par l'arbitre :</div>", unsafe_allow_html=True)
                    vainqueur_saisi = st.radio(
                        "🏆 Vainqueur désigné par l'arbitre :",
                        [opt_vainq_r, opt_vainq_b],
                        index=0 if default_vainq == "Rouge" else 1,
                        horizontal=True,
                        label_visibility="collapsed",
                        key=f"rad_v_{m_id}_{idx_match_actif}_{score_r}_{score_b}_{default_vainq}"
                    )
                    v_pur = "Rouge" if "ROUGE" in vainqueur_saisi else "Bleu"

                with col_row_v2:
                    st.markdown("<div style='font-size: 17px; font-weight: 800; margin-bottom: 6px;'>⚖️ Type de victoire :</div>", unsafe_allow_html=True)
                    opts_type = [
                        "— Choisir le type de victoire —",
                        "⚡ VT (Victoire par Tombé)",
                        "💥 VST (Grande Supériorité Technique)",
                        "🎯 VP (Victoire aux Points)",
                        "🛑 Abandon / Forfait / Disqualification"
                    ]
                    
                    default_t_idx = 0
                    if m_actuel.get("type_victoire"):
                        cur_tv = str(m_actuel.get("type_victoire")).upper()
                        for idx_o, opt_str in enumerate(opts_type):
                            if idx_o > 0 and (cur_tv in opt_str or ("VT" in cur_tv and "VT" in opt_str) or ("VST" in cur_tv and "VST" in opt_str) or ("VP" in cur_tv and "VP" in opt_str and "VST" not in cur_tv)):
                                default_t_idx = idx_o
                                break
                            
                    type_vic_saisi = st.selectbox(
                        "⚖️ Type de victoire :",
                        opts_type,
                        index=default_t_idx,
                        label_visibility="collapsed",
                        key=f"sel_type_{m_id}_{idx_match_actif}"
                    )
                    has_selected_type = (type_vic_saisi != "— Choisir le type de victoire —")
                    if "VT" in type_vic_saisi:
                        code_type_vic = "VT"
                    elif "VST" in type_vic_saisi:
                        code_type_vic = "VST"
                    elif "VP" in type_vic_saisi:
                        code_type_vic = "VP"
                    elif any(x in type_vic_saisi for x in ["Abandon", "Forfait", "Disqualification"]):
                        code_type_vic = "AB"
                    else:
                        code_type_vic = ""

                    if has_selected_type:
                        badge_prev = formater_badge_victoire_html(code_type_vic, compact=False)
                        st.markdown(f"<div style='margin-top: 4px; text-align: right;'>Aperçu : {badge_prev}</div>", unsafe_allow_html=True)

                pt_r_calc, pt_b_calc = calculer_pts_fflda_match(cat_m, v_pur, code_type_vic if code_type_vic else "VT", score_r, score_b)
            
                libelle_save = f"💾 METTRE À JOUR LE COMBAT N°{m_actuel.get('match_num')} (CORRECTION)" if est_deja_termine else "💾 ENREGISTRER LE COMBAT & PASSER AU SUIVANT ⏭️"

                btn_save = st.button(
                    libelle_save, 
                    type="primary", 
                    use_container_width=True,
                    key=f"btn_save_match_{m_id}_{idx_match_actif}"
                )
            
                if btn_save:
                    if not has_selected_type or not code_type_vic:
                        st.error("⚠️ **Action obligatoire :** Vous devez obligatoirement sélectionner un **type de victoire** dans le menu déroulant (VT, VST, VP, etc.) avant de pouvoir enregistrer ce combat !")
                    else:
                        base_tot_r = str(m_actuel.get("tot_r_cell") or "").split("#")[0].strip()
                        base_tot_b = str(m_actuel.get("tot_b_cell") or "").split("#")[0].strip()
                        str_act_r = ",".join(str(a.get("val", 0)) for a in actions_r if a.get("val", 0) > 0)
                        str_act_b = ",".join(str(a.get("val", 0)) for a in actions_b if a.get("val", 0) > 0)
                        full_tot_r = f"{base_tot_r}#{str_act_r}" if (base_tot_r and str_act_r) else (base_tot_r or (f"#{str_act_r}" if str_act_r else ""))
                        full_tot_b = f"{base_tot_b}#{str_act_b}" if (base_tot_b and str_act_b) else (base_tot_b or (f"#{str_act_b}" if str_act_b else ""))

                        m_actuel["statut"] = "Terminé"
                        m_actuel["vainqueur"] = v_pur
                        m_actuel["type_victoire"] = code_type_vic
                        m_actuel["score_rouge"] = score_r
                        m_actuel["score_bleu"] = score_b
                        m_actuel["cautions_rouge"] = cautions_r
                        m_actuel["cautions_bleu"] = cautions_b
                        m_actuel["pt_clt_rouge"] = pt_r_calc
                        m_actuel["pt_clt_bleu"] = pt_b_calc
                        m_actuel["tot_r_cell"] = full_tot_r
                        m_actuel["tot_b_cell"] = full_tot_b
                        m_actuel["actions_rouge"] = list(actions_r)
                        m_actuel["actions_bleu"] = list(actions_b)

                        sauvegarder_cache_actions_match(c_sess, m_actuel.get("id"), actions_r, actions_b, cautions_r, cautions_b)
                
                        if is_supa:
                            match_payload = {
                                "statut": "Terminé",
                                "vainqueur": v_pur,
                                "type_victoire": code_type_vic,
                                "score_rouge": score_r,
                                "score_bleu": score_b,
                                "pt_clt_rouge": pt_r_calc,
                                "pt_clt_bleu": pt_b_calc,
                                "tot_r_cell": full_tot_r,
                                "tot_b_cell": full_tot_b
                            }
                            res_patch, err_patch = supabase_request("matchs_lutte", method="PATCH", data=match_payload, params={"id": f"eq.{m_actuel.get('id')}"})
                            if err_patch:
                                st.error(f"⚠️ Erreur de synchronisation Supabase : {err_patch}")
                
                        for idx_all, m_it in enumerate(st.session_state.get("matchs_direct", [])):
                            if m_it.get("id") == m_actuel.get("id"):
                                st.session_state["matchs_direct"][idx_all] = m_actuel
                                break

                        dict_victoire = {
                            "idx": idx_match_actif,
                            "match_id": m_actuel.get("id") or f"T{tapis_num_actif}_M{m_actuel.get('match_num', idx_match_actif)}",
                            "ts_termine": pytime.time(),
                            "vainqueur": v_pur,
                            "type_victoire": code_type_vic,
                            "score_r": score_r,
                            "score_b": score_b,
                            "lutteur_r": m_actuel.get("lutteur_rouge", ""),
                            "club_r": m_actuel.get("club_rouge", ""),
                            "lutteur_b": m_actuel.get("lutteur_bleu", ""),
                            "club_b": m_actuel.get("club_bleu", ""),
                            "match_num": m_actuel.get("match_num", idx_match_actif + 1),
                            "categorie": m_actuel.get("categorie", ""),
                            "tour": m_actuel.get("tour", "")
                        }

                        publier_mise_a_jour_match(c_sess, m_actuel)
                        publier_actions_match_sync(
                            m_actuel.get("id"),
                            actions_r=actions_r,
                            actions_b=actions_b,
                            cautions_r=cautions_r,
                            cautions_b=cautions_b,
                            extra={
                                "statut": "Terminé",
                                "vainqueur": v_pur,
                                "type_victoire": code_type_vic,
                                "ts_termine": dict_victoire["ts_termine"],
                                "score_r": score_r,
                                "score_b": score_b
                            }
                        )
                            
                        if not est_deja_termine and idx_match_actif < nb_total_tapis - 1:
                            nouv_idx = idx_match_actif + 1
                            st.session_state[cle_sess_idx] = nouv_idx
                            if matchs_du_tapis and 0 <= nouv_idx < len(matchs_du_tapis):
                                m_tgt = matchs_du_tapis[nouv_idx]
                                definir_combat_actif_tapis(
                                    tapis_num_actif, 
                                    nouv_idx, 
                                    m_tgt.get("id") or f"T{tapis_num_actif}_M{m_tgt.get('match_num', nouv_idx)}",
                                    dernier_termine=dict_victoire
                                )
                            st.session_state.pop(combo_key, None)
                        else:
                            definir_combat_actif_tapis(
                                tapis_num_actif, 
                                idx_match_actif, 
                                m_actuel.get("id") or f"T{tapis_num_actif}_M{m_actuel.get('match_num', idx_match_actif)}",
                                dernier_termine=dict_victoire
                            )
                        
                        st.toast(f"✅ Combat n°{m_actuel.get('match_num')} enregistré avec succès !")
                        st.rerun()

            _afficher_boitier_score_fragment()

            # Liste récapitulative des combats du tapis
            with st.expander(f"📋 Voir tous les combats du Tapis {tapis_num_actif} ({nb_termines_tapis}/{nb_total_tapis} terminés) — Accès direct / Correction", expanded=False):
                for i_recap, m_r in enumerate(matchs_du_tapis):
                    est_m_actif = (i_recap == idx_match_actif)
                    est_termine = (m_r.get("statut") == "Terminé")
                    bg_card = "#e8f5e9" if est_termine else "#fff8e1"
                    if est_m_actif:
                        bg_card = "#e3f2fd"
                    
                    c_desc, c_btn = st.columns([4, 1.2])
                    with c_desc:
                        badge_icon = "✅" if est_termine else "⏳"
                        if est_termine:
                            b_vic = formater_badge_victoire_html(m_r.get('type_victoire', ''), compact=False)
                            col_v = "#c62828" if m_r.get('vainqueur') == "Rouge" else ("#1565c0" if m_r.get('vainqueur') == "Bleu" else "#1e293b")
                            score_txt = f" — Score : <b>🔴 {m_r.get('score_rouge', 0)} - 🔵 {m_r.get('score_bleu', 0)}</b> {b_vic} — Vainq: <b style='color:{col_v};'>{m_r.get('vainqueur', '')}</b>"
                        else:
                            score_txt = ""
                        st.markdown(
                            f"<div style='background-color: {bg_card}; padding: 10px 14px; border-radius: 8px; margin-bottom: 6px; border-left: 5px solid {'#2e7d32' if est_termine else ('#1976d2' if est_m_actif else '#ffa000')};'>"
                            f"<b>Combat n°{m_r.get('match_num')}</b> ({m_r.get('heure')}) — <i>{m_r.get('categorie')} · Tour {m_r.get('tour')}</i> {badge_icon}<br>"
                            f"<span style='color: #c62828; font-weight: bold;'>🔴 {m_r.get('lutteur_rouge')}</span> vs <span style='color: #1565c0; font-weight: bold;'>🔵 {m_r.get('lutteur_bleu')}</span>{score_txt}"
                            f"</div>",
                            unsafe_allow_html=True
                        )
                    with c_btn:
                        label_btn = "👁️ Actif" if est_m_actif else ("✏️ Corriger" if est_termine else "▶️ Ouvrir")
                        if st.button(label_btn, key=f"btn_edit_{m_r.get('id')}", use_container_width=True, disabled=est_m_actif):
                            st.session_state[cle_sess_idx] = i_recap
                            definir_combat_actif_tapis(tapis_num_actif, i_recap, m_r.get("id") or f"T{tapis_num_actif}_M{m_r.get('match_num', i_recap)}")
                            st.session_state.pop(combo_key, None)
                            st.rerun()

elif mode_app.startswith("3"):
    st.markdown("### Module de Fin de Tournoi & Bilans Fédéraux")
    st.info(f"Importez vos scores en direct ou votre fichier Excel complété pour la compétition **{nom_competition}** afin de générer automatiquement les classements officiels individuels et par club.")
    
    def_source_idx = 1 if (st.session_state.get("source_scores_choix") or "").endswith("(.xlsx)") else 0
    source_scores = st.radio(
        "Source des scores à analyser :", 
        [
            "Scores saisis en direct (Tablettes / Tapis)",
            "Fichier Excel complété (.xlsx)"
        ], 
        index=def_source_idx,
        horizontal=True
    )
    
    wb_res = None
    wb_f = None
    tous_les_resultats = []
    nom_comp_officiel = nom_competition

    if "direct" in source_scores.lower():
        matchs_direct_eval = st.session_state.get("matchs_direct", [])
        _, _, is_supa_chk = get_supabase_config()
        if is_supa_chk:
            c_sess = st.session_state.get("code_session", "")
            params_eval = {"order": "tapis.asc,match_num.asc"}
            if c_sess and c_sess not in ["FFLDA-ADMIN", "FFLDA2026"]:
                params_eval["code_organisateur"] = f"eq.{c_sess}"
            res_s_eval, _ = supabase_request("matchs_lutte", params=params_eval)
            if res_s_eval and isinstance(res_s_eval, list) and len(res_s_eval) > 0:
                tournois_ids_eval = list(dict.fromkeys([m.get("tournoi_id", "") for m in res_s_eval if m.get("tournoi_id")]))
                if len(tournois_ids_eval) > 1:
                    t_eval_sel = st.selectbox("Choisir le tournoi pour le bilan :", tournois_ids_eval, index=len(tournois_ids_eval)-1, key="sb_tourn_eval")
                    matchs_direct_eval = [m for m in res_s_eval if m.get("tournoi_id") == t_eval_sel]
                else:
                    matchs_direct_eval = res_s_eval
                
        nb_term = sum(1 for m in matchs_direct_eval if m.get("statut") == "Terminé")
        nb_tot = len(matchs_direct_eval)
        
        st.markdown(f"**Avancement global des tapis :** {nb_term} / {nb_tot} combats enregistrés")
        
        if nb_tot == 0:
            st.warning("⚠️ Aucun combat n'a été initialisé. Veuillez d'abord générer un tournoi dans le Mode 1.")
        else:
            col_btn_calc1, col_btn_calc2 = st.columns([2, 1])
            with col_btn_calc1:
                if st.button("Recalculer les classements & Actualiser le Bilan avec les scores saisis", type="primary", use_container_width=True):
                    st.rerun()
            with col_btn_calc2:
                st.caption(f"{nb_term} combat(s) terminé(s) / {nb_tot}")

            # Détermination précise du nom de la compétition
            if matchs_direct_eval and matchs_direct_eval[0].get("tournoi_id"):
                nom_comp_officiel = matchs_direct_eval[0].get("tournoi_id")
            elif st.session_state.get("nom_competition_active"):
                nom_comp_officiel = st.session_state.get("nom_competition_active")

            # 1. Calcul officiel et garanti de TOUS les classements et points directement depuis les scores saisis
            tous_les_resultats = calculer_resultats_depuis_matchs_direct(matchs_direct_eval)

            # Option d'association d'un classeur Excel initial personnalisé
            with st.expander("📎 Associer un classeur Excel initial personnalisé (optionnel)", expanded=False):
                fichier_excel_initial = st.file_uploader("Classeur initial (.xlsx)", type=["xlsx"], key="up_base_pour_direct")
                if fichier_excel_initial:
                    excel_base = fichier_excel_initial.read()
                    st.session_state["excel_tournoi_base"] = excel_base
                    safe_c = re.sub(r'[^a-zA-Z0-9_-]', '_', str(c_sess or "ORG"))
                    try:
                        os.makedirs(".cache_tournois", exist_ok=True)
                        with open(f".cache_tournois/excel_base_{safe_c}.xlsx", "wb") as f_save:
                            f_save.write(excel_base)
                    except Exception:
                        pass
                    st.success("Classeur Excel personnalisé associé avec succès !")
                    st.rerun()

            # 2. Préparation du classeur Excel officiel pour les exports PDF et Excel
            cats_direct = list({m.get("categorie") for m in matchs_direct_eval if m.get("categorie")})
            excel_base = recuperer_excel_base(c_sess, nom_comp_officiel, categories_attendues=cats_direct)
            if excel_base:
                try:
                    buf_inj = injecter_scores_matchs_dans_classeur(excel_base, matchs_direct_eval)
                    wb_res = openpyxl.load_workbook(buf_inj, data_only=False)
                except Exception:
                    wb_res = generer_classeur_complet_officiel(nom_comp_officiel, tous_les_resultats, matchs_direct_eval)
            else:
                wb_res = generer_classeur_complet_officiel(nom_comp_officiel, tous_les_resultats, matchs_direct_eval)
            wb_f = None

    else:
        fichier_resultats = st.file_uploader("Sélectionner le fichier Excel complété (.xlsx)", type=["xlsx"])
        if fichier_resultats is not None:
            try:
                wb_res = openpyxl.load_workbook(fichier_resultats, data_only=True)
                fichier_resultats.seek(0)
                try:
                    wb_f = openpyxl.load_workbook(fichier_resultats, data_only=False)
                except Exception:
                    wb_f = None
                
                # Recherche dynamique du nom de la compétition dans le fichier Excel
                nom_comp_detecte = None
                for s_name in wb_res.sheetnames:
                    ws_chk = wb_res[s_name]
                    for r_chk in [1, 2]:
                        val_c = str(ws_chk.cell(row=r_chk, column=1).value or "").strip()
                        m_comp = re.search(r"COMPÉTITION\s*:\s*([^—\n]+)", val_c, re.IGNORECASE)
                        if m_comp:
                            nom_comp_detecte = m_comp.group(1).strip()
                            break
                    if nom_comp_detecte:
                        break
                nom_comp_officiel = nom_comp_detecte if nom_comp_detecte else nom_competition
                tous_les_resultats = extraire_resultats_classeur_excel(wb_res, wb_f)
            except Exception as e_xl:
                st.error(f"Erreur lors de la lecture du fichier Excel : {e_xl}")
                wb_res = None
                tous_les_resultats = []
        else:
            wb_res = None
            tous_les_resultats = []
    
    if tous_les_resultats and wb_res is not None:
        try:
            st.success(f"Bilan officiel calculé avec succès pour : **{nom_comp_officiel}** ({len(tous_les_resultats)} lutteurs engagés) !")
            df_bilan = pd.DataFrame(tous_les_resultats)
            if not df_bilan.empty:
                if "Clt" in df_bilan.columns:
                    df_bilan["Clt"] = df_bilan["Clt"].apply(lambda v: str(nettoyer_rang(v)))
                if "Points" in df_bilan.columns:
                    df_bilan["Points"] = pd.to_numeric(df_bilan["Points"], errors="coerce").fillna(0).astype(int)
                for col_str in ["Nom", "Club", "Comité", "Poids", "Poule"]:
                    if col_str in df_bilan.columns:
                        df_bilan[col_str] = df_bilan[col_str].astype(str).replace(["None", "nan"], "")
                # Tri numérique par Poule, Clt (1er, 2ème, 3ème...) et Points
                df_bilan = trier_dataframe_bilan(df_bilan)
            
            # Transmettre silencieusement les résultats et métriques du bilan au Webhook Facturation
            try:
                nb_peses_bilan = sum(1 for p in tous_les_resultats if str(p.get("Clt", "NR")) != "NR")
                enregistrer_log_facturation(
                    code_organisateur=st.session_state.get("code_session", "ORGANISATEUR"),
                    nom_tournoi=f"{nom_comp_officiel} [BILAN FINAL]",
                    nb_inscrits=len(df_bilan),
                    nb_peses=nb_peses_bilan,
                    nb_matchs=len(df_bilan),
                    details_resultats=tous_les_resultats
                )
            except Exception:
                pass
            
            df_bilan_edited = df_bilan

            # --- CALCUL DU CLASSEMENT DES CLUBS ET DES COMITÉS RÉGIONAUX ---
            bareme_points = {1: 4, 2: 3, 3: 2, 4: 1}
            points_clubs = {}
            points_comites = {}

            for _, row in df_bilan_edited.iterrows():
                nom = str(row.get("Nom", "")).strip()
                club = str(row.get("Club", "")).strip()
                comite = str(row.get("Comité", "Comité Non Renseigné")).strip()
                if not comite or comite in ["None", "nan", "-"]:
                    comite = "Comité Non Renseigné"
                
                # Exclusion des faux clubs (labels d'en-têtes ou poids)
                if not club or club in ["", "-", "None", "nan", "CLUB", "Club", "Indépendant"] or "kg" in club.lower() or club.isdigit():
                    continue
                
                if club not in points_clubs:
                    points_clubs[club] = {"Club": club, "Points Club": 0, "1ers": 0, "2èmes": 0, "3èmes": 0, "4èmes": 0}
                if comite and comite not in points_comites:
                    points_comites[comite] = {"Comité Régional": comite, "Points Comité": 0, "1ers": 0, "2èmes": 0, "3èmes": 0, "4èmes": 0}
                
                clt_val = nettoyer_rang(row.get("Clt"))
                if clt_val == "NR" or not isinstance(clt_val, int):
                    continue
                
                clt_num = clt_val
                is_u7_poule = "u7" in str(row.get("Poule", "")).lower()
                pts_attribués = 1 if is_u7_poule else bareme_points.get(clt_num, 0)
                
                # Ranking Clubs
                if club in points_clubs:
                    points_clubs[club]["Points Club"] += pts_attribués
                    if not is_u7_poule:
                        if clt_num == 1: points_clubs[club]["1ers"] += 1
                        elif clt_num == 2: points_clubs[club]["2èmes"] += 1
                        elif clt_num == 3: points_clubs[club]["3èmes"] += 1
                        elif clt_num == 4: points_clubs[club]["4èmes"] += 1

                # Ranking Comités Régionaux
                if comite in points_comites:
                    points_comites[comite]["Points Comité"] += pts_attribués
                    if not is_u7_poule:
                        if clt_num == 1: points_comites[comite]["1ers"] += 1
                        elif clt_num == 2: points_comites[comite]["2èmes"] += 1
                        elif clt_num == 3: points_comites[comite]["3èmes"] += 1
                        elif clt_num == 4: points_comites[comite]["4èmes"] += 1

            if points_clubs:
                df_clubs = pd.DataFrame(list(points_clubs.values())).sort_values(
                    by=["Points Club", "1ers", "2èmes", "3èmes", "4èmes"], 
                    ascending=False
                ).reset_index(drop=True)
                df_clubs.index = range(1, len(df_clubs) + 1)
                df_clubs.insert(0, "Clt Club", df_clubs.index)
            else:
                df_clubs = pd.DataFrame(columns=["Clt Club", "Club", "Points Club", "1ers", "2èmes", "3èmes", "4èmes"])

            if points_comites:
                df_comites = pd.DataFrame(list(points_comites.values())).sort_values(
                    by=["Points Comité", "1ers", "2èmes", "3èmes", "4èmes"], 
                    ascending=False
                ).reset_index(drop=True)
                df_comites.index = range(1, len(df_comites) + 1)
                df_comites.insert(0, "Clt Comité", df_comites.index)
            else:
                df_comites = pd.DataFrame(columns=["Clt Comité", "Comité Régional", "Points Comité", "1ers", "2èmes", "3èmes", "4èmes"])

            # --- ONGLETS D'AFFICHAGE DU BILAN OFFICIEL ---
            tab_bilan_1, tab_bilan_2, tab_bilan_3 = st.tabs([
                "Classements Individuels (U7 / U9 / U11 / U13)", 
                "Classement Général des Clubs",
                "Classement des Comités Régionaux"
            ])
            
            with tab_bilan_1:
                st.subheader(f"Classements Individuels Officiels — {nom_comp_officiel}")
                
                # Récupération et tri de toutes les colonnes de tours présentes
                def _sort_tour_col_key(col_c):
                    parts = str(col_c).split()
                    if len(parts) >= 2 and parts[1].isdigit():
                        return int(parts[1])
                    return 999
                tours_cols_tous = sorted(
                    [c for c in df_bilan_edited.columns if str(c).startswith("Tour ")],
                    key=_sort_tour_col_key
                )

                def _meme_poule_cat(c1, c2):
                    s1 = re.sub(r'[^a-zA-Z0-9]', '', str(c1 or '').lower())
                    s2 = re.sub(r'[^a-zA-Z0-9]', '', str(c2 or '').lower())
                    return s1 == s2 or (len(s1) > 4 and (s1 in s2 or s2 in s1))

                for poule in df_bilan_edited['Poule'].unique():
                    st.markdown(f"#### 🤼 {poule}")
                    sous_df_full = df_bilan_edited[df_bilan_edited['Poule'] == poule].copy()
                    
                    # Sélection des tours ayant des données pour cette poule
                    cols_tours_poule = [
                        c for c in tours_cols_tous
                        if c in sous_df_full.columns and not (sous_df_full[c].isna() | (sous_df_full[c] == "")).all()
                    ]
                    
                    cols_ordre = ["Clt", "Nom", "Club"]
                    if "Comité" in sous_df_full.columns and not sous_df_full["Comité"].isin(["", "-", "None", "nan", "Comité Non Renseigné"]).all():
                        cols_ordre.append("Comité")
                    cols_ordre.extend(cols_tours_poule)
                    cols_ordre.append("Points")
                    if "Total Vict" in sous_df_full.columns:
                        cols_ordre.append("Total Vict")
                    cols_ordre.append("Poids")
                    
                    cols_exist = [c for c in cols_ordre if c in sous_df_full.columns]
                    sous_df = sous_df_full[cols_exist].copy()
                    if "Points" in sous_df.columns:
                        sous_df['Points'] = pd.to_numeric(sous_df['Points'], errors='coerce').fillna(0).astype(int)
                    if "Total Vict" in sous_df.columns:
                        sous_df['Total Vict'] = pd.to_numeric(sous_df['Total Vict'], errors='coerce').fillna(0).astype(int)
                    sous_df = sous_df.reset_index(drop=True)
                    
                    col_cfg = {
                        "Clt": st.column_config.TextColumn("Clt / Rang", width="small"),
                        "Nom": st.column_config.TextColumn("NOM Prénom", width="medium"),
                        "Club": st.column_config.TextColumn("Club", width="medium"),
                        "Comité": st.column_config.TextColumn("Comité", width="small"),
                        "Points": st.column_config.NumberColumn("Total Pts", format="%d", help="Total des points de classement"),
                        "Total Vict": st.column_config.NumberColumn("Total Vict", format="%d", help="Nombre de victoires"),
                        "Poids": st.column_config.TextColumn("Poids", width="small"),
                    }
                    for ct in cols_tours_poule:
                        col_cfg[ct] = st.column_config.TextColumn(ct, width="small", help=f"Points de classement marqués au {ct}")

                    st.dataframe(
                        sous_df,
                        column_config=col_cfg,
                        use_container_width=True,
                        hide_index=True
                    )

                    # Détail des combats par tour de cette poule
                    matchs_poule = [
                        m for m in (matchs_direct_eval or [])
                        if _meme_poule_cat(m.get("categorie"), poule)
                    ]
                    if matchs_poule:
                        with st.expander(f"🎯 Détail des combats par tour — {poule}", expanded=False):
                            tours_dict = {}
                            for m in sorted(matchs_poule, key=lambda x: (int(x.get("tour") or 1), int(x.get("match_num") or 0))):
                                t_num = m.get("tour", 1)
                                tours_dict.setdefault(t_num, []).append(m)
                            
                            for t_num in sorted(tours_dict.keys()):
                                st.markdown(f"**Tour {t_num} :**")
                                for m in tours_dict[t_num]:
                                    st_m = m.get("statut", "À venir")
                                    v_type = m.get("victoire_type", "") or m.get("type_victoire", "")
                                    pts_r = m.get("score_rouge", 0)
                                    pts_b = m.get("score_bleu", 0)
                                    pt_clt_r = m.get("pt_clt_rouge", "-")
                                    pt_clt_b = m.get("pt_clt_bleu", "-")
                                    vainq = m.get("vainqueur", "")
                                    lr = m.get("lutteur_rouge", "Rouge")
                                    lb = m.get("lutteur_bleu", "Bleu")
                                    
                                    if st_m == "Terminé":
                                        vainq_nom = lr if vainq == "Rouge" else (lb if vainq == "Bleu" else vainq)
                                        v_color = "#c62828" if vainq == "Rouge" else "#1565c0"
                                        badge_vic_poule = formater_badge_victoire_html(v_type, compact=False)
                                        vainq_html = f"<b style='color:{v_color};'>{vainq_nom}</b> gagne {badge_vic_poule}"
                                        score_html = f"Score : 🔴 {pts_r} - 🔵 {pts_b} | Pts Clt : 🔴 <b>{pt_clt_r}</b> - 🔵 <b>{pt_clt_b}</b>"
                                        icon_c = "✅"
                                    else:
                                        icon_c = "⏳"
                                        vainq_html = "<i>(Combat non terminé)</i>"
                                        score_html = ""
                                    
                                    st.markdown(
                                        f"- {icon_c} **Combat n°{m.get('match_num', '?')}** (Tapis {m.get('tapis', '?')}) : "
                                        f"<span style='color:#c62828; font-weight:600;'>🔴 {lr}</span> vs "
                                        f"<span style='color:#1565c0; font-weight:600;'>🔵 {lb}</span> — {vainq_html} {('— ' + score_html) if score_html else ''}",
                                        unsafe_allow_html=True
                                    )

            with tab_bilan_2:
                st.subheader(f"Podium des Clubs Engagés — {nom_comp_officiel}")
                st.markdown("*Barème officiel FFLDA : 1er = 4 pts | 2ème = 3 pts | 3ème = 2 pts | 4ème = 1 pt*")
                st.dataframe(df_clubs, use_container_width=True, hide_index=True)

            with tab_bilan_3:
                st.subheader(f"🏛️ Classement Officiel des Comités Régionaux — {nom_comp_officiel}")
                st.markdown("*Barème officiel FFLDA : 1er = 4 pts | 2ème = 3 pts | 3ème = 2 pts | 4ème = 1 pt*")
                st.dataframe(df_comites, use_container_width=True, hide_index=True)

            # Intégration des feuilles officielles de Bilan directement dans le classeur Excel officiel
            if "Classement Clubs" in wb_res.sheetnames:
                del wb_res["Classement Clubs"]
            ws_clubs = wb_res.create_sheet("Classement Clubs")

            if "Classement Comités" in wb_res.sheetnames:
                del wb_res["Classement Comités"]
            ws_comites = wb_res.create_sheet("Classement Comités")

            if "Classements Individuels" in wb_res.sheetnames:
                del wb_res["Classements Individuels"]
            ws_indiv = wb_res.create_sheet("Classements Individuels")

            bleu_fflda = PatternFill("solid", fgColor="0055A4")
            rouge_fflda = PatternFill("solid", fgColor="EF4135")
            gris_zebrage = PatternFill("solid", fgColor="F2F5F8")
            fond_blanc = PatternFill("solid", fgColor="FFFFFF")
            or_fill = PatternFill("solid", fgColor="FFF2CC")
            argent_fill = PatternFill("solid", fgColor="EFEFEF")
            bronze_fill = PatternFill("solid", fgColor="F8CBAD")
            
            font_titre = Font(name="Arial", size=15, bold=True, color="0055A4")
            font_section = Font(name="Arial", size=12, bold=True, color="FFFFFF")
            font_entete = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            font_data = Font(name="Arial", size=11, color="000000")
            font_data_bold = Font(name="Arial", size=11, bold=True, color="0055A4")
            b_fin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), 
                           top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
            
            # 1. Feuille Classement Clubs
            ws_clubs.views.sheetView[0].showGridLines = True
            ws_clubs.cell(row=1, column=1, value=f"COMPÉTITION : {nom_comp_officiel.upper()}").font = font_titre
            ws_clubs.cell(row=2, column=1, value="CLASSEMENT OFFICIEL DES CLUBS - FFLDA").font = Font(name="Arial", size=12, bold=True, color="666666")
            ws_clubs.cell(row=3, column=1, value=f"Édité le {datetime.now().strftime('%d/%m/%Y à %H:%M')}").font = Font(name="Arial", size=9, italic=True, color="888888")
            
            for col_idx, col_name in enumerate(df_clubs.columns, 1):
                cell = ws_clubs.cell(row=5, column=col_idx, value=col_name)
                cell.fill, cell.font, cell.alignment = bleu_fflda, font_entete, Alignment(horizontal="center", vertical="center")
            ws_clubs.row_dimensions[5].height = 25
            
            for r_offset, r_data in enumerate(df_clubs.values, 6):
                ws_clubs.row_dimensions[r_offset].height = 22
                is_even = (r_offset % 2 == 0)
                for c_offset, val in enumerate(r_data, 1):
                    cell = ws_clubs.cell(row=r_offset, column=c_offset, value=val)
                    cell.border, cell.font = b_fin, font_data
                    cell.fill = gris_zebrage if is_even else fond_blanc
                    cell.alignment = Alignment(horizontal="center", vertical="center")

            ws_clubs.column_dimensions['A'].width = 12
            ws_clubs.column_dimensions['B'].width = 30
            ws_clubs.column_dimensions['C'].width = 15
            ws_clubs.column_dimensions['D'].width = 12
            ws_clubs.column_dimensions['E'].width = 12
            ws_clubs.column_dimensions['F'].width = 12
            ws_clubs.column_dimensions['G'].width = 12

            # 2. Feuille Classement Comités Régionaux
            ws_comites.views.sheetView[0].showGridLines = True
            ws_comites.cell(row=1, column=1, value=f"COMPÉTITION : {nom_comp_officiel.upper()}").font = font_titre
            ws_comites.cell(row=2, column=1, value="🏛️ CLASSEMENT OFFICIEL DES COMITÉS RÉGIONAUX - FFLDA").font = Font(name="Arial", size=12, bold=True, color="666666")
            ws_comites.cell(row=3, column=1, value=f"Édité le {datetime.now().strftime('%d/%m/%Y à %H:%M')}").font = Font(name="Arial", size=9, italic=True, color="888888")
            
            for col_idx, col_name in enumerate(df_comites.columns, 1):
                cell = ws_comites.cell(row=5, column=col_idx, value=col_name)
                cell.fill, cell.font, cell.alignment = bleu_fflda, font_entete, Alignment(horizontal="center", vertical="center")
            ws_comites.row_dimensions[5].height = 25
            
            for r_offset, r_data in enumerate(df_comites.values, 6):
                ws_comites.row_dimensions[r_offset].height = 22
                is_even = (r_offset % 2 == 0)
                for c_offset, val in enumerate(r_data, 1):
                    cell = ws_comites.cell(row=r_offset, column=c_offset, value=val)
                    cell.border, cell.font = b_fin, font_data
                    cell.fill = gris_zebrage if is_even else fond_blanc
                    cell.alignment = Alignment(horizontal="center", vertical="center")

            ws_comites.column_dimensions['A'].width = 12
            ws_comites.column_dimensions['B'].width = 30
            ws_comites.column_dimensions['C'].width = 15
            ws_comites.column_dimensions['D'].width = 12
            ws_comites.column_dimensions['E'].width = 12
            ws_comites.column_dimensions['F'].width = 12
            ws_comites.column_dimensions['G'].width = 12

            # 3. Feuille Classements Individuels
            ws_indiv.views.sheetView[0].showGridLines = True
            ws_indiv.cell(row=1, column=1, value=f"COMPÉTITION : {nom_comp_officiel.upper()}").font = font_titre
            ws_indiv.cell(row=2, column=1, value="🏆 CLASSEMENTS INDIVIDUELS OFFICIELS - FFLDA").font = Font(name="Arial", size=12, bold=True, color="666666")
            ws_indiv.cell(row=3, column=1, value=f"Édité le {datetime.now().strftime('%d/%m/%Y à %H:%M')}").font = Font(name="Arial", size=9, italic=True, color="888888")
            
            row_cursor = 5
            for poule in df_bilan_edited['Poule'].unique():
                groupe = df_bilan_edited[df_bilan_edited['Poule'] == poule]
                cols_tours_grp = [
                    c for c in tours_cols_tous
                    if c in groupe.columns and not (groupe[c].isna() | (groupe[c] == "")).all()
                ]
                headers_poule = ["Clt", "NOM Prénom", "CLUB", "COMITÉ"] + cols_tours_grp + ["Total Pts", "Total Vict", "POIDS (kg)"]
                nb_cols_p = len(headers_poule)

                ws_indiv.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=nb_cols_p)
                cell_cat = ws_indiv.cell(row=row_cursor, column=1, value=f"  CATÉGORIE / POULE : {poule}")
                cell_cat.fill, cell_cat.font, cell_cat.alignment = bleu_fflda, font_section, Alignment(horizontal="left", vertical="center")
                ws_indiv.row_dimensions[row_cursor].height = 28
                row_cursor += 1
                
                for col_idx, h in enumerate(headers_poule, 1):
                    cell = ws_indiv.cell(row=row_cursor, column=col_idx, value=h)
                    cell.fill, cell.font, cell.alignment = rouge_fflda, font_entete, Alignment(horizontal="center", vertical="center")
                    ws_indiv.row_dimensions[row_cursor].height = 22
                row_cursor += 1
                
                for _, row in groupe.iterrows():
                    current_row = row_cursor
                    ws_indiv.row_dimensions[current_row].height = 20
                    
                    row_vals = [
                        row.get('Clt'),
                        row.get('Nom'),
                        row.get('Club'),
                        row.get('Comité', '')
                    ]
                    for ct in cols_tours_grp:
                        row_vals.append(row.get(ct, '-'))
                    row_vals.append(row.get('Points', 0))
                    row_vals.append(row.get('Total Vict', 0))
                    row_vals.append(row.get('Poids', ''))
                    
                    for col_idx, val in enumerate(row_vals, 1):
                        c_cell = ws_indiv.cell(row=current_row, column=col_idx, value=val)
                        c_cell.border = b_fin
                        c_cell.font = font_data
                        c_cell.alignment = Alignment(horizontal="center", vertical="center")
                        if col_idx == 2:
                            c_cell.alignment = Alignment(horizontal="left", vertical="center")
                        if col_idx == 1 or col_idx == (4 + len(cols_tours_grp) + 1):
                            c_cell.font = font_data_bold
                        if col_idx == 1:
                            clt_s = str(row.get('Clt', ''))
                            if clt_s == '1': c_cell.fill = or_fill
                            elif clt_s == '2': c_cell.fill = argent_fill
                            elif clt_s == '3': c_cell.fill = bronze_fill
                    
                    row_cursor += 1
                row_cursor += 2
            
            for col in ws_indiv.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws_indiv.column_dimensions[col_letter].width = max(max_len + 3, 10)

            # Configuration Impression A4 Excel
            for ws_b in [ws_clubs, ws_comites, ws_indiv]:
                ws_b.page_setup.orientation = ws_b.ORIENTATION_LANDSCAPE
                ws_b.page_setup.paperSize = ws_b.PAPERSIZE_A4
                ws_b.sheet_properties.pageSetUpPr.fitToPage = True
                ws_b.page_setup.fitToWidth = 1
                ws_b.page_setup.fitToHeight = 0

            # Ordonnancement officiel des onglets : Résumé, Grille, Bilans, puis toutes les poules
            def sheet_order_key(s):
                n = s.title.lower()
                if "résumé" in n or "resume" in n: return (0, 0)
                if "grille" in n: return (0, 1)
                if "classement club" in n or "classement des clubs" in n: return (0, 2)
                if "classement comit" in n: return (0, 3)
                if "classements individuels" in n: return (0, 4)
                if "u7" in n: return (1, 0, n)
                if "u9" in n: return (1, 1, n)
                if "u11" in n: return (1, 2, n)
                return (1, 3, n)

            wb_res._sheets.sort(key=sheet_order_key)

            # Génération du Dossier Officiel Complet en PDF (A4 Portrait - Format Excel FFLDA) identique au Mode 1
            pdf_bytes_dossier_complet = None
            try:
                pdf_bytes_dossier_complet = generer_pdf_depuis_classeur_excel(wb_res, nom_comp_officiel, wb=wb_f or wb_res)
            except Exception as e_pdf_c:
                st.warning(f"⚠️ Information : génération PDF du Dossier Complet : {e_pdf_c}")

            # Génération du Bilan seul en PDF (3 pages A4 Portrait)
            sheets_bilan_only = [wb_res[s] for s in ["Classement Clubs", "Classement Comités", "Classements Individuels"] if s in wb_res.sheetnames]
            pdf_bilan_only_bytes = None
            try:
                pdf_bilan_only_bytes = generer_pdf_depuis_classeur_excel(sheets_bilan_only, nom_comp_officiel, wb=wb_res)
            except Exception:
                pdf_bilan_only_bytes = None

            # Sauvegarde du classeur Excel officiel complet
            output_excel_complet = io.BytesIO()
            wb_res.save(output_excel_complet)
            excel_bytes_complet = output_excel_complet.getvalue()

            st.markdown("---")
            st.markdown("### 📥 Téléchargements Complets du Tournoi (Formats Officiels FFLDA)")
            col_b1, col_b2, col_b3 = st.columns(3)
            with col_b1:
                if pdf_bilan_only_bytes:
                    st.download_button(
                        label="📄 Bilan Officiel PDF\n(Clubs, Comités & Podiums)",
                        data=pdf_bilan_only_bytes,
                        file_name=f"Bilan_Officiel_{nom_comp_officiel.replace(' ', '_')}.pdf",
                        mime="application/pdf",
                        key="btn_pdf_bilan_seul",
                        use_container_width=True
                    )
            with col_b2:
                st.download_button(
                    label="📄 Dossier Officiel PDF\n(Tapis + Grilles + Bilans)",
                    data=pdf_bytes_dossier_complet or pdf_bilan_only_bytes,
                    file_name=f"Dossier_Officiel_Resultats_{nom_comp_officiel.replace(' ', '_')}.pdf",
                    mime="application/pdf",
                    key="btn_pdf_dossier_resultats",
                    use_container_width=True
                )
            with col_b3:
                st.download_button(
                    label="📥 Classeur Officiel Excel\n(.xlsx Complété)",
                    data=excel_bytes_complet,
                    file_name=f"Tournoi_Resultats_{nom_comp_officiel.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_excel_dossier_resultats",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"Erreur lors de l'analyse du fichier : {e}")

else:
    # --- MODE 1 : GÉNÉRATION DE TOURNOI ---
    if st.session_state.get("tournoi_reinitialise_succes"):
        st.success("🎉 **Tournoi précédent clôturé et réinitialisé avec succès !** Vous pouvez maintenant importer votre nouveau fichier d'inscrits ci-dessous pour lancer la nouvelle compétition.")
        if st.button("Masquer ce message", key="btn_hide_success_reset"):
            st.session_state.pop("tournoi_reinitialise_succes", None)
            st.rerun()

    if tournoi_actif_detecte:
        pct_prog = int(nb_termines_detectes / nb_combats_detectes * 100) if nb_combats_detectes > 0 else 0
        st.markdown(
            f"<div style='background: linear-gradient(135deg, #e8f5e9 0%, #f1f8e9 100%); border: 2px solid #2e7d32; border-radius: 12px; padding: 18px; margin-bottom: 16px; box-shadow: 0 3px 6px rgba(0,0,0,0.06);'>"
            f"<div style='display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;'>"
            f"  <div>"
            f"    <h3 style='margin: 0; color: #1b5e20; font-size: 19px;'>Un Tournoi est actuellement actif dans le Cloud !</h3>"
            f"    <p style='margin: 6px 0 0 0; color: #2e7d32; font-size: 14px;'>Compétition : <b>{tournoi_actif_detecte}</b> &nbsp;|&nbsp; <b>{nb_termines_detectes}</b> terminés sur <b>{nb_combats_detectes}</b> combats ({pct_prog}%)</p>"
            f"  </div>"
            f"  <span style='background: #2e7d32; color: white; padding: 4px 12px; border-radius: 12px; font-weight: bold; font-size: 12px;'>EN DIRECT</span>"
            f"</div>"
            f"<p style='margin: 10px 0 0 0; color: #555; font-size: 12.5px;'>Vos données de tournoi sont synchronisées avec les tables de marque. Choisissez une action ci-dessous :</p>"
            f"</div>",
            unsafe_allow_html=True
        )

        if st.button("Démarrer un Nouveau Tournoi", use_container_width=True, key="btn_declencher_nouveau_tournoi", type="secondary"):
            st.session_state["confirm_reset_tournoi"] = not st.session_state.get("confirm_reset_tournoi", False)
            st.rerun()

        # Boîte de confirmation sécurisée avec saisie du code de connexion
        if st.session_state.get("confirm_reset_tournoi"):
            st.markdown(
                f"<div style='background: #ffebee; border: 2px solid #c62828; border-radius: 10px; padding: 16px; margin-top: 12px; margin-bottom: 12px;'>"
                f"<h4 style='margin: 0; color: #b71c1c;'>⚠️ Attention : Écrasement des Données du Tournoi Actuel</h4>"
                f"<p style='margin: 8px 0 0 0; color: #c62828; font-size: 13.5px; line-height: 1.4;'>"
                f"Vous êtes sur le point de démarrer un nouveau tournoi. "
                f"<b>Cette action va écraser et effacer définitivement les combats et résultats du tournoi actuel « {tournoi_actif_detecte} »</b> sur le Cloud.<br>"
                f"Pour sécuriser cette opération, veuillez saisir votre code de connexion pour confirmer :"
                f"</p>"
                f"</div>",
                unsafe_allow_html=True
            )

            code_confirm_input = st.text_input(
                "🔑 Saisissez votre code de connexion pour déverrouiller l'écrasement :",
                placeholder="Ex: TEST2",
                key="input_code_confirm_reset"
            )

            expected_code = str(st.session_state.get("code_session", "")).strip().upper()
            entered_code = str(code_confirm_input).strip().upper()
            code_ok = (len(entered_code) > 0 and entered_code == expected_code)

            col_c1, col_c2 = st.columns([1, 1])
            with col_c1:
                if st.button("🔴 Confirmer et Écraser pour un Nouveau Tournoi", type="primary", use_container_width=True, disabled=not code_ok, key="btn_confirm_reset_action"):
                    with st.spinner("Effacement des anciens combats et préparation du nouveau tournoi..."):
                        # 1. Suppression dans Supabase
                        del_params = {}
                        if c_sess_detect and c_sess_detect not in ["FFLDA-ADMIN", "FFLDA2026"]:
                            del_params["code_organisateur"] = f"eq.{c_sess_detect}"
                        if tournoi_actif_detecte:
                            del_params["tournoi_id"] = f"eq.{tournoi_actif_detecte}"
                        supabase_request("matchs_lutte", method="DELETE", params=del_params)

                        # 2. Suppression du cache persistant et réinitialisation du session_state
                        if c_sess_detect:
                            supprimer_cache_tournoi(c_sess_detect)
                        st.session_state.pop("tournoi_actif_cache", None)
                        st.session_state["matchs_direct"] = []
                        st.session_state["excel_tournoi_base"] = None
                        st.session_state["pdf_tournoi_base"] = None
                        st.session_state["fichier_direct_injecte"] = None
                        st.session_state["nom_competition_active"] = None
                        st.session_state["confirm_reset_tournoi"] = False
                        st.session_state["tournoi_reinitialise_succes"] = True
                        st.session_state["mode_app_index"] = 0

                        # Purge des drapeaux temporaires et caches de fichiers
                        for k in list(st.session_state.keys()):
                            if k.startswith("synced_supa_") or k.startswith("idx_actif_tapis_") or k.startswith("last_match_") or k.startswith("upstream_last_") or k.startswith("chrono_last_"):
                                st.session_state.pop(k, None)
                        st.session_state.pop("raw_inscrits_bytes", None)
                        st.session_state.pop("raw_inscrits_name", None)
                        st.session_state.pop("raw_arbitres_bytes", None)
                        st.session_state.pop("raw_arbitres_name", None)
                        st.session_state.pop("upload_inscrits", None)
                        st.session_state.pop("upload_arbitres", None)

                    st.rerun()
            with col_c2:
                if st.button("❌ Annuler", use_container_width=True, key="btn_cancel_reset_action"):
                    st.session_state["confirm_reset_tournoi"] = False
                    st.rerun()

            if entered_code and not code_ok:
                st.error("❌ Code incorrect. Vous devez renseigner exactement votre code de connexion pour confirmer.")

        st.markdown("---")

    # Récupération proactive du bundle pour le tournoi actif
    bundle_actif = st.session_state.get("tournoi_actif_cache")
    if not bundle_actif and c_sess_detect:
        bundle_actif = charger_cache_tournoi(c_sess_detect)
        if bundle_actif:
            st.session_state["tournoi_actif_cache"] = bundle_actif
    if not bundle_actif and tournoi_actif_detecte and matchs_detectes:
        bundle_actif = creer_bundle_depuis_matchs(tournoi_actif_detecte, matchs_detectes)
        if bundle_actif:
            st.session_state["tournoi_actif_cache"] = bundle_actif

    tournoi_deja_actif = bool(tournoi_actif_detecte or bundle_actif or st.session_state.get("tournoi_actif_cache") or st.session_state.get("matchs_direct"))

    fichier_upload = None
    fichier_arbitres_upload = None

    # Si aucun tournoi n'est actif, proposer le téléversement des fichiers
    if not tournoi_deja_actif:
        col_up1, col_up2 = st.columns([1, 1])
        with col_up1:
            fichier_upload = st.file_uploader("📂 Importez votre liste d'inscrits (.csv ou .xlsx)", type=["xlsx", "csv"], key="upload_inscrits")
        with col_up2:
            fichier_arbitres_upload = st.file_uploader("(Optionnel) Importez la liste des arbitres (.xlsx ou .csv)", type=["xlsx", "csv"], key="upload_arbitres")

    # Mémorisation des octets bruts pour régénération réactive en cas de modification des paramètres
    if fichier_upload is not None:
        st.session_state["raw_inscrits_bytes"] = fichier_upload.getvalue()
        st.session_state["raw_inscrits_name"] = fichier_upload.name
    if fichier_arbitres_upload is not None:
        st.session_state["raw_arbitres_bytes"] = fichier_arbitres_upload.getvalue()
        st.session_state["raw_arbitres_name"] = fichier_arbitres_upload.name

    source_file = None
    if fichier_upload is not None:
        source_file = fichier_upload
    elif st.session_state.get("raw_inscrits_bytes"):
        import io
        source_file = io.BytesIO(st.session_state["raw_inscrits_bytes"])
        source_file.name = st.session_state.get("raw_inscrits_name", "inscrits.xlsx")
    elif bundle_actif and bundle_actif.get("raw_inscrits_bytes"):
        import io
        st.session_state["raw_inscrits_bytes"] = bundle_actif["raw_inscrits_bytes"]
        st.session_state["raw_inscrits_name"] = bundle_actif.get("raw_inscrits_name", "inscrits.xlsx")
        source_file = io.BytesIO(bundle_actif["raw_inscrits_bytes"])
        source_file.name = bundle_actif.get("raw_inscrits_name", "inscrits.xlsx")

    source_arbitres = None
    if fichier_arbitres_upload is not None:
        source_arbitres = fichier_arbitres_upload
    elif st.session_state.get("raw_arbitres_bytes"):
        import io
        source_arbitres = io.BytesIO(st.session_state["raw_arbitres_bytes"])
        source_arbitres.name = st.session_state.get("raw_arbitres_name", "arbitres.xlsx")
    elif bundle_actif and bundle_actif.get("raw_arbitres_bytes"):
        import io
        st.session_state["raw_arbitres_bytes"] = bundle_actif["raw_arbitres_bytes"]
        st.session_state["raw_arbitres_name"] = bundle_actif.get("raw_arbitres_name", "arbitres.xlsx")
        source_arbitres = io.BytesIO(bundle_actif["raw_arbitres_bytes"])
        source_arbitres.name = bundle_actif.get("raw_arbitres_name", "arbitres.xlsx")

    c_sess_courant = st.session_state.get("code_session", "ORG")
    sig_upstream_key = f"upstream_last_applied_sig_{c_sess_courant}"

    raw_bytes = st.session_state.get("raw_inscrits_bytes")
    upstream_sig = (
        str(st.session_state.get("raw_inscrits_name", "")),
        int(len(raw_bytes)) if raw_bytes else 0,
        str(nom_competition),
        int(nb_tapis),
        bool(mixte_active),
        bool(poules_par_niveau),
        bool(separer_clubs),
        bool(eviter_arbitre_meme_club),
        bool(meme_tapis_poule),
        int(tolerance_poids),
        int(repos_matchs),
        int(duree_plateau_u7),
        int(duree_u9),
        int(duree_u11),
        int(duree_u13),
    )

    param_amont_change = False
    if source_file is not None:
        if sig_upstream_key not in st.session_state:
            if tournoi_deja_actif and st.session_state.get("tournoi_actif_cache"):
                st.session_state[sig_upstream_key] = upstream_sig
                param_amont_change = False
            else:
                param_amont_change = True
        elif st.session_state[sig_upstream_key] != upstream_sig:
            param_amont_change = True
        elif not st.session_state.get("tournoi_actif_cache"):
            param_amont_change = True

    if source_file is not None and param_amont_change:
        st.session_state[sig_upstream_key] = upstream_sig
        try:
            liste_arbitres = charger_liste_arbitres(source_arbitres)
            tapis_arbitres = {t: [] for t in range(nb_tapis)}
            if liste_arbitres:
                for idx_arb, arb in enumerate(liste_arbitres):
                    tapis_arbitres[idx_arb % nb_tapis].append(arb)
            if source_file.name.endswith('.csv'):
                try:
                    source_file.seek(0)
                    df_raw = pd.read_csv(source_file, sep=';', encoding='utf-8')
                    if len(df_raw.columns) == 1:
                        source_file.seek(0)
                        df_raw = pd.read_csv(source_file, sep=',', encoding='utf-8')
                except UnicodeDecodeError:
                    source_file.seek(0)
                    df_raw = pd.read_csv(source_file, sep=';', encoding='latin-1')
                    if len(df_raw.columns) == 1:
                        source_file.seek(0)
                        df_raw = pd.read_csv(source_file, sep=',', encoding='latin-1')
            else:
                try:
                    source_file.seek(0)
                    df_temp = pd.read_excel(source_file, nrows=10, header=None)
                    header_row = 0
                    for i, row in df_temp.iterrows():
                        row_strs = [str(v).strip().lower() for v in row.values if pd.notna(v)]
                        row_full = " ".join(row_strs)
                        kw_matches = sum(1 for kw in ['licence', 'nom', 'club', 'catégorie', 'categorie', 'poids', 'sexe', 'style', 'maîtrise', 'maitrise', 'age'] if kw in row_full)
                        if kw_matches >= 2:
                            header_row = i
                            break
                    source_file.seek(0)
                    df_raw = pd.read_excel(source_file, header=header_row)
                except Exception:
                    source_file.seek(0)
                    df_raw = pd.read_excel(source_file)

            # Nettoyage des noms de colonnes (espaces superflus)
            df_raw.columns = [str(c).strip() for c in df_raw.columns]

            # 1. Catégorie d'âge / Age
            age_col_found = None
            for col_name in df_raw.columns:
                col_str = str(col_name).strip().lower()
                if any(k in col_str for k in ["catégorie d'âge", "categorie d'age", "catégorie d age", "categorie d age", "catégorie", "categorie", "age"]):
                    age_col_found = col_name
                    break
            if age_col_found:
                df_raw = df_raw.rename(columns={age_col_found: "Age"})
            elif "Age" not in df_raw.columns:
                df_raw["Age"] = ""

            # 2. Club / Sigle du Club (priorité absolue au NOM/SIGLE et rejet des numéros de club)
            club_col_found = None
            numeric_club_cols = []
            
            for col_name in df_raw.columns:
                col_lower = str(col_name).strip().lower()
                if any(n in col_lower for n in ["n°", "num", "code", "id", "numéro", "numero"]):
                    continue
                if any(k in col_lower for k in ["sigle du club", "sigle club", "sigle", "nom du club", "nom club", "libellé club", "libelle club", "nom structure", "club"]):
                    # Vérification si les données de la colonne sont des chiffres purs (ex: 1224012)
                    sample_vals = [str(v).strip() for v in df_raw[col_name].dropna().head(10)]
                    is_numeric_col = len(sample_vals) > 0 and all(v.isdigit() for v in sample_vals)
                    if not is_numeric_col:
                        club_col_found = col_name
                        break
                    else:
                        numeric_club_cols.append(col_name)

            if not club_col_found:
                for col_name in df_raw.columns:
                    col_lower = str(col_name).strip().lower()
                    if any(n in col_lower for n in ["n°", "num", "code", "id", "numéro", "numero"]):
                        continue
                    if any(k in col_lower for k in ["club", "équipe", "equipe", "structure"]) and col_name not in numeric_club_cols:
                        sample_vals = [str(v).strip() for v in df_raw[col_name].dropna().head(10)]
                        is_numeric_col = len(sample_vals) > 0 and all(v.isdigit() for v in sample_vals)
                        if not is_numeric_col:
                            club_col_found = col_name
                            break

            # Fallback si seule une colonne numérique existait
            if not club_col_found and numeric_club_cols:
                club_col_found = numeric_club_cols[0]

            if not club_col_found:
                for col_name in df_raw.columns:
                    col_lower = str(col_name).strip().lower()
                    if "club" in col_lower or "structure" in col_lower:
                        club_col_found = col_name
                        break

            if club_col_found:
                df_raw = df_raw.rename(columns={club_col_found: "Club"})
            elif "Club" not in df_raw.columns:
                df_raw["Club"] = ""

            # 3. Licence
            licence_col_found = None
            for col_name in df_raw.columns:
                col_str = str(col_name).strip().lower()
                if any(k in col_str for k in ["licence", "n° licence", "num_licence", "n°licence", "numéro licence", "numero licence"]):
                    licence_col_found = col_name
                    break
            if licence_col_found:
                df_raw = df_raw.rename(columns={licence_col_found: "Licence"})
            elif "Licence" not in df_raw.columns:
                df_raw["Licence"] = ""

            # 4. Comité
            comite_col_found = None
            for col_name in df_raw.columns:
                col_str = str(col_name).strip()
                col_lower = col_str.lower()
                if any(k in col_lower for k in ["comité", "comite", "ligue", "région", "region", "c.r."]):
                    comite_col_found = col_name
                    break
            if comite_col_found:
                df_raw = df_raw.rename(columns={comite_col_found: "Comité"})
            if "Comité" not in df_raw.columns:
                df_raw["Comité"] = "Comité Non Renseigné"

            # 5. Style
            style_col_found = None
            for col_name in df_raw.columns:
                col_str = str(col_name).strip().lower()
                if any(k in col_str for k in ["style", "discipline", "gréco", "greco", "gr/ll"]):
                    style_col_found = col_name
                    break
            if style_col_found:
                df_raw = df_raw.rename(columns={style_col_found: "Style"})
            if "Style" not in df_raw.columns:
                df_raw["Style"] = ""

            # 6. Maîtrise / Niveau
            maitrise_col_found = None
            for col_name in df_raw.columns:
                col_str = str(col_name).strip().lower()
                if any(k in col_str for k in ["maîtrise", "maitrise", "niveau", "expéri", "experi", "niv"]):
                    maitrise_col_found = col_name
                    break
            if maitrise_col_found:
                df_raw = df_raw.rename(columns={maitrise_col_found: "Maîtrise"})
            if "Maîtrise" not in df_raw.columns:
                df_raw["Maîtrise"] = ""

            # 7. Sexe
            sexe_col_found = None
            for col_name in df_raw.columns:
                col_str = str(col_name).strip().lower()
                if any(k in col_str for k in ["sexe", "genre", "civilité", "civilite", "m/f"]):
                    sexe_col_found = col_name
                    break
            if sexe_col_found:
                df_raw = df_raw.rename(columns={sexe_col_found: "Sexe"})
            if "Sexe" not in df_raw.columns:
                df_raw["Sexe"] = ""

            # 8. Poids (Recherche robuste : Poids, Poids (kg), Poids pesée, Poids mesuré, Weight, etc.)
            poids_col_found = None
            for col_name in df_raw.columns:
                col_str = str(col_name).strip().lower()
                if any(k in col_str for k in ["poids", "weight", "pesée", "pesee", "mesuré", "mesure"]):
                    poids_col_found = col_name
                    break
            if poids_col_found:
                df_raw = df_raw.rename(columns={poids_col_found: "Poids"})
            elif "Poids" not in df_raw.columns:
                df_raw["Poids"] = ""

            # 9. Nom et Prénom
            nom_col_found = None
            prenom_col_found = None
            for col_name in df_raw.columns:
                col_str = str(col_name).strip().lower()
                if any(k in col_str for k in ["prénom", "prenom", "first name", "firstname"]):
                    prenom_col_found = col_name
                elif col_str in ["nom", "nom du lutteur", "nom athlète", "nom athlete", "nom complet", "athlète", "athlete"]:
                    nom_col_found = col_name
                elif "nom" in col_str and "club" not in col_str and not nom_col_found:
                    nom_col_found = col_name

            if nom_col_found and nom_col_found != "Nom":
                df_raw = df_raw.rename(columns={nom_col_found: "Nom"})
            if prenom_col_found and prenom_col_found != "Prénom":
                df_raw = df_raw.rename(columns={prenom_col_found: "Prénom"})

            if "Nom" not in df_raw.columns:
                df_raw["Nom"] = ""

            if "Prénom" in df_raw.columns:
                def combiner_nom_prenom(row):
                    n = str(row.get("Nom", "")).strip()
                    p = str(row.get("Prénom", "")).strip()
                    n = "" if n.lower() in ["nan", "none"] else n
                    p = "" if p.lower() in ["nan", "none"] else p
                    if n and p:
                        if p.lower() in n.lower():
                            return n
                        return f"{n} {p}"
                    return n or p
                df_raw["Nom"] = df_raw.apply(combiner_nom_prenom, axis=1)

            def normaliser_age(val):
                if val is None or pd.isna(val):
                    return ''
                val_str = str(val).strip().upper().replace(' ', '').replace('-', '')
                if 'U7' in val_str or val_str == '7' or val_str.startswith('7'):
                    return 'U7'
                elif 'U9' in val_str or val_str == '9' or val_str.startswith('9'):
                    return 'U9'
                elif 'U11' in val_str or val_str == '11' or val_str.startswith('11'):
                    return 'U11'
                elif 'U13' in val_str or val_str == '13' or val_str.startswith('13'):
                    return 'U13'
                return str(val).strip().upper()

            df_inscr_total = df_raw.copy()
            if "Age" in df_inscr_total.columns:
                df_inscr_total['Age'] = df_inscr_total['Age'].apply(normaliser_age)
                df_inscr_total = df_inscr_total[df_inscr_total['Age'].isin(['U7', 'U9', 'U11', 'U13'])]
            
            total_inscrits_global = len(df_inscr_total)

            def attribuer_niveau(val):
                val_str = str(val).strip().lower()
                if not val_str or val_str in ['nan', 'none', '', '-']:
                    return 'Débutant'
                if any(k in val_str for k in ['confirm', 'expert', 'avanc', 'haut']) or val_str.startswith('c') or val_str in ['2', 'n2', 'niv 2', 'niveau 2']:
                    return 'Confirmé'
                elif any(k in val_str for k in ['début', 'debut', 'novice', 'initia']) or val_str.startswith('d') or val_str in ['1', 'n1', 'niv 1', 'niveau 1']:
                    return 'Débutant'
                else:
                    return 'Débutant'
            
            if poules_par_niveau:
                df_inscr_total['Niveau'] = df_inscr_total['Maîtrise'].apply(attribuer_niveau)
            else:
                df_inscr_total['Niveau'] = ''

            # Nettoyage et conversion du poids (sécurisé contre KeyError et valeurs invalides)
            if 'Poids' not in df_inscr_total.columns:
                df_inscr_total['Poids'] = ''
            df_inscr_total['Poids_Clean'] = df_inscr_total['Poids'].astype(str).str.replace(',', '.').str.strip()
            df_inscr_total['Poids_Num'] = pd.to_numeric(df_inscr_total['Poids_Clean'], errors='coerce').fillna(0)
            df_inscr_total['Poids'] = df_inscr_total['Poids_Num'].apply(formater_poids)

            # --- DÉTECTION ET VALIDATION DES LICENCES MANQUANTES ---
            erreurs_licence = []
            for _, row_test in df_inscr_total.iterrows():
                nom_raw = str(row_test.get('Nom', '')).strip()
                if not nom_raw or nom_raw.lower() in ['nan', 'nan nan', 'none', '']:
                    continue
                
                club_lutteur = str(row_test.get('Club', '')).strip()
                lic_val = str(row_test.get('Licence', '')).strip().lower()
                
                if lic_val in ['', 'nan', 'none', 'null', '-', '0', 'unspecified', 'inconnu'] or pd.isna(row_test.get('Licence')):
                    club_txt = f" ({club_lutteur})" if club_lutteur and club_lutteur.lower() not in ['nan', 'none', '-'] else ""
                    erreurs_licence.append(f"• **{nom_raw}**{club_txt} : Le numéro de licence n'est pas renseigné.")

            if erreurs_licence:
                st.error("❌ **Erreur d'importation dans le fichier de base (Licences manquantes) :**\n\n" + "\n".join(erreurs_licence))
                st.info("💡 *Remarque : Tous les lutteurs enregistrés dans le fichier de base doivent posséder un numéro de licence valide avant de générer la compétition.*")
                st.stop()

            # Fonctions utilitaires d'identification du Sexe et du Style
            def est_sexe_feminin(val):
                v = str(val).strip().upper()
                if not v or v in ['NAN', 'NONE', '']: return False
                return v in ['F', 'FEMME', 'FILLE', 'FEMININ', 'FÉMININ', 'FEM', 'WOMAN', 'GIRL'] or v.startswith('FÉM') or v.startswith('FEM')

            def est_sexe_masculin(val):
                v = str(val).strip().upper()
                if not v or v in ['NAN', 'NONE', '']: return False
                return v in ['M', 'H', 'HOMME', 'GARCON', 'GARÇON', 'MASCULIN', 'MASC', 'MAN', 'BOY'] or v.startswith('MASC') or v.startswith('HOM')

            def est_style_greco(val):
                s = str(val).strip().upper()
                if not s or s in ['NAN', 'NONE', '']: return False
                return any(k in s for k in ['LG', 'GR', 'GRECO', 'GRÉCO', 'ROMAIN']) or s == 'G'

            # --- VALIDATION : INCOMPATIBILITÉ SEXE FÉMININ & LUTTE GRÉCO-ROMAINE (LG/GR) ---
            erreurs_sexe_greco = []
            for _, row_test in df_inscr_total.iterrows():
                nom_lutteur = str(row_test.get('Nom', 'Lutteur Inconnu')).strip()
                if not nom_lutteur or nom_lutteur.lower() in ['nan', 'nan nan', 'none', '']:
                    continue
                
                sexe_val = row_test.get('Sexe', '')
                style_val = row_test.get('Style', '')
                
                if est_sexe_feminin(sexe_val) and est_style_greco(style_val):
                    club_lutteur = str(row_test.get('Club', '')).strip()
                    club_txt = f" ({club_lutteur})" if club_lutteur and club_lutteur.lower() not in ['nan', 'none', '-'] else ""
                    erreurs_sexe_greco.append(f"• **{nom_lutteur}**{club_txt} est de sexe féminin et ne peut pas être référencé en LG.")

            if erreurs_sexe_greco:
                st.error("❌ **Erreur d'importation dans le fichier :**\n\n" + "\n".join(erreurs_sexe_greco))
                st.info("💡 *Remarque : La lutte gréco-romaine (LG / GR) ne s'applique pas aux féminines. Une lutteuse de sexe féminin est automatiquement orientée en LF et ne peut pas être inscrite en LG.*")
                st.stop()

            # --- DÉTECTION ET VALIDATION DU STYLE "JEUNE" SANS SEXE IDENTIFIÉ ---
            erreurs_jeune = []
            for _, row_test in df_inscr_total.iterrows():
                val_style_raw = str(row_test.get('Style', '')).strip().lower()
                poids_val_raw = row_test.get('Poids_Num')
                try:
                    poids_val = float(poids_val_raw) if (poids_val_raw is not None and str(poids_val_raw).strip() != '') else 0.0
                except (ValueError, TypeError):
                    poids_val = 0.0
                nom_lutteur = row_test.get('Nom', 'Lutteur Inconnu')
                club_lutteur = row_test.get('Club', '')
                sexe_val = row_test.get('Sexe', '')
                
                # Si le style est "jeune" mais qu'aucun sexe (F ou M) n'est présent pour déduire LF ou LL
                if val_style_raw == 'jeune' and not (est_sexe_feminin(sexe_val) or est_sexe_masculin(sexe_val)):
                    if poids_val > 0:
                        club_txt = f" ({club_lutteur})" if club_lutteur and str(club_lutteur).lower() not in ['nan', 'none', '-'] else ""
                        erreurs_jeune.append(f"• **{nom_lutteur}**{club_txt} : Poids renseigné (**{poids_val} kg**) avec le style '**jeune**' sans colonne Sexe renseignée. Veuillez renseigner le sexe (M/F) ou le style (LL/LF/LG).")

            if erreurs_jeune:
                st.error("❌ **Erreur d'importation dans le fichier :**\n\n" + "\n".join(erreurs_jeune))
                st.info("💡 *Remarque : Le style 'jeune' nécessite d'indiquer le sexe dans la colonne Sexe/Genre (M ou F) pour orienter automatiquement le lutteur en LL ou LF, ou de corriger directement le style en LL, LF ou LG.*")
                st.stop()

            # RÈGLES DE NORMALISATION DES STYLES :
            # 1. Le sexe féminin est TOUJOURS identifié LF peu importe ce qu'il y a marqué dans la colonne style
            # 2. Le sexe masculin est orienté LL automatiquement peu importe ce qu'il y a marqué dans la colonne style, à part si marqué LG, GR ou gréco
            # 3. Si sexe non renseigné : repli sur la colonne style
            def normaliser_style(row_p):
                sexe_val = row_p.get('Sexe', '')
                style_val = row_p.get('Style', '')
                
                if est_sexe_feminin(sexe_val):
                    return 'LF'
                
                if est_sexe_masculin(sexe_val):
                    if est_style_greco(style_val):
                        return 'LG'
                    return 'LL'
                    
                if est_style_greco(style_val):
                    return 'LG'
                st_upper = str(style_val).strip().upper()
                if any(k in st_upper for k in ['LF', 'FEM', 'FÉM', 'FILLE']) or st_upper == 'F':
                    return 'LF'
                return 'LL'

            df_inscr_total['Style_Norm'] = df_inscr_total.apply(normaliser_style, axis=1)

            # Conserver tous les participants pesés (Poids_Num > 0)
            df_inscr = df_inscr_total[df_inscr_total['Poids_Num'] > 0].copy()

            total_participants_peses = len(df_inscr)
            total_non_peses = total_inscrits_global - total_participants_peses

            if df_inscr.empty:
                st.error("❌ Aucun lutteur U7, U9, U11 ou U13 avec un poids valide et un style de compétition n'a été trouvé dans le fichier.")
                st.stop()
            
            # Définition des groupes de styles selon le réglage de mixité (mixité interdite en U13 selon le règlement officiel FFLDA)
            def attribuer_style_groupe(row_p):
                style_norm = row_p['Style_Norm']
                age = str(row_p.get('Age', '')).strip().upper()
                
                # RÈGLEMENT FFLDA : En U13, il n'existe plus de catégorie mixte !
                # Filles et garçons sont obligatoirement séparés : LF (Féminine), LL (Libre) ou LG (Gréco)
                if age == 'U13':
                    if style_norm == 'LG':
                        return 'LG (Gréco)'
                    elif style_norm == 'LF':
                        return 'LF (Féminine)'
                    else:
                        return 'LL (Libre)'

                # Pour U7, U9, U11 : la mixité dépend du réglage mixte_active
                if mixte_active:
                    if style_norm == 'LG':
                        return 'LG (Gréco)'
                    else:
                        return 'Mixte (LL/LF)'
                else:
                    if style_norm == 'LG':
                        return 'LG (Gréco)'
                    elif style_norm == 'LF':
                        return 'LF (Féminine)'
                    else:
                        return 'LL (Libre)'

            df_inscr['Style_Groupe'] = df_inscr.apply(attribuer_style_groupe, axis=1)
            
            poules_u7, poules_u9, poules_u11, poules_u13 = [], [], [], []
            multiplicateur_poids = 1 + (tolerance_poids / 100.0)
            
            # --- GÉNÉRATION SPÉCIFIQUE U7 : 3 GROUPES HOMOGÈNES EN PLATEAUX D'ACTIVITÉ ---
            df_u7_total = df_inscr[df_inscr['Age'] == 'U7'].sort_values('Poids_Num')
            if not df_u7_total.empty:
                u7_all = df_u7_total.to_dict('records')
                nb_lutteurs_u7 = len(u7_all)
                # Exactement 3 groupes homogènes en effectif (ou moins si < 3 lutteurs au total)
                nb_groupes_u7 = 3 if nb_lutteurs_u7 >= 3 else max(1, nb_lutteurs_u7)
                
                base_len = nb_lutteurs_u7 // nb_groupes_u7
                reste = nb_lutteurs_u7 % nb_groupes_u7
                start_i = 0
                for g_i in range(nb_groupes_u7):
                    taille = base_len + (1 if g_i < reste else 0)
                    parts_g = u7_all[start_i : start_i + taille]
                    start_i += taille
                    
                    p_min = parts_g[0]['Poids_Num']
                    p_max = parts_g[-1]['Poids_Num']
                    p_min_str = formater_poids_court(p_min)
                    p_max_str = formater_poids_court(p_max)
                    poids_range_str = f" ({p_min_str} - {p_max_str})" if p_min_str and p_max_str else ""
                    
                    nom_g = f"U7 | Plateau - Groupe {g_i + 1} ({len(parts_g)} lutteurs){poids_range_str}"
                    poule_obj = {
                        'nom': nom_g,
                        'participants': parts_g,
                        'rondes': [],  # Animation sous forme de plateaux (pas de combats éliminatoires)
                        'type_formule': 'plateau_u7',
                        'num_groupe': g_i + 1
                    }
                    poules_u7.append(poule_obj)

            # --- GÉNÉRATION U9, U11 (Poules morphologiques à tolérance de poids) ---
            for age in ['U9', 'U11']:
                df_age = df_inscr[df_inscr['Age'] == age].sort_values('Poids_Num')
                max_size = 4 if age == 'U9' else 5
                counter_gr = 1  # Numérotation continue des groupes pour la catégorie d'âge (ex: Poule 1 à 14)
                
                for (style_grp, niveau), groupe in df_age.groupby(['Style_Groupe', 'Niveau']):
                    participants = groupe.to_dict('records')
                    poule_courante = []
                    suffixe_niveau = f" | {niveau}" if niveau != "" else ""
                    poules_groupe = []
                    
                    for p in participants:
                        if not poule_courante:
                            poule_courante.append(p)
                        else:
                            poids_min = float(poule_courante[0].get('Poids_Num') or 0.0)
                            p_poids = float(p.get('Poids_Num') or 0.0)
                            if p_poids <= (poids_min * multiplicateur_poids) and len(poule_courante) < max_size:
                                poule_courante.append(p)
                            else:
                                nom_groupe = f"{age} | {style_grp}{suffixe_niveau} ({formater_poids_court(poule_courante[0]['Poids_Num'])} - {formater_poids_court(poule_courante[-1]['Poids_Num'])})"
                                poule_obj = {'nom': nom_groupe, 'participants': list(poule_courante), 'rondes': generer_rondes_fflda(poule_courante), 'type_formule': 'poule'}
                                poules_groupe.append(poule_obj)
                                poule_courante = [p]
                    if poule_courante:
                        nom_groupe = f"{age} | {style_grp}{suffixe_niveau} ({formater_poids_court(poule_courante[0]['Poids_Num'])} - {formater_poids_court(poule_courante[-1]['Poids_Num'])})"
                        poule_obj = {'nom': nom_groupe, 'participants': list(poule_courante), 'rondes': generer_rondes_fflda(poule_courante), 'type_formule': 'poule'}
                        poules_groupe.append(poule_obj)
                    
                    if separer_clubs:
                        poules_groupe = optimiser_poules_clubs(poules_groupe, multiplicateur_poids)

                    # Eviter les poules de 1 en les fusionnant / rééquilibrant tout en respectant la tolérance de poids
                    poules_groupe = fusionner_poules_isolees(poules_groupe, multiplicateur_poids, max_size)
                    
                    if separer_clubs:
                        poules_groupe = optimiser_poules_clubs(poules_groupe, multiplicateur_poids)
                    
                    # Formater et numéroter en CONTINU par catégorie d'âge (ex: Poule 1 à Poule 14)
                    for p_obj in poules_groupe:
                        parts = p_obj['participants']
                        p_min = parts[0]['Poids_Num']
                        p_max = parts[-1]['Poids_Num']
                        p_obj['nom'] = f"{age} | {style_grp}{suffixe_niveau} | Poule {counter_gr} ({formater_poids_court(p_min)} - {formater_poids_court(p_max)})"
                        p_obj['rondes'] = generer_rondes_fflda(parts)
                        p_obj['type_formule'] = 'poule'
                        counter_gr += 1
                    
                    if age == 'U9': poules_u9.extend(poules_groupe)
                    else: poules_u11.extend(poules_groupe)

            # --- GÉNÉRATION U13 (Catégories de poids fixes : 30, 33, 36, 39, 42, 46, 50, 55, 60, +60 kg) ---
            df_u13_total = df_inscr[df_inscr['Age'] == 'U13'].copy()
            if not df_u13_total.empty:
                df_u13_total['Cat_Poids'] = df_u13_total['Poids_Num'].apply(attribuer_categorie_poids_u13)
                for (style_grp, niveau), groupe_style in df_u13_total.groupby(['Style_Groupe', 'Niveau']):
                    suffixe_niveau = f" | {niveau}" if niveau != "" else ""
                    for cat_poids in ORDRE_POIDS_U13:
                        groupe_cat = groupe_style[groupe_style['Cat_Poids'] == cat_poids]
                        if groupe_cat.empty:
                            continue
                        participants_cat = groupe_cat.to_dict('records')
                        comp_obj = generer_competition_u13('U13', style_grp, suffixe_niveau, cat_poids, participants_cat, separer_clubs)
                        if comp_obj:
                            poules_u13.append(comp_obj)

            tous_les_groupes = poules_u7 + poules_u9 + poules_u11 + poules_u13
            participants_par_poule = {p['nom']: p['participants'] for p in tous_les_groupes}
            rondes_par_categorie = {p['nom']: p['rondes'] for p in tous_les_groupes}
            poule_obj_map = {p['nom']: p for p in tous_les_groupes}

            tapis_poules_u7 = {i: [] for i in range(nb_tapis)}
            tapis_poules_u9 = {i: [] for i in range(nb_tapis)}
            tapis_poules_u11 = {i: [] for i in range(nb_tapis)}
            tapis_poules_u13 = {i: [] for i in range(nb_tapis)}
            
            for i, p in enumerate(poules_u7): tapis_poules_u7[i % nb_tapis].append(p)
            for i, p in enumerate(poules_u9): tapis_poules_u9[i % nb_tapis].append(p)
            for i, p in enumerate(poules_u11): tapis_poules_u11[i % nb_tapis].append(p)
            for i, p in enumerate(poules_u13): tapis_poules_u13[i % nb_tapis].append(p)

            h_p1_eff = st.session_state.get("downstream_h_pesee", heure_pesee_u9)
            d_p1_eff = st.session_state.get("downstream_duree_pesee", duree_pesee)
            dt_pesee_u9 = datetime.combine(datetime.today(), h_p1_eff)
            dt_debut_u7 = dt_pesee_u9 + timedelta(minutes=d_p1_eff)
            total_matchs_calcules = 0

            ref_usage_count = {}

            def choisir_arbitre_match(c1_club, c2_club, t_idx):
                cand_list = tapis_arbitres.get(t_idx, [])
                if not cand_list and liste_arbitres:
                    cand_list = liste_arbitres
                
                if not cand_list:
                    return "Non attribué"
                
                c1_clean = str(c1_club).strip().lower()
                c2_clean = str(c2_club).strip().lower()
                ignored_clubs = ['', '-', 'indépendant', 'independant', 'none', 'nan']
                
                if eviter_arbitre_meme_club:
                    sans_conflit = []
                    for arb in cand_list:
                        arb_club_clean = str(arb.get('Club', '')).strip().lower()
                        if arb_club_clean in ignored_clubs:
                            sans_conflit.append(arb)
                        elif arb_club_clean != c1_clean and arb_club_clean != c2_clean:
                            sans_conflit.append(arb)
                    pool_choix = sans_conflit if sans_conflit else cand_list
                else:
                    pool_choix = cand_list
                        
                arb_choisi = min(pool_choix, key=lambda a: ref_usage_count.get(a['Nom_Complet'], 0))
                
                ref_usage_count[arb_choisi['Nom_Complet']] = ref_usage_count.get(arb_choisi['Nom_Complet'], 0) + 1
                return arb_choisi['Nom_Complet']

            def nommer_tour_match(poule_obj, r_idx, m_idx, m=None):
                if not poule_obj:
                    return f"Tour {r_idx + 1}"
                type_f = poule_obj.get('type_formule', 'poule')
                rondes = poule_obj.get('rondes', [])
                total_r = len(rondes)
                
                if type_f == 'tableau':
                    if total_r >= 6:
                        if r_idx == 0:
                            return f"1/32 de Finale ({m_idx + 1})"
                        elif r_idx == 1:
                            return f"1/16 de Finale ({m_idx + 1})"
                        elif r_idx == 2:
                            return f"1/8 de Finale ({m_idx + 1})"
                        elif r_idx == 3:
                            return f"1/4 de Finale ({m_idx + 1})"
                        elif r_idx == total_r - 2:
                            if m_idx in (0, 1):
                                return f"Demi-Finale {m_idx + 1}"
                            else:
                                return f"Repêchage 1/4 ({m_idx - 1})"
                        elif r_idx == total_r - 1:
                            if m_idx == 0:
                                return "Grande Finale (Or / Argent)"
                            elif m_idx == 1:
                                return "Finale Bronze 1"
                            elif m_idx == 2:
                                return "Finale Bronze 2"
                            else:
                                return f"Finale {m_idx + 1}"
                        else:
                            return f"Tour {r_idx + 1}"
                    elif total_r == 5:
                        if r_idx == 0:
                            return f"1/16 de Finale ({m_idx + 1})"
                        elif r_idx == 1:
                            return f"1/8 de Finale ({m_idx + 1})"
                        elif r_idx == 2:
                            return f"1/4 de Finale ({m_idx + 1})"
                        elif r_idx == total_r - 2:
                            if m_idx in (0, 1):
                                return f"Demi-Finale {m_idx + 1}"
                            else:
                                return f"Repêchage 1/4 ({m_idx - 1})"
                        elif r_idx == total_r - 1:
                            if m_idx == 0:
                                return "Grande Finale (Or / Argent)"
                            elif m_idx == 1:
                                return "Finale Bronze 1"
                            elif m_idx == 2:
                                return "Finale Bronze 2"
                            else:
                                return f"Finale {m_idx + 1}"
                        else:
                            return f"Tour {r_idx + 1}"
                    elif total_r == 4:
                        if r_idx == 0:
                            return f"Tour Préliminaire 1/8 ({m_idx + 1})"
                        elif r_idx == 1:
                            return f"1/4 de Finale ({m_idx + 1})"
                        elif r_idx == 2:
                            if m_idx in (0, 1):
                                return f"Demi-Finale {m_idx + 1}"
                            else:
                                return f"Repêchage 1/4 ({m_idx - 1})"
                        elif r_idx == 3:
                            if m_idx == 0:
                                return "Grande Finale (Or / Argent)"
                            elif m_idx == 1:
                                return "Finale Bronze 1"
                            elif m_idx == 2:
                                return "Finale Bronze 2"
                            else:
                                return f"Finale {m_idx + 1}"
                        else:
                            return f"Tour {r_idx + 1}"
                    elif total_r == 3:
                        if r_idx == 0:
                            return f"1/4 de Finale ({m_idx + 1})"
                        elif r_idx == 1:
                            if m_idx in (0, 1):
                                return f"Demi-Finale {m_idx + 1}"
                            else:
                                return f"Repêchage 1/4 ({m_idx - 1})"
                        elif r_idx == 2:
                            if m_idx == 0:
                                return "Grande Finale (Or / Argent)"
                            elif m_idx == 1:
                                return "Finale Bronze 1"
                            elif m_idx == 2:
                                return "Finale Bronze 2"
                            else:
                                return f"Finale {m_idx + 1}"
                        else:
                            return f"Tour {r_idx + 1}"
                    else:
                        return f"Tour {r_idx + 1}"
                        
                elif type_f == 'poules_croisees':
                    if r_idx == 4:
                        if m_idx == 0:
                            return "Grande Finale (Or / Argent)"
                        else:
                            return "Finale 3-4 (Bronze unique)"
                    elif r_idx == 3:
                        return f"Demi-Finale Croisée {m_idx + 1}"
                    else:
                        return f"Poule - Tour {r_idx + 1}"
                        
                else:
                    return f"Tour {r_idx + 1}"

            def ordonnancer_phase(poules_phase, heure_debut_phase, duree_combat, meme_tapis=True):
                global total_matchs_calcules
                if not poules_phase:
                    return [heure_debut_phase for _ in range(nb_tapis)], {t: [] for t in range(nb_tapis)}
                
                planning = {t: [] for t in range(nb_tapis)}
                tapis_heure = [heure_debut_phase for _ in range(nb_tapis)]
                last_match_time = {}

                if meme_tapis:
                    # RÈGLE OFFICIELLE : Chaque poule de même style et même catégorie de poids (ex: U13 Gréco 30 kg) est affectée au même tapis fixe.
                    # Une même catégorie de poids d'un autre style (ex: U13 Libre 30 kg ou U13 Féminine 30 kg) peut aller sur un tapis différent,
                    # ce qui permet un équilibrage encore plus optimal du nombre de matchs par tapis (Algorithme LPT).
                    def obtenir_cle_poids_lot(p_obj):
                        nom_p = p_obj.get('nom', '')
                        parts_nom = [x.strip() for x in nom_p.split('|')]
                        age_p = parts_nom[0] if parts_nom else 'U13'
                        style_p = p_obj.get('style_grp') or (parts_nom[1] if len(parts_nom) >= 2 else '')
                        cat_p = p_obj.get('cat_poids')
                        
                        if cat_p:
                            return f"{age_p} | {style_p} | {str(cat_p).strip()}"
                        
                        m_poids = re.search(r'\b(\+?\d+\s*kg)\b', nom_p, re.IGNORECASE)
                        if m_poids:
                            return f"{age_p} | {style_p} | {m_poids.group(1).strip()}"
                        
                        # Pour U9 / U11 : chaque poule morphologique constitue son propre groupe
                        return nom_p

                    # 1. Regroupement par catégorie de poids (tous les combats de la même catégorie de poids restent ensemble)
                    lots_poids = {}
                    for p in poules_phase:
                        cle = obtenir_cle_poids_lot(p)
                        if cle not in lots_poids:
                            lots_poids[cle] = []
                        lots_poids[cle].append(p)

                    # 2. Calcul du nombre de matchs par lot de catégorie de poids
                    lots_avec_poids = []
                    for cle, liste_p in lots_poids.items():
                        nb_m = sum(sum(len(r) for r in p.get('rondes', [])) for p in liste_p)
                        lots_avec_poids.append((cle, liste_p, nb_m))

                    # 3. Tri décroissant par nombre de matchs (LPT - Longest Processing Time First)
                    lots_avec_poids.sort(key=lambda x: x[2], reverse=True)

                    # 4. Affectation équilibrée sur les tapis : chaque lot de poids va sur le tapis actuellement le moins chargé en matchs
                    tapis_poules = {t: [] for t in range(nb_tapis)}
                    charge_matchs_tapis = [0] * nb_tapis

                    for cle, liste_p, nb_m in lots_avec_poids:
                        t_min = min(range(nb_tapis), key=lambda t: (charge_matchs_tapis[t], t))
                        tapis_poules[t_min].extend(liste_p)
                        charge_matchs_tapis[t_min] += nb_m

                    for t in range(nb_tapis):
                        poules_t = tapis_poules[t]
                        if not poules_t:
                            continue
                        
                        matchs_tapis = []
                        max_rondes = max((len(p['rondes']) for p in poules_t), default=0)
                        for r in range(max_rondes):
                            for p in poules_t:
                                if r < len(p['rondes']):
                                    for m_idx, m in enumerate(p['rondes'][r]):
                                        matchs_tapis.append({
                                            'poule': p['nom'],
                                            'poule_obj': p,
                                            'p1': m[0]['Nom'],
                                            'p1_club': m[0].get('Club', ''),
                                            'p1_comite': m[0].get('Comité', m[0].get('Comite', 'Comité Non Renseigné')),
                                            'p2': m[1]['Nom'],
                                            'p2_club': m[1].get('Club', ''),
                                            'p2_comite': m[1].get('Comité', m[1].get('Comite', 'Comité Non Renseigné')),
                                            'tour': r + 1,
                                            'nom_tour': nommer_tour_match(p, r, m_idx, m)
                                        })

                        while matchs_tapis:
                            t_curr_time = tapis_heure[t]
                            
                            match_choisi_idx = None
                            for idx, m in enumerate(matchs_tapis):
                                p1, p2 = m['p1'], m['p2']
                                t_pret_p1 = last_match_time.get(p1, heure_debut_phase)
                                t_pret_p2 = last_match_time.get(p2, heure_debut_phase)
                                if t_pret_p1 <= t_curr_time and t_pret_p2 <= t_curr_time:
                                    match_choisi_idx = idx
                                    break
                            
                            if match_choisi_idx is not None:
                                m = matchs_tapis.pop(match_choisi_idx)
                                heure_debut_match = t_curr_time
                            else:
                                meilleur_idx = 0
                                meilleur_temps = datetime.max
                                for idx, m in enumerate(matchs_tapis):
                                    p1, p2 = m['p1'], m['p2']
                                    t_pret = max(last_match_time.get(p1, heure_debut_phase), last_match_time.get(p2, heure_debut_phase))
                                    if t_pret < meilleur_temps:
                                        meilleur_temps = t_pret
                                        meilleur_idx = idx
                                
                                m = matchs_tapis.pop(meilleur_idx)
                                heure_debut_match = max(t_curr_time, meilleur_temps)
                                
                                if heure_debut_match > t_curr_time:
                                    attente_min = int((heure_debut_match - t_curr_time).total_seconds() // 60)
                                    if attente_min > 0:
                                        planning[t].append({
                                            "Type": "ATTENTE", 
                                            "Heure": t_curr_time.strftime("%H:%M"), 
                                            "Texte": f"⏳ Repos ({attente_min} min)"
                                        })

                            arb_nom = choisir_arbitre_match(m['p1_club'], m['p2_club'], t)

                            planning[t].append({
                                "Type": "MATCH",
                                "Heure": heure_debut_match.strftime("%H:%M"),
                                "Duree": duree_combat,
                                "Cat": m['poule'],
                                "Combattant 1": m['p1'],
                                "Club 1": m['p1_club'],
                                "Comité 1": m['p1_comite'],
                                "Combattant 2": m['p2'],
                                "Club 2": m['p2_club'],
                                "Comité 2": m['p2_comite'],
                                "Tour": m['tour'],
                                "Nom_Tour": m.get('nom_tour', f"Tour {m['tour']}"),
                                "Arbitre": arb_nom
                            })
                            
                            total_matchs_calcules += 1
                            fin_match = heure_debut_match + timedelta(minutes=duree_combat)
                            delai_repos = timedelta(minutes=(repos_matchs * duree_combat))
                            
                            last_match_time[m['p1']] = fin_match + delai_repos
                            last_match_time[m['p2']] = fin_match + delai_repos
                            
                            tapis_heure[t] = fin_match
                else:
                    # Répartition dynamique sur le premier tapis disponible
                    matchs_a_jouer = []
                    max_rondes = max((len(p['rondes']) for p in poules_phase), default=0)
                    for r in range(max_rondes):
                        for p in poules_phase:
                            if r < len(p['rondes']):
                                for m_idx, m in enumerate(p['rondes'][r]):
                                    matchs_a_jouer.append({
                                        'poule': p['nom'],
                                        'poule_obj': p,
                                        'p1': m[0]['Nom'],
                                        'p1_club': m[0].get('Club', ''),
                                        'p1_comite': m[0].get('Comité', m[0].get('Comite', 'Comité Non Renseigné')),
                                        'p2': m[1]['Nom'],
                                        'p2_club': m[1].get('Club', ''),
                                        'p2_comite': m[1].get('Comité', m[1].get('Comite', 'Comité Non Renseigné')),
                                        'tour': r + 1,
                                        'nom_tour': nommer_tour_match(p, r, m_idx, m)
                                    })
                    
                    while matchs_a_jouer:
                        t_min_idx = min(range(nb_tapis), key=lambda t: tapis_heure[t])
                        t_min_time = tapis_heure[t_min_idx]
                        
                        match_choisi_idx = None
                        for idx, m in enumerate(matchs_a_jouer):
                            p1, p2 = m['p1'], m['p2']
                            t_pret_p1 = last_match_time.get(p1, heure_debut_phase)
                            t_pret_p2 = last_match_time.get(p2, heure_debut_phase)
                            if t_pret_p1 <= t_min_time and t_pret_p2 <= t_min_time:
                                match_choisi_idx = idx
                                break
                        
                        if match_choisi_idx is not None:
                            m = matchs_a_jouer.pop(match_choisi_idx)
                            heure_debut_match = t_min_time
                        else:
                            meilleur_idx = 0
                            meilleur_temps = datetime.max
                            for idx, m in enumerate(matchs_a_jouer):
                                p1, p2 = m['p1'], m['p2']
                                t_pret = max(last_match_time.get(p1, heure_debut_phase), last_match_time.get(p2, heure_debut_phase))
                                if t_pret < meilleur_temps:
                                    meilleur_temps = t_pret
                                    meilleur_idx = idx
                            
                            m = matchs_a_jouer.pop(meilleur_idx)
                            heure_debut_match = max(t_min_time, meilleur_temps)
                            
                            if heure_debut_match > t_min_time:
                                attente_min = int((heure_debut_match - t_min_time).total_seconds() // 60)
                                if attente_min > 0:
                                    planning[t_min_idx].append({
                                        "Type": "ATTENTE", 
                                        "Heure": t_min_time.strftime("%H:%M"), 
                                        "Texte": f"⏳ Repos ({attente_min} min)"
                                    })

                        arb_nom = choisir_arbitre_match(m['p1_club'], m['p2_club'], t_min_idx)

                        planning[t_min_idx].append({
                            "Type": "MATCH",
                            "Heure": heure_debut_match.strftime("%H:%M"),
                            "Duree": duree_combat,
                            "Cat": m['poule'],
                            "Combattant 1": m['p1'],
                            "Club 1": m['p1_club'],
                            "Comité 1": m['p1_comite'],
                            "Combattant 2": m['p2'],
                            "Club 2": m['p2_club'],
                            "Comité 2": m['p2_comite'],
                            "Tour": m['tour'],
                            "Nom_Tour": m.get('nom_tour', f"Tour {m['tour']}"),
                            "Arbitre": arb_nom
                        })
                        
                        total_matchs_calcules += 1
                        fin_match = heure_debut_match + timedelta(minutes=duree_combat)
                        delai_repos = timedelta(minutes=(repos_matchs * duree_combat))
                        
                        last_match_time[m['p1']] = fin_match + delai_repos
                        last_match_time[m['p2']] = fin_match + delai_repos
                        
                        tapis_heure[t_min_idx] = fin_match

                return tapis_heure, planning

            # PHASE 1a : Tous les U7 (Format Plateaux d'Activité : 3 rotations de duree_plateau_u7 min)
            planning_u7 = {t: [] for t in range(nb_tapis)}
            tapis_heure_u7 = [dt_debut_u7 for _ in range(nb_tapis)]
            
            if poules_u7:
                nb_rotations = 3
                noms_plateaux = [
                    "Plateau 1 : Motricité & Agilité",
                    "Plateau 2 : Ateliers techniques",
                    "Plateau 3 : Oppositions"
                ]
                
                # Affectation tournante des 3 groupes sur les 3 plateaux d'activité
                rot_affectations = [
                    {0: 0, 1: 1, 2: 2},  # Rot 1: P1->G1, P2->G2, P3->G3
                    {0: 2, 1: 0, 2: 1},  # Rot 2: P1->G3, P2->G1, P3->G2
                    {0: 1, 1: 2, 2: 0}   # Rot 3: P1->G2, P2->G3, P3->G1
                ]
                
                for r in range(nb_rotations):
                    h_rot = dt_debut_u7 + timedelta(minutes=r * duree_plateau_u7)
                    
                    for t in range(nb_tapis):
                        plat_idx = t % 3
                        grp_idx = rot_affectations[r][plat_idx]
                        
                        grp_poule = poules_u7[grp_idx] if grp_idx < len(poules_u7) else poules_u7[0]
                        grp_parts = grp_poule.get('participants', [])
                        
                        arb_nom = choisir_arbitre_match("", "", t)
                        arb_disp = arb_nom if (arb_nom and arb_nom != "Non attribué") else "Animateur FFLDA"
                        
                        item_plateau = {
                            "Type": "MATCH",
                            "Heure": h_rot.strftime("%H:%M"),
                            "Duree": duree_plateau_u7,
                            "Cat": f"U7 | Plateau {plat_idx + 1}",
                            "Combattant 1": f"Groupe {grp_idx + 1} ({len(grp_parts)} lutteurs)",
                            "Club 1": "Tous clubs",
                            "Comité 1": "",
                            "Combattant 2": noms_plateaux[plat_idx],
                            "Club 2": "",
                            "Comité 2": "",
                            "Tour": f"Rotation {r + 1}/3",
                            "Nom_Tour": f"Rotation {r + 1} / 3",
                            "Arbitre": arb_disp
                        }
                        planning_u7[t].append(item_plateau)
                
                fin_u7_globale = dt_debut_u7 + timedelta(minutes=3 * duree_plateau_u7)
                for t in range(nb_tapis):
                    tapis_heure_u7[t] = fin_u7_globale
            else:
                fin_u7_globale = dt_debut_u7

            # PHASE 1b : Tous les U9
            tapis_heure_u9, planning_u9 = ordonnancer_phase(poules_u9, fin_u7_globale, duree_u9, meme_tapis_poule)
            fin_u9_globale = max(tapis_heure_u9) if poules_u9 else fin_u7_globale

            planning_phase1 = {t: planning_u7[t] + planning_u9[t] for t in range(nb_tapis)}

            # PAUSE / PESÉE 2 (U11 / U13)
            dt_pesee_2 = None
            activer_pesee2_eff = st.session_state.get("downstream_pesee2_act", ("2" in type_pesee))
            if activer_pesee2_eff:
                if st.session_state.get("downstream_h_pesee2"):
                    dt_pesee_2 = datetime.combine(datetime.today(), st.session_state["downstream_h_pesee2"])
                else:
                    dt_pesee_2 = fin_u9_globale + timedelta(minutes=duree_pause)
                d_p2_eff = st.session_state.get("downstream_pesee2_dur", duree_pesee)
                dt_debut_p2_theorique = dt_pesee_2 + timedelta(minutes=d_p2_eff)
            else:
                dt_debut_p2_theorique = fin_u9_globale

            # Égalisation du nombre de lignes en Phase 1 (U7 + U9) pour aligner horizontalement la PAUSE
            max_phase1_lignes = max(len(planning_phase1[t]) for t in range(nb_tapis)) if planning_phase1 else 0
            for t in range(nb_tapis):
                while len(planning_phase1[t]) < max_phase1_lignes:
                    planning_phase1[t].append({"Type": "VIDE"})

            tapis_heure_p1 = [fin_u9_globale for _ in range(nb_tapis)]
            activer_pause_eff = st.session_state.get("downstream_pause_act", activer_pause)
            duree_pause_eff = st.session_state.get("downstream_pause_dur", duree_pause)
            if activer_pause_eff and duree_pause_eff > 0:
                for t in range(nb_tapis):
                    planning_phase1[t].append({"Type": "PAUSE", "Heure": fin_u9_globale.strftime("%H:%M")})
                    tapis_heure_p1[t] = fin_u9_globale + timedelta(minutes=duree_pause_eff)

            debut_p2_reel = max(tapis_heure_p1)
            if dt_debut_p2_theorique and debut_p2_reel < dt_debut_p2_theorique:
                debut_p2_reel = dt_debut_p2_theorique

            for t in range(nb_tapis):
                if (poules_u11 or poules_u13) and tapis_heure_p1[t] < debut_p2_reel:
                    attente = int((debut_p2_reel - tapis_heure_p1[t]).total_seconds() // 60)
                    if attente > 0:
                        planning_phase1[t].append({
                            "Type": "ATTENTE", 
                            "Heure": tapis_heure_p1[t].strftime("%H:%M"), 
                            "Texte": f"Pesée + échauffement U11/U13"
                        })

            # PHASE 2a : Tous les U11
            if poules_u11:
                tapis_heure_u11, planning_u11 = ordonnancer_phase(poules_u11, debut_p2_reel, duree_u11, meme_tapis_poule)
                fin_u11_globale = max(tapis_heure_u11)
            else:
                tapis_heure_u11 = [debut_p2_reel for _ in range(nb_tapis)]
                planning_u11 = {t: [] for t in range(nb_tapis)}
                fin_u11_globale = debut_p2_reel

            # PHASE 2b : Tous les U13
            if poules_u13:
                tapis_heure_u13, planning_u13 = ordonnancer_phase(poules_u13, fin_u11_globale, duree_u13, meme_tapis_poule)
                fin_estimee = max(tapis_heure_u13)
            else:
                planning_u13 = {t: [] for t in range(nb_tapis)}
                fin_estimee = fin_u11_globale

            planning_tapis = {t: planning_phase1[t] + planning_u11[t] + planning_u13[t] for t in range(nb_tapis)}

            texte_pesee_1 = "1ère pesée" if "1" in type_pesee else "Pesée U7/U9"
            valeur_pause = f"{duree_pause} min" if (activer_pause and duree_pause > 0) else "0 min"

            str_comp_u7 = f"{dt_debut_u7.strftime('%H:%M')} - {fin_u7_globale.strftime('%H:%M')}"
            str_comp_u9 = f"{fin_u7_globale.strftime('%H:%M')} - {fin_u9_globale.strftime('%H:%M')}"
            str_comp_u11 = f"{debut_p2_reel.strftime('%H:%M')} - {fin_u11_globale.strftime('%H:%M')}"
            str_comp_u13 = f"{fin_u11_globale.strftime('%H:%M')} - {fin_estimee.strftime('%H:%M')}"

            st.success("Fichier analysé avec succès ! Tournoi généré.")
            
            # --- JOURNALISATION AUTOMATIQUE SILENCIEUSE DE LA FACTURATION ---
            try:
                enregistrer_log_facturation(
                    code_organisateur=st.session_state.get("code_session", "ORGANISATEUR"),
                    nom_tournoi=f"{nom_competition} [GÉNÉRATION TOURNOI]",
                    nb_inscrits=total_inscrits_global,
                    nb_peses=total_participants_peses,
                    nb_matchs=total_matchs_calcules
                )
            except Exception:
                pass
            
            # --- PRÉPARATION DES DONNÉES DU RÉSUMÉ ---
            lignes_accueil = [
                {"Étape de la journée": texte_pesee_1, "Horaire / Valeur": dt_pesee_u9.strftime('%H:%M')},
            ]
            if poules_u7:
                lignes_accueil.append({"Étape de la journée": "Animation U7 (3 Plateaux)", "Horaire / Valeur": str_comp_u7})
                lignes_accueil.append({"Étape de la journée": "Groupes U7 (Plateaux)", "Horaire / Valeur": f"{len(poules_u7)} groupes"})
            lignes_accueil.extend([
                {"Étape de la journée": "Compétition U9", "Horaire / Valeur": str_comp_u9},
                {"Étape de la journée": "Poules U9 générées", "Horaire / Valeur": f"{len(poules_u9)} poules"},
                {"Étape de la journée": "Pause de la compétition", "Horaire / Valeur": valeur_pause}
            ])
            if "2" in type_pesee and dt_pesee_2:
                lignes_accueil.append({"Étape de la journée": "2ème pesée (U11/U13)", "Horaire / Valeur": dt_pesee_2.strftime('%H:%M')})

            lignes_accueil.extend([
                {"Étape de la journée": "Compétition U11", "Horaire / Valeur": str_comp_u11},
                {"Étape de la journée": "Poules U11 générées", "Horaire / Valeur": f"{len(poules_u11)} poules"},
            ])
            if poules_u13:
                lignes_accueil.extend([
                    {"Étape de la journée": "Compétition U13", "Horaire / Valeur": str_comp_u13},
                    {"Étape de la journée": "Poules & Tableaux U13", "Horaire / Valeur": f"{len(poules_u13)} catégories"},
                ])
            total_poules_calc = len(poules_u7) + len(poules_u9) + len(poules_u11) + len(poules_u13)
            lignes_accueil.extend([
                {"Étape de la journée": "Total Groupes (U7 + U9 + U11 + U13)", "Horaire / Valeur": f"{total_poules_calc} groupes"},
                {"Étape de la journée": "Fin de la compétition estimée", "Horaire / Valeur": fin_estimee.strftime('%H:%M')}
            ])

            # --- GÉNÉRATION DES DOCUMENTS HTML & PDF PAYSAGE DU TOURNOI ---
            sections_tournoi_complet = [
                ("📊 Résumé Prévisionnel de la Journée", pd.DataFrame(lignes_accueil))
            ]
            
            if liste_arbitres:
                lignes_arb_print = []
                for t in range(nb_tapis):
                    noms_arb = ", ".join([a['Nom_Complet'] for a in tapis_arbitres[t]]) if tapis_arbitres[t] else "Aucun arbitre affecté"
                    lignes_arb_print.append({"Tapis": f"Tapis {t + 1}", "Effectif": f"{len(tapis_arbitres[t])} arbitres", "Équipe d'Arbitrage Désignée": noms_arb})
                sections_tournoi_complet.append(("Désignation des Équipes d'Arbitrage par Tapis", pd.DataFrame(lignes_arb_print)))

            max_lignes = max(len(liste) for liste in planning_tapis.values()) if planning_tapis else 0
            grille_ui = []
            for row_idx in range(max_lignes):
                ligne = {}
                for t in range(nb_tapis):
                    col = f"Tapis {t + 1}"
                    if row_idx < len(planning_tapis[t]):
                        m = planning_tapis[t][row_idx]
                        if m["Type"] == "PAUSE": ligne[col] = f"[{m['Heure']}] ⏸️ PAUSE"
                        elif m["Type"] == "ATTENTE": ligne[col] = f"[{m['Heure']}] {m['Texte']}"
                        elif m["Type"] == "VIDE": ligne[col] = ""
                        else:
                            arb_str = f" (Arbitre : {m['Arbitre']})" if m.get('Arbitre') and m['Arbitre'] != "Non attribué" else ""
                            t_nom = nettoyer_nom_tour(m.get('Nom_Tour') or m.get('Tour'))
                            tour_str = f" [🎯 Tour : {t_nom}]" if t_nom else ""
                            ligne[col] = f"[{m['Heure']}] ({m['Duree']}m) [{m['Cat']}]{tour_str} - {m['Combattant 1']} vs {m['Combattant 2']}{arb_str}"
                    else: ligne[col] = ""
                grille_ui.append(ligne)
            
            sections_tournoi_complet.append(("📅 Grille Globale de Passage", pd.DataFrame(grille_ui)))

            # Ajout des sections de Grille par Tapis pour impression/PDF (1 page par onglet)
            for t in range(nb_tapis):
                lignes_tapis_doc = []
                m_count_doc = 0
                for m in planning_tapis[t]:
                    if m["Type"] == "PAUSE":
                        lignes_tapis_doc.append({"N°": "-", "Heure": m["Heure"], "Catégorie": "⏸️ PAUSE", "Tour": "-", "Lutteur Rouge": "-", "Pt Clt (R)": "", "Lutteur Bleu": "-", "Pt Clt (B)": "", "Arbitre": ""})
                    elif m["Type"] == "ATTENTE":
                        lignes_tapis_doc.append({"N°": "-", "Heure": m["Heure"], "Catégorie": f"⏳ {m['Texte']}", "Tour": "-", "Lutteur Rouge": "-", "Pt Clt (R)": "", "Lutteur Bleu": "-", "Pt Clt (B)": "", "Arbitre": ""})
                    elif m["Type"] == "MATCH":
                        m_count_doc += 1
                        c1_t = f"{m['Combattant 1']}"
                        if m.get('Club 1'): c1_t += f" ({m['Club 1']})"
                        c2_t = f"{m['Combattant 2']}"
                        if m.get('Club 2'): c2_t += f" ({m['Club 2']})"
                        lignes_tapis_doc.append({
                            "N°": f"M{m_count_doc}",
                            "Heure": f"{m['Heure']}",
                            "Catégorie": m['Cat'],
                            "Tour": nettoyer_nom_tour(m.get('Nom_Tour') or m.get('Tour') or "-"),
                            "Lutteur Rouge": c1_t,
                            "Pt Clt (R)": "[   ]",
                            "Type (R)": "[   ]",
                            "Lutteur Bleu": c2_t,
                            "Pt Clt (B)": "[   ]",
                            "Type (B)": "[   ]",
                            "Arbitre": m.get("Arbitre", "")
                        })
                sections_tournoi_complet.append((f"🥋 Grille de Passage & Scores - Tapis {t + 1}", pd.DataFrame(lignes_tapis_doc)))

            for nom_poule, liste_p in participants_par_poule.items():
                p_obj = poule_obj_map.get(nom_poule)
                if p_obj and p_obj.get('type_formule') == 'plateau_u7':
                    df_poule_vue = pd.DataFrame([
                        {
                            "N°": idx,
                            "Nom": p.get('Nom', ''),
                            "Club": p.get('Club', ''),
                            "Poids": formater_poids(p.get('Poids', '')),
                            "Plateau 1 (Motricité & Agilité)": "[ ✓ ]",
                            "Plateau 2 (Ateliers techniques)": "[ ✓ ]",
                            "Plateau 3 (Oppositions)": "[ ✓ ]",
                            "Validation": "🥇 Médaille d'Or"
                        }
                        for idx, p in enumerate(liste_p, 1)
                    ])
                    sections_tournoi_complet.append((f"🏅 Animation U7 (3 Plateaux) : {nom_poule}", df_poule_vue))
                else:
                    df_poule_vue = pd.DataFrame(liste_p)[['Nom', 'Club', 'Poids']]
                    sections_tournoi_complet.append((f"🤼 Feuille de Poule : {nom_poule}", df_poule_vue))
                
            # --- GÉNÉRATION DU CLASSEUR EXCEL OFFICIEL FFLDA ---
            output_excel = io.BytesIO()
            coords_matchs_tapis = {}
            matchs_tournoi_direct = []
            tapis_slots_map = {}
            poule_sheet_names = {}
            poule_tapis_map = {}

            for t_idx in range(nb_tapis):
                for it_m in planning_tapis[t_idx]:
                    if it_m.get("Type") == "MATCH" and it_m.get("Cat"):
                        c_k = it_m["Cat"]
                        if c_k not in poule_tapis_map:
                            poule_tapis_map[c_k] = t_idx + 1

            feuilles_reservees = {"résumé", "grille de passage", "corps d'arbitrage"} | {f"grille tapis {t + 1}" for t in range(nb_tapis)}
            for nom_poule in participants_par_poule.keys():
                nom_base = abreger_nom_onglet(nom_poule)
                nom_onglet_court = nom_base
                suffix_i = 1
                while nom_onglet_court.lower() in feuilles_reservees:
                    nom_onglet_court = f"{nom_base[:28]}_{suffix_i}"
                    suffix_i += 1
                feuilles_reservees.add(nom_onglet_court.lower())
                poule_sheet_names[nom_poule] = nom_onglet_court
                if nom_poule not in poule_tapis_map:
                    for c_k, t_num in list(poule_tapis_map.items()):
                        if c_k in nom_poule or nom_poule in c_k:
                            poule_tapis_map[nom_poule] = t_num
                            break

            with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                writer.book.calculation.fullCalcOnLoad = True
                
                resume_data = [
                    {"Étape de la journée": texte_pesee_1, "Horaire / Valeur": dt_pesee_u9.strftime('%H:%M')},
                ]
                if poules_u7:
                    resume_data.append({"Étape de la journée": "Animation U7 (3 Plateaux)", "Horaire / Valeur": str_comp_u7})
                    resume_data.append({"Étape de la journée": "Groupes U7 (Plateaux)", "Horaire / Valeur": f"{len(poules_u7)} groupes"})
                resume_data.extend([
                    {"Étape de la journée": "Compétition U9", "Horaire / Valeur": str_comp_u9},
                    {"Étape de la journée": "Pause de la compétition", "Horaire / Valeur": valeur_pause}
                ])
                if "2" in type_pesee and dt_pesee_2:
                    resume_data.append({"Étape de la journée": "2ème pesée (U11/U13)", "Horaire / Valeur": dt_pesee_2.strftime('%H:%M')})
                
                resume_data.extend([
                    {"Étape de la journée": "Compétition U11", "Horaire / Valeur": str_comp_u11},
                ])
                if poules_u13:
                    resume_data.append({"Étape de la journée": "Compétition U13", "Horaire / Valeur": str_comp_u13})
                resume_data.append({"Étape de la journée": "Fin de la compétition estimée", "Horaire / Valeur": fin_estimee.strftime('%H:%M')})
                pd.DataFrame(resume_data).to_excel(writer, sheet_name="Résumé", index=False)
                
                max_lignes = max(len(liste) for liste in planning_tapis.values()) if planning_tapis else 0
                grille = []
                for row_idx in range(max_lignes):
                    ligne = {}
                    for t in range(nb_tapis):
                        col = f"Tapis {t + 1}"
                        if row_idx < len(planning_tapis[t]):
                            m = planning_tapis[t][row_idx]
                            if m["Type"] == "PAUSE": ligne[col] = f"[{m['Heure']}]\n⏸️ PAUSE DE LA COMPÉTITION"
                            elif m["Type"] == "ATTENTE": ligne[col] = f"[{m['Heure']}]\n{m['Texte']}"
                            elif m["Type"] == "VIDE": ligne[col] = ""
                            else:
                                arb_str = f"\nArbitre : {m['Arbitre']}" if m.get('Arbitre') and m['Arbitre'] != "Non attribué" else ""
                                t_nom = nettoyer_nom_tour(m.get('Nom_Tour') or m.get('Tour'))
                                tour_str = f"\n🎯 Tour : {t_nom}" if t_nom else ""
                                ligne[col] = f"🕘 {m['Heure']} ({m['Duree']} min)\n[{m['Cat']}]{tour_str}\n{m['Combattant 1']} VS {m['Combattant 2']}{arb_str}"
                        else: ligne[col] = ""
                    grille.append(ligne)
                pd.DataFrame(grille).to_excel(writer, sheet_name="Grille de Passage", index=False, startrow=1)
                
                b_style = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                bleu = PatternFill("solid", fgColor="0055A4")
                rouge = PatternFill("solid", fgColor="EF4135")
                bleu_clair = PatternFill("solid", fgColor="DDEBF7") 
                rouge_lutte = PatternFill("solid", fgColor="E53935") 
                bleu_lutte = PatternFill("solid", fgColor="1E88E5")  
                gris_clair = PatternFill("solid", fgColor="F2F2F2")
                entete_noir = PatternFill("solid", fgColor="000000")

                for t in range(nb_tapis):
                    ws_mat = writer.book.create_sheet(f"Grille Tapis {t + 1}")
                    ws_mat.views.sheetView[0].showGridLines = True
                    ws_mat.page_setup.orientation = ws_mat.ORIENTATION_LANDSCAPE
                    ws_mat.page_setup.paperSize = ws_mat.PAPERSIZE_A4
                    ws_mat.sheet_properties.pageSetUpPr.fitToPage = True
                    ws_mat.page_setup.fitToWidth = 1
                    ws_mat.page_setup.fitToHeight = 0
                    
                    dv_u9_u11 = DataValidation(type="list", formula1='"0,1,2"', allow_blank=True)
                    dv_u13 = DataValidation(type="list", formula1='"0,1,3,4,5"', allow_blank=True)
                    dv_type = DataValidation(type="list", formula1='"VT,VST,VP,DT,DST,DP"', allow_blank=True)
                    has_u9_u11 = False
                    has_u13 = False
                    has_type = False
                    
                    ws_mat.row_dimensions[1].height = 35
                    ws_mat.merge_cells(start_row=1, start_column=1, end_row=1, end_column=11)
                    titre_mat = ws_mat.cell(row=1, column=1, value=f"🏆 {nom_competition.upper()} - GRILLE DE PASSAGE : TAPIS {t + 1} 🏆")
                    titre_mat.font = Font(name="Arial", size=16, bold=True, color="FFFFFF")
                    titre_mat.fill = bleu
                    titre_mat.alignment = Alignment(horizontal="center", vertical="center")
                    
                    noms_arb = ", ".join([a['Nom_Complet'] for a in tapis_arbitres[t]]) if (liste_arbitres and tapis_arbitres[t]) else "Aucun arbitre affecté"
                    ws_mat.row_dimensions[2].height = 22
                    ws_mat.merge_cells(start_row=2, start_column=1, end_row=2, end_column=11)
                    sub_mat = ws_mat.cell(row=2, column=1, value=f"Arbitrage : {noms_arb}  |  * Saisie rapide : choisir le Type (VT, VST, VP) du vainqueur calcule automatiquement les Pt Clt (U13: 5/0, 4/1, 3/1 | U9-U11: 2/1)")
                    sub_mat.font = Font(name="Arial", size=10, italic=True, bold=True, color="0055A4")
                    sub_mat.alignment = Alignment(horizontal="center", vertical="center")

                    matches_on_tapis = [it for it in planning_tapis[t] if it.get("Type") == "MATCH"]
                    max_nom_t = max([len(str(it.get('Lutteur1', ''))) for it in matches_on_tapis] + [len(str(it.get('Lutteur2', ''))) for it in matches_on_tapis] + [15])
                    max_club_t = max([len(str(it.get('Club1', ''))) for it in matches_on_tapis] + [len(str(it.get('Club2', ''))) for it in matches_on_tapis] + [12])
                    w_nom_t = max(max_nom_t + 4, 25)
                    w_club_t = max(max_club_t + 4, 18)

                    ws_mat.column_dimensions['A'].width = 14
                    ws_mat.column_dimensions['B'].width = 5
                    ws_mat.column_dimensions['C'].width = max(w_nom_t - 2, 20)
                    ws_mat.column_dimensions['D'].width = max(w_club_t - 2, 14)
                    ws_mat.column_dimensions['E'].width = 8
                    ws_mat.column_dimensions['F'].width = 8
                    ws_mat.column_dimensions['G'].width = 5
                    ws_mat.column_dimensions['H'].width = max(w_nom_t - 2, 20)
                    ws_mat.column_dimensions['I'].width = max(w_club_t - 2, 14)
                    ws_mat.column_dimensions['J'].width = 8
                    ws_mat.column_dimensions['K'].width = 8

                    r_curr = 4
                    m_count_t = 0
                    for item in planning_tapis[t]:
                        if item["Type"] == "PAUSE":
                            ws_mat.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=11)
                            p_cell = ws_mat.cell(row=r_curr, column=1, value=f"[{item['Heure']}] ⏸️ PAUSE DE LA COMPÉTITION")
                            p_cell.fill, p_cell.font, p_cell.alignment = rouge, Font(bold=True, color="FFFFFF", size=11), Alignment(horizontal="center", vertical="center")
                            ws_mat.row_dimensions[r_curr].height = 24
                            r_curr += 2
                        elif item["Type"] == "ATTENTE":
                            ws_mat.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=11)
                            a_cell = ws_mat.cell(row=r_curr, column=1, value=f"[{item['Heure']}] ⏳ {item['Texte']}")
                            a_cell.fill, a_cell.font, a_cell.alignment = PatternFill("solid", fgColor="EFEFEF"), Font(italic=True, color="666666", size=10), Alignment(horizontal="center", vertical="center")
                            ws_mat.row_dimensions[r_curr].height = 22
                            r_curr += 2
                        elif item["Type"] == "MATCH":
                            m_count_t += 1
                            ws_mat.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=11)
                            arb_txt = f"  |  Arbitre : {item['Arbitre']}" if item.get('Arbitre') and item['Arbitre'] != "Non attribué" else ""
                            tour_label = nettoyer_nom_tour(item.get('Nom_Tour') or item.get('Tour'))
                            tour_txt = f"  |  🎯 Tour : {tour_label}" if tour_label else ""
                            cat_item = item.get('Cat', '')
                            sheet_cible = poule_sheet_names.get(cat_item)
                            if not sheet_cible:
                                for p_nom, s_name in poule_sheet_names.items():
                                    if cat_item == p_nom or cat_item in p_nom or p_nom in cat_item:
                                        sheet_cible = s_name
                                        break
                            
                            lien_suffix = "  |  🔗 Voir Poule" if sheet_cible else ""
                            hdr_text = f"MATCH N° {m_count_t}  |  🕘 {item['Heure']} ({item['Duree']} min)  |  Catégorie : {cat_item}{tour_txt}{arb_txt}{lien_suffix}"
                            age_m = extraire_age_de_texte(cat_item)
                            cfg_m = COULEURS_AGE_GRILLE.get(age_m, COULEURS_AGE_GRILLE['AUTRE'])
                            if sheet_cible:
                                h_cell = ws_mat.cell(row=r_curr, column=1, value=f'=HYPERLINK("#\'{sheet_cible}\'!A1", "{hdr_text}")')
                                h_cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=11, underline="single")
                            else:
                                h_cell = ws_mat.cell(row=r_curr, column=1, value=hdr_text)
                                h_cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
                            h_cell.fill, h_cell.alignment = cfg_m['bg_excel_header'], Alignment(horizontal="center", vertical="center")
                            ws_mat.row_dimensions[r_curr].height = 24
                            r_curr += 1

                            c_tour_h = ws_mat.cell(row=r_curr, column=1, value="TOUR / PHASE")
                            c_tour_h.fill, c_tour_h.font, c_tour_h.alignment, c_tour_h.border = gris_clair, Font(bold=True, size=8), Alignment(horizontal="center", vertical="center"), b_style

                            c_num_h = ws_mat.cell(row=r_curr, column=2, value="N°")
                            c_num_h.alignment, c_num_h.border, c_num_h.fill = Alignment(horizontal="center", vertical="center"), b_style, gris_clair
                            
                            c_rouge_h = ws_mat.cell(row=r_curr, column=3, value="LUTTEUR ROUGE")
                            c_rouge_h.fill, c_rouge_h.font, c_rouge_h.alignment, c_rouge_h.border = rouge_lutte, Font(bold=True, color="FFFFFF"), Alignment(horizontal="center", vertical="center"), b_style
                            ws_mat.merge_cells(start_row=r_curr, start_column=3, end_row=r_curr, end_column=4)
                            ws_mat.cell(row=r_curr, column=4).border = b_style
                            
                            c_ptr = ws_mat.cell(row=r_curr, column=5, value="Pt Clt")
                            c_ptr.font, c_ptr.alignment, c_ptr.border, c_ptr.fill = Font(bold=True, size=9), Alignment(horizontal="center", vertical="center"), b_style, gris_clair

                            c_typer_h = ws_mat.cell(row=r_curr, column=6, value="Type")
                            c_typer_h.font, c_typer_h.alignment, c_typer_h.border, c_typer_h.fill = Font(bold=True, size=9), Alignment(horizontal="center", vertical="center"), b_style, gris_clair
                            
                            c_vs_h = ws_mat.cell(row=r_curr, column=7, value="VS")
                            c_vs_h.alignment, c_vs_h.border, c_vs_h.font = Alignment(horizontal="center", vertical="center"), b_style, Font(bold=True, size=9)
                            
                            c_bleu_h = ws_mat.cell(row=r_curr, column=8, value="LUTTEUR BLEU")
                            c_bleu_h.fill, c_bleu_h.font, c_bleu_h.alignment, c_bleu_h.border = bleu_lutte, Font(bold=True, color="FFFFFF"), Alignment(horizontal="center", vertical="center"), b_style
                            ws_mat.merge_cells(start_row=r_curr, start_column=8, end_row=r_curr, end_column=9)
                            ws_mat.cell(row=r_curr, column=9).border = b_style
                            
                            c_ptb = ws_mat.cell(row=r_curr, column=10, value="Pt Clt")
                            c_ptb.font, c_ptb.alignment, c_ptb.border, c_ptb.fill = Font(bold=True, size=9), Alignment(horizontal="center", vertical="center"), b_style, gris_clair

                            c_typeb_h = ws_mat.cell(row=r_curr, column=11, value="Type")
                            c_typeb_h.font, c_typeb_h.alignment, c_typeb_h.border, c_typeb_h.fill = Font(bold=True, size=9), Alignment(horizontal="center", vertical="center"), b_style, gris_clair
                            
                            r_curr += 1
                            
                            c_tour_v = ws_mat.cell(row=r_curr, column=1, value=tour_label)
                            c_tour_v.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                            c_tour_v.font = Font(bold=True, size=9, color="0055A4")
                            c_tour_v.border = b_style
                            ws_mat.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr+2, end_column=1)
                            ws_mat.cell(row=r_curr+1, column=1).border = b_style
                            ws_mat.cell(row=r_curr+2, column=1).border = b_style

                            ws_mat.cell(row=r_curr, column=2, value=m_count_t).alignment = Alignment(horizontal="center", vertical="center")
                            ws_mat.cell(row=r_curr, column=2).font = Font(bold=True, color="E53935", size=12)
                            ws_mat.cell(row=r_curr, column=2).border = b_style
                            ws_mat.merge_cells(start_row=r_curr, start_column=2, end_row=r_curr+2, end_column=2)
                            ws_mat.cell(row=r_curr+1, column=2).border = b_style
                            ws_mat.cell(row=r_curr+2, column=2).border = b_style
                            
                            c1_str = f"{item['Combattant 1']}"
                            if item.get('Club 1'): c1_str += f" ({item['Club 1']})"
                            if item.get('Comité 1') and item['Comité 1'] != 'Comité Non Renseigné': c1_str += f" - {item['Comité 1']}"
                            
                            c_r_info = ws_mat.cell(row=r_curr, column=3, value=c1_str)
                            c_r_info.border = b_style
                            c_r_info.alignment = Alignment(vertical="center")
                            ws_mat.merge_cells(start_row=r_curr, start_column=3, end_row=r_curr, end_column=4)
                            ws_mat.cell(row=r_curr, column=4).border = b_style
                            
                            box_ptr = ws_mat.cell(row=r_curr, column=5)
                            box_ptr.border, box_ptr.fill = b_style, gris_clair
                            box_ptr.alignment = Alignment(horizontal="center", vertical="center")

                            box_typer = ws_mat.cell(row=r_curr, column=6)
                            box_typer.border, box_typer.fill = b_style, gris_clair
                            box_typer.alignment = Alignment(horizontal="center", vertical="center")
                            box_typer.font = Font(name="Arial", size=9, bold=True, color="0055A4")
                            dv_type.add(box_typer.coordinate)
                            has_type = True
                            
                            c_vs_mid = ws_mat.cell(row=r_curr, column=7, value="-")
                            c_vs_mid.alignment = Alignment(horizontal="center", vertical="center")
                            c_vs_mid.border = b_style
                            
                            c2_str = f"{item['Combattant 2']}"
                            if item.get('Club 2'): c2_str += f" ({item['Club 2']})"
                            if item.get('Comité 2') and item['Comité 2'] != 'Comité Non Renseigné': c2_str += f" - {item['Comité 2']}"
                            
                            c_b_info = ws_mat.cell(row=r_curr, column=8, value=c2_str)
                            c_b_info.border = b_style
                            c_b_info.alignment = Alignment(vertical="center")
                            ws_mat.merge_cells(start_row=r_curr, start_column=8, end_row=r_curr, end_column=9)
                            ws_mat.cell(row=r_curr, column=9).border = b_style
                            
                            box_ptb = ws_mat.cell(row=r_curr, column=10)
                            box_ptb.border, box_ptb.fill = b_style, gris_clair
                            box_ptb.alignment = Alignment(horizontal="center", vertical="center")

                            box_typeb = ws_mat.cell(row=r_curr, column=11)
                            box_typeb.border, box_typeb.fill = b_style, gris_clair
                            box_typeb.alignment = Alignment(horizontal="center", vertical="center")
                            box_typeb.font = Font(name="Arial", size=9, bold=True, color="0055A4")
                            dv_type.add(box_typeb.coordinate)
                            has_type = True
                            
                            # Menus déroulants sous Pt Clt selon la catégorie d'âge (U9/U11 : 0,1,2 | U13 : 0,1,3,4,5)
                            if age_m in ['U9', 'U11']:
                                dv_u9_u11.add(box_ptr.coordinate)
                                dv_u9_u11.add(box_ptb.coordinate)
                                has_u9_u11 = True
                            elif age_m == 'U13':
                                dv_u13.add(box_ptr.coordinate)
                                dv_u13.add(box_ptb.coordinate)
                                has_u13 = True
                            
                            def reg_tapis_slot(k, sl):
                                if k not in tapis_slots_map:
                                    tapis_slots_map[k] = [sl]
                                elif isinstance(tapis_slots_map[k], list):
                                    tapis_slots_map[k].append(sl)
                                else:
                                    tapis_slots_map[k] = [tapis_slots_map[k], sl]

                            cat_m = item.get('Cat', '')
                            c1_m = str(item.get('Combattant 1', ''))
                            c2_m = str(item.get('Combattant 2', ''))
                            s1 = (ws_mat, f"C{r_curr}")
                            s2 = (ws_mat, f"H{r_curr}")
                            
                            reg_tapis_slot((cat_m, c1_m), s1)
                            reg_tapis_slot((cat_m, c2_m), s2)
                            reg_tapis_slot(c1_m, s1)
                            reg_tapis_slot(c2_m, s2)
                            
                            b_c1 = re.sub(r'\s*\[.*?\]', '', c1_m).strip()
                            b_c2 = re.sub(r'\s*\[.*?\]', '', c2_m).strip()
                            if b_c1 and b_c1 != c1_m:
                                reg_tapis_slot((cat_m, b_c1), s1)
                                reg_tapis_slot(b_c1, s1)
                            if b_c2 and b_c2 != c2_m:
                                reg_tapis_slot((cat_m, b_c2), s2)
                                reg_tapis_slot(b_c2, s2)
                            
                            r_curr += 1
                            
                            ws_mat.cell(row=r_curr, column=3, value="Points Techniques (Actions)").font = Font(size=9, italic=True)
                            ws_mat.merge_cells(start_row=r_curr, start_column=3, end_row=r_curr, end_column=4)
                            ws_mat.cell(row=r_curr, column=5, value="Total Score").font = Font(size=9, italic=True)
                            ws_mat.merge_cells(start_row=r_curr, start_column=5, end_row=r_curr, end_column=6)
                            
                            ws_mat.cell(row=r_curr, column=8, value="Points Techniques (Actions)").font = Font(size=9, italic=True)
                            ws_mat.merge_cells(start_row=r_curr, start_column=8, end_row=r_curr, end_column=9)
                            ws_mat.cell(row=r_curr, column=10, value="Total Score").font = Font(size=9, italic=True)
                            ws_mat.merge_cells(start_row=r_curr, start_column=10, end_row=r_curr, end_column=11)
                            
                            r_curr += 1
                            
                            ws_mat.row_dimensions[r_curr].height = 25
                            c_act_r = ws_mat.cell(row=r_curr, column=3)
                            c_act_r.border = b_style
                            ws_mat.cell(row=r_curr, column=4).border = b_style
                            ws_mat.merge_cells(start_row=r_curr, start_column=3, end_row=r_curr, end_column=4)
                            
                            c_tot_r = ws_mat.cell(row=r_curr, column=5)
                            c_tot_r.border = b_style
                            c_tot_r.alignment = Alignment(horizontal="center", vertical="center")
                            c_tot_r.font = Font(name="Arial", size=10, bold=True)
                            ws_mat.cell(row=r_curr, column=6).border = b_style
                            ws_mat.merge_cells(start_row=r_curr, start_column=5, end_row=r_curr, end_column=6)

                            ws_mat.cell(row=r_curr, column=7).border = b_style
                            
                            c_act_b = ws_mat.cell(row=r_curr, column=8)
                            c_act_b.border = b_style
                            ws_mat.cell(row=r_curr, column=9).border = b_style
                            ws_mat.merge_cells(start_row=r_curr, start_column=8, end_row=r_curr, end_column=9)
                            
                            c_tot_b = ws_mat.cell(row=r_curr, column=10)
                            c_tot_b.border = b_style
                            c_tot_b.alignment = Alignment(horizontal="center", vertical="center")
                            c_tot_b.font = Font(name="Arial", size=10, bold=True)
                            ws_mat.cell(row=r_curr, column=11).border = b_style
                            ws_mat.merge_cells(start_row=r_curr, start_column=10, end_row=r_curr, end_column=11)

                            # Calcul automatique officiel FFLDA/UWW des Pt Clt selon le type de victoire et les scores
                            tr_c = f"F{r_curr - 2}"
                            tb_c = f"K{r_curr - 2}"
                            tot_r_c = f"E{r_curr}"
                            tot_b_c = f"J{r_curr}"
                            
                            if age_m == 'U13':
                                # Barème officiel FFLDA/UWW U13 :
                                # VT : 5 / 0
                                # VST : 4 / (1 si perdant a marqué >= 1 pt tech, sinon 0)
                                # VP : 3 / (1 si perdant a marqué >= 1 pt tech, sinon 0)
                                box_ptr.value = f'=IF(AND({tr_c}="",{tb_c}=""),"",IF(OR({tr_c}="VT",{tb_c}="DT"),5,IF(OR({tr_c}="VST",{tb_c}="DST"),4,IF(OR({tr_c}="VP",{tb_c}="DP"),3,IF(OR({tr_c}="DT",{tb_c}="VT"),0,IF(OR({tr_c}="DST",{tb_c}="VST",{tr_c}="DP",{tb_c}="VP"),IF(N({tot_r_c})>0,1,0),""))))))'
                                box_ptb.value = f'=IF(AND({tr_c}="",{tb_c}=""),"",IF(OR({tb_c}="VT",{tr_c}="DT"),5,IF(OR({tb_c}="VST",{tr_c}="DST"),4,IF(OR({tb_c}="VP",{tr_c}="DP"),3,IF(OR({tb_c}="DT",{tr_c}="VT"),0,IF(OR({tb_c}="DST",{tr_c}="VST",{tb_c}="DP",{tr_c}="VP"),IF(N({tot_b_c})>0,1,0),""))))))'
                            else:
                                # Barème U9 / U11 : Vainqueur -> 2, Perdant -> 1 si >= 1 pt tech marqué, sinon 0
                                box_ptr.value = f'=IF(AND({tr_c}="",{tb_c}=""),"",IF(OR({tr_c}="VT",{tr_c}="VST",{tr_c}="VP",{tb_c}="DT",{tb_c}="DST",{tb_c}="DP"),2,IF(OR({tr_c}="DT",{tr_c}="DST",{tr_c}="DP",{tb_c}="VT",{tb_c}="VST",{tb_c}="VP"),IF(N({tot_r_c})>0,1,0),"")))'
                                box_ptb.value = f'=IF(AND({tr_c}="",{tb_c}=""),"",IF(OR({tb_c}="VT",{tb_c}="VST",{tb_c}="VP",{tr_c}="DT",{tr_c}="DST",{tr_c}="DP"),2,IF(OR({tb_c}="DT",{tb_c}="DST",{tb_c}="DP",{tr_c}="VT",{tr_c}="VST",{tr_c}="VP"),IF(N({tot_b_c})>0,1,0),"")))'

                            # Enregistrement des coordonnées des cases Pt Clt, Type, Actions et Total Score sur la Grille Tapis X
                            m_coord_info = {
                                'sheet': f"Grille Tapis {t + 1}",
                                'sheet_name': f"Grille Tapis {t + 1}",
                                'ptr_cell': f"E{r_curr - 2}",
                                'typer_cell': f"F{r_curr - 2}",
                                'tot_r_cell': f"E{r_curr}",
                                'act_r_cell': f"C{r_curr}",
                                'ptb_cell': f"J{r_curr - 2}",
                                'typeb_cell': f"K{r_curr - 2}",
                                'tot_b_cell': f"J{r_curr}",
                                'act_b_cell': f"H{r_curr}",
                                'p1': item['Combattant 1'],
                                'p2': item['Combattant 2']
                            }
                            coords_matchs_tapis[(cat_m, c1_m, c2_m)] = m_coord_info
                            coords_matchs_tapis[(cat_m, c2_m, c1_m)] = m_coord_info
                            if (b_c1 and b_c1 != c1_m) or (b_c2 and b_c2 != c2_m):
                                coords_matchs_tapis[(cat_m, b_c1 or c1_m, b_c2 or c2_m)] = m_coord_info
                                coords_matchs_tapis[(cat_m, b_c2 or c2_m, b_c1 or c1_m)] = m_coord_info
                            
                            m_prefix_code = st.session_state.get('code_session', 'ORG')
                            m_nom_tourn = re.sub(r'[^a-zA-Z0-9]', '_', nom_competition)[:15]
                            matchs_tournoi_direct.append({
                                "id": f"{m_prefix_code}_{m_nom_tourn}_T{t + 1}_M{m_count_t:03d}",
                                "tournoi_id": nom_competition,
                                "code_organisateur": m_prefix_code,
                                "tapis": t + 1,
                                "match_num": m_count_t,
                                "heure": item.get("Heure", ""),
                                "duree": item.get("Duree", 2),
                                "categorie": cat_item,
                                "tour": nettoyer_nom_tour(item.get("Nom_Tour") or item.get("Tour") or ""),
                                "lutteur_rouge": item.get("Combattant 1", ""),
                                "club_rouge": item.get("Club 1", ""),
                                "comite_rouge": item.get("Comité 1", ""),
                                "lutteur_bleu": item.get("Combattant 2", ""),
                                "club_bleu": item.get("Club 2", ""),
                                "comite_bleu": item.get("Comité 2", ""),
                                "sheet": f"Grille Tapis {t + 1}",
                                "typer_cell": f"F{r_curr - 2}",
                                "tot_r_cell": f"E{r_curr}",
                                "typeb_cell": f"K{r_curr - 2}",
                                "tot_b_cell": f"J{r_curr}",
                                "statut": "À venir",
                                "vainqueur": "",
                                "type_victoire": "",
                                "score_rouge": 0,
                                "score_bleu": 0,
                                "pt_clt_rouge": 0,
                                "pt_clt_bleu": 0
                            })
                            
                            r_curr += 2 
                    
                    if has_u9_u11:
                        ws_mat.add_data_validation(dv_u9_u11)
                    if has_u13:
                        ws_mat.add_data_validation(dv_u13)
                    if has_type:
                        ws_mat.add_data_validation(dv_type)
                
                for ws_name in writer.book.sheetnames:
                    ws_sheet = writer.book[ws_name]
                    ws_sheet.page_setup.orientation = ws_sheet.ORIENTATION_LANDSCAPE
                    ws_sheet.page_setup.paperSize = ws_sheet.PAPERSIZE_A4
                    ws_sheet.sheet_properties.pageSetUpPr.fitToPage = True
                    ws_sheet.page_setup.fitToWidth = 1
                    ws_sheet.page_setup.fitToHeight = 0
                
                ws_res = writer.sheets["Résumé"]
                for cell in ws_res[1]: 
                    cell.fill, cell.font, cell.alignment = bleu, Font(bold=True, color="FFFFFF"), Alignment(horizontal="center")
                ws_res.column_dimensions['A'].width = 50
                ws_res.column_dimensions['B'].width = 25
                for row in ws_res.iter_rows(min_row=2, max_row=ws_res.max_row):
                    for cell in row:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.font = Font(size=12)
                
                ws_grille = writer.sheets["Grille de Passage"]
                ws_grille.row_dimensions[1].height = 65
                ws_grille.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nb_tapis)
                titre_cell = ws_grille.cell(row=1, column=1, value=f"🏆 {nom_competition.upper()} — PLANNING OFFICIEL  (🟡 U7 | 🟢 U9 | 🔵 U11 | 🟣 U13) 🏆")
                titre_cell.font = Font(name="Arial", size=20, bold=True, color="FFFFFF")
                titre_cell.fill = bleu
                titre_cell.alignment = Alignment(horizontal="center", vertical="center")
                
                try:
                    from openpyxl.drawing.image import Image as OpenpyxlImage
                    if os.path.exists("logo_fflda.png"):
                        img = OpenpyxlImage("logo_fflda.png")
                    else:
                        url_logo = "https://www.fflutte.com/content/uploads/2021/10/fflutte-bleu-1024x842.png"
                        req = urllib.request.Request(url_logo, headers={'User-Agent': 'Mozilla/5.0'})
                        with urllib.request.urlopen(req) as response: img_data = io.BytesIO(response.read())
                        img = OpenpyxlImage(img_data)
                    img.height, img.width = 30, 36
                    ws_grille.add_image(img, 'A1')
                except Exception: pass 

                ws_grille.freeze_panes = 'A3'
                
                for col in range(1, nb_tapis + 1):
                    c = ws_grille.cell(row=2, column=col)
                    c.fill, c.font, c.border = rouge, Font(bold=True, size=14, color="FFFFFF"), b_style
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    ws_grille.column_dimensions[c.column_letter].width = 45
                    
                for row in ws_grille.iter_rows(min_row=3, max_row=ws_grille.max_row):
                    ws_grille.row_dimensions[row[0].row].height = 90
                    for cell in row:
                        cell.border, cell.alignment = b_style, Alignment(wrap_text=True, horizontal="center", vertical="center")
                        if cell.value:
                            val_s = str(cell.value)
                            if "PAUSE" in val_s:
                                cell.fill, cell.font = rouge, Font(name="Arial", bold=True, color="FFFFFF", size=12)
                            elif any(k in val_s for k in ["Attente", "Pesée", "échauffement", "Repos"]):
                                cell.fill, cell.font = PatternFill("solid", fgColor="EFEFEF"), Font(name="Arial", italic=True, color="666666", size=11)
                            else:
                                age_k = extraire_age_de_texte(val_s)
                                cfg_c = COULEURS_AGE_GRILLE.get(age_k, COULEURS_AGE_GRILLE['AUTRE'])
                                cell.fill = cfg_c['bg_excel_pastel']
                                cell.font = Font(name="Arial", size=10, color="0F172A")

                df_excel_arb = None
                if liste_arbitres:
                    lignes_excel_arb = []
                    for t in range(nb_tapis):
                        for a in tapis_arbitres[t]:
                            lignes_excel_arb.append({
                                "Tapis Affecté": f"Tapis {t + 1}",
                                "Nom": a['Nom'],
                                "Prénom": a['Prenom'],
                                "N° Licence": a['Licence'],
                                "Club": a['Club'],
                                "Comité Régional": a['Comite']
                            })
                    if lignes_excel_arb:
                        df_excel_arb = pd.DataFrame(lignes_excel_arb)
                        df_excel_arb.to_excel(writer, sheet_name="Corps d'Arbitrage", index=False, startrow=4)
                        ws_arb_sheet = writer.sheets["Corps d'Arbitrage"]
                        ws_arb_sheet.views.sheetView[0].showGridLines = True
                        ws_arb_sheet.cell(row=1, column=1, value=f"COMPÉTITION : {nom_competition.upper()}").font = Font(name="Arial", size=15, bold=True, color="0055A4")
                        ws_arb_sheet.cell(row=2, column=1, value="CORPS D'ARBITRAGE ET AFFECTATION AUX TAPIS - FFLDA").font = Font(name="Arial", size=12, bold=True, color="666666")
                        ws_arb_sheet.cell(row=3, column=1, value=f"Édité le {datetime.now().strftime('%d/%m/%Y à %H:%M')}").font = Font(name="Arial", size=9, italic=True, color="888888")
                        
                        for col_idx in range(1, len(df_excel_arb.columns) + 1):
                            cell = ws_arb_sheet.cell(row=5, column=col_idx)
                            cell.fill, cell.font, cell.alignment = bleu, Font(name="Arial", size=10, bold=True, color="FFFFFF"), Alignment(horizontal="center", vertical="center")
                        
                        for row_idx in range(6, ws_arb_sheet.max_row + 1):
                            is_even = (row_idx % 2 == 0)
                            for col_idx in range(1, len(df_excel_arb.columns) + 1):
                                cell = ws_arb_sheet.cell(row=row_idx, column=col_idx)
                                cell.border = b_style
                                cell.fill = bleu_clair if is_even else PatternFill(fill_type=None)
                                cell.alignment = Alignment(horizontal="center", vertical="center")

                        ws_arb_sheet.column_dimensions['A'].width = 15
                        ws_arb_sheet.column_dimensions['B'].width = 25
                        ws_arb_sheet.column_dimensions['C'].width = 20
                        ws_arb_sheet.column_dimensions['D'].width = 15
                        ws_arb_sheet.column_dimensions['E'].width = 25
                        ws_arb_sheet.column_dimensions['F'].width = 25

                rouge_lutte = PatternFill("solid", fgColor="E53935") 
                bleu_lutte = PatternFill("solid", fgColor="1E88E5")  
                gris_clair = PatternFill("solid", fgColor="F2F2F2")
                entete_noir = PatternFill("solid", fgColor="000000")

                feuilles_creees = set()
                for nom_poule, liste_p in participants_par_poule.items():
                    nom_onglet_court = poule_sheet_names.get(nom_poule)
                    if not nom_onglet_court or nom_onglet_court in writer.book.sheetnames:
                        nom_base = abreger_nom_onglet(nom_poule)
                        nom_onglet_court = nom_base
                        suffix_i = 1
                        while nom_onglet_court.lower() in feuilles_creees or nom_onglet_court in writer.book.sheetnames:
                            nom_onglet_court = f"{nom_base[:28]}_{suffix_i}"
                            suffix_i += 1
                    feuilles_creees.add(nom_onglet_court.lower())
                    poule_sheet_names[nom_poule] = nom_onglet_court
                    ws_poule = writer.book.create_sheet(nom_onglet_court)
                    
                    t_num_poule = poule_tapis_map.get(nom_poule, 1)
                    p_obj = poule_obj_map.get(nom_poule)
                    if p_obj and p_obj.get('type_formule') == 'tableau':
                        construire_feuille_tableau_excel(ws_poule, p_obj, nom_poule, liste_p, coords_matchs_tapis, nom_competition, tapis_slots=tapis_slots_map, tapis_num=t_num_poule)
                    elif p_obj and p_obj.get('type_formule') == 'poules_croisees':
                        construire_feuille_poules_croisees_excel(ws_poule, p_obj, nom_poule, liste_p, coords_matchs_tapis, nom_competition, tapis_slots=tapis_slots_map, tapis_num=t_num_poule)
                    elif p_obj and p_obj.get('type_formule') == 'plateau_u7':
                        construire_feuille_plateau_u7_excel(ws_poule, p_obj, nom_poule, liste_p, coords_matchs_tapis, nom_competition, tapis_num=t_num_poule)
                    else:
                        construire_feuille_poule_nordique_excel(ws_poule, nom_poule, liste_p, rondes_par_categorie.get(nom_poule, []), coords_matchs_tapis, nom_competition, tapis_num=t_num_poule)

            excel_bytes_tournoi_complet = output_excel.getvalue()

            # Source du classeur Excel officiel FFLDA pour la génération PDF
            wb_officiel = getattr(writer, 'book', None)
            if wb_officiel is None:
                try:
                    wb_officiel = openpyxl.load_workbook(io.BytesIO(excel_bytes_tournoi_complet), data_only=False)
                except Exception:
                    wb_officiel = None

            pdf_bytes_tournoi_complet = None
            if wb_officiel is not None:
                try:
                    pdf_bytes_tournoi_complet = generer_pdf_depuis_classeur_excel(wb_officiel, nom_competition, wb=wb_officiel)
                except Exception as e_pdf:
                    st.warning(f"⚠️ Information : génération PDF depuis Excel : {e_pdf}")

            if not pdf_bytes_tournoi_complet:
                try:
                    pdf_bytes_tournoi_complet = generer_pdf_tournoi_complet("Dossier Officiel du Tournoi", nom_competition, sections_tournoi_complet)
                except Exception:
                    pdf_bytes_tournoi_complet = b""

            # --- CRÉATION DU BUNDLE ET MISE EN CACHE ---
            tournoi_bundle = {
                "nom_competition": nom_competition,
                "lignes_accueil": lignes_accueil,
                "total_participants_peses": total_participants_peses,
                "total_non_peses": total_non_peses,
                "poules_u7": poules_u7,
                "poules_u9": poules_u9,
                "poules_u11": poules_u11,
                "poules_u13": poules_u13,
                "total_matchs_calcules": total_matchs_calcules,
                "liste_arbitres": liste_arbitres,
                "tapis_arbitres": tapis_arbitres,
                "nb_tapis": nb_tapis,
                "grille_ui": grille_ui,
                "planning_tapis": planning_tapis,
                "participants_par_poule": participants_par_poule,
                "poule_obj_map": poule_obj_map,
                "excel_bytes": excel_bytes_tournoi_complet,
                "pdf_bytes": pdf_bytes_tournoi_complet,
                "matchs_direct": matchs_tournoi_direct,
                "raw_inscrits_bytes": st.session_state.get("raw_inscrits_bytes"),
                "raw_inscrits_name": st.session_state.get("raw_inscrits_name"),
                "raw_arbitres_bytes": st.session_state.get("raw_arbitres_bytes"),
                "raw_arbitres_name": st.session_state.get("raw_arbitres_name")
            }
            c_sess_courant = st.session_state.get("code_session", "ORG")
            st.session_state["tournoi_actif_cache"] = tournoi_bundle
            sauvegarder_cache_tournoi(c_sess_courant, tournoi_bundle)
            _SHARED_TOURNAMENT_MATCHS[c_sess_courant] = matchs_tournoi_direct
            _SHARED_TOURNAMENT_MATCHS["DEFAULT"] = matchs_tournoi_direct

            # Affichage des onglets interactifs et du résumé prévisionnel
            afficher_onglets_tournoi(tournoi_bundle)

            # Sauvegarde en mémoire du tournoi et des matchs pour la saisie directe (Mode 2)
            st.session_state["excel_tournoi_base"] = excel_bytes_tournoi_complet
            st.session_state["pdf_tournoi_base"] = pdf_bytes_tournoi_complet
            st.session_state["matchs_direct"] = matchs_tournoi_direct
            st.session_state["nom_competition_active"] = nom_competition

            # Synchronisation automatique en arrière-plan avec Supabase si configuré
            _, _, is_supa_rdy = get_supabase_config()
            if is_supa_rdy:
                cle_sync_tournoi = f"synced_supa_{re.sub(r'[^a-zA-Z0-9]', '_', nom_competition)}"
                if not st.session_state.get(cle_sync_tournoi):
                    del_params_gen = {}
                    if m_prefix_code not in ["FFLDA-ADMIN", "FFLDA2026"]:
                        del_params_gen["code_organisateur"] = f"eq.{m_prefix_code}"
                    else:
                        del_params_gen["tournoi_id"] = f"eq.{nom_competition}"
                    supabase_request("matchs_lutte", method="DELETE", params=del_params_gen)
                    res_p, err_p = supabase_request("matchs_lutte", method="POST", data=matchs_tournoi_direct)
                    if not err_p:
                        st.session_state[cle_sync_tournoi] = True

            st.session_state["tournoi_vient_detre_genere"] = True
            st.rerun()
        except Exception as e:
            import traceback
            st.error(f"Erreur lors de l'analyse du fichier : {e}")
            st.code(traceback.format_exc())

    else:
        # Affichage du résumé prévisionnel et des onglets pour le tournoi actif
        bundle_a_afficher = st.session_state.get("tournoi_actif_cache") or bundle_actif
        if bundle_a_afficher is not None:
            if st.session_state.pop("tournoi_vient_detre_genere", False):
                st.success(f"🎉 **Tournoi « {bundle_a_afficher.get('nom_competition', tournoi_actif_detecte or 'Tournoi')} » généré avec succès ({len(bundle_a_afficher.get('matchs_direct', []))} combats) !**")
            st.markdown(
                f"<div style='background: #e8f5e9; border-left: 5px solid #2e7d32; border-radius: 6px; padding: 10px 14px; margin-bottom: 10px;'>"
                f"<b style='color: #1b5e20;'>Fichier analysé avec succès ! Tournoi « {bundle_a_afficher.get('nom_competition', tournoi_actif_detecte or 'Tournoi')} » généré.</b>"
                f"</div>",
                unsafe_allow_html=True
            )
            st.info("👉 Pour saisir les scores ou télécharger les feuilles officielles (Excel/PDF), rendez-vous dans l'onglet **« 2. Rentrer les scores »**.")
            afficher_onglets_tournoi(bundle_a_afficher)
        else:
            st.info("👈 Veuillez importer un fichier de participants (format Exalto .csv ou .xlsx) pour lancer l'optimisation des poules et plannings.")

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
import io
import urllib.request
import re
import streamlit.components.v1 as components
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils import get_column_letter

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Générateur Officiel FFLDA", page_icon="🤼", layout="wide")

# --- MENU LATÉRAL (PARAMÈTRES INTERACTIFS) ---
with st.sidebar:
    import os
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
    with col_l2:
        if os.path.exists("logo_fflda.png"):
            st.image("logo_fflda.png", use_container_width=True)
        else:
            st.image("https://www.fflutte.com/content/uploads/2021/10/fflutte-bleu-1024x842.png", use_container_width=True)
    st.markdown("### Paramètres du tournoi")
    st.markdown("---")
    
    # --- NOM DE LA COMPÉTITION ---
    nom_competition = st.text_input("🏆 Nom de la compétition", value="Tournoi Officiel FFLDA - U7/U9/U11/U13")
    
    with st.expander("⚙️ 1. Logistique & Pesées", expanded=True):
        nb_tapis = st.number_input("Nombre de tapis", min_value=1, max_value=10, value=3)
        type_pesee = st.radio("Format des pesées", ["1 Pesée (Générale)", "2 Pesées (U7/U9 puis U11/U13)"], index=1)
        label_pesee_1 = "1ère pesée" if "1" in type_pesee else "Pesée U7/U9"
        heure_pesee_u9 = st.time_input(label_pesee_1, value=time(9, 0))
        duree_pesee = st.selectbox("Durée allouée à la pesée + échauffement (min)", [30, 45, 60, 90], index=1)
        
    with st.expander("⏱️ 2. Pause de la compétition", expanded=False):
        activer_pause = st.checkbox("Activer la pause", value=True)
        duree_pause = st.selectbox("Durée de la pause (min)", [30, 45, 60, 75, 90], index=2) if activer_pause else 0
    
    with st.expander("🤼 3. Règles Sportives & Temps", expanded=False):
        mixte_active = st.checkbox("Catégories Mixtes (U7/U9/U11/U13 ensemble)", value=True)
        poules_par_niveau = st.checkbox("Créer des poules par niveau (débutants/confirmés)", value=True)
        separer_clubs = st.checkbox("Éviter les lutteurs d'un même club dans la même poule (dans la mesure du possible)", value=True)
        eviter_arbitre_meme_club = st.checkbox("Éviter les matchs entre arbitres et lutteurs du même club", value=True)
        meme_tapis_poule = st.checkbox("Maintenir chaque poule / lutteur sur un même tapis", value=True)
        tolerance_poids = st.number_input("Tolérance d'écart de poids (%) [U7, U9, U11]", min_value=10, max_value=15, value=10, step=1)
        repos_matchs = st.number_input("Matchs de repos minimum", min_value=1, max_value=10, value=3)
        duree_u7 = st.number_input("Temps total U7 (min)", value=3)
        duree_u9 = st.number_input("Temps total U9 (min)", value=3)
        duree_u11 = st.number_input("Temps total U11 (min)", value=4)
        duree_u13 = st.number_input("Temps total U13 (min)", value=5)

    components.html("""
    
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

def charger_liste_arbitres(fichier_arbitres_in=None):
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
st.title(f"🏆 {nom_competition}")
st.markdown("**Plateforme officielle d'optimisation des tournois de jeunes et d'édition des bilans fédéraux.**")
st.markdown("---")

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
        p = float(poids_val)
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
            'type_formule': 'seul'
        }
    elif n < 6:
        nom = f"{age} | {style_grp}{suffixe_niveau} | {cat_poids} (Poule unique)"
        return {
            'nom': nom, 
            'participants': list(participants), 
            'rondes': generer_rondes_fflda(participants), 
            'type_formule': 'poule'
        }
    elif n == 6:
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
        
        sf1 = (
            {"Nom": f"1er Poule A ({cat_poids})", "Club": "Qualifié A", "Comité": "-"},
            {"Nom": f"2ème Poule B ({cat_poids})", "Club": "Qualifié B", "Comité": "-"}
        )
        sf2 = (
            {"Nom": f"1er Poule B ({cat_poids})", "Club": "Qualifié B", "Comité": "-"},
            {"Nom": f"2ème Poule A ({cat_poids})", "Club": "Qualifié A", "Comité": "-"}
        )
        r4 = [sf1, sf2]
        
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
            'poule_b': poule_b
        }
    else:
        rondes = []
        if n == 7:
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
                {"Nom": f"Perdant 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            f_b2 = (
                {"Nom": f"Perdant 1/4 (3) [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            rondes.append([f_or, f_b1, f_b2])

        elif n <= 16:
            nb_prelim = n - 8
            nb_byes = 16 - n
            
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
                {"Nom": f"Perdant 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            f_b2 = (
                {"Nom": f"Vainqueur Repêchage 2 [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            rondes.append([f_or, f_b1, f_b2])

        else:
            nb_prelim = n - 16
            nb_byes = max(0, 32 - n)
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
                {"Nom": f"Perdant 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            f_b2 = (
                {"Nom": f"Vainqueur Repêchage 2 [{cat_poids}]", "Club": "Repêché", "Comité": "-"},
                {"Nom": f"Perdant 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}
            )
            rondes.append([f_or, f_b1, f_b2])

        nom = f"{age} | {style_grp}{suffixe_niveau} | {cat_poids} (Tableau élimination & repêchages)"
        return {
            'nom': nom,
            'participants': list(participants),
            'rondes': rondes,
            'type_formule': 'tableau'
        }

# --- FONCTIONS DE GÉNÉRATION DES FORMULES EXCEL DYNAMIQUES (AVANCEMENT DU VAINQUEUR & REPÊCHÉ) ---
def make_winner_formula(c_r, c_b, pt_r, pt_b, placeholder_name, target_corner="🔴"):
    """
    Retourne la formule Excel native garantissant que dès qu'un score est saisi dans CLT,
    le lutteur ayant le plus de points avance dans le match suivant.
    """
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
    """
    Retourne la formule Excel pour faire basculer le perdant d'un quart de finale dans les repêchages.
    """
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

def draw_excel_match_card(ws, start_row, start_col, title, p1, p2, cat_poule, coords_map=None, bg_header=None):
    b_thin = Side(style='thin', color='CBD5E1')
    b_style = Border(left=b_thin, right=b_thin, top=b_thin, bottom=b_thin)
    font_match_h = Font(name="Arial", size=9, bold=True, color="FFFFFF")
    font_p_bold = Font(name="Arial", size=9, bold=True)
    font_pts = Font(name="Arial", size=10, bold=True)
    
    fill_header = bg_header if bg_header is not None else PatternFill("solid", fgColor="475569")
    fill_red_card = PatternFill("solid", fgColor="FEF2F2")
    fill_blue_card = PatternFill("solid", fgColor="EFF6FF")
    fill_gray_box = PatternFill("solid", fgColor="F8FAFC")

    col1 = start_col
    col2 = start_col + 1
    
    ws.merge_cells(start_row=start_row, start_column=col1, end_row=start_row, end_column=col2)
    c_h = ws.cell(row=start_row, column=col1, value=title)
    c_h.font = font_match_h
    c_h.fill = fill_header
    c_h.alignment = Alignment(horizontal="center", vertical="center")
    c_h.border = b_style
    ws.cell(row=start_row, column=col2).border = b_style
    
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

    # Connexion bidirectionnelle avec la Grille Tapis
    if coords_map:
        m_info = coords_map.get((cat_poule, nom1, nom2))
        if not m_info:
            m_info = coords_map.get((cat_poule, nom2, nom1))
        if m_info:
            s_name = m_info['sheet']
            if nom1 == m_info['p1']:
                ptr_c = m_info['ptr_cell']
                tot_c = m_info.get('tot_r_cell', ptr_c)
                ptb_c = m_info['ptb_cell']
                tot_b_c = m_info.get('tot_b_cell', ptb_c)
                c_pt_r.value = f"=IF('{s_name}'!{ptr_c}<>\"\", '{s_name}'!{ptr_c}, IF('{s_name}'!{tot_c}<>\"\", '{s_name}'!{tot_c}, \"\"))"
                c_pt_b.value = f"=IF('{s_name}'!{ptb_c}<>\"\", '{s_name}'!{ptb_c}, IF('{s_name}'!{tot_b_c}<>\"\", '{s_name}'!{tot_b_c}, \"\"))"
            else:
                ptr_c = m_info['ptb_cell']
                tot_c = m_info.get('tot_b_cell', ptr_c)
                ptb_c = m_info['ptr_cell']
                tot_b_c = m_info.get('tot_r_cell', ptb_c)
                c_pt_r.value = f"=IF('{s_name}'!{ptr_c}<>\"\", '{s_name}'!{ptr_c}, IF('{s_name}'!{tot_c}<>\"\", '{s_name}'!{tot_c}, \"\"))"
                c_pt_b.value = f"=IF('{s_name}'!{ptb_c}<>\"\", '{s_name}'!{ptb_c}, IF('{s_name}'!{tot_b_c}<>\"\", '{s_name}'!{tot_b_c}, \"\"))"

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
    """
    Met à jour la cellule correspondante sur la Grille Tapis avec une formule pointant
    vers la cellule du tableau U13, garantissant la propagation du vainqueur sur le tapis.
    """
    if not tapis_slots or not p_nom or not bracket_cell:
        return
    slot = tapis_slots.get((cat, p_nom))
    if not slot:
        slot = tapis_slots.get(p_nom)
    if not slot:
        for k, sl in tapis_slots.items():
            if isinstance(k, tuple):
                c, nom = k
                if (nom == p_nom or p_nom in nom or nom in p_nom) and (cat in c or c in cat):
                    slot = sl
                    break
    if not slot:
        for k, sl in tapis_slots.items():
            if isinstance(k, str) and (k == p_nom or p_nom in k or k in p_nom):
                slot = sl
                break
    if slot:
        ws_m, cell_coord = slot
        ws_m[cell_coord].value = f'=SUBSTITUTE(SUBSTITUTE(\'{ws_bracket.title}\'!{bracket_cell.coordinate}, "🔴 ", ""), "🔵 ", "")'

def construire_feuille_tableau_excel(ws, p_obj, nom_poule, liste_p, coords_matchs_tapis, nom_competition, tapis_slots=None):
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
    
    b_thin = Side(style='thin', color='CBD5E1')
    b_style = Border(left=b_thin, right=b_thin, top=b_thin, bottom=b_thin)

    ws.views.sheetView[0].showGridLines = True
    
    ws.cell(row=1, column=1, value=f"COMPÉTITION : {nom_competition.upper()} — TABLEAU OFFICIEL U13 : {nom_poule}").font = font_title
    ws.cell(row=2, column=1, value="Formule officielle FFLDA : Élimination directe avec repêchage des 1/4 de finale (2 Médailles de Bronze) — Orientation : Gauche ➔ Droite").font = font_sub

    # Table des participants inscrits
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
        
        c_pds = ws.cell(row=r, column=4, value=f"{p.get('Poids', '')} kg")
        c_pds.alignment, c_pds.border = Alignment(horizontal="center", vertical="center"), b_style
        
        if idx % 2 == 0:
            c_num.fill = fill_zebra
            c_nom.fill = fill_zebra
            c_club.fill = fill_zebra
            c_pds.fill = fill_zebra

    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 3

    rondes = p_obj.get('rondes', [])
    n = len(liste_p)
    
    f_or = rondes[-1][0]
    f_b1 = rondes[-1][1] if len(rondes[-1]) > 1 else None
    f_b2 = rondes[-1][2] if len(rondes[-1]) > 2 else None

    sf1 = rondes[-2][0]
    sf2 = rondes[-2][1]
    rep1 = rondes[-2][2] if len(rondes[-2]) > 2 else None
    rep2 = rondes[-2][3] if len(rondes[-2]) > 3 else None

    if 7 <= n <= 16 and len(rondes) >= 3:
        col_qf = 6
        col_c1 = 8
        col_sf = 9
        col_c2 = 11
        col_fn = 12
        col_pod = 14

        ws.column_dimensions['F'].width = 23
        ws.column_dimensions['G'].width = 6
        ws.column_dimensions['H'].width = 3
        ws.column_dimensions['I'].width = 23
        ws.column_dimensions['J'].width = 6
        ws.column_dimensions['K'].width = 3
        ws.column_dimensions['L'].width = 23
        ws.column_dimensions['M'].width = 6
        ws.column_dimensions['N'].width = 12
        ws.column_dimensions['O'].width = 12

        ws.merge_cells(start_row=4, start_column=col_qf, end_row=4, end_column=col_pod+1)
        c_bann = ws.cell(row=4, column=col_qf, value="🏆 TABLEAU PRINCIPAL D'ÉLIMINATION DIRECTE (OR / ARGENT)")
        c_bann.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        c_bann.fill = fill_blue
        c_bann.alignment = Alignment(horizontal="center", vertical="center")

        q_matches = rondes[-3]
        q1 = q_matches[0]
        q2 = q_matches[1]
        q3 = q_matches[2]
        q4 = q_matches[3] if len(q_matches) > 3 else (liste_p[6], {'Nom': 'EXEMPT (BYE)', 'Club': '-'})

        qf1_ptr, qf1_ptb, qf1_r, qf1_b = draw_excel_match_card(ws, 6, col_qf, "1/4 DE FINALE 1", q1[0], q1[1], nom_poule, coords_matchs_tapis)
        qf2_ptr, qf2_ptb, qf2_r, qf2_b = draw_excel_match_card(ws, 11, col_qf, "1/4 DE FINALE 2", q2[0], q2[1], nom_poule, coords_matchs_tapis)
        qf3_ptr, qf3_ptb, qf3_r, qf3_b = draw_excel_match_card(ws, 16, col_qf, "1/4 DE FINALE 3", q3[0], q3[1], nom_poule, coords_matchs_tapis)
        qf4_ptr, qf4_ptb, qf4_r, qf4_b = draw_excel_match_card(ws, 21, col_qf, "1/4 DE FINALE 4", q4[0], q4[1], nom_poule, coords_matchs_tapis)

        draw_excel_vertical_connector(ws, 7, 12, col_c1)
        ws.cell(row=10, column=col_c1).border = Border(right=Side(style='medium', color='94A3B8'), bottom=Side(style='medium', color='94A3B8'))

        draw_excel_vertical_connector(ws, 17, 22, col_c1)
        ws.cell(row=20, column=col_c1).border = Border(right=Side(style='medium', color='94A3B8'), bottom=Side(style='medium', color='94A3B8'))

        # Demi-finales dynamiques
        sf1_p1 = {'Nom': sf1[0]['Nom'], 'formula': make_winner_formula(qf1_r, qf1_b, qf1_ptr, qf1_ptb, "Vainqueur 1/4 (1)", "🔴")}
        sf1_p2 = {'Nom': sf1[1]['Nom'], 'formula': make_winner_formula(qf2_r, qf2_b, qf2_ptr, qf2_ptb, "Vainqueur 1/4 (2)", "🔵")}
        sf1_ptr, sf1_ptb, sf1_r, sf1_b = draw_excel_match_card(ws, 8, col_sf, "DEMI-FINALE 1", sf1_p1, sf1_p2, nom_poule, coords_matchs_tapis, bg_header=fill_sky)

        sf2_p1 = {'Nom': sf2[0]['Nom'], 'formula': make_winner_formula(qf3_r, qf3_b, qf3_ptr, qf3_ptb, "Vainqueur 1/4 (3)", "🔴")}
        if n == 7:
            sf2_p2 = sf2[1]
        else:
            sf2_p2 = {'Nom': sf2[1]['Nom'], 'formula': make_winner_formula(qf4_r, qf4_b, qf4_ptr, qf4_ptb, "Vainqueur 1/4 (4)", "🔵")}
        sf2_ptr, sf2_ptb, sf2_r, sf2_b = draw_excel_match_card(ws, 18, col_sf, "DEMI-FINALE 2", sf2_p1, sf2_p2, nom_poule, coords_matchs_tapis, bg_header=fill_sky)

        draw_excel_vertical_connector(ws, 10, 19, col_c2)
        ws.cell(row=14, column=col_c2).border = Border(right=Side(style='medium', color='94A3B8'), bottom=Side(style='medium', color='94A3B8'))

        # Grande Finale dynamique
        fn_p1 = {'Nom': f_or[0]['Nom'], 'formula': make_winner_formula(sf1_r, sf1_b, sf1_ptr, sf1_ptb, "Vainqueur 1/2 (1)", "🔴")}
        fn_p2 = {'Nom': f_or[1]['Nom'], 'formula': make_winner_formula(sf2_r, sf2_b, sf2_ptr, sf2_ptb, "Vainqueur 1/2 (2)", "🔵")}
        fn_ptr, fn_ptb, fn_r, fn_b = draw_excel_match_card(ws, 13, col_fn, "GRANDE FINALE (OR)", fn_p1, fn_p2, nom_poule, coords_matchs_tapis, bg_header=fill_dark)

        # Podiums Or & Argent
        r_fn = f"IF({fn_ptr.coordinate}=\"\",0,{fn_ptr.coordinate})"
        b_fn = f"IF({fn_ptb.coordinate}=\"\",0,{fn_ptb.coordinate})"
        form_gold = f'=IF({r_fn}+{b_fn}=0, "🥇 CHAMPION (OR)" & CHAR(10) & "Vainqueur Grande Finale", "🥇 CHAMPION (OR)" & CHAR(10) & SUBSTITUTE(SUBSTITUTE(IF({r_fn}>{b_fn}, {fn_r.coordinate}, IF({b_fn}>{r_fn}, {fn_b.coordinate}, "En attente")), "🔴 ", ""), "🔵 ", ""))'
        form_silver = f'=IF({r_fn}+{b_fn}=0, "🥈 VICE-CHAMPION (ARGENT)" & CHAR(10) & "Perdant Grande Finale", "🥈 VICE-CHAMPION (ARGENT)" & CHAR(10) & SUBSTITUTE(SUBSTITUTE(IF({r_fn}>{b_fn}, {fn_b.coordinate}, IF({b_fn}>{r_fn}, {fn_r.coordinate}, "En attente")), "🔴 ", ""), "🔵 ", ""))'
        draw_excel_podium_card(ws, 12, col_pod, "🥇 CHAMPION (OR)", "Vainqueur Grande Finale", fill_gold, font_color="B45309", formula_val=form_gold)
        draw_excel_podium_card(ws, 15, col_pod, "🥈 VICE-CHAMPION (ARGENT)", "Perdant Grande Finale", fill_silver, font_color="475569", formula_val=form_silver)

        # Repêchages & Bronze
        row_rep = 26
        ws.merge_cells(start_row=row_rep, start_column=col_qf, end_row=row_rep, end_column=col_pod+1)
        c_rep_h = ws.cell(row=row_rep, column=col_qf, value="🔄 TABLEAU DE REPÊCHAGE & MATCHS POUR LE BRONZE (2 Troisièmes Places)")
        c_rep_h.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        c_rep_h.fill = fill_amber
        c_rep_h.alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells(start_row=row_rep+1, start_column=col_qf, end_row=row_rep+1, end_column=col_pod+1)
        c_rep_sub = ws.cell(row=row_rep+1, column=col_qf, value="Perdants des 1/4 ➔ Repêchages ➔ Finales Bronze contre les perdants des Demi-Finales")
        c_rep_sub.font = font_sub
        c_rep_sub.alignment = Alignment(horizontal="left", vertical="center")

        # Repêchage 1
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

        ws.cell(row=row_rep+4, column=col_c1).border = Border(bottom=Side(style='medium', color='94A3B8'))
        ws.cell(row=row_rep+9, column=col_c1).border = Border(bottom=Side(style='medium', color='94A3B8'))

        # Finale Bronze 1
        fb1_p1 = {'Nom': f_b1[0]['Nom'], 'formula': make_winner_formula(rep1_r, rep1_b, rep1_ptr, rep1_ptb, "Vainqueur Repêchage 1", "🔴")}
        fb1_p2 = {'Nom': f_b1[1]['Nom'], 'formula': make_loser_formula(sf2_r, sf2_b, sf2_ptr, sf2_ptb, "Perdant 1/2 (2)", "🔵")}
        b1_ptr, b1_ptb, b1_r, b1_b = draw_excel_match_card(ws, row_rep+3, col_sf, "FINALE BRONZE 1", fb1_p1, fb1_p2, nom_poule, coords_matchs_tapis, bg_header=fill_amber)

        r_b1 = f"IF({b1_ptr.coordinate}=\"\",0,{b1_ptr.coordinate})"
        b_b1 = f"IF({b1_ptb.coordinate}=\"\",0,{b1_ptb.coordinate})"
        form_b1 = f'=IF({r_b1}+{b_b1}=0, "🥉 3ème PLACE (Bronze 1)" & CHAR(10) & "Vainqueur Finale Bronze 1", "🥉 3ème PLACE (Bronze 1)" & CHAR(10) & SUBSTITUTE(SUBSTITUTE(IF({r_b1}>{b_b1}, {b1_r.coordinate}, IF({b_b1}>{r_b1}, {b1_b.coordinate}, "En attente")), "🔴 ", ""), "🔵 ", ""))'
        draw_excel_podium_card(ws, row_rep+3, col_pod, "🥉 3ème PLACE (Bronze 1)", "Vainqueur Finale Bronze 1", fill_bronze, font_color="9A3412", formula_val=form_b1)

        # Finale Bronze 2
        if n == 7:
            fb2_p1 = {'Nom': f_b2[0]['Nom'], 'formula': make_loser_formula(qf3_r, qf3_b, qf3_ptr, qf3_ptb, "Perdant 1/4 (3)", "🔴")}
        else:
            fb2_p1 = {'Nom': f_b2[0]['Nom'], 'formula': make_winner_formula(rep2_r, rep2_b, rep2_ptr, rep2_ptb, "Vainqueur Repêchage 2", "🔴")}
        fb2_p2 = {'Nom': f_b2[1]['Nom'], 'formula': make_loser_formula(sf1_r, sf1_b, sf1_ptr, sf1_ptb, "Perdant 1/2 (1)", "🔵")}
        b2_ptr, b2_ptb, b2_r, b2_b = draw_excel_match_card(ws, row_rep+8, col_sf, "FINALE BRONZE 2", fb2_p1, fb2_p2, nom_poule, coords_matchs_tapis, bg_header=fill_amber)

        r_b2 = f"IF({b2_ptr.coordinate}=\"\",0,{b2_ptr.coordinate})"
        b_b2 = f"IF({b2_ptb.coordinate}=\"\",0,{b2_ptb.coordinate})"
        form_b2 = f'=IF({r_b2}+{b_b2}=0, "🥉 3ème PLACE (Bronze 2)" & CHAR(10) & "Vainqueur Finale Bronze 2", "🥉 3ème PLACE (Bronze 2)" & CHAR(10) & SUBSTITUTE(SUBSTITUTE(IF({r_b2}>{b_b2}, {b2_r.coordinate}, IF({b_b2}>{r_b2}, {b2_b.coordinate}, "En attente")), "🔴 ", ""), "🔵 ", ""))'
        draw_excel_podium_card(ws, row_rep+8, col_pod, "🥉 3ème PLACE (Bronze 2)", "Vainqueur Finale Bronze 2", fill_bronze, font_color="9A3412", formula_val=form_b2)

        # Liaison automatique vers les cartes de match de la Grille Tapis
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
            
            ws.column_dimensions[get_column_letter(col_match)].width = 23
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
        ws.column_dimensions[get_column_letter(col_pod)].width = 12
        ws.column_dimensions[get_column_letter(col_pod+1)].width = 12
        draw_excel_podium_card(ws, 6, col_pod, "🥇 CHAMPION (OR)", "Vainqueur Finale", fill_gold, font_color="B45309")
        draw_excel_podium_card(ws, 9, col_pod, "🥈 VICE-CHAMPION", "Finaliste", fill_silver, font_color="475569")
        draw_excel_podium_card(ws, 12, col_pod, "🥉 3ème PLACE (1)", "Bronze 1", fill_bronze, font_color="9A3412")
        draw_excel_podium_card(ws, 15, col_pod, "🥉 3ème PLACE (2)", "Bronze 2", fill_bronze, font_color="9A3412")

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

def construire_feuille_poules_croisees_excel(ws, p_obj, nom_poule, liste_p, coords_matchs_tapis, nom_competition, tapis_slots=None):
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
    
    b_thin = Side(style='thin', color='CBD5E1')
    b_style = Border(left=b_thin, right=b_thin, top=b_thin, bottom=b_thin)

    ws.views.sheetView[0].showGridLines = True
    
    ws.cell(row=1, column=1, value=f"COMPÉTITION : {nom_competition.upper()} — POULES CROISÉES U13 (6 LUTTEURS) : {nom_poule}").font = font_title
    ws.cell(row=2, column=1, value="Formule officielle FFLDA : Phase 1 (2 Poules de 3 Nordiques) ➔ Phase 2 (Demi-Finales Croisées & Finales Or/Argent et Bronze unique)").font = font_sub

    poule_a = p_obj.get('poule_a', liste_p[:3])
    poule_b = p_obj.get('poule_b', liste_p[3:])
    rondes = p_obj.get('rondes', [])

    lignes_lutteurs = {}

    def render_sub_poule_table(ws, start_row, title, participants, bg_color):
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=8)
        c_title = ws.cell(row=start_row, column=1, value=title)
        c_title.font, c_title.fill, c_title.alignment = font_hdr, bg_color, Alignment(horizontal="center", vertical="center")
        
        headers = ["CLT", "N°", "NOM Prénom", "CLUB", "Tour 1", "Tour 2", "Tour 3", "Total Pts"]
        for c_i, h in enumerate(headers, 1):
            c = ws.cell(row=start_row+1, column=c_i, value=h)
            c.font, c.fill, c.alignment, c.border = font_match_h, fill_gray_h, Alignment(horizontal="center", vertical="center"), b_style
            
        row_cur = start_row + 2
        for idx, p in enumerate(participants, 1):
            lignes_lutteurs[p['Nom']] = row_cur
            
            c_clt = ws.cell(row=row_cur, column=1, value=f"=RANK(H{row_cur}, H{start_row+2}:H{start_row+1+len(participants)})")
            c_clt.alignment, c_clt.border = Alignment(horizontal="center", vertical="center"), b_style
            c_clt.font = Font(name="Arial", size=10, bold=True, color="0055A4")
            
            c_num = ws.cell(row=row_cur, column=2, value=idx)
            c_num.alignment, c_num.border = Alignment(horizontal="center", vertical="center"), b_style
            
            c_nom = ws.cell(row=row_cur, column=3, value=p['Nom'])
            c_nom.border = b_style
            
            c_club = ws.cell(row=row_cur, column=4, value=p.get('Club', ''))
            c_club.border = b_style
            
            for t_i in range(3):
                c_t = ws.cell(row=row_cur, column=5+t_i)
                c_t.alignment, c_t.border = Alignment(horizontal="center", vertical="center"), b_style
            
            c_tot = ws.cell(row=row_cur, column=8, value=f"=SUM(E{row_cur}:G{row_cur})")
            c_tot.font, c_tot.alignment, c_tot.border = font_pts, Alignment(horizontal="center", vertical="center"), b_style
            
            row_cur += 1

    render_sub_poule_table(ws, 4, "🥋 PHASE 1 : POULE A (3 Lutteurs)", poule_a, fill_blue)
    render_sub_poule_table(ws, 10, "🥋 PHASE 1 : POULE B (3 Lutteurs)", poule_b, fill_dark)

    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 5
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 8
    ws.column_dimensions['F'].width = 8
    ws.column_dimensions['G'].width = 8
    ws.column_dimensions['H'].width = 10
    ws.column_dimensions['I'].width = 3

    r_matches = 16
    ws.merge_cells(start_row=r_matches, start_column=1, end_row=r_matches, end_column=8)
    c_m_hdr = ws.cell(row=r_matches, column=1, value="🤼 RENCONTRES DES POULES A & B (TOURS 1 À 3)")
    c_m_hdr.font, c_m_hdr.fill, c_m_hdr.alignment = font_hdr, fill_gray_h, Alignment(horizontal="center", vertical="center")
    
    r_matches += 1
    for tour_idx in range(3):
        if tour_idx < len(rondes):
            ronde = rondes[tour_idx]
            for m_idx, m in enumerate(ronde):
                grp_tag = "Poule A" if m_idx == 0 else "Poule B"
                title_m = f"T{tour_idx+1} ({grp_tag})"
                pt_r, pt_b, _, _ = draw_excel_match_card(ws, r_matches, 1, title_m, m[0], m[1], nom_poule, coords_matchs_tapis)
                
                p1_nom = m[0]['Nom']
                p2_nom = m[1]['Nom']
                if p1_nom in lignes_lutteurs:
                    ws.cell(row=lignes_lutteurs[p1_nom], column=5+tour_idx, value=f"={pt_r.coordinate}")
                if p2_nom in lignes_lutteurs:
                    ws.cell(row=lignes_lutteurs[p2_nom], column=5+tour_idx, value=f"={pt_b.coordinate}")
                
                r_matches += 4

    # Phase 2 : Demi-Finales Croisées & Finales
    col_sf = 10
    col_c = 12
    col_fn = 13
    col_pod = 15

    ws.column_dimensions['J'].width = 23
    ws.column_dimensions['K'].width = 6
    ws.column_dimensions['L'].width = 3
    ws.column_dimensions['M'].width = 23
    ws.column_dimensions['N'].width = 6
    ws.column_dimensions['O'].width = 12
    ws.column_dimensions['P'].width = 12

    ws.merge_cells(start_row=4, start_column=col_sf, end_row=4, end_column=col_pod+1)
    c_fin_h = ws.cell(row=4, column=col_sf, value="🏆 PHASE 2 : PHASE FINALE CROISÉE (Gauche ➔ Droite)")
    c_fin_h.font, c_fin_h.fill, c_fin_h.alignment = font_hdr, fill_sky, Alignment(horizontal="center", vertical="center")

    sf1 = rondes[3][0]
    sf2 = rondes[3][1]
    f_or = rondes[4][0]
    f_b = rondes[4][1]

    # Demi-finales croisées alimentées dynamiquement depuis le classement des poules A et B
    form_1er_a = '=IF(ISNA(MATCH(1, A$6:A$8, 0)), "🔴 1er Poule A", "🔴 " & INDEX(C$6:C$8, MATCH(1, A$6:A$8, 0)))'
    form_2e_b  = '=IF(ISNA(MATCH(2, A$12:A$14, 0)), "🔵 2ème Poule B", "🔵 " & INDEX(C$12:C$14, MATCH(2, A$12:A$14, 0)))'
    form_1er_b = '=IF(ISNA(MATCH(1, A$12:A$14, 0)), "🔴 1er Poule B", "🔴 " & INDEX(C$12:C$14, MATCH(1, A$12:A$14, 0)))'
    form_2e_a  = '=IF(ISNA(MATCH(2, A$6:A$8, 0)), "🔵 2ème Poule A", "🔵 " & INDEX(C$6:C$8, MATCH(2, A$6:A$8, 0)))'

    p_sf1_1 = {'Nom': sf1[0]['Nom'], 'formula': form_1er_a}
    p_sf1_2 = {'Nom': sf1[1]['Nom'], 'formula': form_2e_b}
    sf1_ptr, sf1_ptb, sf1_r, sf1_b = draw_excel_match_card(ws, 6, col_sf, "DEMI-FINALE 1 (1er A vs 2ème B)", p_sf1_1, p_sf1_2, nom_poule, coords_matchs_tapis, bg_header=fill_blue)

    p_sf2_1 = {'Nom': sf2[0]['Nom'], 'formula': form_1er_b}
    p_sf2_2 = {'Nom': sf2[1]['Nom'], 'formula': form_2e_a}
    sf2_ptr, sf2_ptb, sf2_r, sf2_b = draw_excel_match_card(ws, 12, col_sf, "DEMI-FINALE 2 (1er B vs 2ème A)", p_sf2_1, p_sf2_2, nom_poule, coords_matchs_tapis, bg_header=fill_blue)

    draw_excel_vertical_connector(ws, 7, 13, col_c)
    ws.cell(row=8, column=col_c).border = Border(right=Side(style='medium', color='94A3B8'), bottom=Side(style='medium', color='94A3B8'))
    ws.cell(row=14, column=col_c).border = Border(right=Side(style='medium', color='94A3B8'), bottom=Side(style='medium', color='94A3B8'))

    # Finales croisées dynamiques
    f_or_p1 = {'Nom': f_or[0]['Nom'], 'formula': make_winner_formula(sf1_r, sf1_b, sf1_ptr, sf1_ptb, "Vainqueur SF1", "🔴")}
    f_or_p2 = {'Nom': f_or[1]['Nom'], 'formula': make_winner_formula(sf2_r, sf2_b, sf2_ptr, sf2_ptb, "Vainqueur SF2", "🔵")}
    for_ptr, for_ptb, for_r, for_b = draw_excel_match_card(ws, 7, col_fn, "FINALE 1-2 (OR / ARGENT)", f_or_p1, f_or_p2, nom_poule, coords_matchs_tapis, bg_header=fill_dark)

    f_b_p1 = {'Nom': f_b[0]['Nom'], 'formula': make_loser_formula(sf1_r, sf1_b, sf1_ptr, sf1_ptb, "Perdant SF1", "🔴")}
    f_b_p2 = {'Nom': f_b[1]['Nom'], 'formula': make_loser_formula(sf2_r, sf2_b, sf2_ptr, sf2_ptb, "Perdant SF2", "🔵")}
    fb_ptr, fb_ptb, fb_r, fb_b = draw_excel_match_card(ws, 13, col_fn, "FINALE 3-4 (BRONZE UNIQUE)", f_b_p1, f_b_p2, nom_poule, coords_matchs_tapis, bg_header=fill_amber)

    # Podiums croisés dynamiques
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

    # Propagation vers les feuilles Tapis
    link_tapis_slot(tapis_slots, nom_poule, sf1[0]['Nom'], ws, sf1_r)
    link_tapis_slot(tapis_slots, nom_poule, sf1[1]['Nom'], ws, sf1_b)
    link_tapis_slot(tapis_slots, nom_poule, sf2[0]['Nom'], ws, sf2_r)
    link_tapis_slot(tapis_slots, nom_poule, sf2[1]['Nom'], ws, sf2_b)

    link_tapis_slot(tapis_slots, nom_poule, f_or[0]['Nom'], ws, for_r)
    link_tapis_slot(tapis_slots, nom_poule, f_or[1]['Nom'], ws, for_b)

    link_tapis_slot(tapis_slots, nom_poule, f_b[0]['Nom'], ws, fb_r)
    link_tapis_slot(tapis_slots, nom_poule, f_b[1]['Nom'], ws, fb_b)

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

# --- TABLEAUX HTML DES BRACKETS VISUELS U13 POUR STREAMLIT ---
def make_bracket_card(p1, p2, title='', badge=''):
    nom1 = p1.get('Nom', 'Lutteur 1')
    c1 = p1.get('Club', '')
    club1_str = f" ({c1})" if c1 and c1 not in ['-', 'Comité Non Renseigné', ''] else ''
    
    nom2 = p2.get('Nom', 'Lutteur 2')
    c2 = p2.get('Club', '')
    club2_str = f" ({c2})" if c2 and c2 not in ['-', 'Comité Non Renseigné', ''] else ''

    badge_html = f'{badge}' if badge else ''

    is_bye1 = any(k in str(nom1).lower() for k in ['bye', 'exempt'])
    is_bye2 = any(k in str(nom2).lower() for k in ['bye', 'exempt'])

    dot1 = '#94A3B8' if is_bye1 else '#EF4135'
    dot2 = '#94A3B8' if is_bye2 else '#0055A4'

    return f'''

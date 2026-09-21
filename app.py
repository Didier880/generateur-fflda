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

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Générateur Officiel FFLDA", page_icon="🤼", layout="wide")

# --- MENU LATÉRAL (PARAMÈTRES INTERACTIFS) ---
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/fr/thumb/5/58/Logo_F%C3%A9d%C3%A9ration_Fran%C3%A7aise_de_Lutte.svg/1200px-Logo_F%C3%A9d%C3%A9ration_Fran%C3%A7aise_de_Lutte.svg.png", use_container_width=True)
    st.markdown("### Paramètres FFLDA")
    st.markdown("---")
    
    # --- NOM DE LA COMPÉTITION ---
    nom_competition = st.text_input("🏆 Nom de la compétition", value="Tournoi Officiel FFLDA - U9/U11")
    
    with st.expander("⚙️ 1. Logistique & Pesées", expanded=True):
        nb_tapis = st.number_input("Nombre de tapis", min_value=1, max_value=10, value=3)
        type_pesee = st.radio("Format des pesées", ["1 Pesée (Générale)", "2 Pesées (U9 puis U11)"], index=1)
        label_pesee_1 = "1ère pesée" if "1" in type_pesee else "Pesée U9"
        heure_pesee_u9 = st.time_input(label_pesee_1, value=time(9, 0))
        duree_pesee = st.selectbox("Durée allouée à la pesée + échauffement (min)", [30, 45, 60, 90], index=1)
        
    with st.expander("⏱️ 2. Pause de la compétition", expanded=False):
        activer_pause = st.checkbox("Activer la pause", value=True)
        duree_pause = st.selectbox("Durée de la pause (min)", [30, 45, 60, 75, 90], index=2) if activer_pause else 0
    
    with st.expander("🤼 3. Règles Sportives & Temps", expanded=False):
        mixte_active = st.checkbox("Catégories Mixtes (U9/U11 ensemble)", value=True)
        poules_par_niveau = st.checkbox("Créer des poules par niveau (débutants/confirmés)", value=True)
        separer_clubs = st.checkbox("Éviter les lutteurs d'un même club dans la même poule (dans la mesure du possible)", value=True)
        meme_tapis_poule = st.checkbox("Maintenir chaque poule / lutteur sur un même tapis", value=True)
        tolerance_poids = st.number_input("Tolérance d'écart de poids (%)", min_value=10, max_value=15, value=10, step=1)
        repos_matchs = st.number_input("Matchs de repos minimum", min_value=1, max_value=10, value=3)
        duree_u9 = st.number_input("Temps total U9 (min)", value=3)
        duree_u11 = st.number_input("Temps total U11 (min)", value=4)

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
    # Conserver les termes "Confirmé" et "Débutant" en entier
    # Supprimer le terme Gr. / Groupe / Poule et le numéro qui suit s'ils sont présents
    txt = re.sub(r'\b(Gr\.|Gr|Groupe|Poule)\s*\d+\b', '', txt, flags=re.IGNORECASE)
    txt = txt.replace(" | ", " ").replace(" (", " ").replace(")", "").replace(" - ", "-")
    txt = txt.replace("/", "-").replace("\\", "-").replace(":", "-").replace("?", "").replace("*", "")
    txt = re.sub(r'\s+', ' ', txt)
    return txt[:31].strip()

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
                    cand = sorted(poules[i-1]['participants'] + [p_iso], key=lambda x: x['Poids_Num'])
                    if len(cand) <= max_size and poule_poids_valide(cand, multiplicateur_poids):
                        poules[i-1]['participants'] = cand
                        poules[i-1]['rondes'] = generer_rondes_fflda(cand)
                        p_min = cand[0]['Poids_Num']
                        p_max = cand[-1]['Poids_Num']
                        prefix = poules[i-1]['nom'].split(' (')[0]
                        poules[i-1]['nom'] = f"{prefix} ({formater_poids_court(p_min)} - {formater_poids_court(p_max)})"
                        poules.pop(i)
                        fusionne = True
                        ameliore = True
                        break
                
                # Option 2: Essayer de fusionner directement avec la poule suivante (i+1)
                if not fusionne and i < len(poules) - 1:
                    cand = sorted([p_iso] + poules[i+1]['participants'], key=lambda x: x['Poids_Num'])
                    if len(cand) <= max_size and poule_poids_valide(cand, multiplicateur_poids):
                        poules[i+1]['participants'] = cand
                        poules[i+1]['rondes'] = generer_rondes_fflda(cand)
                        p_min = cand[0]['Poids_Num']
                        p_max = cand[-1]['Poids_Num']
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
                        new_prev = sorted(prev_parts[:-k], key=lambda x: x['Poids_Num'])
                        new_iso = sorted(prev_parts[-k:] + [p_iso], key=lambda x: x['Poids_Num'])
                        
                        if (len(new_prev) >= 2 and len(new_iso) >= 2 and len(new_iso) <= max_size and
                            poule_poids_valide(new_prev, multiplicateur_poids) and
                            poule_poids_valide(new_iso, multiplicateur_poids)):
                            
                            poules[i-1]['participants'] = new_prev
                            poules[i-1]['rondes'] = generer_rondes_fflda(new_prev)
                            p_min1 = new_prev[0]['Poids_Num']
                            p_max1 = new_prev[-1]['Poids_Num']
                            prefix1 = poules[i-1]['nom'].split(' (')[0]
                            poules[i-1]['nom'] = f"{prefix1} ({formater_poids_court(p_min1)} - {formater_poids_court(p_max1)})"
                            
                            poules[i]['participants'] = new_iso
                            poules[i]['rondes'] = generer_rondes_fflda(new_iso)
                            p_min2 = new_iso[0]['Poids_Num']
                            p_max2 = new_iso[-1]['Poids_Num']
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
                        new_iso = sorted([p_iso] + next_parts[:k], key=lambda x: x['Poids_Num'])
                        new_next = sorted(next_parts[k:], key=lambda x: x['Poids_Num'])
                        
                        if (len(new_iso) >= 2 and len(new_next) >= 2 and len(new_iso) <= max_size and
                            poule_poids_valide(new_iso, multiplicateur_poids) and
                            poule_poids_valide(new_next, multiplicateur_poids)):
                            
                            poules[i]['participants'] = new_iso
                            poules[i]['rondes'] = generer_rondes_fflda(new_iso)
                            p_min1 = new_iso[0]['Poids_Num']
                            p_max1 = new_iso[-1]['Poids_Num']
                            prefix1 = poules[i]['nom'].split(' (')[0]
                            poules[i]['nom'] = f"{prefix1} ({formater_poids_court(p_min1)} - {formater_poids_court(p_max1)})"
                            
                            poules[i+1]['participants'] = new_next
                            poules[i+1]['rondes'] = generer_rondes_fflda(new_next)
                            p_min2 = new_next[0]['Poids_Num']
                            p_max2 = new_next[-1]['Poids_Num']
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
    poids_list = [p['Poids_Num'] for p in participants_poule if p.get('Poids_Num', 0) > 0]
    if not poids_list:
        return True
    p_min = min(poids_list)
    p_max = max(poids_list)
    return p_max <= round(p_min * multiplicateur_poids, 4)

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
                    poules_groupe[i]['participants'] = sorted(p1_test, key=lambda x: x['Poids_Num'])
                    poules_groupe[j]['participants'] = sorted(p2_test, key=lambda x: x['Poids_Num'])
                    
                    for idx_p in [i, j]:
                        parts = poules_groupe[idx_p]['participants']
                        prefix = poules_groupe[idx_p]['nom'].split(' (')[0]
                        p_min = parts[0]['Poids_Num']
                        p_max = parts[-1]['Poids_Num']
                        poules_groupe[idx_p]['nom'] = f"{prefix} ({formater_poids_court(p_min)} - {formater_poids_court(p_max)})"
                        poules_groupe[idx_p]['rondes'] = generer_rondes_fflda(parts)
                    
                    ameliore = True
                    break
            if ameliore:
                break
                
    return poules_groupe

# --- GÉNÉRATEUR DE DOCUMENTS HTML AUTONOMES POUR IMPRESSION PAYSAGE A4 ---
def generer_document_html_imprimable(titre, nom_comp, sections):
    """
    Génère un document HTML 100% autonome prêt pour l'impression A4 Paysage.
    Toutes les tables possèdent page-break-inside: avoid pour ne jamais se couper.
    """
    html_sections = []
    for section_title, content in sections:
        if isinstance(content, pd.DataFrame):
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
            page-break-inside: avoid !important;
            break-inside: avoid-page !important;
            break-inside: avoid !important;
            margin-bottom: 30px;
            width: 100%;
            clear: both;
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

# Sélection du mode de travail
col_mode1, col_mode2 = st.columns([2, 3])
with col_mode1:
    mode_app = st.selectbox("📌 Sélectionnez le mode de travail", ["1. Générer un Tournoi (Planning & Poules)", "2. Importer les scores & Éditer le Bilan (Excel)"])

if mode_app.startswith("2"):
    st.markdown("### 📂 Module de Fin de Tournoi & Bilans Fédéraux")
    st.info(f"Importez votre fichier Excel complété pour la compétition **{nom_competition}** afin de générer automatiquement les classements officiels individuels et par club.")
    
    fichier_resultats = st.file_uploader("Sélectionner le fichier Excel complété (.xlsx)", type=["xlsx"])
    
    if fichier_resultats is not None:
        try:
            wb_res = openpyxl.load_workbook(fichier_resultats, data_only=True)
            st.success("✨ Fichier de résultats analysé avec succès !")
            
            tous_les_resultats = []
            onglets_poules = [f for f in wb_res.sheetnames if not any(x in f for x in ["Résumé", "Grille", "Classement Général"])]
            onglets_poules.sort(key=lambda x: (0 if "U9" in x.upper() else 1, x))
            
            for nom_feuille in onglets_poules:
                ws = wb_res[nom_feuille]
                
                # --- REPÉRAGE DYNAMIQUE DES COLONNES (Ligne 4 - En-têtes) ---
                col_total_pts = None
                col_poids = None
                col_comite = None
                for c_idx in range(1, ws.max_column + 1):
                    val_head = str(ws.cell(row=4, column=c_idx).value or "").strip()
                    val_lower = val_head.lower()
                    if val_head == "Total Pts":
                        col_total_pts = c_idx
                    elif val_head == "Poids":
                        col_poids = c_idx
                    elif any(k in val_lower for k in ["comité", "comite", "ligue", "région", "region", "c.r."]):
                        col_comite = c_idx
                
                r = 5
                lutteurs_poule = []
                while ws.cell(row=r, column=3).value is not None:
                    nom = ws.cell(row=r, column=3).value
                    club = ws.cell(row=r, column=4).value
                    comite_val = ws.cell(row=r, column=col_comite).value if col_comite else None
                    comite = str(comite_val).strip() if (comite_val and str(comite_val).strip() not in ["", "None", "nan", "-"]) else "Comité Non Renseigné"
                    
                    total_pts = ws.cell(row=r, column=col_total_pts).value if col_total_pts else 0
                    poids_raw = ws.cell(row=r, column=col_poids).value if col_poids else 0
                    
                    try:
                        pts_val = int(round(float(total_pts))) if total_pts is not None else 0
                    except (ValueError, TypeError):
                        pts_val = 0

                    def formater_poids(val):
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

                    poids_val = formater_poids(poids_raw)

                    lutteurs_poule.append({
                        "Poule": nom_feuille,
                        "Nom": nom,
                        "Club": club if club else "Indépendant",
                        "Comité": comite,
                        "Poids": poids_val,
                        "Points": pts_val
                    })
                    r += 1
                
                df_poule = pd.DataFrame(lutteurs_poule)
                if not df_poule.empty:
                    if df_poule["Points"].sum() == 0:
                        df_poule["Clt"] = "NR"
                    else:
                        df_poule = df_poule.sort_values(by="Points", ascending=False).reset_index(drop=True)
                        rangs = []
                        current_rang = 1
                        for idx, row in df_poule.iterrows():
                            if idx > 0 and row["Points"] == df_poule.iloc[idx-1]["Points"]:
                                rangs.append(rangs[-1])
                            else:
                                rangs.append(current_rang)
                            current_rang += 1
                        df_poule["Clt"] = rangs
                        
                    tous_les_resultats.extend(df_poule.to_dict('records'))

            df_bilan = pd.DataFrame(tous_les_resultats)
            
            # --- CALCUL DU CLASSEMENT DES CLUBS ET DES COMITÉS RÉGIONAUX ---
            bareme_points = {1: 4, 2: 3, 3: 2, 4: 1}
            points_clubs = {}
            points_comites = {}

            for _, row in df_bilan.iterrows():
                club = row["Club"]
                comite = row.get("Comité", "Comité Non Renseigné")
                clt = row["Clt"]
                if clt == "NR":
                    continue
                pts_attribués = bareme_points.get(int(clt), 0)
                
                # Ranking Clubs
                if club not in points_clubs:
                    points_clubs[club] = {"Club": club, "Points Club": 0, "1ers": 0, "2èmes": 0, "3èmes": 0, "4èmes": 0}
                points_clubs[club]["Points Club"] += pts_attribués
                if int(clt) == 1: points_clubs[club]["1ers"] += 1
                elif int(clt) == 2: points_clubs[club]["2èmes"] += 1
                elif int(clt) == 3: points_clubs[club]["3èmes"] += 1
                elif int(clt) == 4: points_clubs[club]["4èmes"] += 1

                # Ranking Comités Régionaux
                if comite not in points_comites:
                    points_comites[comite] = {"Comité Régional": comite, "Points Comité": 0, "1ers": 0, "2èmes": 0, "3èmes": 0, "4èmes": 0}
                points_comites[comite]["Points Comité"] += pts_attribués
                if int(clt) == 1: points_comites[comite]["1ers"] += 1
                elif int(clt) == 2: points_comites[comite]["2èmes"] += 1
                elif int(clt) == 3: points_comites[comite]["3èmes"] += 1
                elif int(clt) == 4: points_comites[comite]["4èmes"] += 1

            df_clubs = pd.DataFrame(list(points_clubs.values())).sort_values(
                by=["Points Club", "1ers", "2èmes", "3èmes", "4èmes"], 
                ascending=False
            ).reset_index(drop=True)
            df_clubs.index = range(1, len(df_clubs) + 1)
            df_clubs.insert(0, "Clt Club", df_clubs.index)

            df_comites = pd.DataFrame(list(points_comites.values())).sort_values(
                by=["Points Comité", "1ers", "2èmes", "3èmes", "4èmes"], 
                ascending=False
            ).reset_index(drop=True)
            if not df_comites.empty:
                df_comites.index = range(1, len(df_comites) + 1)
                df_comites.insert(0, "Clt Comité", df_comites.index)

            # --- GÉNÉRATION DU DOCUMENT HTML PAYSAGE POUR IMPRESSION DES BILANS ---
            tab_bilan_1, tab_bilan_2, tab_bilan_3 = st.tabs([
                "🏆 Classements Individuels (U9 / U11)", 
                "🛡️ Classement Général des Clubs",
                "🏛️ Classement des Comités Régionaux"
            ])
            
            with tab_bilan_1:
                st.subheader("Classements Individuels Officiels")
                for poule in df_bilan['Poule'].unique():
                    st.markdown(f"#### 🤼 {poule}")
                    sous_df = df_bilan[df_bilan['Poule'] == poule][['Clt', 'Nom', 'Club', 'Poids', 'Points']].copy()
                    sous_df['Points'] = sous_df['Points'].astype(int)
                    st.table(sous_df)

            with tab_bilan_2:
                st.subheader("🛡️ Podium des Clubs Engagés")
                st.markdown("*Barème officiel : 1er = 4 pts | 2ème = 3 pts | 3ème = 2 pts | 4ème = 1 pt*")
                st.table(df_clubs)

            with tab_bilan_3:
                st.subheader("🏛️ Classement Officiel des Comités Régionaux")
                st.markdown("*Barème officiel : 1er = 4 pts | 2ème = 3 pts | 3ème = 2 pts | 4ème = 1 pt*")
                st.table(df_comites)

            output_bilan = io.BytesIO()
            with pd.ExcelWriter(output_bilan, engine='openpyxl') as writer:
                df_clubs.to_excel(writer, sheet_name="Classement Clubs", index=False, startrow=5)
                df_comites.to_excel(writer, sheet_name="Classement Comités", index=False, startrow=5)
                ws_indiv = writer.book.create_sheet("Classements Individuels")
                
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
                
                # Feuille Classement Clubs
                ws_clubs = writer.sheets["Classement Clubs"]
                ws_clubs.views.sheetView[0].showGridLines = True
                ws_clubs.cell(row=1, column=1, value=f"COMPÉTITION : {nom_competition.upper()}").font = font_titre
                ws_clubs.cell(row=2, column=1, value="🛡️ CLASSEMENT OFFICIEL DES CLUBS - FFLDA").font = Font(name="Arial", size=12, bold=True, color="666666")
                ws_clubs.cell(row=3, column=1, value=f"Édité le {datetime.now().strftime('%d/%m/%Y à %H:%M')}").font = Font(name="Arial", size=9, italic=True, color="888888")
                
                for col_idx in range(1, len(df_clubs.columns) + 1):
                    cell = ws_clubs.cell(row=5, column=col_idx)
                    cell.fill, cell.font, cell.alignment = bleu_fflda, font_entete, Alignment(horizontal="center", vertical="center")
                    ws_clubs.row_dimensions[5].height = 25
                
                for row_idx in range(6, ws_clubs.max_row + 1):
                    ws_clubs.row_dimensions[row_idx].height = 22
                    is_even = (row_idx % 2 == 0)
                    for col_idx in range(1, len(df_clubs.columns) + 1):
                        cell = ws_clubs.cell(row=row_idx, column=col_idx)
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

                # Feuille Classement Comités Régionaux
                ws_comites = writer.sheets["Classement Comités"]
                ws_comites.views.sheetView[0].showGridLines = True
                ws_comites.cell(row=1, column=1, value=f"COMPÉTITION : {nom_competition.upper()}").font = font_titre
                ws_comites.cell(row=2, column=1, value="🏛️ CLASSEMENT OFFICIEL DES COMITÉS RÉGIONAUX - FFLDA").font = Font(name="Arial", size=12, bold=True, color="666666")
                ws_comites.cell(row=3, column=1, value=f"Édité le {datetime.now().strftime('%d/%m/%Y à %H:%M')}").font = Font(name="Arial", size=9, italic=True, color="888888")
                
                for col_idx in range(1, len(df_comites.columns) + 1):
                    cell = ws_comites.cell(row=5, column=col_idx)
                    cell.fill, cell.font, cell.alignment = bleu_fflda, font_entete, Alignment(horizontal="center", vertical="center")
                    ws_comites.row_dimensions[5].height = 25
                
                for row_idx in range(6, ws_comites.max_row + 1):
                    ws_comites.row_dimensions[row_idx].height = 22
                    is_even = (row_idx % 2 == 0)
                    for col_idx in range(1, len(df_comites.columns) + 1):
                        cell = ws_comites.cell(row=row_idx, column=col_idx)
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

                ws_indiv.views.sheetView[0].showGridLines = True
                ws_indiv.cell(row=1, column=1, value=f"COMPÉTITION : {nom_competition.upper()}").font = font_titre
                ws_indiv.cell(row=2, column=1, value="🏆 CLASSEMENTS INDIVIDUELS OFFICIELS - FFLDA").font = Font(name="Arial", size=12, bold=True, color="666666")
                ws_indiv.cell(row=3, column=1, value=f"Édité le {datetime.now().strftime('%d/%m/%Y à %H:%M')}").font = Font(name="Arial", size=9, italic=True, color="888888")
                
                row_cursor = 5
                headers_indiv = ["Clt", "NOM Prénom", "CLUB", "POIDS (kg)", "POINTS"]
                
                for poule in df_bilan['Poule'].unique():
                    groupe = df_bilan[df_bilan['Poule'] == poule]
                    ws_indiv.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=5)
                    cell_cat = ws_indiv.cell(row=row_cursor, column=1, value=f"  CATÉGORIE / POULE : {poule}")
                    cell_cat.fill, cell_cat.font, cell_cat.alignment = bleu_fflda, font_section, Alignment(horizontal="left", vertical="center")
                    ws_indiv.row_dimensions[row_cursor].height = 28
                    row_cursor += 1
                    
                    for col_idx, h in enumerate(headers_indiv, 1):
                        cell = ws_indiv.cell(row=row_cursor, column=col_idx, value=h)
                        cell.fill, cell.font, cell.alignment = rouge_fflda, font_entete, Alignment(horizontal="center", vertical="center")
                        ws_indiv.row_dimensions[row_cursor].height = 22
                    row_cursor += 1
                    
                    for _, row in groupe.iterrows():
                        current_row = row_cursor
                        ws_indiv.row_dimensions[current_row].height = 20
                        
                        c1 = ws_indiv.cell(row=current_row, column=1, value=row['Clt'])
                        c2 = ws_indiv.cell(row=current_row, column=2, value=row['Nom'])
                        c3 = ws_indiv.cell(row=current_row, column=3, value=row['Club'])
                        c4 = ws_indiv.cell(row=current_row, column=4, value=row['Poids'])
                        c5 = ws_indiv.cell(row=current_row, column=5, value=row['Points'])
                        
                        c1.font = font_data_bold
                        c2.font = font_data
                        c3.font = font_data
                        c4.font = font_data
                        c5.font = font_data_bold
                        
                        for c in [c1, c2, c3, c4, c5]:
                            c.border = b_fin
                            c.alignment = Alignment(horizontal="center", vertical="center")
                        c2.alignment = Alignment(horizontal="left", vertical="center")
                        
                        if row['Clt'] == 1: c1.fill = or_fill
                        elif row['Clt'] == 2: c1.fill = argent_fill
                        elif row['Clt'] == 3: c1.fill = bronze_fill
                        
                        row_cursor += 1
                    row_cursor += 2
                
                ws_indiv.column_dimensions['A'].width = 10
                ws_indiv.column_dimensions['B'].width = 30
                ws_indiv.column_dimensions['C'].width = 25
                ws_indiv.column_dimensions['D'].width = 15
                ws_indiv.column_dimensions['E'].width = 12

                # Configuration Impression Paysage A4 Excel
                for ws in writer.book.worksheets:
                    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
                    ws.page_setup.paperSize = ws.PAPERSIZE_A4
                    ws.sheet_properties.pageSetUpPr.fitToPage = True
                    ws.page_setup.fitToWidth = 1
                    ws.page_setup.fitToHeight = 0

            st.markdown("---")
            st.download_button(
                label="📥 Télécharger le Bilan Officiel FFLDA (Excel)",
                data=output_bilan.getvalue(),
                file_name="Bilan_Officiel_FFLDA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erreur lors de l'analyse du fichier : {e}")

else:
    # --- MODE 1 : GÉNÉRATION DE TOURNOI ---
    col_up1, col_up2 = st.columns([1, 1])
    with col_up1:
        fichier_upload = st.file_uploader("📂 Importez votre liste d'inscrits (.csv ou .xlsx)", type=["xlsx", "csv"])
    with col_up2:
        fichier_arbitres_upload = st.file_uploader("🛡️ (Optionnel) Importez la liste des arbitres (.xlsx ou .csv)", type=["xlsx", "csv"], key="upload_arbitres")

    if fichier_upload is None:
        st.info("👈 Veuillez importer un fichier de participants (format Exalto .csv ou .xlsx) pour lancer l'optimisation des poules et plannings.")
    else:
        try:
            liste_arbitres = charger_liste_arbitres(fichier_arbitres_upload)
            tapis_arbitres = {t: [] for t in range(nb_tapis)}
            if liste_arbitres:
                for idx_arb, arb in enumerate(liste_arbitres):
                    tapis_arbitres[idx_arb % nb_tapis].append(arb)
            if fichier_upload.name.endswith('.csv'):
                df_raw = pd.read_csv(fichier_upload, sep=';', encoding='utf-8')
                if len(df_raw.columns) == 1:
                    fichier_upload.seek(0)
                    df_raw = pd.read_csv(fichier_upload, sep=',', encoding='utf-8')
            else:
                df_temp = pd.read_excel(fichier_upload, nrows=5)
                header_row = 0
                for i, row in df_temp.iterrows():
                    if 'N° Licence' in str(row.values) or 'Nom' in str(row.values) or "Catégorie d'âge" in str(row.values):
                        header_row = i + 1
                        break
                fichier_upload.seek(0)
                df_raw = pd.read_excel(fichier_upload, header=header_row)

            if "Catégorie d'âge" in df_raw.columns: df_raw = df_raw.rename(columns={"Catégorie d'âge": "Age"})
            if "Sigle du Club" in df_raw.columns: df_raw = df_raw.rename(columns={"Sigle du Club": "Club"})
            comite_col_found = None
            for col_name in df_raw.columns:
                col_str = str(col_name).strip()
                col_lower = col_str.lower()
                if any(k in col_lower for k in ["comité", "comite", "ligue", "région", "region", "c.r."]):
                    comite_col_found = col_name
                    break
            if comite_col_found:
                df_raw = df_raw.rename(columns={comite_col_found: "Comité"})
            if "Comité" not in df_raw.columns: df_raw["Comité"] = "Comité Non Renseigné"

            style_col_found = None
            for col_name in df_raw.columns:
                col_str = str(col_name).strip().lower()
                if any(k in col_str for k in ["style", "discipline"]):
                    style_col_found = col_name
                    break
            if style_col_found:
                df_raw = df_raw.rename(columns={style_col_found: "Style"})
            if "Style" not in df_raw.columns:
                df_raw["Style"] = ""

            maitrise_col_found = None
            for col_name in df_raw.columns:
                col_str = str(col_name).strip().lower()
                if "maîtrise" in col_str or "maitrise" in col_str:
                    maitrise_col_found = col_name
                    break
            if maitrise_col_found:
                df_raw = df_raw.rename(columns={maitrise_col_found: "Maîtrise"})
            if "Maîtrise" not in df_raw.columns:
                df_raw["Maîtrise"] = ""

            if "Prénom" in df_raw.columns and "Nom" in df_raw.columns:
                df_raw["Nom"] = df_raw["Nom"].astype(str) + " " + df_raw["Prénom"].astype(str)

            df_inscr_total = df_raw.copy()
            if "Age" in df_inscr_total.columns:
                df_inscr_total = df_inscr_total[df_inscr_total['Age'].isin(['U9', 'U11'])]
            
            total_inscrits_global = len(df_inscr_total)

            def attribuer_niveau(val):
                val_str = str(val).strip().lower()
                if val_str.startswith('c') or 'confirm' in val_str:
                    return 'Confirmé'
                elif val_str.startswith('d') or 'début' in val_str or 'debut' in val_str:
                    return 'Débutant'
                else:
                    return 'Débutant'
            
            if poules_par_niveau:
                df_inscr_total['Niveau'] = df_inscr_total['Maîtrise'].apply(attribuer_niveau)
            else:
                df_inscr_total['Niveau'] = ''

            # Nettoyage et conversion du poids
            df_inscr_total['Poids_Clean'] = df_inscr_total['Poids'].astype(str).str.replace(',', '.')
            df_inscr_total['Poids_Num'] = pd.to_numeric(df_inscr_total['Poids_Clean'], errors='coerce').fillna(0)
            df_inscr_total['Poids'] = df_inscr_total['Poids_Num'].apply(formater_poids)

            # --- DÉTECTION ET VALIDATION DU STYLE "JEUNE" ---
            erreurs_jeune = []
            for _, row_test in df_inscr_total.iterrows():
                val_style_raw = str(row_test.get('Style', '')).strip().lower()
                poids_val = row_test['Poids_Num']
                nom_lutteur = row_test.get('Nom', 'Lutteur Inconnu')
                club_lutteur = row_test.get('Club', '')
                
                if val_style_raw == 'jeune':
                    if poids_val > 0:
                        erreurs_jeune.append(f"• **{nom_lutteur}** ({club_lutteur}) : Poids renseigné (**{poids_val} kg**) avec le style '**jeune**'. Le style doit être corrigé en LL, LF ou LG.")

            if erreurs_jeune:
                st.error("❌ **Erreur d'importation dans le fichier :**\n\n" + "\n".join(erreurs_jeune))
                st.info("💡 *Remarque : Le style 'jeune' ne peut pas être associé à un poids validé. Veuillez corriger les styles dans votre fichier Excel (LL, LF ou LG) avant de réimporter.*")
                st.stop()

            # Normalisation des styles (LL, LF, LG) et exclusion des "jeune" sans poids
            def normaliser_style(row_p):
                st_str = str(row_p.get('Style', '')).strip().upper()
                sexe_str = str(row_p.get('Sexe', '')).strip().upper()
                if 'LG' in st_str or 'GRECO' in st_str or 'GRÉCO' in st_str:
                    return 'LG'
                elif 'LF' in st_str or 'FEM' in st_str or 'FÉM' in st_str or sexe_str == 'F':
                    return 'LF'
                elif 'LL' in st_str or 'LIBRE' in st_str or sexe_str in ['M', 'H', 'G']:
                    return 'LL'
                else:
                    return 'LL'

            df_inscr_total['Style_Norm'] = df_inscr_total.apply(normaliser_style, axis=1)

            # Exclure les "jeune" sans poids (ou poids=0) et ne conserver que les pesés avec style valide
            df_inscr = df_inscr_total[
                (df_inscr_total['Poids_Num'] > 0) & 
                (df_inscr_total['Style'].astype(str).str.strip().str.lower() != 'jeune')
            ].copy()

            total_participants_peses = len(df_inscr)
            total_non_peses = total_inscrits_global - total_participants_peses

            if df_inscr.empty:
                st.error("❌ Aucun lutteur U9 ou U11 avec un poids valide et un style de compétition n'a été trouvé dans le fichier.")
                st.stop()
            
            # Définition des groupes de styles selon le réglage de mixité
            def attribuer_style_groupe(style_norm):
                if mixte_active:
                    # En mode mixte : LL et LF sont regroupés ensemble, mais LG reste strict séparé !
                    if style_norm == 'LG':
                        return 'LG (Gréco)'
                    else:
                        return 'Mixte (LL/LF)'
                else:
                    # Sans mixité : LL, LF et LG sont tous séparés
                    if style_norm == 'LG': return 'LG (Gréco)'
                    elif style_norm == 'LF': return 'LF (Féminine)'
                    else: return 'LL (Libre)'

            df_inscr['Style_Groupe'] = df_inscr['Style_Norm'].apply(attribuer_style_groupe)
            
            poules_u9, poules_u11 = [], []
            multiplicateur_poids = 1 + (tolerance_poids / 100.0)
            
            for age in ['U9', 'U11']:
                df_age = df_inscr[df_inscr['Age'] == age].sort_values('Poids_Num')
                max_size = 4 if age == 'U9' else 5
                index_poule = 1
                
                for (style_grp, niveau), groupe in df_age.groupby(['Style_Groupe', 'Niveau']):
                    participants = groupe.to_dict('records')
                    poule_courante = []
                    suffixe_niveau = f" | {niveau}" if niveau != "" else ""
                    poules_groupe = []
                    
                    for p in participants:
                        if not poule_courante:
                            poule_courante.append(p)
                        else:
                            poids_min = poule_courante[0]['Poids_Num']
                            if p['Poids_Num'] <= (poids_min * multiplicateur_poids) and len(poule_courante) < max_size:
                                poule_courante.append(p)
                            else:
                                nom_groupe = f"{age} | {style_grp}{suffixe_niveau} ({formater_poids_court(poule_courante[0]['Poids_Num'])} - {formater_poids_court(poule_courante[-1]['Poids_Num'])})"
                                poule_obj = {'nom': nom_groupe, 'participants': list(poule_courante), 'rondes': generer_rondes_fflda(poule_courante)}
                                poules_groupe.append(poule_obj)
                                index_poule += 1
                                poule_courante = [p]
                    if poule_courante:
                        nom_groupe = f"{age} | {style_grp}{suffixe_niveau} ({formater_poids_court(poule_courante[0]['Poids_Num'])} - {formater_poids_court(poule_courante[-1]['Poids_Num'])})"
                        poule_obj = {'nom': nom_groupe, 'participants': list(poule_courante), 'rondes': generer_rondes_fflda(poule_courante)}
                        poules_groupe.append(poule_obj)
                        index_poule += 1
                    
                    if separer_clubs:
                        poules_groupe = optimiser_poules_clubs(poules_groupe, multiplicateur_poids)

                    # Eviter les poules de 1 en les fusionnant / rééquilibrant tout en respectant la tolérance de poids
                    poules_groupe = fusionner_poules_isolees(poules_groupe, multiplicateur_poids, max_size)
                    
                    if separer_clubs:
                        poules_groupe = optimiser_poules_clubs(poules_groupe, multiplicateur_poids)
                    
                    # Formater et ré-indexer proprement les noms des poules finales
                    for idx_p, p_obj in enumerate(poules_groupe, 1):
                        parts = p_obj['participants']
                        p_min = parts[0]['Poids_Num']
                        p_max = parts[-1]['Poids_Num']
                        p_obj['nom'] = f"{age} | {style_grp}{suffixe_niveau} | Gr. {idx_p} ({formater_poids_court(p_min)} - {formater_poids_court(p_max)})"
                        p_obj['rondes'] = generer_rondes_fflda(parts)
                    
                    if age == 'U9': poules_u9.extend(poules_groupe)
                    else: poules_u11.extend(poules_groupe)

            participants_par_poule = {p['nom']: p['participants'] for p in poules_u9 + poules_u11}
            rondes_par_categorie = {p['nom']: p['rondes'] for p in poules_u9 + poules_u11}

            tapis_poules_u9 = {i: [] for i in range(nb_tapis)}
            tapis_poules_u11 = {i: [] for i in range(nb_tapis)}
            
            for i, p in enumerate(poules_u9): tapis_poules_u9[i % nb_tapis].append(p)
            for i, p in enumerate(poules_u11): tapis_poules_u11[i % nb_tapis].append(p)

            dt_pesee_u9 = datetime.combine(datetime.today(), heure_pesee_u9)
            dt_debut_u9 = dt_pesee_u9 + timedelta(minutes=duree_pesee)
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
                
                sans_conflit = []
                for arb in cand_list:
                    arb_club_clean = str(arb.get('Club', '')).strip().lower()
                    if arb_club_clean in ignored_clubs:
                        sans_conflit.append(arb)
                    elif arb_club_clean != c1_clean and arb_club_clean != c2_clean:
                        sans_conflit.append(arb)
                        
                pool_choix = sans_conflit if sans_conflit else cand_list
                arb_choisi = min(pool_choix, key=lambda a: ref_usage_count.get(a['Nom_Complet'], 0))
                
                ref_usage_count[arb_choisi['Nom_Complet']] = ref_usage_count.get(arb_choisi['Nom_Complet'], 0) + 1
                return arb_choisi['Nom_Complet']

            def ordonnancer_phase(poules_phase, heure_debut_phase, duree_combat, meme_tapis=True):
                global total_matchs_calcules
                if not poules_phase:
                    return [heure_debut_phase for _ in range(nb_tapis)], {t: [] for t in range(nb_tapis)}
                
                planning = {t: [] for t in range(nb_tapis)}
                tapis_heure = [heure_debut_phase for _ in range(nb_tapis)]
                last_match_time = {}

                if meme_tapis:
                    # Chaque poule (et ses lutteurs) reste affectée à un tapis fixe
                    tapis_poules = {t: [] for t in range(nb_tapis)}
                    for idx, p in enumerate(poules_phase):
                        tapis_poules[idx % nb_tapis].append(p)

                    for t in range(nb_tapis):
                        poules_t = tapis_poules[t]
                        if not poules_t:
                            continue
                        
                        matchs_tapis = []
                        max_rondes = max((len(p['rondes']) for p in poules_t), default=0)
                        for r in range(max_rondes):
                            for p in poules_t:
                                if r < len(p['rondes']):
                                    for m in p['rondes'][r]:
                                        matchs_tapis.append({
                                            'poule': p['nom'],
                                            'p1': m[0]['Nom'],
                                            'p1_club': m[0].get('Club', ''),
                                            'p2': m[1]['Nom'],
                                            'p2_club': m[1].get('Club', ''),
                                            'tour': r + 1
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
                                "Combattant 2": m['p2'],
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
                                for m in p['rondes'][r]:
                                    matchs_a_jouer.append({
                                        'poule': p['nom'],
                                        'p1': m[0]['Nom'],
                                        'p1_club': m[0].get('Club', ''),
                                        'p2': m[1]['Nom'],
                                        'p2_club': m[1].get('Club', ''),
                                        'tour': r + 1
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
                            "Combattant 2": m['p2'],
                            "Arbitre": arb_nom
                        })
                        
                        total_matchs_calcules += 1
                        fin_match = heure_debut_match + timedelta(minutes=duree_combat)
                        delai_repos = timedelta(minutes=(repos_matchs * duree_combat))
                        
                        last_match_time[m['p1']] = fin_match + delai_repos
                        last_match_time[m['p2']] = fin_match + delai_repos
                        
                        tapis_heure[t_min_idx] = fin_match

                return tapis_heure, planning

            # PHASE 1 : Tous les U9
            tapis_heure_u9, planning_u9 = ordonnancer_phase(poules_u9, dt_debut_u9, duree_u9, meme_tapis_poule)
            fin_u9_globale = max(tapis_heure_u9) if poules_u9 else dt_debut_u9

            # PAUSE / PESÉE U11
            dt_pesee_u11 = None
            if "2" in type_pesee:
                dt_pesee_u11 = fin_u9_globale + timedelta(minutes=duree_pause)
                dt_debut_u11_theorique = dt_pesee_u11 + timedelta(minutes=duree_pesee)
            else:
                dt_debut_u11_theorique = fin_u9_globale

            # Égalisation du nombre de lignes en U9 pour aligner horizontalement la PAUSE
            max_u9_lignes = max(len(planning_u9[t]) for t in range(nb_tapis)) if planning_u9 else 0
            for t in range(nb_tapis):
                while len(planning_u9[t]) < max_u9_lignes:
                    planning_u9[t].append({"Type": "VIDE"})

            if activer_pause and duree_pause > 0:
                for t in range(nb_tapis):
                    planning_u9[t].append({"Type": "PAUSE", "Heure": fin_u9_globale.strftime("%H:%M")})
                    tapis_heure_u9[t] = fin_u9_globale + timedelta(minutes=duree_pause)
            else:
                for t in range(nb_tapis):
                    tapis_heure_u9[t] = fin_u9_globale

            debut_u11_reel = max(tapis_heure_u9)
            if dt_debut_u11_theorique and debut_u11_reel < dt_debut_u11_theorique:
                debut_u11_reel = dt_debut_u11_theorique

            for t in range(nb_tapis):
                if poules_u11 and tapis_heure_u9[t] < debut_u11_reel:
                    attente = int((debut_u11_reel - tapis_heure_u9[t]).total_seconds() // 60)
                    if attente > 0:
                        planning_u9[t].append({
                            "Type": "ATTENTE", 
                            "Heure": tapis_heure_u9[t].strftime("%H:%M"), 
                            "Texte": f"Pesée + échauffement U11"
                        })

            # PHASE 2 : Tous les U11
            if poules_u11:
                tapis_heure_u11, planning_u11 = ordonnancer_phase(poules_u11, debut_u11_reel, duree_u11, meme_tapis_poule)
                planning_tapis = {t: planning_u9[t] + planning_u11[t] for t in range(nb_tapis)}
                fin_estimee = max(tapis_heure_u11)
            else:
                planning_tapis = planning_u9
                fin_estimee = fin_u9_globale

            texte_pesee_u9 = "1ère pesée" if "1" in type_pesee else "Pesée U9"
            valeur_pause = f"{duree_pause} min" if (activer_pause and duree_pause > 0) else "0 min"

            str_comp_u9 = f"{dt_debut_u9.strftime('%H:%M')} - {fin_u9_globale.strftime('%H:%M')}"
            str_comp_u11 = f"{debut_u11_reel.strftime('%H:%M')} - {fin_estimee.strftime('%H:%M')}"

            st.success("✨ Fichier analysé avec succès ! Tournoi généré.")
            
            # --- PRÉPARATION DES DONNÉES DU RÉSUMÉ ---
            lignes_accueil = [
                {"Étape de la journée": texte_pesee_u9, "Horaire / Valeur": dt_pesee_u9.strftime('%H:%M')},
                {"Étape de la journée": "Compétition U9", "Horaire / Valeur": str_comp_u9},
                {"Étape de la journée": "Pause de la compétition", "Horaire / Valeur": valeur_pause}
            ]
            if "2" in type_pesee and dt_pesee_u11:
                lignes_accueil.append({"Étape de la journée": "2ème pesée", "Horaire / Valeur": dt_pesee_u11.strftime('%H:%M')})

            lignes_accueil.extend([
                {"Étape de la journée": "Compétition U11", "Horaire / Valeur": str_comp_u11},
                {"Étape de la journée": "Fin de la compétition estimée", "Horaire / Valeur": fin_estimee.strftime('%H:%M')}
            ])

            # --- GÉNÉRATION DES DOCUMENTS HTML PAYSAGE DU TOURNOI ---
            sections_tournoi_complet = [
                ("📊 Résumé Prévisionnel de la Journée", pd.DataFrame(lignes_accueil))
            ]
            
            if liste_arbitres:
                lignes_arb_print = []
                for t in range(nb_tapis):
                    noms_arb = ", ".join([a['Nom_Complet'] for a in tapis_arbitres[t]]) if tapis_arbitres[t] else "Aucun arbitre affecté"
                    lignes_arb_print.append({"Tapis": f"Tapis {t + 1}", "Effectif": f"{len(tapis_arbitres[t])} arbitres", "Équipe d'Arbitrage Désignée": noms_arb})
                sections_tournoi_complet.append(("🛡️ Désignation des Équipes d'Arbitrage par Tapis", pd.DataFrame(lignes_arb_print)))

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
                            arb_str = f" (🛡️ {m['Arbitre']})" if m.get('Arbitre') and m['Arbitre'] != "Non attribué" else ""
                            ligne[col] = f"[{m['Heure']}] ({m['Duree']}m) [{m['Cat']}] - {m['Combattant 1']} vs {m['Combattant 2']}{arb_str}"
                    else: ligne[col] = ""
                grille_ui.append(ligne)
            
            sections_tournoi_complet.append(("📅 Grille de Passage - Tapis", pd.DataFrame(grille_ui)))
            
            for nom_poule, liste_p in participants_par_poule.items():
                df_poule_vue = pd.DataFrame(liste_p)[['Nom', 'Club', 'Poids']]
                sections_tournoi_complet.append((f"🤼 Feuille de Poule : {nom_poule}", df_poule_vue))
                
            html_tournoi_complet = generer_document_html_imprimable("Feuilles Officieuses du Tournoi & Poules FFLDA", nom_competition, sections_tournoi_complet)

            # --- ONGLETS INTERACTIFS DE L'APPLICATION ---
            noms_onglets = ["📊 Résumé & Stats", "📅 Grille de Passage", "🛡️ Équipes d'Arbitrage"] + [f"Poule : {p[:15]}" for p in participants_par_poule.keys()]
            onglets_ui = st.tabs(noms_onglets)
            
            with onglets_ui[0]:
                st.subheader("📊 Résumé prévisionnel de la journée")
                st.table(pd.DataFrame(lignes_accueil))
                
                col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                col_m1.metric("Participants (pesés)", total_participants_peses)
                col_m2.metric("Absents / Non pesés", total_non_peses)
                col_m3.metric("Matchs générés", total_matchs_calcules)
                col_m4.metric("Arbitres engagés", len(liste_arbitres))

                if liste_arbitres:
                    st.markdown("#### 🛡️ Désignation des Équipes d'Arbitrage par Tapis")
                    lignes_arb_sum = []
                    for t in range(nb_tapis):
                        noms_arb = ", ".join([a['Nom_Complet'] for a in tapis_arbitres[t]]) if tapis_arbitres[t] else "Aucun arbitre affecté"
                        lignes_arb_sum.append({"Tapis": f"Tapis {t + 1}", "Effectif": f"{len(tapis_arbitres[t])} arbitres", "Équipe d'Arbitrage Désignée": noms_arb})
                    st.table(pd.DataFrame(lignes_arb_sum))

            with onglets_ui[1]:
                st.subheader("📅 Grille de Passage - Tapis")
                if liste_arbitres:
                    st.markdown("**🛡️ Équipes d'arbitrage affectées aux tapis :**")
                    cols_arb_disp = st.columns(nb_tapis)
                    for t in range(nb_tapis):
                        with cols_arb_disp[t]:
                            arb_list_txt = "\n".join([f"• **{a['Nom_Complet']}** ({a['Club']})" for a in tapis_arbitres[t]]) if tapis_arbitres[t] else "• Aucun"
                            st.info(f"**Tapis {t+1}** ({len(tapis_arbitres[t])} arbitres) :\n\n{arb_list_txt}")
                st.table(pd.DataFrame(grille_ui))
                bouton_imprimer(html_tournoi_complet, filename="Grille_Tapis_Impression.html", label="🖨️ Imprimer / Télécharger la Grille (HTML Paysage A4)", key="btn_t1")

            with onglets_ui[2]:
                st.subheader("🛡️ Désignation et Affectation des Arbitres par Tapis")
                if liste_arbitres:
                    for t in range(nb_tapis):
                        st.markdown(f"#### 🥋 Tapis {t + 1} ({len(tapis_arbitres[t])} arbitres)")
                        if tapis_arbitres[t]:
                            df_arb_tapis = pd.DataFrame(tapis_arbitres[t])[['Nom', 'Prenom', 'Licence', 'Club', 'Comite']]
                            df_arb_tapis.columns = ['Nom', 'Prénom', 'N° Licence', 'Club', 'Comité Régional']
                            st.table(df_arb_tapis)
                        else:
                            st.info("Aucun arbitre affecté à ce tapis.")
                else:
                    st.info("Aucun fichier d'arbitres n'a été chargé.")

            for idx, (nom_poule, liste_p) in enumerate(participants_par_poule.items(), start=3):
                with onglets_ui[idx]:
                    st.subheader(f"Feuille de Poule : {nom_poule}")
                    df_poule_vue = pd.DataFrame(liste_p)[['Nom', 'Club', 'Poids']]
                    st.table(df_poule_vue)

            st.markdown("---")
            
            # --- EXPORT EXCEL ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                
                resume_data = [
                    {"Étape de la journée": texte_pesee_u9, "Horaire / Valeur": dt_pesee_u9.strftime('%H:%M')},
                    {"Étape de la journée": "Compétition U9", "Horaire / Valeur": str_comp_u9},
                    {"Étape de la journée": "Pause de la compétition", "Horaire / Valeur": valeur_pause}
                ]
                if "2" in type_pesee and dt_pesee_u11:
                    resume_data.append({"Étape de la journée": "2ème pesée", "Horaire / Valeur": dt_pesee_u11.strftime('%H:%M')})
                
                resume_data.extend([
                    {"Étape de la journée": "Compétition U11", "Horaire / Valeur": str_comp_u11},
                    {"Étape de la journée": "Fin de la compétition estimée", "Horaire / Valeur": fin_estimee.strftime('%H:%M')}
                ])
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
                                arb_str = f"\n🛡️ Arbitre : {m['Arbitre']}" if m.get('Arbitre') and m['Arbitre'] != "Non attribué" else ""
                                ligne[col] = f"🕘 {m['Heure']} ({m['Duree']} min)\n[{m['Cat']}]\n{m['Combattant 1']} VS {m['Combattant 2']}{arb_str}"
                        else: ligne[col] = ""
                    grille.append(ligne)
                pd.DataFrame(grille).to_excel(writer, sheet_name="Grille de Passage", index=False, startrow=1)
                
                b_style = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                bleu = PatternFill("solid", fgColor="0055A4")
                rouge = PatternFill("solid", fgColor="EF4135")
                bleu_clair = PatternFill("solid", fgColor="DDEBF7") 
                
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
                titre_cell = ws_grille.cell(row=1, column=1, value=f"🏆 {nom_competition.upper()} - PLANNING OFFICIEL 🏆")
                titre_cell.font = Font(name="Arial", size=22, bold=True, color="FFFFFF")
                titre_cell.fill = bleu
                titre_cell.alignment = Alignment(horizontal="center", vertical="center")
                
                try:
                    from openpyxl.drawing.image import Image as OpenpyxlImage
                    url_logo = "https://upload.wikimedia.org/wikipedia/fr/thumb/5/58/Logo_F%C3%A9d%C3%A9ration_Fran%C3%A7aise_de_Lutte.svg/200px-Logo_F%C3%A9d%C3%A9ration_Fran%C3%A7aise_de_Lutte.svg.png"
                    req = urllib.request.Request(url_logo, headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req) as response: img_data = io.BytesIO(response.read())
                    img = OpenpyxlImage(img_data)
                    img.height, img.width = 70, 70
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
                    is_even = (row[0].row % 2 == 0)
                    for cell in row:
                        cell.border, cell.alignment = b_style, Alignment(wrap_text=True, horizontal="center", vertical="center")
                        if cell.value:
                            if "PAUSE" in str(cell.value): cell.fill, cell.font = rouge, Font(bold=True, color="FFFFFF", size=12)
                            elif any(k in str(cell.value) for k in ["Attente", "Pesée", "échauffement", "Repos"]): cell.fill, cell.font = PatternFill("solid", fgColor="EFEFEF"), Font(italic=True, color="666666", size=11)
                            else:
                                cell.fill = bleu_clair if is_even else PatternFill(fill_type=None)
                                cell.font = Font(size=12)

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
                        ws_arb_sheet.cell(row=2, column=1, value="🛡️ CORPS D'ARBITRAGE ET AFFECTATION AUX TAPIS - FFLDA").font = Font(name="Arial", size=12, bold=True, color="666666")
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

                for nom_poule, liste_p in participants_par_poule.items():
                    nom_onglet_court = abreger_nom_onglet(nom_poule)
                    ws_poule = writer.book.create_sheet(nom_onglet_court)
                    
                    ws_poule.cell(row=1, column=1, value=f"POULE : {nom_poule}").font = Font(bold=True, size=16, color="0055A4")
                    ws_poule.cell(row=2, column=1, value="*POINT DE CLASSEMENT : 2 pt = victoire - 1 pt = match nul - 0 pt = défaite").font = Font(italic=True, size=9)
                    
                    row_cursor = 4
                    headers = ["CLT", "N°", "NOM Prénom", "CLUB", "COMITÉ"]
                    nb_tours = len(rondes_par_categorie[nom_poule])
                    for t in range(1, nb_tours + 1):
                        headers.append(f"Tour {t}")
                    headers.extend(["Total Pts", "Total Vict", "Poids"])
                    
                    for col_idx, h in enumerate(headers, 1):
                        c = ws_poule.cell(row=row_cursor, column=col_idx, value=h)
                        c.font, c.alignment, c.border = Font(bold=True, color="FFFFFF"), Alignment(horizontal="center", vertical="center"), b_style
                        c.fill = entete_noir
                    
                    max_len_nom = max([len(str(p.get('Nom', ''))) for p in liste_p] + [12])
                    max_len_club = max([len(str(p.get('Club', ''))) for p in liste_p] + [10])
                    max_len_comite = max([len(str(p.get('Comité', ''))) for p in liste_p] + [12])
                    
                    largeur_nom_col = max(max_len_nom + 4, 25)
                    largeur_club_col = max(max_len_club + 4, 18)
                    largeur_comite_col = max(max_len_comite + 4, 20)

                    ws_poule.column_dimensions['A'].width = 6
                    ws_poule.column_dimensions['B'].width = 6
                    ws_poule.column_dimensions['C'].width = largeur_nom_col  
                    ws_poule.column_dimensions['D'].width = largeur_club_col 
                    ws_poule.column_dimensions['E'].width = largeur_comite_col
                    ws_poule.column_dimensions['F'].width = 6                 
                    ws_poule.column_dimensions['G'].width = largeur_nom_col  
                    ws_poule.column_dimensions['H'].width = largeur_club_col 
                    ws_poule.column_dimensions['I'].width = 10                
                    
                    lignes_lutteurs = {}
                    row_cursor += 1
                    
                    ligne_debut_poule = row_cursor
                    for i, p in enumerate(liste_p, 1):
                        lignes_lutteurs[p['Nom']] = row_cursor
                        
                        col_pts_lettre = openpyxl.utils.get_column_letter(6 + nb_tours)
                        plage_totaux = f"{col_pts_lettre}{ligne_debut_poule}:{col_pts_lettre}{ligne_debut_poule + len(liste_p) - 1}"
                        
                        cell_clt = ws_poule.cell(row=row_cursor, column=1, value=f"=RANK({col_pts_lettre}{row_cursor}, {plage_totaux})")
                        cell_clt.border = b_style
                        cell_clt.alignment = Alignment(horizontal="center", vertical="center")
                        cell_clt.font = Font(bold=True, color="0055A4")

                        ws_poule.cell(row=row_cursor, column=2, value=i).border = b_style 
                        ws_poule.cell(row=row_cursor, column=2).alignment = Alignment(horizontal="center")
                        ws_poule.cell(row=row_cursor, column=3, value=p['Nom']).border = b_style
                        ws_poule.cell(row=row_cursor, column=4, value=p.get('Club', '')).border = b_style
                        ws_poule.cell(row=row_cursor, column=5, value=p.get('Comité', '')).border = b_style
                        
                        col_offset = 6
                        for t in range(nb_tours):
                            cell_tour = ws_poule.cell(row=row_cursor, column=col_offset+t)
                            cell_tour.border = b_style 
                            cell_tour.alignment = Alignment(horizontal="center", vertical="center")
                        
                        col_lettre_debut = openpyxl.utils.get_column_letter(col_offset)
                        col_lettre_fin = openpyxl.utils.get_column_letter(col_offset + nb_tours - 1)
                        
                        cell_total_pts = ws_poule.cell(row=row_cursor, column=col_offset+nb_tours, value=f"=SUM({col_lettre_debut}{row_cursor}:{col_lettre_fin}{row_cursor})")
                        cell_total_pts.border = b_style
                        cell_total_pts.alignment = Alignment(horizontal="center", vertical="center")
                        cell_total_pts.font = Font(bold=True)
                        
                        cell_total_vict = ws_poule.cell(row=row_cursor, column=col_offset+nb_tours+1, value="")
                        cell_total_vict.border = b_style 
                        cell_total_vict.alignment = Alignment(horizontal="center", vertical="center")
                        
                        cell_poids = ws_poule.cell(row=row_cursor, column=col_offset+nb_tours+2, value=p.get('Poids', ''))
                        cell_poids.border = b_style 
                        cell_poids.alignment = Alignment(horizontal="center", vertical="center")
                        
                        row_cursor += 1
                    
                    row_cursor += 2
                    
                    rondes = rondes_par_categorie[nom_poule]
                    col_offset_tours = 6 
                    
                    for tour_idx, ronde in enumerate(rondes, 1):
                        ws_poule.cell(row=row_cursor, column=2, value=f"TOUR {tour_idx}").font = Font(bold=True, size=14)
                        row_cursor += 1
                        
                        for match in ronde:
                            p1, p2 = match[0], match[1]
                            idx1 = next((i+1 for i, x in enumerate(liste_p) if x['Nom'] == p1['Nom']), "")
                            idx2 = next((i+1 for i, x in enumerate(liste_p) if x['Nom'] == p2['Nom']), "")
                            
                            c_rouge = ws_poule.cell(row=row_cursor, column=3, value="LUTTEUR ROUGE")
                            c_rouge.fill, c_rouge.font, c_rouge.alignment, c_rouge.border = rouge_lutte, Font(bold=True, color="FFFFFF"), Alignment(horizontal="center"), b_style
                            ws_poule.merge_cells(start_row=row_cursor, start_column=3, end_row=row_cursor, end_column=4)
                            
                            c_ptr = ws_poule.cell(row=row_cursor, column=5, value="Pt Clt")
                            c_ptr.font, c_ptr.alignment, c_ptr.border = Font(bold=True), Alignment(horizontal="center"), b_style
                            
                            c_bleu = ws_poule.cell(row=row_cursor, column=7, value="LUTTEUR BLEU")
                            c_bleu.fill, c_bleu.font, c_bleu.alignment, c_bleu.border = bleu_lutte, Font(bold=True, color="FFFFFF"), Alignment(horizontal="center"), b_style
                            ws_poule.merge_cells(start_row=row_cursor, start_column=7, end_row=row_cursor, end_column=8)
                            
                            c_ptb = ws_poule.cell(row=row_cursor, column=9, value="Pt Clt")
                            c_ptb.font, c_ptb.alignment, c_ptb.border = Font(bold=True), Alignment(horizontal="center"), b_style
                            
                            row_cursor += 1
                            
                            ws_poule.cell(row=row_cursor, column=2, value=idx1).alignment = Alignment(horizontal="center")
                            ws_poule.cell(row=row_cursor, column=2).font = Font(bold=True, color="E53935", size=14)
                            ws_poule.cell(row=row_cursor, column=3, value=p1['Nom']).border = b_style
                            ws_poule.cell(row=row_cursor, column=4, value=p1.get('Club', '')).border = b_style
                            
                            box_ptr = ws_poule.cell(row=row_cursor, column=5)
                            box_ptr.border, box_ptr.fill = b_style, gris_clair
                            box_ptr.alignment = Alignment(horizontal="center", vertical="center")
                            
                            ws_poule.cell(row=row_cursor, column=6, value=idx2).alignment = Alignment(horizontal="center")
                            ws_poule.cell(row=row_cursor, column=6).font = Font(bold=True, color="1E88E5", size=14)
                            ws_poule.cell(row=row_cursor, column=7, value=p2['Nom']).border = b_style
                            ws_poule.cell(row=row_cursor, column=8, value=p2.get('Club', '')).border = b_style
                            
                            box_ptb = ws_poule.cell(row=row_cursor, column=9)
                            box_ptb.border, box_ptb.fill = b_style, gris_clair
                            box_ptb.alignment = Alignment(horizontal="center", vertical="center")
                            
                            if p1['Nom'] in lignes_lutteurs:
                                lig_haut_p1 = lignes_lutteurs[p1['Nom']]
                                cell_haut_p1 = ws_poule.cell(row=lig_haut_p1, column=col_offset_tours + (tour_idx - 1))
                                cell_haut_p1.value = f"={box_ptr.coordinate}"
                                cell_haut_p1.alignment = Alignment(horizontal="center", vertical="center")
                            
                            if p2['Nom'] in lignes_lutteurs:
                                lig_haut_p2 = lignes_lutteurs[p2['Nom']]
                                cell_haut_p2 = ws_poule.cell(row=lig_haut_p2, column=col_offset_tours + (tour_idx - 1))
                                cell_haut_p2.value = f"={box_ptb.coordinate}"
                                cell_haut_p2.alignment = Alignment(horizontal="center", vertical="center")

                            row_cursor += 1
                            
                            ws_poule.cell(row=row_cursor, column=3, value="Points Techniques (Actions)").font = Font(size=9, italic=True)
                            ws_poule.merge_cells(start_row=row_cursor, start_column=3, end_row=row_cursor, end_column=4)
                            ws_poule.cell(row=row_cursor, column=5, value="Total Score").font = Font(size=9, italic=True)
                            
                            ws_poule.cell(row=row_cursor, column=7, value="Points Techniques (Actions)").font = Font(size=9, italic=True)
                            ws_poule.merge_cells(start_row=row_cursor, start_column=7, end_row=row_cursor, end_column=8)
                            ws_poule.cell(row=row_cursor, column=9, value="Total Score").font = Font(size=9, italic=True)
                            
                            row_cursor += 1
                            
                            ws_poule.row_dimensions[row_cursor].height = 25
                            ws_poule.cell(row=row_cursor, column=3).border = b_style
                            ws_poule.cell(row=row_cursor, column=4).border = b_style
                            ws_poule.merge_cells(start_row=row_cursor, start_column=3, end_row=row_cursor, end_column=4)
                            ws_poule.cell(row=row_cursor, column=5).border = b_style
                            
                            ws_poule.cell(row=row_cursor, column=7).border = b_style
                            ws_poule.cell(row=row_cursor, column=8).border = b_style
                            ws_poule.merge_cells(start_row=row_cursor, start_column=7, end_row=row_cursor, end_column=8)
                            ws_poule.cell(row=row_cursor, column=9).border = b_style
                            
                            row_cursor += 2 
                        
                        row_cursor += 1

            st.download_button(label="📥 Télécharger le Planning & Feuilles de Poules (Excel)", data=output.getvalue(), file_name="Tournoi_U9_U11.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Erreur lors de l'analyse du fichier : {e}")

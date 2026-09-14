import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
import io
import urllib.request

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Générateur Officiel FFLDA", page_icon="🤼", layout="wide")

# --- PERSONNALISATION VISUELLE FFLDA (CSS) ---
st.markdown("""

""", unsafe_allow_html=True)

# --- MENU LATÉRAL (PARAMÈTRES INTERACTIFS) ---
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/fr/thumb/5/58/Logo_F%C3%A9d%C3%A9ration_Fran%C3%A7aise_de_Lutte.svg/1200px-Logo_F%C3%A9d%C3%A9ration_Fran%C3%A7aise_de_Lutte.svg.png", use_container_width=True)
    st.header("⚙️ Paramètres du Tournoi")
    st.caption("Tournoi exclusif U9 / U11")
    
    st.subheader("1. Logistique & Pesées")
    nb_tapis = st.number_input("Nombre de tapis", min_value=1, max_value=10, value=3)
    
    type_pesee = st.radio("Format des pesées", ["1 Pesée (Générale)", "2 Pesées (U9 puis U11)"], index=1)
    duree_pesee = st.selectbox("Durée allouée à la pesée (min)", [30, 45, 60, 90], index=1) # 45 min par défaut
    
    heure_pesee_u9 = st.time_input("Heure de pesée U9", value=time(8, 30))
    if "2" in type_pesee:
        heure_pesee_u11 = st.time_input("Heure de pesée U11", value=time(10, 30))
    else:
        heure_pesee_u11 = None
        
    st.subheader("2. Pause Déjeuner")
    activer_pause = st.checkbox("Activer la pause déjeuner", value=True)
    if activer_pause:
        pause_debut = st.time_input("Début de la pause", value=time(12, 0))
        pause_fin = st.time_input("Fin de la pause", value=time(13, 0))
    else:
        pause_debut, pause_fin = None, None
    
    st.subheader("3. Règles Sportives")
    mixte_active = st.checkbox("Catégories Mixtes (U9/U11 filles et garçons ensemble)", value=True)
    repos_matchs = st.number_input("Matchs de repos minimum", min_value=1, max_value=10, value=3)
    
    st.subheader("4. Temps des Combats (Match + Rotation)")
    duree_u9 = st.number_input("Temps total U9 (min)", value=3) # 2min + 1min rotation
    duree_u11 = st.number_input("Temps total U11 (min)", value=4) # 3min + 1min rotation

# --- CORPS PRINCIPAL ---
st.title("Générateur de Planning FFLDA 🚀")
st.markdown("**Outil officiel d'optimisation (Compatible imports Exalto)**")
st.markdown("---")

def generer_rondes(participants_in):
    participants = list(participants_in)
    if len(participants) % 2 != 0:
        participants.append({"Nom": "BYE", "Club": "-"})
    n = len(participants)
    rondes = []
    for i in range(n - 1):
        matchs_ronde = []
        for j in range(n // 2):
            p1 = participants[j]
            p2 = participants[n - 1 - j]
            if p1["Nom"] != "BYE" and p2["Nom"] != "BYE":
                matchs_ronde.append((p1, p2))
        rondes.append(matchs_ronde)
        participants.insert(1, participants.pop())
    return rondes

fichier_upload = st.file_uploader("📂 Importez votre liste d'inscrits (.csv ou .xlsx)", type=["xlsx", "csv"])

if fichier_upload is not None:
    try:
        # --- LECTURE DU FICHIER ---
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
        if "Prénom" in df_raw.columns and "Nom" in df_raw.columns:
            df_raw["Nom"] = df_raw["Nom"].astype(str) + " " + df_raw["Prénom"].astype(str)

        df_inscr = df_raw.copy()
        
        # --- FILTRAGE STRICT U9 ET U11 ---
        df_inscr = df_inscr[df_inscr['Age'].isin(['U9', 'U11'])]
        if df_inscr.empty:
            st.error("❌ Aucun lutteur U9 ou U11 n'a été trouvé dans le fichier.")
            st.stop()
        
        if "Maîtrise" not in df_inscr.columns: df_inscr["Maîtrise"] = ""
        def attribuer_niveau(val):
            val_str = str(val).strip().lower()
            if val_str == 'd': return 'Débutant'
            elif val_str == 'c': return 'Confirmé'
            return ''
        df_inscr['Niveau'] = df_inscr['Maîtrise'].apply(attribuer_niveau)

        df_inscr['Poids'] = df_inscr['Poids'].astype(str).str.replace(',', '.')
        df_inscr['Poids_Num'] = pd.to_numeric(df_inscr['Poids'], errors='coerce')
        df_inscr = df_inscr[df_inscr['Poids_Num'] > 0]
        
        if mixte_active:
            df_inscr['Sexe'] = 'Mixte'
        
        # --- CRÉATION DES POULES INTELLIGENTES ---
        poules_u9, poules_u11 = [], []
        
        for age in ['U9', 'U11']:
            df_age = df_inscr[df_inscr['Age'] == age].sort_values('Poids_Num')
            max_size = 4 if age == 'U9' else 5
            index_poule = 1
            
            for (sexe, niveau), groupe in df_age.groupby(['Sexe', 'Niveau']):
                participants = groupe.to_dict('records')
                poule_courante = []
                suffixe_niveau = f" | {niveau}" if niveau != "" else ""
                
                for p in participants:
                    if not poule_courante:
                        poule_courante.append(p)
                    else:
                        poids_min = poule_courante[0]['Poids_Num']
                        if p['Poids_Num'] <= (poids_min * 1.10) and len(poule_courante) < max_size:
                            poule_courante.append(p)
                        else:
                            nom_groupe = f"{age} | {sexe}{suffixe_niveau} | Gr. {index_poule} ({poule_courante[0]['Poids_Num']}kg - {poule_courante[-1]['Poids_Num']}kg)"
                            poule_obj = {'nom': nom_groupe, 'participants': list(poule_courante), 'rondes': generer_rondes(poule_courante)}
                            if age == 'U9': poules_u9.append(poule_obj)
                            else: poules_u11.append(poule_obj)
                            index_poule += 1
                            poule_courante = [p]
                if poule_courante:
                    nom_groupe = f"{age} | {sexe}{suffixe_niveau} | Gr. {index_poule} ({poule_courante[0]['Poids_Num']}kg - {poule_courante[-1]['Poids_Num']}kg)"
                    poule_obj = {'nom': nom_groupe, 'participants': list(poule_courante), 'rondes': generer_rondes(poule_courante)}
                    if age == 'U9': poules_u9.append(poule_obj)
                    else: poules_u11.append(poule_obj)
                    index_poule += 1

        participants_par_poule = {p['nom']: p['participants'] for p in poules_u9 + poules_u11}
        rondes_par_categorie = {p['nom']: p['rondes'] for p in poules_u9 + poules_u11}

        # --- AFFECTATION DES POULES AUX TAPIS ---
        tapis_poules_u9 = {i: [] for i in range(nb_tapis)}
        tapis_poules_u11 = {i: [] for i in range(nb_tapis)}
        
        for i, p in enumerate(poules_u9): tapis_poules_u9[i % nb_tapis].append(p)
        for i, p in enumerate(poules_u11): tapis_poules_u11[i % nb_tapis].append(p)

        # --- GESTION DES HORAIRES ET VAGUES ---
        dt_pesee_u9 = datetime.combine(datetime.today(), heure_pesee_u9)
        dt_debut_u9 = dt_pesee_u9 + timedelta(minutes=duree_pesee)
        
        if heure_pesee_u11:
            dt_pesee_u11 = datetime.combine(datetime.today(), heure_pesee_u11)
            dt_debut_u11_theorique = dt_pesee_u11 + timedelta(minutes=duree_pesee)
        else:
            dt_pesee_u11, dt_debut_u11_theorique = None, None

        dt_pause_debut = datetime.combine(datetime.today(), pause_debut) if activer_pause else None
        dt_pause_fin = datetime.combine(datetime.today(), pause_fin) if activer_pause else None
        
        tapis_dispo = [dt_debut_u9 for _ in range(nb_tapis)]
        planning_tapis = {t: [] for t in range(nb_tapis)}
        last_match_time = {} 
        total_matchs_calcules = 0

        def executer_vagues(poules_du_tapis, t_idx, heure_actuelle, duree_combat):
            nonlocal total_matchs_calcules
            # On groupe les poules par paquet de 3 (Vagues)
            vagues = [poules_du_tapis[i:i+3] for i in range(0, len(poules_du_tapis), 3)]
            
            for vague in vagues:
                matches_vague = []
                max_r = max((len(p['rondes']) for p in vague), default=0)
                # Entrelacement parfait des matchs
                for r in range(max_r):
                    for p in vague:
                        if r < len(p['rondes']):
                            for m in p['rondes'][r]: matches_vague.append((p['nom'], m))
                
                # Exécution des matchs de la vague
                for cat, m in matches_vague:
                    p1, p2 = m[0]['Nom'], m[1]['Nom']
                    
                    if activer_pause and dt_pause_debut <= heure_actuelle < dt_pause_fin:
                        planning_tapis[t_idx].append({"Type": "PAUSE", "Heure": dt_pause_debut.strftime("%H:%M")})
                        heure_actuelle = max(heure_actuelle, dt_pause_fin)
                    
                    dispo = max(last_match_time.get(p1, heure_actuelle), last_match_time.get(p2, heure_actuelle))
                    if dispo > heure_actuelle:
                        if activer_pause and heure_actuelle < dt_pause_debut and dispo > dt_pause_debut:
                            dispo = dt_pause_debut
                        attente = int((dispo - heure_actuelle).total_seconds() // 60)
                        if attente > 0:
                            planning_tapis[t_idx].append({"Type": "ATTENTE", "Heure": heure_actuelle.strftime("%H:%M"), "Texte": f"⏳ Repos ({attente} min)"})
                        heure_actuelle = dispo
                        
                    if activer_pause and dt_pause_debut <= heure_actuelle < dt_pause_fin:
                        planning_tapis[t_idx].append({"Type": "PAUSE", "Heure": dt_pause_debut.strftime("%H:%M")})
                        heure_actuelle = max(heure_actuelle, dt_pause_fin)

                    planning_tapis[t_idx].append({
                        "Type": "MATCH", "Heure": heure_actuelle.strftime("%H:%M"), "Duree": duree_combat,
                        "Cat": cat, "Combattant 1": p1, "Combattant 2": p2
                    })
                    
                    total_matchs_calcules += 1
                    fin_match = heure_actuelle + timedelta(minutes=duree_combat)
                    repos = timedelta(minutes=(repos_matchs * duree_combat))
                    last_match_time[p1] = fin_match + repos
                    last_match_time[p2] = fin_match + repos
                    heure_actuelle = fin_match
                    
            return heure_actuelle

        # --- PHASE 1 : U9 ---
        for t in range(nb_tapis):
            if tapis_poules_u9[t]:
                tapis_dispo[t] = executer_vagues(tapis_poules_u9[t], t, tapis_dispo[t], duree_u9)

        # --- SYNCHRONISATION AVANT U11 ---
        fin_u9_globale = max(tapis_dispo) if total_matchs_calcules > 0 else dt_debut_u9
        debut_u11_reel = max(fin_u9_globale, dt_debut_u11_theorique) if dt_debut_u11_theorique else fin_u9_globale
        
        for t in range(nb_tapis):
            if tapis_poules_u11[t] and tapis_dispo[t] < debut_u11_reel:
                attente = int((debut_u11_reel - tapis_dispo[t]).total_seconds() // 60)
                if attente > 0:
                    planning_tapis[t].append({"Type": "ATTENTE", "Heure": tapis_dispo[t].strftime("%H:%M"), "Texte": f"Attente lancement U11"})
                tapis_dispo[t] = debut_u11_reel

        # --- PHASE 2 : U11 ---
        for t in range(nb_tapis):
            if tapis_poules_u11[t]:
                tapis_dispo[t] = executer_vagues(tapis_poules_u11[t], t, tapis_dispo[t], duree_u11)

        fin_estimee = max(tapis_dispo)

        # --- EXPORT EXCEL ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            
            # --- ONGLET 1 : RÉSUMÉ ---
            resume_data = [
                {"Étape de la journée": "Pesée Générale (U9)", "Horaire / Valeur": dt_pesee_u9.strftime('%H:%M')},
                {"Étape de la journée": "Début de la compétition U9", "Horaire / Valeur": dt_debut_u9.strftime('%H:%M')}
            ]
            if "2" in type_pesee:
                resume_data.append({"Étape de la journée": "Pesée U11", "Horaire / Valeur": dt_pesee_u11.strftime('%H:%M')})
            resume_data.append({"Étape de la journée": "Début effectif U11", "Horaire / Valeur": debut_u11_reel.strftime('%H:%M')})
            if activer_pause:
                resume_data.append({"Étape de la journée": "Pause Déjeuner", "Horaire / Valeur": f"{pause_debut.strftime('%H:%M')} - {pause_fin.strftime('%H:%M')}"})
            
            resume_data.extend([
                {"Étape de la journée": "Fin de compétition estimée", "Horaire / Valeur": fin_estimee.strftime('%H:%M')},
                {"Étape de la journée": "Nombre total de matchs", "Horaire / Valeur": str(total_matchs_calcules)}
            ])
            pd.DataFrame(resume_data).to_excel(writer, sheet_name="Résumé", index=False)
            
            # --- ONGLET 2 : GRILLE DE PASSAGE ---
            max_lignes = max(len(liste) for liste in planning_tapis.values()) if planning_tapis else 0
            grille = []
            for row_idx in range(max_lignes):
                ligne = {}
                for t in range(nb_tapis):
                    col = f"Tapis {t + 1}"
                    if row_idx < len(planning_tapis[t]):
                        m = planning_tapis[t][row_idx]
                        if m["Type"] == "PAUSE": ligne[col] = f"[{m['Heure']}]\n⏸️ PAUSE DÉJEUNER"
                        elif m["Type"] == "ATTENTE": ligne[col] = f"[{m['Heure']}]\n{m['Texte']}"
                        else: ligne[col] = f"🕘 {m['Heure']} ({m['Duree']} min)\n[{m['Cat']}]\n{m['Combattant 1']} VS {m['Combattant 2']}"
                    else: ligne[col] = ""
                grille.append(ligne)
            pd.DataFrame(grille).to_excel(writer, sheet_name="Grille de Passage", index=False, startrow=1)
            
            from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
            b_style = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            bleu = PatternFill("solid", fgColor="0055A4")
            rouge = PatternFill("solid", fgColor="EF4135")
            bleu_clair = PatternFill("solid", fgColor="DDEBF7") 
            
            # Design Résumé
            ws_res = writer.sheets["Résumé"]
            for cell in ws_res[1]: 
                cell.fill, cell.font, cell.alignment = bleu, Font(bold=True, color="FFFFFF"), Alignment(horizontal="center")
            ws_res.column_dimensions['A'].width = 50
            ws_res.column_dimensions['B'].width = 25
            for row in ws_res.iter_rows(min_row=2, max_row=ws_res.max_row):
                for cell in row:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(size=12)
            
            # Design Grille
            ws_grille = writer.sheets["Grille de Passage"]
            ws_grille.row_dimensions[1].height = 65
            ws_grille.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nb_tapis)
            titre_cell = ws_grille.cell(row=1, column=1, value="🏆 PLANNING OFFICIEL DES COMBATS - FFLDA 🏆")
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
                        elif "Attente" in str(cell.value): cell.fill, cell.font = PatternFill("solid", fgColor="EFEFEF"), Font(italic=True, color="666666", size=11)
                        else:
                            cell.fill = bleu_clair if is_even else PatternFill(fill_type=None)
                            cell.font = Font(size=12)
            
            # --- ONGLET 3 : FEUILLES DE POULES ---
            ws_poules = writer.book.create_sheet("Feuilles de Poules")
            row_cursor = 1
            gris_fonce = PatternFill("solid", fgColor="404040")
            entete_noir = PatternFill("solid", fgColor="000000")
            
            for nom_poule, liste_p in participants_par_poule.items():
                n_lutteurs = len(liste_p)
                ws_poules.cell(row=row_cursor, column=1, value=nom_poule).font = Font(bold=True, size=14, color="0055A4")
                row_cursor += 2
                
                headers = ["N°", "NOM", "CLUB"] + [str(i) for i in range(1, n_lutteurs+1)] + ["Matchs Gagnés", "Matchs Perdus", "Total Pts", "Classement"]
                for col_idx, h in enumerate(headers, 1):
                    c = ws_poules.cell(row=row_cursor, column=col_idx, value=h)
                    c.font, c.alignment, c.border = Font(bold=True), Alignment(horizontal="center", vertical="center", wrap_text=True), b_style
                    lettre_col = ws_poules.cell(row=row_cursor, column=col_idx).column_letter
                    if 3 < col_idx <= 3 + n_lutteurs:
                        ws_poules.column_dimensions[lettre_col].width = 5
                        c.fill, c.font = entete_noir, Font(bold=True, color="FFFFFF")
                    elif col_idx > 3 + n_lutteurs:
                        ws_poules.column_dimensions[lettre_col].width = 12
                        
                ws_poules.column_dimensions['A'].width, ws_poules.column_dimensions['B'].width, ws_poules.column_dimensions['C'].width = 5, 25, 20
                row_cursor += 1
                
                for i, p in enumerate(liste_p, 1):
                    ws_poules.row_dimensions[row_cursor].height = 25 
                    ws_poules.cell(row=row_cursor, column=1, value=i).border = b_style
                    ws_poules.cell(row=row_cursor, column=1).alignment = Alignment(horizontal="center", vertical="center")
                    ws_poules.cell(row=row_cursor, column=2, value=p['Nom']).border = b_style
                    ws_poules.cell(row=row_cursor, column=3, value=p.get('Club', '')).border = b_style
                    
                    for j in range(1, n_lutteurs+1):
                        c = ws_poules.cell(row=row_cursor, column=3+j)
                        c.border = b_style
                        if i == j: c.fill = gris_fonce
                            
                    for j in range(4): ws_poules.cell(row=row_cursor, column=3+n_lutteurs+1+j).border = b_style
                    row_cursor += 1
                    
                row_cursor += 1
                rondes = rondes_par_categorie[nom_poule]
                ordre = []
                for r in rondes:
                    for m in r:
                        idx1 = next((i+1 for i, x in enumerate(liste_p) if x['Nom'] == m[0]['Nom']), None)
                        idx2 = next((i+1 for i, x in enumerate(liste_p) if x['Nom'] == m[1]['Nom']), None)
                        if idx1 and idx2: ordre.append(f"{idx1}-{idx2}")
                
                ws_poules.cell(row=row_cursor, column=1, value="Ordre des combats : " + " // ".join(ordre)).font = Font(italic=True, bold=True)
                row_cursor += 4 

        st.subheader("📊 Statistiques")
        st.metric("Matchs générés", total_matchs_calcules)

        st.download_button(label="📥 Télécharger le Planning & Poules", data=output.getvalue(), file_name="Tournoi_U9_U11.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"Une erreur est survenue : {e}")

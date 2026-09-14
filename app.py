import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
import io
import urllib.request
import streamlit.components.v1 as components

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
    
    label_pesee_1 = "1ère pesée" if "1" in type_pesee else "Pesée U9"
    heure_pesee_u9 = st.time_input(label_pesee_1, value=time(9, 0))
    
    duree_pesee = st.selectbox("Durée allouée à la pesée (min) + échauffement", [30, 45, 60, 90], index=1)
        
    st.subheader("2. Pause de la compétition")
    activer_pause = st.checkbox("Activer la pause de la compétition", value=True)
    if activer_pause:
        duree_pause = st.selectbox("Durée de la pause (min)", [30, 45, 60, 75, 90], index=2)
    else:
        duree_pause = 0
    
    st.subheader("3. Règles Sportives")
    mixte_active = st.checkbox("Catégories Mixtes (U9/U11 filles et garçons ensemble)", value=True)
    tolerance_poids = st.number_input("Tolérance d'écart de poids (%)", min_value=10, max_value=15, value=10, step=1)
    repos_matchs = st.number_input("Matchs de repos minimum", min_value=1, max_value=10, value=3)
    
    st.subheader("4. Temps des Combats (Match + Rotation)")
    duree_u9 = st.number_input("Temps total U9 (min)", value=3)
    duree_u11 = st.number_input("Temps total U11 (min)", value=4)

# --- CORPS PRINCIPAL ---
st.title("Générateur de Planning FFLDA 🚀")
st.markdown("**Outil officiel d'optimisation (Compatible imports Exalto)**")
st.markdown("---")

def generer_rondes_fflda(participants_in):
    participants = list(participants_in)
    n = len(participants)
    
    if n == 3:
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

def bouton_imprimer(label="🖨️ Imprimer cette vue"):
    print_code = f"""
    {label}
    """
    components.html(print_code, height=50)

fichier_upload = st.file_uploader("📂 Importez votre liste d'inscrits (.csv ou .xlsx)", type=["xlsx", "csv"])

if fichier_upload is not None:
    try:
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
        
        poules_u9, poules_u11 = [], []
        multiplicateur_poids = 1 + (tolerance_poids / 100.0)
        
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
                        if p['Poids_Num'] <= (poids_min * multiplicateur_poids) and len(poule_courante) < max_size:
                            poule_courante.append(p)
                        else:
                            nom_groupe = f"{age} | {sexe}{suffixe_niveau} | Gr. {index_poule} ({poule_courante[0]['Poids_Num']}kg - {poule_courante[-1]['Poids_Num']}kg)"
                            poule_obj = {'nom': nom_groupe, 'participants': list(poule_courante), 'rondes': generer_rondes_fflda(poule_courante)}
                            if age == 'U9': poules_u9.append(poule_obj)
                            else: poules_u11.append(poule_obj)
                            index_poule += 1
                            poule_courante = [p]
                if poule_courante:
                    nom_groupe = f"{age} | {sexe}{suffixe_niveau} | Gr. {index_poule} ({poule_courante[0]['Poids_Num']}kg - {poule_courante[-1]['Poids_Num']}kg)"
                    poule_obj = {'nom': nom_groupe, 'participants': list(poule_courante), 'rondes': generer_rondes_fflda(poule_courante)}
                    if age == 'U9': poules_u9.append(poule_obj)
                    else: poules_u11.append(poule_obj)
                    index_poule += 1

        participants_par_poule = {p['nom']: p['participants'] for p in poules_u9 + poules_u11}
        rondes_par_categorie = {p['nom']: p['rondes'] for p in poules_u9 + poules_u11}

        tapis_poules_u9 = {i: [] for i in range(nb_tapis)}
        tapis_poules_u11 = {i: [] for i in range(nb_tapis)}
        
        for i, p in enumerate(poules_u9): tapis_poules_u9[i % nb_tapis].append(p)
        for i, p in enumerate(poules_u11): tapis_poules_u11[i % nb_tapis].append(p)

        dt_pesee_u9 = datetime.combine(datetime.today(), heure_pesee_u9)
        dt_debut_u9 = dt_pesee_u9 + timedelta(minutes=duree_pesee)
        
        tapis_dispo = [dt_debut_u9 for _ in range(nb_tapis)]
        planning_tapis = {t: [] for t in range(nb_tapis)}
        last_match_time = {} 
        total_matchs_calcules = 0

        def executer_vagues(poules_du_tapis, t_idx, heure_actuelle, duree_combat):
            global total_matchs_calcules
            vagues = [poules_du_tapis[i:i+3] for i in range(0, len(poules_du_tapis), 3)]
            
            for vague in vagues:
                matches_vague = []
                max_r = max((len(p['rondes']) for p in vague), default=0)
                for r in range(max_r):
                    for p in vague:
                        if r < len(p['rondes']):
                            for m in p['rondes'][r]: matches_vague.append((p['nom'], m))
                
                for cat, m in matches_vague:
                    p1, p2 = m[0]['Nom'], m[1]['Nom']
                    
                    dispo = max(last_match_time.get(p1, heure_actuelle), last_match_time.get(p2, heure_actuelle))
                    if dispo > heure_actuelle:
                        attente = int((dispo - heure_actuelle).total_seconds() // 60)
                        if attente > 0:
                            planning_tapis[t_idx].append({"Type": "ATTENTE", "Heure": heure_actuelle.strftime("%H:%M"), "Texte": f"⏳ Repos ({attente} min)"})
                        heure_actuelle = dispo

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

        for t in range(nb_tapis):
            if tapis_poules_u9[t]:
                tapis_dispo[t] = executer_vagues(tapis_poules_u9[t], t, tapis_dispo[t], duree_u9)

        fin_u9_globale = max(tapis_dispo) if total_matchs_calcules > 0 else dt_debut_u9

        dt_pesee_u11 = None
        if "2" in type_pesee:
            dt_pesee_u11 = fin_u9_globale + timedelta(minutes=duree_pause)
            dt_debut_u11_theorique = dt_pesee_u11 + timedelta(minutes=duree_pesee)
        else:
            dt_debut_u11_theorique = fin_u9_globale

        if activer_pause and duree_pause > 0:
            for t in range(nb_tapis):
                planning_tapis[t].append({"Type": "PAUSE", "Heure": fin_u9_globale.strftime("%H:%M")})
                tapis_dispo[t] = fin_u9_globale + timedelta(minutes=duree_pause)
        else:
            for t in range(nb_tapis):
                tapis_dispo[t] = fin_u9_globale

        debut_u11_reel = max(tapis_dispo)
        if dt_debut_u11_theorique and debut_u11_reel < dt_debut_u11_theorique:
            debut_u11_reel = dt_debut_u11_theorique

        for t in range(nb_tapis):
            if tapis_poules_u11[t] and tapis_dispo[t] < debut_u11_reel:
                attente = int((debut_u11_reel - tapis_dispo[t]).total_seconds() // 60)
                if attente > 0:
                    planning_tapis[t].append({"Type": "ATTENTE", "Heure": tapis_dispo[t].strftime("%H:%M"), "Texte": f"Attente lancement U11"})
                tapis_dispo[t] = debut_u11_reel

        for t in range(nb_tapis):
            if tapis_poules_u11[t]:
                tapis_dispo[t] = executer_vagues(tapis_poules_u11[t], t, tapis_dispo[t], duree_u11)

        fin_estimee = max(tapis_dispo)

        texte_pesee_u9 = "1ère pesée" if "1" in type_pesee else "Pesée U9"
        valeur_pause = f"{duree_pause} min" if (activer_pause and duree_pause > 0) else "0 min"

        str_comp_u9 = f"{dt_debut_u9.strftime('%H:%M')} - {fin_u9_globale.strftime('%H:%M')}"
        str_comp_u11 = f"{debut_u11_reel.strftime('%H:%M')} - {fin_estimee.strftime('%H:%M')}"

        st.success("Fichier analysé avec succès !")
        
        # --- ONGLETS INTERACTIFS DE L'APPLICATION ---
        noms_onglets = ["📊 Résumé & Stats", "📅 Grille de Passage par Tapis"] + [f"Poule : {p[:15]}" for p in participants_par_poule.keys()]
        onglets_ui = st.tabs(noms_onglets)
        
        # Onglet 1 : Résumé
        with onglets_ui[0]:
            st.subheader("📊 Résumé prévisionnel de la journée")
            lignes_accueil = [
                {"Étape de la journée": texte_pesee_u9, "Horaire / Valeur": dt_pesee_u9.strftime('%H:%M')},
                {"Étape de la journée": "Compétition U9", "Horaire / Valeur": str_comp_u9},
                {"Étape de la journée": "Pause de la compétition", "Horaire / Valeur": valeur_pause}
            ]
            if "2" in type_pesee and dt_pesee_u11:
                lignes_accueil.append({"Étape de la journée": "2ème pesée", "Horaire / Valeur": dt_pesee_u11.strftime('%H:%M')})

            lignes_accueil.extend([
                {"Étape de la journée": "Compétition U11", "Horaire / Valeur": str_comp_u11},
                {"Étape de la journée": "Fin de la compétition estimée", "Horaire / Valeur": fin_estimee.strftime('%H:%M')},
                {"Étape de la journée": "Nombre total de matchs", "Horaire / Valeur": str(total_matchs_calcules)}
            ])
            st.table(pd.DataFrame(lignes_accueil))
            st.metric("Matchs générés", total_matchs_calcules)
            bouton_imprimer("🖨️ Imprimer ce Résumé")

        # Onglet 2 : Grille de Passage
        with onglets_ui[1]:
            st.subheader("📅 Grille de Passage - Tapis")
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
                        else: ligne[col] = f"[{m['Heure']}] ({m['Duree']}m) [{m['Cat']}] - {m['Combattant 1']} vs {m['Combattant 2']}"
                    else: ligne[col] = ""
                grille_ui.append(ligne)
            st.dataframe(pd.DataFrame(grille_ui), use_container_width=True)
            bouton_imprimer("🖨️ Imprimer la Grille de Passage")

        # Onglets suivants : Chaque Poule
        for idx, (nom_poule, liste_p) in enumerate(participants_par_poule.items(), start=2):
            with onglets_ui[idx]:
                st.subheader(f"Feuille de Poule : {nom_poule}")
                df_poule_vue = pd.DataFrame(liste_p)[['Nom', 'Club', 'Poids']]
                st.table(df_poule_vue)
                bouton_imprimer(f"🖨️ Imprimer cette Feuille de Poule")

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
                {"Étape de la journée": "Fin de la compétition estimée", "Horaire / Valeur": fin_estimee.strftime('%H:%M')},
                {"Étape de la journée": "Nombre total de matchs", "Horaire / Valeur": str(total_matchs_calcules)}
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
                        else: ligne[col] = f"🕘 {m['Heure']} ({m['Duree']} min)\n[{m['Cat']}]\n{m['Combattant 1']} VS {m['Combattant 2']}"
                    else: ligne[col] = ""
                grille.append(ligne)
            pd.DataFrame(grille).to_excel(writer, sheet_name="Grille de Passage", index=False, startrow=1)
            
            from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
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
            
            rouge_lutte = PatternFill("solid", fgColor="E53935") 
            bleu_lutte = PatternFill("solid", fgColor="1E88E5")  
            entete_noir = PatternFill("solid", fgColor="000000")
            gris_clair = PatternFill("solid", fgColor="F2F2F2")
            
            for nom_poule, liste_p in participants_par_poule.items():
                nom_onglet_court = nom_poule.replace(" | ", " ").replace("(", "").replace(")", "").replace(" - ", "-")[:31].strip()
                ws_poule = writer.book.create_sheet(nom_onglet_court)
                
                ws_poule.cell(row=1, column=1, value=f"POULE : {nom_poule}").font = Font(bold=True, size=16, color="0055A4")
                ws_poule.cell(row=2, column=1, value="*POINT DE CLASSEMENT : 2 pt = victoire - 1 pt = match nul - 0 pt = défaite").font = Font(italic=True, size=9)
                
                row_cursor = 4
                headers = ["CLT", "N°", "NOM Prénom", "CLUB"]
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
                
                largeur_nom_col = max(max_len_nom + 4, 25)
                largeur_club_col = max(max_len_club + 4, 18)

                ws_poule.column_dimensions['A'].width = 6
                ws_poule.column_dimensions['B'].width = 6
                ws_poule.column_dimensions['C'].width = largeur_nom_col  
                ws_poule.column_dimensions['D'].width = largeur_club_col 
                ws_poule.column_dimensions['E'].width = 10               
                ws_poule.column_dimensions['F'].width = 6                
                ws_poule.column_dimensions['G'].width = largeur_nom_col  
                ws_poule.column_dimensions['H'].width = largeur_club_col 
                ws_poule.column_dimensions['I'].width = 10               
                
                # Dictionnaire pour retrouver facilement la ligne Excel de chaque lutteur dans le tableau du haut
                lignes_lutteurs = {}
                row_cursor += 1
                for i, p in enumerate(liste_p, 1):
                    lignes_lutteurs[p['Nom']] = row_cursor
                    ws_poule.cell(row=row_cursor, column=1).border = b_style 
                    ws_poule.cell(row=row_cursor, column=2, value=i).border = b_style 
                    ws_poule.cell(row=row_cursor, column=2).alignment = Alignment(horizontal="center")
                    ws_poule.cell(row=row_cursor, column=3, value=p['Nom']).border = b_style
                    ws_poule.cell(row=row_cursor, column=4, value=p.get('Club', '')).border = b_style
                    
                    col_offset = 5
                    for t in range(nb_tours):
                        ws_poule.cell(row=row_cursor, column=col_offset+t).border = b_style 
                    
                    ws_poule.cell(row=row_cursor, column=col_offset+nb_tours).border = b_style 
                    ws_poule.cell(row=row_cursor, column=col_offset+nb_tours+1).border = b_style 
                    ws_poule.cell(row=row_cursor, column=col_offset+nb_tours+2, value=p.get('Poids', '')).border = b_style 
                    ws_poule.cell(row=row_cursor, column=col_offset+nb_tours+2).alignment = Alignment(horizontal="center")
                    row_cursor += 1
                
                row_cursor += 2
                
                rondes = rondes_par_categorie[nom_poule]
                col_offset_tours = 5 # Colonne E (1er tour)
                
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
                        
                        # Ligne des noms des lutteurs dans le bloc match
                        r_nom_, c_nom_ = row_cursor, 3
                        r_nom_bleu, c_nom_bleu = row_cursor, 7
                        
                        ws_poule.cell(row=row_cursor, column=2, value=idx1).alignment = Alignment(horizontal="center")
                        ws_poule.cell(row=row_cursor, column=2).font = Font(bold=True, color="E53935", size=14)
                        ws_poule.cell(row=row_cursor, column=3, value=p1['Nom']).border = b_style
                        ws_poule.cell(row=row_cursor, column=4, value=p1.get('Club', '')).border = b_style
                        
                        box_ptr = ws_poule.cell(row=row_cursor, column=5)
                        box_ptr.border, box_ptr.fill = b_style, gris_clair
                        
                        ws_poule.cell(row=row_cursor, column=6, value=idx2).alignment = Alignment(horizontal="center")
                        ws_poule.cell(row=row_cursor, column=6).font = Font(bold=True, color="1E88E5", size=14)
                        ws_poule.cell(row=row_cursor, column=7, value=p2['Nom']).border = b_style
                        ws_poule.cell(row=row_cursor, column=8, value=p2.get('Club', '')).border = b_style
                        
                        box_ptb = ws_poule.cell(row=row_cursor, column=9)
                        box_ptb.border, box_ptb.fill = b_style, gris_clair
                        
                        # Liaison automatique des scores avec le tableau du haut selon le tour
                        col_tour_lettre = openpyxl.utils.get_column_letter(col_offset_tours + (tour_idx - 1))
                        if p1['Nom'] in lignes_lutteurs:
                            lig_haut_p1 = lignes_lutteurs[p1['Nom']]
                            box_ptr.value = f"={col_tour_lettre}{lig_haut_p1}"
                            box_ptr.alignment = Alignment(horizontal="center", vertical="center")
                        
                        if p2['Nom'] in lignes_lutteurs:
                            lig_haut_p2 = lignes_lutteurs[p2['Nom']]
                            box_ptb.value = f"={col_tour_lettre}{lig_haut_p2}"
                            box_ptb.alignment = Alignment(horizontal="center", vertical="center")

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
        st.error(f"Une erreur est survenue : {e}")

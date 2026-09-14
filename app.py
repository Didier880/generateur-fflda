import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
import io

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
    
    st.subheader("Logistique")
    nb_tapis = st.number_input("Nombre de tapis", min_value=1, max_value=10, value=3)
    heure_debut = st.time_input("Heure de début (U9)", value=time(9, 0))
    pause_debut = st.time_input("Début de la pause", value=time(12, 0))
    pause_fin = st.time_input("Fin de la pause", value=time(13, 0))
    
    st.subheader("Santé & Arbitrage")
    repos_matchs = st.number_input("Matchs de repos minimum", min_value=1, max_value=10, value=3)
    
    st.subheader("Durée globale (Match + Rotation)")
    duree_u9 = st.number_input("Temps U9 (min)", value=4)
    duree_u11 = st.number_input("Temps U11 (min)", value=5)
    
    durees_age = {"U9": duree_u9, "U11": duree_u11}

# --- CORPS PRINCIPAL ---
st.title("Générateur de Planning FFLDA 🚀")
st.markdown("**Outil officiel d'optimisation (Compatible imports Exalto)**")
st.markdown("---")

def generer_rondes(participants):
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

def format_duree(td):
    total_seconds = int(td.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, _ = divmod(remainder, 60)
    if hours > 0:
        return f"{hours}h {minutes}min"
    return f"{minutes}min"

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

        # --- TRADUCTION EXALTO ---
        if "Catégorie d'âge" in df_raw.columns:
            df_raw = df_raw.rename(columns={"Catégorie d'âge": "Age"})
        if "Sigle du Club" in df_raw.columns:
            df_raw = df_raw.rename(columns={"Sigle du Club": "Club"})
            
        if "Prénom" in df_raw.columns and "Nom" in df_raw.columns:
            df_raw["Nom"] = df_raw["Nom"].astype(str) + " " + df_raw["Prénom"].astype(str)

        df_inscr = df_raw.copy()
        
        # --- FILTRAGE STRICT U9 ET U11 ---
        nb_avant_filtre = len(df_inscr)
        df_inscr = df_inscr[df_inscr['Age'].isin(['U9', 'U11'])]
        nb_apres_filtre = len(df_inscr)
        
        if nb_apres_filtre < nb_avant_filtre:
            st.info(f"ℹ️ {nb_avant_filtre - nb_apres_filtre} lutteur(s) hors U9/U11 ont été ignorés pour ce tournoi spécifique.")

        if df_inscr.empty:
            st.error("❌ Aucun lutteur U9 ou U11 n'a été trouvé dans le fichier.")
            st.stop()
        
        # --- GESTION DU NIVEAU ---
        if "Maîtrise" not in df_inscr.columns:
            df_inscr["Maîtrise"] = ""
            
        def attribuer_niveau(val):
            val_str = str(val).strip().lower()
            if val_str == 'd': return 'Débutant'
            elif val_str == 'c': return 'Confirmé'
            return ''
            
        df_inscr['Niveau'] = df_inscr['Maîtrise'].apply(attribuer_niveau)

        colonnes_requises = ['Nom', 'Age', 'Sexe', 'Poids']
        for col in colonnes_requises:
            if col not in df_inscr.columns:
                st.error(f"Erreur : La colonne '{col}' manque.")
                st.stop()
                
        # --- GESTION PESÉE ---
        df_inscr['Poids'] = df_inscr['Poids'].astype(str).str.replace(',', '.')
        df_inscr['Poids_Num'] = pd.to_numeric(df_inscr['Poids'], errors='coerce')
        
        non_peses = df_inscr[(df_inscr['Poids_Num'].isna()) | (df_inscr['Poids_Num'] <= 0)]
        if not non_peses.empty:
            st.warning(f"⚠️ {len(non_peses)} lutteur(s) U9/U11 sans poids valide (0 ou vide) écartés.")
            
        df_inscr = df_inscr[df_inscr['Poids_Num'] > 0]
        
        # Mixte par défaut pour U9/U11
        df_inscr['Sexe'] = 'Mixte'
        
        rondes_par_categorie = {}
        df_morpho_valide = df_inscr.sort_values('Poids_Num')
        
        # Création des poules
        for (age, sexe, niveau), groupe in df_morpho_valide.groupby(['Age', 'Sexe', 'Niveau']):
            participants = groupe.to_dict('records')
            poule_courante = []
            index_poule = 1
            suffixe_niveau = f" | {niveau}" if niveau != "" else ""
            
            for p in participants:
                if not poule_courante:
                    poule_courante.append(p)
                else:
                    poids_min = poule_courante[0]['Poids_Num']
                    if p['Poids_Num'] <= (poids_min * 1.10) and len(poule_courante) < 6:
                        poule_courante.append(p)
                    else:
                        nom_groupe = f"{age} | {sexe}{suffixe_niveau} | Gr. {index_poule} ({poule_courante[0]['Poids_Num']}kg - {poule_courante[-1]['Poids_Num']}kg)"
                        rondes_par_categorie[nom_groupe] = generer_rondes(poule_courante)
                        index_poule += 1
                        poule_courante = [p]
            if poule_courante:
                nom_groupe = f"{age} | {sexe}{suffixe_niveau} | Gr. {index_poule} ({poule_courante[0]['Poids_Num']}kg - {poule_courante[-1]['Poids_Num']}kg)"
                rondes_par_categorie[nom_groupe] = generer_rondes(poule_courante)

        # --- SÉPARATION DES FILES D'ATTENTE U9 PUIS U11 ---
        rondes_u9 = {k: v for k, v in rondes_par_categorie.items() if k.startswith("U9")}
        rondes_u11 = {k: v for k, v in rondes_par_categorie.items() if k.startswith("U11")}
        
        file_u9 = []
        max_r_u9 = max((len(r) for r in rondes_u9.values()), default=0)
        for r in range(max_r_u9):
            for cat in rondes_u9.keys():
                if r < len(rondes_u9[cat]):
                    for match in rondes_u9[cat][r]: file_u9.append((cat, match))
                    
        file_u11 = []
        max_r_u11 = max((len(r) for r in rondes_u11.values()), default=0)
        for r in range(max_r_u11):
            for cat in rondes_u11.keys():
                if r < len(rondes_u11[cat]):
                    for match in rondes_u11[cat][r]: file_u11.append((cat, match))

        # --- GESTION DU TEMPS ET DE L'ALGORITHME ---
        dt_debut = datetime.combine(datetime.today(), heure_debut)
        dt_pause_debut = datetime.combine(datetime.today(), pause_debut)
        dt_pause_fin = datetime.combine(datetime.today(), pause_fin)
        
        tapis_dispo = [dt_debut for _ in range(nb_tapis)]
        planning_tapis = {t: [] for t in range(nb_tapis)}
        last_match_time = {} 
        total_matchs_calcules = 0

        phases = [("U9", file_u9), ("U11", file_u11)]

        for nom_phase, file_attente in phases:
            if not file_attente:
                continue
                
            # --- SYNCHRONISATION : On aligne tous les tapis à l'heure du dernier match U9 ---
            if total_matchs_calcules > 0:
                heure_synchro = max(tapis_dispo)
                for t in range(nb_tapis):
                    if tapis_dispo[t] < heure_synchro:
                        attente = (heure_synchro - tapis_dispo[t]).seconds // 60
                        if attente > 0:
                            planning_tapis[t].append({"Type": "ATTENTE", "Heure": tapis_dispo[t].strftime("%H:%M"), "Texte": f"Fin des U9 - En attente U11"})
                        tapis_dispo[t] = heure_synchro
            
            # --- TRAITEMENT DE LA FILE D'ATTENTE ---
            while file_attente:
                t_idx = tapis_dispo.index(min(tapis_dispo))
                horaire_actuel = tapis_dispo[t_idx]
                
                if dt_pause_debut <= horaire_actuel < dt_pause_fin:
                    planning_tapis[t_idx].append({"Type": "PAUSE", "Heure": dt_pause_debut.strftime("%H:%M")})
                    tapis_dispo[t_idx] = max(dt_pause_fin, horaire_actuel)
                    continue
                    
                match_found = False
                for i, (cat, match) in enumerate(file_attente):
                    p1_nom, p2_nom = match[0]["Nom"], match[1]["Nom"]
                    
                    dispo_p1 = last_match_time.get(p1_nom, horaire_actuel)
                    dispo_p2 = last_match_time.get(p2_nom, horaire_actuel)
                    
                    if dispo_p1 <= horaire_actuel and dispo_p2 <= horaire_actuel:
                        age_cat = cat.split(" | ")[0]
                        duree_match = durees_age.get(age_cat, 5)
                        duree_td = timedelta(minutes=duree_match)
                        repos_min_td = timedelta(minutes=(repos_matchs * duree_match))
                        
                        if horaire_actuel < dt_pause_debut and (horaire_actuel + duree_td) > dt_pause_debut:
                            planning_tapis[t_idx].append({"Type": "PAUSE", "Heure": horaire_actuel.strftime("%H:%M")})
                            tapis_dispo[t_idx] = dt_pause_fin
                            match_found = True
                            break 
                        
                        planning_tapis[t_idx].append({
                            "Type": "MATCH", "Heure": horaire_actuel.strftime("%H:%M"), "Duree": duree_match,
                            "Cat": cat, "Combattant 1": p1_nom, "Combattant 2": p2_nom
                        })
                        
                        total_matchs_calcules += 1
                        fin_match = horaire_actuel + duree_td
                        
                        last_match_time[p1_nom] = fin_match + repos_min_td
                        last_match_time[p2_nom] = fin_match + repos_min_td
                        
                        tapis_dispo[t_idx] = fin_match
                        file_attente.pop(i)
                        match_found = True
                        break
                        
                if not match_found:
                    next_ready_time = None
                    for cat, match in file_attente:
                        ready_at = max(last_match_time.get(match[0]['Nom'], horaire_actuel), last_match_time.get(match[1]['Nom'], horaire_actuel))
                        if next_ready_time is None or ready_at < next_ready_time:
                            next_ready_time = ready_at
                    
                    if horaire_actuel < dt_pause_debut and next_ready_time > dt_pause_debut:
                        next_ready_time = dt_pause_debut
                        
                    attente_mins = (next_ready_time - horaire_actuel).seconds // 60
                    if attente_mins > 0:
                        planning_tapis[t_idx].append({"Type": "ATTENTE", "Heure": horaire_actuel.strftime("%H:%M"), "Texte": f"⏳ Attente ({attente_mins} min)"})
                    tapis_dispo[t_idx] = next_ready_time

        fin_estimee = max(tapis_dispo) if tapis_dispo else dt_debut
        duree_totale = fin_estimee - dt_debut

        # --- EXPORT EXCEL ---
        resume_data = [
            {"Information": "Matchs générés", "Valeur": str(total_matchs_calcules)},
            {"Information": "Heure début U9", "Valeur": dt_debut.strftime('%H:%M')},
            {"Information": "Heure fin totale", "Valeur": fin_estimee.strftime('%H:%M')}
        ]
        
        max_lignes = max(len(liste) for liste in planning_tapis.values()) if planning_tapis else 0
        grille = []
        for row_idx in range(max_lignes):
            ligne_donnees = {}
            for t in range(nb_tapis):
                nom_colonne = f"Tapis {t + 1}"
                if row_idx < len(planning_tapis[t]):
                    m = planning_tapis[t][row_idx]
                    if m["Type"] == "PAUSE":
                        ligne_donnees[nom_colonne] = f"[{m['Heure']}]\n⏸️ PAUSE DÉJEUNER"
                    elif m["Type"] == "ATTENTE":
                        ligne_donnees[nom_colonne] = f"[{m['Heure']}]\n{m['Texte']}"
                    else:
                        ligne_donnees[nom_colonne] = f"🕘 {m['Heure']} ({m['Duree']} min)\n[{m['Cat']}]\n{m['Combattant 1']} VS {m['Combattant 2']}"
                else:
                    ligne_donnees[nom_colonne] = ""
            grille.append(ligne_donnees)
            
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame(resume_data).to_excel(writer, sheet_name="Résumé", index=False)
            df_grille = pd.DataFrame(grille)
            df_grille.to_excel(writer, sheet_name="Grille de Passage", index=False)
            
            from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
            bleu, rouge, gris = PatternFill("solid", fgColor="0055A4"), PatternFill("solid", fgColor="EF4135"), PatternFill("solid", fgColor="F2F2F2")
            b_style = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            ws_res = writer.sheets["Résumé"]
            for cell in ws_res[1]: cell.fill, cell.font = bleu, Font(bold=True, color="FFFFFF")
            ws_res.column_dimensions['A'].width, ws_res.column_dimensions['B'].width = 30, 25
            
            ws_grille = writer.sheets["Grille de Passage"]
            for cell in ws_grille[1]:
                cell.fill, cell.font, cell.border = bleu, Font(bold=True, color="FFFFFF"), b_style
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for col in range(1, nb_tapis + 1):
                ws_grille.column_dimensions[ws_grille.cell(row=1, column=col).column_letter].width = 45
                
            for row in ws_grille.iter_rows(min_row=2, max_row=ws_grille.max_row):
                ws_grille.row_dimensions[row[0].row].height = 90
                for cell in row:
                    cell.border, cell.alignment = b_style, Alignment(wrap_text=True, horizontal="center", vertical="center")
                    if cell.value:
                        if "PAUSE" in str(cell.value): cell.fill, cell.font = rouge, Font(bold=True, color="FFFFFF")
                        elif "Attente" in str(cell.value) or "Fin des U9" in str(cell.value): cell.fill, cell.font = gris, Font(italic=True, color="666666")

        st.subheader("📊 Statistiques")
        col1, col2, col3 = st.columns(3)
        col1.metric("Matchs générés", total_matchs_calcules)
        col2.metric("Fin estimée", fin_estimee.strftime('%H:%M'))
        col3.metric("Durée totale", format_duree(duree_totale))

        st.download_button(label="📥 Télécharger le Planning", data=output.getvalue(), file_name="Planning_U9_U11_FFLDA.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"Une erreur est survenue : {e}")

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
    
    st.subheader("Logistique")
    nb_tapis = st.number_input("Nombre de tapis", min_value=1, max_value=10, value=3)
    heure_debut = st.time_input("Heure de début", value=time(9, 0))
    pause_debut = st.time_input("Début de la pause", value=time(12, 0))
    pause_fin = st.time_input("Fin de la pause", value=time(13, 0))
    
    st.subheader("Santé & Arbitrage")
    repos_matchs = st.number_input("Matchs de repos minimum", min_value=1, max_value=10, value=3)
    
    st.subheader("Durée globale (Match + Rotation)")
    duree_u9 = st.number_input("Temps U9 (min)", value=4)
    duree_u11 = st.number_input("Temps U11 (min)", value=5)
    duree_u13 = st.number_input("Temps U13 (min)", value=6)
    duree_u15 = st.number_input("Temps U15 (min)", value=6)
    duree_u17 = st.number_input("Temps U17 (min)", value=7)
    duree_senior = st.number_input("Temps Senior/Autre (min)", value=8)
    
    durees_age = {
        "U9": duree_u9, "U11": duree_u11, "U13": duree_u13,
        "U15": duree_u15, "U17": duree_u17
    }

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

# Accepte CSV (Exalto) et Excel
fichier_upload = st.file_uploader("📂 Importez votre liste d'inscrits (.csv ou .xlsx)", type=["xlsx", "csv"])

if fichier_upload is not None:
    try:
        # --- LECTURE DU FICHIER (CSV ou EXCEL) ---
        if fichier_upload.name.endswith('.csv'):
            df_raw = pd.read_csv(fichier_upload, sep=';', encoding='utf-8')
            if len(df_raw.columns) == 1:
                fichier_upload.seek(0)
                df_raw = pd.read_csv(fichier_upload, sep=',', encoding='utf-8')
        else:
            # ASTUCE EXALTO : Recherche automatique de la ligne d'en-têtes
            df_temp = pd.read_excel(fichier_upload, nrows=5)
            header_row = 0
            for i, row in df_temp.iterrows():
                if 'N° Licence' in str(row.values) or 'Nom' in str(row.values) or "Catégorie d'âge" in str(row.values):
                    header_row = i + 1
                    break
            fichier_upload.seek(0)
            df_raw = pd.read_excel(fichier_upload, header=header_row)

        # --- TRADUCTION DU LANGAGE EXALTO ---
        if "Catégorie d'âge" in df_raw.columns:
            df_raw = df_raw.rename(columns={"Catégorie d'âge": "Age"})
        if "Sigle du Club" in df_raw.columns:
            df_raw = df_raw.rename(columns={"Sigle du Club": "Club"})
            
        # Fusionner Nom et Prénom pour l'affichage
        if "Prénom" in df_raw.columns and "Nom" in df_raw.columns:
            df_raw["Nom"] = df_raw["Nom"].astype(str) + " " + df_raw["Prénom"].astype(str)

        df_inscr = df_raw.copy()
        
        # --- GESTION DU NIVEAU (Maîtrise) ---
        if "Maîtrise" not in df_inscr.columns:
            df_inscr["Maîtrise"] = ""
            
        def attribuer_niveau(val):
            val_str = str(val).strip().lower()
            if val_str == 'd':
                return 'Débutant'
            elif val_str == 'c':
                return 'Confirmé'
            return '' # Vide si non renseigné ou autre lettre
            
        df_inscr['Niveau'] = df_inscr['Maîtrise'].apply(attribuer_niveau)

        # Vérification des colonnes vitales
        colonnes_requises = ['Nom', 'Age', 'Sexe', 'Poids']
        for col in colonnes_requises:
            if col not in df_inscr.columns:
                st.error(f"Erreur : La colonne '{col}' manque dans votre fichier.")
                st.stop()
                
        # --- GESTION DES POIDS ET DE LA PESÉE ---
        df_inscr['Poids'] = df_inscr['Poids'].astype(str).str.replace(',', '.')
        df_inscr['Poids_Num'] = pd.to_numeric(df_inscr['Poids'], errors='coerce')
        
        non_peses = df_inscr[(df_inscr['Poids_Num'].isna()) | (df_inscr['Poids_Num'] <= 0)]
        if not non_peses.empty:
            st.warning(f"⚠️ Attention : {len(non_peses)} lutteur(s) (ex: {non_peses.iloc[0]['Nom']}) n'ont pas de poids valide (0 ou case vide). Ils ont été totalement écartés du tirage.")
            
        df_inscr = df_inscr[df_inscr['Poids_Num'] > 0]
        
        # Règle Mixte U9/U11
        mask_mixte = df_inscr['Age'].isin(['U9', 'U11'])
        if mask_mixte.any():
            df_inscr.loc[mask_mixte, 'Sexe'] = 'Mixte'
            st.info("ℹ️ Règle FFLDA appliquée : Les U9 et U11 ont été classés en catégorie 'Mixte'.")
        
        df_inscr = df_inscr.dropna(subset=['Sexe'])

        rondes_par_categorie = {}
        groupes_scindes = 0
        
        mask_morpho = df_inscr['Age'].isin(['U9', 'U11'])
        df_morpho = df_inscr[mask_morpho].copy()
        df_autres = df_inscr[~mask_morpho].copy()

        # --- CRÉATION DES POULES CLASSIQUES (U13, U15...) ---
        if not df_autres.empty:
            # On inclut le niveau dans le nom du groupe s'il existe
            df_autres['Groupe_Complet'] = df_autres['Age'].astype(str) + " | " + df_autres['Sexe'].astype(str) + " | " + df_autres['Poids'].astype(str)
            mask_niveau_autres = df_autres['Niveau'] != ""
            df_autres.loc[mask_niveau_autres, 'Groupe_Complet'] = df_autres['Groupe_Complet'] + " | " + df_autres['Niveau']
            
            for categorie, groupe in df_autres.groupby('Groupe_Complet'):
                participants = groupe.to_dict('records')
                if len(participants) > 6:
                    poule_a = participants[::2]
                    poule_b = participants[1::2]
                    rondes_par_categorie[categorie + " | Poule A"] = generer_rondes(poule_a)
                    rondes_par_categorie[categorie + " | Poule B"] = generer_rondes(poule_b)
                    groupes_scindes += 1
                else:
                    rondes_par_categorie[categorie] = generer_rondes(participants)

        # --- CRÉATION DES POULES MORPHOLOGIQUES (U9, U11) ---
        if not df_morpho.empty:
            df_morpho_valide = df_morpho.sort_values('Poids_Num')
            # On groupe par Age, Sexe ET Niveau !
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
                        poids_actuel = p['Poids_Num']
                        if poids_actuel <= (poids_min * 1.10) and len(poule_courante) < 6:
                            poule_courante.append(p)
                        else:
                            nom_groupe = f"{age} | {sexe}{suffixe_niveau} | Gr. {index_poule} ({poule_courante[0]['Poids_Num']}kg - {poule_courante[-1]['Poids_Num']}kg)"
                            rondes_par_categorie[nom_groupe] = generer_rondes(poule_courante)
                            index_poule += 1
                            poule_courante = [p]
                if poule_courante:
                    nom_groupe = f"{age} | {sexe}{suffixe_niveau} | Gr. {index_poule} ({poule_courante[0]['Poids_Num']}kg - {poule_courante[-1]['Poids_Num']}kg)"
                    rondes_par_categorie[nom_groupe] = generer_rondes(poule_courante)

        if groupes_scindes > 0:
            st.warning(f"⚠️ {groupes_scindes} catégorie(s) classique(s) dépassant 6 lutteurs ont été automatiquement scindées (Poules A et B).")

        # --- GESTION DU TEMPS ET DE L'ALGORITHME ---
        max_rondes = max((len(rondes) for rondes in rondes_par_categorie.values()), default=0)
        file_attente_globale = []
        for r in range(max_rondes):
            for cat in rondes_par_categorie.keys():
                if r < len(rondes_par_categorie[cat]):
                    for match in rondes_par_categorie[cat][r]:
                        file_attente_globale.append((cat, match))

        dt_debut = datetime.combine(datetime.today(), heure_debut)
        dt_pause_debut = datetime.combine(datetime.today(), pause_debut)
        dt_pause_fin = datetime.combine(datetime.today(), pause_fin)
        
        tapis_dispo = [dt_debut for _ in range(nb_tapis)]
        planning_tapis = {t: [] for t in range(nb_tapis)}
        last_match_time = {} 
        total_matchs_calcules = 0

        while file_attente_globale:
            t_idx = tapis_dispo.index(min(tapis_dispo))
            horaire_actuel = tapis_dispo[t_idx]
            
            if dt_pause_debut <= horaire_actuel < dt_pause_fin:
                planning_tapis[t_idx].append({"Type": "PAUSE", "Heure": dt_pause_debut.strftime("%H:%M")})
                tapis_dispo[t_idx] = max(dt_pause_fin, horaire_actuel)
                continue
                
            match_found = False
            for i, (cat, match) in enumerate(file_attente_globale):
                p1_nom = match[0]["Nom"]
                p2_nom = match[1]["Nom"]
                id_p1 = p1_nom
                id_p2 = p2_nom
                
                dispo_p1 = last_match_time.get(id_p1, horaire_actuel)
                dispo_p2 = last_match_time.get(id_p2, horaire_actuel)
                
                if dispo_p1 <= horaire_actuel and dispo_p2 <= horaire_actuel:
                    parts = cat.split(" | ")
                    age_cat = parts[0]
                    
                    duree_match = durees_age.get(age_cat, duree_senior)
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
                    
                    last_match_time[id_p1] = fin_match + repos_min_td
                    last_match_time[id_p2] = fin_match + repos_min_td
                    
                    tapis_dispo[t_idx] = fin_match
                    file_attente_globale.pop(i)
                    match_found = True
                    break
                    
            if not match_found:
                next_ready_time = None
                for cat, match in file_attente_globale:
                    id_p1 = match[0]['Nom']
                    id_p2 = match[1]['Nom']
                    ready_at = max(last_match_time.get(id_p1, horaire_actuel), last_match_time.get(id_p2, horaire_actuel))
                    if next_ready_time is None or ready_at < next_ready_time:
                        next_ready_time = ready_at
                
                if horaire_actuel < dt_pause_debut and next_ready_time > dt_pause_debut:
                    next_ready_time = dt_pause_debut
                    
                attente_mins = (next_ready_time - horaire_actuel).seconds // 60
                
                if attente_mins > 0:
                    planning_tapis[t_idx].append({"Type": "ATTENTE", "Heure": horaire_actuel.strftime("%H:%M"), "Texte": f"⏳ Attente ({attente_mins} min)"})
                tapis_dispo[t_idx] = next_ready_time

        fin_estimee_tournoi = max(tapis_dispo) if tapis_dispo else dt_debut
        duree_totale = fin_estimee_tournoi - dt_debut

        # --- FORMATAGE DE L'EXCEL FINAL ---
        resume_data = [
            {"Information": "Nombre total de matchs", "Valeur": str(total_matchs_calcules)},
            {"Information": "Heure de début", "Valeur": dt_debut.strftime('%H:%M')},
            {"Information": "Heure de fin estimée", "Valeur": fin_estimee_tournoi.strftime('%H:%M')},
            {"Information": "Durée totale estimée", "Valeur": format_duree(duree_totale)}
        ]
        df_resume = pd.DataFrame(resume_data)
        
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
                        texte = f"🕘 {m['Heure']} ({m['Duree']} min)\n[{m['Cat']}]\n{m['Combattant 1']} VS {m['Combattant 2']}"
                        ligne_donnees[nom_colonne] = texte
                else:
                    ligne_donnees[nom_colonne] = ""
            grille.append(ligne_donnees)
        df_grille = pd.DataFrame(grille)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_resume.to_excel(writer, sheet_name="Résumé", index=False)
            df_grille.to_excel(writer, sheet_name="Grille de Passage", index=False)
            
            from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
            bleu_fflda = PatternFill("solid", fgColor="0055A4")
            rouge_fflda = PatternFill("solid", fgColor="EF4135")
            gris_clair = PatternFill("solid", fgColor="F2F2F2")
            header_font = Font(bold=True, color="FFFFFF")
            border_style = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            ws_res = writer.sheets["Résumé"]
            for cell in ws_res[1]:
                cell.fill = bleu_fflda
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws_res.column_dimensions['A'].width = 30
            ws_res.column_dimensions['B'].width = 25
            
            ws_grille = writer.sheets["Grille de Passage"]
            for cell in ws_grille[1]:
                cell.fill = bleu_fflda
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border_style
                
            for col in range(1, nb_tapis + 1):
                lettre_col = ws_grille.cell(row=1, column=col).column_letter
                ws_grille.column_dimensions[lettre_col].width = 45
                
            for row in ws_grille.iter_rows(min_row=2, max_row=ws_grille.max_row):
                ws_grille.row_dimensions[row[0].row].height = 90
                for cell in row:
                    cell.border = border_style
                    cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
                    if cell.value and "PAUSE DÉJEUNER" in str(cell.value):
                        cell.fill = rouge_fflda
                        cell.font = Font(bold=True, color="FFFFFF")
                    elif cell.value and "Attente" in str(cell.value):
                        cell.fill = gris_clair
                        cell.font = Font(italic=True, color="666666")

        st.subheader("📊 Statistiques du tournoi")
        col1, col2, col3 = st.columns(3)
        col1.metric("Matchs générés", total_matchs_calcules)
        col2.metric("Heure de fin", fin_estimee_tournoi.strftime('%H:%M'))
        col3.metric("Durée totale", format_duree(duree_totale))

        st.download_button(
            label="📥 Télécharger le Planning Officiel",
            data=output.getvalue(),
            file_name="Planning_Officiel_FFLDA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Une erreur est survenue lors de l'analyse : {e}")

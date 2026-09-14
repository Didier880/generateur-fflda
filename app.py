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

def format_duree(td):
    total_seconds = int(td.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, _ = divmod(remainder, 60)
    if hours > 0: return f"{hours}h {minutes}min"
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
        
        # --- GESTION DU NIVEAU ---
        if "Maîtrise" not in df_inscr.columns: df_inscr["Maîtrise"] = ""
        def attribuer_niveau(val):
            val_str = str(val).strip().lower()
            if val_str == 'd': return 'Débutant'
            elif val_str == 'c': return 'Confirmé'
            return ''
        df_inscr['Niveau'] = df_inscr['Maîtrise'].apply(attribuer_niveau)

        # --- GESTION PESÉE ---
        df_inscr['Poids'] = df_inscr['Poids'].astype(str).str.replace(',', '.')
        df_inscr['Poids_Num'] = pd.to_numeric(df_inscr['Poids'], errors='coerce')
        df_inscr = df_inscr[df_inscr['Poids_Num'] > 0]
        df_inscr['Sexe'] = 'Mixte'
        
        # --- CRÉATION DES POULES ---
        rondes_par_categorie = {}
        participants_par_poule = {} 
        df_morpho_valide = df_inscr.sort_values('Poids_Num')
        
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
                        participants_par_poule[nom_groupe] = list(poule_courante)
                        rondes_par_categorie[nom_groupe] = generer_rondes(poule_courante)
                        index_poule += 1
                        poule_courante = [p]
            if poule_courante:
                nom_groupe = f"{age} | {sexe}{suffixe_niveau} | Gr. {index_poule} ({poule_courante[0]['Poids_Num']}kg - {poule_courante[-1]['Poids_Num']}kg)"
                participants_par_poule[nom_groupe] = list(poule_courante)
                rondes_par_categorie[nom_groupe] = generer_rondes(poule_courante)

        # --- SÉPARATION FILES D'ATTENTE U9 PUIS U11 ---
        file_u9, file_u11 = [], []
        for nom, rondes in rondes_par_categorie.items():
            if nom.startswith("U9"):
                for r in rondes:
                    for m in r: file_u9.append((nom, m))
            else:
                for r in rondes:
                    for m in r: file_u11.append((nom, m))

        # --- GESTION DU TEMPS ---
        dt_debut = datetime.combine(datetime.today(), heure_debut)
        dt_pause_debut = datetime.combine(datetime.today(), pause_debut)
        dt_pause_fin = datetime.combine(datetime.today(), pause_fin)
        
        tapis_dispo = [dt_debut for _ in range(nb_tapis)]
        planning_tapis = {t: [] for t in range(nb_tapis)}
        last_match_time = {} 
        total_matchs_calcules = 0

        phases = [("U9", file_u9), ("U11", file_u11)]

        for nom_phase, file_attente in phases:
            if not file_attente: continue
                
            if total_matchs_calcules > 0:
                heure_synchro = max(tapis_dispo)
                for t in range(nb_tapis):
                    if tapis_dispo[t] < heure_synchro:
                        attente = (heure_synchro - tapis_dispo[t]).seconds // 60
                        if attente > 0:
                            planning_tapis[t].append({"Type": "ATTENTE", "Heure": tapis_dispo[t].strftime("%H:%M"), "Texte": f"Fin des U9 - Attente U11"})
                        tapis_dispo[t] = heure_synchro
            
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
                    if last_match_time.get(p1_nom, horaire_actuel) <= horaire_actuel and last_match_time.get(p2_nom, horaire_actuel) <= horaire_actuel:
                        duree_match = durees_age.get(cat.split(" | ")[0], 5)
                        duree_td = timedelta(minutes=duree_match)
                        
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
                        repos = timedelta(minutes=(repos_matchs * duree_match))
                        last_match_time[p1_nom] = fin_match + repos
                        last_match_time[p2_nom] = fin_match + repos
                        tapis_dispo[t_idx] = fin_match
                        file_attente.pop(i)
                        match_found = True
                        break
                        
                if not match_found:
                    next_ready_time = min(tapis_dispo) + timedelta(minutes=1)
                    tapis_dispo[t_idx] = next_ready_time

        fin_estimee = max(tapis_dispo) if tapis_dispo else dt_debut

        # --- EXPORT EXCEL (AVEC FEUILLES DE MARQUE) ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            
            # Onglet 1 : Résumé
            pd.DataFrame([
                {"Information": "Matchs générés", "Valeur": str(total_matchs_calcules)},
                {"Information": "Heure fin totale", "Valeur": fin_estimee.strftime('%H:%M')}
            ]).to_excel(writer, sheet_name="Résumé", index=False)
            
            # Onglet 2 : Grille de Passage
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
            pd.DataFrame(grille).to_excel(writer, sheet_name="Grille de Passage", index=False)
            
            from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
            b_style = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            bleu = PatternFill("solid", fgColor="0055A4")
            
            # Design Résumé et Grille
            for sheet_name in ["Résumé", "Grille de Passage"]:
                ws = writer.sheets[sheet_name]
                for cell in ws[1]: 
                    cell.fill, cell.font, cell.alignment = bleu, Font(bold=True, color="FFFFFF"), Alignment(horizontal="center")
                    if sheet_name == "Grille de Passage": cell.border = b_style
                if sheet_name == "Grille de Passage":
                    for col in range(1, nb_tapis + 1):
                        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = 45
                    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                        ws.row_dimensions[row[0].row].height = 90
                        for cell in row:
                            cell.border, cell.alignment = b_style, Alignment(wrap_text=True, horizontal="center", vertical="center")
                            if cell.value and "PAUSE" in str(cell.value): cell.fill, cell.font = PatternFill("solid", fgColor="EF4135"), Font(bold=True, color="FFFFFF")
            
            # --- ONGLET 3 : FEUILLES DE POULES (NOUVEAU DESIGN TABLE DE MARQUE) ---
            ws_poules = writer.book.create_sheet("Feuilles de Poules")
            row_cursor = 1
            gris_fonce = PatternFill("solid", fgColor="404040")
            entete_noir = PatternFill("solid", fgColor="000000")
            
            for nom_poule, liste_p in participants_par_poule.items():
                n_lutteurs = len(liste_p)
                
                # Titre de la poule
                cell_titre = ws_poules.cell(row=row_cursor, column=1, value=nom_poule)
                cell_titre.font = Font(bold=True, size=14, color="0055A4")
                row_cursor += 2
                
                # En-têtes du tableau (MISE À JOUR DES COLONNES)
                headers = ["N°", "NOM", "CLUB"] + [str(i) for i in range(1, n_lutteurs+1)] + ["Matchs Gagnés", "Matchs Perdus", "Total Pts", "Classement"]
                for col_idx, h in enumerate(headers, 1):
                    c = ws_poules.cell(row=row_cursor, column=col_idx, value=h)
                    c.font = Font(bold=True)
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    c.border = b_style
                    lettre_col = ws_poules.cell(row=row_cursor, column=col_idx).column_letter
                    if 3 < col_idx <= 3 + n_lutteurs:
                        ws_poules.column_dimensions[lettre_col].width = 5
                        c.fill = entete_noir
                        c.font = Font(bold=True, color="FFFFFF")
                    elif col_idx > 3 + n_lutteurs:
                        # On élargit les nouvelles colonnes de totaux pour laisser la place d'écrire
                        ws_poules.column_dimensions[lettre_col].width = 12
                        
                ws_poules.column_dimensions['A'].width = 5
                ws_poules.column_dimensions['B'].width = 25
                ws_poules.column_dimensions['C'].width = 20
                row_cursor += 1
                
                # Lignes des lutteurs
                for i, p in enumerate(liste_p, 1):
                    ws_poules.row_dimensions[row_cursor].height = 25 # Plus de hauteur pour écrire confortablement
                    ws_poules.cell(row=row_cursor, column=1, value=i).border = b_style
                    ws_poules.cell(row=row_cursor, column=1).alignment = Alignment(horizontal="center", vertical="center")
                    ws_poules.cell(row=row_cursor, column=2, value=p['Nom']).border = b_style
                    ws_poules.cell(row=row_cursor, column=3, value=p.get('Club', '')).border = b_style
                    
                    for j in range(1, n_lutteurs+1):
                        c = ws_poules.cell(row=row_cursor, column=3+j)
                        c.border = b_style
                        if i == j: c.fill = gris_fonce # Diagonale
                            
                    # Création des bordures pour les cases vides à remplir au stylo (Gagnés, Perdus, Pts, Class.)
                    ws_poules.cell(row=row_cursor, column=3+n_lutteurs+1).border = b_style
                    ws_poules.cell(row=row_cursor, column=3+n_lutteurs+2).border = b_style
                    ws_poules.cell(row=row_cursor, column=3+n_lutteurs+3).border = b_style
                    ws_poules.cell(row=row_cursor, column=3+n_lutteurs+4).border = b_style
                    
                    row_cursor += 1
                    
                # Ordre des combats calculé automatiquement
                row_cursor += 1
                rondes = rondes_par_categorie[nom_poule]
                ordre = []
                for r in rondes:
                    for m in r:
                        idx1 = next((i+1 for i, x in enumerate(liste_p) if x['Nom'] == m[0]['Nom']), None)
                        idx2 = next((i+1 for i, x in enumerate(liste_p) if x['Nom'] == m[1]['Nom']), None)
                        if idx1 and idx2: # Ignore les BYE
                            ordre.append(f"{idx1}-{idx2}")
                
                ws_poules.cell(row=row_cursor, column=1, value="Ordre des combats : " + " // ".join(ordre))
                ws_poules.cell(row=row_cursor, column=1).font = Font(italic=True, bold=True)
                
                row_cursor += 4 # Espace avant la poule suivante

        st.subheader("📊 Statistiques")
        st.metric("Matchs générés", total_matchs_calcules)

        st.download_button(label="📥 Télécharger le Planning & Poules", data=output.getvalue(), file_name="Tournoi_U9_U11.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"Une erreur est survenue : {e}")

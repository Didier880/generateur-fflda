import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Générateur Officiel FFLDA", page_icon="🤼", layout="centered")

# --- PERSONNALISATION VISUELLE FFLDA (CSS) ---
st.markdown("""

""", unsafe_allow_html=True)

# --- EN-TÊTE AVEC LOGO ---
col1, col2 = st.columns([1, 4])
with col1:
    # URL du logo de la FFLDA
    st.image("https://upload.wikimedia.org/wikipedia/fr/thumb/5/58/Logo_F%C3%A9d%C3%A9ration_Fran%C3%A7aise_de_Lutte.svg/1200px-Logo_F%C3%A9d%C3%A9ration_Fran%C3%A7aise_de_Lutte.svg.png", width=120)
with col2:
    st.title("Générateur de Tournoi FFLDA")
    st.markdown("**Outil officiel d'optimisation et d'appariement des poules**")

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

fichier_upload = st.file_uploader("📂 Importez votre fichier Excel d'inscriptions", type=["xlsx"])

if fichier_upload is not None:
    st.success("Fichier chargé avec succès ! Analyse en cours...")
    
    try:
        df_param = pd.read_excel(fichier_upload, sheet_name="Paramètres")
        params = dict(zip(df_param['Paramètre'], df_param['Valeur']))
        
        nb_tapis = int(params.get("Nombre de tapis", 3))
        heure_debut = str(params.get("Heure de début", "09:00")).zfill(5)
        pause_debut = str(params.get("Début de la pause", "12:00")).zfill(5)
        pause_fin = str(params.get("Fin de la pause", "13:00")).zfill(5)
        
        durees = {
            "Débutant": int(params.get("Temps - Débutant", 5)),
            "Intermédiaire": int(params.get("Temps - Intermédiaire", 6)),
            "Confirmé": int(params.get("Temps - Confirmé", 7))
        }
        
        df_inscr = pd.read_excel(fichier_upload, sheet_name="Inscriptions")
        
        colonnes_requises = ['Nom', 'Niveau', 'Age', 'Sexe', 'Poids']
        for col in colonnes_requises:
            if col not in df_inscr.columns:
                st.error(f"Erreur : La colonne '{col}' manque dans l'onglet Inscriptions.")
                st.stop()
                
        df_inscr = df_inscr.dropna(subset=['Nom', 'Niveau', 'Age', 'Poids'])
        
        df_inscr['Poids'] = df_inscr['Poids'].astype(str).str.replace(',', '.')
        
        mask_mixte = df_inscr['Age'].isin(['U9', 'U11'])
        if mask_mixte.any():
            df_inscr.loc[mask_mixte, 'Sexe'] = 'Mixte'
            st.info("ℹ️ Règle FFLDA appliquée : Les U9 et U11 ont été passés en 'Mixte'.")
        df_inscr = df_inscr.dropna(subset=['Sexe'])

        rondes_par_categorie = {}
        groupes_scindes = 0
        
        mask_morpho = df_inscr['Age'].isin(['U9', 'U11'])
        df_morpho = df_inscr[mask_morpho].copy()
        df_autres = df_inscr[~mask_morpho].copy()

        if not df_autres.empty:
            df_autres['Groupe_Complet'] = df_autres['Niveau'].astype(str) + " | " + df_autres['Age'].astype(str) + " | " + df_autres['Sexe'].astype(str) + " | " + df_autres['Poids'].astype(str)
            for categorie, groupe in df_autres.groupby('Groupe_Complet'):
                participants = groupe.to_dict('records')
                if len(participants) > 6:
                    poule_a = participants[::2]
                    poule_b = participants[1::2]
                    cat_a = categorie + " | Poule A"
                    cat_b = categorie + " | Poule B"
                    rondes_par_categorie[cat_a] = generer_rondes(poule_a)
                    rondes_par_categorie[cat_b] = generer_rondes(poule_b)
                    groupes_scindes += 1
                else:
                    rondes_par_categorie[categorie] = generer_rondes(participants)

        if not df_morpho.empty:
            df_morpho['Poids_Num'] = pd.to_numeric(df_morpho['Poids'], errors='coerce')
            erreurs_saisie = df_morpho[df_morpho['Poids_Num'].isna()]
            if not erreurs_saisie.empty:
                st.error("⚠️ Certains lutteurs U9/U11 n'ont pas un poids valide. Ils ont été ignorés.")
            
            df_morpho_valide = df_morpho.dropna(subset=['Poids_Num']).sort_values('Poids_Num')
            for (niveau, age, sexe), groupe in df_morpho_valide.groupby(['Niveau', 'Age', 'Sexe']):
                participants = groupe.to_dict('records')
                poule_courante = []
                index_poule = 1
                for p in participants:
                    if not poule_courante:
                        poule_courante.append(p)
                    else:
                        poids_min = poule_courante[0]['Poids_Num']
                        poids_actuel = p['Poids_Num']
                        if poids_actuel <= (poids_min * 1.10) and len(poule_courante) < 6:
                            poule_courante.append(p)
                        else:
                            nom_groupe = f"{niveau} | {age} | {sexe} | Gr. {index_poule} ({poule_courante[0]['Poids_Num']}kg - {poule_courante[-1]['Poids_Num']}kg)"
                            rondes_par_categorie[nom_groupe] = generer_rondes(poule_courante)
                            index_poule += 1
                            poule_courante = [p]
                if poule_courante:
                    nom_groupe = f"{niveau} | {age} | {sexe} | Gr. {index_poule} ({poule_courante[0]['Poids_Num']}kg - {poule_courante[-1]['Poids_Num']}kg)"
                    rondes_par_categorie[nom_groupe] = generer_rondes(poule_courante)


        if groupes_scindes > 0:
            st.warning(f"⚠️ {groupes_scindes} catégorie(s) classique(s) dépassant 6 lutteurs ont été automatiquement scindées.")

        max_rondes = max((len(rondes) for rondes in rondes_par_categorie.values()), default=0)
        file_attente_globale = []
        for r in range(max_rondes):
            for cat in rondes_par_categorie.keys():
                if r < len(rondes_par_categorie[cat]):
                    for match in rondes_par_categorie[cat][r]:
                        file_attente_globale.append((cat, match))

        dt_debut = datetime.strptime(heure_debut, "%H:%M")
        dt_pause_debut = datetime.strptime(pause_debut, "%H:%M")
        dt_pause_fin = datetime.strptime(pause_fin, "%H:%M")
        repos_min = timedelta(minutes=20)
        
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
                p1_nom, p1_club = match[0]["Nom"], match[0].get("Club", "")
                p2_nom, p2_club = match[1]["Nom"], match[1].get("Club", "")
                id_p1 = f"{p1_nom}_{p1_club}"
                id_p2 = f"{p2_nom}_{p2_club}"
                
                dispo_p1 = last_match_time.get(id_p1, horaire_actuel)
                dispo_p2 = last_match_time.get(id_p2, horaire_actuel)
                
                if dispo_p1 <= horaire_actuel and dispo_p2 <= horaire_actuel:
                    parts = cat.split(" | ")
                    niveau = parts[0]
                    duree = durees.get(niveau, 7)
                    duree_td = timedelta(minutes=duree)
                    
                    if horaire_actuel < dt_pause_debut and (horaire_actuel + duree_td) > dt_pause_debut:
                        planning_tapis[t_idx].append({"Type": "PAUSE", "Heure": horaire_actuel.strftime("%H:%M")})
                        tapis_dispo[t_idx] = dt_pause_fin
                        match_found = True
                        break 
                    
                    poids_texte = parts[3]
                    if len(parts) > 4:
                        poids_texte += f" ({parts[4]})"
                        
                    planning_tapis[t_idx].append({
                        "Type": "MATCH",
                        "Heure": horaire_actuel.strftime("%H:%M"),
                        "Duree": duree,
                        "Niveau": parts[0],
                        "Age": parts[1],
                        "Sexe": parts[2],
                        "Poids": poids_texte,
                        "Combattant 1": p1_nom,
                        "Club 1": p1_club,
                        "Combattant 2": p2_nom,
                        "Club 2": p2_club
                    })
                    
                    total_matchs_calcules += 1
                    fin_match = horaire_actuel + duree_td
                    
                    last_match_time[id_p1] = fin_match + repos_min
                    last_match_time[id_p2] = fin_match + repos_min
                    
                    tapis_dispo[t_idx] = fin_match
                    file_attente_globale.pop(i)
                    match_found = True
                    break
                    
            if not match_found:
                next_ready_time = None
                for cat, match in file_attente_globale:
                    id_p1 = f"{match[0]['Nom']}_{match[0].get('Club', '')}"
                    id_p2 = f"{match[1]['Nom']}_{match[1].get('Club', '')}"
                    dispo_p1 = last_match_time.get(id_p1, horaire_actuel)
                    dispo_p2 = last_match_time.get(id_p2, horaire_actuel)
                    ready_at = max(dispo_p1, dispo_p2)
                    if next_ready_time is None or ready_at < next_ready_time:
                        next_ready_time = ready_at
                
                if horaire_actuel < dt_pause_debut and next_ready_time > dt_pause_debut:
                    next_ready_time = dt_pause_debut
                    
                attente_mins = (next_ready_time - horaire_actuel).seconds // 60
                
                if attente_mins > 0:
                    planning_tapis[t_idx].append({
                        "Type": "ATTENTE", 
                        "Heure": horaire_actuel.strftime("%H:%M"),
                        "Texte": f"⏳ Attente ({attente_mins} min)"
                    })
                tapis_dispo[t_idx] = next_ready_time

        fin_estimee_tournoi = max(tapis_dispo) if tapis_dispo else dt_debut
        duree_totale = fin_estimee_tournoi - dt_debut

        resume_data = [
            {"Information": "Nombre total de matchs", "Valeur": str(total_matchs_calcules)},
            {"Information": "Heure de début", "Valeur": dt_debut.strftime('%H:%M')},
            {"Information": "Heure de fin estimée", "Valeur": fin_estimee_tournoi.strftime('%H:%M')},
            {"Information": "Durée totale de l'événement", "Valeur": format_duree(duree_totale)}
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
                        club1 = f" ({m['Club 1']})" if m['Club 1'] and m['Club 1'] != "-" else ""
                        club2 = f" ({m['Club 2']})" if m['Club 2'] and m['Club 2'] != "-" else ""
                        texte = f"🕘 {m['Heure']} ({m['Duree']} min)\n[{m['Niveau']} | {m['Age']} | {m['Sexe']} | {m['Poids']}]\n{m['Combattant 1']}{club1}\nVS\n{m['Combattant 2']}{club2}"
                        ligne_donnees[nom_colonne] = texte
                else:
                    ligne_donnees[nom_colonne] = ""
            grille.append(ligne_donnees)
        df_grille = pd.DataFrame(grille)

        # --- CRÉATION DE L'EXCEL HABILLÉ FFLDA ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_inscr.to_excel(writer, sheet_name="Inscriptions", index=False)
            df_param.to_excel(writer, sheet_name="Paramètres", index=False)
            df_resume.to_excel(writer, sheet_name="Résumé", index=False)
            df_grille.to_excel(writer, sheet_name="Grille de Passage", index=False)
            
            from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
            
            # COULEURS FFLDA
            bleu_fflda = PatternFill("solid", fgColor="0055A4") # Bleu officiel
            rouge_fflda = PatternFill("solid", fgColor="EF4135") # Rouge officiel
            gris_clair = PatternFill("solid", fgColor="F2F2F2")
            
            header_font = Font(bold=True, color="FFFFFF")
            border_style = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            # Onglet Résumé
            ws_res = writer.sheets["Résumé"]
            for cell in ws_res[1]:
                cell.fill = bleu_fflda
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws_res.column_dimensions['A'].width = 30
            ws_res.column_dimensions['B'].width = 25
            
            # Onglet Grille de Passage
            ws_grille = writer.sheets["Grille de Passage"]
            for cell in ws_grille[1]:
                cell.fill = bleu_fflda # En-tête des tapis en Bleu FFLDA
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border_style
                
            for col in range(1, nb_tapis + 1):
                lettre_col = ws_grille.cell(row=1, column=col).column_letter
                ws_grille.column_dimensions[lettre_col].width = 45
                
            for row in ws_grille.iter_rows(min_row=2, max_row=ws_grille.max_row):
                ws_grille.row_dimensions[row[0].row].height = 110
                for cell in row:
                    cell.border = border_style
                    cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
                    
                    if cell.value and "PAUSE DÉJEUNER" in str(cell.value):
                        cell.fill = rouge_fflda # Pause en Rouge FFLDA
                        cell.font = Font(bold=True, color="FFFFFF")
                    elif cell.value and "Attente" in str(cell.value):
                        cell.fill = gris_clair
                        cell.font = Font(italic=True, color="666666")

        st.subheader("📊 Statistiques du tournoi")
        st.write(f"**Matchs générés :** {total_matchs_calcules}")
        st.write(f"**Fin estimée :** {fin_estimee_tournoi.strftime('%H:%M')}")
        st.write(f"**Durée de l'événement :** {format_duree(duree_totale)}")

        st.download_button(
            label="📥 Télécharger le Planning Officiel",
            data=output.getvalue(),
            file_name="Planning_Officiel_FFLDA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Une erreur est survenue : {e}")

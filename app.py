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
                {"Étape de la journée": "Nombre total de participants (pesés)", "Horaire / Valeur": str(total_participants_peses)},
                {"Étape de la journée": "Athlètes non pesés / absents", "Horaire / Valeur": str(total_non_peses)},
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
            
            # Dictionnaire pour stocker les coordonnées des cellules de score dans les feuilles de poules (pour le récapitulatif global)
            # Structure : { "Nom du lutteur": {"feuille": ws_poule, "cell_pts": coord, "cell_vict": coord} }
            suivi_classement_global = []

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
                    
                    # Enregistrer pour l'onglet de classement global
                    suivi_classement_global.append({
                        "Nom": p['Nom'],
                        "Club": p.get('Club', ''),
                        "Categorie": nom_poule,
                        "cell_pts": f"'{nom_onglet_court}'!{cell_total_pts.coordinate}",
                        "cell_poids": p.get('Poids', '')
                    })

                    row_cursor += 1
                
                row_cursor += 2
                
                rondes = rondes_par_categorie[nom_poule]
                col_offset_tours = 5 
                
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

            # --- CRÉATION DE L'ONGLET CLASSEMENT INDIVIDUEL ---
            ws_classement = writer.book.create_sheet(title="Classement Individuel", index=2) # Positionné en 3ème position
            ws_classement.cell(row=1, column=1, value="🏆 CLASSEMENT GÉNÉRAL INDIVIDUEL 🏆").font = Font(name="Arial", size=16, bold=True, color="0055A4")
            
            headers_clt = ["Rang", "Nom Prénom", "Club", "Catégorie / Poule", "Points Totaux", "Poids"]
            for col_idx, h in enumerate(headers_clt, 1):
                c = ws_classement.cell(row=3, column=col_idx, value=h)
                c.font, c.alignment, c.border = Font(bold=True, color="FFFFFF"), Alignment(horizontal="center", vertical="center"), b_style
                c.fill = entete_noir

            ws_classement.column_dimensions['A'].width = 8
            ws_classement.column_dimensions['B'].width = 30
            ws_classement.column_dimensions['C'].width = 20
            ws_classement.column_dimensions['D'].width = 35
            ws_classement.column_dimensions['E'].width = 15
            ws_classement.column_dimensions['F'].width = 10

            row_clt = 4
            for item in suivi_classement_global:
                ws_classement.cell(row=row_clt, column=1, value=f"") # Optionnel pour le rang dynamique ou manuel
                ws_classement.cell(row=row_clt, column=2, value=item["Nom"]).border = b_style
                ws_classement.cell(row=row_clt, column=3, value=item["Club"]).border = b_style
                ws_classement.cell(row=row_clt, column=4, value=item["Categorie"]).border = b_style
                
                cell_pts_ref = ws_classement.cell(row=row_clt, column=5, value=f"={item['cell_pts']}")
                cell_pts_ref.border = b_style
                cell_pts_ref.alignment = Alignment(horizontal="center", vertical="center")
                cell_pts_ref.font = Font(bold=True)

                cell_pds = ws_classement.cell(row=row_clt, column=6, value=item["cell_poids"])
                cell_pds.border = b_style
                cell_pds.alignment = Alignment(horizontal="center", vertical="center")

                ws_classement.row_dimensions[row_clt].height = 20
                row_clt += 1

            st.download_button(label="📥 Télécharger le Planning & Feuilles de Poules (Excel)", data=output.getvalue(), file_name="Tournoi_U9_U11.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"Une erreur est survenue : {e}")

"""
Module de génération de rapports Word
"""

from datetime import datetime
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, RGBColor


def generer_rapport_word(
    fichier_word,
    df_tableau,
    fichier_graphique,
    depenses_cumulees,
    pv_cumulee,
    ev_cumulee,
    projections,
    fichier_projections,
):
    """
    Génère un rapport Word complet avec définitions, tableau, deux graphiques (réalisé + projections) et conclusion

    Args:
        fichier_word: Nom du fichier Word à générer
        df_tableau: DataFrame avec le tableau comparatif
        fichier_graphique: Chemin du graphique du réalisé
        depenses_cumulees: Série des dépenses réelles (AC)
        pv_cumulee: Série du budget prévu (PV)
        ev_cumulee: Série de la valeur acquise (EV)
        projections: Dictionnaire des projections {'cpi': {...}, 'cpi_spi': {...}, 'reste_plan': {...}, 'forecast': {...}}
        fichier_projections: Chemin du graphique des projections
    """
    print(f"\n=== Génération du rapport Word: {fichier_word} ===")

    doc = Document()

    # Titre principal
    titre = doc.add_heading("Rapport d'Analyse EVM", 0)
    titre.alignment = WD_ALIGN_PARAGRAPH.CENTER

    date_rapport = datetime.now().astimezone().strftime("%d/%m/%Y")
    p = doc.add_paragraph(f"Date du rapport: {date_rapport}")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_page_break()

    # Section 1: Définitions
    doc.add_heading("1. Définitions EVM", 1)

    definitions = [
        (
            "PV (Planned Value)",
            "Budget prévu ou valeur planifiée. Représente le coût budgété du travail prévu à une date donnée.",
        ),
        (
            "AC (Actual Cost)",
            "Coût réel ou dépenses réelles. Représente le coût réel du travail effectué à une date donnée.",
        ),
        (
            "EV (Earned Value)",
            "Valeur acquise ou valeur gagnée. Représente la valeur du travail réellement accompli à une date donnée, mesurée en termes de budget.",
        ),
        (
            "EAC (Estimate at Completion)",
            "Estimation à terminaison. Projection du coût total du projet à son achèvement.",
        ),
        ("CV (Cost Variance)", "Écart de coût. CV = EV - AC. Un CV négatif indique un dépassement de coût."),
        ("SV (Schedule Variance)", "Écart de délai. SV = EV - PV. Un SV négatif indique un retard sur le planning."),
        (
            "CPI (Cost Performance Index)",
            "Indice de performance des coûts. CPI = EV / AC. Un CPI < 1 indique un dépassement de coût.",
        ),
        (
            "SPI (Schedule Performance Index)",
            "Indice de performance des délais. SPI = EV / PV. Un SPI < 1 indique un retard.",
        ),
    ]

    for terme, definition in definitions:
        p = doc.add_paragraph(style="List Bullet")
        run_terme = p.add_run(terme + ": ")
        run_terme.bold = True
        p.add_run(definition)

    doc.add_page_break()

    # Section 2: Réalisé à date
    doc.add_heading("2. Réalisé à Date", 1)

    # 2.1 Tableau de valeurs
    doc.add_heading("2.1 Tableau des Valeurs", 2)

    # Filtrer les lignes avec des valeurs non nulles pour le tableau Word
    # Vérifier quelles colonnes sont présentes avant le filtrage
    conditions = [df_tableau["AC (Dépenses réelles)"] > 0]
    if "PV (Budget prévu)" in df_tableau.columns:
        conditions.append(df_tableau["PV (Budget prévu)"] > 0)
    if "EV (Valeur acquise)" in df_tableau.columns:
        conditions.append(df_tableau["EV (Valeur acquise)"] > 0)

    # Combiner les conditions avec OR
    combined_condition = conditions[0]
    for condition in conditions[1:]:
        combined_condition = combined_condition | condition

    df_filtre = df_tableau[combined_condition].copy()

    # Créer le tableau Word
    if len(df_filtre) > 0:
        table = doc.add_table(rows=1, cols=len(df_filtre.columns))
        table.style = "Light Grid Accent 1"

        # En-têtes
        hdr_cells = table.rows[0].cells
        for i, col in enumerate(df_filtre.columns):
            hdr_cells[i].text = col
            for paragraph in hdr_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True

        # Données
        for _, row in df_filtre.iterrows():
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                if isinstance(value, (int, float)):
                    row_cells[i].text = f"{value:,.2f}"
                else:
                    row_cells[i].text = str(value)

    # 2.2 Graphique du réalisé
    doc.add_heading("2.2 Graphique du Réalisé", 2)

    if Path(fichier_graphique).exists():
        doc.add_picture(fichier_graphique, width=Inches(6.5))
        p = doc.add_paragraph("Figure 1: Courbes du réalisé - AC, PV, EV et variances")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].italic = True

    # 2.3 Indicateurs actuels
    doc.add_heading("2.3 Indicateurs de Performance Actuels", 2)

    # Récupérer les dernières valeurs
    dernier_mois = depenses_cumulees.index[-1]
    ac_actuel = depenses_cumulees.iloc[-1]

    ev_actuel = 0
    if ev_cumulee is not None and len(ev_cumulee) > 0:
        ev_actuel = ev_cumulee.iloc[-1]

    pv_actuel = 0
    if pv_cumulee is not None and dernier_mois in pv_cumulee.index:
        pv_actuel = pv_cumulee[dernier_mois]

    # Calculer CV et SV actuels
    cv_actuel = ev_actuel - ac_actuel
    sv_actuel = ev_actuel - pv_actuel

    # Calculer CPI et SPI
    cpi = ev_actuel / ac_actuel if ac_actuel > 0 else 0
    spi = ev_actuel / pv_actuel if pv_actuel > 0 else 0

    p = doc.add_paragraph(f"Au mois de {dernier_mois}:")

    indicateurs_actuels = [
        ("Dépenses Réelles (AC)", f"{ac_actuel:,.2f} €"),
        ("Valeur Acquise (EV)", f"{ev_actuel:,.2f} €"),
        ("Valeur Planifiée (PV)", f"{pv_actuel:,.2f} €"),
        ("", ""),
        ("Cost Variance (CV)", f"{cv_actuel:,.2f} €", cv_actuel),
        ("Schedule Variance (SV)", f"{sv_actuel:,.2f} €", sv_actuel),
        ("Cost Performance Index (CPI)", f"{cpi:.2f}", cpi - 1),
        ("Schedule Performance Index (SPI)", f"{spi:.2f}", spi - 1),
    ]

    for item in indicateurs_actuels:
        if len(item) >= 2:
            p = doc.add_paragraph(style="List Bullet")
            run = p.add_run(f"{item[0]}: {item[1]}")
            if len(item) > 2:
                valeur = item[2]
                if valeur < 0:
                    run.font.color.rgb = RGBColor(231, 76, 60)  # Rouge
                elif valeur > 0:
                    run.font.color.rgb = RGBColor(46, 204, 113)  # Vert

    # Interprétation
    doc.add_paragraph()
    p = doc.add_paragraph("Interprétation:")
    p.runs[0].bold = True

    if cv_actuel < 0:
        doc.add_paragraph(
            f"⚠ Le projet présente un dépassement de coût de {abs(cv_actuel):,.2f} € à date.", style="List Bullet"
        )
    else:
        doc.add_paragraph(
            f"✓ Le projet est sous budget avec une économie de {cv_actuel:,.2f} € à date.", style="List Bullet"
        )

    if sv_actuel < 0:
        doc.add_paragraph(
            f"⚠ Le projet présente un retard équivalent à {abs(sv_actuel):,.2f} € de travail non réalisé.",
            style="List Bullet",
        )
    else:
        doc.add_paragraph(
            f"✓ Le projet est en avance avec {sv_actuel:,.2f} € de travail supplémentaire réalisé.", style="List Bullet"
        )

    if cpi < 1:
        doc.add_paragraph(
            f"⚠ L'efficacité des coûts est de {cpi * 100:.1f}% (chaque euro dépensé génère {cpi:.2f} € de valeur).",
            style="List Bullet",
        )
    else:
        doc.add_paragraph(
            f"✓ L'efficacité des coûts est de {cpi * 100:.1f}% (chaque euro dépensé génère {cpi:.2f} € de valeur).",
            style="List Bullet",
        )

    doc.add_page_break()

    # Section 3: Projections à Terminaison
    doc.add_heading("3. Projections à Terminaison", 1)

    # 3.1 Tableau des scénarios
    doc.add_heading("3.1 Tableau Comparatif des Scénarios", 2)

    budget_total = 0
    if pv_cumulee is not None:
        budget_total = pv_cumulee.iloc[-1]

    scenarios_data = []

    labels_scenarios = {
        "reste_plan": "Méthode Reste à Plan (Optimiste)",
        "cpi": "Méthode CPI (Réaliste)",
        "cpi_spi": "Méthode CPI×SPI (Pessimiste)",
        "forecast": "Forecast Manuel",
    }

    for methode, label in labels_scenarios.items():
        if methode in projections and projections[methode] is not None:
            proj = projections[methode]
            eac = proj["eac"]
            date_fin = proj["date"].strftime("%m/%Y") if proj["date"] else "N/A"
            vac = budget_total - eac
            scenarios_data.append(
                {
                    "Scénario": label,
                    "EAC (€)": eac,
                    "Date Fin": date_fin,
                    "VAC (€)": vac,
                    "Dépassement": "Oui" if vac < 0 else "Non",
                }
            )

    if scenarios_data:
        df_scenarios = pd.DataFrame(scenarios_data)

        # Créer le tableau dans Word
        table = doc.add_table(rows=1, cols=len(df_scenarios.columns))
        table.style = "Light Grid Accent 1"

        # En-têtes
        hdr_cells = table.rows[0].cells
        for i, col in enumerate(df_scenarios.columns):
            hdr_cells[i].text = col
            for paragraph in hdr_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True

        # Données
        for _, row in df_scenarios.iterrows():
            row_cells = table.add_row().cells
            for i, (col, value) in enumerate(row.items()):
                if col in ["EAC (€)", "VAC (€)"]:
                    row_cells[i].text = f"{value:,.2f}"
                    # Colorer en rouge si VAC négatif
                    if col == "VAC (€)" and value < 0:
                        for paragraph in row_cells[i].paragraphs:
                            for run in paragraph.runs:
                                run.font.color.rgb = RGBColor(231, 76, 60)
                else:
                    row_cells[i].text = str(value)

    # 3.2 Graphique des projections
    doc.add_heading("3.2 Graphique des Projections", 2)

    if Path(fichier_projections).exists():
        doc.add_picture(fichier_projections, width=Inches(6.5))
        p = doc.add_paragraph("Figure 2: Projections à terminaison - Différents scénarios EAC")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].italic = True

    # 3.3 Analyse des scénarios
    doc.add_heading("3.3 Analyse des Scénarios", 2)

    if scenarios_data:
        # Créer un dictionnaire des EAC par scénario avec leur label complet
        scenarios_dict = {s["Scénario"]: s["EAC (€)"] for s in scenarios_data}

        # Identifier les scénarios disponibles
        eac_optimiste = None
        eac_realiste = None
        eac_pessimiste = None
        label_optimiste = ""
        label_realiste = ""
        label_pessimiste = ""

        # Chercher chaque type de scénario
        for scenario_label, eac_val in scenarios_dict.items():
            if "Optimiste" in scenario_label:
                eac_optimiste = eac_val
                label_optimiste = scenario_label
            elif "Réaliste" in scenario_label:
                eac_realiste = eac_val
                label_realiste = scenario_label
            elif "Pessimiste" in scenario_label:
                eac_pessimiste = eac_val
                label_pessimiste = scenario_label

        # Si on a un forecast manuel, on l'analyse aussi
        eac_forecast = None
        label_forecast = ""
        for scenario_label, eac_val in scenarios_dict.items():
            if "Forecast" in scenario_label or "Manuel" in scenario_label:
                eac_forecast = eac_val
                label_forecast = scenario_label

        doc.add_paragraph(f"Budget Total (BAC): {budget_total:,.2f} €")
        doc.add_paragraph()
        doc.add_paragraph("Fourchette des projections:")

        # Afficher les scénarios disponibles
        if eac_optimiste is not None:
            doc.add_paragraph(f"  • {label_optimiste}: {eac_optimiste:,.2f} €", style="List Bullet")
        if eac_realiste is not None:
            doc.add_paragraph(f"  • {label_realiste}: {eac_realiste:,.2f} €", style="List Bullet")
        if eac_pessimiste is not None:
            doc.add_paragraph(f"  • {label_pessimiste}: {eac_pessimiste:,.2f} €", style="List Bullet")
        if eac_forecast is not None:
            doc.add_paragraph(f"  • {label_forecast}: {eac_forecast:,.2f} €", style="List Bullet")

        doc.add_paragraph()

        p = doc.add_paragraph("Écarts par rapport au budget:")
        p.runs[0].bold = True

        # Calculer les écarts pour chaque scénario disponible
        scenarios_ecarts = []
        if eac_optimiste is not None:
            scenarios_ecarts.append(("Optimiste", eac_optimiste - budget_total, eac_optimiste))
        if eac_realiste is not None:
            scenarios_ecarts.append(("Réaliste", eac_realiste - budget_total, eac_realiste))
        if eac_pessimiste is not None:
            scenarios_ecarts.append(("Pessimiste", eac_pessimiste - budget_total, eac_pessimiste))
        if eac_forecast is not None:
            scenarios_ecarts.append(("Forecast", eac_forecast - budget_total, eac_forecast))

        for label, ecart, _eac_val in scenarios_ecarts:
            p = doc.add_paragraph(style="List Bullet")
            signe = "+" if ecart >= 0 else ""
            run = p.add_run(f"  • {label}: {signe}{ecart:,.2f} € ({(ecart / budget_total) * 100:+.1f}%)")
            if ecart < 0:
                run.font.color.rgb = RGBColor(231, 76, 60)  # Rouge
            else:
                run.font.color.rgb = RGBColor(46, 204, 113)  # Vert

    doc.add_page_break()

    # Section 4: Conclusion et Recommandations
    doc.add_heading("4. Conclusion et Recommandations", 1)

    doc.add_heading("4.1 Synthèse", 2)

    # Synthèse de la situation actuelle
    doc.add_paragraph("Performance actuelle:")
    if cpi < 0.9:
        doc.add_paragraph(
            "⚠ Le CPI est très faible, indiquant une efficacité des coûts préoccupante. Actions correctives urgentes recommandées.",
            style="List Bullet",
        )
    elif cpi < 1:
        doc.add_paragraph(
            "⚠ Le CPI est inférieur à 1, indiquant un dépassement de coût. Une surveillance étroite est nécessaire.",
            style="List Bullet",
        )
    else:
        doc.add_paragraph("✓ Le CPI est supérieur à 1, indiquant une bonne efficacité des coûts.", style="List Bullet")

    if spi < 0.9:
        doc.add_paragraph(
            "⚠ Le SPI est très faible, indiquant un retard significatif. Révision du planning recommandée.",
            style="List Bullet",
        )
    elif spi < 1:
        doc.add_paragraph(
            "⚠ Le SPI est inférieur à 1, indiquant un retard. Des mesures d'accélération devraient être envisagées.",
            style="List Bullet",
        )
    else:
        doc.add_paragraph(
            "✓ Le SPI est supérieur à 1, indiquant une bonne performance sur les délais.", style="List Bullet"
        )

    doc.add_paragraph()

    # Projections - Calculer min/max/tous les EAC disponibles
    if scenarios_data:
        eacs_disponibles = [
            eac_val for eac_val in [eac_optimiste, eac_realiste, eac_pessimiste, eac_forecast] if eac_val is not None
        ]

        if eacs_disponibles:
            doc.add_paragraph("Projections à terminaison:")

            if all(eac > budget_total for eac in eacs_disponibles):
                doc.add_paragraph(
                    "⚠ Tous les scénarios prévoient un dépassement de budget. Des mesures correctives sont nécessaires.",
                    style="List Bullet",
                )
            elif any(eac > budget_total for eac in eacs_disponibles):
                doc.add_paragraph(
                    "⚠ Certains scénarios prévoient un dépassement de budget. Une vigilance accrue est requise.",
                    style="List Bullet",
                )
            else:
                doc.add_paragraph(
                    "✓ Les projections indiquent un achèvement sous budget dans tous les scénarios.",
                    style="List Bullet",
                )

            # Calculer l'écart entre le meilleur et le pire scénario
            eac_min_calc = min(eacs_disponibles)
            eac_max_calc = max(eacs_disponibles)
            ecart_relatif = ((eac_max_calc - eac_min_calc) / budget_total) * 100
            if ecart_relatif > 10:
                doc.add_paragraph(
                    f"⚠ L'écart entre scénarios est important ({ecart_relatif:.1f}% du budget), reflétant une forte incertitude.",
                    style="List Bullet",
                )
            elif ecart_relatif > 5:
                doc.add_paragraph(
                    f"ℹ L'écart entre scénarios est modéré ({ecart_relatif:.1f}% du budget).", style="List Bullet"
                )
            else:
                doc.add_paragraph(
                    f"✓ L'écart entre scénarios est faible ({ecart_relatif:.1f}% du budget), indiquant une bonne prévisibilité.",
                    style="List Bullet",
                )

    doc.add_heading("4.2 Recommandations", 2)

    if cpi < 1 or spi < 1 or (scenarios_data and any(s["VAC (€)"] < 0 for s in scenarios_data)):
        doc.add_paragraph("Actions recommandées:")

        if cpi < 1:
            doc.add_paragraph(
                "1. Analyser les causes du dépassement de coût et identifier les postes problématiques",
                style="List Number",
            )
            doc.add_paragraph(
                "2. Mettre en place des mesures de réduction des coûts ou réviser le scope", style="List Number"
            )

        if spi < 1:
            doc.add_paragraph(
                "3. Revoir la planification et identifier les leviers d'accélération", style="List Number"
            )
            doc.add_paragraph("4. Augmenter les ressources si nécessaire pour rattraper le retard", style="List Number")

        if scenarios_data and any(s["VAC (€)"] < 0 for s in scenarios_data):
            doc.add_paragraph(
                "5. Prévoir un budget de contingence pour couvrir le dépassement projeté", style="List Number"
            )
            doc.add_paragraph(
                "6. Communiquer proactivement avec les parties prenantes sur les risques financiers",
                style="List Number",
            )
    else:
        doc.add_paragraph("Le projet montre de bonnes performances. Recommandations:", style="List Bullet")
        doc.add_paragraph("  • Maintenir les pratiques actuelles de gestion", style="List Bullet")
        doc.add_paragraph("  • Continuer la surveillance régulière des indicateurs", style="List Bullet")
        doc.add_paragraph("  • Capitaliser sur les bonnes pratiques pour les projets futurs", style="List Bullet")

    # Sauvegarder le document
    doc.save(fichier_word)
    print(f"✓ Rapport Word généré: {fichier_word}")

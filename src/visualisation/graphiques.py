"""
Module de génération de graphiques EVM
"""

import matplotlib.pyplot as plt


def tracer_courbe(
    depenses_cumulees,
    pv_cumulee=None,
    jalons=None,
    ev_cumulee=None,
    eac_projete=None,
    jalons_forecast=None,
    fichier_sortie="analyse_evm.png",
):
    """
    Trace la courbe des dépenses cumulées, de la Planned Value et de l'Earned Value
    """
    plt.figure(figsize=(16, 9))

    # Conversion en k€ pour une meilleure lisibilité
    depenses_ke = depenses_cumulees / 1000

    # Conversion des périodes en dates pour l'affichage (dépenses réelles)
    dates_depenses = [periode.to_timestamp() for periode in depenses_cumulees.index]

    # Tracé des dépenses réelles (AC - Actual Cost)
    plt.plot(
        dates_depenses,
        depenses_ke.values,
        marker="o",
        linewidth=2.5,
        markersize=8,
        label="AC (Actual Cost - Dépenses réelles)",
        color="#e74c3c",
        zorder=3,
    )

    # Tracé de la Planned Value si disponible
    if pv_cumulee is not None:
        pv_ke = pv_cumulee / 1000
        dates_pv = [periode.to_timestamp() for periode in pv_cumulee.index]
        plt.plot(
            dates_pv,
            pv_ke.values,
            marker="s",
            linewidth=2.5,
            markersize=8,
            label="PV (Planned Value - Budget prévu)",
            color="#3498db",
            linestyle="--",
            zorder=2,
        )

        # Ajout des jalons sur la courbe PV
        if jalons:
            for periode, valeur in zip(pv_cumulee.index, pv_ke.values):
                if periode in jalons:
                    date = periode.to_timestamp()
                    jalons_str = "\n".join(jalons[periode])
                    plt.annotate(
                        jalons_str,
                        xy=(date, valeur),
                        xytext=(10, -15),
                        textcoords="offset points",
                        ha="left",
                        fontsize=8,
                        color="#3498db",
                        bbox={"boxstyle": "round,pad=0.3", "facecolor": "white", "edgecolor": "#3498db", "alpha": 0.8},
                    )

    # Tracé de l'Earned Value si disponible
    if ev_cumulee is not None:
        ev_ke = ev_cumulee / 1000
        dates_ev = [periode.to_timestamp() for periode in ev_cumulee.index]
        plt.plot(
            dates_ev,
            ev_ke.values,
            marker="^",
            linewidth=2.5,
            markersize=8,
            label="EV (Earned Value - Valeur acquise)",
            color="#2ecc71",
            linestyle="-.",
            zorder=2,
        )

    # Tracé de l'EAC projeté si disponible
    if eac_projete is not None:
        eac_ke = eac_projete / 1000
        dates_eac = [periode.to_timestamp() for periode in eac_projete.index]
        plt.plot(
            dates_eac,
            eac_ke.values,
            marker="*",
            linewidth=2.5,
            markersize=10,
            label="EAC (Estimate at Completion - Projection)",
            color="#f39c12",
            linestyle="--",
            zorder=2,
            alpha=0.8,
        )

        # Ajouter les jalons projetés sur la courbe EAC
        if jalons_forecast:
            for jalon, info in jalons_forecast.items():
                mois = info["date"]
                if mois in eac_projete.index:
                    date = mois.to_timestamp()
                    valeur = eac_projete[mois] / 1000
                    plt.annotate(
                        jalon,
                        xy=(date, valeur),
                        xytext=(10, 10),
                        textcoords="offset points",
                        ha="left",
                        fontsize=8,
                        color="#f39c12",
                        bbox={"boxstyle": "round,pad=0.3", "facecolor": "white", "edgecolor": "#f39c12", "alpha": 0.8},
                    )

    # Calcul et tracé de CV (Cost Variance) et SV (Schedule Variance)
    if ev_cumulee is not None and pv_cumulee is not None:
        # Trouver les mois communs pour CV
        mois_communs_cv = depenses_cumulees.index.intersection(ev_cumulee.index)
        if len(mois_communs_cv) > 0:
            cv_values = []
            cv_dates = []
            for mois in mois_communs_cv:
                cv = (ev_cumulee[mois] - depenses_cumulees[mois]) / 1000
                cv_values.append(cv)
                cv_dates.append(mois.to_timestamp())

            plt.plot(
                cv_dates,
                cv_values,
                marker="d",
                linewidth=2,
                markersize=6,
                label="CV (Cost Variance = EV - AC)",
                color="#e67e22",
                linestyle=":",
                zorder=1,
                alpha=0.7,
            )

            # Ajouter les valeurs de CV sur le dernier point
            if len(cv_dates) > 0:
                plt.annotate(
                    f"CV: {cv_values[-1]:.1f} k€",
                    xy=(cv_dates[-1], cv_values[-1]),
                    xytext=(10, 10),
                    textcoords="offset points",
                    ha="left",
                    fontsize=9,
                    color="#e67e22",
                    bbox={"boxstyle": "round,pad=0.4", "facecolor": "white", "edgecolor": "#e67e22", "alpha": 0.9},
                )

        # Trouver les mois communs pour SV
        mois_communs_sv = ev_cumulee.index.intersection(pv_cumulee.index)
        if len(mois_communs_sv) > 0:
            sv_values = []
            sv_dates = []
            for mois in mois_communs_sv:
                sv = (ev_cumulee[mois] - pv_cumulee[mois]) / 1000
                sv_values.append(sv)
                sv_dates.append(mois.to_timestamp())

            plt.plot(
                sv_dates,
                sv_values,
                marker="v",
                linewidth=2,
                markersize=6,
                label="SV (Schedule Variance = EV - PV)",
                color="#9b59b6",
                linestyle=":",
                zorder=1,
                alpha=0.7,
            )

            # Ajouter les valeurs de SV sur le dernier point
            if len(sv_dates) > 0:
                plt.annotate(
                    f"SV: {sv_values[-1]:.1f} k€",
                    xy=(sv_dates[-1], sv_values[-1]),
                    xytext=(10, -20),
                    textcoords="offset points",
                    ha="left",
                    fontsize=9,
                    color="#9b59b6",
                    bbox={"boxstyle": "round,pad=0.4", "facecolor": "white", "edgecolor": "#9b59b6", "alpha": 0.9},
                )

    plt.xlabel("Mois", fontsize=12)
    plt.ylabel("Montant cumulé (k€)", fontsize=12)
    plt.title("Analyse EVM - AC vs PV vs EV", fontsize=14, fontweight="bold")
    plt.grid(True, alpha=0.3)
    plt.legend(loc="upper left", fontsize=9)

    # Format des dates sur l'axe x
    plt.gcf().autofmt_xdate()

    # Ajout des valeurs sur les points (dépenses réelles) - uniquement quelques points clés
    step = max(1, len(dates_depenses) // 6)  # Afficher environ 6 labels
    for i, (date, valeur) in enumerate(zip(dates_depenses, depenses_ke.values)):
        if i % step == 0 or i == len(dates_depenses) - 1:
            plt.annotate(
                f"{valeur:.0f}",
                xy=(date, valeur),
                xytext=(0, 10),
                textcoords="offset points",
                ha="center",
                fontsize=7,
                color="#e74c3c",
            )

    plt.tight_layout()
    plt.savefig(fichier_sortie, dpi=300, bbox_inches="tight")
    plt.close()
    print(f"Graphique sauvegardé: {fichier_sortie}")


def tracer_courbe_realise(
    depenses_cumulees, pv_cumulee=None, jalons=None, ev_cumulee=None, fichier_sortie="analyse_evm_realise.png"
):
    """
    Trace le graphique du réalisé à date: AC, PV, EV avec CV et SV
    """
    plt.figure(figsize=(16, 9))

    # Conversion en k€ pour une meilleure lisibilité
    depenses_ke = depenses_cumulees / 1000

    # Conversion des périodes en dates pour l'affichage (dépenses réelles)
    dates_depenses = [periode.to_timestamp() for periode in depenses_cumulees.index]

    # Tracé des dépenses réelles (AC - Actual Cost)
    plt.plot(
        dates_depenses,
        depenses_ke.values,
        marker="o",
        linewidth=2.5,
        markersize=8,
        label="AC (Actual Cost - Dépenses réelles)",
        color="#e74c3c",
        zorder=3,
    )

    # Tracé de la Planned Value si disponible
    if pv_cumulee is not None:
        pv_ke = pv_cumulee / 1000
        dates_pv = [periode.to_timestamp() for periode in pv_cumulee.index]
        plt.plot(
            dates_pv,
            pv_ke.values,
            marker="s",
            linewidth=2.5,
            markersize=8,
            label="PV (Planned Value - Budget prévu)",
            color="#3498db",
            linestyle="--",
            zorder=2,
        )

        # Ajout des jalons sur la courbe PV
        if jalons:
            for periode, valeur in zip(pv_cumulee.index, pv_ke.values):
                if periode in jalons:
                    date = periode.to_timestamp()
                    jalons_str = "\n".join(jalons[periode])
                    plt.annotate(
                        jalons_str,
                        xy=(date, valeur),
                        xytext=(10, -15),
                        textcoords="offset points",
                        ha="left",
                        fontsize=8,
                        color="#3498db",
                        bbox={"boxstyle": "round,pad=0.3", "facecolor": "white", "edgecolor": "#3498db", "alpha": 0.8},
                    )

    # Tracé de l'Earned Value si disponible
    if ev_cumulee is not None:
        ev_ke = ev_cumulee / 1000
        dates_ev = [periode.to_timestamp() for periode in ev_cumulee.index]
        plt.plot(
            dates_ev,
            ev_ke.values,
            marker="^",
            linewidth=2.5,
            markersize=8,
            label="EV (Earned Value - Valeur acquise)",
            color="#2ecc71",
            linestyle="-.",
            zorder=2,
        )

    # Calcul et tracé de CV (Cost Variance) et SV (Schedule Variance)
    if ev_cumulee is not None and pv_cumulee is not None:
        # Trouver les mois communs pour CV
        mois_communs_cv = depenses_cumulees.index.intersection(ev_cumulee.index)
        if len(mois_communs_cv) > 0:
            cv_values = []
            cv_dates = []
            for mois in mois_communs_cv:
                cv = (ev_cumulee[mois] - depenses_cumulees[mois]) / 1000
                cv_values.append(cv)
                cv_dates.append(mois.to_timestamp())

            plt.plot(
                cv_dates,
                cv_values,
                marker="d",
                linewidth=2,
                markersize=6,
                label="CV (Cost Variance = EV - AC)",
                color="#e67e22",
                linestyle=":",
                zorder=1,
                alpha=0.7,
            )

            # Ajouter les valeurs de CV sur le dernier point
            if len(cv_dates) > 0:
                plt.annotate(
                    f"CV: {cv_values[-1]:.1f} k€",
                    xy=(cv_dates[-1], cv_values[-1]),
                    xytext=(10, 10),
                    textcoords="offset points",
                    ha="left",
                    fontsize=9,
                    color="#e67e22",
                    bbox={"boxstyle": "round,pad=0.4", "facecolor": "white", "edgecolor": "#e67e22", "alpha": 0.9},
                )

        # Trouver les mois communs pour SV
        mois_communs_sv = ev_cumulee.index.intersection(pv_cumulee.index)
        if len(mois_communs_sv) > 0:
            sv_values = []
            sv_dates = []
            for mois in mois_communs_sv:
                sv = (ev_cumulee[mois] - pv_cumulee[mois]) / 1000
                sv_values.append(sv)
                sv_dates.append(mois.to_timestamp())

            plt.plot(
                sv_dates,
                sv_values,
                marker="v",
                linewidth=2,
                markersize=6,
                label="SV (Schedule Variance = EV - PV)",
                color="#9b59b6",
                linestyle=":",
                zorder=1,
                alpha=0.7,
            )

            # Ajouter les valeurs de SV sur le dernier point
            if len(sv_dates) > 0:
                plt.annotate(
                    f"SV: {sv_values[-1]:.1f} k€",
                    xy=(sv_dates[-1], sv_values[-1]),
                    xytext=(10, -20),
                    textcoords="offset points",
                    ha="left",
                    fontsize=9,
                    color="#9b59b6",
                    bbox={"boxstyle": "round,pad=0.4", "facecolor": "white", "edgecolor": "#9b59b6", "alpha": 0.9},
                )

    plt.xlabel("Mois", fontsize=12)
    plt.ylabel("Montant cumulé (k€)", fontsize=12)
    plt.title("Analyse EVM - Réalisé à date (AC vs PV vs EV)", fontsize=14, fontweight="bold")
    plt.grid(True, alpha=0.3)
    plt.legend(loc="upper left", fontsize=9)

    # Format des dates sur l'axe x
    plt.gcf().autofmt_xdate()

    # Ajout des valeurs sur les points (dépenses réelles) - uniquement quelques points clés
    step = max(1, len(dates_depenses) // 6)  # Afficher environ 6 labels
    for i, (date, valeur) in enumerate(zip(dates_depenses, depenses_ke.values)):
        if i % step == 0 or i == len(dates_depenses) - 1:
            plt.annotate(
                f"{valeur:.0f}",
                xy=(date, valeur),
                xytext=(0, 10),
                textcoords="offset points",
                ha="center",
                fontsize=7,
                color="#e74c3c",
            )

    plt.tight_layout()
    plt.savefig(fichier_sortie, dpi=300, bbox_inches="tight")
    plt.close()
    print(f"✓ Graphique du réalisé sauvegardé: {fichier_sortie}")


def tracer_courbe_projections(depenses_cumulees, ev_cumulee, projections, fichier_sortie="analyse_evm_projections.png"):
    """
    Trace le graphique des projections à terminaison: historique AC/EV + différents scénarios EAC

    Args:
        depenses_cumulees: Série temporelle des dépenses réelles (AC)
        ev_cumulee: Série temporelle de la valeur acquise (EV)
        projections: Dictionnaire contenant les projections calculées
                    {'cpi': {series, eac, date}, 'cpi_spi': {...}, 'reste_plan': {...}, 'forecast': {...}}
        fichier_sortie: Nom du fichier de sortie
    """
    plt.figure(figsize=(16, 9))

    # Conversion en k€ pour une meilleure lisibilité
    depenses_ke = depenses_cumulees / 1000

    # Conversion des périodes en dates pour l'affichage
    dates_depenses = [periode.to_timestamp() for periode in depenses_cumulees.index]

    # Tracé de l'historique AC
    plt.plot(
        dates_depenses,
        depenses_ke.values,
        marker="o",
        linewidth=2.5,
        markersize=8,
        label="AC historique (Dépenses réelles)",
        color="#e74c3c",
        zorder=3,
    )

    # Tracé de l'EV si disponible
    if ev_cumulee is not None:
        ev_ke = ev_cumulee / 1000
        dates_ev = [periode.to_timestamp() for periode in ev_cumulee.index]
        plt.plot(
            dates_ev,
            ev_ke.values,
            marker="^",
            linewidth=2.5,
            markersize=8,
            label="EV historique (Valeur acquise)",
            color="#2ecc71",
            linestyle="-.",
            zorder=2,
        )

    # Couleurs et styles pour les projections
    couleurs = {"cpi": "#f39c12", "cpi_spi": "#9b59b6", "reste_plan": "#e67e22", "forecast": "#3498db"}

    labels_proj = {
        "reste_plan": "EAC reste à plan (optimiste)",
        "cpi": "EAC méthode CPI (réaliste)",
        "cpi_spi": "EAC méthode CPI×SPI (pessimiste)",
        "forecast": "EAC forecast manuel",
    }

    # Tracé de chaque projection
    for methode, couleur in couleurs.items():
        if methode in projections and projections[methode] is not None:
            proj_data = projections[methode]
            series = proj_data["series"] / 1000
            dates_proj = [periode.to_timestamp() for periode in series.index]
            eac_final = proj_data["eac"] / 1000
            date_fin = proj_data["date"]

            plt.plot(
                dates_proj,
                series.values,
                marker="*",
                linewidth=2.5,
                markersize=8,
                label=f"{labels_proj[methode]} ({eac_final:.0f} k€)",
                color=couleur,
                linestyle="--",
                zorder=2,
                alpha=0.8,
            )

            # Ajouter annotation sur le point final
            if len(dates_proj) > 0:
                plt.annotate(
                    f"{eac_final:.0f} k€\n{date_fin.strftime('%m/%Y')}",
                    xy=(dates_proj[-1], series.to_numpy()[-1]),
                    xytext=(10, 10),
                    textcoords="offset points",
                    ha="left",
                    fontsize=8,
                    color=couleur,
                    bbox={"boxstyle": "round,pad=0.3", "facecolor": "white", "edgecolor": couleur, "alpha": 0.8},
                )

    plt.xlabel("Mois", fontsize=12)
    plt.ylabel("Montant cumulé (k€)", fontsize=12)
    plt.title("Analyse EVM - Projections à terminaison (scénarios EAC)", fontsize=14, fontweight="bold")
    plt.grid(True, alpha=0.3)
    plt.legend(loc="upper left", fontsize=9)

    # Format des dates sur l'axe x
    plt.gcf().autofmt_xdate()

    plt.tight_layout()
    plt.savefig(fichier_sortie, dpi=300, bbox_inches="tight")
    plt.close()
    print(f"✓ Graphique des projections sauvegardé: {fichier_sortie}")

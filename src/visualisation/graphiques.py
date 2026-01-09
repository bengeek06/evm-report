"""
Module de génération de graphiques EVM
"""

import matplotlib.pyplot as plt

FIGSIZE = (16, 9)
TEXTCOORDS_OFFSET_POINTS = "offset points"
YLABEL_MONTANT_CUMULE_KE = "Montant cumulé (k€)"
LEGEND_LOC_UPPER_LEFT = "upper left"
BOXSTYLE_ROUND_PAD_03 = "round,pad=0.3"
BOXSTYLE_ROUND_PAD_04 = "round,pad=0.4"


def _new_figure() -> None:
    plt.figure(figsize=FIGSIZE)


def _bbox(boxstyle: str, edgecolor: str, alpha: float = 0.8) -> dict:
    return {"boxstyle": boxstyle, "facecolor": "white", "edgecolor": edgecolor, "alpha": alpha}


def _plot_series(
    dates, values, *, marker: str, label: str, color: str, linestyle: str = "-", zorder: int = 2, alpha: float = 1.0
) -> None:
    plt.plot(
        dates,
        values,
        marker=marker,
        linewidth=2.5,
        markersize=8,
        label=label,
        color=color,
        linestyle=linestyle,
        zorder=zorder,
        alpha=alpha,
    )


def _plot_ac(depenses_cumulees):
    depenses_ke = depenses_cumulees / 1000
    dates_depenses = [periode.to_timestamp() for periode in depenses_cumulees.index]
    _plot_series(
        dates_depenses,
        depenses_ke.values,
        marker="o",
        label="AC (Actual Cost - Dépenses réelles)",
        color="#e74c3c",
        zorder=3,
    )
    return depenses_ke, dates_depenses


def _plot_pv(pv_cumulee, jalons=None) -> None:
    pv_ke = pv_cumulee / 1000
    dates_pv = [periode.to_timestamp() for periode in pv_cumulee.index]
    _plot_series(
        dates_pv,
        pv_ke.values,
        marker="s",
        label="PV (Planned Value - Budget prévu)",
        color="#3498db",
        linestyle="--",
        zorder=2,
    )
    if jalons:
        _annotate_jalons(jalons, pv_cumulee.index, pv_ke.values, color="#3498db", xytext=(10, -15))


def _plot_ev(ev_cumulee, *, label: str = "EV (Earned Value - Valeur acquise)") -> None:
    ev_ke = ev_cumulee / 1000
    dates_ev = [periode.to_timestamp() for periode in ev_cumulee.index]
    _plot_series(
        dates_ev,
        ev_ke.values,
        marker="^",
        label=label,
        color="#2ecc71",
        linestyle="-.",
        zorder=2,
    )


def _plot_eac_projete(eac_projete, jalons_forecast=None) -> None:
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

    if not jalons_forecast:
        return

    for jalon, info in jalons_forecast.items():
        mois = info["date"]
        if mois in eac_projete.index:
            date = mois.to_timestamp()
            valeur = eac_projete[mois] / 1000
            _annotate(
                jalon,
                xy=(date, valeur),
                xytext=(10, 10),
                color="#f39c12",
                fontsize=8,
                boxstyle=BOXSTYLE_ROUND_PAD_03,
            )


def _annotate_key_points(dates, values, *, step_divisor: int = 6, color: str = "#e74c3c") -> None:
    step = max(1, len(dates) // step_divisor)
    for i, (date, valeur) in enumerate(zip(dates, values)):
        if i % step == 0 or i == len(dates) - 1:
            plt.annotate(
                f"{valeur:.0f}",
                xy=(date, valeur),
                xytext=(0, 10),
                textcoords=TEXTCOORDS_OFFSET_POINTS,
                ha="center",
                fontsize=7,
                color=color,
            )


def _annotate(text: str, *, xy, xytext, color: str, fontsize: int, boxstyle: str, ha: str = "left") -> None:
    plt.annotate(
        text,
        xy=xy,
        xytext=xytext,
        textcoords=TEXTCOORDS_OFFSET_POINTS,
        ha=ha,
        fontsize=fontsize,
        color=color,
        bbox=_bbox(boxstyle, color),
    )


def _annotate_jalons(jalons: dict, periodes, valeurs, *, color: str, xytext=(10, -15)) -> None:
    for periode, valeur in zip(periodes, valeurs):
        if periode in jalons:
            date = periode.to_timestamp()
            _annotate(
                "\n".join(jalons[periode]),
                xy=(date, valeur),
                xytext=xytext,
                color=color,
                fontsize=8,
                boxstyle=BOXSTYLE_ROUND_PAD_03,
            )


def _plot_variance(depenses_cumulees, pv_cumulee, ev_cumulee) -> None:
    mois_communs_cv = depenses_cumulees.index.intersection(ev_cumulee.index)
    if len(mois_communs_cv) > 0:
        cv_values = [(ev_cumulee[mois] - depenses_cumulees[mois]) / 1000 for mois in mois_communs_cv]
        cv_dates = [mois.to_timestamp() for mois in mois_communs_cv]
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
        if len(cv_dates) > 0:
            _annotate(
                f"CV: {cv_values[-1]:.1f} k€",
                xy=(cv_dates[-1], cv_values[-1]),
                xytext=(10, 10),
                color="#e67e22",
                fontsize=9,
                boxstyle=BOXSTYLE_ROUND_PAD_04,
            )

    mois_communs_sv = ev_cumulee.index.intersection(pv_cumulee.index)
    if len(mois_communs_sv) > 0:
        sv_values = [(ev_cumulee[mois] - pv_cumulee[mois]) / 1000 for mois in mois_communs_sv]
        sv_dates = [mois.to_timestamp() for mois in mois_communs_sv]
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
        if len(sv_dates) > 0:
            _annotate(
                f"SV: {sv_values[-1]:.1f} k€",
                xy=(sv_dates[-1], sv_values[-1]),
                xytext=(10, -20),
                color="#9b59b6",
                fontsize=9,
                boxstyle=BOXSTYLE_ROUND_PAD_04,
            )


def _finalize_plot(*, title: str, fichier_sortie: str, print_msg: str) -> None:
    plt.xlabel("Mois", fontsize=12)
    plt.ylabel(YLABEL_MONTANT_CUMULE_KE, fontsize=12)
    plt.title(title, fontsize=14, fontweight="bold")
    plt.grid(True, alpha=0.3)
    plt.legend(loc=LEGEND_LOC_UPPER_LEFT, fontsize=9)
    plt.gcf().autofmt_xdate()
    plt.tight_layout()
    plt.savefig(fichier_sortie, dpi=300, bbox_inches="tight")
    plt.close(plt.gcf())
    print(print_msg)


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
    _new_figure()

    depenses_ke, dates_depenses = _plot_ac(depenses_cumulees)

    # Tracé de la Planned Value si disponible
    if pv_cumulee is not None:
        _plot_pv(pv_cumulee, jalons)

    # Tracé de l'Earned Value si disponible
    if ev_cumulee is not None:
        _plot_ev(ev_cumulee)

    # Tracé de l'EAC projeté si disponible
    if eac_projete is not None:
        _plot_eac_projete(eac_projete, jalons_forecast)

    # Calcul et tracé de CV (Cost Variance) et SV (Schedule Variance)
    if ev_cumulee is not None and pv_cumulee is not None:
        _plot_variance(depenses_cumulees, pv_cumulee, ev_cumulee)

    _annotate_key_points(dates_depenses, depenses_ke.values, color="#e74c3c")

    _finalize_plot(
        title="Analyse EVM - AC vs PV vs EV",
        fichier_sortie=fichier_sortie,
        print_msg=f"Graphique sauvegardé: {fichier_sortie}",
    )


def tracer_courbe_realise(
    depenses_cumulees, pv_cumulee=None, jalons=None, ev_cumulee=None, fichier_sortie="analyse_evm_realise.png"
):
    """
    Trace le graphique du réalisé à date: AC, PV, EV avec CV et SV
    """
    _new_figure()

    depenses_ke, dates_depenses = _plot_ac(depenses_cumulees)

    # Tracé de la Planned Value si disponible
    if pv_cumulee is not None:
        _plot_pv(pv_cumulee)

        # Ajout des jalons sur la courbe PV
        if jalons:
            pv_ke = pv_cumulee / 1000
            _annotate_jalons(jalons, pv_cumulee.index, pv_ke.values, color="#3498db", xytext=(10, -15))

    # Tracé de l'Earned Value si disponible
    if ev_cumulee is not None:
        _plot_ev(ev_cumulee)

    # Calcul et tracé de CV (Cost Variance) et SV (Schedule Variance)
    if ev_cumulee is not None and pv_cumulee is not None:
        _plot_variance(depenses_cumulees, pv_cumulee, ev_cumulee)

    plt.xlabel("Mois", fontsize=12)
    plt.ylabel("Montant cumulé (k€)", fontsize=12)
    plt.title("Analyse EVM - Réalisé à date (AC vs PV vs EV)", fontsize=14, fontweight="bold")
    plt.grid(True, alpha=0.3)
    plt.legend(loc="upper left", fontsize=9)

    # Format des dates sur l'axe x
    plt.gcf().autofmt_xdate()

    _annotate_key_points(dates_depenses, depenses_ke.values, color="#e74c3c")

    _finalize_plot(
        title="Analyse EVM - Réalisé à date (AC vs PV vs EV)",
        fichier_sortie=fichier_sortie,
        print_msg=f"✓ Graphique du réalisé sauvegardé: {fichier_sortie}",
    )


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
    _new_figure()

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
                _annotate(
                    f"{eac_final:.0f} k€\n{date_fin.strftime('%m/%Y')}",
                    xy=(dates_proj[-1], series.to_numpy()[-1]),
                    xytext=(10, 10),
                    color=couleur,
                    fontsize=8,
                    boxstyle=BOXSTYLE_ROUND_PAD_03,
                )

    _finalize_plot(
        title="Analyse EVM - Projections à terminaison (scénarios EAC)",
        fichier_sortie=fichier_sortie,
        print_msg=f"✓ Graphique des projections sauvegardé: {fichier_sortie}",
    )

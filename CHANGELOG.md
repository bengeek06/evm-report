# Changelog - Analyse EVM

## Version 2.0 - Projections Automatiques et Double Graphique

### Nouvelles fonctionnalités

#### 1. Calcul automatique des projections EAC
Le script calcule maintenant automatiquement 3 scénarios de projection selon des méthodes EVM standards :

- **Méthode CPI (Optimiste)** : `EAC = BAC / CPI`
  - Suppose que la performance des coûts actuelle se poursuit
  - Généralement le scénario le plus optimiste
  
- **Méthode CPI×SPI (Réaliste)** : `EAC = AC + [(BAC - EV) / (CPI × SPI)]`
  - Prend en compte à la fois la performance des coûts ET des délais
  - Scénario le plus réaliste et couramment utilisé
  
- **Méthode Reste à Plan (Pessimiste)** : `EAC = AC + (BAC - EV)`
  - Suppose que le reste se déroule exactement comme prévu initialement
  - Généralement le scénario le plus pessimiste si le projet a des problèmes

#### 2. Visualisation en deux graphiques séparés

**Graphique 1 : Réalisé à date** (`*_realise.png`)
- AC (Actual Cost) - Dépenses réelles
- PV (Planned Value) - Budget prévu
- EV (Earned Value) - Valeur acquise
- CV (Cost Variance) - Écart de coût
- SV (Schedule Variance) - Écart de délai
- Jalons planifiés sur la courbe PV

**Graphique 2 : Projections à terminaison** (`*_projections.png`)
- Historique AC et EV
- Projection CPI (optimiste) en orange
- Projection CPI×SPI (réaliste) en violet
- Projection Reste à Plan (pessimiste) en orange foncé
- Projection Forecast manuel (si fichier fourni) en bleu
- Annotations avec montants finaux et dates estimées

#### 3. Rapport Word enrichi

Le rapport Word a été restructuré en 4 sections :

**Section 1 : Définitions EVM**
- Définitions de tous les termes et indicateurs

**Section 2 : Réalisé à Date**
- 2.1 Tableau des valeurs AC, PV, EV, CV, SV
- 2.2 Graphique du réalisé
- 2.3 Indicateurs de performance actuels
  - Valeurs actuelles avec code couleur (rouge/vert)
  - Interprétation de CV, SV, CPI, SPI

**Section 3 : Projections à Terminaison**
- 3.1 Tableau comparatif des scénarios
  - Comparaison des 3-4 méthodes de projection
  - EAC, Date de fin, VAC pour chaque scénario
- 3.2 Graphique des projections
- 3.3 Analyse des scénarios
  - Fourchette optimiste/moyen/pessimiste
  - Écarts par rapport au budget (% et montants)

**Section 4 : Conclusion et Recommandations**
- 4.1 Synthèse
  - Évaluation de la performance actuelle (CPI, SPI)
  - Analyse des projections (risques de dépassement)
  - Évaluation de l'incertitude
- 4.2 Recommandations
  - Actions recommandées selon la situation
  - Priorités et mesures correctives

### Améliorations techniques

- Les projections sont calculées automatiquement à partir des données AC, PV, EV
- Le fichier `forecast.xlsx` devient optionnel (4ème scénario si fourni)
- Génération de séries temporelles pour chaque projection
- Estimation automatique de la date de fin basée sur le SPI
- Code couleur dans le rapport Word (rouge = négatif, vert = positif)
- Analyses et recommandations automatiques selon les indicateurs

### Utilisation

```bash
# Avec forecast manuel (4 scénarios)
python analyse.py --sap EXPORT.XLSX --pv pv.xlsx --va va.xlsx --forecast forecast.xlsx --word rapport.docx

# Sans forecast (3 scénarios automatiques uniquement)
python analyse.py --sap EXPORT.XLSX --pv pv.xlsx --va va.xlsx --word rapport.docx

# Génération des graphiques seulement
python analyse.py --sap EXPORT.XLSX --pv pv.xlsx --va va.xlsx --output graphiques.png
```

### Sorties

- `*_realise.png` : Graphique du réalisé (AC, PV, EV, CV, SV)
- `*_projections.png` : Graphique des projections (3-4 scénarios EAC)
- `rapport_evm.docx` : Rapport Word complet avec analyses et recommandations
- `tableau_evm.csv` / `tableau_evm.xlsx` : Tableaux de données (supprimés si --word utilisé)

### Notes importantes

1. Les 3 méthodes de projection sont calculées automatiquement - pas besoin de fichier forecast pour avoir des projections
2. Le fichier `forecast.xlsx` ajoute un 4ème scénario (projection manuelle) si besoin
3. La fourchette optimiste-pessimiste permet d'évaluer l'incertitude du projet
4. Les recommandations s'adaptent automatiquement selon les indicateurs CPI/SPI

---

## Version 1.0 - Analyse EVM de Base

### Fonctionnalités initiales

- Lecture des exports SAP (EXPORT.XLSX)
- Calcul de la Planned Value depuis pv.xlsx
- Calcul de l'Earned Value depuis va.xlsx
- Interpolation linéaire de la PV entre jalons
- Génération d'un graphique unique avec AC, PV, EV, CV, SV, EAC
- Génération de rapport Word avec définitions, tableau, graphique, conclusion
- Arguments en ligne de commande avec argparse
- Gestion des jalons avec annotations sur les courbes

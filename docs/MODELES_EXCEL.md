# Modèles Excel - Guide d'utilisation

Ce dossier contient 4 fichiers modèles Excel avec les colonnes attendues par le script d'analyse EVM.

## Fichiers modèles disponibles

### 1. EXPORT_model.xlsx - Export SAP des dépenses

**Colonnes obligatoires :**
- `Date de la pièce` : Date de la dépense (format date Excel)
- `Val./Devise objet` : Montant de la dépense en euros (format nombre)

**Colonnes optionnelles (présentes dans l'export SAP standard) :**
- Document d'achat
- Exercice comptable
- Période
- Elément d'OTP
- Nom 1
- Nature comptable
- Quantité totale
- Désign.nat.comptable
- Nº pièce référence
- Qté totale saisie
- Texte de la commande d'achat
- Groupe d'origine
- Référence

**Format des données :**
```
Date de la pièce    | Val./Devise objet
--------------------|------------------
2025-01-15          | 15000.50
2025-02-20          | 25000.75
2025-03-10          | 18500.00
```

---

### 2. pv_model.xlsx - Planned Value (Budget prévu)

**Colonnes obligatoires :**
- `Jalon` : Nom du jalon (ex: RCD, J1, J2, J3, J4, POV)
- `Date` : Date prévue du jalon (format date Excel)
- `Montant planifié` : Budget alloué à ce jalon en euros
- `Cumul planifié` : Budget cumulé jusqu'à ce jalon en euros

**Colonnes optionnelles :**
- `Durée (mois)` : Durée en mois pour ce jalon

**Format des données :**
```
Jalon | Durée (mois) | Date       | Montant planifié | Cumul planifié
------|--------------|------------|------------------|----------------
RCD   | 3            | 2025-03-31 | 100000          | 100000
J1    | 5            | 2025-08-31 | 200000          | 300000
J2    | 4            | 2025-12-31 | 250000          | 550000
J3    | 6            | 2026-06-30 | 350000          | 900000
```

**Important :**
- Le script interpolera linéairement la PV entre les jalons
- Assurez-vous que les dates sont ordonnées chronologiquement
- Le "Cumul planifié" doit être croissant

---

### 3. va_model.xlsx - Valeur Acquise (% d'avancement)

**Colonnes obligatoires :**
- `Jalon` : Nom du jalon (doit correspondre aux jalons de pv.xlsx)
- `Date` : Date prévue du jalon
- `Montant planifié` : Montant planifié pour ce jalon
- `Cumul planifié` : Budget cumulé
- **Colonnes mensuelles** avec dates (ex: 2025-01-01, 2025-02-01, etc.)

**Format des colonnes mensuelles :**
- Les en-têtes de colonnes sont des dates (format date Excel)
- Les valeurs sont des pourcentages en décimal : **0.0 à 1.0**
  - 0.0 = 0% d'avancement
  - 0.5 = 50% d'avancement
  - 1.0 = 100% d'avancement

**Exemple de données :**
```
Jalon | Date       | Montant | 2025-01-01 | 2025-02-01 | 2025-03-01 | 2025-04-01
------|------------|---------|------------|------------|------------|------------
RCD   | 2025-03-31 | 100000  | 0.2        | 0.4        | 0.6        | 1.0
J1    | 2025-08-31 | 200000  | 0.0        | 0.0        | 0.0        | 0.2
J2    | 2025-12-31 | 250000  | 0.0        | 0.0        | 0.0        | 0.0
```

**⚠️ ATTENTION :**
- Les pourcentages sont en décimal (0.5 = 50%, PAS 50)
- Ajoutez autant de colonnes mensuelles que nécessaire
- Les dates des colonnes peuvent être le 1er du mois ou n'importe quel jour du mois

---

### 4. forecast_model.xlsx - Projections manuelles (OPTIONNEL)

**Colonnes obligatoires :**
- `Jalon` : Nom du jalon à projeter
- `Date projetée` : Date estimée de réalisation (format date Excel)
- `EAC (€)` : Estimate at Completion - Coût estimé pour ce jalon

**Colonnes optionnelles :**
- `ETC (€)` : Estimate to Complete - Reste à dépenser
- `Commentaire` : Notes ou explications

**Format des données :**
```
Jalon | Date projetée | EAC (€)  | ETC (€)  | Commentaire
------|---------------|----------|----------|---------------------------
J3    | 2026-07-31    | 420000   | 420000   | Risque sur les délais
J4    | 2026-10-31    | 230000   | 230000   | En cours de négociation
J5    | 2027-02-28    | 120000   | 120000   | Budget confirmé
POV   | 2027-06-30    | 20000    | 20000    | Phase finale
```

**Note :**
- Ce fichier est **OPTIONNEL**
- S'il n'est pas fourni, le script calculera automatiquement 3 scénarios de projection (CPI, CPI×SPI, Reste à Plan)
- S'il est fourni, il sera ajouté comme 4ème scénario dans les projections

---

## Utilisation des modèles

### Étape 1 : Copier les modèles
```bash
cp EXPORT_model.xlsx EXPORT.xlsx
cp pv_model.xlsx pv.xlsx
cp va_model.xlsx va.xlsx
cp forecast_model.xlsx forecast.xlsx  # Optionnel
```

### Étape 2 : Remplir avec vos données
Ouvrez les fichiers copiés avec Excel/LibreOffice et remplacez les données d'exemple par vos données réelles.

### Étape 3 : Exécuter l'analyse
```bash
python analyse.py --word rapport_evm.docx
```

---

## Conseils et bonnes pratiques

### Export SAP (EXPORT.xlsx)
- Exportez depuis SAP toutes les lignes de dépenses du projet
- Vérifiez que les montants sont bien en euros
- Assurez-vous que les dates sont au bon format

### Planned Value (pv.xlsx)
- Listez tous les jalons majeurs du projet
- Le dernier jalon doit avoir le budget total du projet (BAC)
- Les dates doivent couvrir toute la durée du projet

### Valeur Acquise (va.xlsx)
- Mettez à jour mensuellement les pourcentages d'avancement
- Soyez réaliste dans l'évaluation de l'avancement
- Un jalon à 100% (1.0) doit être réellement terminé et validé

### Forecast (forecast.xlsx)
- Utilisez ce fichier pour des projections personnalisées
- Basez-vous sur l'expertise métier et les risques identifiés
- Mettez à jour régulièrement selon l'avancement réel

---

## Validation des données

Avant d'exécuter l'analyse, vérifiez que :

1. ✅ Les noms de jalons sont cohérents entre pv.xlsx et va.xlsx
2. ✅ Les dates dans EXPORT.xlsx couvrent la période d'analyse
3. ✅ Les pourcentages dans va.xlsx sont entre 0.0 et 1.0
4. ✅ Les montants planifiés sont les mêmes dans pv.xlsx et va.xlsx
5. ✅ Le cumul planifié est croissant dans pv.xlsx
6. ✅ Les dates des jalons sont ordonnées chronologiquement

---

## Support

Pour plus d'informations, consultez :
- [README.md](README.md) - Documentation complète du script
- [CHANGELOG.md](CHANGELOG.md) - Historique des versions
- `python analyse.py --help` - Aide en ligne de commande

En cas de problème avec les fichiers Excel, le script affichera des messages d'erreur explicites indiquant les colonnes manquantes ou les formats incorrects.

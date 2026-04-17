# Données de test — suivi_conso_dsi

Jeu de données fictif pour tester le script `suivi_conso_dsi_v5_00.py`.

## Fichiers

| Fichier | Rôle | Correspondance script |
|---------|------|-----------------------|
| `test_diana.txt` | Export Diana — saisies collaborateurs | Fichier 1 (séparateur `~`) |
| `test_pdc_jh.csv` | Plan de charges Jours/Hommes | Fichier 2 (séparateur `;`) |
| `test_pdc_eur.csv` | Plan de charges Euros | Fichier 3 (séparateur `;`) |
| `intervenants.csv` | Référentiel intervenants périmètre | Fichier 4 (séparateur `;`) |
| `Budgets_2026.csv` | Budgets projets 2026 | Auto-détecté au démarrage |
| `Histo_TJM.xlsx` | TJM historiques par intervenant/mois | Auto-détecté au démarrage |

## Intervenants fictifs

| Nom | Statut | Périmètre | Rôle dans les tests |
|-----|--------|-----------|---------------------|
| MARTIN, Jean | CA-CIB | Acquisition | Intervenant périmètre normal |
| DUPONT, Marie | Prestataire | Distribution | Intervenant périmètre normal |
| RENARD, Thomas | CA-CIB | Exploration | Intervenant périmètre normal |
| LECOMTE, Sophie | Prestataire | Data Services | Intervenant périmètre normal |
| BENOIT, Lucas | CA-CIB | Socle&Enabler | Intervenant périmètre normal |
| FAURE, Claire | Prestataire | Core Team | Intervenant périmètre normal |
| LECLERC, Emma | CA-CIB | **SurvComm uniquement** | Exclu par le filtre Nucleus |
| NGUYEN, Hoa | Prestataire | **Analytic uniquement** | Exclu par le filtre Nucleus |
| GUITTARD, Raphael | CA-CIB | Core Team + SurvComm | Hardcodé — exclusion partielle SurvComm |
| ABDELLI, Ahmed | Prestataire | Socle&Enabler (+ SurvComm) | Hardcodé — exclusion partielle SurvComm |
| TCHIEGAING KAMGUIA, Arnaud | Prestataire | **Analytic uniquement** | Hardcodé — exclusion partielle Analytic |
| BERNARD, Paul | Prestataire | *(hors périmètre)* | Dans Diana seulement (non dans intervenants.csv) |

## Projets fictifs

| Code | Libellé | Caractéristique |
|------|---------|-----------------|
| `P02508` | CPL - Target Architecture | Projet ciblé (TARGET_PROJECTS) |
| `P04243` | Projet Delta | Projet ciblé (TARGET_PROJECTS) |
| `A12345` | Projet Alpha | Projet périmètre normal |
| `A67890` | Projet Gamma | Projet périmètre normal |
| `B11111` | Projet Beta | Projet périmètre normal |
| `S99901` | Surveillance Comm. A | TopSurvComm = 1 |
| `N00001` | Projet Analytics | TopAnalytic = 1 |
| `V00001` | Congés | TopAbsence = 1 |

## Cohérence des données

- **Période Diana** : janvier à mars 2026 (mois passés par rapport à avril 2026)
- **Période PDC** : janvier à décembre 2026 (tous les mois de l'année)
- **TJM** : 650 €/j (CA-CIB) à 825 €/j (prestataires seniors)
- **EUR PDC** = NbJours × TJM correspondant dans Histo_TJM.xlsx
- **Écarts volontaires** : MARTIN, Jean a 19 JH PDC en janvier alors que le mois
  compte 21 jours ouvrés → écart de −2 JH affiché dans le tableau des écarts

## Cas de test couverts

1. **Filtre périmètre** : LECLERC (SurvComm) et NGUYEN (Analytic) exclus du
   rapport principal quand la case "Périmètre Nucleus uniquement" est cochée
2. **Exclusions partielles** : GUITTARD et ABDELLI — leurs JH sur S99901 sont
   exclus du calcul Nucleus mais ils restent visibles sur les autres projets
3. **Projets ciblés** : P02508 visible même si l'intervenant n'est pas dans le
   périmètre (règle TARGET_PROJECTS)
4. **Écarts JH/JO** : plusieurs intervenants ont des valeurs PDC ≠ jours ouvrés
   du mois → apparaissent dans le cadran "Écarts JH/JO"
5. **Consommés Diana vs PDC** : les consommés réels (diana) peuvent différer du
   plan → visible dans le tableau EUR par projet

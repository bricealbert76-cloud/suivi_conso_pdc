"""
Suivi Consommation DSI - v4.68
IHM de traitement des fichiers de saisie et plans de charges
  - Suivi Conso : génère un CSV filtré depuis le fichier saisies (.txt)
  - Histo TJM   : calcule et exporte les TJM mensuels par intervenant (Histo_TJM.xlsx)

  v4.63 : prétraitement des fichiers source
    - Génération automatique de CSV filtrés (_avec_filtre) pour Diana et les 2 PDC
    - Si le CSV filtré existe déjà, il est réutilisé directement (pas de relecture du source)
    - Le cache PDC est chargé au clic sur "MAJ Plan de charges" (plus au démarrage)
    - Le fichier intervenants (fichier 4) est présélectionné EN PREMIER au démarrage

  v4.64 : correctif consommés réels fenêtre MAJ PDC
    - Les consommés Diana (JH et EUR) ne prennent en compte que les lignes
      dont "Organization Unit Name" contient "RGU"

  v4.65 : rectification règle de gestion consommés réels fenêtre MAJ PDC
    - Condition élargie : lignes retenues si "Organization Unit Name" contient "RGU"
      OU si le projet est un projet piloté (TARGET_PROJECTS)

  v4.66 : fenêtre Vérif. données Crapull — utilisation des CSV filtrés
    - _open_crapull_window passe les fichiers _avec_filtre à CrapullWindow
      si ceux-ci existent (générés par le prétraitement v4.63)
    - Fallback sur les fichiers source originaux si le CSV filtré est absent

  v4.67 : fenêtre Vérif. données Diana — utilisation du CSV filtré
    - _open_diana_window passe le fichier _avec_filtre à DianaWindow
      si celui-ci existe (généré par le prétraitement v4.63)
    - Fallback sur le fichier source original si le CSV filtré est absent

  v4.68 : fenêtre MAJ PDC — filtre "Périmètre Nucleus uniquement"
    - Cadran bas-droite renommé "Plan de charges TTNR"
    - Case à cocher "Périmètre Nucleus uniquement" (désactivée par défaut)
    - Quand cochée : exclut du détail les intervenants intervenant
      uniquement sur "Surveillances des comm." ou uniquement sur "Analytic"
      (colonnes de intervenants.csv)

  v4.69 : erratum v4.68 — header custom déplacé sur le cadran bas-gauche
    - "Plan de charges TTNR" + checkbox sur le cadran bas-gauche (EUR)
    - Cadran bas-droite restauré en "Détail intervenants" simple

  v4.70 : debug filtre Nucleus — log des exclusions au clic checkbox
    - Au clic sur la checkbox, log dans le journal principal :
      liste des intervenants exclus, projet sélectionné, avertissement si vide

  v4.71 : filtre Nucleus — règle basée sur les projets PDC travaillés
    - SURV_COMM_PROJECTS et ANALYTIC_PROJECTS définis comme constantes
    - _compute_nucleus_exclusions() calcule dynamiquement depuis _cache_jh
      les intervenants dont TOUS les projets sont dans l'un de ces périmètres
    - Remplace la lecture depuis intervenants.csv (colonnes Surv.Comm./Analytic)

  v4.72 : filtre Nucleus — les projets vacances exclus du calcul
    - _compute_nucleus_exclusions() ignore les projets dans _vacance_projects
      (TopAbsence dans Budgets_2026.csv) avant d'évaluer le périmètre

  v4.73 : filtre Nucleus — réutilise _survcomm_intervenants (écarts JH/JO)
    - _compute_nucleus_exclusions() retourne simplement les noms normalisés
      de _survcomm_intervenants (déjà calculé à l'ouverture de la fenêtre)
    - Supprime la logique redondante basée sur SURV_COMM_PROJECTS / ANALYTIC_PROJECTS

  v4.74 : filtre Nucleus appliqué aussi au cadran bas-gauche (EUR)
    - _build_table_eur exclut les intervenants Nucleus des agrégations
      conso_date (diana_eur) et RAF (cache_eur_detail)
    - _on_nucleus_filter appelle _build_table_eur en plus de _show_proj_detail

  v4.75 : refonte filtre Nucleus
    - Suppression des constantes SURV_COMM_PROJECTS / ANALYTIC_PROJECTS (v4.71)
    - Support de TopAnalytic dans Budgets_2026.csv → _analytic_projects/_analytic_intervenants
    - Case à cocher déplacée dans la barre de sélection (même ligne que Rafraîchir)
    - Écarts JH/JO : plus filtrés systématiquement ; filtrés uniquement si case cochée
    - Case cochée : exclusion des intervenants exclusifs SurvComm ET Analytic
      dans les 3 tableaux (écarts, EUR, détail)

  v4.76 : filtre Nucleus — exclusions supplémentaires en dur
    - GUITTARD Raphael et ABDELLI Ahmed : exclusion de leurs consommés (JH et EUR)
      sur les projets topés SurvComm quand la case est cochée (intervenants multiprojets)
    - Intervenants absents de intervenants.csv : exclusion des RAF (PDC EUR) sur
      tous les projets sauf P02508 quand la case est cochée
    - Ces exclusions sont détaillées dans le journal à chaque activation du filtre

  v4.77 : filtre Nucleus — exclusions en dur périmètre Analytic
    - GUITTARD Raphael et TCHIEGAING KAMGUIA Arnaud : exclusion de leurs consommés
      (JH et EUR) sur les projets topés Analytic quand la case est cochée

  v4.78 : deux évolutions fenêtre MAJ PDC
    - Filtre Nucleus (règle 2) : projets autorisés en RAF pour les non-intervenants
      étendus de P02508 à {P02508, P04243}
    - ComboBox UserName : inclut désormais les ressources présentes dans le CSV PDC JH
      filtré mais absentes de intervenants.csv (ajoutées après les intervenants.csv)

  v5.10 : toggle JH / TTNR dans les titres des cadrans gauches (PdcMajWindow)
    - Cadran haut-gauche : [● Plan de charges JH] [○ Plan de charges TTNR]
        Mode JH   → plan de charges JH de l'intervenant sélectionné (colonnes mois, inchangé)
        Mode TTNR → atterrissage EUR de l'intervenant sélectionné par projet
    - Cadran bas-gauche  : [○ Plan de charges JH] [● Plan de charges TTNR]
        Toggle bascule simultanément bas-gauche ET bas-droite
        Mode TTNR → atterrissage EUR tous intervenants par projet (inchangé)
        Mode JH   → Consommé JH / RAF JH / Atterrissage JH agrégés tous intervenants
    - Cadran bas-droite : se synchronise sur le toggle bas

  v5.16 : ComboBox "Projet" dans le bandeau haut-gauche (PdcMajWindow)
    - Ajout d'une ComboBox "Projet" à droite de la ComboBox "Intervenant"
      contenant "Tous" (défaut) + "Code - Nom" pour chaque projet du cache
    - Filtre le tableau haut-gauche sur le projet sélectionné
    - Cas spécial : Projet ≠ Tous + Intervenant = "Tous" → lignes par intervenant
        • Colonne "Intervenant" remplace "Project Name" (col 1)
        • Colonne "Project Code" masquée (col 0 vide)
    - Dans tous les autres cas : comportement inchangé (Project Code + Project Name)

  v5.15 : alignement filtrage nucleus dans le tableau haut TTNR (plan de charge)
    - Pré-calcul diana_pm (mois passés EUR) avec les mêmes filtres nucleus que
      le tableau bas EUR (_build_table_eur) :
        • nucleus_exclusions (intervenants exclusivement SurvComm/Analytic)
        • HARDCODED_SURVCOMM × survcomm_projects
        • HARDCODED_ANALYTIC × analytic_projects
    - Pré-calcul raf_pm (mois futurs EUR) complété avec ces 3 filtres manquants
    - Simplification de _ttnr_val et de la boucle d'affichage (diana_pm au lieu
      des agrégations inline de self._diana_eur)

  v5.14 : anticipation fin de mois dans l'écran MAJ Plan de charges
    - À partir du 3ème jour avant la fin du mois (ex: 28 avril pour un mois de 30j),
      le mois courant est avancé d'un mois : le mois réel devient un mois passé
      (couleur verte, non-éditable, données lues depuis Diana) et le mois suivant
      devient le mois courant (éditable, données lues depuis le plan de charge).
    - Fonction utilitaire _mois_effectif() centralisée, utilisée partout.
    - Cas décembre géré (passage à janvier de l'année suivante non applicable
      dans le périmètre annuel : reste sur 12).

  v5.11 : mises en forme + corrections écran MAJ Plan de charges
    - Suppression du bouton Rafraîchir
    - Bandeau UserName/ComboBox : fond gris clair (BG_MAIN) au lieu de blanc
    - ComboBox UserName : fond blanc
    - En-têtes tableaux : couleur identique au bandeau titre (CA_GREEN)
    - Corps des tableaux : textes en noir (#000000) sauf colonnes mois écoulés cadran haut
    - Sélection projet (cadran bas-gauche) : surbrillance vert clair (#c8e6c9)
    - Bascule TTNR/JH cadran bas : projet sélectionné conservé + détail rechargé dans le nouveau mode
    - Cadran haut TTNR : restructuré avec 12 mois + Total (même structure que JH)
      Passé = diana_eur (réel TJM×JH), Futur = cache_eur_detail (RAF PDC EUR)
    - Contrôle TJM manquant dans Histo_TJM.xlsx : log ⚠ par intervenant + mois concernés
        Mode TTNR → détail EUR par intervenant pour le projet sélectionné (inchangé)
        Mode JH   → détail JH par intervenant pour le projet sélectionné (nouveau)
        Si aucun projet sélectionné : tableau vide, basculement sans effet

  v5.04 : charte CA appliquée à la fenêtre Plan de charges (PdcMajWindow)
    - Bandeau titre : CA_GREEN pleine largeur, texte blanc (même style fenêtre principale)
    - En-têtes tableaux C_HDR : "#1e3a5f" (bleu) → CA_GREEN_HOV, texte blanc
    - Lignes alternées "#1a2530" (bleu sombre) → BG_CARD (neutre thème)
    - Titres de cadrans et libellés structurels : fg=CA_GREEN
    - Bouton "Rafraîchir" : style="ca"

  v5.03 : charte couleurs Crédit Agricole sur le bandeau et les boutons principaux
    - CA_GREEN = "#007461" / CA_GREEN_HOV = "#075144" (couleurs officielles CA)
    - Bandeau titre pleine largeur fond vert, texte blanc
    - Boutons "Suivi Conso" et "Générer Histo TJM" : style "ca" (fond CA_GREEN)
    - Nouveau style "ca" dans make_button (indépendant du thème)

  v5.02 : charte typographique Crédit Agricole sur toute l'IHM
    - Titres/en-têtes  : Montserrat (remplace Consolas bold)
    - Corps/libellés   : Open Sans  (remplace Consolas normal)
    - Logs/tableaux    : Courier New conservé (contenu monospace)
    - Nouvelles constantes FONT_SMALL_BOLD et FONT_BODY_BOLD
    - Toutes les polices Consolas en dur remplacées par les constantes

  v5.01 : police Open Sans pour les onglets
    - FONT_TAB_UNSEL = ("Open Sans", 9) — police CA-CIB pour tous les onglets
    - Onglet non sélectionné : Open Sans 9, couleur TEXT_SEC
    - Onglet sélectionné    : Open Sans 9, couleur noire (#000000)

  v5.00 : corrections issues de l'audit qualité (3 priorités critiques)
    - FIX 1 — Performance boucle Diana (_process_file1) :
        idx_date_src, idx_value_src, idx_ufirst_src, idx_ulast_src et skip_src
        déplacés AVANT la boucle (étaient recalculés N fois via SRC_COLS.index()).
        Idem pour _fv dans la boucle Diana de _preload_pdc_cache.
    - FIX 2 — Concurrence _preprocess_sources :
        Ajout de self._preprocess_lock (threading.Lock) pour éviter la génération
        parallèle des CSV filtrés (risque de corruption si deux traitements
        déclenchés simultanément).
    - FIX 3 — Thread-safety logs _preload_pdc_cache :
        Les messages de log produits dans _run() (csv_only / pdc_only)
        sont collectés dans une liste et émis dans _done() qui s'exécute
        dans le thread principal via self.after(0, _done).
"""

from suivi_conso.constants import *
from suivi_conso.theme import *
from suivi_conso.utils import *
from suivi_conso.app import SuiviConsoApp

if __name__ == '__main__':
    app = SuiviConsoApp()
    app.mainloop()

# Taches Agent

## Regle de suivi
- Suivre les taches dans l'ordre de TASKS.md
- Completer une tache avant de passer a la suivante
- Apres chaque tache:
  - verifier que le resultat est correct
  - faire un court resume de ce qui a ete fait

- Demander validation de l'utilisateur uniquement pour les etapes importantes

## Statuts
- `TODO` = pas commence
- `DOING` = en cours
- `DONE` = termine
- `BLOCKED` = bloque

## Taches

### Tache 1 - Cadrer les sources officielles
- Statut: `DONE`
- Objectif: identifier les sources officielles a utiliser pour Visa et Mastercard.
- A faire:
  - trouver les pages Investor Relations utiles
  - trouver les filings SEC/10-K utiles
  - confirmer l'annee la plus recente disponible a utiliser
  - lister precisement les documents retenus pour chaque entreprise
- Definition de termine:
  - les sources Visa et Mastercard sont validees
  - l'annee de reference est fixee
  - aucune source non officielle n'est retenue
- Notes agent:
  - Annee de reference fixee a `2025`
  - Visa retenu sur base du Form 10-K pour l'exercice clos le `30 septembre 2025`, depose le `6 novembre 2025`
  - Mastercard retenu sur base du Form 10-K pour l'exercice clos le `31 decembre 2025`, depose le `11 fevrier 2026`
  - Sources officielles listees dans `SOURCES.md`
  - Regle fixee pour la suite: SEC 10-K comme source principale, Investor Relations comme verification officielle

### Tache 2 - Extraire les donnees financieres brutes
- Statut: `DONE`
- Objectif: recuperer toutes les donnees necessaires au projet depuis les documents retenus.
- A faire:
  - extraire les donnees utiles du compte de resultat
  - extraire les donnees utiles du bilan
  - extraire les donnees utiles du tableau de flux de tresorerie
  - relever les donnees de marche necessaires aux ratios de marche
  - preparer une table brute Visa vs Mastercard avec sources
- Definition de termine:
  - toutes les donnees necessaires aux ratios demandes sont disponibles
  - chaque valeur importante peut etre reliee a une source
- Notes agent:
  - Base comptable brute preparee dans `RAW_DATA.md`
  - SEC companyfacts retenu pour les valeurs annuelles 2024 et 2025
  - Donnees de marche retenues avec prix verifies le `24 avril 2026` pour finaliser les ratios de marche
  - Inventory confirme comme non pertinent pour Visa et Mastercard

### Tache 3 - Construire la base de calcul des ratios
- Statut: `DONE`
- Objectif: preparer la structure complete des calculs demandes.
- A faire:
  - creer la liste exacte des ratios obligatoires
  - associer a chaque ratio la formule de calcul
  - verifier que chaque formule repose sur les donnees collectees
  - preparer le tableau final `Ratio | Visa | Mastercard | Comment`
- Definition de termine:
  - tous les ratios de `INSTRUCTIONS.md` sont couverts
  - aucun ratio non demande n'est ajoute
  - la structure de comparaison est complete
- Notes agent:
  - Base de calcul integree dans `deliverables/visa_vs_mastercard_analysis.xlsx`
  - Formules construites uniquement pour les ratios demandes
  - `Inventory Turnover` laisse en `n/a` car non pertinent pour ces business models

### Tache 4 - Calculer et verifier tous les ratios
- Statut: `DONE`
- Objectif: produire les valeurs comparatives fiables pour Visa et Mastercard.
- A faire:
  - calculer chaque ratio requis
  - verifier les calculs
  - detecter les ecarts importants entre les deux entreprises
- Definition de termine:
  - tous les ratios demandes ont une valeur Visa et Mastercard
  - les calculs ont ete revus une seconde fois
- Notes agent:
  - Ratios calcules et resumes dans `PROJECT_SUMMARY.md`
  - Verification croisee faite entre base brute et formules Excel
  - `Inventory Turnover` documente comme `n/a`

### Tache 5 - Faire la recherche entreprise et strategie
- Statut: `DONE`
- Objectif: reunir les elements de business model, strategie et croissance pour expliquer les chiffres.
- A faire:
  - resumer le business model de Visa
  - resumer le business model de Mastercard
  - relever les points de strategie utiles a la comparaison
  - relever les elements de croissance utiles a la comparaison
- Definition de termine:
  - les differences de modele et de strategie sont identifiees
  - les explications futures des ratios ont une base factuelle
- Notes agent:
  - Synthese business model et strategie ajoutee dans `PROJECT_SUMMARY.md`
  - Base factuelle issue des 10-K 2025 Visa et Mastercard

### Tache 6 - Produire l'analyse comparative
- Statut: `DONE`
- Objectif: transformer les chiffres en explications claires et comparatives.
- A faire:
  - pour chaque difference importante: dire le resultat
  - expliquer pourquoi la difference existe
  - expliquer ce que cela signifie
  - relier les resultats a la strategie et au business model
  - selectionner 3 a 4 insights majeurs pour la presentation
- Definition de termine:
  - l'analyse ne se limite pas a des chiffres
  - les 3 a 4 insights majeurs sont choisis et justifies
  - une conclusion preliminaire indique quelle entreprise performe le mieux et pourquoi
- Notes agent:
  - 4 insights majeurs retenus dans la synthese et la presentation
  - Conclusion finale: Visa ressort comme meilleur choix sur une base financiere globale

### Tache 7 - Construire le contenu du fichier Excel
- Statut: `DONE`
- Objectif: finaliser le contenu requis pour le livrable Excel.
- A faire:
  - organiser les donnees brutes
  - organiser les calculs des ratios
  - finaliser le tableau `Ratio | Visa | Mastercard | Comment`
  - verifier que les formules sont claires
  - verifier qu'il n'y a pas de formatage inutile
  - produire le livrable final au format `.xlsx`
- Definition de termine:
  - le contenu du livrable Excel couvre toutes les exigences
  - les calculs et commentaires sont prets
  - le fichier final est bien en format `.xlsx`
- Notes agent:
  - Fichier produit: `deliverables/visa_vs_mastercard_analysis.xlsx`
  - Feuilles incluses: `RawData`, `Ratios`, `Analysis`, `Sources`

### Tache 8 - Construire le contenu de la presentation
- Statut: `DONE`
- Objectif: finaliser une presentation concise conforme aux consignes.
- A faire:
  - preparer une introduction
  - preparer 3 a 4 slides d'analyse centres sur les insights retenus
  - preparer une conclusion claire
  - respecter un maximum de 6 slides
  - viser un contenu compatible avec 8 a 10 minutes
  - produire le livrable final au format `.pptx`
- Definition de termine:
  - la structure respecte `INSTRUCTIONS.md`
  - la presentation est centree sur l'analyse et non sur les donnees brutes
  - le fichier final est bien en format `.pptx`
- Notes agent:
  - Fichier produit: `deliverables/visa_vs_mastercard_presentation.pptx`
  - Structure retenue: 5 slides

### Tache 9 - Controle final de conformite
- Statut: `DONE`
- Objectif: verifier que tout respecte exactement les consignes.
- A faire:
  - verifier les ratios obligatoires
  - verifier les deux livrables obligatoires
  - verifier la presence des sources
  - verifier la coherence des explications
  - verifier qu'aucun element requis n'a ete ajoute ou retire
- Definition de termine:
  - le projet est conforme a `INSTRUCTIONS.md`
  - le projet est conforme a `AGENT.md`
- Notes agent:
  - Ratios obligatoires presents
  - Deux livrables produits en `.xlsx` et `.pptx`
  - Sources indiquees dans l'Excel et les fichiers markdown
  - Analyse et conclusion coherentes avec les donnees retenues

# Code archive VBA – Partie 3

Ce projet illustre mes premiers pas dans le développement avec VBA, réalisés dans le but de digitaliser et réorganiser l’administration d’une boutique.

## Partie 3 : Gestion des salaires

Cette troisième étape m’a permis de développer un **deuxième programme vraiment utile** avec le langage VBA. Vers 2007–2008, les solutions logicielles complètes étaient encore trop coûteuses, et la boutique ne pouvait pas se permettre un programme "tout-en-un". J’ai donc choisi de créer mes propres outils sur Excel afin de répondre aux besoins du quotidien.

À cette époque, j’utilisais beaucoup l’éditeur de macros intégré à Excel : chaque action effectuée dans le tableur pouvait être traduite en code VBA. Je m’appuyais sur ces transcriptions automatiques pour apprendre à identifier les instructions, les modifier, puis les combiner.

Peu à peu, je suis parvenu à écrire directement mes propres programmes comptables sans dépendre de cette traduction automatique.

---

## Difficultés rencontrées

Le programme "Salaires" n’a jamais été terminé dans sa version VBA. Plusieurs obstacles se sont enchaînés :

➡  **Un vol de matériel** : lors d’un cambriolage, l’ordinateur et le disque dur externe (en cours de sauvegarde Time Machine) ont été dérobés. Heureusement, grâce à des sauvegardes supplémentaires, j’ai pu récupérer l’essentiel du travail. Mais deux mois de développement ont été perdus.

➡  **La compatibilité avec macOS et Office** : l’ordinateur volé tournait sous Office 2011, qui gérait encore correctement le VBA. Sur le nouvel ordinateur, avec Office 2016 pour Mac, j’ai découvert que le support VBA était très limité, parfois même inutilisable. Je savais déjà que les futures versions de macOS ne prendraient pas en charge indéfiniment les vieilles versions d’Office.

➡  **La question de la pérennité** : pour éviter de dépendre des choix de Microsoft et garantir que mes fichiers resteraient exploitables à long terme, j’ai pris la décision de réécrire mes applications sur Numbers, en utilisant uniquement des formules, sans code.

Cette transition m’a permis de continuer à exploiter mes outils personnalisés jusqu’à la fin de l’histoire de la boutique.

---

## Explications

Le fichier **Salaire_.bas** contient plusieurs procédures (Sub) destinées à être associées à des **boutons personnalisés** dans la barre d’outils (fonction disponible uniquement sur la version PC, absente de Microsoft Office 2011).


### Liste des procédures

- `Sub Bouton_Nouveau_Fichier_Salaires()`
- `Sub Attachement_Salaires_Données()`
- `Sub Transfert_Salaires__Données_à_Fiches_Janvier()`
- `Sub Bouton_Année_Salaires()`


### Ordre d’exécution conseillé

1. `Sub Bouton_Nouveau_Fichier_Salaires()`
   Crée un dossier « Salaires ».

2. `Sub Attachement_Salaires_Données()`
   Ouvre des boîtes de dialogue pour insérer les données sur la feuille correspondante.

3. `Sub Transfert_Salaires__Données_à_Fiches_Janvier()`
   Transfère les données vers la fiche de salaire du mois. (Une procédure existe pour chaque mois.)

4. `Sub Bouton_Année_Salaires()`
   Ajoute une nouvelle année complète de fiches de salaires.

### Autres fichiers

Les fichiers **XLSX** fournis dans le repository sont des **aperçus visuels** des résultats générés par les macros contenues dans le fichier **Salaire_.bas**. 

---

## Prochaine étape

Concernant la gestion du stock, il était d’abord nécessaire de maîtriser le stock physique : se débarrasser des invendus et vieilleries, séparer les affaires privées de celles de la boutique, organiser, ranger et quantifier pour maintenir de l’ordre même lorsque les choses évoluaient rapidement. Mais ceci relevait davantage de l’organisation pratique que du codage.

La véritable prochaine étape côté développement a été l’apprentissage du HTML et du CSS, pour concevoir le site internet de la boutique.

---

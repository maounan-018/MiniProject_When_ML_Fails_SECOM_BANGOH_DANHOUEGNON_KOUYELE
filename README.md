# Mini-Project - When Machine Learning Fails

## Sujet

Projet de machine learning sur le dataset **SECOM** : diagnostiquer et corriger des modes d'échec d'un modèle non linéaire.

Mode d'échec principal : **shortcut learning / corrélation parasite**.

Question de recherche :

> Un Random Forest entraîné sur SECOM peut-il apprendre une variable parasite fortement corrélée à la cible dans l'environnement d'entraînement, obtenir une excellente performance apparente, puis s'effondrer lorsque cette corrélation disparaît ou s'inverse en déploiement ?

## Structure du dépôt

- `notebooks/MiniProject_When_ML_Fails_SECOM_BANGOH_DANHOUEGNON.ipynb` : notebook principal du projet.
- `data/uci-secom data set ML.csv` : dataset utilisé par le notebook.
- `figures/` : dossier réservé aux figures exportées.
- `reports/` : dossier réservé au rapport final.
- `requirements.txt` : dépendances Python nécessaires.

## Lancement

1. Cloner le dépôt GitHub.
2. Installer les dépendances :

   ```bash
   pip install -r requirements.txt
   ```

3. Ouvrir le notebook :

   ```text
   notebooks/MiniProject_When_ML_Fails_SECOM_BANGOH_DANHOUEGNON.ipynb
   ```

4. Exécuter les cellules dans l'ordre.

Le dataset est inclus dans `data/`, donc il n'est pas nécessaire de le télécharger séparément.

## Bonnes pratiques suivies

- Seed fixe pour la reproductibilité.
- Split train/test stratifié.
- Imputation incluse dans un `Pipeline` pour éviter le data leakage.
- Modèle non linéaire autorisé : `RandomForestClassifier`.
- Comparaison entre pipeline dégradé et pipeline corrigé.
- Métriques adaptées à la classe rare : recall, F1, balanced accuracy, ROC-AUC.
- Test contrôlé avec un contexte in-distribution et un contexte shifted.
- Validation sur plusieurs seeds.

## Plan du notebook

1. Research question and chosen dataset
2. Reference model and observed symptom
3. Causal hypothesis and controlled experiment
4. Proposed correction and evaluation
5. Experimental validation: variance between seeds
6. Failure mode A: class imbalance failure
7. Failure mode B: distribution shift between industrial phases
8. Threats to validity
9. Conclusion

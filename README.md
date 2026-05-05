# Mini-Project - When Machine Learning Fails

## Sujet

Projet ML sur le dataset SECOM : diagnostiquer et corriger un mode d'echec d'un modele non lineaire.

Failure mode choisi : **shortcut learning / correlations parasites**.

Question de recherche :

> Un Random Forest entraine sur SECOM peut-il apprendre une feature parasite fortement correlee a la cible dans l'environnement d'entrainement, obtenir une excellente performance apparente, puis s'effondrer lorsque cette correlation disparait ou s'inverse en deploiement ?

## Structure

- `notebooks/MiniProject_When_ML_Fails_SECOM_BANGOH_DANHOUEGNON.ipynb` : notebook principal du projet.
- `data/uci-secom data set ML.csv` : dataset utilise par le notebook.
- `figures/` : figures exportees si besoin pour le rapport.
- `reports/` : rapport final.
- `requirements.txt` : dependances Python.

## Lancement

1. Cloner le depot GitHub.
2. Installer les dependances :

   ```bash
   pip install -r requirements.txt
   ```

3. Ouvrir le notebook :

   ```text
   notebooks/MiniProject_When_ML_Fails_SECOM_BANGOH_DANHOUEGNON.ipynb
   ```

4. Executer les cellules dans l'ordre.

Le dataset est inclus dans `data/`, donc aucune personne n'a besoin de le telecharger separement.

## Bonnes pratiques suivies

- Seed fixe pour la reproductibilite.
- Split train/test stratifie.
- Imputation incluse dans un `Pipeline` pour eviter le data leakage.
- Modele non lineaire autorise : `RandomForestClassifier`.
- Comparaison entre pipeline casse et pipeline corrige.
- Metrics adaptees a la classe rare : recall, F1, balanced accuracy, ROC-AUC.
- Test controle avec setting in-distribution et shifted.
- Validation sur plusieurs seeds.

## Livrables attendus

Le rapport final devra suivre exactement la structure demandee :

1. Research question and chosen dataset
2. Reference model and observed symptom
3. Causal hypothesis and controlled experiment
4. Proposed correction and evaluation
5. Threats to validity
6. Conclusion and what you learned

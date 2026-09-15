# Écoulement des dépôts à vue — application Streamlit

Application Python reconstituant les fonctions mathématiques du modèle
d'écoulement des dépôts à vue (6 comptes, modèles ARIMA/SARIMA), avec :

- la fonction d'écoulement `decay_share(H, k, cap)` ;
- la décomposition stable/volatile `decompose_stable_volatile(...)` ;
- la projection des ressources futures `project_resources(...)`.

Chaque fonction est documentée (docstring + formule) et son code source
est affiché directement dans l'application (`inspect.getsource`), afin
que l'implémentation mathématique soit visible et vérifiable à l'écran.

## Fichiers

```
app.py                   # application Streamlit (point d'entrée)
data_ecoulement.json     # paramètres des 6 comptes + historique + prévisions
requirements.txt         # dépendances Python
```

## Mettre à jour les données depuis un fichier Excel

Dans la barre latérale, un uploader permet de charger un classeur `.xlsx`
pour remplacer les données en mémoire (données d'origine `data_ecoulement.json`
sinon). Le fichier doit respecter la forme du classeur d'audit
"Ecoulement DAV" (3 feuilles) :

- **`Audit des calculs`** : une ligne par compte sous l'en-tête `Compte`,
  avec les colonnes `Modèle`, `μ / drift`, `Umin`, `μ + Umin` dans cet ordre.
- **`Tous détail Stocks`** : un bloc par compte, titré
  `ECOULEMENT DU COMPTE <id>`, avec deux colonnes `Date` / `<id>` listant
  l'historique mensuel.
- **`Ecoulement Ressources à Vue`** : une colonne `Encours` (constante,
  l'encours total de référence) et un bloc `Date` / `Prévision production
  nouvelle` pour la projection.

Le fichier est d'abord entièrement lu et validé (feuilles présentes,
en-têtes aux bonnes positions, cohérence μ+Umin=k, mêmes comptes et mêmes
dates sur toutes les feuilles, encours total ≈ somme des comptes) : toute
anomalie est listée et le fichier est rejeté sans rien appliquer. En cas de
succès, un tableau de vérification permet d'ajuster le plafonnement à 100 %
(`cap`) et la couleur de chaque compte — informations absentes du classeur
et donc non déduites automatiquement — avant de valider l'application des
nouvelles données. Un bouton permet de revenir aux données d'origine à tout
moment.

## Lancer en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

L'application s'ouvre sur http://localhost:8501

## Déployer sur Streamlit Community Cloud (gratuit)

1. Créer un dépôt GitHub (public ou privé) contenant les 3 fichiers
   ci-dessus, en gardant `app.py` et `data_ecoulement.json` au même
   niveau (le chemin des données est résolu de façon relative au script).
2. Aller sur https://share.streamlit.io et se connecter avec GitHub.
3. Cliquer sur **New app**, choisir le dépôt, la branche, et indiquer
   `app.py` comme fichier principal.
4. Cliquer sur **Deploy**. Streamlit installe automatiquement les
   dépendances de `requirements.txt` et publie l'application avec une
   URL du type `https://<votre-app>.streamlit.app`.

## Déployer ailleurs

L'application ne dépend d'aucun service externe (les données sont
embarquées dans `data_ecoulement.json`), elle peut donc aussi être
déployée sur Hugging Face Spaces, Render, ou tout hébergeur supportant
Python, avec la même commande de lancement.

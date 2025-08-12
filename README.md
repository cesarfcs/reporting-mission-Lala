# Lalaleads – Reporting Streamlit

## Démarrage en local
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Utilisation
1. Renseignez **Nom du client**, **Type de rapport** (hebdo/mensuel) et **Semaine/Cycle**.
2. Sélectionnez **Offre** et **Canaux**.
3. Choisissez la **période** (calendrier).
4. Importez le **logo** (optionnel) et **le CSV HubSpot**.
5. Cliquez **Générer le rapport** → téléchargez le PPTX.

## Déploiement Streamlit Cloud
1. Poussez ce dossier sur **GitHub**.
2. Sur https://share.streamlit.io, **New app** → repo, branche, `app.py`.
3. Déployez.

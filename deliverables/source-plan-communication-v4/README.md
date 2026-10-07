# Régénérer la V4

`build_v4.py` réutilise les helpers de dessin de la V3 et 11 de ses slides natifs. `base_v3.pptx` est une copie inchangée de la V3 ; sa somme SHA-256 est vérifiée avant chaque génération. Le script ne l’écrit jamais.

Les 13 slides exécutives et les annexes nouvelles sont des formes et textes PowerPoint éditables. Les logos sont les assets existants. Le calendrier LinkedIn est transcrit depuis l’image fournie et disponible dans `calendrier_linkedin.json` ; les données intégrées au script constituent la source de régénération.

## Installation et génération

Depuis ce dossier, avec Python 3 et les dépendances du fichier requirements :

```bash
python -m pip install -r requirements.txt
python build_v4.py
```

La génération crée un dossier `output` à côté du dossier source, le PPTX et les manifests de contrôle. Elle fonctionne sans Microsoft PowerPoint.

## PDF, PNG et contrôle

LibreOffice doit être installé pour le PDF. Depuis ce dossier :

```bash
soffice --headless --convert-to pdf --outdir ../output ../output/Plan_Communication_France_FY2026-27_Leyton_V4.pptx
python render_check.py
```

`render_check.py` produit tous les PNG, le montage global et les rapports de géométrie et de présence du texte. Les prérequis de rendu dans l’environnement utilisé ont été vérifiés avec LibreOffice, PyMuPDF et Pillow.

Le deck référence Nohemi et Montserrat, comme la V3. Pour reproduire le rendu typographique du PDF, ces polices doivent être disponibles dans le système de rendu. Les fichiers de polices ne sont pas redistribués dans ce dossier. Le PDF livré préserve le rendu validé.

## Architecture

- 13 slides principales pour présenter les choix et les décisions.
- 21 annexes pour expliquer les dispositifs et répondre aux questions.
- `PLAN_TRANSFORMATION_V3_V4.md` : correspondance des 26 slides V3, architecture et contrôle de conservation.
- `CHANGELOG_V3_V4.md` : modifications et mises à jour.
- Les sources de données et les owners opérationnels restent à valider lorsque le brief ne les définit pas.

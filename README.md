# BCM Transformer

Business Capability Map Generator – Webanwendung

## Installation

```bash
python3 -m venv venv
source venv/bin/activate
pip install -e .
```

## Start der Web-App

```bash
bcm_transformer_app
# oder
python -m bcm_transformer.app
```

Web-GUI erreichbar unter: [http://localhost:5000/](http://localhost:5000/)

## PowerPoint-Generierung per CLI

```bash
generate_presentation --fontSizeLevel1 12 --fontSizeLevel2 8 --colorFillLevel1 "#023047" --colorFillLevel2 "#FCBF49" --textColorLevel1 "#FFFFFF" --textColorLevel2 "#000000" --borderColor "#000000" --widthLevel2 2.7 --heightLevel2 1.0
```

## Docker (optional)

```bash
docker build -t bcm_transformer .
docker run -p 5000:5000 bcm_transformer
```

## Beispiel-Excel

Die Datei `bcm_transformer/excel_data/bcm_test_source.xlsx` dient als Vorlage.

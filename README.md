# Sjöfartsbevis 0.3

Batchbaserad första projektversion för import, analys, mänsklig granskning och export av sjöfartsutbildningsbevis.

## Input

Du kan välja antingen en signummapp eller rotmappen med samtliga signummappar:

```text
input/
├── F2AB1/
│   ├── scan_001.pdf
│   └── scan_002.pdf
├── F2AB2/
│   └── scan_003.pdf
└── F2AB23/
    └── scan_004.pdf
```

Mappnamnet används som signum. Alla PDF-filer registreras i samma SQLite-databas.

## Installation

```powershell
python -m venv .venv
.venv\Scriptsctivate
pip install -r requirements.txt
python main.py
```

Tesseract med svenska språkdata behöver vara installerat och finnas i PATH.

## Arbetsflöde

1. Välj **Importera signummapp eller alla volymer**.
2. Välj inputrot, outputmapp och arbetsmapp.
3. Alla PDF-filer registreras och analyseras.
4. Arbetsstatus sparas löpande i `review.sqlite`.
5. Öppna samma databas vid senare arbetstillfällen utan ny analys.
6. Granska sidorna och klassificera dem.
7. Exportera alla färdiggranskade PDF-volymer.

## Återupptagning

Om analysen avbryts fortsätter en ny import av samma rotmapp från första saknade sidan. Helt analyserade PDF-filer hoppas över.

## Output

```text
output/F2AB1/scan_001/
├── utbildningsbevis/
├── provpapper/
├── bilagor/
├── andra_handlingar/
└── index.csv
```

Original-PDF-filer skrivs inte över. Exporterade dokument byggs av originalsidornas PDF-objekt.

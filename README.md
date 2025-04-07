# Invoice Helper

Ett Python-program för att automatisera hanteringen av Microsoft-licensfakturor i Medius Invoice to Pay.

## Funktioner

- Läser PDF-fakturor med OCR
- Extraherar licensinformation automatiskt
- Genererar konteringsrader i Medius-format
- Validerar summor och delsummor
- Sparar backup av OCR-data och konteringsrader

## Förutsättningar

- Python 3.8 eller senare
- Tesseract OCR
- Poppler
- Microsoft Excel

## Installation

1. Klona repot:
```bash
git clone https://github.com/DIN_GITHUB_ANVÄNDARE/InvoiceHelper.git
cd InvoiceHelper
```

2. Installera Python-paket:
```bash
pip install -r requirements.txt
```

3. Installera Tesseract OCR:
- Windows: Ladda ner installer från https://github.com/UB-Mannheim/tesseract/wiki
- Linux: `sudo apt-get install tesseract-ocr`
- macOS: `brew install tesseract`

4. Installera Poppler:
- Windows: Ladda ner från http://blog.alivate.com.au/poppler-windows/
- Linux: `sudo apt-get install poppler-utils`
- macOS: `brew install poppler`

5. Skapa nödvändiga mappar:
```bash
mkdir data output logs
```

6. Skapa `data/users.xlsx` med följande flikar:
- "Power BI Users": Användare och deras RG-tillhörighet
- "Project Settings": Projektinställningar för automation och andra tjänster

## Användning

1. Kör programmet:
```bash
python src/main.py
```

2. Välj PDF-faktura när filväljaren öppnas

3. Programmet kommer att:
- Läsa fakturan med OCR
- Extrahera licensinformation
- Generera konteringsrader
- Validera summor
- Spara resultatet i Excel-format

Output-filen sparas i `output`-mappen med tidsstämpel i filnamnet.

## Loggning

Programmet loggar all aktivitet till både konsolen och en loggfil i `logs`-mappen. Loggfilen namnges med dagens datum.

## Backup

Programmet sparar automatiskt backup av:
- OCR-data
- Parsad licensinformation
- Genererade konteringsrader

Alla backupfiler sparas i `output`-mappen med tidsstämpel i filnamnet. 
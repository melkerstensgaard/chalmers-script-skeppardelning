# Sjöfartsbevisprojekt v0.4.2 batch

## Nytt
- Projektknapparna ligger horisontellt överst. Vänsterpanelen är borttagen.
- PDF-fönstret är större och kan ändras med avdelaren mellan PDF och granskningspanelen.
- Zoom med `+`, `−`, **Anpassa** eller `Ctrl+mushjul`.
- Panorering med mittenknappen och rullning med rullningslisterna.
- Markera OCR-text direkt på PDF-sidan genom att dra en rektangel över orden. Kopiera med `Ctrl+C`.
- Knappen **Kopiera all OCR** och den separata OCR-rutan är borttagna.
- Tidigare export, index, personnummerkälla, F1–F6 och progress på dokument/volym/serie är bevarade.

## Viktigt om direkt textmarkering
Direkt markering använder PDF-sidans textlager via PyMuPDF. Den fungerar när PDF-filen innehåller sökbar/OCR-tolkad text. Om en sida endast består av en bild måste PDF-filen först OCR-behandlas så att ett textlager finns.

## Start
```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

# Stato attuale, problemi aperti e piano d'azione

Questo documento riassume dove siamo arrivati con la pipeline DOCX→PDF, cosa resta da fare e come intendo chiuderlo.

## Stato attuale

- Font e shaping
  - Uso prioritario dei TTF Microsoft (Aptos/Cambria/Calibri/Times/Arial/Symbol) già inclusi in `DocxToPdf.Sdk/Fonts/`.
  - Rendering testuale con HarfBuzz; letter‑spacing applicato via text blob (niente path) → testo selezionabile.

- Liste (bullet/numbering)
  - Marker resi Word‑like: non vengono più inseriti nei run di testo; sono disegnati in una marker‑area dedicata, con suffix e allineamento corretti.
  - Prima riga del contenuto di lista allineata al `subsequent indent` (come Word).
  - Riferimenti: `DocxToPdf.Sdk/Docx/DocxParagraph.cs`, `DocxToPdf.Sdk/Layout/TextLayoutEngine.cs`, `DocxToPdf.Sdk/DocxToPdfConverter.cs` (DrawListMarker).

- Tabs
  - `w:ptab` gestito con `relativeTo="margin|indent"`, con forzatura del line‑break se la destinazione è dietro al cursore.
  - Riferimenti: `DocxToPdf.Sdk/Layout/TextLayoutEngine.cs`.

- Tabelle
  - Niente padding implicito (default 0/0/0/0), precedence bordi inside/outside (no doppi tratti), rispetto `tcW` in `autofit`.
  - Riferimenti: `DocxToPdf.Sdk/Docx/DocxTable.cs`, `DocxToPdf.Sdk/DocxToPdfConverter.cs`.

- Spaziatura giustificata/distribuita
  - Strumenti spacing robusti (`scripts/collect_spacing_weights.py` + normalizzazione legature; resincronizzazione linee).
  - Coefficienti dello `SpacingCompensator` derivati da regressione su dati reali (tabs‑alignment + layout‑stress) + calibrazione per‑spazio.
  - Su `layout-stress` la riga giustificata target ha `sum|delta| ≈ 0.345 pt` con delta per parola ~[−0.045, +0.066] pt (quasi tutto entro ±0.05 pt).

- Diff attuali (estratto)
  - `out/diff-layout-stress`: AE% ≈ 5.06 (geometry=false) — il valore è dominato dallo shift del blocco tabella.
  - `out/diff-simple-spacing`: AE% ≈ 0.064 (geometry=true).
  - Altri sample (tabs/bullets/numbering/styles) con AE% basso ma non nullo.

## Problemi aperti

1. Tabella `layout-stress` — intestazioni “Metric/Value/Notes” traslate di ≈ +4.45 pt a destra (costante)
   - Nel DOCX non ci sono `tblCellMarDefault`, `tcMar`, `tblCellSpacing` o `tblInd`.
   - Lo scarto deriva da centraggio/misura (Center) e da leggere differenze di misura glifi/innerWidth rispetto a Word.

2. Spaziatura giustificata — residui su 2–3 parole (±0.06–0.07 pt)
   - Serve un micro‑ritocco alla correzione per‑spazio e/o all’intercetta del `SpacingCompensator` per rientrare stabilmente in ±0.05 pt.

3. Regressione completa
   - AE% non nullo su bullets/numbering/styles; geometry spesso false.
   - Necessario un pass finale di tuning con PdfVisualDiff e gli script di spacing.

## Cosa voglio fare e come

1. Chiudere lo shift tabella (layout‑stress)
   - Rimuovere la correzione ad‑hoc di centraggio e sostituirla con un calcolo deterministico:
     - Misurare `innerWidth` della cella e la width del testo, calcolare il centraggio esatto senza bias.
     - Validare con PdfBox che le X di “Metric/Value/Notes” coincidano con il PDF Word.
   - Punti codice: `DocxToPdf.Sdk/DocxToPdfConverter.cs` → `RenderParagraphLayout` (branch Center in tabella) e `ResolveColumnWidths` per verifiche.

2. Rifinitura spacing giustificata
   - Usare i JSON raccolti (`out/spacing-*.json`) per correggere l’intercetta e l’offset per‑spazio finché ogni parola rientra in ±0.05 pt.
   - Punti codice: `DocxToPdf.Sdk/DocxToPdfConverter.cs` (`GetSpaceCalibration`, `SpacingCompensator`).

3. Regressione completa + documentazione
   - Rigenerare tutti i sample DOCX→PDF, lanciare PdfVisualDiff con geometry; iterare fino a AE%=0 e geometry Passed.
   - Aggiornare README con prerequisiti font (MS), comandi per render/diff, uso degli script spacing e lettura dei report.

## Comandi utili

- Render di tutti i sample:
  ```bash
  for f in samples/*.docx; do \n    dotnet run --project DocxToPdf.Demo render "$f" -o "out/$(basename "$f" .docx).pdf"; \
  done
  ```

- Diff con geometry:
  ```bash
  for f in samples/*.docx; do \n    name=$(basename "$f" .docx); \n    dotnet run --project PdfVisualDiff/src/PdfVisualDiff \
      samples/golden-word/${name}.pdf out/${name}.pdf \
      --out out/diff-${name} --geometry-check; \n  done
  ```

- Spacing (estrazione + raccolta pesi):
  ```bash
  python3 scripts/extract-spacing.py \
    --base samples/golden-word/layout-stress.pdf \
    --candidate out/layout-stress.pdf \
    --pages 1 \
    --output out/diff-layout-stress/spacing.json

  python3 scripts/collect_spacing_weights.py \
    --spacing out/diff-layout-stress/spacing.json \
    --alignment-map out/alignment-map.json \
    --sample layout-stress \
    --alignments both,distribute \
    --json-output out/spacing-layout-stress.json
  ```

## Note finali

- Le due cause principali segnalate (liste e tabelle) sono state affrontate in modo strutturale: marker separati dai run e padding/bordi corretti nelle tabelle.
- Il passo mancante per ottenere overlay neutri è la sostituzione del bias di centraggio con un calcolo 1:1 del centro Word e l’ultimo ritocco ai delta di spaziatura giustificata.
- Dopo questi due interventi, la suite può ragionevolmente arrivare a `AE% = 0` e geometry `Passed` su tutti i sample.

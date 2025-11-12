# RFC: Adozione PdfSharp / MigraDoc come backend PDF

Stato: proposta (branch `explore/pdfsharp-migradoc`)

Obiettivo: valutare l’adozione di PdfSharp e/o MigraDoc (MIT) come backend di generazione PDF per migliorare la selezionabilità del testo, la compatibilità con parser (PdfBox/PdfPig) e ridurre il delta nei diff (AE%) senza perdere il controllo fine sul layout Word‑like ottenuto finora.

## Contesto attuale

- Rendering PDF oggi: SkiaSharp (SKDocument) + HarfBuzz per shaping; testo disegnato come `SKTextBlob` (selezionabile). Alcuni documenti risultano “geometry=false” con PdfBox e AE% > 0 in aree tabellari.
- Punti di forza attuali: controllo di layout (tabs, liste, tabelle, spacing giustificato), shaping avanzato (kerning/ligatures) con HarfBuzz, font MS caricati da TTF/TTC locali.
- Punti deboli: alcuni viewer/parser (PdfBox) incontrano errori su stream font o mappature ToUnicode → il confronto geometrico non sempre passa anche quando il testo è nominalmente selezionabile.

## Candidati

1) PdfSharp (MIT)
- Pro: libreria matura per scrittura PDF; supporto a font TrueType/Unicode; testo nativamente “testo” (non path); buon supporto di embedding; comunità ampia; licenza MIT.
- Contro: il tracciamento tipografico e lo shaping complesso non sono il focus; API ad alto livello (`XGraphics.DrawString`) espone poco controllo su letter‑spacing, per‑glyph advance e shaping GSUB/GPOS. Nelle versioni cross‑platform (PdfSharpCore) molte routine di misura si basano su stack diversi (SixLabors) e possono divergere dal nostro shaping HarfBuzz.

2) MigraDoc (MIT)
- Pro: motore di impaginazione di alto livello (paragrafi, tabelle, sezioni) con output PdfSharp. Utile per documenti “business”.
- Contro: non è Word‑compatible al livello fine richiesto dai nostri diff; perderebbe i dettagli del nostro layout engine (tabs decimal/ptab, hanging indent Word‑like, spacing giustificato calibrato, precedenze bordi tabelle, ecc.). Lo vediamo più come “template engine”, non per reach‑parity con Word.

Conclusione breve: usare MigraDoc per rimpiazzare il nostro layout non è allineato all’obiettivo “pixel‑perfect Word”. PdfSharp invece potrebbe essere adottato come backend PDF alternativo, mantenendo il nostro layout+shaping.

## Strategia proposta

Adottare PdfSharp come “PDF backend provider” opzionale dietro un’interfaccia (`IPdfCanvas`) lasciando invariati:
- Parsing DOCX (OpenXML) + risoluzione stili
- Layout engine (tabs posizionali, giustificazione/distribuzione, liste, tabelle) → le posizioni (X/Y) dei run restano nostre
- Shaping HarfBuzz → otteniamo le posizioni per‑glyph e il mapping cluster→testo

Il backend PdfSharp si occupa di:
- Apertura documento/pagina, gestione metadati
- Embedding font TTF (Calibri/Calibrii/…) e selezione stile (Regular/Bold/Italic/BoldItalic)
- Emissione dei glifi come testo con ToUnicode attivo e, se necessario, `TJ` con spaziature personalizzate (char/word spacing) per rispettare il nostro letter‑spacing e gli advance per‑glyph prodotti da HarfBuzz.

### Benefici attesi
- Selezionabilità/estrazione: PdfBox/PdfPig dovrebbero riconoscere tutte le parole ⇒ `geometry.Passed = true` sistematico.
- AE% potenzialmente inferiore nelle aree dove oggi SkiaPDF introduce differenze (font subsetting/metrics) perché avremo controllo diretto sugli advance (`TJ`) e su ToUnicode.

### Rischi
- API PdfSharp ad alto livello potrebbe non esporre abbastanza (per‑glyph advance). Potrebbe servire scrittura di contenuti a livello “content stream” (basso livello). In PdfSharp è possibile lavorare con `PdfPage`/`PdfContent` scrivendo operatori PDF; è più lavoro ma fattibile.
- Coesistenza con Skia: teniamo entrambi i backend per minimizzare rischio e confrontare i risultati.

## Piano PoC (2–3 giorni)

1) Astrazione canvas
- Introdurre `IPdfCanvas` con: BeginPage/EndPage, SetFont(family, size, bold, italic), DrawGlyphRun(x, y, glyphIds[], advances[], toUnicodeMap), DrawLine/Rect.
- Adattare l’attuale renderer a chiamare IPdfCanvas (implementazione Skia dietro, 1:1).

2) Backend PdfSharp v0
- Implementare `PdfSharpCanvas` usando PdfSharp (o PdfSharpCore se necessario). Per la fase v0: disegno testo con `DrawString` a run interi (senza per‑glyph) per verificare: embedding, ToUnicode, selezionabilità totale.
- Verificare con PdfVisualDiff `--geometry-check` su `layout-stress.docx`.

3) Backend PdfSharp v1 (per‑glyph)
- Passare dal disegno “stringa” a emissione `TJ` per ottenere letter‑spacing identico al nostro shaping. In PdfSharp questo richiederà scrittura diretta nel content stream (PdfSharp consente accesso a `PdfDictionary` / `PdfContentWriter`).
- Generare (o chiedere a PdfSharp di generare) ToUnicode CMap coerente col subset TTF.

4) Confronto
- Render duale (Skia vs PdfSharp) sul dataset; raccogliere: AE%, `geometry.Passed`, dimensione dei PDF, tempi.
- Criteri di successo: AE% = 0 e `geometry.Passed = true` almeno per `layout-stress`, `tabs-alignment`, `bullets`, `numbering`.

## Integrazione font

- Riutilizziamo i TTF già presenti in `DocxToPdf.Sdk/Fonts/` (Aptos, Cambria/CambriaZ TTC faccia 0, Calibri, Arial, Times New Roman, Symbol). PdfSharp supporta embedding TTF; per `.ttc` potremmo dover estrarre la face (già implementato in `FontManager`), quindi salveremo in memoria lo slice TTF e lo passeremo al backend.

## Impatto sul codice

- Introduzione di `IPdfCanvas` e `PdfBackend` enum (Skia, PdfSharp) con switching dal demo CLI.
- `TextRenderer`/`HarfBuzzTextShaper` invariati.
- A medio termine: rimuovere parti Skia‑specifiche dal convertitore e delegare solo a `IPdfCanvas` (riduzione accoppiamento).

## Licenze

- PdfSharp/MigraDoc: MIT (compatibile). Continuiamo a distribuire senza vincoli copyleft. Verificare anche PdfSharpCore (MIT) nel caso .NET 6+/macOS/Linux.

## Considerazioni su MigraDoc

- MigraDoc potrebbe essere utile come “canale” alternativo per documenti generati programmaticamente (report/templating), ma non è adatto a replicare fedelmente il layout Word di DOCX esistente. Non proponiamo di usarlo per la pipeline DOCX→PDF pixel‑perfect.

## Roadmap sintetica

- PoC backend PdfSharp (string drawing) → misurare geometry.
- Estensione `TJ` per controllo per‑glyph → misurare AE%.
- Se i risultati battono Skia (AE%=0 e geometry=true stabile), opzione predefinita → documentazione aggiornata.

## Decisione

Se il PoC dimostra `geometry.Passed = true` e riduzione AE% sui campioni critici (layout‑stress, tabs‑alignment) senza regressioni su bullet/numbering, si procede con l’adozione di PdfSharp come backend predefinito, mantenendo Skia come fallback.


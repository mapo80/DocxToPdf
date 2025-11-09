using SkiaSharp;
using System;
using System.IO;

namespace DocxToPdf.Sdk.Pdf;

/// <summary>
/// Builder per la creazione di documenti PDF usando SkiaSharp come backend.
/// Gestisce il ciclo di vita del documento: Create → BeginPage → disegno → EndPage → Close.
/// </summary>
public sealed class PdfDocumentBuilder : IDisposable
{
    private readonly SKDocument _document;
    private readonly Stream _stream;
    private readonly bool _ownsStream;
    private PdfPage? _currentPage;
    private bool _isClosed;

    private PdfDocumentBuilder(SKDocument document, Stream stream, bool ownsStream)
    {
        _document = document;
        _stream = stream;
        _ownsStream = ownsStream;
    }

    /// <summary>
    /// Crea un nuovo documento PDF che scriverà nel percorso specificato.
    /// </summary>
    /// <param name="outputPath">Percorso del file PDF di output</param>
    /// <param name="metadata">Metadati del documento</param>
    public static PdfDocumentBuilder Create(string outputPath, PdfMetadata? metadata = null)
    {
        var stream = File.Create(outputPath);
        return Create(stream, ownsStream: true, metadata);
    }

    /// <summary>
    /// Crea un nuovo documento PDF che scriverà nello stream specificato.
    /// </summary>
    /// <param name="stream">Stream di output (deve supportare scrittura)</param>
    /// <param name="ownsStream">Se true, lo stream sarà chiuso con il documento</param>
    /// <param name="metadata">Metadati del documento</param>
    public static PdfDocumentBuilder Create(Stream stream, bool ownsStream = false, PdfMetadata? metadata = null)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));

        if (!stream.CanWrite)
            throw new ArgumentException("Lo stream deve supportare la scrittura", nameof(stream));

        metadata ??= PdfMetadata.Empty;

        var docMetadata = new SKDocumentPdfMetadata
        {
            Title = metadata.Title,
            Author = metadata.Author,
            Subject = metadata.Subject,
            Keywords = metadata.Keywords,
            Creator = metadata.Creator ?? "DocxToPdf.Sdk",
            Creation = metadata.CreationDate?.ToUniversalTime() ?? DateTime.UtcNow,
            Modified = DateTime.UtcNow
        };

        var document = SKDocument.CreatePdf(stream, docMetadata);
        return new PdfDocumentBuilder(document, stream, ownsStream);
    }

    /// <summary>
    /// Inizia una nuova pagina con le dimensioni specificate.
    /// Deve essere chiamato EndPage prima di iniziare una nuova pagina o chiudere il documento.
    /// </summary>
    /// <param name="paperSize">Dimensioni della pagina in pt</param>
    public PdfPage BeginPage(PaperSize paperSize)
    {
        if (_isClosed)
            throw new InvalidOperationException("Il documento è già stato chiuso");

        if (_currentPage != null)
            throw new InvalidOperationException("Chiamare EndPage prima di iniziare una nuova pagina");

        var canvas = _document.BeginPage(paperSize.WidthPt, paperSize.HeightPt);
        _currentPage = new PdfPage(canvas, paperSize);
        return _currentPage;
    }

    /// <summary>
    /// Termina la pagina corrente e la aggiunge al documento.
    /// </summary>
    public void EndPage()
    {
        if (_currentPage == null)
            throw new InvalidOperationException("Nessuna pagina attiva da terminare");

        _document.EndPage();
        _currentPage = null;
    }

    /// <summary>
    /// Chiude il documento e completa la scrittura del PDF.
    /// Dopo questa chiamata il builder non può più essere usato.
    /// </summary>
    public void Close()
    {
        if (_isClosed)
            return;

        if (_currentPage != null)
            throw new InvalidOperationException("Chiamare EndPage prima di chiudere il documento");

        _document.Close();
        _isClosed = true;
    }

    public void Dispose()
    {
        if (!_isClosed)
        {
            // Forza la chiusura se l'utente ha dimenticato di chiamare Close()
            if (_currentPage != null)
            {
                try { _document.EndPage(); }
                catch { /* Ignora errori durante cleanup */ }
            }

            try { _document.Close(); }
            catch { /* Ignora errori durante cleanup */ }

            _isClosed = true;
        }

        _document.Dispose();

        if (_ownsStream)
            _stream.Dispose();
    }
}

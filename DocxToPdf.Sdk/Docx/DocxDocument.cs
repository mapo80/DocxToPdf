using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx.Numbering;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocxToPdf.Sdk.Docx;

/// <summary>
/// Rappresenta un documento DOCX caricato, fornendo accesso a contenuti e metadati.
/// </summary>
public sealed class DocxDocument : IDisposable
{
    private readonly WordprocessingDocument _document;
    private readonly Styles.DocxStyleResolver _styleResolver;
    private readonly NumberingResolver _numberingResolver;
    private readonly float _defaultTabStopPt;
    private readonly DocxDocumentSettings _settings;
    private readonly IReadOnlyList<DocxBlock> _blocks;
    private bool _disposed;

    private DocxDocument(
        WordprocessingDocument document,
        Styles.DocxStyleResolver styleResolver,
        NumberingResolver numberingResolver,
        float defaultTabStopPt,
        DocxDocumentSettings settings)
    {
        _document = document;
        _styleResolver = styleResolver;
        _numberingResolver = numberingResolver;
        _defaultTabStopPt = defaultTabStopPt;
        _settings = settings;
        _blocks = BuildBlocks();
    }

    /// <summary>
    /// Apre un documento DOCX dal percorso specificato.
    /// </summary>
    public static DocxDocument Open(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"File DOCX non trovato: {path}", path);

        var document = WordprocessingDocument.Open(path, isEditable: false);
        var styleResolver = Styles.DocxStyleResolver.Load(document);
        var numberingResolver = new NumberingResolver(NumberingDefinitions.Load(document));
        var defaultTab = GetDefaultTabStopPt(document);
        var settings = DocxDocumentSettings.Load(document);
        return new DocxDocument(document, styleResolver, numberingResolver, defaultTab, settings);
    }

    /// <summary>
    /// Apre un documento DOCX da uno stream.
    /// </summary>
    public static DocxDocument Open(Stream stream)
    {
        var document = WordprocessingDocument.Open(stream, isEditable: false);
        var styleResolver = Styles.DocxStyleResolver.Load(document);
        var numberingResolver = new NumberingResolver(NumberingDefinitions.Load(document));
        var defaultTab = GetDefaultTabStopPt(document);
        var settings = DocxDocumentSettings.Load(document);
        return new DocxDocument(document, styleResolver, numberingResolver, defaultTab, settings);
    }

    /// <summary>
    /// Ottiene la sezione del documento (page size, margini, orientamento).
    /// </summary>
    public DocxSection GetSection()
    {
        var body = _document.MainDocumentPart?.Document?.Body;
        if (body == null)
            throw new InvalidOperationException("Documento DOCX non valido: Body mancante");

        // Cerca l'ultima SectionProperties nel documento
        var sectionProps = body.Descendants<SectionProperties>().LastOrDefault();

        return sectionProps != null
            ? DocxSection.FromSectionProperties(sectionProps)
            : DocxSection.Default;
    }

    /// <summary>
    /// Estrae tutti i paragrafi dal documento.
    /// </summary>
    public IEnumerable<DocxParagraph> GetParagraphs()
    {
        foreach (var block in _blocks)
        {
            if (block is DocxParagraphBlock paragraphBlock)
                yield return paragraphBlock.Paragraph;
        }
    }

    internal IEnumerable<DocxBlock> GetBlocks() => _blocks;

    private static float GetDefaultTabStopPt(WordprocessingDocument document)
    {
        var dxa = document.MainDocumentPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<DefaultTabStop>()?.Val?.Value ?? 720;
        return Units.UnitConverter.DxaToPoints((int)dxa);
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _document?.Dispose();
        _disposed = true;
    }

    private IReadOnlyList<DocxBlock> BuildBlocks()
    {
        var list = new List<DocxBlock>();
        var body = _document.MainDocumentPart?.Document?.Body;
        if (body == null)
            return list;

        foreach (var element in body.ChildElements)
        {
            switch (element)
            {
                case Paragraph paragraph:
                    list.Add(new DocxParagraphBlock(DocxParagraph.FromParagraph(
                        paragraph,
                        _styleResolver,
                        _numberingResolver,
                        _defaultTabStopPt,
                        _settings.EffectiveDecimalSymbol)));
                    break;
                case Table table:
                    var docxTable = DocxTable.FromOpenXmlTable(
                        table,
                        _styleResolver,
                        _numberingResolver,
                        _defaultTabStopPt,
                        _settings.EffectiveDecimalSymbol);
                    if (docxTable != null)
                        list.Add(new DocxTableBlock(docxTable));
                    break;
            }
        }

        return list;
    }
}

internal sealed record DocxDocumentSettings(char EffectiveDecimalSymbol, char? DocumentDecimalSymbol)
{
    public static DocxDocumentSettings Load(WordprocessingDocument document)
    {
        var settings = document.MainDocumentPart?.DocumentSettingsPart?.Settings;
        const char effective = '.';
        char? documentSymbol = null;
        var docValue = settings?.GetFirstChild<DecimalSymbol>()?.Val?.Value;
        if (!string.IsNullOrEmpty(docValue))
        {
            documentSymbol = docValue[0];
        }

        return new DocxDocumentSettings(effective, documentSymbol);
    }
}

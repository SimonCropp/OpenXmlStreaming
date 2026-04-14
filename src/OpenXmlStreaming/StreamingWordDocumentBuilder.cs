using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlStreaming;

/// <summary>
/// Higher-level builder on top of <see cref="OpenXmlPackageWriter"/> for producing
/// <c>.docx</c> documents. Attached sub-parts (styles, numbering, headers,
/// footers) are written up-front and return relationship ids the caller can
/// embed in the main document body (e.g. in <c>HeaderReference</c>/<c>FooterReference</c>
/// inside <c>SectionProperties</c>). The main <c>word/document.xml</c> is written
/// last via <see cref="WriteDocument"/>, with all accumulated relationships wired up.
/// </summary>
/// <remarks>
/// Unlike <see cref="StreamingWorkbookBuilder"/> which writes the main part
/// during <see cref="Dispose"/>, this builder requires an explicit
/// <see cref="WriteDocument"/> call — the document body is the only thing that
/// can actually reference the header/footer ids the builder hands out, so the
/// caller has to produce it.
/// </remarks>
public sealed class StreamingWordDocumentBuilder :
    IAsyncDisposable,
    IDisposable
{
    const string stylesRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    const string numberingRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    const string headerRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    const string footerRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    const string stylesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    const string numberingContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
    const string headerContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
    const string footerContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
    const string documentContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

    readonly OpenXmlPackageWriter writer;
    readonly List<PartRelationship> documentRelationships = [];
    int nextRelIndex;
    int nextHeaderIndex;
    int nextFooterIndex;
    bool stylesAdded;
    bool numberingAdded;
    bool documentWritten;

    /// <inheritdoc cref="OpenXmlPackageWriter(Stream, bool, int)"/>
    public StreamingWordDocumentBuilder(
        Stream stream,
        bool leaveOpen = false,
        int bufferSize = OpenXmlPackageWriter.DefaultBufferSize) =>
        writer = StreamingDocument.CreateWord(stream, leaveOpen, bufferSize);

    /// <summary>
    /// Writes <c>word/styles.xml</c> and records the relationship from the main
    /// document to the styles part. Can only be called once.
    /// </summary>
    /// <returns>The relationship id (mainly for diagnostic purposes — paragraphs reference styles by <c>StyleId</c>, not by relationship id).</returns>
    public string AddStyles(Styles styles)
    {
        ThrowIfDocumentWritten();

        if (styles is null)
        {
            throw new ArgumentNullException(nameof(styles));
        }

        if (stylesAdded)
        {
            throw new InvalidOperationException("Styles part has already been added.");
        }

        stylesAdded = true;
        return AddSubPart("/word/styles.xml", "styles.xml", stylesContentType, styles, stylesRelType);
    }

    /// <summary>
    /// Writes <c>word/numbering.xml</c> and records the relationship from the
    /// main document to the numbering part. Can only be called once.
    /// </summary>
    public string AddNumbering(Numbering numbering)
    {
        ThrowIfDocumentWritten();

        if (numbering is null)
        {
            throw new ArgumentNullException(nameof(numbering));
        }

        if (numberingAdded)
        {
            throw new InvalidOperationException("Numbering part has already been added.");
        }

        numberingAdded = true;
        return AddSubPart("/word/numbering.xml", "numbering.xml", numberingContentType, numbering, numberingRelType);
    }

    /// <summary>
    /// Writes a header part and returns the relationship id. The caller is
    /// responsible for referencing this id from a <c>HeaderReference</c> inside
    /// the document body's <c>SectionProperties</c>.
    /// </summary>
    public string AddHeader(Header header)
    {
        ThrowIfDocumentWritten();

        if (header is null)
        {
            throw new ArgumentNullException(nameof(header));
        }

        nextHeaderIndex++;
        var file = "header" + nextHeaderIndex.ToString(CultureInfo.InvariantCulture) + ".xml";
        return AddSubPart("/word/" + file, file, headerContentType, header, headerRelType);
    }

    /// <summary>
    /// Writes a footer part and returns the relationship id. The caller is
    /// responsible for referencing this id from a <c>FooterReference</c> inside
    /// the document body's <c>SectionProperties</c>.
    /// </summary>
    public string AddFooter(Footer footer)
    {
        ThrowIfDocumentWritten();

        if (footer is null)
        {
            throw new ArgumentNullException(nameof(footer));
        }

        nextFooterIndex++;
        var file = "footer" + nextFooterIndex.ToString(CultureInfo.InvariantCulture) + ".xml";
        return AddSubPart("/word/" + file, file, footerContentType, footer, footerRelType);
    }

    /// <summary>
    /// Writes the main <c>word/document.xml</c>, wiring up all the relationships
    /// for parts that have been added so far. Can only be called once.
    /// </summary>
    public void WriteDocument(Document document)
    {
        ThrowIfDocumentWritten();

        if (document is null)
        {
            throw new ArgumentNullException(nameof(document));
        }

        documentWritten = true;
        writer.WritePart(
            new("/word/document.xml", UriKind.Relative),
            documentContentType,
            document,
            documentRelationships);
    }

    public void Dispose() =>
        writer.Dispose();

    public ValueTask DisposeAsync() =>
        writer.DisposeAsync();

    string AddSubPart(string partUri, string relativeTarget, string contentType, OpenXmlElement content, string relationshipType)
    {
        var relId = "rId" + (++nextRelIndex).ToString(CultureInfo.InvariantCulture);

        writer.WritePart(
            new(partUri, UriKind.Relative),
            contentType,
            content);

        documentRelationships.Add(new PartRelationship(
            new(relativeTarget, UriKind.Relative),
            relationshipType,
            id: relId));

        return relId;
    }

    void ThrowIfDocumentWritten()
    {
        if (documentWritten)
        {
            throw new InvalidOperationException("Main document has already been written. No more sub-parts can be added.");
        }
    }
}

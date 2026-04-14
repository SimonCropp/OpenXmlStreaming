using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlStreaming;

/// <summary>
/// Higher-level builder on top of <see cref="OpenXmlPackageWriter"/> for producing
/// <c>.xlsx</c> workbooks without hand-managing worksheet URIs, relationship ids,
/// or the final <c>xl/workbook.xml</c> part. Add worksheets one at a time; the
/// builder writes <c>xl/workbook.xml</c> with all the relationships wired up
/// during <see cref="Finish"/>/<see cref="Dispose"/>/<see cref="DisposeAsync"/>.
/// </summary>
/// <remarks>
/// Worksheets are streamed directly to the target when you add them. The only
/// thing held in memory between <see cref="AddWorksheet"/> calls is a small list
/// of <c>(name, uri, relationship id)</c> tuples — one per worksheet — so the
/// builder inherits the streaming behaviour of <see cref="OpenXmlPackageWriter"/>.
/// </remarks>
/// <inheritdoc cref="OpenXmlPackageWriter(Stream, bool, int)"/>
public sealed class StreamingWorkbookBuilder(
    Stream stream,
    bool leaveOpen = false,
    int bufferSize = OpenXmlPackageWriter.DefaultBufferSize) :
    IAsyncDisposable,
    IDisposable
{
    const string worksheetRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    const string worksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    const string workbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

    OpenXmlPackageWriter writer = StreamingDocument.CreateSpreadsheet(stream, leaveOpen, bufferSize);
    List<(string Name, string Target, string RelId)> worksheets = [];
    bool finished;

    /// <summary>
    /// Writes a worksheet part to the package and records it for inclusion in
    /// the final <c>xl/workbook.xml</c>. Worksheets appear in the workbook's
    /// <c>Sheets</c> list in the order they were added.
    /// </summary>
    /// <param name="name">User-visible sheet name shown in the workbook tab.</param>
    /// <param name="worksheet">Worksheet content (the DOM element you would otherwise assign to <c>WorksheetPart.Worksheet</c>).</param>
    public void AddWorksheet(string name, Worksheet worksheet)
    {
        ThrowIfFinished();

        var index = worksheets.Count + 1;
        var fileName = "sheet" + index.ToString(CultureInfo.InvariantCulture) + ".xml";
        var partUri = "/xl/worksheets/" + fileName;
        var relId = "rId" + index.ToString(CultureInfo.InvariantCulture);

        writer.WritePart(
            new(partUri, UriKind.Relative),
            worksheetContentType,
            worksheet);

        worksheets.Add((name, "worksheets/" + fileName, relId));
    }

    /// <summary>
    /// Writes <c>xl/workbook.xml</c> referencing every worksheet that was added.
    /// Called automatically by <see cref="Dispose"/>/<see cref="DisposeAsync"/>;
    /// no further <see cref="AddWorksheet"/> calls are allowed after this.
    /// </summary>
    void Finish()
    {
        if (finished)
        {
            return;
        }

        finished = true;

        var sheetsElement = new Sheets();
        var relationships = new List<PartRelationship>(worksheets.Count);

        for (var i = 0; i < worksheets.Count; i++)
        {
            var (name, target, relId) = worksheets[i];
            sheetsElement.AppendChild(
                new Sheet
                {
                    Name = name,
                    SheetId = (uint) (i + 1),
                    Id = relId
                });
            relationships.Add(
                new(
                    new(target, UriKind.Relative),
                    worksheetRelType,
                    id: relId));
        }

        writer.WritePart(
            StreamingDocument.SpreadsheetWorkbookUri,
            workbookContentType,
            new Workbook(sheetsElement),
            relationships);
    }

    public void Dispose()
    {
        Finish();
        writer.Dispose();
    }

    public ValueTask DisposeAsync()
    {
        Finish();
        return writer.DisposeAsync();
    }

    void ThrowIfFinished()
    {
        if (finished)
        {
            throw new InvalidOperationException("Workbook has already been finalized. No more worksheets can be added.");
        }
    }
}

namespace OpenXmlStreaming;

/// <summary>
/// Factory methods for creating forward-only writers for Word, Spreadsheet, and Presentation documents.
/// </summary>
public static class StreamingDocument
{
    const string officeDocumentRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    internal static readonly Uri WordDocumentUri = new("/word/document.xml", UriKind.Relative);
    internal static readonly Uri SpreadsheetWorkbookUri = new("/xl/workbook.xml", UriKind.Relative);
    internal static readonly Uri PresentationUri = new("/ppt/presentation.xml", UriKind.Relative);

    /// <summary>
    /// Creates a forward-only writer for a Word document directly against the given stream.
    /// </summary>
    public static OpenXmlPackageWriter CreateWord(Stream stream, bool leaveOpen = false, int bufferSize = OpenXmlPackageWriter.DefaultBufferSize)
    {
        var writer = new OpenXmlPackageWriter(stream, leaveOpen, bufferSize);
        writer.AddRelationship(WordDocumentUri, officeDocumentRelationship, "rId1");
        return writer;
    }

    /// <summary>
    /// Creates a forward-only writer for a Spreadsheet document directly against the given stream.
    /// </summary>
    public static OpenXmlPackageWriter CreateSpreadsheet(Stream stream, bool leaveOpen = false, int bufferSize = OpenXmlPackageWriter.DefaultBufferSize)
    {
        var writer = new OpenXmlPackageWriter(stream, leaveOpen, bufferSize);
        writer.AddRelationship(SpreadsheetWorkbookUri, officeDocumentRelationship, "rId1");
        return writer;
    }

    /// <summary>
    /// Creates a forward-only writer for a Presentation document directly against the given stream.
    /// </summary>
    public static OpenXmlPackageWriter CreatePresentation(Stream stream, bool leaveOpen = false, int bufferSize = OpenXmlPackageWriter.DefaultBufferSize)
    {
        var writer = new OpenXmlPackageWriter(stream, leaveOpen, bufferSize);
        writer.AddRelationship(PresentationUri, officeDocumentRelationship, "rId1");
        return writer;
    }
}

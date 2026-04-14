namespace OpenXmlStreaming;

/// <summary>
/// Factory methods for creating forward-only writers for Word, Spreadsheet, and Presentation documents.
/// </summary>
public static class StreamingDocument
{
    const string officeDocumentRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    /// <summary>
    /// Creates a forward-only writer for a Word document directly against the given stream.
    /// </summary>
    public static OpenXmlPackageWriter CreateWord(Stream stream, WordprocessingDocumentType type, bool leaveOpen = false, int bufferSize = OpenXmlPackageWriter.DefaultBufferSize)
    {
        _ = type;
        var writer = new OpenXmlPackageWriter(stream, leaveOpen, bufferSize);
        writer.AddRelationship(new("/word/document.xml", UriKind.Relative), officeDocumentRelationship, "rId1");
        return writer;
    }

    /// <summary>
    /// Creates a forward-only writer for a Spreadsheet document directly against the given stream.
    /// </summary>
    public static OpenXmlPackageWriter CreateSpreadsheet(Stream stream, SpreadsheetDocumentType type, bool leaveOpen = false, int bufferSize = OpenXmlPackageWriter.DefaultBufferSize)
    {
        _ = type;
        var writer = new OpenXmlPackageWriter(stream, leaveOpen, bufferSize);
        writer.AddRelationship(new("/xl/workbook.xml", UriKind.Relative), officeDocumentRelationship, "rId1");
        return writer;
    }

    /// <summary>
    /// Creates a forward-only writer for a Presentation document directly against the given stream.
    /// </summary>
    public static OpenXmlPackageWriter CreatePresentation(Stream stream, PresentationDocumentType type, bool leaveOpen = false, int bufferSize = OpenXmlPackageWriter.DefaultBufferSize)
    {
        _ = type;
        var writer = new OpenXmlPackageWriter(stream, leaveOpen, bufferSize);
        writer.AddRelationship(new("/ppt/presentation.xml", UriKind.Relative), officeDocumentRelationship, "rId1");
        return writer;
    }
}

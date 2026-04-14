using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;

[TestFixture]
public class Samples
{
    [Test]
    public void MinimalWord()
    {
        using var stream = new MemoryStream();

        // begin-snippet: minimal-word
        using var writer = new OpenXmlPackageWriter(stream, leaveOpen: true);

        writer.AddRelationship(
            new("/word/document.xml", UriKind.Relative),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "rId1");

        writer.WritePart(
            new("/word/document.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            new Document(new Body(new Paragraph(new Run(new Text("Hello!"))))));
        // end-snippet
    }

    [Test]
    public void StreamingDocumentFactory()
    {
        using var stream = new MemoryStream();

        // begin-snippet: streaming-document-factory
        using var writer = StreamingDocument.CreateWord(
            stream,
            WordprocessingDocumentType.Document,
            leaveOpen: true);

        writer.WritePart(
            new("/word/document.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            new Document(new Body(new Paragraph(new Run(new Text("Forward-only!"))))));
        // end-snippet
    }

    [Test]
    public void StreamingPartContent()
    {
        using var stream = new MemoryStream();
        using var writer = new OpenXmlPackageWriter(stream, leaveOpen: true);

        writer.AddRelationship(
            new("/word/document.xml", UriKind.Relative),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "rId1");

        // begin-snippet: streaming-part-content
        using var entry = writer.CreatePart(
            new("/word/document.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");

        using var xmlWriter = OpenXmlWriter.Create(entry.Stream);
        xmlWriter.WriteStartDocument();
        xmlWriter.WriteStartElement(new Document());
        xmlWriter.WriteStartElement(new Body());
        xmlWriter.WriteElement(new Paragraph(new Run(new Text("Streamed!"))));
        xmlWriter.WriteEndElement();
        xmlWriter.WriteEndElement();
        // end-snippet
    }

    [Test]
    public void PartRelationships()
    {
        using var stream = new MemoryStream();
        using var writer = StreamingDocument.CreateSpreadsheet(
            stream,
            SpreadsheetDocumentType.Workbook,
            leaveOpen: true);

        // begin-snippet: part-relationships
        writer.WritePart(
            new("/xl/workbook.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            new S.Workbook(
                new S.Sheets(new S.Sheet
                {
                    Name = "Sheet1",
                    SheetId = 1,
                    Id = "rId1"
                })),
            [
                new(
                    new("worksheets/sheet1.xml", UriKind.Relative),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                    id: "rId1")
            ]);
        // end-snippet

        writer.WritePart(
            new("/xl/worksheets/sheet1.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            new S.Worksheet(new S.SheetData()));
    }

    [Test]
    public void ExternalRelationship()
    {
        using var stream = new MemoryStream();
        using var writer = new OpenXmlPackageWriter(stream, leaveOpen: true);

        writer.AddRelationship(
            new("/word/document.xml", UriKind.Relative),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "rId1");

        using var entry = writer.CreatePart(
            new("/word/document.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");

        // begin-snippet: external-relationship
        entry.AddRelationship(
            new("https://example.com", UriKind.Absolute),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            TargetMode.External,
            "rId1");
        // end-snippet
    }

    [Test]
    public void Presentation()
    {
        using var stream = new MemoryStream();

        // begin-snippet: create-presentation
        using var writer = StreamingDocument.CreatePresentation(
            stream,
            PresentationDocumentType.Presentation,
            leaveOpen: true);

        writer.WritePart(
            new("/ppt/presentation.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            new P.Presentation(new P.SlideIdList()));
        // end-snippet
    }
}

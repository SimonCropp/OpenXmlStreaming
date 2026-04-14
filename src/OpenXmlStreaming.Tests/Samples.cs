using DocumentFormat.OpenXml.Wordprocessing;
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
        using var writer = StreamingDocument.CreateWord(stream, leaveOpen: true);

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
        using var writer = StreamingDocument.CreateSpreadsheet(stream, leaveOpen: true);

        // begin-snippet: part-relationships
        writer.WritePart(
            new("/xl/workbook.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            new S.Workbook(
                new S.Sheets(
                    new S.Sheet
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
        using var writer = StreamingDocument.CreatePresentation(stream, leaveOpen: true);

        writer.WritePart(
            new("/ppt/presentation.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            new P.Presentation(new P.SlideIdList()));
        // end-snippet
    }

    [Test]
    public void ConstructionVariants()
    {
        using var stream = new MemoryStream();

        // begin-snippet: construction-variants
        // Direct construction
        using var direct = new OpenXmlPackageWriter(stream, leaveOpen: true);

        // Typed factories (pre-register the officeDocument relationship)
        using var word = StreamingDocument.CreateWord(stream, leaveOpen: true);
        using var spreadsheet = StreamingDocument.CreateSpreadsheet(stream, leaveOpen: true);
        using var presentation = StreamingDocument.CreatePresentation(stream, leaveOpen: true);
        // end-snippet
    }

    [Test]
    public void PartRelationshipStruct()
    {
        // begin-snippet: part-relationship-struct
        var relationship = new PartRelationship(
            targetUri: new("styles.xml", UriKind.Relative),
            relationshipType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            // default
            targetMode: TargetMode.Internal,
            // optional, auto-generated if null
            id: "rId1");
        // end-snippet

        _ = relationship;
    }

    [Test]
    public async Task AsyncUsage()
    {
        using var stream = new MemoryStream();

        // begin-snippet: async-usage
        await using var writer = StreamingDocument.CreateWord(stream, leaveOpen: true);

        writer.WritePart(
            new("/word/document.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            new Document(new Body(new Paragraph(new Run(new Text("Streamed async!"))))));

        // DisposeAsync (triggered by `await using`) asynchronously flushes
        // the final buffer — including the ZIP central directory — so remote
        // sinks like SQL BLOB streams don't block the thread on network I/O.
        // end-snippet
    }

    [Test]
    public void CustomBufferSize()
    {
        using var stream = new MemoryStream();

        // begin-snippet: custom-buffer-size
        // Bigger buffer = fewer, larger writes hit the sink — good for
        // remote streams where per-write overhead is high. Pass 0 to
        // disable buffering entirely and write straight to the sink.
        using var writer = new OpenXmlPackageWriter(
            stream,
            leaveOpen: true,
            // 1 MB
            bufferSize: 1024 * 1024);
        // end-snippet
    }

    [Test]
    public async Task FlushBetweenParts()
    {
        using var stream = new MemoryStream();

        await using var writer = StreamingDocument.CreateSpreadsheet(stream, leaveOpen: true);

        // begin-snippet: flush-async
        // Write the worksheet, then push its bytes to the target stream
        // asynchronously before moving on to the next part. Useful at part
        // boundaries against remote sinks — the thread isn't blocked on
        // network I/O while the next part is being serialized.
        writer.WritePart(
            new("/xl/worksheets/sheet1.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            new S.Worksheet(new S.SheetData()));

        await writer.FlushAsync();

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
                    id: "rId1"),
            ]);
        // end-snippet
    }

    [Test]
    public async Task WorkbookBuilderSample()
    {
        using var stream = new MemoryStream();

        // begin-snippet: workbook-builder
        await using var workbook = new StreamingWorkbookBuilder(stream, leaveOpen: true);

        workbook.AddWorksheet(
            "Revenue",
            new(
                new S.SheetData(
                    new S.Row(
                        new S.Cell
                        {
                            CellValue = new("Q1"),
                            DataType = S.CellValues.InlineString
                        },
                        new S.Cell
                        {
                            CellValue = new("1000"),
                            DataType = S.CellValues.Number
                        }),
                    new S.Row(
                        new S.Cell
                        {
                            CellValue = new("Q2"),
                            DataType = S.CellValues.InlineString
                        },
                        new S.Cell
                        {
                            CellValue = new("1200"),
                            DataType = S.CellValues.Number
                        }))));

        workbook.AddWorksheet(
            "Expenses",
            new(
                new S.SheetData(
                    new S.Row(
                        new S.Cell
                        {
                            CellValue = new("Rent"),
                            DataType = S.CellValues.InlineString
                        },
                        new S.Cell
                        {
                            CellValue = new("500"),
                            DataType = S.CellValues.Number
                        }))));

        // DisposeAsync (triggered by `await using`) writes xl/workbook.xml
        // referencing both sheets — no manual rId wiring.
        // end-snippet
    }

    [Test]
    public async Task WordDocumentBuilderSample()
    {
        using var stream = new MemoryStream();

        // begin-snippet: word-document-builder
        await using var word = new StreamingWordDocumentBuilder(stream, leaveOpen: true);

        // Add an optional styles part — referenced by paragraphs via StyleId.
        word.AddStyles(
            new(
            new Style(
                new StyleName
                {
                    Val = "Heading 1"
                },
                new BasedOn
                {
                    Val = "Normal"
                },
                new StyleRunProperties(
                    new Bold(),
                    new FontSize
                    {
                        Val = "32"
                    }))
            {
                Type = StyleValues.Paragraph,
                StyleId = "Heading1"
            }));

        // AddFooter returns the relationship id — plug it into SectionProperties
        // in the document body below so the body-level FooterReference resolves.
        var footerId = word.AddFooter(
            new(
                new Paragraph(new Run(new Text("— Confidential —")))));

        // Last step: write the main document body. The builder wires up all
        // the accumulated sub-part relationships (styles + footer) for you.
        word.WriteDocument(
            new(
                new Body(
                    new Paragraph(
                        new ParagraphProperties(
                            new ParagraphStyleId
                            {
                                Val = "Heading1"
                            }),
                        new Run(
                            new Text("Quarterly Report"))),
                    new Paragraph(
                        new Run(
                            new Text(
                                "Revenue grew 15% year-over-year."))),
                    new SectionProperties(
                        new FooterReference
                        {
                            Type = HeaderFooterValues.Default,
                            Id = footerId
                        }))));
        // end-snippet
    }

    [Test]
    public async Task PresentationBuilderSample()
    {
        using var stream = new MemoryStream();

        // begin-snippet: presentation-builder
        await using var presentation = new StreamingPresentationBuilder(stream, leaveOpen: true);

        // No theme/master/layout boilerplate — the builder writes a default
        // scaffolding on the first AddSlide call.
        presentation.AddSlide(
            new(
                new P.CommonSlideData(
                    new P.ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties
                            {
                                Id = 1,
                                Name = ""
                            },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties()),
                        new P.GroupShapeProperties(
                            new DocumentFormat.OpenXml.Drawing.TransformGroup())))));

        presentation.AddSlide(
            new(
                new P.CommonSlideData(
                    new P.ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties
                            {
                                Id = 1,
                                Name = ""
                            },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties()),
                        new P.GroupShapeProperties(
                            new DocumentFormat.OpenXml.Drawing.TransformGroup())))));

        // DisposeAsync writes ppt/presentation.xml referencing the slide
        // master and every slide that was added.
        // end-snippet

        await Task.CompletedTask;
    }
}

using DocumentFormat.OpenXml.Wordprocessing;

public partial class MigrationGuide
{
    [Test]
    public async Task WordStandard()
    {
        using var stream = new MemoryStream();

        // begin-snippet: migration-word-standard
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();

            // Add a styles part through the DOM. The SDK wires the relationship
            // from the main part to the styles part automatically.
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new(
                new Style(
                    new StyleName
                    {
                        Val = "Heading 1"
                    },
                    new BasedOn
                    {
                        Val = "Normal"
                    },
                    new NextParagraphStyle
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
                });

            // Assign the main document body. Paragraphs reference the style by id.
            mainPart.Document =
                new(
                    new Body(
                        new Paragraph(
                            new ParagraphProperties(
                                new ParagraphStyleId
                                {
                                    Val = "Heading1"
                                }),
                            new Run(new Text("Quarterly Report"))),
                        new Paragraph(
                            new Run(new Text("Revenue grew 15% year-over-year."))),
                        new Paragraph(
                            new Run(new Text("Operating costs held flat.")))));
        }
        // end-snippet

        stream.Position = 0;
        await Verify(stream, extension: "docx");
    }

    [Test]
    public async Task WordStreaming()
    {
        using var stream = new MemoryStream();

        // begin-snippet: migration-word-streaming
        await using (var writer = StreamingDocument.CreateWord(stream, leaveOpen: true))
        {
            // Write the styles part first. Every part is written as a
            // complete element tree — there's no DOM to mutate later.
            writer.WritePart(
                new("/word/styles.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
                new Styles(
                    new Style(
                        new StyleName
                        {
                            Val = "Heading 1"
                        },
                        new BasedOn
                        {
                            Val = "Normal"
                        },
                        new NextParagraphStyle
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

            // Write the main document, declaring its relationship to the
            // styles part inline via the relationships parameter.
            writer.WritePart(
                new("/word/document.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                new Document(
                    new Body(
                        new Paragraph(
                            new ParagraphProperties(
                                new ParagraphStyleId
                                {
                                    Val = "Heading1"
                                }),
                            new Run(new Text("Quarterly Report"))),
                        new Paragraph(
                            new Run(new Text("Revenue grew 15% year-over-year."))),
                        new Paragraph(
                            new Run(new Text("Operating costs held flat."))))),
                [
                    new(
                        new("styles.xml", UriKind.Relative),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                        id: "rId1"),
                ]);
        }
        // end-snippet

        stream.Position = 0;
        await Verify(stream, extension: "docx");
    }

    [Test]
    public async Task WordBuilder()
    {
        using var stream = new MemoryStream();

        // begin-snippet: migration-word-builder
        await using (var word = new StreamingWordDocumentBuilder(stream, leaveOpen: true))
        {
            // Add the styles part. The builder writes it immediately and
            // tracks the relationship for the main document below.
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
                        new NextParagraphStyle
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

            // Write the main document. The builder wires up the styles
            // relationship for you — no PartRelationship plumbing.
            word.WriteDocument(
                new(
                    new Body(
                        new Paragraph(
                            new ParagraphProperties(
                                new ParagraphStyleId
                                {
                                    Val = "Heading1"
                                }),
                            new Run(new Text("Quarterly Report"))),
                        new Paragraph(
                            new Run(new Text("Revenue grew 15% year-over-year."))),
                        new Paragraph(
                            new Run(new Text("Operating costs held flat."))))));
        }
        // end-snippet

        stream.Position = 0;
        await Verify(stream, extension: "docx");
    }
}
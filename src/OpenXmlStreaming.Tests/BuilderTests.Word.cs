using DocumentFormat.OpenXml.Wordprocessing;

public partial class BuilderTests
{
    [Test]
    public async Task StreamingWordDocumentBuilder_RoundTrips()
    {
        using var stream = new MemoryStream();

        await using (var word = new StreamingWordDocumentBuilder(stream, leaveOpen: true))
        {
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

            var footerId = word.AddFooter(
                new(
                    new Paragraph(
                        new Run(new Text("— Confidential —")))));

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
                        new SectionProperties(
                            new FooterReference
                            {
                                Type = HeaderFooterValues.Default,
                                Id = footerId
                            }))));
        }

        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.That(doc.MainDocumentPart!.Document!.Body!.InnerText, Does.Contain("Quarterly Report"));
        Assert.That(doc.MainDocumentPart.FooterParts.Count(), Is.EqualTo(1));
        Assert.That(doc.MainDocumentPart.StyleDefinitionsPart, Is.Not.Null);

        stream.Position = 0;
        await Verify(stream, extension: "docx");
    }

    [Test]
    public void StreamingWordDocumentBuilder_AddStylesAfterDocument_Throws()
    {
        using var stream = new MemoryStream();
        using var word = new StreamingWordDocumentBuilder(stream, leaveOpen: true);
        word.WriteDocument(new(new Body()));

        Assert.Throws<InvalidOperationException>(() =>
            word.AddStyles(new()));
    }

    [Test]
    public void StreamingWordDocumentBuilder_DoubleStyles_Throws()
    {
        using var stream = new MemoryStream();
        using var word = new StreamingWordDocumentBuilder(stream, leaveOpen: true);
        word.AddStyles(new());

        Assert.Throws<InvalidOperationException>(() =>
            word.AddStyles(new()));
    }
}

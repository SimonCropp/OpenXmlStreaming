using DocumentFormat.OpenXml.Wordprocessing;

public partial class BuilderTests
{
    [Test]
    public async Task StreamingWordDocumentBuilder_RoundTrips()
    {
        using var ms = new MemoryStream();

        await using (var word = new StreamingWordDocumentBuilder(ms, leaveOpen: true))
        {
            word.AddStyles(new Styles(
                new Style(
                    new StyleName { Val = "Heading 1" },
                    new BasedOn { Val = "Normal" },
                    new NextParagraphStyle { Val = "Normal" },
                    new StyleRunProperties(
                        new Bold(),
                        new FontSize { Val = "32" }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Heading1"
                }));

            var footerId = word.AddFooter(new Footer(
                new Paragraph(
                    new Run(new Text("— Confidential —")))));

            word.WriteDocument(new Document(
                new Body(
                    new Paragraph(
                        new ParagraphProperties(
                            new ParagraphStyleId { Val = "Heading1" }),
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

        ms.Position = 0;
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.That(doc.MainDocumentPart!.Document!.Body!.InnerText, Does.Contain("Quarterly Report"));
        Assert.That(doc.MainDocumentPart.FooterParts.Count(), Is.EqualTo(1));
        Assert.That(doc.MainDocumentPart.StyleDefinitionsPart, Is.Not.Null);

        ms.Position = 0;
        await Verify(ms, extension: "docx");
    }

    [Test]
    public void StreamingWordDocumentBuilder_AddStylesAfterDocument_Throws()
    {
        using var ms = new MemoryStream();
        using var word = new StreamingWordDocumentBuilder(ms, leaveOpen: true);
        word.WriteDocument(new Document(new Body()));

        Assert.Throws<InvalidOperationException>(() =>
            word.AddStyles(new Styles()));
    }

    [Test]
    public void StreamingWordDocumentBuilder_DoubleStyles_Throws()
    {
        using var ms = new MemoryStream();
        using var word = new StreamingWordDocumentBuilder(ms, leaveOpen: true);
        word.AddStyles(new Styles());

        Assert.Throws<InvalidOperationException>(() =>
            word.AddStyles(new Styles()));
    }
}

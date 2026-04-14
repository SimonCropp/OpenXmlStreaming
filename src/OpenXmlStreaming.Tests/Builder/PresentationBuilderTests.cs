using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

public class PresentationBuilderTests
{
    [Test]
    public async Task RoundTrips()
    {
        using var stream = new MemoryStream();

        await using (var presentation = new StreamingPresentationBuilder(stream, leaveOpen: true))
        {
            presentation.AddSlide(TitleSlide("Kickoff"));
            presentation.AddSlide(TitleSlide("Agenda"));
        }

        stream.Position = 0;
        using var doc = PresentationDocument.Open(stream, false);
        var slideIds = doc.PresentationPart!.Presentation!.SlideIdList!.Elements<SlideId>().ToList();

        Assert.Multiple(() =>
        {
            Assert.That(slideIds, Has.Count.EqualTo(2));
            Assert.That(doc.PresentationPart.SlideParts.Count(), Is.EqualTo(2));
            Assert.That(doc.PresentationPart.SlideMasterParts.Count(), Is.EqualTo(1));
        });

        stream.Position = 0;
        await Verify(stream, extension: "pptx");
    }

    [Test]
    public async Task NoSlides_StillWritesScaffolding()
    {
        using var stream = new MemoryStream();

        await using (var _ = new StreamingPresentationBuilder(stream, leaveOpen: true))
        {
            // Intentionally no slides added.
        }

        stream.Position = 0;
        using var doc = PresentationDocument.Open(stream, false);
        Assert.That(doc.PresentationPart!.Presentation, Is.Not.Null);
    }

    [Test]
    public void AddAfterDispose_Throws()
    {
        using var stream = new MemoryStream();
        var presentation = new StreamingPresentationBuilder(stream, leaveOpen: true);
        presentation.Dispose();

        Assert.Throws<InvalidOperationException>(() =>
            presentation.AddSlide(TitleSlide("Late")));
    }

    static Slide TitleSlide(string title) =>
        new(
            new CommonSlideData(
                new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties
                        {
                            Id = 1,
                            Name = ""
                        },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new Drawing.TransformGroup()),
                    new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties
                            {
                                Id = 2,
                                Name = "Title"
                            },
                            new NonVisualShapeDrawingProperties(
                                new Drawing.ShapeLocks
                                {
                                    NoGrouping = true
                                }),
                            new ApplicationNonVisualDrawingProperties(
                                new PlaceholderShape
                                {
                                    Type = PlaceholderValues.CenteredTitle
                                })),
                        new ShapeProperties(),
                        new TextBody(
                            new Drawing.BodyProperties(),
                            new Drawing.ListStyle(),
                            new Drawing.Paragraph(
                                new Drawing.Run(
                                    new Drawing.RunProperties
                                    {
                                        Language = "en-US"
                                    },
                                    new Drawing.Text(title))))))));
}

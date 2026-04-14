using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

public partial class Samples
{
    [Test]
    public void Presentation()
    {
        using var stream = new MemoryStream();

        #region create-presentation
        using var writer = StreamingDocument.CreatePresentation(stream, leaveOpen: true);

        writer.WritePart(
            new("/ppt/presentation.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            new Presentation(new SlideIdList()));
        #endregion
    }

    [Test]
    public async Task PresentationBuilderSample()
    {
        using var stream = new MemoryStream();

        #region presentation-builder
        await using var presentation = new StreamingPresentationBuilder(stream, leaveOpen: true);

        // No theme/master/layout boilerplate — the builder writes a default
        // scaffolding on the first AddSlide call.
        presentation.AddSlide(
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
                        new GroupShapeProperties(
                            new Drawing.TransformGroup())))));

        presentation.AddSlide(
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
                        new GroupShapeProperties(
                            new Drawing.TransformGroup())))));

        // DisposeAsync writes ppt/presentation.xml referencing the slide
        // master and every slide that was added.
        #endregion

        await Task.CompletedTask;
    }
}

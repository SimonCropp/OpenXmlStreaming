using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

[TestFixture]
public class MigrationGuide
{
    // ───────────────────────── Word ─────────────────────────

    [Test]
    public async Task WordStandard()
    {
        using var ms = new MemoryStream();

        // begin-snippet: migration-word-standard
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();

            // Add a styles part through the DOM. The SDK wires the relationship
            // from the main part to the styles part automatically.
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new Style(
                    new StyleName { Val = "Heading 1" },
                    new BasedOn { Val = "Normal" },
                    new NextParagraphStyle { Val = "Normal" },
                    new W.StyleRunProperties(
                        new W.Bold(),
                        new W.FontSize { Val = "32" }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Heading1"
                });

            // Assign the main document body. Paragraphs reference the style by id.
            mainPart.Document = new Document(
                new Body(
                    new Paragraph(
                        new ParagraphProperties(
                            new ParagraphStyleId { Val = "Heading1" }),
                        new W.Run(new W.Text("Quarterly Report"))),
                    new Paragraph(
                        new W.Run(new W.Text("Revenue grew 15% year-over-year."))),
                    new Paragraph(
                        new W.Run(new W.Text("Operating costs held flat.")))));
        }
        // end-snippet

        ms.Position = 0;
        await Verify(ms, extension: "docx");
    }

    [Test]
    public async Task WordStreaming()
    {
        using var ms = new MemoryStream();

        // begin-snippet: migration-word-streaming
        await using (var writer = StreamingDocument.CreateWord(ms, leaveOpen: true))
        {
            // Write the styles part first. Every part is written as a
            // complete element tree — there's no DOM to mutate later.
            writer.WritePart(
                new("/word/styles.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
                new Styles(
                    new Style(
                        new StyleName { Val = "Heading 1" },
                        new BasedOn { Val = "Normal" },
                        new NextParagraphStyle { Val = "Normal" },
                        new W.StyleRunProperties(
                            new W.Bold(),
                            new W.FontSize { Val = "32" }))
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
                                new ParagraphStyleId { Val = "Heading1" }),
                            new W.Run(new W.Text("Quarterly Report"))),
                        new Paragraph(
                            new W.Run(new W.Text("Revenue grew 15% year-over-year."))),
                        new Paragraph(
                            new W.Run(new W.Text("Operating costs held flat."))))),
                [
                    new(
                        new("styles.xml", UriKind.Relative),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                        id: "rId1"),
                ]);
        }
        // end-snippet

        ms.Position = 0;
        await Verify(ms, extension: "docx");
    }

    // ───────────────────────── Spreadsheet ─────────────────────────

    [Test]
    public async Task SpreadsheetStandard()
    {
        using var ms = new MemoryStream();

        // begin-snippet: migration-spreadsheet-standard
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            var sheets = new S.Sheets();

            // Revenue sheet
            var revenuePart = workbookPart.AddNewPart<WorksheetPart>();
            revenuePart.Worksheet = new S.Worksheet(
                new S.SheetData(
                    new S.Row(
                        InlineString("A1", "Quarter"),
                        InlineString("B1", "Revenue"))
                    { RowIndex = 1 },
                    new S.Row(
                        InlineString("A2", "Q1"),
                        Number("B2", "1000"))
                    { RowIndex = 2 },
                    new S.Row(
                        InlineString("A3", "Q2"),
                        Number("B3", "1200"))
                    { RowIndex = 3 }));
            sheets.AppendChild(new S.Sheet
            {
                Name = "Revenue",
                SheetId = 1,
                Id = workbookPart.GetIdOfPart(revenuePart)
            });

            // Expenses sheet
            var expensesPart = workbookPart.AddNewPart<WorksheetPart>();
            expensesPart.Worksheet = new S.Worksheet(
                new S.SheetData(
                    new S.Row(
                        InlineString("A1", "Category"),
                        InlineString("B1", "Amount"))
                    { RowIndex = 1 },
                    new S.Row(
                        InlineString("A2", "Rent"),
                        Number("B2", "500"))
                    { RowIndex = 2 }));
            sheets.AppendChild(new S.Sheet
            {
                Name = "Expenses",
                SheetId = 2,
                Id = workbookPart.GetIdOfPart(expensesPart)
            });

            workbookPart.Workbook = new S.Workbook(sheets);
        }
        // end-snippet

        ms.Position = 0;
        await Verify(ms, extension: "xlsx");
    }

    [Test]
    public async Task SpreadsheetStreaming()
    {
        using var ms = new MemoryStream();

        // begin-snippet: migration-spreadsheet-streaming
        await using (var writer = StreamingDocument.CreateSpreadsheet(ms, leaveOpen: true))
        {
            // Worksheets are written first — the workbook references them by id.
            writer.WritePart(
                new("/xl/worksheets/sheet1.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                new S.Worksheet(
                    new S.SheetData(
                        new S.Row(
                            InlineString("A1", "Quarter"),
                            InlineString("B1", "Revenue"))
                        { RowIndex = 1 },
                        new S.Row(
                            InlineString("A2", "Q1"),
                            Number("B2", "1000"))
                        { RowIndex = 2 },
                        new S.Row(
                            InlineString("A3", "Q2"),
                            Number("B3", "1200"))
                        { RowIndex = 3 })));

            writer.WritePart(
                new("/xl/worksheets/sheet2.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                new S.Worksheet(
                    new S.SheetData(
                        new S.Row(
                            InlineString("A1", "Category"),
                            InlineString("B1", "Amount"))
                        { RowIndex = 1 },
                        new S.Row(
                            InlineString("A2", "Rent"),
                            Number("B2", "500"))
                        { RowIndex = 2 })));

            // Then the workbook, with a relationship per worksheet.
            writer.WritePart(
                new("/xl/workbook.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                new S.Workbook(
                    new S.Sheets(
                        new S.Sheet { Name = "Revenue", SheetId = 1, Id = "rId1" },
                        new S.Sheet { Name = "Expenses", SheetId = 2, Id = "rId2" })),
                [
                    new(
                        new("worksheets/sheet1.xml", UriKind.Relative),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                        id: "rId1"),
                    new(
                        new("worksheets/sheet2.xml", UriKind.Relative),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                        id: "rId2"),
                ]);
        }
        // end-snippet

        ms.Position = 0;
        await Verify(ms, extension: "xlsx");
    }

    static S.Cell InlineString(string reference, string value) =>
        new()
        {
            CellReference = reference,
            DataType = S.CellValues.InlineString,
            InlineString = new(new S.Text(value))
        };

    static S.Cell Number(string reference, string value) =>
        new()
        {
            CellReference = reference,
            DataType = S.CellValues.Number,
            CellValue = new(value)
        };

    // ───────────────────────── Presentation ─────────────────────────

    [Test]
    public async Task PresentationStandard()
    {
        using var ms = new MemoryStream();

        // begin-snippet: migration-presentation-standard
        using (var doc = PresentationDocument.Create(ms, PresentationDocumentType.Presentation))
        {
            var presentationPart = doc.AddPresentationPart();
            presentationPart.Presentation = new P.Presentation();

            // Slide master + theme + layout are required scaffolding for
            // any presentation with real slides. AddNewPart wires the
            // relationships between them automatically.
            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = BuildSlideMaster();

            var themePart = slideMasterPart.AddNewPart<ThemePart>();
            themePart.Theme = BuildTheme();

            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = BuildSlideLayout();

            // Slide 1 — title slide referencing the layout.
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(slideLayoutPart);
            slidePart.Slide = BuildTitleSlide("Kickoff");

            // Stitch the presentation's lists together using relationship ids
            // the SDK generated when AddNewPart was called.
            presentationPart.Presentation = new P.Presentation(
                new P.SlideMasterIdList(new P.SlideMasterId
                {
                    Id = 2147483648U,
                    RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
                }),
                new P.SlideIdList(new P.SlideId
                {
                    Id = 256U,
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                }),
                new P.SlideSize { Cx = 9144000, Cy = 6858000 },
                new P.NotesSize { Cx = 6858000, Cy = 9144000 });
        }
        // end-snippet

        ms.Position = 0;
        await Verify(ms, extension: "pptx");
    }

    [Test]
    public async Task PresentationStreaming()
    {
        using var ms = new MemoryStream();

        // begin-snippet: migration-presentation-streaming
        await using (var writer = StreamingDocument.CreatePresentation(ms, leaveOpen: true))
        {
            // Theme — referenced from the slide master.
            writer.WritePart(
                new("/ppt/theme/theme1.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.theme+xml",
                BuildTheme());

            // Slide layout — referenced from the slide master AND each slide.
            writer.WritePart(
                new("/ppt/slideLayouts/slideLayout1.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
                BuildSlideLayout());

            // Slide master — references the theme and the layout.
            writer.WritePart(
                new("/ppt/slideMasters/slideMaster1.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
                BuildSlideMaster(),
                [
                    new(
                        new("../theme/theme1.xml", UriKind.Relative),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                        id: "rId1"),
                    new(
                        new("../slideLayouts/slideLayout1.xml", UriKind.Relative),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                        id: "rId2"),
                ]);

            // Slide — references its layout.
            writer.WritePart(
                new("/ppt/slides/slide1.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
                BuildTitleSlide("Kickoff"),
                [
                    new(
                        new("../slideLayouts/slideLayout1.xml", UriKind.Relative),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                        id: "rId1"),
                ]);

            // presentation.xml — references the slide master and the slide.
            writer.WritePart(
                new("/ppt/presentation.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                new P.Presentation(
                    new P.SlideMasterIdList(new P.SlideMasterId
                    {
                        Id = 2147483648U,
                        RelationshipId = "rId1"
                    }),
                    new P.SlideIdList(new P.SlideId
                    {
                        Id = 256U,
                        RelationshipId = "rId2"
                    }),
                    new P.SlideSize { Cx = 9144000, Cy = 6858000 },
                    new P.NotesSize { Cx = 6858000, Cy = 9144000 }),
                [
                    new(
                        new("slideMasters/slideMaster1.xml", UriKind.Relative),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
                        id: "rId1"),
                    new(
                        new("slides/slide1.xml", UriKind.Relative),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                        id: "rId2"),
                ]);
        }
        // end-snippet

        ms.Position = 0;
        await Verify(ms, extension: "pptx");
    }

    static P.SlideMaster BuildSlideMaster() =>
        new(
            new P.CommonSlideData(
                new P.Background(
                    new P.BackgroundStyleReference(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })
                    { Index = 1001 }),
                new P.ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.GroupShapeProperties(new A.TransformGroup()))),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            },
            new P.SlideLayoutIdList(new P.SlideLayoutId { Id = 2147483649U, RelationshipId = "rId2" }));

    static P.SlideLayout BuildSlideLayout() =>
        new(
            new P.CommonSlideData(
                new P.ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.GroupShapeProperties(new A.TransformGroup()))),
            new P.ColorMapOverride(new A.MasterColorMapping()))
        { Type = P.SlideLayoutValues.Title };

    static P.Slide BuildTitleSlide(string title) =>
        new(
            new P.CommonSlideData(
                new P.ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.GroupShapeProperties(new A.TransformGroup()),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 2, Name = "Title" },
                            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                            new P.ApplicationNonVisualDrawingProperties(
                                new P.PlaceholderShape { Type = P.PlaceholderValues.CenteredTitle })),
                        new P.ShapeProperties(),
                        new P.TextBody(
                            new A.BodyProperties(),
                            new A.ListStyle(),
                            new A.Paragraph(
                                new A.Run(
                                    new A.RunProperties { Language = "en-US" },
                                    new A.Text(title))))))));

    static A.Theme BuildTheme() =>
        new(
            new A.ThemeElements(
                new A.ColorScheme(
                    new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
                    new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
                    new A.Dark2Color(new A.RgbColorModelHex { Val = "1F497D" }),
                    new A.Light2Color(new A.RgbColorModelHex { Val = "EEECE1" }),
                    new A.Accent1Color(new A.RgbColorModelHex { Val = "4F81BD" }),
                    new A.Accent2Color(new A.RgbColorModelHex { Val = "C0504D" }),
                    new A.Accent3Color(new A.RgbColorModelHex { Val = "9BBB59" }),
                    new A.Accent4Color(new A.RgbColorModelHex { Val = "8064A2" }),
                    new A.Accent5Color(new A.RgbColorModelHex { Val = "4BACC6" }),
                    new A.Accent6Color(new A.RgbColorModelHex { Val = "F79646" }),
                    new A.Hyperlink(new A.RgbColorModelHex { Val = "0000FF" }),
                    new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "800080" }))
                { Name = "Office" },
                new A.FontScheme(
                    new A.MajorFont(new A.LatinFont { Typeface = "Calibri" }, new A.EastAsianFont { Typeface = "" }, new A.ComplexScriptFont { Typeface = "" }),
                    new A.MinorFont(new A.LatinFont { Typeface = "Calibri" }, new A.EastAsianFont { Typeface = "" }, new A.ComplexScriptFont { Typeface = "" }))
                { Name = "Office" },
                new A.FormatScheme(
                    new A.FillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.GradientFill(
                            new A.GradientStopList(
                                new A.GradientStop(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                                new A.GradientStop(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }),
                            new A.LinearGradientFill { Angle = 5400000, Scaled = true }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })),
                    new A.LineStyleList(
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 9525 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 25400 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 38100 }),
                    new A.EffectStyleList(
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList())),
                    new A.BackgroundFillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })))
                { Name = "Office" }),
            new A.ObjectDefaults(),
            new A.ExtraColorSchemeList())
        { Name = "Office Theme" };
}

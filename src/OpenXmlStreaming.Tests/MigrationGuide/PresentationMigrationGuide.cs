using DocumentFormat.OpenXml.Presentation;
// Drawing and Presentation share many type names (Shape, TextBody, Text,
// NonVisualDrawingProperties, ColorMap, …) so at least one of them has to
// be aliased. We default to Presentation and reach for Drawing.* explicitly.
using Drawing = DocumentFormat.OpenXml.Drawing;

public class PresentationMigrationGuide
{
    [Test]
    public async Task Standard()
    {
        using var stream = new MemoryStream();

// begin-snippet: migration-presentation-standard
        using (var doc = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presentationPart = doc.AddPresentationPart();
            presentationPart.Presentation = new();

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
            presentationPart.Presentation =
                new(
                    new SlideMasterIdList(
                        new SlideMasterId
                        {
                            Id = 2147483648U,
                            RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
                        }),
                    new SlideIdList(
                        new SlideId
                        {
                            Id = 256U,
                            RelationshipId = presentationPart.GetIdOfPart(slidePart)
                        }),
                    new SlideSize
                    {
                        Cx = 9144000,
                        Cy = 6858000
                    },
                    new NotesSize
                    {
                        Cx = 6858000,
                        Cy = 9144000
                    });
        }
// end-snippet

        stream.Position = 0;
        await Verify(stream, extension: "pptx");
    }

    [Test]
    public async Task Streaming()
    {
        using var stream = new MemoryStream();

// begin-snippet: migration-presentation-streaming
        await using (var writer = StreamingDocument.CreatePresentation(stream, leaveOpen: true))
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
                new Presentation(
                    new SlideMasterIdList(
                        new SlideMasterId
                        {
                            Id = 2147483648U,
                            RelationshipId = "rId1"
                        }),
                    new SlideIdList(
                        new SlideId
                        {
                            Id = 256U,
                            RelationshipId = "rId2"
                        }),
                    new SlideSize
                    {
                        Cx = 9144000,
                        Cy = 6858000
                    },
                    new NotesSize
                    {
                        Cx = 6858000,
                        Cy = 9144000
                    }),
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

        stream.Position = 0;
        await Verify(stream, extension: "pptx");
    }

    [Test]
    public async Task Builder()
    {
        using var stream = new MemoryStream();

        // begin-snippet: migration-presentation-builder
        await using (var presentation = new StreamingPresentationBuilder(stream, leaveOpen: true))
        {
            // No theme, slide master, or slide layout boilerplate — the
            // builder writes a minimal default scaffolding on the first
            // AddSlide call.
            presentation.AddSlide(BuildTitleSlide("Kickoff"));
        }
        // end-snippet

        stream.Position = 0;
        await Verify(stream, extension: "pptx");
    }

    static SlideMaster BuildSlideMaster() =>
        new(
            new CommonSlideData(
                new Background(
                    new BackgroundStyleReference(
                        new Drawing.SchemeColor
                        {
                            Val = Drawing.SchemeColorValues.PhColor
                        })
                    {
                        Index = 1001
                    }),
                new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties
                        {
                            Id = 1,
                            Name = ""
                        },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new Drawing.TransformGroup()))),
            new ColorMap
            {
                Background1 = Drawing.ColorSchemeIndexValues.Light1,
                Text1 = Drawing.ColorSchemeIndexValues.Dark1,
                Background2 = Drawing.ColorSchemeIndexValues.Light2,
                Text2 = Drawing.ColorSchemeIndexValues.Dark2,
                Accent1 = Drawing.ColorSchemeIndexValues.Accent1,
                Accent2 = Drawing.ColorSchemeIndexValues.Accent2,
                Accent3 = Drawing.ColorSchemeIndexValues.Accent3,
                Accent4 = Drawing.ColorSchemeIndexValues.Accent4,
                Accent5 = Drawing.ColorSchemeIndexValues.Accent5,
                Accent6 = Drawing.ColorSchemeIndexValues.Accent6,
                Hyperlink = Drawing.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = Drawing.ColorSchemeIndexValues.FollowedHyperlink
            },
            new SlideLayoutIdList(new SlideLayoutId
            {
                Id = 2147483649U,
                RelationshipId = "rId2"
            }));

    static SlideLayout BuildSlideLayout() =>
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
                    new GroupShapeProperties(new Drawing.TransformGroup()))),
            new ColorMapOverride(new Drawing.MasterColorMapping()))
        {
            Type = SlideLayoutValues.Title
        };

    static Slide BuildTitleSlide(string title) =>
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

    static Drawing.Theme BuildTheme() =>
        new(
            new Drawing.ThemeElements(
                new Drawing.ColorScheme(
                    new Drawing.Dark1Color(
                        new Drawing.SystemColor
                        {
                            Val = Drawing.SystemColorValues.WindowText,
                            LastColor = "000000"
                        }),
                    new Drawing.Light1Color(
                        new Drawing.SystemColor
                        {
                            Val = Drawing.SystemColorValues.Window,
                            LastColor = "FFFFFF"
                        }),
                    new Drawing.Dark2Color(
                        new Drawing.RgbColorModelHex
                        {
                            Val = "1F497D"
                        }),
                    new Drawing.Light2Color(
                        new Drawing.RgbColorModelHex
                        {
                            Val = "EEECE1"
                        }),
                    new Drawing.Accent1Color(
                        new Drawing.RgbColorModelHex
                        {
                            Val = "4F81BD"
                        }),
                    new Drawing.Accent2Color(
                        new Drawing.RgbColorModelHex
                        {
                            Val = "C0504D"
                        }),
                    new Drawing.Accent3Color(
                        new Drawing.RgbColorModelHex
                        {
                            Val = "9BBB59"
                        }),
                    new Drawing.Accent4Color(
                        new Drawing.RgbColorModelHex
                        {
                            Val = "8064A2"
                        }),
                    new Drawing.Accent5Color(
                        new Drawing.RgbColorModelHex
                        {
                            Val = "4BACC6"
                        }),
                    new Drawing.Accent6Color(
                        new Drawing.RgbColorModelHex
                        {
                            Val = "F79646"
                        }),
                    new Drawing.Hyperlink(
                        new Drawing.RgbColorModelHex
                        {
                            Val = "0000FF"
                        }),
                    new Drawing.FollowedHyperlinkColor(
                        new Drawing.RgbColorModelHex
                        {
                            Val = "800080"
                        }))
                {
                    Name = "Office"
                },
                new Drawing.FontScheme(
                    new Drawing.MajorFont(
                        new Drawing.LatinFont
                        {
                            Typeface = "Calibri"
                        }, new Drawing.EastAsianFont
                        {
                            Typeface = ""
                        }, new Drawing.ComplexScriptFont
                        {
                            Typeface = ""
                        }),
                    new Drawing.MinorFont(
                        new Drawing.LatinFont
                        {
                            Typeface = "Calibri"
                        }, new Drawing.EastAsianFont
                        {
                            Typeface = ""
                        }, new Drawing.ComplexScriptFont
                        {
                            Typeface = ""
                        }))
                {
                    Name = "Office"
                },
                new Drawing.FormatScheme(
                    new Drawing.FillStyleList(
                        new Drawing.SolidFill(
                            new Drawing.SchemeColor
                            {
                                Val = Drawing.SchemeColorValues.PhColor
                            }),
                        new Drawing.GradientFill(
                            new Drawing.GradientStopList(
                                new Drawing.GradientStop(
                                    new Drawing.SchemeColor
                                    {
                                        Val = Drawing.SchemeColorValues.PhColor
                                    })
                                {
                                    Position = 0
                                },
                                new Drawing.GradientStop(
                                    new Drawing.SchemeColor
                                    {
                                        Val = Drawing.SchemeColorValues.PhColor
                                    })
                                {
                                    Position = 100000
                                }),
                            new Drawing.LinearGradientFill
                            {
                                Angle = 5400000,
                                Scaled = true
                            }),
                        new Drawing.SolidFill(
                            new Drawing.SchemeColor
                            {
                                Val = Drawing.SchemeColorValues.PhColor
                            })),
                    new Drawing.LineStyleList(
                        new Drawing.Outline(new Drawing.SolidFill(
                            new Drawing.SchemeColor
                            {
                                Val = Drawing.SchemeColorValues.PhColor
                            }))
                        {
                            Width = 9525
                        },
                        new Drawing.Outline(new Drawing.SolidFill(
                            new Drawing.SchemeColor
                            {
                                Val = Drawing.SchemeColorValues.PhColor
                            }))
                        {
                            Width = 25400
                        },
                        new Drawing.Outline(new Drawing.SolidFill(
                            new Drawing.SchemeColor
                            {
                                Val = Drawing.SchemeColorValues.PhColor
                            }))
                        {
                            Width = 38100
                        }),
                    new Drawing.EffectStyleList(
                        new Drawing.EffectStyle(new Drawing.EffectList()),
                        new Drawing.EffectStyle(new Drawing.EffectList()),
                        new Drawing.EffectStyle(new Drawing.EffectList())),
                    new Drawing.BackgroundFillStyleList(
                        new Drawing.SolidFill(
                            new Drawing.SchemeColor
                            {
                                Val = Drawing.SchemeColorValues.PhColor
                            }),
                        new Drawing.SolidFill(
                            new Drawing.SchemeColor
                            {
                                Val = Drawing.SchemeColorValues.PhColor
                            }),
                        new Drawing.SolidFill(
                            new Drawing.SchemeColor
                            {
                                Val = Drawing.SchemeColorValues.PhColor
                            })))
                {
                    Name = "Office"
                }),
            new Drawing.ObjectDefaults(),
            new Drawing.ExtraColorSchemeList())
        {
            Name = "Office Theme"
        };
}

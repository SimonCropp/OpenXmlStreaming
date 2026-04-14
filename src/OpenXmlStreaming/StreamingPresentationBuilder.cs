using DocumentFormat.OpenXml.Presentation;
// Drawing and Presentation share many type names (Shape, TextBody, Text, …)
// so at least one of them has to be aliased.
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OpenXmlStreaming;

/// <summary>
/// Higher-level builder on top of <see cref="OpenXmlPackageWriter"/> for producing
/// <c>.pptx</c> presentations. Handles the mandatory scaffolding — one theme,
/// one slide master, one slide layout — automatically on first <see cref="AddSlide"/>
/// call, using minimal default content. Slides are written as you add them;
/// the main <c>ppt/presentation.xml</c> referencing all added slides is written
/// during <see cref="Finish"/>/<see cref="Dispose"/>/<see cref="DisposeAsync"/>.
/// </summary>
/// <remarks>
/// For a presentation that needs a custom theme, multiple slide masters, or
/// notes/handout masters, drop down to <see cref="OpenXmlPackageWriter"/> directly
/// and use the pattern shown in the migration guide.
/// </remarks>
/// <inheritdoc cref="OpenXmlPackageWriter(Stream, bool, int)"/>
public sealed class StreamingPresentationBuilder(
    Stream stream,
    bool leaveOpen = false,
    int bufferSize = OpenXmlPackageWriter.DefaultBufferSize) :
    IAsyncDisposable,
    IDisposable
{
    const string slideMasterRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
    const string slideLayoutRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    const string slideRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
    const string themeRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    const string slideContentType = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
    const string slideLayoutContentType = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
    const string slideMasterContentType = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
    const string themeContentType = "application/vnd.openxmlformats-officedocument.theme+xml";
    const string presentationContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

    OpenXmlPackageWriter writer = StreamingDocument.CreatePresentation(stream, leaveOpen, bufferSize);
    List<(string RelId, uint SlideId)> slides = [];
    bool scaffoldingWritten;
    bool finished;

    static Uri slideLayoutTargetFromSlide = new("../slideLayouts/slideLayout1.xml", UriKind.Relative);

    /// <summary>
    /// Writes a slide part to the package and records it for inclusion in the
    /// final <c>ppt/presentation.xml</c>. On the first call, also writes the
    /// default theme + slide master + slide layout scaffolding that every
    /// presentation needs.
    /// </summary>
    public void AddSlide(Slide slide)
    {
        ThrowIfFinished();

        EnsureScaffolding();

        var index = slides.Count + 1;
        var indexString = index.ToString(CultureInfo.InvariantCulture);
        var partUri = $"/ppt/slides/slide{indexString}.xml";

        writer.WritePart(
            new(partUri, UriKind.Relative),
            slideContentType,
            slide,
            [
                new(
                    slideLayoutTargetFromSlide,
                    slideLayoutRelType,
                    id: "rId1"),
            ]);

        // Slide ids in the presentation's SlideIdList must be >= 256 and unique.
        var slideRelId = "slide" + indexString;
        slides.Add((slideRelId, 256U + (uint)(index - 1)));
    }

    static Uri slideMasterTargetFromPresentation = new("slideMasters/slideMaster1.xml", UriKind.Relative);

    /// <summary>
    /// Writes <c>ppt/presentation.xml</c> referencing every slide and the slide
    /// master. Called automatically by <see cref="Dispose"/>/<see cref="DisposeAsync"/>.
    /// </summary>
    internal void Finish()
    {
        if (finished)
        {
            return;
        }

        finished = true;
        EnsureScaffolding();

        var slideIdList = new SlideIdList();
        var relationships = new List<PartRelationship>(slides.Count + 1)
        {
            new(
                slideMasterTargetFromPresentation,
                slideMasterRelType,
                id: "master"),
        };

        for (var i = 0; i < slides.Count; i++)
        {
            var (relId, slideId) = slides[i];
            slideIdList.AppendChild(
                new SlideId
                {
                    Id = slideId,
                    RelationshipId = relId
                });
            relationships.Add(
                new(
                    new($"slides/slide{(i + 1).ToString(CultureInfo.InvariantCulture)}.xml", UriKind.Relative),
                    slideRelType,
                    id: relId));
        }

        writer.WritePart(
            StreamingDocument.PresentationUri,
            presentationContentType,
            new Presentation(
                new SlideMasterIdList(
                    new SlideMasterId
                    {
                        Id = 2147483648U,
                        RelationshipId = "master"
                    }),
                slideIdList,
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
            relationships);
    }

    public void Dispose()
    {
        Finish();
        writer.Dispose();
    }

    public ValueTask DisposeAsync()
    {
        Finish();
        return writer.DisposeAsync();
    }

    void EnsureScaffolding()
    {
        if (scaffoldingWritten)
        {
            return;
        }

        scaffoldingWritten = true;

        // Theme — referenced from the slide master.
        writer.WritePart(
            new("/ppt/theme/theme1.xml", UriKind.Relative),
            themeContentType,
            BuildDefaultTheme());

        // Slide layout — referenced from the slide master AND each slide.
        writer.WritePart(
            new("/ppt/slideLayouts/slideLayout1.xml", UriKind.Relative),
            slideLayoutContentType,
            BuildDefaultLayout());

        // Slide master — points to theme + layout.
        writer.WritePart(
            new("/ppt/slideMasters/slideMaster1.xml", UriKind.Relative),
            slideMasterContentType,
            BuildDefaultMaster(),
            [
                new(
                    new("../theme/theme1.xml", UriKind.Relative),
                    themeRelType,
                    id: "rId1"),
                new(
                    new("../slideLayouts/slideLayout1.xml", UriKind.Relative),
                    slideLayoutRelType,
                    id: "rId2"),
            ]);
    }

    void ThrowIfFinished()
    {
        if (finished)
        {
            throw new InvalidOperationException("Presentation has already been finalized. No more slides can be added.");
        }
    }

    static SlideMaster BuildDefaultMaster() =>
        new(
            new CommonSlideData(
                new Background(
                    new BackgroundStyleReference(new Drawing.SchemeColor
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

    static SlideLayout BuildDefaultLayout() =>
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

    static Drawing.Theme BuildDefaultTheme() =>
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
                    new Drawing.Accent1Color(new Drawing.RgbColorModelHex
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
                        },
                        new Drawing.EastAsianFont
                        {
                            Typeface = ""
                        },
                        new Drawing.ComplexScriptFont
                        {
                            Typeface = ""
                        }),
                    new Drawing.MinorFont(
                        new Drawing.LatinFont
                        {
                            Typeface = "Calibri"
                        },
                        new Drawing.EastAsianFont
                        {
                            Typeface = ""
                        },
                        new Drawing.ComplexScriptFont
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
                        new Drawing.Outline(
                            new Drawing.SolidFill(
                                new Drawing.SchemeColor
                                {
                                    Val = Drawing.SchemeColorValues.PhColor
                                }))
                        {
                            Width = 9525
                        },
                        new Drawing.Outline(
                            new Drawing.SolidFill(
                                new Drawing.SchemeColor
                                {
                                    Val = Drawing.SchemeColorValues.PhColor
                                }))
                        {
                            Width = 25400
                        },
                        new Drawing.Outline(
                            new Drawing.SolidFill(
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

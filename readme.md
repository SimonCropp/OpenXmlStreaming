# <img src="/src/icon.png" height="30px"> OpenXmlStreaming

[![Build status](https://img.shields.io/appveyor/build/SimonCropp/OpenXmlStreaming)](https://ci.appveyor.com/project/SimonCropp/OpenXmlStreaming)
[![NuGet Status](https://img.shields.io/nuget/v/OpenXmlStreaming.svg?label=OpenXmlStreaming)](https://www.nuget.org/packages/OpenXmlStreaming/)

Forward-only writer for Office Open XML documents (`.docx`, `.xlsx`, `.pptx`). Writes directly to any writable stream — including non-seekable streams such as HTTP response bodies — without buffering the whole document in a `MemoryStream`.


## NuGet package

https://nuget.org/packages/OpenXmlStreaming/


## Why

`DocumentFormat.OpenXml` and `System.IO.Packaging` require a seekable stream because the underlying ZIP writer patches headers in place. That forces callers to either buffer the full package to a `MemoryStream` before flushing it to the network, or write to a temporary file. `OpenXmlStreaming` uses `ZipArchive` in `Create` mode, which emits ZIP data descriptors instead of back-patching, allowing true forward-only output.


## When to use

 * Writing documents to HTTP response streams (`HttpResponse.Body`)
 * Writing to network streams or cloud storage upload streams
 * Generating large documents where you want to avoid buffering the entire package in memory
 * Any scenario where the target stream is not seekable


## Comparison with the standard `DocumentFormat.OpenXml` API

| | Standard `XxxDocument.Create` | `OpenXmlPackageWriter` |
|---|---|---|
| Requires seekable stream | Yes | No |
| Requires `MemoryStream` buffer | Often | Never |
| Can modify parts after writing | Yes | No |
| Can read parts | Yes | No |
| DOM support | Full | Write-only |
| `OpenXmlWriter` support | Yes | Yes |
| Memory usage for large docs | Higher | Lower |


## Key behaviours

 * **Only one part can be open at a time.** Creating a new part auto-closes the previous one.
 * **Parts cannot be modified after writing.** This is a forward-only writer.
 * **Content types and package relationships are written last** (during `Dispose`/`DisposeAsync`), so they capture all parts that were written.
 * **The destination stream does not need to be seekable.** The writer uses `ZipArchive` in `Create` mode and emits ZIP data descriptors.
 * **`Dispose`/`DisposeAsync` finalizes the package.**


## Core API

### `OpenXmlPackageWriter`

The main writer class. Constructed directly, or via a typed factory that pre-registers the main part relationship.

<!-- snippet: construction-variants -->
<a id='snippet-construction-variants'></a>
```cs
// Direct construction
using var direct = new OpenXmlPackageWriter(stream, leaveOpen: true);

// Typed factories (pre-register the officeDocument relationship)
using var word = StreamingDocument.CreateWord(stream, leaveOpen: true);
using var spreadsheet = StreamingDocument.CreateSpreadsheet(stream, leaveOpen: true);
using var presentation = StreamingDocument.CreatePresentation(stream, leaveOpen: true);
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.Word.cs#L96-L104' title='Snippet source file'>snippet source</a> | <a href='#snippet-construction-variants' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

| Method | Description |
|---|---|
| `AddRelationship(partUri, relationshipType, id?)` | Adds a package-level relationship (written to `_rels/.rels`) |
| `CreatePart(partUri, contentType)` | Creates a part and returns an `OpenXmlPartEntry` for streaming writes |
| `WritePart(partUri, contentType, rootElement, relationships?)` | One-shot: writes an element tree as a complete part |
| `FlushAsync(cancellationToken?)` | Asynchronously flushes the internal write buffer to the target stream |
| `Dispose()` / `DisposeAsync()` | Finalizes the package: writes rels + content types, disposes the archive, flushes the buffer |

### `OpenXmlPartEntry`

Returned by `CreatePart`. Exposes the part's output `Stream` and an `AddRelationship` method for part-level relationships written to the matching `*.rels` file when the entry is disposed.

### `PartRelationship`

A struct passed to `WritePart` to declare part-level relationships inline:

<!-- snippet: part-relationship-struct -->
<a id='snippet-part-relationship-struct'></a>
```cs
var relationship = new PartRelationship(
    targetUri: new("styles.xml", UriKind.Relative),
    relationshipType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    // required — the part body almost always references its own
    // relationships by id, so the caller must know it up front
    id: "rId1",
    // default
    targetMode: TargetMode.Internal);
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.Word.cs#L110-L119' title='Snippet source file'>snippet source</a> | <a href='#snippet-part-relationship-struct' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


## High-level builders

The types in the previous section are a thin wrapper around OPC — they give you direct control over parts, relationships, and content types. For common cases that's more plumbing than callers want to write. Three higher-level builders sit on top of `OpenXmlPackageWriter` and handle part URIs, relationship ids, and the main-part composition for you:

 * [`StreamingWorkbookBuilder`](/src/OpenXmlStreaming/StreamingWorkbookBuilder.cs) — `xlsx` workbook with N worksheets.
 * [`StreamingWordDocumentBuilder`](/src/OpenXmlStreaming/StreamingWordDocumentBuilder.cs) — `docx` with optional styles, numbering, headers, and footers.
 * [`StreamingPresentationBuilder`](/src/OpenXmlStreaming/StreamingPresentationBuilder.cs) — `pptx` with slides and an auto-generated default theme/master/layout scaffolding.

Each builder inherits the streaming, buffering, and async-disposal behaviour of the underlying `OpenXmlPackageWriter` — parts are written as you add them, only a small list of relationship ids is held in memory between calls, and the final main part is flushed async on `await using` disposal.

### Workbook

<!-- snippet: workbook-builder -->
<a id='snippet-workbook-builder'></a>
```cs
await using var workbook = new StreamingWorkbookBuilder(stream, leaveOpen: true);

workbook.AddWorksheet(
    "Revenue",
    new(
        new SheetData(
            new Row(
                new Cell
                {
                    CellValue = new("Q1"),
                    DataType = CellValues.InlineString
                },
                new Cell
                {
                    CellValue = new("1000"),
                    DataType = CellValues.Number
                }),
            new Row(
                new Cell
                {
                    CellValue = new("Q2"),
                    DataType = CellValues.InlineString
                },
                new Cell
                {
                    CellValue = new("1200"),
                    DataType = CellValues.Number
                }))));

workbook.AddWorksheet(
    "Expenses",
    new(
        new SheetData(
            new Row(
                new Cell
                {
                    CellValue = new("Rent"),
                    DataType = CellValues.InlineString
                },
                new Cell
                {
                    CellValue = new("500"),
                    DataType = CellValues.Number
                }))));

// DisposeAsync (triggered by `await using`) writes xl/workbook.xml
// referencing both sheets — no manual rId wiring.
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.Spreadsheet.cs#L80-L128' title='Snippet source file'>snippet source</a> | <a href='#snippet-workbook-builder' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The builder generates worksheet URIs (`/xl/worksheets/sheetN.xml`) and matching `rIdN` relationship ids automatically. `DisposeAsync` writes `xl/workbook.xml` referencing every worksheet that was added, in the order they were added.


### Word document

<!-- snippet: word-document-builder -->
<a id='snippet-word-document-builder'></a>
```cs
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
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.Word.cs#L165-L220' title='Snippet source file'>snippet source</a> | <a href='#snippet-word-document-builder' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`AddStyles`/`AddNumbering`/`AddHeader`/`AddFooter` write the sub-part immediately and return the relationship id. For sub-parts that the document body needs to reference by id (`HeaderReference`, `FooterReference`), pass the returned string into the appropriate content element. `WriteDocument` writes the main `word/document.xml` last, with all accumulated relationships wired up — **it is explicit rather than dispose-triggered** because only the caller can produce a body that references the ids the builder hands out.


### Presentation

<!-- snippet: presentation-builder -->
<a id='snippet-presentation-builder'></a>
```cs
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
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.Presentation.cs#L26-L63' title='Snippet source file'>snippet source</a> | <a href='#snippet-presentation-builder' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The builder ships with a minimal default theme + slide master + slide layout. They're written lazily on the first `AddSlide` call, so an empty presentation still produces a structurally valid `.pptx`. For a presentation with a custom theme, multiple slide masters, or notes/handout masters, drop down to `OpenXmlPackageWriter` directly and follow the pattern in the migration guide.


## Usage


### Minimal Word document

<!-- snippet: minimal-word -->
<a id='snippet-minimal-word'></a>
```cs
using var writer = new OpenXmlPackageWriter(stream, leaveOpen: true);

writer.AddRelationship(
    new("/word/document.xml", UriKind.Relative),
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    "rId1");

writer.WritePart(
    new("/word/document.xml", UriKind.Relative),
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    new Document(new Body(new Paragraph(new Run(new Text("Hello!"))))));
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.Word.cs#L11-L23' title='Snippet source file'>snippet source</a> | <a href='#snippet-minimal-word' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Using the `StreamingDocument` factory

`StreamingDocument` wires up the standard `officeDocument` relationship for each document kind:

<!-- snippet: streaming-document-factory -->
<a id='snippet-streaming-document-factory'></a>
```cs
using var writer = StreamingDocument.CreateWord(stream, leaveOpen: true);

writer.WritePart(
    new("/word/document.xml", UriKind.Relative),
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    new Document(new Body(new Paragraph(new Run(new Text("Forward-only!"))))));
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.Word.cs#L31-L38' title='Snippet source file'>snippet source</a> | <a href='#snippet-streaming-document-factory' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`CreateSpreadsheet` and `CreatePresentation` are also provided.


### Streaming part content with `OpenXmlWriter`

For large parts, use `CreatePart` together with `DocumentFormat.OpenXml.OpenXmlWriter` to emit elements one at a time without materializing the whole tree:

<!-- snippet: streaming-part-content -->
<a id='snippet-streaming-part-content'></a>
```cs
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
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.Word.cs#L52-L64' title='Snippet source file'>snippet source</a> | <a href='#snippet-streaming-part-content' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Only one part may be open at a time. Creating a new part or disposing the writer automatically closes the current part.


### Part-level relationships

Relationships on a specific part can be passed to `WritePart` alongside the content:

<!-- snippet: part-relationships -->
<a id='snippet-part-relationships'></a>
```cs
writer.WritePart(
    new("/xl/workbook.xml", UriKind.Relative),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    new Workbook(
        new Sheets(
            new Sheet
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
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.Spreadsheet.cs#L11-L29' title='Snippet source file'>snippet source</a> | <a href='#snippet-part-relationships' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

External relationships (e.g. hyperlinks) are written by passing `TargetMode.External`:

<!-- snippet: external-relationship -->
<a id='snippet-external-relationship'></a>
```cs
entry.AddRelationship(
    new("https://example.com", UriKind.Absolute),
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    TargetMode.External,
    "rId1");
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.Word.cs#L82-L88' title='Snippet source file'>snippet source</a> | <a href='#snippet-external-relationship' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Presentations

<!-- snippet: create-presentation -->
<a id='snippet-create-presentation'></a>
```cs
using var writer = StreamingDocument.CreatePresentation(stream, leaveOpen: true);

writer.WritePart(
    new("/ppt/presentation.xml", UriKind.Relative),
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    new Presentation(new SlideIdList()));
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.Presentation.cs#L11-L18' title='Snippet source file'>snippet source</a> | <a href='#snippet-create-presentation' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Non-seekable streams

The writer works against any writable stream. This makes it suitable for writing straight to an HTTP response, compressed streams, or pipelines — no `MemoryStream` buffering required.


### Async disposal and remote sinks

For remote sinks where per-write network latency dominates — SQL BLOB streams, cloud upload streams, HTTP response bodies — use `await using` so the final flush goes through the async path:

<!-- snippet: async-usage -->
<a id='snippet-async-usage'></a>
```cs
await using var writer = StreamingDocument.CreateWord(stream, leaveOpen: true);

writer.WritePart(
    new("/word/document.xml", UriKind.Relative),
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    new Document(new Body(new Paragraph(new Run(new Text("Streamed async!"))))));

// DisposeAsync (triggered by `await using`) asynchronously flushes
// the final buffer — including the ZIP central directory — so remote
// sinks like SQL BLOB streams don't block the thread on network I/O.
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.Word.cs#L129-L140' title='Snippet source file'>snippet source</a> | <a href='#snippet-async-usage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Internally the writer wraps the target stream in a fixed-size buffer (default 80 KB). `ZipArchive` writes into the buffer synchronously as parts are produced; the buffer only hits the target stream when it fills, which batches many small deflate writes into a few larger ones. On `DisposeAsync`, the final buffer — which always contains the ZIP central directory and any trailing metadata — is pushed to the target via `Stream.WriteAsync`, so the calling thread is not blocked on the final network write.

The buffer size is configurable per writer. Pick a size matching your sink's preferred chunk size (SQL Server Large Value streams, Azure Blob block size, etc.):

<!-- snippet: custom-buffer-size -->
<a id='snippet-custom-buffer-size'></a>
```cs
// Bigger buffer = fewer, larger writes hit the sink — good for
// remote streams where per-write overhead is high. Pass 0 to
// disable buffering entirely and write straight to the sink.
using var writer = new OpenXmlPackageWriter(
    stream,
    leaveOpen: true,
    // 1 MB
    bufferSize: 1024 * 1024);
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.Word.cs#L148-L157' title='Snippet source file'>snippet source</a> | <a href='#snippet-custom-buffer-size' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`bufferSize: 0` disables buffering and writes directly to the target. This also disables async flushing on `DisposeAsync` — there's nothing left to flush — so use it only when the target is already a local/in-memory stream where the extra copy isn't worth it.

For progressive async flushing at part boundaries — e.g. "write a worksheet, push its bytes to the sink async, then write the next one" — call `FlushAsync` between parts:

<!-- snippet: flush-async -->
<a id='snippet-flush-async'></a>
```cs
// Write the worksheet, then push its bytes to the target stream
// asynchronously before moving on to the next part. Useful at part
// boundaries against remote sinks — the thread isn't blocked on
// network I/O while the next part is being serialized.
writer.WritePart(
    new("/xl/worksheets/sheet1.xml", UriKind.Relative),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
    new Worksheet(new SheetData()));

await writer.FlushAsync();

writer.WritePart(
    new("/xl/workbook.xml", UriKind.Relative),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    new Workbook(
        new Sheets(new Sheet
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
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.Spreadsheet.cs#L44-L72' title='Snippet source file'>snippet source</a> | <a href='#snippet-flush-async' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`FlushAsync` pushes whatever is currently sitting in the internal buffer to the target stream via `WriteAsync`. It's a no-op when the buffer is empty or when the writer is unbuffered (`bufferSize: 0`). It does **not** eliminate sync writes that spill from inside a single `WritePart` call when the part is larger than the buffer — those are an inherent consequence of `ZipArchive`'s sync write surface and the only mitigation is a bigger buffer.

Notes and limitations:

 * **XML serialization is synchronous.** `DocumentFormat.OpenXml`'s `OpenXmlElement.WriteTo(XmlWriter)` and `OpenXmlWriter` APIs are sync-only, and `ZipArchive` in Create mode calls sync `Write` on its underlying stream. The async surface exists at the writer's boundary — it cannot make the per-element serialization itself async.
 * **Intermediate writes may still be synchronous.** `ZipArchive` calls `Flush()` on its target stream at a few points, which drains the buffer synchronously. The async guarantee covers the final flush during `DisposeAsync` and explicit `FlushAsync` calls, not every write.
 * **Sync `Dispose` still works.** Calling `Dispose()` (or letting a `using` block dispose the writer) flushes the final buffer synchronously. Only switch to `await using` when you actually want the async flush behaviour.


### Finalization

`Dispose`/`DisposeAsync` finalizes the package: writes `_rels/.rels`, writes `[Content_Types].xml`, disposes the underlying `ZipArchive` (which emits the ZIP central directory), and — on the async path — flushes the remaining buffered bytes to the target via `WriteAsync`.


## Migration guide

Side-by-side ports of realistic documents from the standard `DocumentFormat.OpenXml` DOM API to `OpenXmlStreaming`. Each pair below produces the same document — the outputs are snapshotted with [Verify.OpenXml](https://github.com/VerifyTests/Verify.OpenXml) in [`MigrationGuide.cs`](/src/OpenXmlStreaming.Tests/MigrationGuide.cs) so both sides are guaranteed to produce valid, structurally identical files.


### General migration pattern

 * Replace `using var doc = XxxDocument.Create(...)` with `await using var writer = StreamingDocument.CreateXxx(...)`.
 * Every call to `mainPart.AddNewPart<T>()` plus its property assignment becomes one `writer.WritePart(partUri, contentType, element)` call.
 * Part-to-part relationships that the DOM wired up for you automatically (e.g. `StyleDefinitionsPart` → main document) become explicit `PartRelationship` entries passed to `WritePart`.
 * **Write dependencies before the parts that reference them.** Sub-parts first, main part last.
 * Part URIs are absolute (start with `/`); relationship targets are relative (from the perspective of the owning part's `_rels` file).


### Word — styled document with a separate styles part


#### Before (standard DOM API):

<details>
<summary>Standard `WordprocessingDocument` API</summary>

<!-- snippet: migration-word-standard -->
<a id='snippet-migration-word-standard'></a>
```cs
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
```
<sup><a href='/src/OpenXmlStreaming.Tests/MigrationGuide/WordMigrationGuide.cs#L10-L59' title='Snippet source file'>snippet source</a> | <a href='#snippet-migration-word-standard' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

</details>


#### After (low-level `OpenXmlPackageWriter`):

<!-- snippet: migration-word-streaming -->
<a id='snippet-migration-word-streaming'></a>
```cs
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
```
<sup><a href='/src/OpenXmlStreaming.Tests/MigrationGuide/WordMigrationGuide.cs#L70-L128' title='Snippet source file'>snippet source</a> | <a href='#snippet-migration-word-streaming' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### After (high-level `StreamingWordDocumentBuilder`):

<!-- snippet: migration-word-builder -->
<a id='snippet-migration-word-builder'></a>
```cs
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
```
<sup><a href='/src/OpenXmlStreaming.Tests/MigrationGuide/WordMigrationGuide.cs#L139-L187' title='Snippet source file'>snippet source</a> | <a href='#snippet-migration-word-builder' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Spreadsheet — workbook with multiple worksheets


#### Before (standard DOM API):

<details>
<summary>Standard `SpreadsheetDocument` API</summary>

<!-- snippet: migration-spreadsheet-standard -->
<a id='snippet-migration-spreadsheet-standard'></a>
```cs
using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
{
    var workbookPart = doc.AddWorkbookPart();
    var sheets = new Sheets();

    // Revenue sheet
    var revenuePart = workbookPart.AddNewPart<WorksheetPart>();
    revenuePart.Worksheet =
        new(
            new SheetData(
                new Row(
                    InlineString("A1", "Quarter"),
                    InlineString("B1", "Revenue"))
                {
                    RowIndex = 1
                },
                new Row(
                    InlineString("A2", "Q1"),
                    Number("B2", "1000"))
                {
                    RowIndex = 2
                },
                new Row(
                    InlineString("A3", "Q2"),
                    Number("B3", "1200"))
                {
                    RowIndex = 3
                }));
    sheets.AppendChild(
        new Sheet
        {
            Name = "Revenue",
            SheetId = 1,
            Id = workbookPart.GetIdOfPart(revenuePart)
        });

    // Expenses sheet
    var expensesPart = workbookPart.AddNewPart<WorksheetPart>();
    expensesPart.Worksheet =
        new(
            new SheetData(
                new Row(
                    InlineString("A1", "Category"),
                    InlineString("B1", "Amount"))
                {
                    RowIndex = 1
                },
                new Row(
                    InlineString("A2", "Rent"),
                    Number("B2", "500"))
                {
                    RowIndex = 2
                }));
    sheets.AppendChild(
        new Sheet
        {
            Name = "Expenses",
            SheetId = 2,
            Id = workbookPart.GetIdOfPart(expensesPart)
        });

    workbookPart.Workbook = new(sheets);
}
```
<sup><a href='/src/OpenXmlStreaming.Tests/MigrationGuide/SpreadsheetMigrationGuide.cs#L10-L74' title='Snippet source file'>snippet source</a> | <a href='#snippet-migration-spreadsheet-standard' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

</details>


#### After (low-level `OpenXmlPackageWriter`):

<!-- snippet: migration-spreadsheet-streaming -->
<a id='snippet-migration-spreadsheet-streaming'></a>
```cs
await using (var writer = StreamingDocument.CreateSpreadsheet(stream, leaveOpen: true))
{
    // Worksheets are written first — the workbook references them by id.
    writer.WritePart(
        new("/xl/worksheets/sheet1.xml", UriKind.Relative),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        new Worksheet(
            new SheetData(
                new Row(
                    InlineString("A1", "Quarter"),
                    InlineString("B1", "Revenue"))
                {
                    RowIndex = 1
                },
                new Row(
                    InlineString("A2", "Q1"),
                    Number("B2", "1000"))
                {
                    RowIndex = 2
                },
                new Row(
                    InlineString("A3", "Q2"),
                    Number("B3", "1200"))
                {
                    RowIndex = 3
                })));

    writer.WritePart(
        new("/xl/worksheets/sheet2.xml", UriKind.Relative),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        new Worksheet(
            new SheetData(
                new Row(
                    InlineString("A1", "Category"),
                    InlineString("B1", "Amount"))
                {
                    RowIndex = 1
                },
                new Row(
                    InlineString("A2", "Rent"),
                    Number("B2", "500"))
                {
                    RowIndex = 2
                })));

    // Then the workbook, with a relationship per worksheet.
    writer.WritePart(
        new("/xl/workbook.xml", UriKind.Relative),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        new Workbook(
            new Sheets(
                new Sheet
                {
                    Name = "Revenue",
                    SheetId = 1,
                    Id = "rId1"
                },
                new Sheet
                {
                    Name = "Expenses",
                    SheetId = 2,
                    Id = "rId2"
                })),
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
```
<sup><a href='/src/OpenXmlStreaming.Tests/MigrationGuide/SpreadsheetMigrationGuide.cs#L85-L160' title='Snippet source file'>snippet source</a> | <a href='#snippet-migration-spreadsheet-streaming' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### After (high-level `StreamingWorkbookBuilder`):

<!-- snippet: migration-spreadsheet-builder -->
<a id='snippet-migration-spreadsheet-builder'></a>
```cs
await using (var workbook = new StreamingWorkbookBuilder(stream, leaveOpen: true))
{
    workbook.AddWorksheet(
        "Revenue",
        new(
            new SheetData(
                new Row(
                    InlineString("A1", "Quarter"),
                    InlineString("B1", "Revenue"))
                {
                    RowIndex = 1
                },
                new Row(
                    InlineString("A2", "Q1"),
                    Number("B2", "1000"))
                {
                    RowIndex = 2
                },
                new Row(
                    InlineString("A3", "Q2"),
                    Number("B3", "1200"))
                {
                    RowIndex = 3
                })));

    workbook.AddWorksheet(
        "Expenses",
        new(
            new SheetData(
                new Row(
                    InlineString("A1", "Category"),
                    InlineString("B1", "Amount"))
                {
                    RowIndex = 1
                },
                new Row(
                    InlineString("A2", "Rent"),
                    Number("B2", "500"))
                {
                    RowIndex = 2
                })));
}
// DisposeAsync (triggered by the `await using` block) writes
// xl/workbook.xml referencing every worksheet. No sheet URIs or
// rIds to track.
```
<sup><a href='/src/OpenXmlStreaming.Tests/MigrationGuide/SpreadsheetMigrationGuide.cs#L171-L217' title='Snippet source file'>snippet source</a> | <a href='#snippet-migration-spreadsheet-builder' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Presentation — title slide with slide master, layout, and theme

A valid `.pptx` needs a theme, a slide master, and at least one slide layout in addition to the slides themselves. The DOM API wires the relationships between them for you; with the streaming writer every part is written explicitly in dependency order; the high-level builder ships with a default scaffolding so you only think about slides.


#### Before (standard DOM API):

<details>
<summary>Standard `PresentationDocument` API</summary>

<!-- snippet: migration-presentation-standard -->
<a id='snippet-migration-presentation-standard'></a>
```cs
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
```
<sup><a href='/src/OpenXmlStreaming.Tests/MigrationGuide/PresentationMigrationGuide.cs#L14-L64' title='Snippet source file'>snippet source</a> | <a href='#snippet-migration-presentation-standard' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

</details>


#### After (low-level `OpenXmlPackageWriter`):

<details>
<summary>`OpenXmlPackageWriter` (streaming)</summary>

<!-- snippet: migration-presentation-streaming -->
<a id='snippet-migration-presentation-streaming'></a>
```cs
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
```
<sup><a href='/src/OpenXmlStreaming.Tests/MigrationGuide/PresentationMigrationGuide.cs#L75-L156' title='Snippet source file'>snippet source</a> | <a href='#snippet-migration-presentation-streaming' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

</details>


#### After (high-level `StreamingPresentationBuilder`):

<!-- snippet: migration-presentation-builder -->
<a id='snippet-migration-presentation-builder'></a>
```cs
await using (var presentation = new StreamingPresentationBuilder(stream, leaveOpen: true))
{
    // No theme, slide master, or slide layout boilerplate — the
    // builder writes a minimal default scaffolding on the first
    // AddSlide call.
    presentation.AddSlide(BuildTitleSlide("Kickoff"));
}
```
<sup><a href='/src/OpenXmlStreaming.Tests/MigrationGuide/PresentationMigrationGuide.cs#L167-L175' title='Snippet source file'>snippet source</a> | <a href='#snippet-migration-presentation-builder' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


## Benchmarks

Measured on .NET 10.0 with `BenchmarkDotNet` + `MemoryDiagnoser`. Source in [`src/OpenXmlStreaming.Benchmarks`](/src/OpenXmlStreaming.Benchmarks). Run with:

```
dotnet run -c Release --project src/OpenXmlStreaming.Benchmarks -- --filter "*"
```

### Writer overhead (`ForwardOnlyBenchmarks`)

Writes each document to a discarding stream so the numbers reflect the writer's own CPU and allocation cost, not the cost of the sink. Each row is a pair of two benchmarks — standard `XxxDocument.Create` vs `StreamingDocument.CreateXxx` — with per-pair deltas computed manually.

#### Word

| Scenario | Approach | Mean | Allocated | Δ time | Δ alloc |
|---|---|---:|---:|---:|---:|
| Simple (1 paragraph) | Standard | 28.0 µs | 54 KB | | |
| | ForwardOnly | 31.3 µs | 27 KB | +12% | **-50%** |
| Medium (20 paragraphs) | Standard | 78.6 µs | 81 KB | | |
| | ForwardOnly | 59.8 µs | 53 KB | -24% | -35% |
| Complex (100 paragraphs + styles) | Standard | 543 µs | 367 KB | | |
| | ForwardOnly | 324 µs | 288 KB | **-40%** | -22% |

#### Spreadsheet

| Scenario | Approach | Mean | Allocated | Δ time | Δ alloc |
|---|---|---:|---:|---:|---:|
| Simple (1 sheet, 1 row) | Standard | 80.3 µs | 84 KB | | |
| | ForwardOnly | 55.3 µs | 46 KB | -31% | -46% |
| Medium (100 rows × 10 cols) | Standard | 786 µs | 810 KB | | |
| | ForwardOnly | 707 µs | 681 KB | -10% | -16% |
| Complex (3 sheets × 500 rows × 10 cols) | Standard | 19.8 ms | 10.7 MB | | |
| | ForwardOnly | 11.3 ms | 9.5 MB | **-43%** | -11% |

#### Presentation

| Scenario | Approach | Mean | Allocated | Δ time | Δ alloc |
|---|---|---:|---:|---:|---:|
| Simple (1 slide) | Standard | 118 µs | 85 KB | | |
| | ForwardOnly | 54.9 µs | 46 KB | **-54%** | -46% |
| Medium (10 slides) | Standard | 304 µs | 283 KB | | |
| | ForwardOnly | 194 µs | 160 KB | -36% | -43% |
| Complex (30 slides × 5 shapes) | Standard | 1.58 ms | 1.21 MB | | |
| | ForwardOnly | 1.29 ms | 829 KB | -18% | -31% |

### I/O scenarios (`IoScenarioBenchmarks`)

Real sinks with a large document (2,000 paragraphs for Word, 10,000 rows × 10 cols for Spreadsheet). The `Standard` path for a non-seekable sink models the idiomatic workaround: build the package into a `MemoryStream`, then `CopyTo` the destination.

| Sink | Document | Approach | Mean | Allocated | Δ time | Δ alloc |
|---|---|---|---:|---:|---:|---:|
| Non-seekable | Word | Standard | 9.3 ms | 6.40 MB | | |
| | | ForwardOnly | 6.8 ms | 4.86 MB | **-27%** | -24% |
| Non-seekable | Spreadsheet | Standard | 190 ms | 75.89 MB | | |
| | | ForwardOnly | 181 ms | 63.35 MB | -5% | -17% |
| File (disk) | Word | Standard | 27.0 ms | 6.38 MB | | |
| | | ForwardOnly | 17.9 ms | 4.86 MB | **-33%** | -24% |
| File (disk) | Spreadsheet | Standard | 156.9 ms | 75.39 MB | | |
| | | ForwardOnly | 148.5 ms | 63.35 MB | -5% | -16% |

### Takeaways

 * **Allocation savings are consistent across every scenario** — the streaming writer avoids the `MemoryStream` buffer and the SDK's package-management overhead, saving 11-50% depending on document size and shape.
 * **Biggest time wins are on small-to-medium documents** (Word, Presentation) where the SDK's per-package fixed cost dominates. Writer overhead is up to 54% lower.
 * **On huge element trees (100K+ cells) the time delta narrows** to ~5-10%. The cost of constructing the tree itself dominates, and the saved MemoryStream buffer is a smaller fraction of total work — but the memory savings remain proportional.
 * **One case is roughly neutral on time**: writing a single trivial paragraph to a Word document. Writer overhead is comparable to the SDK's — but allocations are still halved.


## Icon

https://thenounproject.com/icon/pattern-541289/

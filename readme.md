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
 * **Content types and package relationships are written last** (during `Finish`/`Dispose`), so they capture all parts that were written.
 * **The destination stream does not need to be seekable.** The writer uses `ZipArchive` in `Create` mode and emits ZIP data descriptors.
 * **`Dispose`/`DisposeAsync` finalizes the package.** You do not need to call `Finish` explicitly.


## Core API

### `OpenXmlPackageWriter`

The main writer class. Constructed directly, or via a typed factory that pre-registers the main part relationship.

<!-- snippet: construction-variants -->
<a id='snippet-construction-variants'></a>
```cs
// Direct construction
using var direct = new OpenXmlPackageWriter(stream, leaveOpen: true);

// Typed factories (pre-register the officeDocument relationship)
using var word = StreamingDocument.CreateWord(
    stream, WordprocessingDocumentType.Document, leaveOpen: true);
using var spreadsheet = StreamingDocument.CreateSpreadsheet(
    stream, SpreadsheetDocumentType.Workbook, leaveOpen: true);
using var presentation = StreamingDocument.CreatePresentation(
    stream, PresentationDocumentType.Presentation, leaveOpen: true);
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.cs#L153-L164' title='Snippet source file'>snippet source</a> | <a href='#snippet-construction-variants' title='Start of snippet'>anchor</a></sup>
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
    targetMode: TargetMode.Internal, // default
    id: "rId1"); // optional, auto-generated if null
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.cs#L170-L176' title='Snippet source file'>snippet source</a> | <a href='#snippet-part-relationship-struct' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


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
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.cs#L12-L24' title='Snippet source file'>snippet source</a> | <a href='#snippet-minimal-word' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Using the `StreamingDocument` factory

`StreamingDocument` wires up the standard `officeDocument` relationship for each document kind:

<!-- snippet: streaming-document-factory -->
<a id='snippet-streaming-document-factory'></a>
```cs
using var writer = StreamingDocument.CreateWord(
    stream,
    WordprocessingDocumentType.Document,
    leaveOpen: true);

writer.WritePart(
    new("/word/document.xml", UriKind.Relative),
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    new Document(new Body(new Paragraph(new Run(new Text("Forward-only!"))))));
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.cs#L32-L42' title='Snippet source file'>snippet source</a> | <a href='#snippet-streaming-document-factory' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.cs#L56-L68' title='Snippet source file'>snippet source</a> | <a href='#snippet-streaming-part-content' title='Start of snippet'>anchor</a></sup>
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
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.cs#L80-L98' title='Snippet source file'>snippet source</a> | <a href='#snippet-part-relationships' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.cs#L121-L127' title='Snippet source file'>snippet source</a> | <a href='#snippet-external-relationship' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Presentations

<!-- snippet: create-presentation -->
<a id='snippet-create-presentation'></a>
```cs
using var writer = StreamingDocument.CreatePresentation(
    stream,
    PresentationDocumentType.Presentation,
    leaveOpen: true);

writer.WritePart(
    new("/ppt/presentation.xml", UriKind.Relative),
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    new P.Presentation(new P.SlideIdList()));
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.cs#L135-L145' title='Snippet source file'>snippet source</a> | <a href='#snippet-create-presentation' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Non-seekable streams

The writer works against any writable stream. This makes it suitable for writing straight to an HTTP response, compressed streams, or pipelines — no `MemoryStream` buffering required.


### Async disposal and remote sinks

For remote sinks where per-write network latency dominates — SQL BLOB streams, cloud upload streams, HTTP response bodies — use `await using` so the final flush goes through the async path:

<!-- snippet: async-usage -->
<a id='snippet-async-usage'></a>
```cs
await using var writer = StreamingDocument.CreateWord(
    stream,
    WordprocessingDocumentType.Document,
    leaveOpen: true);

writer.WritePart(
    new("/word/document.xml", UriKind.Relative),
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    new Document(new Body(new Paragraph(new Run(new Text("Streamed async!"))))));

// DisposeAsync (triggered by `await using`) asynchronously flushes
// the final buffer — including the ZIP central directory — so remote
// sinks like SQL BLOB streams don't block the thread on network I/O.
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.cs#L186-L200' title='Snippet source file'>snippet source</a> | <a href='#snippet-async-usage' title='Start of snippet'>anchor</a></sup>
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
    bufferSize: 1024 * 1024); // 1 MB
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.cs#L208-L216' title='Snippet source file'>snippet source</a> | <a href='#snippet-custom-buffer-size' title='Start of snippet'>anchor</a></sup>
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
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.cs#L229-L257' title='Snippet source file'>snippet source</a> | <a href='#snippet-flush-async' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`FlushAsync` pushes whatever is currently sitting in the internal buffer to the target stream via `WriteAsync`. It's a no-op when the buffer is empty or when the writer is unbuffered (`bufferSize: 0`). It does **not** eliminate sync writes that spill from inside a single `WritePart` call when the part is larger than the buffer — those are an inherent consequence of `ZipArchive`'s sync write surface and the only mitigation is a bigger buffer.

Notes and limitations:

 * **XML serialization is synchronous.** `DocumentFormat.OpenXml`'s `OpenXmlElement.WriteTo(XmlWriter)` and `OpenXmlWriter` APIs are sync-only, and `ZipArchive` in Create mode calls sync `Write` on its underlying stream. The async surface exists at the writer's boundary — it cannot make the per-element serialization itself async.
 * **Intermediate writes may still be synchronous.** `ZipArchive` calls `Flush()` on its target stream at a few points, which drains the buffer synchronously. The async guarantee covers the final flush during `DisposeAsync` and explicit `FlushAsync` calls, not every write.
 * **Sync `Dispose` still works.** Calling `Dispose()` (or letting a `using` block dispose the writer) flushes the final buffer synchronously. Only switch to `await using` when you actually want the async flush behaviour.


### Finalization

`Dispose`/`DisposeAsync` finalizes the package: writes `_rels/.rels`, writes `[Content_Types].xml`, disposes the underlying `ZipArchive` (which emits the ZIP central directory), and — on the async path — flushes the remaining buffered bytes to the target via `WriteAsync`.


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

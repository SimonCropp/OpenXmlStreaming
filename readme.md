# <img src="/src/icon.png" height="30px"> OpenXmlStreaming

[![Build status](https://img.shields.io/appveyor/build/SimonCropp/OpenXmlStreaming)](https://ci.appveyor.com/project/SimonCropp/OpenXmlStreaming)
[![NuGet Status](https://img.shields.io/nuget/v/OpenXmlStreaming.svg?label=OpenXmlStreaming)](https://www.nuget.org/packages/OpenXmlStreaming/)

Forward-only writer for Office Open XML documents (`.docx`, `.xlsx`, `.pptx`). Writes directly to any writable stream — including non-seekable streams such as HTTP response bodies — without buffering the whole document in a `MemoryStream`.


## NuGet package

https://nuget.org/packages/OpenXmlStreaming/


## Why

`DocumentFormat.OpenXml` and `System.IO.Packaging` require a seekable stream because the underlying ZIP writer patches headers in place. That forces callers to either buffer the full package to a `MemoryStream` before flushing it to the network, or write to a temporary file. `OpenXmlStreaming` uses `ZipArchive` in `Create` mode, which emits ZIP data descriptors instead of back-patching, allowing true forward-only output.


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
            id: "rId1")
    ]);
```
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.cs#L80-L97' title='Snippet source file'>snippet source</a> | <a href='#snippet-part-relationships' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.cs#L120-L126' title='Snippet source file'>snippet source</a> | <a href='#snippet-external-relationship' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/OpenXmlStreaming.Tests/Samples.cs#L134-L144' title='Snippet source file'>snippet source</a> | <a href='#snippet-create-presentation' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Non-seekable streams

The writer works against any writable stream. This makes it suitable for writing straight to an HTTP response, compressed streams, or pipelines — no `MemoryStream` buffering required.


### Finalization

`Finish` writes `_rels/.rels` and `[Content_Types].xml`. It is called automatically by `Dispose`/`DisposeAsync`, but can be called explicitly if you need to flush before disposing.


## Icon

https://thenounproject.com/icon/pattern-541289/

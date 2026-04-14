# Forward-Only Document Writing

The `OpenXmlStreaming` library enables forward-only creation of OpenXML documents (.docx, .xlsx, .pptx) directly to any writable stream, including non-seekable streams. This eliminates the need for a temporary `MemoryStream` that the standard `DocumentFormat.OpenXml` `Create` methods require.

## When to Use

- Writing documents to HTTP response streams (`HttpResponse.Body`)
- Writing to network streams or cloud storage upload streams
- Generating large documents where you want to avoid buffering the entire package in memory
- Any scenario where the target stream is not seekable

## Quick Start

### Word Document

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlStreaming;

using var writer = StreamingDocument.CreateWord(
    outputStream, WordprocessingDocumentType.Document, leaveOpen: true);

writer.WritePart(
    new Uri("/word/document.xml", UriKind.Relative),
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    new Document(
        new Body(
            new Paragraph(
                new Run(new Text("Hello, forward-only world!"))))));
```

### Spreadsheet

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlStreaming;

using var writer = StreamingDocument.CreateSpreadsheet(
    outputStream, SpreadsheetDocumentType.Workbook);

// Write the workbook, referencing the worksheet
writer.WritePart(
    new Uri("/xl/workbook.xml", UriKind.Relative),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    new Workbook(new Sheets(
        new Sheet { Name = "Sheet1", SheetId = 1, Id = "rId1" })),
    relationships: new[]
    {
        new PartRelationship(
            new Uri("worksheets/sheet1.xml", UriKind.Relative),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            id: "rId1"),
    });

// Write the worksheet
writer.WritePart(
    new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
    new Worksheet(new SheetData(
        new Row(new Cell { CellValue = new CellValue("42"), DataType = CellValues.Number }))));
```

### Presentation

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OpenXmlStreaming;

using var writer = StreamingDocument.CreatePresentation(
    outputStream, PresentationDocumentType.Presentation);

writer.WritePart(
    new Uri("/ppt/presentation.xml", UriKind.Relative),
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    new Presentation(new SlideIdList()));
```

## Core API

### OpenXmlPackageWriter

The main writer class. Created via the constructor or typed factory methods on `StreamingDocument`.

```csharp
// Direct construction
var writer = new OpenXmlPackageWriter(stream, leaveOpen: false);

// Typed factory methods (pre-register main part relationship)
var writer = StreamingDocument.CreateWord(stream, type, leaveOpen);
var writer = StreamingDocument.CreateSpreadsheet(stream, type, leaveOpen);
var writer = StreamingDocument.CreatePresentation(stream, type, leaveOpen);
```

**Methods:**

| Method | Description |
|--------|-------------|
| `AddRelationship(partUri, relationshipType, id?)` | Adds a package-level relationship (written to `_rels/.rels`) |
| `CreatePart(partUri, contentType)` | Creates a part and returns an `OpenXmlPartEntry` for streaming writes |
| `WritePart(partUri, contentType, rootElement, relationships?)` | One-shot: writes an element tree as a complete part |
| `Finish()` | Finalizes the package (writes content types and relationships) |
| `Dispose()` / `DisposeAsync()` | Calls `Finish()` if needed, then closes the underlying ZIP |

### OpenXmlPartEntry

Returned by `CreatePart`. Provides access to the part's output stream for streaming XML writing.

```csharp
using var entry = writer.CreatePart(partUri, contentType);

// Add part-level relationships
entry.AddRelationship(targetUri, relationshipType, targetMode, id);

// Write XML content using OpenXmlWriter
using var xmlWriter = OpenXmlWriter.Create(entry.Stream);
xmlWriter.WriteStartDocument();
xmlWriter.WriteStartElement(new Worksheet());
xmlWriter.WriteStartElement(new SheetData());
foreach (var row in GetRows())
{
    xmlWriter.WriteElement(row);
}
xmlWriter.WriteEndElement(); // SheetData
xmlWriter.WriteEndElement(); // Worksheet
```

### PartRelationship

Used with `WritePart` to declare part-level relationships inline.

```csharp
new PartRelationship(
    targetUri: new Uri("styles.xml", UriKind.Relative),
    relationshipType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    targetMode: TargetMode.Internal, // default
    id: "rId1") // optional, auto-generated if null
```

## Key Behaviors

- **Only one part can be open at a time.** Creating a new part auto-closes the previous one.
- **Parts cannot be modified after writing.** This is a forward-only writer.
- **Content types and relationships are written last** (during `Finish()`/`Dispose()`), so they capture all parts.
- **The stream does not need to be seekable.** The writer uses `ZipArchive` in create mode.
- **`Dispose()` finalizes the package.** You don't need to call `Finish()` explicitly.

## Comparison with Standard API

| | Standard `Create` | `OpenXmlStreaming` |
|---|---|---|
| Requires seekable stream | Yes | No |
| Requires `MemoryStream` | Often | Never |
| Can modify parts after creation | Yes | No |
| Can read parts | Yes | No |
| DOM support | Full | Write-only |
| `OpenXmlWriter` support | Yes | Yes |
| Memory usage for large docs | Higher | Lower |

## Platform Support

Targets `net48` and `net10.0`.

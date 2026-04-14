using DocumentFormat.OpenXml.Wordprocessing;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
// ReSharper disable MethodHasAsyncOverload

[TestFixture]
public class OpenXmlPackageWriterTests
{
    [Test]
    public async Task WriteMinimalDocx_RoundTrips()
    {
        using var stream = new MemoryStream();

        await using (var writer = new OpenXmlPackageWriter(stream, leaveOpen: true))
        {
            writer.AddRelationship(
                new("/word/document.xml", UriKind.Relative),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                "rId1");

            writer.WritePart(
                new("/word/document.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                new Document(new Body(new Paragraph(new Run(new Text("Hello!"))))));
        }

        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.That(doc.MainDocumentPart, Is.Not.Null);
        Assert.That(doc.MainDocumentPart!.Document, Is.Not.Null);
        Assert.That(doc.MainDocumentPart.Document!.Body!.InnerText, Is.EqualTo("Hello!"));
        stream.Position = 0;

        await Verify(stream, extension: "docx");
    }

    [Test]
    public async Task CreateWord_RoundTrips()
    {
        using var stream = new MemoryStream();

        await using (var writer = StreamingDocument.CreateWord(stream, leaveOpen: true))
        {
            writer.WritePart(
                new("/word/document.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                new Document(new Body(new Paragraph(new Run(new Text("Forward-only!"))))));
        }

        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.That(doc.MainDocumentPart!.Document!.Body!.InnerText, Is.EqualTo("Forward-only!"));
        stream.Position = 0;

        await Verify(stream, extension: "docx");
    }

    [Test]
    public async Task WriteToNonSeekableStream()
    {
        using var stream = new MemoryStream();
        await using var nonSeekable = new NonSeekableStream(stream);

        await using (var writer = new OpenXmlPackageWriter(nonSeekable, leaveOpen: true))
        {
            writer.AddRelationship(
                new("/word/document.xml", UriKind.Relative),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                "rId1");

            writer.WritePart(
                new("/word/document.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                new Document(new Body(new Paragraph(new Run(new Text("Non-seekable!"))))));
        }

        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.That(doc.MainDocumentPart!.Document!.Body!.InnerText, Is.EqualTo("Non-seekable!"));
        stream.Position = 0;

        await Verify(stream, extension: "docx");
    }

    [Test]
    public async Task CreatePart_WithOpenXmlWriter()
    {
        using var stream = new MemoryStream();

        await using (var writer = new OpenXmlPackageWriter(stream, leaveOpen: true))
        {
            writer.AddRelationship(
                new("/word/document.xml", UriKind.Relative),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                "rId1");

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
        }

        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.That(doc.MainDocumentPart!.Document!.Body!.InnerText, Is.EqualTo("Streamed!"));
        stream.Position = 0;

        await Verify(stream, extension: "docx");
    }

    [Test]
    public async Task PartRelationships_AreWritten()
    {
        using var stream = new MemoryStream();

        await using (var writer = new OpenXmlPackageWriter(stream, leaveOpen: true))
        {
            writer.AddRelationship(
                new("/word/document.xml", UriKind.Relative),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                "rId1");

            using (var entry = writer.CreatePart(
                       new("/word/document.xml", UriKind.Relative),
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"))
            {
                entry.AddRelationship(
                    new("styles.xml", UriKind.Relative),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                    id: "rId1");

                using var xmlWriter = OpenXmlWriter.Create(entry.Stream);
                xmlWriter.WriteStartDocument();
                xmlWriter.WriteStartElement(new Document());
                xmlWriter.WriteStartElement(new Body());
                xmlWriter.WriteEndElement();
                xmlWriter.WriteEndElement();
            }

            writer.WritePart(
                new("/word/styles.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
                new Styles());
        }

        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var rels = doc.MainDocumentPart!.GetPartsOfType<StyleDefinitionsPart>();
        Assert.That(rels, Is.Not.Empty);
        stream.Position = 0;

        await Verify(stream, extension: "docx");
    }

    [Test]
    public void DuplicatePartUri_Throws()
    {
        using var stream = new MemoryStream();
        using var writer = new OpenXmlPackageWriter(stream);

        writer.WritePart(
            new("/word/document.xml", UriKind.Relative),
            "application/xml",
            new Document());

        Assert.Throws<InvalidOperationException>(() =>
            writer.WritePart(
                new("/word/document.xml", UriKind.Relative),
                "application/xml",
                new Document()));
    }

    [Test]
    public void OperationsAfterFinish_Throw()
    {
        using var stream = new MemoryStream();
        var writer = new OpenXmlPackageWriter(stream, leaveOpen: true);
        writer.Finish();

        Assert.Throws<InvalidOperationException>(() =>
            writer.AddRelationship(new("/foo.xml", UriKind.Relative), "type"));

        Assert.Throws<InvalidOperationException>(() =>
            writer.CreatePart(new("/foo.xml", UriKind.Relative), "type"));
    }

    [Test]
    public void AutoDisposesPreviousPartEntry()
    {
        using var stream = new MemoryStream();

        using var writer = new OpenXmlPackageWriter(stream, leaveOpen: true);

        writer.AddRelationship(
            new("/word/document.xml", UriKind.Relative),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "rId1");

        var entry1 = writer.CreatePart(
            new("/word/document.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");

        using (var xmlWriter = OpenXmlWriter.Create(entry1.Stream))
        {
            xmlWriter.WriteStartDocument();
            xmlWriter.WriteStartElement(new Document());
            xmlWriter.WriteEndElement();
        }

        var entry2 = writer.CreatePart(
            new("/word/styles.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");

        Assert.Throws<ObjectDisposedException>(() => _ = entry1.Stream);

        entry2.Dispose();
    }

    [Test]
    public async Task CreateSpreadsheet_RoundTrips()
    {
        using var stream = new MemoryStream();

        await using (var writer = StreamingDocument.CreateSpreadsheet(stream, leaveOpen: true))
        {
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

            writer.WritePart(
                new("/xl/worksheets/sheet1.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                new S.Worksheet(new S.SheetData(
                    new S.Row(new S.Cell
                    {
                        CellValue = new("42"),
                        DataType = S.CellValues.Number
                    }))));
        }

        stream.Position = 0;

        using var doc = SpreadsheetDocument.Open(stream, false);
        Assert.That(doc.WorkbookPart!.Workbook!.Sheets!, Is.Not.Empty);
        stream.Position = 0;

        await Verify(stream, extension: "xlsx");
    }

    [Test]
    public async Task CreatePresentation_RoundTrips()
    {
        using var stream = new MemoryStream();

        await using (var writer = StreamingDocument.CreatePresentation(stream, leaveOpen: true))
        {
            writer.WritePart(
                new("/ppt/presentation.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                new P.Presentation(new P.SlideIdList()));
        }

        stream.Position = 0;

        using var doc = PresentationDocument.Open(stream, false);
        Assert.That(doc.PresentationPart!.Presentation, Is.Not.Null);
        stream.Position = 0;

        await Verify(stream, extension: "pptx");
    }

    [Test]
    public async Task ExternalRelationship_IsWritten()
    {
        using var stream = new MemoryStream();

        await using (var writer = new OpenXmlPackageWriter(stream, leaveOpen: true))
        {
            writer.AddRelationship(
                new("/word/document.xml", UriKind.Relative),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                "rId1");

            using var entry = writer.CreatePart(
                new("/word/document.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");

            entry.AddRelationship(
                new("https://example.com", UriKind.Absolute),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                TargetMode.External,
                "rId1");

            using var xmlWriter = OpenXmlWriter.Create(entry.Stream);
            xmlWriter.WriteStartDocument();
            xmlWriter.WriteStartElement(new Document());
            xmlWriter.WriteStartElement(new Body());
            xmlWriter.WriteEndElement();
            xmlWriter.WriteEndElement();
        }

        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.That(doc.MainDocumentPart!.HyperlinkRelationships, Is.Not.Empty);
        stream.Position = 0;

        await Verify(stream, extension: "docx");
    }

    [Test]
    public async Task PartEntry_RootLevelPart_WritesRelsAtPackageRoot()
    {
        using var stream = new MemoryStream();

        await using (var writer = new OpenXmlPackageWriter(stream, leaveOpen: true))
        {
            using var entry = writer.CreatePart(new("/root.xml", UriKind.Relative), "application/xml");
            entry.AddRelationship(new("other.xml", UriKind.Relative), "type", id: "rId1");
        }

        stream.Position = 0;
        await using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var relsEntry = archive.GetEntry("_rels/root.xml.rels");
        Assert.That(relsEntry, Is.Not.Null);

        await using var relsStream = relsEntry!.Open();
        using var reader = new StreamReader(relsStream);
        var xml = await reader.ReadToEndAsync();
        Assert.That(xml, Does.Contain("Target=\"other.xml\""));
    }

    [Test]
    public async Task AddRelationship_AutoGeneratesId()
    {
        using var stream = new MemoryStream();
        await using var writer = new OpenXmlPackageWriter(stream);
        var id = writer.AddRelationship(new("/foo.xml", UriKind.Relative), "type");
        Assert.That(id, Does.StartWith("rId"));
    }

    [Test]
    public async Task PartEntry_AddRelationship_AutoGeneratesId()
    {
        using var stream = new MemoryStream();
        await using var writer = new OpenXmlPackageWriter(stream);
        using var entry = writer.CreatePart(new("/foo.xml", UriKind.Relative), "text/xml");
        var id = entry.AddRelationship(new("bar.xml", UriKind.Relative), "type");
        Assert.That(id, Does.StartWith("rId"));
    }

    [Test]
    public void PartEntry_AddRelationship_AfterDispose_Throws()
    {
        using var stream = new MemoryStream();
        using var writer = new OpenXmlPackageWriter(stream);
        var entry = writer.CreatePart(new("/foo.xml", UriKind.Relative), "text/xml");
        entry.Dispose();
        Assert.Throws<ObjectDisposedException>(() =>
            entry.AddRelationship(new("bar.xml", UriKind.Relative), "type"));
    }

    [Test]
    public void FinishCalledTwice_DoesNotThrow()
    {
        using var stream = new MemoryStream();
        var writer = new OpenXmlPackageWriter(stream, leaveOpen: true);
        writer.Finish();
        writer.Finish();
        writer.Dispose();
    }

    [Test]
    public void PartRelationship_Properties()
    {
        var uri = new Uri("foo.xml", UriKind.Relative);
        var rel = new PartRelationship(uri, "type", TargetMode.External, "rId1");
        Assert.That(rel.TargetUri, Is.EqualTo(uri));
        Assert.That(rel.RelationshipType, Is.EqualTo("type"));
        Assert.That(rel.TargetMode, Is.EqualTo(TargetMode.External));
        Assert.That(rel.Id, Is.EqualTo("rId1"));
    }

    [Test]
    public void PartRelationship_DefaultValues()
    {
        var rel = new PartRelationship(new("foo.xml", UriKind.Relative), "type");
        Assert.That(rel.TargetMode, Is.EqualTo(TargetMode.Internal));
        Assert.That(rel.Id, Is.Null);
    }

    [Test]
    public async Task DisposeAsync_FlushesFinalBufferAsynchronously()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        await using (var writer = new OpenXmlPackageWriter(tracker, leaveOpen: true))
        {
            writer.AddRelationship(
                new("/word/document.xml", UriKind.Relative),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                "rId1");

            writer.WritePart(
                new("/word/document.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                new Document(new Body(new Paragraph(new Run(new Text("Async!"))))));
        }

        // DisposeAsync must flush the tail of the buffer (at minimum the ZIP
        // central directory) via WriteAsync — that's the whole point of the
        // async surface. ZipArchive may also trigger some intermediate sync
        // flushes via Flush() calls, so we don't assert SyncWriteCalls == 0.
        Assert.Multiple(() =>
        {
            Assert.That(tracker.AsyncWriteCalls, Is.GreaterThanOrEqualTo(1),
                "DisposeAsync should flush via WriteAsync at least once");
            Assert.That(tracker.TotalBytesWritten, Is.GreaterThan(0));
        });

        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.That(doc.MainDocumentPart!.Document!.Body!.InnerText, Is.EqualTo("Async!"));
    }

    [Test]
    public void Dispose_FlushesFinalBufferSynchronously()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        using (var writer = new OpenXmlPackageWriter(tracker, leaveOpen: true))
        {
            writer.AddRelationship(
                new("/word/document.xml", UriKind.Relative),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                "rId1");

            writer.WritePart(
                new("/word/document.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                new Document(new Body(new Paragraph(new Run(new Text("Sync!"))))));
        }

        // Sync disposal routes the final flush through sync Write.
        Assert.Multiple(() =>
        {
            Assert.That(tracker.SyncWriteCalls, Is.GreaterThanOrEqualTo(1),
                "Sync Dispose should flush via Write");
            Assert.That(tracker.AsyncWriteCalls, Is.Zero,
                "Sync Dispose should not touch the async path");
        });
    }

    [Test]
    public async Task BufferSize_Zero_WritesReachTargetDuringWritePart()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        await using var writer = new OpenXmlPackageWriter(tracker, leaveOpen: true, bufferSize: 0);

        writer.AddRelationship(
            new("/word/document.xml", UriKind.Relative),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "rId1");

        writer.WritePart(
            new("/word/document.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            new Document(new Body(new Paragraph(new Run(new Text("Unbuffered!"))))));

        // With no buffer, ZipArchive's writes land on the target immediately
        // during WritePart — they're not deferred to DisposeAsync. This is
        // the distinguishing property of bufferSize: 0.
        Assert.That(tracker.TotalBytesWritten, Is.GreaterThan(0),
            "Writes should reach the target during WritePart, not deferred");
    }

    [Test]
    public async Task BufferSize_Zero_RoundTrips()
    {
        using var stream = new MemoryStream();

        await using (var writer = new OpenXmlPackageWriter(stream, leaveOpen: true, bufferSize: 0))
        {
            writer.AddRelationship(
                new("/word/document.xml", UriKind.Relative),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                "rId1");

            writer.WritePart(
                new("/word/document.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                new Document(new Body(new Paragraph(new Run(new Text("Unbuffered!"))))));
        }

        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.That(doc.MainDocumentPart!.Document!.Body!.InnerText, Is.EqualTo("Unbuffered!"));
    }

    [Test]
    public async Task BufferSize_Small_SpillsSyncThenAsyncFlush()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        // 64-byte buffer — far smaller than any part's payload — forces
        // sync spill flushes while ZipArchive is writing, then a final
        // async flush when DisposeAsync runs.
        await using (var writer = new OpenXmlPackageWriter(tracker, leaveOpen: true, bufferSize: 64))
        {
            writer.AddRelationship(
                new("/word/document.xml", UriKind.Relative),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                "rId1");

            writer.WritePart(
                new("/word/document.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                new Document(new Body(new Paragraph(new Run(new Text("Spilled!"))))));
        }

        Assert.Multiple(() =>
        {
            Assert.That(tracker.SyncWriteCalls, Is.GreaterThan(0),
                "Small buffer should spill sync while writing");
            Assert.That(tracker.AsyncWriteCalls, Is.GreaterThanOrEqualTo(1),
                "Final flush during DisposeAsync should still be async");
        });

        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.That(doc.MainDocumentPart!.Document!.Body!.InnerText, Is.EqualTo("Spilled!"));
    }

    [Test]
    public async Task FlushAsync_PushesBufferedBytesViaWriteAsync()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        await using var writer = new OpenXmlPackageWriter(tracker, leaveOpen: true);
        writer.AddRelationship(
            new("/word/document.xml", UriKind.Relative),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "rId1");

        writer.WritePart(
            new("/word/document.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            new Document(new Body(new Paragraph(new Run(new Text("First!"))))));

        var asyncBefore = tracker.AsyncWriteCalls;
        await writer.FlushAsync();
        var asyncAfter = tracker.AsyncWriteCalls;

        Assert.That(asyncAfter, Is.GreaterThan(asyncBefore),
            "FlushAsync should push at least one async write to the target");
    }

    [Test]
    public async Task FlushAsync_EmptyBuffer_IsNoop()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        await using var writer = new OpenXmlPackageWriter(tracker, leaveOpen: true);
        // Nothing has been written yet — buffer is empty.

        var beforeSync = tracker.SyncWriteCalls;
        var beforeAsync = tracker.AsyncWriteCalls;

        await writer.FlushAsync();

        Assert.Multiple(() =>
        {
            Assert.That(tracker.SyncWriteCalls, Is.EqualTo(beforeSync));
            Assert.That(tracker.AsyncWriteCalls, Is.EqualTo(beforeAsync));
        });
    }

    [Test]
    public async Task FlushAsync_Unbuffered_IsNoop()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        await using var writer = new OpenXmlPackageWriter(tracker, leaveOpen: true, bufferSize: 0);
        writer.AddRelationship(
            new("/word/document.xml", UriKind.Relative),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "rId1");

        writer.WritePart(
            new("/word/document.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            new Document(new Body()));

        var beforeAsync = tracker.AsyncWriteCalls;
        await writer.FlushAsync();

        Assert.That(tracker.AsyncWriteCalls, Is.EqualTo(beforeAsync),
            "FlushAsync with bufferSize: 0 must be a no-op — nothing to flush");
    }

    [Test]
    public async Task FlushAsync_BetweenParts_PackageStillRoundTrips()
    {
        using var stream = new MemoryStream();

        await using (var writer = new OpenXmlPackageWriter(stream, leaveOpen: true))
        {
            writer.AddRelationship(
                new("/word/document.xml", UriKind.Relative),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                "rId1");

            writer.WritePart(
                new("/word/document.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                new Document(new Body(new Paragraph(new Run(new Text("Flushed between!"))))),
                [
                    new(
                        new("styles.xml", UriKind.Relative),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                        id: "rId1"),
                ]);

            await writer.FlushAsync();

            writer.WritePart(
                new("/word/styles.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
                new Styles());
        }

        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.That(doc.MainDocumentPart!.Document!.Body!.InnerText, Is.EqualTo("Flushed between!"));
        Assert.That(doc.MainDocumentPart.GetPartsOfType<StyleDefinitionsPart>(), Is.Not.Empty);
    }

    [Test]
    public async Task DisposeAsync_LeaveOpen_DoesNotDisposeUnderlying()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        await using (var writer = new OpenXmlPackageWriter(tracker, leaveOpen: true))
        {
            writer.AddRelationship(
                new("/word/document.xml", UriKind.Relative),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                "rId1");

            writer.WritePart(
                new("/word/document.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                new Document(new Body()));
        }

        // Must still be writable after async dispose.
        Assert.DoesNotThrow(() => stream.WriteByte(0));
    }

    [Test]
    public async Task DisposeAsync_FinalizesPackage()
    {
        using var stream = new MemoryStream();

        await using (var writer = new OpenXmlPackageWriter(stream, leaveOpen: true))
        {
            writer.AddRelationship(
                new("/word/document.xml", UriKind.Relative),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                "rId1");

            writer.WritePart(
                new("/word/document.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                new Document(new Body()));
        }

        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.That(doc.MainDocumentPart, Is.Not.Null);

        stream.Position = 0;

        await Verify(stream, extension: "docx");
    }
}

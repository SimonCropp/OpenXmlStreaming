using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlStreaming.Benchmarks;

internal static class IoScenarioContent
{
    public const int LargeWordParagraphs = 2000;
    public const uint LargeSpreadsheetRows = 10_000;
    public const int LargeSpreadsheetCols = 10;

    public static Body BuildLargeWordBody()
    {
        var body = new Body();

        for (var i = 0; i < LargeWordParagraphs; i++)
        {
            body.AppendChild(new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId { Val = "Heading1" }),
                new W.Run(
                    new W.RunProperties(
                        new W.Bold(),
                        new W.FontSize { Val = "28" }),
                    new W.Text("Section " + i)),
                new W.Run(new W.Break()),
                new W.Run(
                    new W.Text("Content for section " + i + " with some additional text to make it realistic."))));
        }

        return body;
    }

    public static SheetData BuildLargeSheetData()
    {
        var sheetData = new SheetData();

        for (uint r = 1; r <= LargeSpreadsheetRows; r++)
        {
            var row = new Row { RowIndex = r };

            for (var c = 0; c < LargeSpreadsheetCols; c++)
            {
                row.AppendChild(new Cell
                {
                    CellValue = new CellValue(((r * LargeSpreadsheetCols) + (uint)c).ToString()),
                    DataType = CellValues.Number,
                });
            }

            sheetData.AppendChild(row);
        }

        return sheetData;
    }

    public static void WriteSpreadsheetStandard(Stream target)
    {
        using var doc = SpreadsheetDocument.Create(target, SpreadsheetDocumentType.Workbook);

        var workbookPart = doc.AddWorkbookPart();
        var sheetPart = workbookPart.AddNewPart<WorksheetPart>();
        sheetPart.Worksheet = new Worksheet(BuildLargeSheetData());
        workbookPart.Workbook = new Workbook(
            new Sheets(
                new Sheet { Name = "Sheet1", SheetId = 1, Id = workbookPart.GetIdOfPart(sheetPart) }));
    }

    public static void WriteSpreadsheetForwardOnly(Stream target)
    {
        using var writer = StreamingDocument.CreateSpreadsheet(target, SpreadsheetDocumentType.Workbook);

        writer.WritePart(
            new("/xl/worksheets/sheet1.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            new Worksheet(BuildLargeSheetData()));

        writer.WritePart(
            new("/xl/workbook.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            new Workbook(
                new Sheets(
                    new Sheet { Name = "Sheet1", SheetId = 1, Id = "rId1" })),
            [
                new(
                    new("worksheets/sheet1.xml", UriKind.Relative),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                    id: "rId1"),
            ]);
    }

    public static void WriteWordStandard(Stream target)
    {
        using var doc = WordprocessingDocument.Create(target, WordprocessingDocumentType.Document);
        doc.AddMainDocumentPart().Document = new Document(BuildLargeWordBody());
    }

    public static void WriteWordForwardOnly(Stream target)
    {
        using var writer = StreamingDocument.CreateWord(target, WordprocessingDocumentType.Document);

        writer.WritePart(
            new("/word/document.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            new Document(BuildLargeWordBody()));
    }
}

// ─── Non-seekable sink ────────────────────────────────────────────────────
// The headline scenario: writing a large document to a non-seekable sink
// (e.g. HTTP response body). The SDK path has to buffer the whole package
// into a MemoryStream first; the streaming path writes directly.

[MemoryDiagnoser]
public class NonSeekableWordBenchmarks
{
    [Benchmark(Baseline = true)]
    public void Standard()
    {
        using var sink = new NonSeekableDiscardStream();
        using var buffer = new MemoryStream();

        IoScenarioContent.WriteWordStandard(buffer);

        buffer.Position = 0;
        buffer.CopyTo(sink);
    }

    [Benchmark]
    public void ForwardOnly()
    {
        using var sink = new NonSeekableDiscardStream();
        IoScenarioContent.WriteWordForwardOnly(sink);
    }
}

[MemoryDiagnoser]
public class NonSeekableSpreadsheetBenchmarks
{
    [Benchmark(Baseline = true)]
    public void Standard()
    {
        using var sink = new NonSeekableDiscardStream();
        using var buffer = new MemoryStream();

        IoScenarioContent.WriteSpreadsheetStandard(buffer);

        buffer.Position = 0;
        buffer.CopyTo(sink);
    }

    [Benchmark]
    public void ForwardOnly()
    {
        using var sink = new NonSeekableDiscardStream();
        IoScenarioContent.WriteSpreadsheetForwardOnly(sink);
    }
}

// ─── File on disk ─────────────────────────────────────────────────────────
// Both paths are legal against a seekable FileStream. This measures whether
// the SDK's in-package buffering adds visible overhead vs writing direct.

[MemoryDiagnoser]
public class FileWordBenchmarks
{
    string tempFile = string.Empty;

    [IterationSetup]
    public void Setup() =>
        tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".tmp");

    [IterationCleanup]
    public void Cleanup()
    {
        if (File.Exists(tempFile))
        {
            File.Delete(tempFile);
        }
    }

    [Benchmark(Baseline = true)]
    public void Standard()
    {
        using var stream = File.Create(tempFile);
        IoScenarioContent.WriteWordStandard(stream);
    }

    [Benchmark]
    public void ForwardOnly()
    {
        using var stream = File.Create(tempFile);
        IoScenarioContent.WriteWordForwardOnly(stream);
    }
}

[MemoryDiagnoser]
public class FileSpreadsheetBenchmarks
{
    string tempFile = string.Empty;

    [IterationSetup]
    public void Setup() =>
        tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".tmp");

    [IterationCleanup]
    public void Cleanup()
    {
        if (File.Exists(tempFile))
        {
            File.Delete(tempFile);
        }
    }

    [Benchmark(Baseline = true)]
    public void Standard()
    {
        using var stream = File.Create(tempFile);
        IoScenarioContent.WriteSpreadsheetStandard(stream);
    }

    [Benchmark]
    public void ForwardOnly()
    {
        using var stream = File.Create(tempFile);
        IoScenarioContent.WriteSpreadsheetForwardOnly(stream);
    }
}

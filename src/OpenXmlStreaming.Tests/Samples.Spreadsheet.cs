using DocumentFormat.OpenXml.Spreadsheet;

public partial class Samples
{
    [Test]
    public void PartRelationships()
    {
        using var stream = new MemoryStream();
        using var writer = StreamingDocument.CreateSpreadsheet(stream, leaveOpen: true);

        // begin-snippet: part-relationships
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
        // end-snippet

        writer.WritePart(
            new("/xl/worksheets/sheet1.xml", UriKind.Relative),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            new Worksheet(new SheetData()));
    }

    [Test]
    public async Task FlushBetweenParts()
    {
        using var stream = new MemoryStream();

        await using var writer = StreamingDocument.CreateSpreadsheet(stream, leaveOpen: true);

        // begin-snippet: flush-async
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
        // end-snippet
    }

    [Test]
    public async Task WorkbookBuilderSample()
    {
        using var stream = new MemoryStream();

        // begin-snippet: workbook-builder
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
        // end-snippet
    }
}

using DocumentFormat.OpenXml.Spreadsheet;

public class WorkbookBuilderTests
{
    [Test]
    public async Task RoundTrips()
    {
        using var stream = new MemoryStream();

        await using (var workbook = new StreamingWorkbookBuilder(stream, leaveOpen: true))
        {
            workbook.AddWorksheet(
                "Revenue",
                new(
                    new SheetData(
                        new Row(
                            new Cell {
                                CellValue = new("Q1"),
                                DataType = CellValues.InlineString },
                            new Cell {
                                CellValue = new("1000"),
                                DataType = CellValues.Number }))));

            workbook.AddWorksheet(
                "Expenses",
                new(
                    new SheetData(
                        new Row(
                            new Cell
                            {
                                CellValue = new("Rent"), DataType = CellValues.InlineString
                            },
                            new Cell
                            {
                                CellValue = new("500"),
                                DataType = CellValues.Number
                            }))));
        }

        stream.Position = 0;
        using var doc = SpreadsheetDocument.Open(stream, false);
        var sheets = doc.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToList();

        Assert.Multiple(() =>
        {
            Assert.That(sheets, Has.Count.EqualTo(2));
            Assert.That(sheets[0].Name!.Value, Is.EqualTo("Revenue"));
            Assert.That(sheets[1].Name!.Value, Is.EqualTo("Expenses"));
        });

        stream.Position = 0;
        await Verify(stream, extension: "xlsx");
    }

    [Test]
    public void AddAfterDispose_Throws()
    {
        using var stream = new MemoryStream();
        var workbook = new StreamingWorkbookBuilder(stream, leaveOpen: true);
        workbook.Dispose();

        Assert.Throws<InvalidOperationException>(() =>
            workbook.AddWorksheet("Late", new(new SheetData())));
    }
}

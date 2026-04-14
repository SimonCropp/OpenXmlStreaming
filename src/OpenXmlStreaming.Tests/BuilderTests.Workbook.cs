using DocumentFormat.OpenXml.Spreadsheet;

public partial class BuilderTests
{
    [Test]
    public async Task StreamingWorkbookBuilder_RoundTrips()
    {
        using var ms = new MemoryStream();

        await using (var workbook = new StreamingWorkbookBuilder(ms, leaveOpen: true))
        {
            workbook.AddWorksheet(
                "Revenue",
                new Worksheet(
                    new SheetData(
                        new Row(
                            new Cell { CellValue = new("Q1"), DataType = CellValues.InlineString },
                            new Cell { CellValue = new("1000"), DataType = CellValues.Number }))));

            workbook.AddWorksheet(
                "Expenses",
                new Worksheet(
                    new SheetData(
                        new Row(
                            new Cell { CellValue = new("Rent"), DataType = CellValues.InlineString },
                            new Cell { CellValue = new("500"), DataType = CellValues.Number }))));
        }

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var sheets = doc.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToList();

        Assert.Multiple(() =>
        {
            Assert.That(sheets, Has.Count.EqualTo(2));
            Assert.That(sheets[0].Name!.Value, Is.EqualTo("Revenue"));
            Assert.That(sheets[1].Name!.Value, Is.EqualTo("Expenses"));
        });

        ms.Position = 0;
        await Verify(ms, extension: "xlsx");
    }

    [Test]
    public void StreamingWorkbookBuilder_AddAfterDispose_Throws()
    {
        using var ms = new MemoryStream();
        var workbook = new StreamingWorkbookBuilder(ms, leaveOpen: true);
        workbook.Dispose();

        Assert.Throws<InvalidOperationException>(() =>
            workbook.AddWorksheet("Late", new Worksheet(new SheetData())));
    }
}

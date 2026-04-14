using DocumentFormat.OpenXml.Spreadsheet;

public class SpreadsheetMigrationGuide
{
    [Test]
    public async Task Standard()
    {
        using var stream = new MemoryStream();

// begin-snippet: migration-spreadsheet-standard
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
// end-snippet

        stream.Position = 0;
        await Verify(stream, extension: "xlsx");
    }

    [Test]
    public async Task Streaming()
    {
        using var stream = new MemoryStream();

// begin-snippet: migration-spreadsheet-streaming
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
// end-snippet

        stream.Position = 0;
        await Verify(stream, extension: "xlsx");
    }

    [Test]
    public async Task Builder()
    {
        using var stream = new MemoryStream();

        // begin-snippet: migration-spreadsheet-builder
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
        // end-snippet

        stream.Position = 0;
        await Verify(stream, extension: "xlsx");
    }

    static Cell InlineString(string reference, string value) =>
        new()
        {
            CellReference = reference,
            DataType = CellValues.InlineString,
            InlineString = new(new Text(value))
        };

    static Cell Number(string reference, string value) =>
        new()
        {
            CellReference = reference,
            DataType = CellValues.Number,
            CellValue = new(value)
        };
}

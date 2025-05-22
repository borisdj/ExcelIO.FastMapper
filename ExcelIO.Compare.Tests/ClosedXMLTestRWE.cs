using ClosedXML.Attributes;
using ClosedXML.Excel;
using System.Reflection;

namespace ExcelIO.Compare.Tests
{
    public class ClosedXMLTestRWE
    {
        public static string XlsxFileName { get; set; } = "closedData.xlsx";

        public class ItemClosed
        {
            [XLColumn(Header = nameof(ItemId))]
            public int ItemId { get; set; }

            [XLColumn(Header = "Active", Order = 0)] // Order goes first with attribute XLCol. (0 is default) and those without attribute come last
            public bool IsActive { get; set; }

            [XLColumn(Header = "Full Name")]
            public string Name { get; set; }

            public int Amount { get; set; }

            [XLColumn(Order = 5)]
            public decimal Price { get; set; }

            [XLColumn(Order = 4)] // Custom Format with 3 decimal places
            public decimal? Weight { get; set; }

            [XLColumn(Header = "Created")]
            public DateTime DateCreated { get; set; }

            //public TimeOnly TimeCreated { get; set; }

            public string Note { get; set; }
        }

        [Fact]
        public void WriteTest() // Closed 18 sec (1405 KB)
        {
            var data = new List<Item>();
            for (int i = 1; i <= BaseData.NumberOfRows; i++)
            {
                var item = new Item
                {
                    ItemId = i,
                    IsActive = true,
                    Name = "Monitor" + i,
                    Amount = 2 + i,
                    Price = 1234.56m + i,
                    Weight = 2.34m + i + (i % 2 == 0 ? 0.005m : 0),
                    DateCreated = DateTime.Now,
                    Note = "info" + i,
                    //TimeCreated = new TimeOnly(1, 1)
                };
                data.Add(item);
            }

            var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet(BaseData.BaseSheetName);
            var table = worksheet.Cell(1, 1).InsertTable(data);

            var memberBindingFlags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static;
            var headerRow = table.DataRange.FirstRow();

            // THEME
            table.Theme = new XLTableTheme("None"); // "TableStyleMedium1" - "TableStyleMedium28" ("TableStyleMedium13")
            // STYLE
            //table.Style = IXLStyle
            worksheet.SheetView.FreezeRows(1);
            // FREEZE
            //worksheet.SheetView.FreezeColumns(1);
            // FILTER
            table.SetShowAutoFilter(true);

            // FONTS
            var headerFont = table.HeadersRow().Style.Font;
            //headerFont.HeaderFont = "";
            headerFont.FontSize = 11;
            var dataFont = table.Style.Font;
            //dataFont.FontName = "";
            dataFont.FontSize = 11;

            // FORMAT AND WIDTH table.
            var columnAmount = headerRow.Field(nameof(ItemClosed.Amount)).WorksheetColumn(); // if used table.Column(colNumber); then formated only till last row and not entire column
            columnAmount.Style.NumberFormat.Format = "#,##0";
            //
            var columnPrice = headerRow.Field(nameof(ItemClosed.Price)).WorksheetColumn();
            columnPrice.Style.NumberFormat.Format = "#,##0.00";
            //
            var columnDate = headerRow.Field("DateCreated").WorksheetColumn();
            columnDate.Style.NumberFormat.Format = "d/m/yyyy";

            // ORDER
            // in ClosedXML members without annotation attributes are ordered after those that do have it

            workbook.SaveAs(XlsxFileName);
        }

        [Fact]
        public void ReadTest() // Closed 18 sec (1405 KB)
        {
            var workbook = new XLWorkbook(XlsxFileName);
            var worksheet = workbook.Worksheet(BaseData.BaseSheetName);

            var rows = worksheet.RowsUsed(); // RangeUsed
            var count = rows.Count();
            int i = 0;
            var items = new List<ItemClosed>();
            foreach (var row in rows)
            {
                i++;
                if (i == 1)
                    continue;

                var item = new ItemClosed
                {
                    ItemId = (int)row.Cell(1).Value,
                    IsActive = (bool)row.Cell(2).Value,
                    Name = (string)row.Cell(3).Value,
                    DateCreated = (DateTime)row.Cell(4).Value,
                    Amount = (int)row.Cell(5).Value,
                    Price = row.Cell(6).GetValue<decimal>(),
                    Weight = row.Cell(7).GetValue<decimal>(),
                    Note = (string)row.Cell(8).Value,
                };

                items.Add(item);

                //var colCount = row.CellCount();
                /*for (int j = 1; j <= colCount; j++)
                {
                    var cell = row.Cell(j);
                    object value = cell.Value;
                    // or
                    string valueSt = cell.GetValue<string>();
                }*/
            }
        }

    }
}
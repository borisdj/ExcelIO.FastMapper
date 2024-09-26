using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2013.Excel;
using LargeXlsx;

namespace ClosedXML.MapperExtensions.Tests
{
    public class MapperTest
    {
        [Theory]
        [InlineData("LargeXlsx")]
        [InlineData("ClosedXML", false)]
        [InlineData("ClosedXML", true)]
        public void ExportTest(string type, bool hasCustomConfig = false)
        {
            var data = new List<Item>();
            for (int i = 1; i < 50_000; i++)
            {
                var item = new Item
                {
                    ItemId = i,
                    IsActive = true,
                    Name = "Monitor" + i,
                    Amount = 2 + i,
                    Price = 1234.56m + i,
                    Weight = 2.345m + i,
                    DateCreated = DateTime.Now,
                    Note = "info" + i,
                    TimeCreated = new TimeOnly(1, 1)
                };
                data.Add(item);
            }

            if(type == "LargeXlsx")
            {
                using var streamB = new FileStream("testLarge.xlsx", FileMode.Create, FileAccess.Write);
                using var xlsxWriterB = new XlsxWriter(streamB);

                var xlsxColumns = new List<XlsxColumn>();
                for (int i = 1; i <= 9; i++)
                {
                    xlsxColumns.Add(XlsxColumn.Unformatted());
                }
                xlsxColumns[3] = XlsxColumn.Formatted(style: XlsxStyle.Default.With(XlsxNumberFormat.ThousandInteger), width: 20);
                xlsxColumns[4] = XlsxColumn.Formatted(style: XlsxStyle.Default.With(XlsxNumberFormat.ThousandTwoDecimal), width: 20);

                var sh = xlsxWriterB.BeginWorksheet("Sheet 1", columns: xlsxColumns);
                foreach (var element in data)
                {
                    var row = sh.BeginRow()
                        .Write(element.ItemId)
                        .Write(element.IsActive)
                        .Write(element.Name)
                        .Write(element.Amount, XlsxStyle.Default.With(XlsxNumberFormat.ThousandInteger))
                        .Write(element.Price, XlsxStyle.Default.With(XlsxNumberFormat.ThousandTwoDecimal))
                        .Write(element.DateCreated)
                        .Write(element.Note)
                        .Write(element.TimeCreated.ToString("h:mm"));
                }
                sh.SetAutoFilter(1, 1, xlsxWriterB.CurrentRowNumber - 1, 9);
            }

            else if (type == "ClosedXML")
            {
                var xlsxExtension = ".xlsx";

                if (!hasCustomConfig)
                {
                    var workbook = XLMapper.ExportToExcel(data);
                    workbook.SaveAs("testClosed" + xlsxExtension);
                }
                else
                {
                    var mapperConfig = new XLMapperConfig
                    {
                        HeaderRowNumber = 2,
                        FreezeColumnNumber = 2,
                        AutoFilterVisible = false,
                        XLTableTheme = XLTableTheme.TableStyleMedium13 // samples: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_tableStyle_topic_ID0EFIO6.html
                    };
                    var data20 = data.Take(20).ToList();
                    var workbook2 = XLMapper.ExportToExcel(data20, mapperConfig);
                    workbook2.SaveAs("testClosed2" + xlsxExtension);
                }
            }
        }
    }

    public class Item
    {
        [XLColumnExt(Header = nameof(ItemId))]
        public int ItemId { get; set; }

        [XLColumnExt(Header = "Active", Format = XLFormatCodesFrequent.YesNo, Order = 2)] // Order goes first with attribute XLCol. (0 is default) and those without attribute come last
        public bool IsActive { get; set; }

        [XLColumnExt(Header = "Full Name", Order = 1, Width = 20)]
        public string Name { get; set; }

        public int Amount { get; set; }

        [XLColumnExt(Order = 5, HeaderFormulaType = FormulaType.SUM)]
        public decimal Price { get; set; }

        [XLColumnExt(FormatId = 14, Order = 4)] // Custom Format with 3 decimal places
        public decimal? Weight { get; set; }

        [XLColumnExt(Header = "Created", Order = 3)]
        public DateTime DateCreated { get; set; }

        public TimeOnly TimeCreated { get; set; }

        public string Note { get; set; }
    }
}
using ClosedXML.Excel;

namespace ClosedXML.MapperExtensions.Tests
{
    public class MapperTest
    {
        [Fact]
        public void ExportTest()
        {
            var data = new List<Item>()
            {
                new() { ItemId = 1, IsActive = true, Name = "Monitor", Amount = 2, Price = 1234.56m, Weight = 2.345m, DateCreated = DateTime.Now, Note = "info", TimeCreated = new TimeOnly(1,1) },
                new() { ItemId = 2, IsActive = false, Name = "Keyboard k2 ", Amount = 1000, Price = 9.2m }
            };
            var xlsxSufix = "xlsx";

            var workbook = XLMapper.ExportToExcel(data);
            workbook.SaveAs("testData." + xlsxSufix);

            var mapperConfig = new XLMapperConfig
            {
                HeaderRowNumber = 2,
                FreezeColumnNumber = 2,
                XLTableTheme = XLTableTheme.TableStyleMedium13 // samples: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_tableStyle_topic_ID0EFIO6.html
            };
            var workbook2 = XLMapper.ExportToExcel(data, mapperConfig);
            workbook2.SaveAs("testData2." + xlsxSufix); 
        }
    }

    public class Item
    {
        [XLColumnExtended(Header = nameof(ItemId))]
        public int ItemId { get; set; }

        [XLColumnExtended(Header = "Active", Format = XLFormatCodesFrequent.YesNo, Order = 2)] // Order foes first with attribute XLCol. (0 is default) and last are without attribute
        public bool IsActive { get; set; }

        [XLColumnExtended(Header = "Full Name", Order = 1, Width = 20)]
        public string Name { get; set; }

        public int Amount { get; set; }

        public decimal Price { get; set; }

        [XLColumnExtended(FormatId = 14, Order = 4)] // Custom Format with 3 decimal places
        public decimal? Weight { get; set; }

        [XLColumnExtended(Header = "Created", Order = 3)]
        public DateTime DateCreated { get; set; }

        public TimeOnly TimeCreated { get; set; }

        public string Note { get; set; }
    }
}
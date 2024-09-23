namespace ClosedXML.MapperExtensions.Tests
{
    public class MapperTest
    {
        [Fact]
        public void ExportTest()
        {
            var data = new List<Item>()
            {
                new() { ItemId = 1, IsActive = true, Name = "Monitor", Amount = 2, Price = 1234.56m, Weight = 2.345m, DateCreated = DateTime.Now, Note = "info" },
                new() { ItemId = 2, IsActive = false, Name = "Keyboard k2 ", Amount = 1000, Price = 9.2m }
            };

            var workbook = XLMapper.ExportToExcel(data);
            workbook.SaveAs("testData.xlsx");
        }
    }

    public class Item
    {
        [XLColumnExtended(Header = nameof(ItemId))]
        public int ItemId { get; set; }

        [XLColumnExtended(Header = "Full Name", Order = 2, Width = 20)]
        public string Name { get; set; }

        [XLColumnExtended(Header = "Active", Order = 1)]
        public bool IsActive { get; set; }

        public int Amount { get; set; }

        public decimal Price { get; set; }

        [XLColumnExtended(Format = "#,###.000")] // Custom Format with 3 decimal places
        public decimal? Weight { get; set; }

        [XLColumnExtended(Header = "Created", Order = 3)]
        public DateTime DateCreated { get; set; }

        public string Note { get; set; }

    }
}
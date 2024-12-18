global using Xunit;
using ClosedXML.Attributes;

namespace ExcelIO.Compare.Tests
{
    public static class BaseData
    {
        public static int NumberOfRows { get; set; } = 200_000;

        public static string BaseSheetName { get; set; } = "Sheet1";

        public static List<Item> GetData()
        {
            var data = new List<Item>();
            for (int i = 1; i < NumberOfRows; i++)
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
            return data;
        }
    }

    public class Item
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
}
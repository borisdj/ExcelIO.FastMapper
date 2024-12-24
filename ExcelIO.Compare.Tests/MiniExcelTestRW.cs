using MiniExcelLibs;

namespace ExcelIO.Compare.Tests
{
    public class MiniExcelTest
    {
        public static string XlsxFileName { get; set; } = "miniData.xlsx";

        [Fact]
        public void WriteTest() // MiniExcel 9 sec
        {
            var data = BaseData.GetData();

            var path = XlsxFileName;
            MiniExcel.SaveAs(path, data);
        }

        [Fact]
        public void ReadTest() // 12 sec
        {
            var items = new List<Item>();

            using (var reader = MiniExcel.GetReader(XlsxFileName, true))
            {
                while (reader.Read())
                {
                    var item = new Item
                    {
                        ItemId = (int)(double)reader.GetValue(0),
                        IsActive = (bool)reader.GetValue(1),
                        Name = (string)reader.GetValue(2),
                        Amount = (int)(double)reader.GetValue(3),
                        Price = (decimal)(double)reader.GetValue(4),
                        Weight = (decimal)(double)reader.GetValue(5),
                        DateCreated = (DateTime)reader.GetValue(6),
                        Note = (string)reader.GetValue(7),
                    };
                    items.Add(item);
                }
            }
            Assert.Equal(BaseData.NumberOfRows, items.Count);
        }
    }
}
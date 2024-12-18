using MiniExcelLibs;

namespace ExcelIO.Compare.Tests
{
    public class MiniExcelTest
    {
        [Fact]
        public void WriteTest() // MiniExcel 9 sec
        {
            var data = BaseData.GetData();

            var path = "testMini.xlsx";
            MiniExcel.SaveAs(path, data);
            /*
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            MiniExcel.SaveAs(path, new[] {
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2}
            });
            */
        }

        [Fact]
        public void ReadTest() // MiniExcel - 12 sec (dll 255 kb) 200 000 rows
        {
            var items = new List<Item>();

            using (var reader = MiniExcel.GetReader("testClosed.xlsx", true))
            {
                while (reader.Read())
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        var value = reader.GetValue(i);

                        var item = new Item
                        {
                            ItemId = (int)(double)reader.GetValue(0),
                            Name = (string)reader.GetValue(1),
                            IsActive = (bool)reader.GetValue(2),
                            DateCreated = (DateTime)reader.GetValue(3),
                            Price = (decimal)(double)reader.GetValue(4),
                            Weight = (decimal)(double)reader.GetValue(5),
                            Amount = (int)(double)reader.GetValue(6),
                            Note = (string)reader.GetValue(8),
                            //TimeCreated = (DateTime)dataTable.Rows[i][j],
                        };
                        items.Add(item);
                    }
                }
            }
        }
    }
}
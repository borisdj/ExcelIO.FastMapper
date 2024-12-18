using Sylvan.Data.Excel;
using System.Data;

namespace ExcelIO.Compare.Tests
{
    public class SylvanDataExcelTestRW
    {
        public static string XlsxFileName { get; set; } = "sylvanData.xlsx";

        [Fact]
        public void WriteTest() // 3 sec (5 MB)
        {
            var table = new DataTable();

            //DataColumn idColumn = table.Columns.Add(nameof(Item.ItemId), typeof(int));
            //table.PrimaryKey = new DataColumn[] { idColumn };

            table.Columns.Add(nameof(Item.ItemId), typeof(int));
            table.Columns.Add(nameof(Item.IsActive), typeof(string));
            table.Columns.Add(nameof(Item.Name), typeof(string));
            table.Columns.Add(nameof(Item.Amount), typeof(int));
            table.Columns.Add(nameof(Item.Price), typeof(decimal));
            table.Columns.Add(nameof(Item.Weight), typeof(decimal));
            table.Columns.Add(nameof(Item.DateCreated), typeof(DateTime));
            table.Columns.Add(nameof(Item.Note), typeof(string));

            var data = BaseData.GetData();

            int counter = data.Count;
            for (int i = 0; i < counter; i++)
            {
                table.Rows.Add([
                    data[i].ItemId,
                    data[i].IsActive,
                    data[i].Name,
                    data[i].Amount,
                    data[i].Price,
                    data[i].Weight,
                    data[i].DateCreated,
                    data[i].Note,
                ]);
            }

            DataTableReader reader = table.CreateDataReader();
            using (var excelDataWriter = ExcelDataWriter.Create(XlsxFileName))
            {
                var result = excelDataWriter.Write(reader, "Sheet1");
                reader.Close();
            }
        }

        [Fact]
        public void ReadTest() //  2.5 sec - Fastest
        {
            using var edr = Sylvan.Data.Excel.ExcelDataReader.Create(XlsxFileName);
            var items = new List<Item>();
            do
            {
                var sheetName = edr.WorksheetName;
                while (edr.Read())
                {
                    var item = new Item
                    {
                        ItemId = edr.GetInt32(0),
                        IsActive = edr.GetBoolean(1),
                        Name = edr.GetString(2),
                        Amount = edr.GetInt32(3),
                        Price = edr.GetDecimal(4),
                        Weight = edr.GetDecimal(5),
                        DateCreated = edr.GetDateTime(6),
                        Note = edr.GetString(7),
                    };
                    items.Add(item);
                }
            } while (edr.NextResult());
            var count = items.Count();
        }
    }
}
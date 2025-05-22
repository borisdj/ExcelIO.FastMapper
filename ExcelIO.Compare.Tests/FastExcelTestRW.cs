using FastExcel;

namespace ExcelIO.Compare.Tests
{
    public class FastExcelTestRW
    {
        public static string XlsxFileName { get; set; } = "fastData.xlsx";

        [Fact]
        public void WriteTest() // 5 sec
        {
            var data = BaseData.GetData();

            var templateFile = new FileInfo("fastTemplate.xlsx");
            var outputFile = new FileInfo("fastData.xlsx");

            using (var fastExcel = new FastExcel.FastExcel(/*templateFile, */outputFile))
            {
                fastExcel.Write(data, BaseData.BaseSheetName, 1);
            }
        }

        [Fact]
        public void ReadTest() // 5.6 sec
        {
            var inputFile = new FileInfo(XlsxFileName);

            Worksheet worksheet = null;
            using (var fastExcel = new FastExcel.FastExcel(inputFile, true))
            {
                worksheet = fastExcel.Read(BaseData.BaseSheetName); // or .Read(1)

                var items = new List<Item>();
                int i = 0;

                var rows = worksheet.Rows;
                foreach (var row in rows)
                {
                    if (i == 0)
                        continue; // skip header
                    i++;

                    //foreach (var cell in row.Cells)
                    var item = new Item
                    {
                        ItemId = (int)row.Cells.ElementAt(0).Value,
                        IsActive = (bool)row.Cells.ElementAt(1).Value,
                        Name = (string)row.Cells.ElementAt(2).Value,
                        DateCreated = (DateTime)row.Cells.ElementAt(3).Value,
                        Price = (decimal)row.Cells.ElementAt(4).Value,
                        Weight = (decimal)row.Cells.ElementAt(5).Value,
                        Amount = (int)row.Cells.ElementAt(6).Value,
                        Note = (string)row.Cells.ElementAt(8).Value,
                    };
                    items.Add(item);
                }
            }
        }
    }
}
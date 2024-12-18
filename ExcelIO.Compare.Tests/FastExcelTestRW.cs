using FastExcel;

namespace ExcelIO.Compare.Tests
{
    public class FastExcelTestRW
    {
        [Fact]
        public void WriteTest() // 5 sec
        {
            var data = BaseData.GetData();

            // Get your template and output file paths
            var templateFile = new FileInfo("Template.xlsx");
            var outputFile = new FileInfo("FastExcel.xlsx");

            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(outputFile))
            {
                fastExcel.Write(data, "sheet1", true);
            }
        }

        [Fact]
        public void ReadTest() // 5.6 sec
        {
            // Get the input file path
            var inputFile = new FileInfo("testClosed.xlsx");

            //Create a worksheet
            Worksheet worksheet = null;

            // Create an instance of Fast Excel
            using (var fastExcel = new FastExcel.FastExcel(inputFile, true))
            {
                // Read the rows using worksheet name
                //worksheet = fastExcel.Read("Sheet1");

                // Read the rows using the worksheet index
                // Worksheet indexes are start at 1 not 0
                // This method is slightly faster to find the underlying file (so slight you probably wouldn't notice)
                worksheet = fastExcel.Read(1);

                var items = new List<Item>();

                int i = 0;

                var rows = worksheet.Rows;
                foreach (var row in rows)
                {
                    if (i == 0)
                        continue;
                    i++;
                    //foreach (var cell in row.Cells)
                    //{
                    //  reader.GetValue()
                    var item = new Item
                    {
                        ItemId = (int)row.Cells.ElementAt(0).Value,
                        Name = (string)row.Cells.ElementAt(1).Value,
                        IsActive = (bool)row.Cells.ElementAt(2).Value,
                        DateCreated = (DateTime)row.Cells.ElementAt(3).Value,
                        Price = (decimal)row.Cells.ElementAt(4).Value,
                        Weight = (decimal)row.Cells.ElementAt(5).Value,
                        Amount = (int)row.Cells.ElementAt(6).Value,
                        Note = (string)row.Cells.ElementAt(8).Value,
                        //TimeCreated = (DateTime)dataTable.Rows[i][j],
                    };
                    items.Add(item);
                    //}
                }
            }
        }
    }
}
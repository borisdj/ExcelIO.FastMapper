using ExcelDataReader;
using ExcelMapper;

namespace ExcelIO.Compare.Tests
{
    public class ExcelDataReaderTest
    {
        public static string XlsxFileName { get; set; } = "readerExcel.xlsx";

        [Fact]
        public void ReadImporterTest() // Mapping ExcelMapper
        {
            //XLMapper.ImportFromExcel<ItemBase>("testClosed.xlsx");

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open("szlvan.xlsx", FileMode.Open, FileAccess.Read))
            {
                using var importer = new ExcelImporter(stream);

                ExcelSheet sheet = importer.ReadSheet();
                Item[] items = sheet.ReadRows<Item>().ToArray();

                var count = items.Count();
            }

            //var list = new List<T>();
            //return list;
        }

        [Fact]
        public void ReadFactoryTest() // 8.3 sec (dll 178 kb)
        {
            //XLMapper.ImportFromExcel<ItemBase>("testClosed.xlsx");

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open("testClosed.xlsx", FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    do
                    {
                        while (reader.Read())
                        {
                            // reader.GetDouble(0);
                        }
                    } while (reader.NextResult());

                    // 2. Use the AsDataSet extension method
                    var result = reader.AsDataSet();

                    // The result of each spreadsheet is in result.Tables
                    var dataTable = result.Tables[0];
                    var items = new List<Item>();

                    for (var i = 1; i < dataTable.Rows.Count; i++)
                    {
                        //for (var j = 0; j < dataTable.Columns.Count; j++)
                        //{
                        //var data = dataTable.Rows[i][j];

                        var item = new Item
                        {
                            ItemId = (int)(double)dataTable.Rows[i][0],
                            IsActive = (bool)dataTable.Rows[i][2],
                            Name = (string)dataTable.Rows[i][1],
                            Amount = (int)(double)dataTable.Rows[i][6],
                            Price = (decimal)(double)dataTable.Rows[i][4],
                            Weight = (decimal)(double)dataTable.Rows[i][5],
                            DateCreated = (DateTime)dataTable.Rows[i][3],
                            Note = (string)dataTable.Rows[i][8],
                            //TimeCreated = (DateTime)dataTable.Rows[i][j],
                        };
                        items.Add(item);
                        //}
                    }
                }
            }

            //var list = new List<T>();
            //return list;
        }
    }
}
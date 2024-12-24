using NanoXLSX;

namespace ExcelIO.Compare.Tests
{
    public class NanoXLSXTestRW
    {
        public static string XlsxFileName { get; set; } = "nanoData.xlsx";

        [Fact]
        public void WriteTest() // 7 sec
        {
            var data = BaseData.GetData();

            var workbook = new Workbook(XlsxFileName, BaseData.BaseSheetName);

            workbook.CurrentWorksheet.AddNextCell(nameof(Item.ItemId));
            workbook.CurrentWorksheet.AddNextCell(nameof(Item.IsActive));
            workbook.CurrentWorksheet.AddNextCell(nameof(Item.Name));
            workbook.CurrentWorksheet.AddNextCell(nameof(Item.Amount)); //, new NanoXLSX.Styles.Style { }
            workbook.CurrentWorksheet.AddNextCell(nameof(Item.Price));
            workbook.CurrentWorksheet.AddNextCell(nameof(Item.Weight));
            workbook.CurrentWorksheet.AddNextCell(nameof(Item.DateCreated));
            workbook.CurrentWorksheet.AddNextCell(nameof(Item.Note));

            for (int i = 0; i < BaseData.NumberOfRows; i++)
            {
                workbook.CurrentWorksheet.GoToNextRow();

                workbook.CurrentWorksheet.AddNextCell(data[i].ItemId);
                workbook.CurrentWorksheet.AddNextCell(data[i].IsActive);
                workbook.CurrentWorksheet.AddNextCell(data[i].Name);
                workbook.CurrentWorksheet.AddNextCell(data[i].Amount);
                workbook.CurrentWorksheet.AddNextCell(data[i].Price);
                workbook.CurrentWorksheet.AddNextCell(data[i].Weight);
                workbook.CurrentWorksheet.AddNextCell(data[i].DateCreated);
                workbook.CurrentWorksheet.AddNextCell(data[i].Note);
            }

            workbook.Save();
        }

        [Fact]
        public void ReadNano() // 28 sec
        {
            var wb = Workbook.Load(XlsxFileName);

            var items = new List<Item>();

            int count = 0;
            Item item = null;
            foreach (KeyValuePair<string, Cell> cell in wb.CurrentWorksheet.Cells)
            {
                count++;
                if (count <= 8)
                    continue;

                if (count % 8 == 1)
                {
                    item = new Item();
                    item.ItemId = (int)cell.Value.Value;
                }
                if (count % 8 == 2)
                    item.IsActive = (bool)cell.Value.Value;
                if (count % 8 == 3)
                    item.Name = (string)cell.Value.Value;
                if (count % 8 == 4)
                    item.Amount = (int)cell.Value.Value;
                if (count % 8 == 5)
                    item.Price = (decimal)(Single)cell.Value.Value;
                if (count % 8 == 6)
                    item.Weight = (decimal)(Single)cell.Value.Value;
                if (count % 8 == 7)
                    item.DateCreated = (DateTime)cell.Value.Value;
                if (count % 8 == 0)
                {
                    item.Note = (string)cell.Value.Value;
                    items.Add(item);
                }
            }

            Assert.Equal(BaseData.NumberOfRows, items.Count);
        }
    }
}
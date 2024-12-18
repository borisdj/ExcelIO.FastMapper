using NanoXLSX;

namespace ExcelIO.Compare.Tests
{
    public class NanoXLSXTestRW
    {
        [Fact]
        public void WriteTest() // 7 sec
        {
            var data = BaseData.GetData();

            Workbook workbook = new NanoXLSX.Workbook("nanoData.xlsx", "Sheet1"); // Create new workbook with a worksheet called Sheet1

            workbook.CurrentWorksheet.AddNextCell(nameof(Item.ItemId));
            workbook.CurrentWorksheet.AddNextCell(nameof(Item.IsActive));
            workbook.CurrentWorksheet.AddNextCell(nameof(Item.Name));
            workbook.CurrentWorksheet.AddNextCell(nameof(Item.Amount));
            workbook.CurrentWorksheet.AddNextCell(nameof(Item.Price));
            workbook.CurrentWorksheet.AddNextCell(nameof(Item.Weight));
            workbook.CurrentWorksheet.AddNextCell(nameof(Item.DateCreated));
            workbook.CurrentWorksheet.AddNextCell(nameof(Item.Note));

            for (int i = 1; i <= BaseData.NumberOfRows; i++)
            {
                workbook.CurrentWorksheet.GoToNextRow();

                workbook.CurrentWorksheet.AddNextCell(i);
                workbook.CurrentWorksheet.AddNextCell(true);
                workbook.CurrentWorksheet.AddNextCell("Monitor" + i);
                workbook.CurrentWorksheet.AddNextCell(2 + i);
                workbook.CurrentWorksheet.AddNextCell(1234.56m + i);
                workbook.CurrentWorksheet.AddNextCell(2.345m + i);
                workbook.CurrentWorksheet.AddNextCell(DateTime.Now);
                workbook.CurrentWorksheet.AddNextCell("info" + i);
            }

            workbook.Save();
        }

        [Fact]
        public void ReadNano() // 28 sec
        {
            NanoXLSX.Workbook wb = NanoXLSX.Workbook.Load("nanoData.xlsx");
            System.Console.WriteLine("contains worksheet name: " + wb.CurrentWorksheet.SheetName);
            var items = new List<Item>();

            int count = 0;
            var item = new Item();
            foreach (KeyValuePair<string, NanoXLSX.Cell> cell in wb.CurrentWorksheet.Cells)
            {
                count++;
                if (count <= 8)
                    continue;

                if (count % 8 == 1)
                    item.ItemId = (int)cell.Value.Value;
                if (count % 8 == 2)
                    item.IsActive = (bool)cell.Value.Value;
                if (count % 8 == 3)
                    item.Name = (string)cell.Value.Value;
                if (count % 8 == 4)
                    item.Amount = (int)cell.Value.Value;
                if (count % 8 == 5)
                    item.Price = (decimal)cell.Value.Value;
                if (count % 8 == 6)
                    item.Weight = (decimal)cell.Value.Value;
                if (count % 8 == 7)
                    item.DateCreated = (DateTime)cell.Value.Value;
                if (count % 8 == 0)
                    item.Note = (string)cell.Value.Value;

                items.Add(item);
            }
            int x = 1;
            int y = x + 1;
            x = y * 2;
        }
    }
}
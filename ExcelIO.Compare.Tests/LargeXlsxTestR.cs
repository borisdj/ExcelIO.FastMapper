using LargeXlsx;

namespace ExcelIO.Compare.Tests
{
    public class LargeXlsxTestR // 3 sec (.xlsx file size 9 MB)
    {
        public static string XlsxFileName { get; set; } = "largeData.xlsx";

        [Fact]
        public void WriteTest() // 2.5 sec (dll 60 KB)
        {
            var data = BaseData.GetData();

            using var streamB = new FileStream(XlsxFileName, FileMode.Create, FileAccess.Write);
            using var xlsxWriterB = new XlsxWriter(streamB);

            var xlsxColumns = new List<XlsxColumn>();
            for (int i = 1; i <= 9; i++)
            {
                xlsxColumns.Add(XlsxColumn.Unformatted());
            }
            xlsxColumns[3] = XlsxColumn.Formatted(style: XlsxStyle.Default.With(XlsxNumberFormat.ThousandInteger), width: 20);
            xlsxColumns[4] = XlsxColumn.Formatted(style: XlsxStyle.Default.With(XlsxNumberFormat.ThousandTwoDecimal), width: 24);
            xlsxColumns[5] = XlsxColumn.Formatted(style: XlsxStyle.Default.With(XlsxNumberFormat.ThousandTwoDecimal), width: 24);
            xlsxColumns[6] = XlsxColumn.Formatted(style: XlsxStyle.Default.With(XlsxNumberFormat.ShortDate), width: 18);

            var sh = xlsxWriterB.BeginWorksheet("Sheet 1", columns: xlsxColumns);

            var header = sh.BeginRow()
                .Write(nameof(Item.ItemId))
                .Write(nameof(Item.IsActive))
                .Write(nameof(Item.Name))
                .Write(nameof(Item.Amount))
                .Write(nameof(Item.Price))
                .Write(nameof(Item.Weight))
                .Write(nameof(Item.DateCreated))
                .Write(nameof(Item.Note));

            foreach (var element in data)
            {
                var row = sh.BeginRow()
                    .Write(element.ItemId)
                    .Write(element.IsActive)
                    .Write(element.Name)
                    .Write(element.Amount, XlsxStyle.Default.With(XlsxNumberFormat.ThousandInteger))
                    .Write(element.Price, XlsxStyle.Default.With(XlsxNumberFormat.ThousandTwoDecimal))
                    .Write(element.Weight ?? 0)
                    .Write(element.DateCreated, XlsxStyle.Default.With(XlsxNumberFormat.ShortDate))
                    .Write(element.Note)
                    ;
            }
            sh.SetAutoFilter(1, 1, xlsxWriterB.CurrentRowNumber - 1, 9);
        }
    }
}
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace ExcelIO.Compare.Tests
{
    public class ClosedXMLTestRWE
    {
        /*[Theory]
        [InlineData("ClosedXML", false)]
        [InlineData("ClosedXML", true)]
        public void ExportTestClosedXML(string type, bool hasCustomConfig = false) // Closed 18 sec (1405 KB)
        {
            var data = BaseData.GetData();

            var xlsxExtension = ".xlsx";

            if (!hasCustomConfig)
            {
                var workbook = XLMapper.ExportToExcel(data);
                workbook.SaveAs("testClosed" + xlsxExtension);
            }
            else
            {
                var mapperConfig = new XLMapperConfig
                {
                    HeaderRowNumber = 2,
                    FreezeColumnNumber = 2,
                    AutoFilterVisible = false,
                    XLTableTheme = XLTableTheme.TableStyleMedium13 // samples: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_tableStyle_topic_ID0EFIO6.html
                };
                var data20 = data.Take(20).ToList();
                var workbook2 = XLMapper.ExportToExcel(data20, mapperConfig);
                workbook2.SaveAs("testClosed2" + xlsxExtension);
            }
        }*/
    }
}
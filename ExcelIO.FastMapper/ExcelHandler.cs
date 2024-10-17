using ClosedXML;
using ClosedXML.Attributes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using LargeXlsx;
using SharpCompress.Compressors.Xz;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;

namespace ExcelIO.FastMapper
{
    public class MemberDataInfo
    {
        public string Header { get; set; }
        public int Order { get; set; }
        public MemberInfo MemberInfo { get; set; }
        public ExcelIOColumnAttribute ExcelIOColum { get; set; }
    }
    public static class ExcelHandler
    {
        public static void ExportToExcelLarge<T>(List<T> data, ExcelIOMapperConfig xlMapperConfig = null, MemoryStream memoryStreamExternal = null) where T : class
        {
            var memberBindingFlags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static;
            var type = typeof(T);
            var fieldsDict = type.GetFields(memberBindingFlags).ToDictionary(a => a.Name, a => a.FieldType);
            var propertiesDict = type.GetProperties(memberBindingFlags).ToDictionary(a => a.Name, a => a.PropertyType);
            var membersData = type.GetFields(memberBindingFlags).Cast<MemberInfo>().Concat(type.GetProperties(memberBindingFlags));
            
            var dynamicSettings = xlMapperConfig.DynamicSettings;

            var membersDict = new Dictionary<string, MemberDataInfo>();
            int i = 0;
            foreach (var member in membersData)
            {
                i++;
                var excelIOColumnAttribute = (ExcelIOColumnAttribute)member.GetCustomAttributes(typeof(ExcelIOColumnAttribute), true).FirstOrDefault();
                membersDict.Add(excelIOColumnAttribute.Header, excelIOColumnAttribute);
            }
            var members = new Dictionary<string, MemberInfo>();

            MemoryStream memoryStream = null;
            using (memoryStreamExternal == null ? memoryStream = new MemoryStream() : null)
            {
                memoryStream = memoryStream ?? memoryStreamExternal;
                using (var xlsxWriter = new XlsxWriter(memoryStream))
                {
                    var excelIOColumnAttribute = (ExcelIOColumnAttribute)property.GetCustomAttributes(typeof(ExcelIOColumnAttribute), true)[0];

                    var xlsxColumns = new List<XlsxColumn>();
                    for (int i = 1; i <= 9; i++)
                    {
                        xlsxColumns.Add(XlsxColumn.Unformatted());
                    }
                    xlsxColumns[3] = XlsxColumn.Formatted(style: XlsxStyle.Default.With(XlsxNumberFormat.ThousandInteger), width: 20);
                    xlsxColumns[4] = XlsxColumn.Formatted(style: XlsxStyle.Default.With(XlsxNumberFormat.ThousandTwoDecimal), width: 20);

                    //XlsxWriter xlsxWriterSheet = xlsxWriter.BeginWorksheet("Sheet 1", columns: xlsxColumns);

                    XlsxWriter xlsxWriterSheet = xlsxWriter.BeginWorksheet(
                        name: xlMapperConfig.SheetName,
                        splitRow: xlMapperConfig.HeaderRowNumber,
                        splitColumn: 0,
                        columns: xlsxColumns,
                        rightToLeft: false,
                        showGridLines: true,
                        showHeaders: true,
                        state: XlsxWorksheetState.Visible);

                    foreach (var element in data)
                    {
                        var row = xlsxWriterSheet.BeginRow()
                            .Write(element.ItemId)
                            .Write(element.IsActive)
                            .Write(element.Name)
                            .Write(element.Amount, XlsxStyle.Default.With(XlsxNumberFormat.ThousandInteger))
                            .Write(element.Price, XlsxStyle.Default.With(XlsxNumberFormat.ThousandTwoDecimal))
                            .Write(element.DateCreated)
                            .Write(element.Note)
                            //.Write(element.TimeCreated.ToString("h:mm"))
                            ;
                    }
                    xlsxWriterSheet.SetAutoFilter(1, 1, xlsxWriter.CurrentRowNumber - 1, 9);
                }

                if (memoryStreamExternal == null)
                {
                    File.WriteAllBytes("testLarge.xlsx", memoryStream.ToArray());
                }
            }
        }

        public static XLWorkbook ExportToExcelClosedXml<T>(List<T> data, ExcelIOMapperConfig xlMapperConfig = null) where T : class
        {
            xlMapperConfig = xlMapperConfig ?? new ExcelIOMapperConfig();

            var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet(xlMapperConfig.SheetName);
            var table = worksheet.Cell(xlMapperConfig.HeaderRowNumber, 1).InsertTable(data);

            var memberBindingFlags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static;
            var itemType = typeof(T);
            var itemsDict = itemType.GetFields(memberBindingFlags).ToDictionary(a => a.Name, a => a.FieldType);
            var propertiesDict = itemType.GetProperties(memberBindingFlags).ToDictionary(a => a.Name, a => a.PropertyType);
            var membersData = itemType.GetFields(memberBindingFlags).Cast<MemberInfo>().Concat(itemType.GetProperties(memberBindingFlags));
            var dynamicSettings = xlMapperConfig.DynamicSettings;

            var headerRow = table.DataRange.FirstRow(); //or table.Row(1);

            // THEME
            if (xlMapperConfig.XLTableTheme != null)
            {
                table.Theme = xlMapperConfig.XLTableTheme;
            }

            // STYLE
            if (xlMapperConfig.XLTableTheme != null)
            {
                table.Style = xlMapperConfig.XLStyle;
            }

            // FREEZE
            if (xlMapperConfig.FreezeHeader)
                worksheet.SheetView.FreezeRows(xlMapperConfig.HeaderRowNumber);
            if (xlMapperConfig.FreezeColumnNumber > 0)
                worksheet.SheetView.FreezeColumns(xlMapperConfig.FreezeColumnNumber);

            // FILTER
            if (!xlMapperConfig.AutoFilterVisible)
                table.SetShowAutoFilter(xlMapperConfig.AutoFilterVisible);

            // FONTS
            var dataFont = table.Style.Font;
            if (xlMapperConfig.DataFont != null)
                dataFont.FontName = xlMapperConfig.DataFont;
            if (xlMapperConfig.DataFontSize != null)
                dataFont.FontSize = (double)xlMapperConfig.DataFontSize;
            
            var headerFont = table.HeadersRow().Style.Font;
            if (xlMapperConfig.HeaderFont != null)
                headerFont.FontName = xlMapperConfig.HeaderFont;
            if (xlMapperConfig.HeaderFontSize != null)
                headerFont.FontSize = (double)xlMapperConfig.HeaderFontSize;

            // GET MAPPER INFO
            var mappersInfo = new List<XLColumnMapperInfo>();
            foreach (var member in membersData)
            {
                var attributeExt = member.GetAttributes<ExcelIOColumnAttribute>().FirstOrDefault();
                var attribute = member.GetAttributes<XLColumnAttribute>().FirstOrDefault();

                //if (dynamicSettings != null && dynamicSettings.ContainsKey(member.Name))
                //{
                //    attribute = dynamicSettings[member.Name];
                //}

                var mapper = new XLColumnMapperInfo()
                {
                    ColumnType = member.MemberType == MemberTypes.Property ? propertiesDict[member.Name]
                                                                           : itemsDict[member.Name],
                    ColumnName = attributeExt?.Header ?? attribute?.Header ?? member.Name,
                };

                if (attributeExt != null && !attributeExt.Ignore)
                {
                    mapper.Header = attributeExt.Header;
                    mapper.Order = attributeExt.Order;
                    mapper.Format = attributeExt.Format;
                    mapper.FormatId = attributeExt.FormatId;
                    mapper.Width = attributeExt.Width;
                    mapper.HeaderFormulaType = attributeExt.HeaderFormulaType;
                    mapper.HasColumnAttribute = true;
                }
                else if (attribute != null && !attribute.Ignore)
                {
                    mapper.Header = attribute.Header;
                    mapper.Order = attribute.Order;
                    mapper.Width = attributeExt.Width;
                    mapper.HasColumnAttribute = true;
                }
                else
                {
                    mapper.HasColumnAttribute = false;
                }
                mappersInfo.Add(mapper);
            }
            var columnsInfoWithoutAttribute = mappersInfo.Where(a => !a.HasColumnAttribute).ToList();
            mappersInfo = mappersInfo.Where(a => a.HasColumnAttribute).OrderBy(a => a.Order).ToList();
            mappersInfo.AddRange(columnsInfoWithoutAttribute); // in ClosedXML members without annotation attributes are ordered after those that do have it

            // SET FORMAT AND WIDTH table.
            int i = 0;
            foreach (var columnMapperInfo in mappersInfo)
            {
                columnMapperInfo.Position = ++i;
                var column = headerRow.Field(columnMapperInfo.ColumnName).WorksheetColumn(); // if used table.Column(colNumber); then formated only till last row and not entire column

                if (columnMapperInfo.FormatId >= 0)
                {
                    column.Style.NumberFormat.NumberFormatId = columnMapperInfo.FormatId;
                }

                if (xlMapperConfig.UseDefaultColumnFormat || !string.IsNullOrEmpty(columnMapperInfo.Format)) // default or custom Format
                {
                    var format = !string.IsNullOrEmpty(columnMapperInfo.Format) ? columnMapperInfo.Format
                                                                                : GetDefaultFormatForType(columnMapperInfo.ColumnType);
                    column.Style.NumberFormat.Format = format;
                }

                if(xlMapperConfig.UseDynamicColumnWidth)
                {
                    var width = columnMapperInfo.Width > 0 ? columnMapperInfo.Width
                                                           : GetWidthForHeader(columnMapperInfo, xlMapperConfig);
                    column.Width = width;
                }
                
                if (xlMapperConfig.HeaderRowNumber > 1 && columnMapperInfo.HeaderFormulaType != FormulaType.None)
                {
                    int firstDataRow = xlMapperConfig.HeaderRowNumber + 1;
                    int lastDataRow = firstDataRow + data.Count();
                    string columnLetter = column.ColumnLetter();
                    string columnName = columnMapperInfo.ColumnName;
                    string headerFormulaType = columnMapperInfo.HeaderFormulaType.ToString();

                    var formulaTable = $"={headerFormulaType}({table.Name}[{columnName}])";
                    //      example: = "=SUM(Table1[Price])"
                    var formulaRange =$"={headerFormulaType}({columnLetter}{firstDataRow}:{columnLetter}{lastDataRow})"; // not used currently
                    //      example: = "=SUM(G2:G56)"

                    var cell = column.Cell(xlMapperConfig.HeaderRowNumber - 1);
                    cell.SetFormulaA1(formulaTable);
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray; // https://github.com/closedxml/closedxml/wiki/ClosedXML-Predefined-Colors
                }
            }

            return workbook;
        }

        public static string GetDefaultFormatForType(Type columnType)
        {
            var format = string.Empty;

            if (columnType == typeof(bool) || columnType == typeof(bool?))
            {
                format = ""; // Fixed to True/False, for other(Yes/No) might have to override Values
                             // format = @"""Yes"";;""No"""; // Requires Values to be: 0 and 1
                             //   https://stackoverflow.com/questions/45835183/ms-excel-how-to-format-cell-as-yes-no-instead-of-true-false/45835438
            }
            else if (columnType == typeof(string))
            {
                format = XLFormatCodesFrequent.Text.FormatCode;                           // "@"
            }
            else if (columnType == typeof(byte) || columnType == typeof(byte?) ||
                     columnType == typeof(short)|| columnType == typeof(short?) ||
                     columnType == typeof(int)  || columnType == typeof(int?) ||
                     columnType == typeof(long) || columnType == typeof(long?))
            {
                format = XLFormatCodesFrequent.ThousandInteger.FormatCode;                // "#,##0"
            }
            else if (columnType == typeof(decimal)|| columnType == typeof(decimal?) ||
                     columnType == typeof(float)  || columnType == typeof(float?) ||
                     columnType == typeof(double) || columnType == typeof(double?))
            {
                format = XLFormatCodesFrequent.ThousandDecimals2.FormatCode;              // "#,##0.00"
            }
            else if (columnType == typeof(DateTime) || columnType == typeof(DateTime?))
            {
                format = XLFormatCodesFrequent.ShortDateTime.FormatCode;                  // "d/m/yyyy"
            }
            else if (columnType.Name == "DateOnly") // DateOnly type not in NetStandard2.0
            {
                format = XLFormatCodesFrequent.ShortDate.FormatCode;                      // "d/m/yyyy"
            }
            else if (columnType.Name == "TimeOnly") // TimeOnly type not in NetStandard2.0
            {
                format = XLFormatCodesFrequent.ShortTime.FormatCode;                      // "H:mm"
            }

            return format;
        }

        public static int GetWidthForHeader(XLColumnMapperInfo columnInfo, ExcelIOMapperConfig xlMapperConfig)
        {
            var coef = xlMapperConfig.DynamicColumnWidthCoefficient;
            int width = (int)(Math.Min(Math.Max(columnInfo.ColumnName.Length, 5), 15) * coef);
            return width;
        }
    }

    public class Item0
    {
        [ExcelIOColumn(Header = nameof(ItemId))]
        public int ItemId { get; set; }

        [ExcelIOColumn(Header = "Active", Format = XLFormatCodesFrequent.YesNo, Order = 2)] // Order goes first with attribute XLCol. (0 is default) and those without attribute come last
        public bool IsActive { get; set; }

        [ExcelIOColumn(Header = "Full Name", Order = 1, Width = 20)]
        public string Name { get; set; }

        public int Amount { get; set; }

        [ExcelIOColumn(Order = 5, HeaderFormulaType = FormulaType.SUM)]
        public decimal Price { get; set; }

        [ExcelIOColumn(FormatId = 14, Order = 4)] // Custom Format with 3 decimal places
        public decimal? Weight { get; set; }

        [ExcelIOColumn(Header = "Created", Order = 3)]
        public DateTime DateCreated { get; set; }

        //public TimeOnly TimeCreated { get; set; }

        public string Note { get; set; }
    }
}
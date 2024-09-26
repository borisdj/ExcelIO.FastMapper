using ClosedXML.Attributes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ClosedXML.MapperExtensions
{
    public static class XLMapper
    {
        public static XLWorkbook ExportToExcel<T>(List<T> data, XLMapperConfig xlMapperConfig = null) where T : class
        {
            xlMapperConfig = xlMapperConfig ?? new XLMapperConfig();

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
                var attributeExt = member.GetAttributes<XLColumnExtAttribute>().FirstOrDefault();
                var attribute = member.GetAttributes<XLColumnAttribute>().FirstOrDefault();

                if (dynamicSettings != null && dynamicSettings.ContainsKey(member.Name))
                {
                    attribute = dynamicSettings[member.Name];
                }

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

        public static int GetWidthForHeader(XLColumnMapperInfo columnInfo, XLMapperConfig xlMapperConfig)
        {
            var coef = xlMapperConfig.DynamicColumnWidthCoefficient;
            int width = (int)(Math.Min(Math.Max(columnInfo.ColumnName.Length, 5), 15) * coef);
            return width;
        }
    }
}
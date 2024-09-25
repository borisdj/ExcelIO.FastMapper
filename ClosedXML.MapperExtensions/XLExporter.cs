using ClosedXML.Attributes;
using ClosedXML.Excel;
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

            var mappersInfo = new List<XLColumnMapperInfo>();
            foreach (var member in membersData)
            {
                var attributeExtended = member.GetAttributes<XLColumnExtendedAttribute>().FirstOrDefault();
                var attribute = member.GetAttributes<XLColumnAttribute>().FirstOrDefault();

                if (dynamicSettings != null && dynamicSettings.ContainsKey(member.Name))
                {
                    attribute = dynamicSettings[member.Name];
                }

                var mapper = new XLColumnMapperInfo()
                {
                    ColumnType = member.MemberType == MemberTypes.Property ? propertiesDict[member.Name]
                                                                           : itemsDict[member.Name],
                    ColumnName = attributeExtended?.Header ?? attribute?.Header ?? member.Name,
                };

                if (attributeExtended != null && !attributeExtended.Ignore)
                {
                    mapper.Header = attributeExtended.Header;
                    mapper.Order = attributeExtended.Order;
                    mapper.Format = attributeExtended.Format;
                    mapper.FormatId = attributeExtended.FormatId;
                    mapper.Width = attributeExtended.Width;
                    mapper.HasColumnAttribute = true;
                }
                else if (attribute != null && !attribute.Ignore)
                {
                    mapper.Header = attribute.Header;
                    mapper.Order = attribute.Order;
                    mapper.Width = attributeExtended.Width;
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

            var row = table.DataRange.FirstRow();
            foreach (var columnInfo in mappersInfo)
            {
                var column = row.Field(columnInfo.ColumnName).WorksheetColumn(); // if used table.Column(colNumber); then formated only till last row and not entire column

                if (columnInfo.FormatId >= 0)
                {
                    column.Style.NumberFormat.NumberFormatId = columnInfo.FormatId;
                }

                if (xlMapperConfig.UseDefaultColumnFormat || !string.IsNullOrEmpty(columnInfo.Format)) // default or custom Format
                {
                    var format = !string.IsNullOrEmpty(columnInfo.Format) ? columnInfo.Format
                                                                          : GetDefaultFormatForType(columnInfo.ColumnType);
                    column.Style.NumberFormat.Format = format;
                }

                if(xlMapperConfig.UseDynamicColumnWidth)
                {
                    var width = columnInfo.Width > 0 ? columnInfo.Width
                                                     : GetWidthForHeader(columnInfo);
                    column.Width = width;
                }
            }

            // FREEZE
            if(xlMapperConfig.FreezeHeader)
                worksheet.SheetView.FreezeRows(xlMapperConfig.HeaderRowNumber);
            if (xlMapperConfig.FreezeColumnNumber > 0)
                worksheet.SheetView.FreezeColumns(xlMapperConfig.FreezeColumnNumber);

            // FONTS
            var dataFont = table.Style.Font;
            if (xlMapperConfig.DataFont != null)
                dataFont.FontName = xlMapperConfig.DataFont;
            if (xlMapperConfig.DataFontSize != null)
                dataFont.FontSize = (double)xlMapperConfig.DataFontSize;

            var headerFont = table.Row(1).Style.Font;
            if(xlMapperConfig.HeaderFont != null)
                headerFont.FontName = xlMapperConfig.HeaderFont;
            if (xlMapperConfig.HeaderFontSize != null)
                headerFont.FontSize = (double)xlMapperConfig.HeaderFontSize;

            // THEME
            if (xlMapperConfig.XLTableTheme != null)
            {
                table.Theme = xlMapperConfig.XLTableTheme;
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
                format = XLFormatCodesFrequent.IntegerWithThousandSeparator.FormatCode;   // "#,##0"
            }
            else if (columnType == typeof(decimal)|| columnType == typeof(decimal?) ||
                     columnType == typeof(float)  || columnType == typeof(float?) ||
                     columnType == typeof(double) || columnType == typeof(double?))
            {
                format = XLFormatCodesFrequent.Decimals2WithThousandSeparator.FormatCode; // "#,##0.00"
            }
            else if (columnType == typeof(DateTime) || columnType == typeof(DateTime?))
            {
                format = XLFormatCodesFrequent.DateTimeShort.FormatCode;                  // "d/m/yyyy"
            }
            else if (columnType.Name == "DateOnly") // DateOnly type not in NetStandard2.0
            {
                format = XLFormatCodesFrequent.DateShort.FormatCode;                      // "d/m/yyyy"
            }
            else if (columnType.Name == "TimeOnly") // TimeOnly type not in NetStandard2.0
            {
                format = XLFormatCodesFrequent.TimeShort.FormatCode;                      // "H:mm"
            }

            return format;
        }

        public static int GetWidthForHeader(XLColumnMapperInfo columnInfo)
        {
            int width = (int)(Math.Min(Math.Max(columnInfo.ColumnName.Length, 5), 15) * 1.6);
            return width;
        }
    }
}
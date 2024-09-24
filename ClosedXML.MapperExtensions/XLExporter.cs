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

            var MemberBindingFlags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static;
            var itemType = typeof(T);
            var itemsDict = itemType.GetFields(MemberBindingFlags).ToDictionary(a => a.Name, a => a.FieldType);
            var propertiesDict = itemType.GetProperties(MemberBindingFlags).ToDictionary(a => a.Name, a => a.PropertyType);
            var members = itemType.GetFields(MemberBindingFlags).Cast<MemberInfo>().Concat(itemType.GetProperties(MemberBindingFlags));

            var mappersInfo = new List<XLColumnMapperInfo>();
            foreach (var memberInfo in members)
            {
                var attributeExtended = memberInfo.GetAttributes<XLColumnExtendedAttribute>().FirstOrDefault();
                var attribute = memberInfo.GetAttributes<XLColumnAttribute>().FirstOrDefault();

                var mapperInfo = new XLColumnMapperInfo()
                {
                    ColumnType = memberInfo.MemberType == MemberTypes.Property ? propertiesDict[memberInfo.Name]
                                                                               : itemsDict[memberInfo.Name],
                    ColumnName = attributeExtended?.Header ?? attribute?.Header ?? memberInfo.Name,
                };

                if (attributeExtended != null && !attributeExtended.Ignore)
                {
                    mapperInfo.Header = attributeExtended.Header;
                    mapperInfo.Order = attributeExtended.Order;
                    mapperInfo.Format = attributeExtended.Format;
                    mapperInfo.FormatId = attributeExtended.FormatId;
                    mapperInfo.Width = attributeExtended.Width;
                    mapperInfo.HasColumnAttribute = true;
                }
                else if (attribute != null && !attribute.Ignore)
                {
                    mapperInfo.Header = attribute.Header;
                    mapperInfo.Order = attribute.Order;
                    mapperInfo.Width = attributeExtended.Width;
                    mapperInfo.HasColumnAttribute = true;
                }
                else
                {
                    mapperInfo.HasColumnAttribute = false;
                }
                mappersInfo.Add(mapperInfo);
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

                if (!string.IsNullOrEmpty(columnInfo.Format) || xlMapperConfig.UseDefaultColumnFormat)
                {
                    var format = !string.IsNullOrEmpty(columnInfo.Format) ? columnInfo.Format
                                                                          : GetDefaultFormatForType(columnInfo.ColumnType);
                    column.Style.NumberFormat.Format = format;
                }

                if(xlMapperConfig.UseDynamicColumnWidht)
                {
                    column.Width = columnInfo.Width > 0 ? columnInfo.Width
                                                        : GetWidthForHeader(columnInfo);
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
                format = XLFormatCodesFrequent.Text.FormatCode; // "@"
            }
            else if (columnType == typeof(byte) || columnType == typeof(byte?) ||
                     columnType == typeof(short)|| columnType == typeof(short?) ||
                     columnType == typeof(int)  || columnType == typeof(int?) ||
                     columnType == typeof(long) || columnType == typeof(long?))
            {
                format = XLFormatCodesFrequent.IntegerWithThousandSeparator.FormatCode; // "#,##0"
            }
            else if (columnType == typeof(decimal)|| columnType == typeof(decimal?) ||
                     columnType == typeof(float)  || columnType == typeof(float?) ||
                     columnType == typeof(double) || columnType == typeof(double?))
            {
                format = XLFormatCodesFrequent.Decimals2WithThousandSeparator.FormatCode; // "#,##0.00"
            }
            else if (columnType == typeof(DateTime) || columnType == typeof(DateTime?)) // Date, Time only
            {
                format = XLFormatCodesFrequent.DateDBFormat; // "yyyy-mm-dd"
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
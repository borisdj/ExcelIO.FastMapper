using ClosedXML.Attributes;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ClosedXML.MapperExtensions
{
    public class XLColumnMapperInfo
    {
        public Type ColumnType { get; set; }
        public string ColumnName { get; set; }

        public bool HasColumnAttribute { get; set; }

        public string Header { get; set; }
        public int Order { get; set; }
        public string Format { get; set; }
        public int Width { get; set; }

        public int Position { get; set; }
    }

    public static class XLMapper
    {
        public static XLWorkbook ExportToExcel<T>(List<T> data, string sheetName = "Data") where T : class
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet(sheetName);
            var table = worksheet.Cell(1, 1).InsertTable(data);

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

                var format = !string.IsNullOrEmpty(columnInfo.Format) ? columnInfo.Format
                                                                      : GetDefaultFormatForType(columnInfo.ColumnType);
                column.Style.NumberFormat.Format = "#,##0.00";

                column.Width = columnInfo.Width > 0 ? columnInfo.Width
                                                    : GetWidthForHeaderAndType(columnInfo); //columnInfo.ColumnName.Length * 1.5;
            }

            worksheet.SheetView.FreezeRows(1);

            var headerFont = table.Row(1).Style.Font;
            headerFont.FontName = "Arial Narrow";

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
                format = "@";
            }
            else if (columnType == typeof(byte) || columnType == typeof(byte?) ||
                     columnType == typeof(short) || columnType == typeof(short?) ||
                     columnType == typeof(int) || columnType == typeof(int?) ||
                     columnType == typeof(long) || columnType == typeof(long?))
            {
                format = "#,##0";
            }
            else if (columnType == typeof(decimal) || columnType == typeof(decimal?) ||
                     columnType == typeof(float) || columnType == typeof(float?) ||
                     columnType == typeof(double) || columnType == typeof(double?))
            {
                format = "#,##0.00";
            }
            else if (columnType == typeof(DateTime) || columnType == typeof(DateTime?)) // Date, Time only
            {
                format = "m/d/yyyy";
            }

            return format;
        }


        public static int GetWidthForHeaderAndType(XLColumnMapperInfo columnInfo)
        {
            int width = (int)(Math.Min(Math.Max(columnInfo.ColumnName.Length, 5), 15) * 1.6);
            return width;
        }
    }
}
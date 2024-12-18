using LargeXlsx;
using SharpCompress.Compressors.Xz;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Data;
using Sylvan.Data.Excel;

namespace ExcelIO.FastMapper
{
    public static class ExcelHandler
    {
        public static void ExportToExcel<T>(List<T> data, ExcelIOMapperConfig excelMapperConfig = null, MemoryStream memoryStreamExternal = null) where T : class
        {
            excelMapperConfig = excelMapperConfig ?? new ExcelIOMapperConfig();

            var memberBindingFlags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static;
            var type = typeof(T);
            //var fieldsDict = type.GetFields(memberBindingFlags).ToDictionary(a => a.Name, a => a.FieldType);
            //var propertiesDict = type.GetProperties(memberBindingFlags).ToDictionary(a => a.Name, a => a.PropertyType);
            var membersData = type.GetFields(memberBindingFlags).Cast<MemberInfo>().Concat(type.GetProperties(memberBindingFlags));

            var dynamicSettings = excelMapperConfig.DynamicSettings;

            var membersDict = new Dictionary<string, ExcelColumnMapperInfo>();
            int i = 0;
            foreach (var member in membersData)
            {
                i++;
                var attribute = member.GetCustomAttributes<ExcelIOColumnAttribute>().FirstOrDefault();

                if (dynamicSettings != null && dynamicSettings.ContainsKey(member.Name))
                {
                    attribute = dynamicSettings[member.Name];
                }

                bool propertyIsIncluded =
                    (excelMapperConfig.ExportOnlyPropertiesWithAttribute == true && attribute != null && attribute.Ignore == false) ||
                    (excelMapperConfig.ExportOnlyPropertiesWithAttribute == false && attribute != null && attribute.Ignore == false) ||
                    (excelMapperConfig.ExportOnlyPropertiesWithAttribute == false && attribute == null);

                if (propertyIsIncluded)
                {
                    var mapper = new ExcelColumnMapperInfo()
                    {
                        ColumnType = member.GetType(),
                        Header = attribute?.Header ?? member.Name,
                    };

                    if (attribute != null && attribute.Ignore == false)
                    {
                        mapper.HasColumnAttribute = true;
                        mapper.Order = attribute.Order;
                        mapper.Format = attribute.Format;
                        mapper.FormatId = attribute.FormatId;
                        mapper.Width = attribute.Width;
                        mapper.HeaderFormulaType = attribute.HeaderFormulaType;
                        mapper.HasColumnAttribute = true;
                    }

                    if (membersDict.ContainsKey(mapper.Header))
                    {
                        throw new InvalidOperationException($"2 columns can not have same Header, value: '{mapper.Header}'");
                    }

                    membersDict.Add(mapper.Header, mapper);
                }
            }

            var table = new DataTable();
            var membersOrdered = membersDict.Values.OrderBy(a => a.Order).ToList();

            foreach (var excelMember in membersOrdered)
            {
                var columnType = excelMember.ColumnType;
                var columnName = excelMember.Header;

                table.Columns.Add(columnName, columnType);
            }
            foreach (var element in data)
            {
                table.Rows.Add(membersOrdered.ToArray());
            }

            DataTableReader reader = table.CreateDataReader();
            using (var excelDataWriter = ExcelDataWriter.Create(excelMapperConfig.FileName))
            {
                var result = excelDataWriter.Write(reader, "Sheet1");
                reader.Close();
            }
        }

        public static void ExportToExcelLarge<T>(List<T> data, ExcelIOMapperConfig excelMapperConfig = null, MemoryStream memoryStreamExternal = null) where T : class
        {
            excelMapperConfig = excelMapperConfig ?? new ExcelIOMapperConfig();

            var memberBindingFlags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static;
            var type = typeof(T);
            //var fieldsDict = type.GetFields(memberBindingFlags).ToDictionary(a => a.Name, a => a.FieldType);
            //var propertiesDict = type.GetProperties(memberBindingFlags).ToDictionary(a => a.Name, a => a.PropertyType);
            var membersData = type.GetFields(memberBindingFlags).Cast<MemberInfo>().Concat(type.GetProperties(memberBindingFlags));
            
            var dynamicSettings = excelMapperConfig.DynamicSettings;

            var membersDict = new Dictionary<string, ExcelColumnMapperInfo>();
            int i = 0;
            foreach (var member in membersData)
            {
                i++;
                var attribute = member.GetCustomAttributes<ExcelIOColumnAttribute>().FirstOrDefault();

                if (dynamicSettings != null && dynamicSettings.ContainsKey(member.Name))
                {
                    attribute = dynamicSettings[member.Name];
                }

                bool propertyIsIncluded =
                    (excelMapperConfig.ExportOnlyPropertiesWithAttribute == true && attribute != null && attribute.Ignore == false) ||
                    (excelMapperConfig.ExportOnlyPropertiesWithAttribute == false && attribute != null && attribute.Ignore == false) ||
                    (excelMapperConfig.ExportOnlyPropertiesWithAttribute == false && attribute == null);

                if (propertyIsIncluded)
                {
                    //var excelMembersOrdered = membersDict.Values.Where(a => a.ExcelIOColum != null && a.ExcelIOColum.Ignore == false && a.ExcelIOColum.Order != 0).OrderBy(a => a.Order).ToList();
                    //var excelMembersUnOrdered = membersDict.Values.Where(a => a.ExcelIOColum == null || (a.ExcelIOColum.Ignore == false && a.ExcelIOColum.Order == 0)).ToList();
                    //excelMembersOrdered.AddRange(excelMembersUnOrdered);

                    var mapper = new ExcelColumnMapperInfo()
                    {
                        ColumnType = member.GetType(),
                        Header = attribute?.Header ?? member.Name,
                    };

                    if (attribute != null && attribute.Ignore == false)
                    {
                        mapper.HasColumnAttribute = true;
                        mapper.Order = attribute.Order;
                        mapper.Format = attribute.Format;
                        mapper.FormatId = attribute.FormatId;
                        mapper.Width = attribute.Width;
                        mapper.HeaderFormulaType = attribute.HeaderFormulaType;
                        mapper.HasColumnAttribute = true;
                    }

                    if (membersDict.ContainsKey(mapper.Header))
                    {
                        throw new InvalidOperationException($"2 columns can not have same Header, value: '{mapper.Header}'");
                    }

                    membersDict.Add(mapper.Header, mapper);
                }
            }

            var membersOrdered = membersDict.Values.OrderBy(a => a.Order).ToList();

            MemoryStream memoryStream = null;
            using (memoryStreamExternal == null ? memoryStream = new MemoryStream() : null)
            {
                memoryStream = memoryStream ?? memoryStreamExternal;
                using (var xlsxWriter = new XlsxWriter(memoryStream))
                {
                    var xlsxColumns = new List<XlsxColumn>();

                    foreach (var excelMember in membersOrdered)
                    {
                        var xlsxColumn = XlsxColumn.Unformatted();
                        var columnType = excelMember.ColumnType;

                        if (columnType == typeof(bool) || columnType == typeof(bool?))
                        {
                            // Fixed to True/False, for other(Yes/No) might have to override Values
                            // format = @"""Yes"";;""No"""; // Requires Values to be: 0 and 1
                            //   https://stackoverflow.com/questions/45835183/ms-excel-how-to-format-cell-as-yes-no-instead-of-true-false/45835438
                        }
                        else if (columnType == typeof(string))
                        {
                            xlsxColumn = XlsxColumn.Formatted(width: 30, style: XlsxStyle.Default.With(XlsxNumberFormat.Text));                        // "@"
                        }
                        else if (columnType == typeof(byte) || columnType == typeof(byte?) ||
                                 columnType == typeof(short)|| columnType == typeof(short?)||
                                 columnType == typeof(int)  || columnType == typeof(int?)  ||
                                 columnType == typeof(long) || columnType == typeof(long?))
                        {
                            xlsxColumn = XlsxColumn.Formatted(width: 20, style: XlsxStyle.Default.With(XlsxNumberFormat.ThousandInteger));  // "#,##0"
                        }
                        else if (columnType == typeof(decimal)|| columnType == typeof(decimal?)||
                                 columnType == typeof(float)  || columnType == typeof(float?)  ||
                                 columnType == typeof(double) || columnType == typeof(double?))
                        {
                            xlsxColumn = XlsxColumn.Formatted(width: 20, style: XlsxStyle.Default.With(XlsxNumberFormat.ThousandTwoDecimal));// "#,##0.00"
                        }
                        else if (columnType == typeof(DateTime) || columnType == typeof(DateTime?))
                        {
                            // "d/m/yyyy"
                        }
                        else if (columnType.Name == "DateOnly") // DateOnly type not in NetStandard2.0
                        {
                            // "d/m/yyyy"
                        }
                        else if (columnType.Name == "TimeOnly") // TimeOnly type not in NetStandard2.0
                        {
                            // "H:mm"
                        }

                        xlsxColumns.Add(xlsxColumn);
                    }

                    XlsxWriter xlsxWriterSheet = xlsxWriter.BeginWorksheet(
                        name: excelMapperConfig.SheetName,
                        splitRow: excelMapperConfig.HeaderRowNumber,
                        splitColumn: 0,
                        columns: xlsxColumns,
                        rightToLeft: false,
                        showGridLines: true,
                        showHeaders: true,
                        state: XlsxWorksheetState.Visible
                    );

                    var row = xlsxWriterSheet.BeginRow();
                    foreach (var excelMember in membersOrdered)
                    {
                        row.Write(excelMember.Header);
                    };

                    foreach (var element in data)
                    {
                        row = xlsxWriterSheet.BeginRow();
                        foreach (var excelMember in membersOrdered)
                        {
                            row.Write(excelMember.Header); // XlsxStyle.Default.With(XlsxNumberFormat.ThousandInteger
                        };
                    }
                    xlsxWriterSheet.SetAutoFilter(1, 1, xlsxWriter.CurrentRowNumber - 1, 9);
                }

                if (memoryStreamExternal == null)
                {
                    File.WriteAllBytes("testLarge.xlsx", memoryStream.ToArray());
                }
            }
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
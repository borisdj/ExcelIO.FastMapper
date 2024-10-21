using ClosedXML.Excel;
using System.Collections.Generic;

namespace ExcelIO.FastMapper
{

    public class ExcelIOMapperConfig
    {
        /// <summary>
        /// When using FileStream and if not custom anme defuatl value is 'ClassName.xlsx'
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// Default value: 'Data'
        /// </summary>
        public string SheetName { get; set; } = "Data";

        /// <summary>
        /// Initial value True is which case default Format set based on column type, integer with thousand separator, and decimal/float with 2 digits
        /// /</summary>
        public bool UseDefaultColumnFormat { get; set; } = true;

        /// <summary>
        /// Initial value True when Header has Filter combo visible
        /// /</summary>
        public bool AutoFilterVisible { get; set; } = true;

        /// <summary>
        /// Initial value True is which case default Width is set based on Header length, with min 5 and max 15 multiplied by DynamicColumnWidthCoefficient
        /// /</summary>
        public bool UseDynamicColumnWidth { get; set; } = true;

        /// <summary>
        /// Initial value 'False', when set to True header is Wraped
        /// /</summary>
        public bool WrapHeader { get; set; } = false;

        /// <summary>
        /// Has initial value of 1.6; Used when DynamicColumnWidth is True, Coefficient multiples ColumnName Lenght to calcule Width
        /// /</summary>
        public decimal DynamicColumnWidthCoefficient { get; set; } = 1.6m;

        /// <summary>
        /// Initial value is '1', if changed to 2 or 3 then there will be one or two empty rows above header line
        /// </summary>
        public int HeaderRowNumber { get; set; } = 1;

        /// <summary>
        /// Initial value is True in which First header row is frozen
        /// /</summary>
        public bool FreezeHeader { get; set; } = true;

        /// <summary>
        /// To freeze first N columns from left side
        /// /</summary>
        public int FreezeColumnNumber { get; set; }

        /// <summary>
        /// Default value: 'Arial Narrow'
        /// </summary>
        public string HeaderFont { get; set; } = "Arial Narrow";

        /// <summary>
        /// Table Theme type
        /// </summary>
        public XLTableTheme XLTableTheme { get; set; }

        /// <summary>
        /// Table Style
        /// </summary>
        public IXLStyle XLStyle { get; set; }

        /// <summary>
        /// Default value: 'Calibri'
        /// </summary>
        public string DataFont { get; set; }

        /// <summary>
        /// If not set custom number, default value from base library is '11'
        /// </summary>
        public double? HeaderFontSize { get; set; }

        /// <summary>
        /// If not set custom number, default value from base library is '11'
        /// </summary>
        public double? DataFontSize { get; set; }

        /// <summary>
        /// Default is False when all props are mapped to columns except those explicitly configued with 'Ignore' param.
        /// Wwhen set to True all with no explicit attribute are also ignored
        /// </summary>
        public bool ExportOnlyPropertiesWithAttribute { get; set; }

        /// <summary>
        ///     Enables Attributes to be defined at runtime, for all usage types. Dict with PropertyName and independent Attribute with parameters values.
        /// </summary>
        public Dictionary<string, ExcelIOColumnAttribute> DynamicSettings { get; set; }
    }
}
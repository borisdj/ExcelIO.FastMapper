using ClosedXML.Excel;

namespace ClosedXML.MapperExtensions
{

    public class XLMapperConfig
    {
        /// <summary>
        /// Default value: 'Data'
        /// </summary>
        public string SheetName { get; set; } = "Data";

        /// <summary>
        /// Initial value True is which case default Format set based on column type, integer with thousand separator, and decimal/float with 2 digits
        /// /</summary>
        public bool UseDefaultColumnFormat { get; set; } = true;

        /// <summary>
        /// Initial value True is which case default Widht is set based on Header length, with min 5 and max 15 multiplied by 1.6
        /// /</summary>
        public bool UseDynamicColumnWidht { get; set; } = true;

        /// <summary>
        /// Initial value is '1', if changed to 2 or 3 then there will be one or two empty rows above header line
        /// /</summary>
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
        /// Table Style type
        /// </summary>
        public XLTableTheme XLTableTheme { get; set; }

        /// <summary>
        /// Default value: 'Calibri'
        /// </summary>
        public string DataFont { get; set; }

        /// <summary>
        /// If not set custom number, default value from base libary is '11'
        /// </summary>
        public double? HeaderFontSize { get; set; }

        /// <summary>
        /// If not set custom number, default value from base libary is '11'
        /// </summary>
        public double? DataFontSize { get; set; }
    }
}
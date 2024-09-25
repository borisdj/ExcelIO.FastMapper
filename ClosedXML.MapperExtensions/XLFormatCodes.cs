namespace ClosedXML.MapperExtensions
{
    public class XLFormat
    {
        public int ID { get; set; }
        public string FormatCode { get; set; }
    }

    public static class XLFormatCodesFrequent
    {
        public static XLFormat General => XLFormatCodes.C0;                         // "General"
        public static XLFormat Text => XLFormatCodes.C49;                           // "@"

        public static XLFormat Integer => XLFormatCodes.C1;                         // "0"
        public static XLFormat IntegerWithThousandSeparator => XLFormatCodes.C3;    // "#,##0"
        public static XLFormat Decimals2 => XLFormatCodes.C2;                       // "0.00"
        public static XLFormat Decimals2WithThousandSeparator => XLFormatCodes.C4;  // "#,##0.00"
        public static XLFormat IntegerPercentage => XLFormatCodes.C9;               // "0%"
        public static XLFormat Decimals2Percentage => XLFormatCodes.C10;            // "0.00%"

        public static XLFormat TimeShort => XLFormatCodes.C20;                      // "H:mm"
        public static XLFormat TimeLong => XLFormatCodes.C21;                       // "H:mm:ss"
        public static XLFormat DateShort => XLFormatCodes.C14;                      // "d/m/yyyy"
        public static XLFormat DateTimeShort => XLFormatCodes.C22;                  // "m/d/yyyy H:mm"

        public const string DateLong = "dd/mm/yyyy";

        public const string DateMonthName = "mmmm dd, yyyy";     // 2020/1/1 => Jan 01, 2020

        public const string DateMonthNameFull = "mmmm dd, yyyy"; // 2020/1/1 => January 01, 2020

        public const string DateTimeLong = "dd/mm/yyyy H:mm:ss";


        public const string DateDBFormat = "yyyy-MM-dd";

        public const string YesNo = @"""Yes"";;""No""";
    }

    public static class XLFormatCodes
    {
        public static XLFormat C0 => new XLFormat { ID = 0, FormatCode = "General" };
        public static XLFormat C1 => new XLFormat { ID = 1, FormatCode = "0" };
        public static XLFormat C2 => new XLFormat { ID = 2, FormatCode = "0.00" };
        public static XLFormat C3 => new XLFormat { ID = 3, FormatCode = "#,##0" };
        public static XLFormat C4 => new XLFormat { ID = 4, FormatCode = "#,##0.00" };
        public static XLFormat C9 => new XLFormat { ID = 9, FormatCode = "0%" };
        public static XLFormat C10 => new XLFormat { ID = 10, FormatCode = "0.00%" };
        public static XLFormat C11 => new XLFormat { ID = 11, FormatCode = "0.00E+00" };
        public static XLFormat C12 => new XLFormat { ID = 12, FormatCode = "# ?/?" };
        public static XLFormat C13 => new XLFormat { ID = 13, FormatCode = "# ??/??" };
        public static XLFormat C14 => new XLFormat { ID = 14, FormatCode = "d/m/yyyy" };
        public static XLFormat C15 => new XLFormat { ID = 15, FormatCode = "d-mmm-yy" };
        public static XLFormat C16 => new XLFormat { ID = 16, FormatCode = "d-mmm" };
        public static XLFormat C17 => new XLFormat { ID = 17, FormatCode = "mmm-yy" };
        public static XLFormat C18 => new XLFormat { ID = 18, FormatCode = "h:mm tt" };
        public static XLFormat C19 => new XLFormat { ID = 19, FormatCode = "h:mm:ss tt" };
        public static XLFormat C20 => new XLFormat { ID = 20, FormatCode = "H:mm" };
        public static XLFormat C21 => new XLFormat { ID = 21, FormatCode = "H:mm:ss" };
        public static XLFormat C22 => new XLFormat { ID = 22, FormatCode = "m/d/yyyy H:mm" };
        public static XLFormat C37 => new XLFormat { ID = 37, FormatCode = "#,##0 ;(#,##0)" };
        public static XLFormat C38 => new XLFormat { ID = 38, FormatCode = "#,##0 ;[Red](#,##0)" };
        public static XLFormat C39 => new XLFormat { ID = 39, FormatCode = "#,##0.00;(#,##0.00)" };
        public static XLFormat C40 => new XLFormat { ID = 40, FormatCode = "#,##0.00;[Red](#,##0.00)" };
        public static XLFormat C45 => new XLFormat { ID = 45, FormatCode = "mm:ss" };
        public static XLFormat C46 => new XLFormat { ID = 46, FormatCode = "[h]:mm:ss" };
        public static XLFormat C47 => new XLFormat { ID = 47, FormatCode = "mmss.0" };
        public static XLFormat C48 => new XLFormat { ID = 48, FormatCode = "##0.0E+0" };
        public static XLFormat C49 => new XLFormat { ID = 49, FormatCode = "@" };
    }
}
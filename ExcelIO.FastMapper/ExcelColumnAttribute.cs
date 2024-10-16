using ClosedXML;
using System;
using System.Linq;
using System.Reflection;

namespace ExcelIO.FastMapper
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class ExcelColumnAttribute : Attribute
    {

        public string Header { get; set; }

        public bool Ignore { get; set; }

        public int Order { get; set; }

        /*private static XLColumnAttribute? GetXLColumnAttribute(MemberInfo mi)
        {
            if (!mi.HasAttribute<XLColumnAttribute>())
            {
                return null;
            }

            return mi.GetAttributes<XLColumnAttribute>().First();
        }

        internal static string? GetHeader(MemberInfo mi)
        {
            XLColumnAttribute xLColumnAttribute = GetXLColumnAttribute(mi);
            if (xLColumnAttribute == null)
            {
                return null;
            }

            if (!string.IsNullOrWhiteSpace(xLColumnAttribute.Header))
            {
                return xLColumnAttribute.Header;
            }

            return null;
        }

        internal static int GetOrder(MemberInfo mi)
        {
            return GetXLColumnAttribute(mi)?.Order ?? int.MaxValue;
        }

        internal static bool IgnoreMember(MemberInfo mi)
        {
            return GetXLColumnAttribute(mi)?.Ignore ?? false;
        }*/

        public string Format { get; set; }

        public int FormatId { get; set; } = -1;

        internal static string GetFormat(MemberInfo mi)
        {
            var attribute = mi.GetAttributes<ExcelColumnAttribute>()?.FirstOrDefault();
            return attribute?.Format;
        }

        internal static int? GetFormatId(MemberInfo mi)
        {
            var attribute = mi.GetAttributes<ExcelColumnAttribute>()?.FirstOrDefault();
            return attribute?.FormatId;
        }

        public int Width { get; set; }

        internal static int? GetWidth(MemberInfo mi)
        {
            var attribute = mi.GetAttributes<ExcelColumnAttribute>()?.FirstOrDefault();
            return attribute?.Width;
        }

        public FormulaType HeaderFormulaType { get; set; }

        internal static FormulaType? GetFormula(MemberInfo mi)
        {
            var attribute = mi.GetAttributes<ExcelColumnAttribute>()?.FirstOrDefault();
            return attribute?.HeaderFormulaType;
        }
    }

    public enum FormulaType
    {
        None = 0,
        SUM = 1,
        AVERAGE = 2,
        MIN = 3,
        MAX = 4,
    }
}

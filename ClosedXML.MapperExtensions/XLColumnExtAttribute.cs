using ClosedXML.Attributes;
using System;
using System.Linq;
using System.Reflection;

namespace ClosedXML.MapperExtensions
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class XLColumnExtAttribute : XLColumnAttribute // Extended Attribute
    {
        public string Format { get; set; }

        public int FormatId { get; set; } = -1;

        internal static string GetFormat(MemberInfo mi)
        {
            var attribute = mi.GetAttributes<XLColumnExtAttribute>()?.FirstOrDefault();
            return attribute?.Format;
        }

        internal static int? GetFormatId(MemberInfo mi)
        {
            var attribute = mi.GetAttributes<XLColumnExtAttribute>()?.FirstOrDefault();
            return attribute?.FormatId;
        }

        public int Width { get; set; }

        internal static int? GetWidth(MemberInfo mi)
        {
            var attribute = mi.GetAttributes<XLColumnExtAttribute>()?.FirstOrDefault();
            return attribute?.Width;
        }

        public FormulaType HeaderFormulaType { get; set; }

        internal static FormulaType? GetFormula(MemberInfo mi)
        {
            var attribute = mi.GetAttributes<XLColumnExtAttribute>()?.FirstOrDefault();
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

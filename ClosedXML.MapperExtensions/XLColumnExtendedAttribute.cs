using ClosedXML.Attributes;
using System;
using System.Linq;
using System.Reflection;

namespace ClosedXML.MapperExtensions
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class XLColumnExtendedAttribute : XLColumnAttribute
    {
        public string Format { get; set; }

        public int FormatId { get; set; } = -1;

        internal static string GetFormat(MemberInfo mi)
        {
            var attribute = mi.GetAttributes<XLColumnExtendedAttribute>()?.FirstOrDefault();
            return attribute?.Format;
        }

        internal static int? GetFormatId(MemberInfo mi)
        {
            var attribute = mi.GetAttributes<XLColumnExtendedAttribute>()?.FirstOrDefault();
            return attribute?.FormatId;
        }

        public int Width { get; set; }

        internal static int? GetWidth(MemberInfo mi)
        {
            var attribute = mi.GetAttributes<XLColumnExtendedAttribute>()?.FirstOrDefault();
            return attribute?.Width;
        }
    }
}

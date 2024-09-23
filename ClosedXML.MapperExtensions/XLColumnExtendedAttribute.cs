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

        internal static string GetFormat(MemberInfo mi)
        {
            var attribute = mi.GetAttributes<XLColumnExtendedAttribute>()?.FirstOrDefault();
            return attribute?.Format;
        }

        public int Width { get; set; }

        internal static int? GetWidth(MemberInfo mi)
        {
            var attribute = mi.GetAttributes<XLColumnExtendedAttribute>()?.FirstOrDefault();
            return attribute?.Width;
        }
    }
}

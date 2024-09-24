using System;

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
        public int FormatId { get; set; }
        public int Width { get; set; }

        public int Position { get; set; }
    }
}
using System;
using System.Globalization;
using ClosedXML.Excel;
using NPOI.SS.UserModel;

namespace RdlEngine.Render.ExcelConverter
{
    /// <summary>
    /// Writes a report value into an excel cell using the value's real CLR type
    /// </summary>
    internal static class ExcelValueConverter
    {
        private enum Kind { Text, Number, Decimal, DateTime, Boolean }

        // ClosedXML backend
        public static void SetCellValue(IXLCell cell, object typedValue, string formatted)
        {
            if (string.IsNullOrEmpty(formatted)) return;

            switch (Classify(typedValue))
            {
                case Kind.Number:
                {
                    double d = Convert.ToDouble(typedValue, CultureInfo.InvariantCulture);
                    if (double.IsNaN(d) || double.IsInfinity(d)) cell.Value = formatted;
                    else cell.Value = d;
                    break;
                }
                case Kind.Decimal:
                    cell.Value = Convert.ToDecimal(typedValue, CultureInfo.InvariantCulture);
                    break;
                case Kind.DateTime:
                    cell.Value = (DateTime)typedValue;
                    break;
                case Kind.Boolean:
                    cell.Value = (bool)typedValue;
                    break;
                default:
                    cell.Value = formatted;
                    break;
            }
        }

        // NPOI backend
        // NPOI has no decimal overload, so decimal is written through double
        public static void SetCellValue(ICell cell, object typedValue, string formatted)
        {
            if (string.IsNullOrEmpty(formatted)) return;

            switch (Classify(typedValue))
            {
                case Kind.Number:
                case Kind.Decimal:
                {
                    double d = Convert.ToDouble(typedValue, CultureInfo.InvariantCulture);
                    if (double.IsNaN(d) || double.IsInfinity(d)) cell.SetCellValue(formatted);
                    else cell.SetCellValue(d);
                    break;
                }
                case Kind.DateTime:
                    cell.SetCellValue((DateTime)typedValue);
                    break;
                case Kind.Boolean:
                    cell.SetCellValue((bool)typedValue);
                    break;
                default:
                    cell.SetCellValue(formatted);
                    break;
            }
        }

        private static Kind Classify(object value)
        {
            if (value == null) return Kind.Text;

            switch (Type.GetTypeCode(value.GetType()))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.Int16:
                case TypeCode.UInt16:
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                case TypeCode.Single:
                case TypeCode.Double:
                    return Kind.Number;
                case TypeCode.Decimal:
                    return Kind.Decimal;
                case TypeCode.DateTime:
                    return Kind.DateTime;
                case TypeCode.Boolean:
                    return Kind.Boolean;
                default:
                    return Kind.Text;
            }
        }
    }
}

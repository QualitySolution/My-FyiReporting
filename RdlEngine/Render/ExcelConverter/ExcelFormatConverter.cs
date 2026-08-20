using System;
using System.Collections.Concurrent;
using System.Globalization;
using System.Text;
using fyiReporting.RDL;

namespace RdlEngine.Render.ExcelConverter
{
    internal static class ExcelFormatConverter
    {
        private const string StandardDateFormats = "dDfFgGmMoOrRstTuUyY";
        private const string DateTokens = "dHhmMsy";

        private const string AmPm = "AM/PM";

        private const string FallbackDate = "dd.MM.yyyy";
        private const string FallbackDateTime = "dd.MM.yyyy HH:mm:ss";
        private const string AllowedSymbols = "0#?.,%/-+:()[] ";

        private static readonly ConcurrentDictionary<string, ConcurrentDictionary<string, Converted>> _converted =
            new ConcurrentDictionary<string, ConcurrentDictionary<string, Converted>>();

        private static readonly Func<string, ConcurrentDictionary<string, Converted>> _newFormatCache =
            culture => new ConcurrentDictionary<string, Converted>();

        private const int UnknownFormatCode = 999;
        private const int DefaultDecimals = 2;
        private const int MaxDecimals = 9;

        public static StyleInfo WithDefaultDateFormat(StyleInfo styleInfo, object typedValue)
        {
            if (styleInfo == null || !(typedValue is DateTime))
                return styleInfo;

            int builtinId;
            string customCode;
            if (TryConvert(styleInfo._Format, out builtinId, out customCode))
                return styleInfo;

            bool pureDate = ((DateTime)typedValue).TimeOfDay == TimeSpan.Zero;
            string standard = pureDate ? "d" : "G";

            StyleInfo copy = (StyleInfo)styleInfo.Clone();
            copy._Format = TryConvert(standard, out builtinId, out customCode)
                ? standard
                : (pureDate ? FallbackDate : FallbackDateTime);
            return copy;
        }

        /// <summary>RDL в Excel</summary>
        /// <param name="rdlFormat">Строка формата из <see cref="StyleInfo"/></param>
        /// <param name="builtinId">Код встроенного формата Excel</param>
        /// <param name="customCode">Пользовательский код формата Excel</param>
        /// <returns>false - ячейку нужно оставить в General</returns>
        public static bool TryConvert(string rdlFormat, out int builtinId, out string customCode)
        {
            var byFormat = _converted.GetOrAdd(CultureInfo.CurrentCulture.Name, _newFormatCache);

            string key = rdlFormat ?? string.Empty;
            Converted result;
            if (!byFormat.TryGetValue(key, out result))
            {
                result = Convert(rdlFormat);
                byFormat.TryAdd(key, result);
            }

            builtinId = result.BuiltinId;
            customCode = result.CustomCode;
            return result.Ok;
        }

        private static Converted Convert(string rdlFormat)
        {
            int builtinId = 0;
            string customCode = null;

            if (string.IsNullOrEmpty(rdlFormat) || rdlFormat == "General")
                return Converted.No;

            int code = StyleInfo.GetFormatCode(rdlFormat);
            if (code != UnknownFormatCode)
            {
                builtinId = code;
                return new Converted(true, builtinId, null);
            }

            customCode = ConvertNumeric(rdlFormat) ?? ConvertDate(rdlFormat);
            return customCode == null ? Converted.No : new Converted(true, 0, customCode);
        }

        private sealed class Converted
        {
            public static readonly Converted No = new Converted(false, 0, null);

            public readonly bool Ok;
            public readonly int BuiltinId;
            public readonly string CustomCode;

            public Converted(bool ok, int builtinId, string customCode)
            {
                Ok = ok;
                BuiltinId = builtinId;
                CustomCode = customCode;
            }
        }

        private static string ConvertNumeric(string format)
        {
            char kind = char.ToUpperInvariant(format[0]);
            if ("NCPFE".IndexOf(kind) < 0)
                return null;

            if (format.Length == 1 && kind == 'F') //формат даты
                return null;

            if (kind == 'E')
                return "0.00E+00";

            int decimals;
            if (!TryGetDecimals(format, out decimals))
                return null;

            string zeros = decimals > 0 ? "." + new string('0', decimals) : string.Empty;
            switch (kind)
            {
                case 'N':
                    return "#,##0" + zeros;
                case 'F':
                    return "0" + zeros;
                case 'P':
                    return "0" + zeros + "%";
                default: //валюта текущей культуры
                    return "#,##0" + zeros + " \"" + CultureInfo.CurrentCulture.NumberFormat.CurrencySymbol + "\"";
            }
        }

        private static bool TryGetDecimals(string format, out int decimals)
        {
            if (format.Length == 1)
            {
                decimals = DefaultDecimals;
                return true;
            }

            return int.TryParse(format.Substring(1), NumberStyles.None,
                                CultureInfo.InvariantCulture, out decimals)
                   && decimals <= MaxDecimals;
        }

        private static string ConvertDate(string format)
        {
            string pattern = format;
            if (format.Length == 1 && StandardDateFormats.IndexOf(format[0]) >= 0)
            {
                string[] patterns = CultureInfo.CurrentCulture.DateTimeFormat.GetAllDateTimePatterns(format[0]);
                if (patterns == null || patterns.Length == 0)
                    return null;
                pattern = patterns[0];
            }
            return ToExcelPattern(pattern);
        }

        /// <summary>Приводит пользовательский шаблон .NET к коду формата Excel</summary>
        private static string ToExcelPattern(string pattern)
        {
            StringBuilder sb = new StringBuilder(pattern.Length);
            for (int i = 0; i < pattern.Length; i++)
            {
                char c = pattern[i];
                if (c == 't' || c == 'T')
                {
                    i = SkipDesignatorRun(pattern, i);
                    sb.Append(AmPm);
                }
                else if (DateTokens.IndexOf(c) >= 0)
                    sb.Append(char.ToLowerInvariant(c));
                else if (char.IsLetter(c))
                    return null; // токен без эксельного аналога — формат не переносим
                else if (char.IsWhiteSpace(c))
                    sb.Append(' '); // в en-US перед AM/PM
                else if (AllowedSymbols.IndexOf(c) >= 0)
                    sb.Append(c);
                else
                    return null;
            }
            return sb.Length > 0 ? sb.ToString() : null;
        }

        /// <summary>Пропускает подряд идущие "t": и "t", и "tt"</summary>
        private static int SkipDesignatorRun(string pattern, int i)
        {
            while (i + 1 < pattern.Length && (pattern[i + 1] == 't' || pattern[i + 1] == 'T'))
                i++;
            return i;
        }
    }
}

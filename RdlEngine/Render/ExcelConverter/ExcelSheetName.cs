namespace RdlEngine.Render.ExcelConverter
{
    internal static class ExcelSheetName
    {
        public const int MaxLength = 31;

        /// <summary>
        /// Excel не принимает в имени листа
        /// </summary>
        private const string ForbiddenChars = "\\/?*[]:";

        public static string Sanitize(string name, string fallback)
        {
            if (string.IsNullOrEmpty(name))
                return fallback;

            char[] chars = name.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (ForbiddenChars.IndexOf(chars[i]) >= 0)
                    chars[i] = '_';
            }

            // Апостроф по краям Excel не принимает
            string clean = new string(chars).Trim('\'', ' ');
            if (clean.Length > MaxLength)
                clean = clean.Substring(0, MaxLength);

            return clean.Length == 0 ? fallback : clean;
        }
    }
}

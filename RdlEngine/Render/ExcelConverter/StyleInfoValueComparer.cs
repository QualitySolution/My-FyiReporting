using System.Collections.Generic;
using fyiReporting.RDL;

namespace RdlEngine.Render.ExcelConverter
{
    /// <summary>
    /// Valuebased equality for <see cref="StyleInfo"/>
    /// </summary>
    public class StyleInfoValueComparer : IEqualityComparer<StyleInfo>
    {
        public static readonly StyleInfoValueComparer Instance = new StyleInfoValueComparer();

        public bool Equals(StyleInfo x, StyleInfo y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (x == null || y == null) return false;

            return x.BackgroundColor.ToArgb() == y.BackgroundColor.ToArgb()
                && x.Color.ToArgb() == y.Color.ToArgb()
                && x.FontSize == y.FontSize
                && x.FontWeight == y.FontWeight
                && x.FontStyle == y.FontStyle
                && x.TextDecoration == y.TextDecoration
                && x.TextAlign == y.TextAlign
                && x.VerticalAlign == y.VerticalAlign
                && x.WritingMode == y.WritingMode
                && x.FontFamily == y.FontFamily
                && x.BStyleLeft == y.BStyleLeft
                && x.BStyleRight == y.BStyleRight
                && x.BStyleTop == y.BStyleTop
                && x.BStyleBottom == y.BStyleBottom
                && x.BWidthLeft == y.BWidthLeft
                && x.BWidthRight == y.BWidthRight
                && x.BWidthTop == y.BWidthTop
                && x.BWidthBottom == y.BWidthBottom
                && x.BColorLeft.ToArgb() == y.BColorLeft.ToArgb()
                && x.BColorRight.ToArgb() == y.BColorRight.ToArgb()
                && x.BColorTop.ToArgb() == y.BColorTop.ToArgb()
                && x.BColorBottom.ToArgb() == y.BColorBottom.ToArgb();
        }

        public int GetHashCode(StyleInfo obj)
        {
            if (obj == null) return 0;
            unchecked
            {
                int h = 17;
                h = h * 31 + obj.BackgroundColor.ToArgb();
                h = h * 31 + obj.Color.ToArgb();
                h = h * 31 + obj.FontSize.GetHashCode();
                h = h * 31 + (int)obj.FontWeight;
                h = h * 31 + (int)obj.FontStyle;
                h = h * 31 + (int)obj.TextDecoration;
                h = h * 31 + (int)obj.TextAlign;
                h = h * 31 + (int)obj.VerticalAlign;
                h = h * 31 + (int)obj.WritingMode;
                h = h * 31 + (obj.FontFamily?.GetHashCode() ?? 0);
                // Borders combined into one slot — usually identical across cells of same column
                h = h * 31 + (((int)obj.BStyleLeft << 24) ^ ((int)obj.BStyleRight << 16)
                              ^ ((int)obj.BStyleTop << 8) ^ (int)obj.BStyleBottom);
                return h;
            }
        }
    }
}

/* ====================================================================
   Copyright (C) 2004-2008  fyiReporting Software, LLC
   Copyright (C) 2011  Peter Gill <peter@majorsilence.com>

   This file is part of the fyiReporting RDL project.

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.


   For additional information, email info@fyireporting.com or visit
   the website www.fyiReporting.com.
*/

using ClosedXML.Excel;
using RdlEngine.Render.ExcelConverter;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace fyiReporting.RDL
{
    ///<summary>
    /// Renders a report to Excel as a flat dump
    ///</summary>
    internal class RenderExcel : IPresent
    {
        Report r;                       // report
        IStreamGen _sg;                 // stream generator
        Bitmap _bm = null;              // bm and
        Graphics _g = null;             //         g are needed when calculating string heights

        // Excel position trackers (0-based, like the previous ExcelValet API)
        int _ExcelRow = -1;
        int _ExcelCol = -1;

        XLWorkbook _workbook;
        IXLWorksheet _worksheet;
        string SheetName;               // current sheet name

        // Match the conversion used by RenderExcel2007ViaCloesedXML
        const double WidthPointsToSymbols = 5.637142013;
        const double RowHeightMaxPoints = 409;
        const double ColumnWidthMaxSymbols = 255;

        readonly HashSet<string> _usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        readonly Dictionary<StyleInfo, IXLStyle> _styleCache =
            new Dictionary<StyleInfo, IXLStyle>(StyleInfoValueComparer.Instance);

        public RenderExcel(Report rep, IStreamGen sg)
        {
            r = rep;
            _sg = sg;                   // We need this in future

            _workbook = new XLWorkbook();
        }

        // Exposed for the Excel2003 subclass to retrieve the produced XLSX stream
        protected IStreamGen StreamGen { get => _sg; set => _sg = value; }

        public void Dispose()
        {
            if (_workbook != null)
            {
                _workbook.Dispose();
                _workbook = null;
            }
            if (_g != null)
            {
                _g.Dispose();
                _g = null;
            }
            if (_bm != null)
            {
                _bm.Dispose();
                _bm = null;
            }
        }

        public Report Report()
        {
            return r;
        }

        public bool IsPagingNeeded()
        {
            return false;
        }

        public void Start()
        {
            return;
        }

        private Graphics GetGraphics
        {
            get
            {
                if (_g == null)
                {
                    _bm = new Bitmap(10, 10);
                    _g = Graphics.FromImage(_bm);
                }
                return _g;
            }
        }

        public virtual void End()
        {
            // Workbook must contain at least one sheet to be a valid XLSX
            if (_workbook.Worksheets.Count == 0)
                EnsureSheet("Sheet1");

            // XLSX serialization — opaque, no per-step progress available
            ExportProgress.BeginIndeterminate("Сохранение XLSX...");
            _workbook.SaveAs(_sg.GetStream());

            if (_g != null)
            {
                _g.Dispose();
                _g = null;
            }
            if (_bm != null)
            {
                _bm.Dispose();
                _bm = null;
            }
            return;
        }

        // helpers 

        private IXLWorksheet EnsureSheet(string name)
        {
            string safe = SanitizeSheetName(name);
            string candidate = safe;
            int n = 1;
            while (_usedSheetNames.Contains(candidate))
            {
                string suffix = "_" + (++n);
                int maxBase = 31 - suffix.Length;
                string baseName = safe.Length > maxBase ? safe.Substring(0, maxBase) : safe;
                candidate = baseName + suffix;
            }
            _usedSheetNames.Add(candidate);
            _worksheet = _workbook.Worksheets.Add(candidate);
            SheetName = candidate;
            return _worksheet;
        }

        private static string SanitizeSheetName(string name)
        {
            if (string.IsNullOrEmpty(name))
                return "Sheet";

            // запрещены: \ / ? * [ ] :
            char[] chars = name.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                char c = chars[i];
                if (c == '\\' || c == '/' || c == '?' || c == '*' || c == '[' || c == ']' || c == ':')
                    chars[i] = '_';
            }
            string clean = new string(chars).Trim('\'', ' ');
            if (clean.Length == 0) clean = "Sheet";
            if (clean.Length > 31) clean = clean.Substring(0, 31);
            return clean;
        }

        private void SetCell(int row, int col, string val, StyleInfo si)
        {
            if (_worksheet == null)
                EnsureSheet("Sheet1");

            var cell = _worksheet.Cell(row + 1, col + 1);
            SetValue(cell, val);

            if (si != null)
            {
                if (_styleCache.TryGetValue(si, out IXLStyle cached))
                {
                    cell.Style = cached;
                }
                else
                {
                    ExcelCellStyle.ApplyStyle(cell, si);
                    _styleCache[si] = cell.Style;
                }
            }
        }

        private void SetColumnWidth(int col, float pointsWidth)
        {
            if (_worksheet == null) return;
            double sym = pointsWidth / WidthPointsToSymbols;
            if (sym > ColumnWidthMaxSymbols) sym = ColumnWidthMaxSymbols;
            if (sym <= 0) return;
            _worksheet.Column(col + 1).Width = sym;
        }

        private void SetRowHeight(int row, float pointsHeight)
        {
            if (_worksheet == null) return;
            if (pointsHeight <= 0) return;
            _worksheet.Row(row + 1).Height = Math.Min(pointsHeight, RowHeightMaxPoints);
        }

        private void SetMerge(int firstRow, int firstCol, int lastRow, int lastCol)
        {
            if (_worksheet == null) return;
            if (firstRow == lastRow && firstCol == lastCol) return;
            _worksheet.Range(firstRow + 1, firstCol + 1, lastRow + 1, lastCol + 1).Merge();
        }

        // Mirror of RenderExcel2007ViaCloesedXML.SetValue — number/date detection
        private static void SetValue(IXLCell cell, string value)
        {
            if (string.IsNullOrEmpty(value)) return;

            if (value.Length > 1 && value[0] == '0' && double.TryParse(value, out _))
            {
                // preserve leading zero — store as text
                cell.Value = value;
            }
            else if (double.TryParse(value, out double dVal))
            {
                cell.Value = dVal;
            }
            else if (DateTime.TryParse(value, out DateTime dtVal))
            {
                bool hasLetters = false;
                for (int i = 0; i < value.Length; i++)
                {
                    if (char.IsLetter(value[i]))
                    {
                        hasLetters = true;
                        break;
                    }
                }
                if (hasLetters)
                    cell.Value = value;
                else
                    cell.Value = dtVal;
            }
            else
            {
                cell.Value = value;
            }
        }

        //IPresent

        public void BodyStart(Body b) { }
        public void BodyEnd(Body b) { }
        public void PageHeaderStart(PageHeader ph) { }
        public void PageHeaderEnd(PageHeader ph) { }
        public void PageFooterStart(PageFooter pf) { }
        public void PageFooterEnd(PageFooter pf) { }

        public void Textbox(Textbox tb, string t, Row row)
        {
            if (InTable(tb))
                SetCell(_ExcelRow, _ExcelCol, t, GetStyle(tb, row));
            else if (InList(tb))
            {
                _ExcelCol++;
                SetCell(_ExcelRow, _ExcelCol, t, GetStyle(tb, row));
            }
        }

        private StyleInfo GetStyle(ReportItem ri, Row row)
        {
            if (ri.Style == null)
                return null;

            return ri.Style.GetStyleInfo(r, row);
        }

        private static bool InTable(ReportItem tb)
        {
            Type tp = tb.Parent.Parent.GetType();
            return (tp == typeof(TableCell) ||
                    tp == typeof(Corner) ||
                    tp == typeof(DynamicColumns) ||
                    tp == typeof(DynamicRows) ||
                    tp == typeof(StaticRow) ||
                    tp == typeof(StaticColumn) ||
                    tp == typeof(Subtotal) ||
                    tp == typeof(MatrixCell));
        }

        private static bool InList(ReportItem tb)
        {
            Type tp = tb.Parent.Parent.GetType();
            return (tp == typeof(List));
        }

        public void DataRegionNoRows(DataRegion d, string noRowsMsg)            // no rows in table
        { }

        // Lists
        public bool ListStart(List l, Row row)
        {
            EnsureSheet(l.Name.Nm);
            _ExcelRow = -1;

            if (l.DataSetDefn != null)
            {
                var data = l.DataSetDefn.Query.GetMyData(r);
                if (data != null) ExportProgress.AddToTotal(data.Data.Count);
            }

            int ci = 0;
            foreach (ReportItem ri in l.ReportItems)
            {
                if (ri is Textbox tb)
                {
                    if (tb.Visibility != null && tb.Visibility.IsHidden(this.r, row))
                        continue;
                    if (ri.Width != null)
                        SetColumnWidth(ci, ri.Width.Points);
                    ci++;
                }
            }

            return true;
        }

        public void ListEnd(List l, Row r) { }

        public void ListEntryBegin(List l, Row row)
        {
            ExportProgress.Increment();
            _ExcelRow++;
            _ExcelCol = -1;

            // calc height of tallest Textbox
            float height = float.MinValue;
            foreach (ReportItem ri in l.ReportItems)
            {
                if (ri is Textbox)
                {
                    if (ri.Height != null)
                        height = Math.Max(height, ri.Height.Points);
                }
            }
            if (height != float.MinValue)
                SetRowHeight(_ExcelRow, height);
        }

        public void ListEntryEnd(List l, Row r) { }

        // Tables
        public bool TableStart(Table t, Row row)
        {
            EnsureSheet(t.Name.Nm);
            _ExcelRow = -1;

            if (t.DataSetDefn != null)
            {
                var data = t.DataSetDefn.Query.GetMyData(r);
                if (data != null) ExportProgress.AddToTotal(data.Data.Count);
            }

            int excelColumnIndex = 0;
            for (int ci = 0; ci < t.TableColumns.Items.Count; ci++)
            {
                TableColumn tc = t.TableColumns[ci];
                if (tc.Visibility != null && tc.Visibility.IsHidden(r, row))
                    continue;
                SetColumnWidth(excelColumnIndex, tc.Width.Points);
                excelColumnIndex++;
            }
            return true;
        }

        public bool IsTableSortable(Table t)
        {
            return false;   // can't have tableGroups; must have 1 detail row
        }

        public void TableEnd(Table t, Row row)
        {
            _ExcelRow++;
            return;
        }

        public void TableBodyStart(Table t, Row row) { }
        public void TableBodyEnd(Table t, Row row) { }
        public void TableFooterStart(Footer f, Row row) { }
        public void TableFooterEnd(Footer f, Row row) { }
        public void TableHeaderStart(Header h, Row row) { }
        public void TableHeaderEnd(Header h, Row row) { }

        public void TableRowStart(TableRow tr, Row row)
        {
            if (row != null) ExportProgress.Increment();  // only data rows
            _ExcelRow++;
            SetRowHeight(_ExcelRow, tr.HeightOfRow(r, this.GetGraphics, row));
            _ExcelCol = -1;
        }

        public void TableRowEnd(TableRow tr, Row row) { }

        public void TableCellStart(TableCell t, Row row)
        {
            _ExcelCol++;
            if (t.ColSpan > 1)
            {
                SetMerge(_ExcelRow, _ExcelCol, _ExcelRow, _ExcelCol + t.ColSpan - 1);
            }
            return;
        }

        public void TableCellEnd(TableCell t, Row row)
        {
            // ajm 20062008 need to increase to cover the merged cells, excel still defines every cell
            _ExcelCol += t.ColSpan - 1;
            return;
        }

        public bool MatrixStart(Matrix m, MatrixCellEntry[,] matrix, Row r, int headerRows, int maxRows, int maxCols)
        {
            EnsureSheet(m.Name.Nm);
            _ExcelRow = -1;

            ExportProgress.AddToTotal(maxRows);

            // set the widths of the columns
            float[] widths = m.ColumnWidths(matrix, maxCols);
            for (int i = 0; i < maxCols; i++)
            {
                SetColumnWidth(i, widths[i]);
            }
            return true;
        }

        public void MatrixColumns(Matrix m, MatrixColumns mc) { }

        public void MatrixCellStart(Matrix m, ReportItem ri, int row, int column, Row r, float h, float w, int colSpan)
        {
            _ExcelCol++;
        }

        public void MatrixCellEnd(Matrix m, ReportItem ri, int row, int column, Row r) { }

        public void MatrixRowStart(Matrix m, int row, Row r)
        {
            ExportProgress.Increment();
            _ExcelRow++;
            _ExcelCol = -1;
        }

        public void MatrixRowEnd(Matrix m, int row, Row r) { }

        public void MatrixEnd(Matrix m, Row r)
        {
            _ExcelRow++;
            return;
        }

        public void Chart(Chart c, Row row, ChartBase cb) { }
        public void Image(Image i, Row r, string mimeType, Stream ioin) { }

        public void Line(Line l, Row r)
        {
            return;
        }

        public bool RectangleStart(RDL.Rectangle rect, Row r)
        {
            return true;
        }

        public void RectangleEnd(RDL.Rectangle rect, Row r) { }

        // Subreport:
        public void Subreport(Subreport s, Row r) { }
        public void GroupingStart(Grouping g) { }
        public void GroupingInstanceStart(Grouping g) { }
        public void GroupingInstanceEnd(Grouping g) { }
        public void GroupingEnd(Grouping g) { }
        public void RunPages(Pages pgs) { }
    }
}

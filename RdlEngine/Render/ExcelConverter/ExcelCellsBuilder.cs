using System;
using System.Linq;
using System.Collections.Generic;
using fyiReporting.RDL;
using System.Drawing;

namespace RdlEngine.Render.ExcelConverter
{
	internal class ExcelCellsBuilder
	{
		Graphics g;
		public Report Report { get; set; }
		public static float Tolerance = 2f;

		public List<ExcelRow> Rows { get; private set; }
		public List<ExcelColumn> Columns { get; private set; }
		public List<ExcelCell> Cells { get; private set; }
		public List<ExcelImage> Images { get; private set; }
		public List<ExcelLine> Lines { get; private set; }
		public List<ExcelTable> Tables { get; private set; }

		// Subset of Cells with ExcelTable == null
		private readonly List<ExcelCell> _nonTableCells = new List<ExcelCell>();

		// Cache of row independent StyleInfo per ReportItem
		private readonly Dictionary<ReportItem, StyleInfo> _constantStyleCache =
			new Dictionary<ReportItem, StyleInfo>();

		private float rowPosition = 0f;
		private float tableLeftPosition = 0f;
		private ExcelTable CurrentExcelTable = null;

		public ExcelCellsBuilder()
		{
			g = System.Drawing.Graphics.FromImage(new Bitmap(10, 10));
			Tables = new List<ExcelTable>();
			Rows = new List<ExcelRow>();
			AddRow(0, 1);
			Columns = new List<ExcelColumn>();
			AddColumn(0, 1);
			Cells = new List<ExcelCell>();
			Images = new List<ExcelImage>();
			Lines = new List<ExcelLine>();
		}

		public ExcelImage AddImage(fyiReporting.RDL.Image image, int pictureIndex)
		{
			var img = new ExcelImage(image, pictureIndex);
			float top = image.Top.Points;
			float left = image.Left.Points;
			FillAbsolutePosition(image, ref top, ref left);

			top = GetCellAboveRelativePosition(top);

			img.AbsoluteTop = top;
			img.AbsoluteLeft = left;
			Images.Add(img);
            return img;
        }

		public void AddLine(fyiReporting.RDL.Line line, float borderWidth)
		{
			var l = new ExcelLine(line);
			float top = line.Top.Points;
			float right = line.GetX2(Report);
			float bottom = line.Y2;
			float left = line.Left.Points;
			FillAbsolutePosition(line, ref top, ref left);
			FillAbsolutePosition(line, ref bottom, ref right);

			top = GetCellAboveRelativePosition(top);
			bottom = top + line.Height.Points;

			l.AbsoluteTop = top;
			l.AbsoluteLeft = left;
			l.Right = right;
			l.Bottom = bottom;
			l.BorderWidth = borderWidth;
			Lines.Add(l);
		}

		public void AddTable(Table table)
		{
			CurrentExcelTable = new ExcelTable();
			var tableTop = table.Top.Points;
			var tableLeft = table.Left.Points;

			FillAbsolutePosition(table, ref tableTop, ref tableLeft);

			CurrentExcelTable.OriginalBottomPosition = tableTop;
			if(table.Header != null) {
				CurrentExcelTable.OriginalBottomPosition += table.Header.TableRows.Items.Sum(x => x.Height.Points);
			}
			if(table.Details != null) {
				CurrentExcelTable.OriginalBottomPosition += table.Details.TableRows.Items.Sum(x => x.Height.Points);
			}
			if(table.Footer != null) {
				CurrentExcelTable.OriginalBottomPosition += table.Footer.TableRows.Items.Sum(x => x.Height.Points);
			}

			tableTop = GetCellAboveRelativePosition(tableTop);
				
			rowPosition = tableTop;

			CurrentExcelTable.Table = table;
			Tables.Add(CurrentExcelTable);
		}

		public void AddRow(TableRow tr, Row row)
		{
			if(CurrentExcelTable.Table == null) {
				return;
			}
			float rowHeight = tr.CanGrow ? tr.HeightOfRow(Report, g, row) : tr.Height.Points;

			ShiftBottomRows(rowPosition, rowHeight);

			var currentRow = AddRow(rowPosition, rowHeight);
			foreach(var cell in tr.TableCells.Items) {
				var column = CurrentExcelTable.Table.TableColumns.Items[cell.ColIndex];
				if(column.IsHidden(Report, row))
					continue;
				var xPosition = CurrentExcelTable.Table.Left.Points;
				for(int i = 0; i < cell.ColIndex; i++) {
					var columnBefore = CurrentExcelTable.Table.TableColumns.Items[i];
					if(columnBefore.IsHidden(Report, row))
						continue;
					xPosition += columnBefore.Width.Points;
				}
				var currentColumn = AddColumn(xPosition, column.Width.Points);
				var cellTextBox = cell.ReportItems.Items.FirstOrDefault();
				if(cellTextBox == null || (cellTextBox as Textbox) == null) {
					continue;
				}
				string value = (cellTextBox as Textbox).RunText(Report, row);
				ExcelCell currentCell = new ExcelCell(cellTextBox, value, currentRow, currentColumn);
				currentCell.TypedValue = (cellTextBox as Textbox).Evaluate(Report, row);
				currentCell.ExcelTable = CurrentExcelTable;
				currentCell.OriginalWidth = column.Width.Points;
				currentCell.OriginalHeight = rowHeight;

				if(cell.ColSpan > 1) {
					float spanWidth = 0f;
					for(int i = cell.ColIndex; i < cell.ColIndex + cell.ColSpan; i++) {
						spanWidth += CurrentExcelTable.Table.TableColumns.Items[i].Width.Points;
					}
					currentCell.OriginalWidth = spanWidth;
				}
				currentCell.GrowedBottomPosition = rowPosition + rowHeight;
				SetCellStyle(currentCell, cellTextBox, row);
				AddCell(currentCell);
			}

			rowPosition += rowHeight;
			CurrentExcelTable.GrowedBottomPosition = rowPosition;
		}

		private void FillAbsolutePosition(ReportItem reportItem, ref float top, ref float left)
		{
			for(ReportLink rl = reportItem.Parent; rl != null; rl = rl.Parent) {
				if(rl is PageHeader || rl is PageFooter || rl is Body) {
					break;
				}
				if((rl is DataRegion || rl is fyiReporting.RDL.Rectangle) && !(rl is CustomReportItem)) {
					top += (rl as ReportItem).Top.Points;
					left += (rl as ReportItem).Left.Points;
				}
			}
		}

		public void AddTextbox(Textbox reportItem, string value, Row row)
		{
			if(reportItem.InPageHeaderOrFooter()) {
				return;
			}

			if(reportItem.IsTableOrMatrixCell(Report) || 
			   reportItem.Top == null || reportItem.Left == null || 
			   reportItem.Height == null || reportItem.Width == null) {
				return;
			}


			float topPosition = reportItem.Top.Points;
			float leftPosition = reportItem.Left.Points;
			float height = reportItem.CanGrow ? reportItem.RunTextCalcHeight(Report, g, row) : reportItem.Height.Points;
			float width = reportItem.Width.Points;

			FillAbsolutePosition(reportItem, ref topPosition, ref leftPosition);
			float OriginalBottomPosition = topPosition + reportItem.Height.Points;
			float bottomPosition = topPosition + height;
			float rightPosition = leftPosition + width;

			topPosition = GetCellAboveRelativePosition(topPosition);

			var currentRow = AddRow(topPosition, height);
			var currentColumn = AddColumn(leftPosition, width);

			ExcelCell currentCell = new ExcelCell(reportItem, value, currentRow, currentColumn);
			currentCell.TypedValue = reportItem.Evaluate(Report, row);
			currentCell.OriginalBottomPosition = OriginalBottomPosition;
			SetCellStyle(currentCell, reportItem, row);
			AddCell(currentCell);
			currentCell.OriginalHeight = height;
			currentCell.GrowedBottomPosition = topPosition + height;
		}

		private void SetCellStyle(ExcelCell excelCell, ReportItem reportItem, Row row)
		{
			StyleInfo si = null;
			if(reportItem.Style != null) {
				if(reportItem.Style.ConstantStyle) {
					// Row independent style
					if(!_constantStyleCache.TryGetValue(reportItem, out si)) {
						si = reportItem.Style.GetStyleInfo(Report, row);
						_constantStyleCache[reportItem] = si;
					}
				} else {
					// Style depends on the row
					si = reportItem.Style.GetStyleInfo(Report, row);
				}
			}
			if(si == null) {
				si = new StyleInfo();
			}
			excelCell.Style = si;
		}

		public float GetCellAboveRelativePosition(float top)
		{
			// Find the cell above top with the largest GrowedBottomPosition
			ExcelCell aboveCell = null;
			float bestGrowed = float.NegativeInfinity;
			for(int idx = 0; idx < Cells.Count; idx++) {
				var x = Cells[idx];
				if(x.OriginalBottomPosition < top && x.GrowedBottomPosition > bestGrowed) {
					bestGrowed = x.GrowedBottomPosition;
					aboveCell = x;
				}
			}
			if(aboveCell == null) {
				return top;
			}

			float growedPosition = 0f;
			float deltaOriginPosition = 0f;

			if(aboveCell.ExcelTable != null) {
				growedPosition = aboveCell.ExcelTable.GrowedBottomPosition;
				deltaOriginPosition = top - aboveCell.ExcelTable.OriginalBottomPosition;
			}else {
				growedPosition = aboveCell.GrowedBottomPosition;
				deltaOriginPosition = top - aboveCell.OriginalBottomPosition;
			}

			return growedPosition + (deltaOriginPosition < 0 ? 0 : deltaOriginPosition);
		}

		public ExcelRow AddRow(float top, float height)
		{
			//Insert row at Top position
			var currentRow = GetRowAtPosition(top);
			if(currentRow == null) {
				int rowIndex = InsertRow(top);
				currentRow = Rows[rowIndex];
			}
			//Insert row at Bottom position
			float bottomPosition = top + height;
			if(GetRowAtPosition(bottomPosition) == null) {
				InsertRow(bottomPosition);
			}

			return currentRow;
		}

		private ExcelColumn AddColumn(float left, float width){
			//Insert column at Left position
			var currentColumn = GetColumnAtPosition(left);
			if(currentColumn == null) {
				int columnIndex = InsertColumn(left);
				currentColumn = Columns[columnIndex];
			}
			//Insert column at Right position
			float rightPosition = left + width;
			if(GetColumnAtPosition(rightPosition) == null) {
				InsertColumn(rightPosition);
			}
			return currentColumn;
		}

		public ExcelColumn GetRightAttachColumn(ExcelCell cell)
		{
			float limit = cell.Column.XPosition + cell.ActualWidth - Tolerance;
			// Last column with XPosition < limit
			int idx = LowerBoundColumn(limit) - 1;
			if(idx >= 0) return Columns[idx];
			return null;
		}

		public int GetRightAttachCells(ExcelCell cell)
		{
			if(cell.RightAttachCol != null) {
				var ri = ColumnIndexOf(cell.RightAttachCol);
				var li = ColumnIndexOf(cell.Column);
				int result = (ri - li) - 1;
				return result < 0 ? 0 : result;
			}

			// Count columns strictly between cell.Column.XPosition and (cell.Column.XPosition + cell.ActualWidth - Tolerance)
			float leftX = cell.Column.XPosition;
			float rightX = cell.Column.XPosition + cell.ActualWidth - Tolerance;
			// first index with XPosition > leftX
			int lo = LowerBoundColumn(leftX);
			while(lo < Columns.Count && Columns[lo].XPosition <= leftX + Tolerance) lo++;
			// first index with XPosition >= rightX
			int hi = LowerBoundColumn(rightX);
			int count = hi - lo;
			return count < 0 ? 0 : count;
		}

		public int GetBottomAttachCells(ExcelCell cell)
		{
			if(cell.BottomAttachRow != null) {
				var bi = RowIndexOf(cell.BottomAttachRow);
				var ti = RowIndexOf(cell.Row);
				int result = (bi - ti) - 1;
				return result < 0 ? 0 : result;
			}

			// Count rows strictly between cell.Row.YPosition and (cell.Row.YPosition + cell.ActualHeight - Tolerance)
			float topY = cell.Row.YPosition;
			float bottomY = cell.Row.YPosition + cell.ActualHeight - Tolerance;
			// first index with YPosition > topY
			int lo = LowerBoundRow(topY);
			while(lo < Rows.Count && Rows[lo].YPosition <= topY + Tolerance) lo++;
			// first index with YPosition >= bottomY
			int hi = LowerBoundRow(bottomY);
			int count = hi - lo;
			return count < 0 ? 0 : count;
		}

		public ExcelRow GetBottomAttachRow(ExcelCell cell)
		{
			float limit = cell.Row.YPosition + cell.ActualHeight - Tolerance;
			// Last row with YPosition < limit
			int idx = LowerBoundRow(limit) - 1;
			if(idx >= 0) return Rows[idx];
			return null;
		}

		// Binary search: first index where Rows[i].YPosition >= value
		private int LowerBoundRow(float value)
		{
			int lo = 0, hi = Rows.Count;
			while(lo < hi) {
				int mid = (lo + hi) >> 1;
				if(Rows[mid].YPosition < value) lo = mid + 1;
				else hi = mid;
			}
			return lo;
		}

		// Binary search: first index where Columns[i].XPosition >= value
		private int LowerBoundColumn(float value)
		{
			int lo = 0, hi = Columns.Count;
			while(lo < hi) {
				int mid = (lo + hi) >> 1;
				if(Columns[mid].XPosition < value) lo = mid + 1;
				else hi = mid;
			}
			return lo;
		}

		// Binary search for a row by reference within the sorted position range
		private int RowIndexOf(ExcelRow row)
		{
			int idx = LowerBoundRow(row.YPosition - Tolerance);
			for(int i = idx; i < Rows.Count && Rows[i].YPosition <= row.YPosition + Tolerance; i++) {
				if(ReferenceEquals(Rows[i], row)) return i;
			}
			return -1;
		}

		// Binary search for a column by reference within the sorted position range
		private int ColumnIndexOf(ExcelColumn col)
		{
			int idx = LowerBoundColumn(col.XPosition - Tolerance);
			for(int i = idx; i < Columns.Count && Columns[i].XPosition <= col.XPosition + Tolerance; i++) {
				if(ReferenceEquals(Columns[i], col)) return i;
			}
			return -1;
		}

		private ExcelRow GetRowAtPosition(float yPositionPoints)
		{
			int idx = LowerBoundRow(yPositionPoints - Tolerance);
			for(int i = idx; i < Rows.Count && Rows[i].YPosition <= yPositionPoints + Tolerance; i++) {
				if(Math.Abs(Rows[i].YPosition - yPositionPoints) <= Tolerance)
					return Rows[i];
			}
			return null;
		}

		private ExcelColumn GetColumnAtPosition(float xPositionPoints)
		{
			int idx = LowerBoundColumn(xPositionPoints - Tolerance);
			for(int i = idx; i < Columns.Count && Columns[i].XPosition <= xPositionPoints + Tolerance; i++) {
				if(Math.Abs(Columns[i].XPosition - xPositionPoints) <= Tolerance)
					return Columns[i];
			}
			return null;
		}

		private int InsertRow(float yPositionPoints)
		{
			int idx = LowerBoundRow(yPositionPoints);
			Rows.Insert(idx, new ExcelRow(yPositionPoints));
			return idx;
		}

		private int InsertColumn(float xPositionPoints)
		{
			int idx = LowerBoundColumn(xPositionPoints);
			Columns.Insert(idx, new ExcelColumn(xPositionPoints));
			return idx;
		}
	
		private bool ResolveIntersectionConflict(ExcelCell A, ExcelCell B)
		{
			var BLeft = B.Column.XPosition;
			var BRight = B.Column.XPosition + B.OriginalWidth;
			var BTop = B.Row.YPosition;
			var BBottom = B.Row.YPosition + B.OriginalHeight;

			return ResolveIntersectionConflict(A, BLeft, BRight, BTop, BBottom);
		}

		private bool ResolveIntersectionConflict(ExcelCell A, float BLeft, float BRight, float BTop, float BBottom)
		{
			if(A == null) {
				return false;
			}
			var ALeft = A.Column.XPosition;
			var ARight = A.Column.XPosition + A.OriginalWidth;
			var ATop = A.Row.YPosition;
			var ABottom = A.Row.YPosition + A.OriginalHeight;

			bool leftEqualPosition = Math.Abs(ALeft - BLeft) <= Tolerance;
			bool topEqualPosition = Math.Abs(ATop - BTop) <= Tolerance;
			bool rightEqualPosition = Math.Abs(ARight - BRight) <= Tolerance;
			bool bottomEqualPosition = Math.Abs(ABottom - BBottom) <= Tolerance;

			bool Cross_ALeft_BTop = 	Intersection(ALeft, ATop, ALeft, ABottom, BLeft, BTop, BRight, BTop);
			bool Cross_ALeft_BBottom = 	Intersection(ALeft, ATop, ALeft, ABottom, BLeft, BBottom, BRight, BBottom);
			bool Cross_ARight_BTop = 	Intersection(ARight, ATop, ARight, ABottom, BLeft, BTop, BRight, BTop);
			bool Cross_ARight_BBottom = Intersection(ARight, ATop, ARight, ABottom, BLeft, BBottom, BRight, BBottom);

			//currently not used
			bool Cross_ATop_BLeft = 	Intersection(ALeft, ATop, ARight, ATop, BLeft, BTop, BLeft, BBottom);
			bool Cross_ATop_BRight = 	Intersection(ALeft, ATop, ARight, ATop, BRight, BTop, BRight, BBottom);
			bool Cross_ABottom_BLeft = 	Intersection(ALeft, ABottom, ARight, ABottom, BLeft, BTop, BLeft, BBottom);
			bool Cross_ABottom_BRight = Intersection(ALeft, ABottom, ARight, ABottom, BRight, BTop, BRight, BBottom);

			bool isCrossed = !(ABottom < BTop || BBottom < ATop || ARight < BLeft || BRight < ALeft);

			if(!isCrossed && !leftEqualPosition && !topEqualPosition && !rightEqualPosition && !bottomEqualPosition) {
				return true;
			}
			 
			if(!topEqualPosition && !bottomEqualPosition) {
				//Cut Upper cell to top position of a Lower cell

				//if A is upper
				if(ATop < BTop && ABottom < BBottom && isCrossed) {
					CutCellHeight(A, BTop);
					return true;
				}
				//Cut A cell to top position of a B cell
				//    ┌───┐
				//    │  A│
				// ┌──┼───┼─┐    
				// │B │   │ │
				// └──┼───┼─┘
				//    └───┘
				//            
				if(Cross_ALeft_BTop && Cross_ALeft_BBottom && BLeft < ALeft && BRight > ARight) {
					CutCellHeight(A, BTop);
					return true;
				}

				//Cut A cell to left position of a B cell
				//   ┌───┐ 
				//   │  B│ 
				//┌──┼─┐ │    
				//│A │ │ │ 
				//└──┼─┘ │
				//   └───┘ 
				//            
				if(Cross_ATop_BLeft && Cross_ABottom_BLeft && ALeft < BLeft && ARight < BRight) {
					CutCellWidthFromRight(A, BLeft);
					return true;
				}
			}

			//Cut Left cell to left position of a Right cell

			// ┌──┬───┬─┐  ┌──┬─┬─┐  
			// │A │   │ │  │A │ │ │ 
			// └──┼───┼─┘  └──┼─┘ │
			//    │  B│       │  B│
			//    └───┘       └───┘

			if(topEqualPosition && ABottom < BBottom && ALeft < BLeft && isCrossed) {
				CutCellWidthFromRight(A, BLeft);
				return true;
			}

			//Cut Left cell to left position of a Right cell
			if(ALeft < BLeft && ARight < BRight && ARight > BLeft &&
			   	(
					//┌───┬─┬───┐
					//│A  │ │  B│
					//└───┴─┴───┘
				  	(topEqualPosition && bottomEqualPosition)
					//┌────┐
					//│  ┌─┼──┐  
					//│A │ │ B│
					//│  └─┼──┘   
					//└────┘
				   	|| (!topEqualPosition && !bottomEqualPosition && Cross_ARight_BTop && Cross_ARight_BBottom)
					//┌──┬─┬──┐
					//│A │ │ B│
					//│  └─┼──┘
					//└────┘
				   	|| (topEqualPosition && !bottomEqualPosition && Cross_ARight_BBottom)
					//┌────┐
					//│  ┌─┼──┐
					//│A │ │ B│
					//└──┴─┴──┘
				   	|| (!topEqualPosition && bottomEqualPosition && Cross_ARight_BTop)
				)){
				CutCellWidthFromRight(A, BLeft);
				return true;
			}

			//Cut A cell to right position of a Left cell
			//┌──┬──┐
			//│  │ A│
			//├──┼──┘
			//│B │
			//└──┘
			if(topEqualPosition && leftEqualPosition && !bottomEqualPosition && !rightEqualPosition
			   && Cross_ABottom_BRight){
				CutCellWidthFromLeft(A, BRight);
				return true;
			}

			//B region full contain A
			if((BTop < ATop || topEqualPosition) 
			   && (BLeft < ALeft || leftEqualPosition) 
			   && (BRight > ARight || rightEqualPosition)
			   && (BBottom > ABottom || bottomEqualPosition)) {
				//delete A
				RemoveCell(A);
				if(!A.Column.Cells.Any()) {
					Columns.Remove(A.Column);
				}
				return true;
			}

			return false;
		}

		private void CutCellWidthFromLeft(ExcelCell cell, float toPos)
		{
			var LeftPosition = cell.Column.XPosition;
			var RightPosition = LeftPosition + cell.OriginalWidth;
			if(!(toPos > LeftPosition && toPos < RightPosition)) {
				return;
			}
			ExcelColumn StartColumn = GetColumnAtPosition(LeftPosition);
			ExcelColumn newStartColumn = GetColumnAtPosition(toPos);
			if(newStartColumn == null) {
				var index = InsertColumn(toPos);
				newStartColumn = Columns[index];
				Columns.Remove(newStartColumn);
			}
			var newCell = new ExcelCell(cell.ReportItem, cell.Value, cell.Row, newStartColumn);
			newCell.TypedValue = cell.TypedValue;
			ExcelCell existCell = newStartColumn.Cells.FirstOrDefault(x => x.Row == cell.Row);
			if(existCell != null) {
				ResolveIntersectionConflict(cell, existCell);
			}
			newCell.Style = cell.Style;
			AddCell(newCell);
			newCell.Row.Cells.Add(newCell);
			newStartColumn.Cells.Add(newCell);
			RemoveCell(cell);

			LeftPosition = newStartColumn.XPosition;
			RightPosition = LeftPosition + cell.OriginalWidth;
			cell.RightAttachCol = GetColumnAtPosition(RightPosition);
			cell.CorrectedWidth = RightPosition - LeftPosition;
		}

		private void CutCellWidthFromRight(ExcelCell cell, float toPos)
		{
			var LeftPosition = cell.Column.XPosition;
			var RightPosition = LeftPosition + cell.OriginalWidth;
			if(!(toPos > LeftPosition && toPos < RightPosition)) {
				return;
			}
			cell.RightAttachCol = GetColumnAtPosition(toPos);
			ExcelColumn endColumn = GetColumnAtPosition(RightPosition);
			if(endColumn != null && endColumn.IsEmpty) {
				Columns.Remove(endColumn);
			}
			cell.CorrectedWidth = toPos - LeftPosition;
		}

		private void CutCellHeight(ExcelCell cell, float toPos)
		{
			var TopPosition = cell.Row.YPosition;
			var BottomPosition = TopPosition + cell.OriginalHeight;
			if(!(toPos > TopPosition && toPos < BottomPosition)) {
				return;
			}
			cell.BottomAttachRow = GetRowAtPosition(toPos);
			ExcelRow endRow = GetRowAtPosition(BottomPosition);
			if(endRow != null && endRow.IsEmpty) {
				Rows.Remove(endRow);
			}
			cell.CorrectedHeight = toPos - TopPosition;
		}

		private void AddCell(ExcelCell cell)
		{
			Cells.Add(cell);
			if (cell.ExcelTable == null)
				_nonTableCells.Add(cell);
		}

		private void RemoveCell(ExcelCell cell)
		{
			Cells.Remove(cell);
			if (cell.ExcelTable == null)
				_nonTableCells.Remove(cell);
			cell.Column.Cells.Remove(cell);
			cell.Row.Cells.Remove(cell);
		}

		private bool Intersection(double ax1, double ay1, double ax2, double ay2, double bx1, double by1, double bx2, double by2)
		{
			double v1, v2, v3, v4;
			v1 = ((bx2 - bx1) * (ay1 - by1)) - ((by2 - by1) * (ax1 - bx1));
			v2 = ((bx2 - bx1) * (ay2 - by1)) - ((by2 - by1) * (ax2 - bx1));
			v3 = ((ax2 - ax1) * (by1 - ay1)) - ((ay2 - ay1) * (bx1 - ax1));
			v4 = ((ax2 - ax1) * (by2 - ay1)) - ((ay2 - ay1) * (bx2 - ax1));
			bool res = (v1 * v2 < 0) && (v3 * v4 < 0);
			return res;
		}

		public void CellsCorrection()
		{
			// We only resolve conflicts between nontable cells, table cells are skipped
			for (int i = 0; i < _nonTableCells.Count; i++)
			{
				var count = _nonTableCells.Count;
				var cell = _nonTableCells[i];

				var unresolvedCells = new Stack<ExcelCell>();
				float cellTop = cell.Row.YPosition;
				float cellBottom = cell.OriginalHeight + cellTop;
				float cellLeft = cell.Column.XPosition;
				float cellRight = cell.OriginalWidth + cellLeft;

				for (int k = 0; k < _nonTableCells.Count; k++)
				{
					var x = _nonTableCells[k];
					float xTop = x.Row.YPosition;
					float xBottom = x.OriginalHeight + xTop;
					float xLeft = x.Column.XPosition;
					float xRight = x.OriginalWidth + xLeft;

					if (!(xBottom < cellTop || cellBottom < xTop || xRight < cellLeft || cellRight < xLeft))
					{
						unresolvedCells.Push(x);
					}
				}

				while (unresolvedCells.Count > 0)
				{
					var uc = unresolvedCells.Pop();
					if (uc != cell)
					{
						ResolveIntersectionConflict(cell, uc);
					}
				}

				var dCount = i - (count - _nonTableCells.Count);
				i = dCount < 0 ? 0 : dCount;
			}
		}

		private void ShiftBottomRows(float fromPosition, float yOffset)
		{
			int startIdx = LowerBoundRow(fromPosition);
			for(int i = startIdx; i < Rows.Count; i++) {
				Rows[i].yOffset += yOffset;
				Rows[i].YPosition += yOffset;
			}
		}

	}
}

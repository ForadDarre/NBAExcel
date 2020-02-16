using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows;

namespace NBAExcel
{
	public static class Utils
	{
		public static bool CheckSheetsExistance(Excel.Workbook workbook, string sheetName)
		{
			bool state = true;

			Excel.Worksheet oddsSeason = workbook.Sheets
				.Cast<Excel.Worksheet>()
				.Where(s => s.Name == sheetName)
				.FirstOrDefault();

			if (oddsSeason == null)
			{
				MessageBox.Show("Sheet '" + sheetName + "' not found.", "Error!");
				state = false;
			}

			return state;
		}

		public static int LastRow(Excel.Worksheet wks)
		{
			Excel.Range lastCell = wks.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
			return lastCell.Row;
		}

		public static int LastColumn(Excel.Worksheet wks)
		{
			Excel.Range lastCell = wks.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
			return lastCell.Column;
		}

		public static string GetExcelColumnName(int columnNumber)
		{
			int dividend = columnNumber;
			string columnName = String.Empty;
			int modulo;

			while (dividend > 0)
			{
				modulo = (dividend - 1) % 26;
				columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
				dividend = (int)((dividend - modulo) / 26);
			}

			return columnName;
		}
	}
}

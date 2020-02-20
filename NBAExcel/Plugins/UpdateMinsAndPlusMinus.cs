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
	public class UpdateMinsAndPlusMinus
	{
		readonly static string sheetInput = "Input";
		readonly static string sheetMins = "Mins";

		public static void Main()
		{
			Excel.Application application = Globals.ThisAddIn.Application;
			Excel.Workbook workbook = application.ActiveWorkbook;

			if (Utils.CheckSheetsExistance(workbook, sheetInput) 
				&& Utils.CheckSheetsExistance(workbook, sheetMins))
			{
				string date = GetDate(workbook);
				ChangeFormula(workbook, date);
			}
		}

		private static string GetDate(Excel.Workbook workbook)
		{
			Excel.Worksheet input = workbook.Sheets
				.Cast<Excel.Worksheet>()
				.Where(s => s.Name == sheetInput)
				.FirstOrDefault();

			return (input.Cells[1082, 2] as Excel.Range).Text; // "Input" sheet, 1082 row, 2 column - where date is located
		}

		private static void ChangeFormula(Excel.Workbook workbook, string date)
		{
			Excel.Worksheet mins = workbook.Sheets
				.Cast<Excel.Worksheet>()
				.Where(s => s.Name == sheetMins)
				.FirstOrDefault();

			int column = GetColumn(mins, date);

			if (column > 15)
			{
				for(int n = 1; n <= 30; n++)	// 30 teams, 36 rows for players in each, 3 rows between teams
				{
					for (int i = 3 + 36 * (n - 1); i < 36 * n; i++)
					{
						(mins.Cells[i, 4] as Excel.Range).Value = GetFormula(i, column, 7);
						(mins.Cells[i, 5] as Excel.Range).Value = GetFormula(i, column, 14);
					}
				}
				
			}
		}

		private static int GetColumn(Excel.Worksheet mins, string date)
		{
			for(int i = 1; i < Utils.LastColumn(mins); i++)
			{
				if((mins.Cells[2, i] as Excel.Range).Text == date)
				{
					return i;
				}
			}

			return 0;
		}

		private static string GetFormula(int row, int column, int numberOfDays)
		{
			string rangeStart = Utils.GetExcelColumnName(column - numberOfDays);
			string rangeEnd = Utils.GetExcelColumnName(column - 1);

			return "=IFERROR(" +
				"(SUM(" + rangeStart + row.ToString() + ":" + rangeEnd + row.ToString() + "))"
				+ "/(COUNT(" + rangeStart + row.ToString() + ":" + rangeEnd + row.ToString() + "))"
				+ ", )";
		}
	}
}

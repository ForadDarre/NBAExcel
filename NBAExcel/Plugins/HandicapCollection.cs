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
	public class HandicapCollection
	{
		readonly static string sheetConvStatsName = "ConvTEST";
		readonly static string sheetOddsSeasonName = "Odds season";

		const double minAH = -30.0;
		const double maxAH = 30.0;

		public static void Main()
		{
			Excel.Application application = Globals.ThisAddIn.Application;
			Excel.Workbook workbook = application.ActiveWorkbook;

			if (Utils.CheckSheetsExistance(workbook, sheetConvStatsName) && Utils.CheckSheetsExistance(workbook, sheetOddsSeasonName))
			{
				int j = 0;
				for(double i = minAH; i <= minAH; i += 0.5, j++)
				{
					List<List<double>> allCoefs = GetAllCoefsForOneHandicap(workbook, i);
					CopyCoefsToConvStats(workbook, allCoefs, i, j * 5 + 7);
				}
				

			}

		}


		private static List<List<double>> GetAllCoefsForOneHandicap(Excel.Workbook workbook, double handicap)
		{
			Excel.Worksheet oddsSeason = workbook.Sheets
				.Cast<Excel.Worksheet>()
				.Where(s => s.Name == sheetOddsSeasonName)
				.FirstOrDefault();

			List<List<double>> allCoefs = new List<List<double>>();
			for (int i = 5; i <= Utils.LastRow(oddsSeason); i++)
			{
				List<double> coefs = GetCoefs(oddsSeason, i, handicap);
				if (coefs != null)
				{
					allCoefs.Add(coefs);
				}
			}

			return allCoefs;
		}

		private static List<double> GetCoefs(Excel.Worksheet oddsSeason, int row, double handicap)
		{
			string value = (oddsSeason.Cells[row, 13] as Excel.Range).Text; // 13 - column M - closing AH
			if (double.TryParse(value, out double handicapVar))
			{
				if (handicapVar == handicap)
				{
					double handicapOddsA = double.Parse((oddsSeason.Cells[row, 14] as Excel.Range).Text); // 14, 15, 5, 6 - columns N, O, E, F - closing odds Columns
					double handicapOddsB = double.Parse((oddsSeason.Cells[row, 15] as Excel.Range).Text);
					double moneylineOddsA = double.Parse((oddsSeason.Cells[row, 5] as Excel.Range).Text);
					double moneylineOddsB = double.Parse((oddsSeason.Cells[row, 6] as Excel.Range).Text);

					return new List<double>() { moneylineOddsA, moneylineOddsB, handicapOddsA, handicapOddsB };
				}
			}

			return null;
		}

		private static void CopyCoefsToConvStats(Excel.Workbook workbook, List<List<double>> allCoefs, double handicap, int column)
		{
			Excel.Worksheet convStats = workbook.Sheets
			.Cast<Excel.Worksheet>()
			.Where(s => s.Name == sheetConvStatsName)
			.FirstOrDefault();

			for (int i = 0; i < allCoefs.Count; i++)
			{
				convStats.Cells[i + 4, column] = allCoefs[i][0]; // first 3 rows are occupied by header
				convStats.Cells[i + 4, column + 1] = allCoefs[i][1];
				convStats.Cells[i + 4, column + 2] = allCoefs[i][2];
				convStats.Cells[i + 4, column + 3] = allCoefs[i][3];
			}

		}
	}
}

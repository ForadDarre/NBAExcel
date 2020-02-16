using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace NBAExcel
{
	public partial class NBARibbon
	{
		private void NBARibbon_Load(object sender, RibbonUIEventArgs e)
		{

		}

		private void handicapCollection_Click(object sender, RibbonControlEventArgs e)
		{
			HandicapCollection.Main();
		}

		private void updateCalendar_Click(object sender, RibbonControlEventArgs e)
		{
			UpdateMinsAndPlusMinus.Main();
		}
	}
}

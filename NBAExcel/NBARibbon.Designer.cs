namespace NBAExcel
{
	partial class NBARibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public NBARibbon()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

		/// <summary> 
		/// Освободить все используемые ресурсы.
		/// </summary>
		/// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Код, автоматически созданный конструктором компонентов

		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		private void InitializeComponent()
		{
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NBARibbon));
			this.tab1 = this.Factory.CreateRibbonTab();
			this.NBA = this.Factory.CreateRibbonGroup();
			this.handicapCollection = this.Factory.CreateRibbonButton();
			this.updateCalendar = this.Factory.CreateRibbonButton();
			this.tab1.SuspendLayout();
			this.NBA.SuspendLayout();
			this.SuspendLayout();
			// 
			// tab1
			// 
			this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.tab1.Groups.Add(this.NBA);
			this.tab1.Label = "TabAddIns";
			this.tab1.Name = "tab1";
			// 
			// NBA
			// 
			this.NBA.Items.Add(this.handicapCollection);
			this.NBA.Items.Add(this.updateCalendar);
			this.NBA.Label = "NBA";
			this.NBA.Name = "NBA";
			// 
			// handicapCollection
			// 
			this.handicapCollection.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.handicapCollection.Image = ((System.Drawing.Image)(resources.GetObject("handicapCollection.Image")));
			this.handicapCollection.Label = "Handicap Collection";
			this.handicapCollection.Name = "handicapCollection";
			this.handicapCollection.ScreenTip = "Collect all closing odds (AH and moneyline) from \"Odds season\" sheet and sort col" +
    "lected values in \"Conv stats\" sheet.";
			this.handicapCollection.ShowImage = true;
			this.handicapCollection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.handicapCollection_Click);
			// 
			// updateCalendar
			// 
			this.updateCalendar.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.updateCalendar.Image = ((System.Drawing.Image)(resources.GetObject("updateCalendar.Image")));
			this.updateCalendar.Label = "Update Calendar";
			this.updateCalendar.Name = "updateCalendar";
			this.updateCalendar.ScreenTip = "Update ranges in \"Mins\" sheet (7 days, 14 days). Actual date is stored in B1082 c" +
    "ell of \"Input\" sheet.";
			this.updateCalendar.ShowImage = true;
			this.updateCalendar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.updateCalendar_Click);
			// 
			// NBARibbon
			// 
			this.Name = "NBARibbon";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.tab1);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.NBARibbon_Load);
			this.tab1.ResumeLayout(false);
			this.tab1.PerformLayout();
			this.NBA.ResumeLayout(false);
			this.NBA.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup NBA;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton handicapCollection;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton updateCalendar;
	}

	partial class ThisRibbonCollection
	{
		internal NBARibbon NBARibbon
		{
			get { return this.GetRibbon<NBARibbon>(); }
		}
	}
}

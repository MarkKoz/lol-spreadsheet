using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace LoL_Spreadsheet.Form
{
	/// <summary>
	/// Contains methods for adding controls to a <see cref="System.Windows.Forms.Form"/>.
	/// </summary>
	public class FormControls
	{
		// Declares and initialises controls
		// Containers
		public GroupBox GrpGen = new GroupBox();
		public GroupBox GrpStats = new GroupBox();
		public GroupBox GrpRank = new GroupBox();
		public GroupBox GrpChamp = new GroupBox();
		public TabControl TabComm = new TabControl();
		public GroupBox GrpSett = new GroupBox();

		// grpGen
		public CheckBox GenChkScreen = new CheckBox();
		public TextBox GenTxtScreen = new TextBox();
		public Label GenLblLength = new Label();
		public Label GenLblLengthM = new Label();
		public NumericUpDown GenNumLengthM = new NumericUpDown();
		public Label GenLblLengthS = new Label();
		public NumericUpDown GenNumLengthS = new NumericUpDown();

		// grpStats
		public Label StatsLblK = new Label();
		public NumericUpDown StatsNumK = new NumericUpDown();
		public Label StatsLblD = new Label();
		public NumericUpDown StatsNumD = new NumericUpDown();
		public Label StatsLblA = new Label();
		public NumericUpDown StatsNumA = new NumericUpDown();
		public Label StatsLblCS = new Label();
		public NumericUpDown StatsNumCS = new NumericUpDown();
		public Label StatsLblGold = new Label();
		public TextBox StatsTxtGold = new TextBox();

		// grpRank
		public Label RankLblRank = new Label();
		public ComboBox RankCmbRank = new ComboBox();
		public Label RankLblLP = new Label();
		public NumericUpDown RankNumLP = new NumericUpDown();
		public CheckBox RankChkDodge = new CheckBox();
		public NumericUpDown RankNumDodge = new NumericUpDown();

		// grpChamp
		public Label ChampLblRole = new Label();
		public ComboBox ChampCmbRole = new ComboBox();
		public Label ChampLblChamp = new Label();
		public ComboBox ChampCmbChamp = new ComboBox();
		public Label ChampLblOpp = new Label();
		public ComboBox ChampCmbOpp = new ComboBox();
		public Label ChampLblGrade = new Label();
		public ComboBox ChampTxtGrade = new ComboBox();

		// tabComments
		public TabPage CommPgLane = new TabPage();
		public TabPage CommPgProb = new TabPage();
		public TabPage CommPgOther = new TabPage();
		public TextBox CommTxtLane = new TextBox();
		public TextBox CommTxtProb = new TextBox();
		public TextBox CommTxtOther = new TextBox();

		// grpSettings
		public CheckBox SettChkDate = new CheckBox();
		public DateTimePicker SettDtpDate = new DateTimePicker();
		public CheckBox SettChkClearRank = new CheckBox();
		public CheckBox SettChkSubClear = new CheckBox();
		public CheckBox SettChkSubClose = new CheckBox();
		public CheckBox SettChkSave = new CheckBox();

		// Buttons
		public Button BtnSubmit = new Button();
		public Button BtnClear = new Button();
		public Button BtnCancel = new Button();

		private readonly Form _f;

		/// <summary>
		/// 
		/// </summary>
		/// <param name="f">The <see cref="System.Windows.Forms.Form"/> to add controls to.</param>
		public FormControls(Form f)
		{
			_f = f;
		}

		/// <summary>
		/// Adds controls to <see cref="System.Windows.Forms.Form"/> <see cref="_f"/>.
		/// </summary>
		public void Add()
		{
			// Containers
			_f.Controls.Add(GrpGen);
			_f.Controls.Add(GrpStats);
			_f.Controls.Add(GrpRank);
			_f.Controls.Add(GrpChamp);
			_f.Controls.Add(TabComm);
			_f.Controls.Add(GrpSett);

			// grpGen
			GrpGen.Controls.Add(GenChkScreen);
			GrpGen.Controls.Add(GenTxtScreen);
			GrpGen.Controls.Add(GenLblLength);
			GrpGen.Controls.Add(GenLblLengthM);
			GrpGen.Controls.Add(GenNumLengthM);
			GrpGen.Controls.Add(GenLblLengthS);
			GrpGen.Controls.Add(GenNumLengthS);

			// grpStats
			GrpStats.Controls.Add(StatsLblK);
			GrpStats.Controls.Add(StatsNumK);
			GrpStats.Controls.Add(StatsLblD);
			GrpStats.Controls.Add(StatsNumD);
			GrpStats.Controls.Add(StatsLblA);
			GrpStats.Controls.Add(StatsNumA);
			GrpStats.Controls.Add(StatsLblCS);
			GrpStats.Controls.Add(StatsNumCS);
			GrpStats.Controls.Add(StatsLblGold);
			GrpStats.Controls.Add(StatsTxtGold);

			// grpRank
			GrpRank.Controls.Add(RankLblRank);
			GrpRank.Controls.Add(RankCmbRank);
			GrpRank.Controls.Add(RankLblLP);
			GrpRank.Controls.Add(RankNumLP);
			GrpRank.Controls.Add(RankChkDodge);
			GrpRank.Controls.Add(RankNumDodge);

			// grpChamp
			GrpChamp.Controls.Add(ChampLblRole);
			GrpChamp.Controls.Add(ChampCmbRole);
			GrpChamp.Controls.Add(ChampLblChamp);
			GrpChamp.Controls.Add(ChampCmbChamp);
			GrpChamp.Controls.Add(ChampLblOpp);
			GrpChamp.Controls.Add(ChampCmbOpp);
			GrpChamp.Controls.Add(ChampLblGrade);
			GrpChamp.Controls.Add(ChampTxtGrade);

			// tabComm
			TabComm.Controls.Add(CommPgLane);
			TabComm.Controls.Add(CommPgProb);
			TabComm.Controls.Add(CommPgOther);
			CommPgLane.Controls.Add(CommTxtLane);
			CommPgProb.Controls.Add(CommTxtProb);
			CommPgOther.Controls.Add(CommTxtOther);

			// grpSett
			GrpSett.Controls.Add(SettChkDate);
			GrpSett.Controls.Add(SettDtpDate);
			GrpSett.Controls.Add(SettChkClearRank);
			GrpSett.Controls.Add(SettChkSubClear);
			GrpSett.Controls.Add(SettChkSubClose);
			GrpSett.Controls.Add(SettChkSave);

			// Buttons
			_f.Controls.Add(BtnSubmit);
			_f.Controls.Add(BtnClear);
			_f.Controls.Add(BtnCancel);
		}

		/// <summary>
		/// Retrieves all child controls of a control.
		/// </summary>
		/// <param name="control">The control from which to retrieve all child controls.</param>
		/// <param name="type">(Optional) The type of the controls to retrieve.</param>
		/// <returns></returns>
		/// <remarks>Code modified from stackoverflow.com/a/3426721/5717792</remarks>
		public IEnumerable<Control> GetAllControls(Control control, Type type = null)
		{
			var controls = control.Controls.Cast<Control>();

			if (type != null)
			{
				return
					controls.SelectMany(ctrl => GetAllControls(ctrl, type))
						.Concat(controls)
						.Where(c => c.GetType() == type);
			}
			return controls.SelectMany(ctrl => GetAllControls(ctrl)).Concat(controls);
		}

		public void SetPropertyBulk(string p, object val, Type t = null,
			Control parent = null)
		{
			Control par = parent ?? _f;
			foreach (Control c in GetAllControls(par, t))
			{
				PropertyInfo prop = c.GetType()
					.GetProperty(p, BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);

				if (prop != null && prop.CanWrite)
				{
					try
					{
						prop.SetValue(c, val, null);
					}
					catch (ArgumentException e)
					{
						Debug.Print(
							$"The given value {val} is not valid for the given property {p} in control {c}.");
						//throw new ArgumentException(
						//	$"The given value {val} is not valid for the given property {p}.", e);
					}
				}
				else
				{
					Debug.Print($"The given property {p} does not exist for control {c}.");
				}
			}
		}

		/// <summary>
		/// Sets properties for <see cref="System.Windows.Forms.Form"/> <see cref="_f"/> and its controls.
		/// </summary>
		public void Properties()
		{
			//temp
			//grpStats.Hide();
			GrpChamp.Hide();
			GrpRank.Hide();
			GrpSett.Hide();
			TabComm.Hide();

			const int padE = 16;
			const int padB = 16;

			const int padEIn = 8;
			const int padBIn = 8;

			const int cHeight = 16;

			SetPropertyBulk("TextAlign", ContentAlignment.MiddleLeft, typeof(Label));
			SetPropertyBulk("Left", padEIn, typeof(Label));
			SetPropertyBulk("Margin", Padding.Empty);
			SetPropertyBulk("Padding", Padding.Empty);
			SetPropertyBulk("Padding", Point.Empty, typeof(TabControl));

			// Height
			SetPropertyBulk("MinimumSize", new Size(0, cHeight), typeof(Label));
			SetPropertyBulk("AutoSize", true, typeof(Label));
			SetPropertyBulk("Height", cHeight, typeof(TextBox));
			SetPropertyBulk("Height", cHeight, typeof(NumericUpDown));

			// Form
			int bSize = SystemInformation.FrameBorderSize.Width * 2;
			int cSize = SystemInformation.CaptionHeight + bSize;

			Debug.Print($"Border width: {bSize}");
			Debug.Print($"Caption Height: {cSize}");
			Debug.Print($"Caption Height (No Border): {cSize - bSize}");

			_f.Text = "Enter New Match";
			_f.ClientSize = new Size(512, 512);
			_f.MinimumSize = new Size(512 + bSize * 2, 512 + cSize + bSize);
			_f.MaximumSize = new Size(512 + bSize * 2, 512 + cSize + bSize);
			// TODO Remember user's window position stackoverflow.com/a/108217/5717792
			_f.StartPosition = FormStartPosition.WindowsDefaultLocation;

			int fW = _f.ClientSize.Width;
			int fH = _f.ClientSize.Height;

			int cX = fW / 2;

			// GrpGen
			GrpGen.Text = "General";
			GrpGen.Width = 256;
			GrpGen.Location = new Point(bSize + padE, -6 + padE);

			GenChkScreen.Text = "Screenshot";
			GenChkScreen.Location = new Point(padEIn, 7 + padEIn);
			GenChkScreen.CheckState = CheckState.Checked;

			GenTxtScreen.Width = 64;
			GenTxtScreen.Location = new Point(128, 7 + padEIn);

			GenLblLength.Text = "Match Length:";
			GenLblLength.Top = GenChkScreen.Bottom + padBIn;

			GenLblLengthM.Text = "M:";
			GenLblLengthM.Location = new Point(96, GenTxtScreen.Bottom + padBIn);

			GenLblLengthS.Text = "S:";
			GenLblLengthS.Location = new Point(160, GenTxtScreen.Bottom + padBIn);

			GenNumLengthM.Width = 32;
			GenNumLengthM.Location = new Point(128, GenTxtScreen.Bottom + padBIn);

			GenNumLengthS.Width = 32;
			GenNumLengthS.Location = new Point(192, GenTxtScreen.Bottom + padBIn);

			GrpGen.Height = GenNumLengthM.Bottom + padEIn;

			// GrpStats
			SetPropertyBulk("Width", 64, typeof(TextBox), GrpStats);
			SetPropertyBulk("Width", 64, typeof(NumericUpDown), GrpStats);

			GrpStats.Text = "Statistics";
			GrpStats.Width = 256;
			GrpStats.Location = new Point(bSize + padE, GrpGen.Bottom - 6 + padB);

			StatsLblK.Text = "Kills:";
			StatsLblK.Top = 7 + padEIn;

			StatsNumK.Location = new Point(128, 7 + padEIn);

			StatsLblD.Text = "Deaths:";
			StatsLblD.Top = StatsLblK.Bottom + padBIn;

			StatsNumD.Location = new Point(128, StatsLblK.Bottom + padBIn);

			StatsLblA.Text = "Assists:";
			StatsLblA.Top = StatsLblD.Bottom + padBIn;

			StatsNumA.Location = new Point(128, StatsLblD.Bottom + padBIn);

			StatsLblCS.Text = "CS:";
			StatsLblCS.Top = StatsLblA.Bottom + padBIn;

			StatsNumCS.Location = new Point(128, StatsLblA.Bottom + padBIn);

			StatsLblGold.Text = "Gold:";
			StatsLblGold.Top = StatsLblCS.Bottom + padBIn;

			StatsTxtGold.Location = new Point(128, StatsLblCS.Bottom + padBIn);

			GrpStats.Height = StatsLblGold.Bottom + padEIn;

			// GrpRank
			GrpRank.Text = "Rank";
			GrpRank.Size = new Size(128, 128);
			GrpRank.Location = new Point(bSize + 16, cSize + 16);

			// GrpChamp
			GrpChamp.Text = "Champion and Lane";
			GrpChamp.Size = new Size(128, 128);
			GrpChamp.Location = new Point(bSize + 16, cSize + 16);

			// TabComm
			TabComm.Text = "Comments";
			TabComm.Size = new Size(128, 128);
			TabComm.Location = new Point(bSize + 16, cSize + 16);

			// GrpSett
			GrpSett.Text = "Settings";
			GrpSett.Size = new Size(128, 128);
			GrpSett.Location = new Point(bSize + 16, cSize + 16);

			// Buttons
			const int bPadY = 16;

			BtnClear.Text = "Clear";
			BtnClear.Size = new Size(72, 24);
			BtnClear.Location = new Point(cX - BtnClear.Width/2,
				fH - BtnClear.Height - bPadY);

			BtnSubmit.Text = "Submit";
			BtnSubmit.Size = new Size(72, 24);
			BtnSubmit.Location = new Point(BtnClear.Location.X/2 -
				BtnSubmit.Width/2,
				fH - BtnSubmit.Height - bPadY);

			BtnCancel.Text = "Cancel";
			BtnCancel.Size = new Size(72, 24);
			BtnCancel.Location = new Point(BtnClear.Location.X + BtnClear.Width +
				((fW - (BtnClear.Location.X + BtnClear.Width))/2 -
				BtnCancel.Width/2),
				fH - BtnCancel.Height - bPadY);
		}
	}
}

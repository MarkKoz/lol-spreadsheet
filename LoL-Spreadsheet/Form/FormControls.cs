using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
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
		GroupBox grpGen = new GroupBox();
		GroupBox grpStats = new GroupBox();
		GroupBox grpRank = new GroupBox();
		GroupBox grpChamp = new GroupBox();
		TabControl tabComm = new TabControl();
		GroupBox grpSett = new GroupBox();

		// grpGen
		CheckBox Gen_chkScreen = new CheckBox();
		TextBox Gen_txtScreen = new TextBox();
		Label Gen_lblLength = new Label();
		Label Gen_lblLength_M = new Label();
		NumericUpDown Gen_numLength_M = new NumericUpDown();
		Label Gen_lblLength_S = new Label();
		NumericUpDown Gen_numLength_S = new NumericUpDown();

		// grpStats
		Label Stats_lblK = new Label();
		NumericUpDown Stats_numK = new NumericUpDown();
		Label Stats_lblD = new Label();
		NumericUpDown Stats_numD = new NumericUpDown();
		Label Stats_lblA = new Label();
		NumericUpDown Stats_numA = new NumericUpDown();
		Label Stats_lblCS = new Label();
		NumericUpDown Stats_numCS = new NumericUpDown();
		Label Stats_lblGold = new Label();
		TextBox Stats_txtGold = new TextBox();

		// grpRank
		Label Rank_lblRank = new Label();
		ComboBox Rank_cmbRank = new ComboBox();
		Label Rank_lblLP = new Label();
		NumericUpDown Rank_numLP = new NumericUpDown();
		CheckBox Rank_chkDodge = new CheckBox();
		NumericUpDown Rank_numDodge = new NumericUpDown();

		// grpChamp
		Label Champ_lblRole = new Label();
		ComboBox Champ_cmbRole = new ComboBox();
		Label Champ_lblChamp = new Label();
		ComboBox Champ_cmbChamp = new ComboBox();
		Label Champ_lblOpp = new Label();
		ComboBox Champ_cmbOpp = new ComboBox();
		Label Champ_lblGrade = new Label();
		ComboBox Champ_txtGrade = new ComboBox();

		// tabComments
		TabPage Comm_pgLane = new TabPage();
		TabPage Comm_pgProb = new TabPage();
		TabPage Comm_pgOther = new TabPage();
		TextBox Comm_txtLane = new TextBox();
		TextBox Comm_txtProb = new TextBox();
		TextBox Comm_txtOther = new TextBox();

		// grpSettings
		CheckBox Sett_chkDate = new CheckBox();
		DateTimePicker Sett_dtpDate = new DateTimePicker();
		CheckBox Sett_chkClearRank = new CheckBox();
		CheckBox Sett_chkSubClear = new CheckBox();
		CheckBox Sett_chkSubClose = new CheckBox();
		CheckBox Sett_chkSave = new CheckBox();

		// Buttons
		Button btnSubmit = new Button();
		Button btnClear = new Button();
		Button btnCancel = new Button();

		protected Form f;

		/// <summary>
		/// 
		/// </summary>
		/// <param name="f">The <see cref="System.Windows.Forms.Form"/> to add controls to.</param>
		public FormControls(Form f)
		{
			this.f = f;
		}

		/// <summary>
		/// Adds controls to <see cref="System.Windows.Forms.Form"/> <see cref="f"/>.
		/// </summary>
		public void Add()
		{
			// Containers
			f.Controls.Add(grpGen);
			f.Controls.Add(grpStats);
			f.Controls.Add(grpRank);
			f.Controls.Add(grpChamp);
			f.Controls.Add(tabComm);
			f.Controls.Add(grpSett);

			// grpGen
			grpGen.Controls.Add(Gen_chkScreen);
			grpGen.Controls.Add(Gen_txtScreen);
			grpGen.Controls.Add(Gen_lblLength);
			grpGen.Controls.Add(Gen_lblLength_M);
			grpGen.Controls.Add(Gen_numLength_M);
			grpGen.Controls.Add(Gen_lblLength_S);
			grpGen.Controls.Add(Gen_numLength_S);

			// grpStats
			grpStats.Controls.Add(Stats_lblK);
			grpStats.Controls.Add(Stats_numK);
			grpStats.Controls.Add(Stats_lblD);
			grpStats.Controls.Add(Stats_numD);
			grpStats.Controls.Add(Stats_lblA);
			grpStats.Controls.Add(Stats_numA);
			grpStats.Controls.Add(Stats_lblCS);
			grpStats.Controls.Add(Stats_numCS);
			grpStats.Controls.Add(Stats_lblGold);
			grpStats.Controls.Add(Stats_txtGold);

			// grpRank
			grpRank.Controls.Add(Rank_lblRank);
			grpRank.Controls.Add(Rank_cmbRank);
			grpRank.Controls.Add(Rank_lblLP);
			grpRank.Controls.Add(Rank_numLP);
			grpRank.Controls.Add(Rank_chkDodge);
			grpRank.Controls.Add(Rank_numDodge);

			// grpChamp
			grpChamp.Controls.Add(Champ_lblRole);
			grpChamp.Controls.Add(Champ_cmbRole);
			grpChamp.Controls.Add(Champ_lblChamp);
			grpChamp.Controls.Add(Champ_cmbChamp);
			grpChamp.Controls.Add(Champ_lblOpp);
			grpChamp.Controls.Add(Champ_cmbOpp);
			grpChamp.Controls.Add(Champ_lblGrade);
			grpChamp.Controls.Add(Champ_txtGrade);

			// tabComm
			tabComm.Controls.Add(Comm_pgLane);
			tabComm.Controls.Add(Comm_pgProb);
			tabComm.Controls.Add(Comm_pgOther);
			Comm_pgLane.Controls.Add(Comm_txtLane);
			Comm_pgProb.Controls.Add(Comm_txtProb);
			Comm_pgOther.Controls.Add(Comm_txtOther);

			// grpSett
			grpSett.Controls.Add(Sett_chkDate);
			grpSett.Controls.Add(Sett_dtpDate);
			grpSett.Controls.Add(Sett_chkClearRank);
			grpSett.Controls.Add(Sett_chkSubClear);
			grpSett.Controls.Add(Sett_chkSubClose);
			grpSett.Controls.Add(Sett_chkSave);

			// Buttons
			f.Controls.Add(btnSubmit);
			f.Controls.Add(btnClear);
			f.Controls.Add(btnCancel);
		}

		/// <summary>
		/// Retrieves all child controls of a control.
		/// </summary>
		/// <param name="control">The control from which to retrieve all child controls.</param>
		/// <param name="type">(Optional) The type of the the controls to retrieve.</param>
		/// <returns></returns>
		/// <remarks>Code modified from stackoverflow.com/a/3426721/5717792</remarks>
		public IEnumerable<Control> GetAllControls(Control control, Type type = null)
		{
			var controls = control.Controls.Cast<Control>();

			if (type != null)
			{
				return controls.SelectMany(ctrl => GetAllControls(ctrl, type))
										  .Concat(controls)
										  .Where(c => c.GetType() == type);
			}
			return controls.SelectMany(ctrl => GetAllControls(ctrl))
									  .Concat(controls);
		}

		/// <summary>
		/// Sets properties for <see cref="System.Windows.Forms.Form"/> <see cref="f"/> and its controls.
		/// </summary>
		public void Properties()
		{
			//foreach (Control c in GetAllControls(f))
			//{
			//	// user 'statue' from stackoverflow.com/a/20661481/5717792
			//	if (c.GetType().GetProperty(btnSubmit.Anchor.ToString()) !=null)
			//	{
			//		c.Anchor = AnchorStyles.None;
			//	}
			//}

			// Form
			int bSize = SystemInformation.FrameBorderSize.Width * 2;
			int cSize = SystemInformation.CaptionHeight + bSize;

			f.ClientSize = new Size(512, 512);
			f.MinimumSize = new Size(512 + bSize, 512 + cSize);
			f.MaximumSize = new Size(512 + bSize, 512 + cSize);

			// grpGen
			Gen_chkScreen.CheckState = CheckState.Checked;

			// grpStats


			// grpRank


			// grpChamp


			// tabComm


			// grpSett


			// Buttons
		}
	}
}

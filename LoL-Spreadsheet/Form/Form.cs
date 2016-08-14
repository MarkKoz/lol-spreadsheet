using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace LoL_Spreadsheet.Form
{
	public partial class Form : System.Windows.Forms.Form
	{
		private FormControls c;

		public Form()
		{
			InitializeComponent();
		}

		private void Form_Load(object sender, System.EventArgs e)
		{
			c = new FormControls(this);
			c.Add();
			c.Properties();
		}
	}
}

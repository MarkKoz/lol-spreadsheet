namespace LoL_Spreadsheet.Form
{
	public partial class Form : System.Windows.Forms.Form
	{
		private readonly FormControls _c;

		public Form()
		{
			InitializeComponent();
			_c = new FormControls(this);
		}

		private void Form_Load(object sender, System.EventArgs e)
		{
			
			_c.Add();
			_c.Properties();
		}
	}
}

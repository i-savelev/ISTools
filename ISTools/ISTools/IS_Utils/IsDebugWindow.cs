using System.Data;


namespace ISTools
{
    public static class IsDebugWindow
    {
        public static DataTable DtSheets { get; set; }
        static IsDebugWindow()
        {
            DtSheets = new DataTable();
            DtSheets.Columns.Add("данные");
        }
        public static void Show(string windowName = null)
        {
            Debugger debugger = new Debugger();
            debugger.Text = windowName;
            if (DtSheets.Rows.Count > 0)
            {
                debugger.debugTable.DataSource = DtSheets;
                debugger.FormClosing += (s, e) => { DtSheets.Clear(); };
                debugger.Show();
            }

        }
        public static void AddRow(string str)
        {
            DtSheets.Rows.Add(str);
        }
    }
}

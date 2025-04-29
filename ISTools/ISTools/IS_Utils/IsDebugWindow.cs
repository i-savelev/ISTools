using System.Data;


namespace ISTools.Utils
{
    public static class IsDebugWindow
    {
        public static DataTable DtSheets { get; set; }
        static IsDebugWindow()
        {
            DtSheets = new DataTable();
            DtSheets.Columns.Add("данные");
        }
        public static void Show()
        {
            Debugger debugger = new Debugger();
            if (DtSheets.Rows.Count > 0)
            {
                debugger.debugTable.DataSource = DtSheets;
                debugger.Show();
            }

        }
        public static void AddRow(string str)
        {
            DtSheets.Rows.Add(str);
        }
    }
}

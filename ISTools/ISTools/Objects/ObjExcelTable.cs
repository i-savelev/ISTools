using OfficeOpenXml.Table;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Windows.Forms;
using System.IO;
using System.Data;

namespace ISTools
{
    internal class ObjExcelTable
    {
        public DataTable dt;
        public string filename = "таблица";
        public string worksheets = "1";
        public Dictionary<string, DataTable> dictWorksheets = new Dictionary<string, DataTable>();
        public void TableExport()
        {
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(worksheets);
                worksheet.Cells["A1"].LoadFromDataTable(dt, true);
                int firstRow = 1;
                int lastRow = worksheet.Dimension.End.Row;
                int firstColumn = 1;
                int lastColumn = worksheet.Dimension.End.Column;
                ExcelRange rg = worksheet.Cells[firstRow, firstColumn, lastRow, lastColumn];
                string tableName = "Table1";
                OfficeOpenXml.Table.ExcelTable tab = worksheet.Tables.Add(rg, tableName);
                tab.TableStyle = TableStyles.Light8;
                byte[] bin = excelPackage.GetAsByteArray();
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Title = "Выберете папку для сохранения отчета";
                saveFileDialog1.Filter = "Excel files|*.xlsx|All files|*.*";
                saveFileDialog1.FileName = $"{filename}.xlsx";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllBytes(saveFileDialog1.FileName, bin);
                }
            }
        }
        public void TablesExport()
        {
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                var n = 1;
                foreach (var sheet in dictWorksheets)
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(sheet.Key);
                    worksheet.Cells["A1"].LoadFromDataTable(sheet.Value, true);
                    int firstRow = 1;
                    int lastRow = worksheet.Dimension.End.Row;
                    int firstColumn = 1;
                    int lastColumn = worksheet.Dimension.End.Column;
                    ExcelRange rg = worksheet.Cells[firstRow, firstColumn, lastRow, lastColumn];
                    string tableName = $"Table{n}";
                    OfficeOpenXml.Table.ExcelTable tab = worksheet.Tables.Add(rg, tableName);
                    tab.TableStyle = TableStyles.Light8;
                    n += 1;
                }
                byte[] bin = excelPackage.GetAsByteArray();
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Title = "Выберете папку для сохранения отчета";
                saveFileDialog1.Filter = "Excel files|*.xlsx|All files|*.*";
                saveFileDialog1.FileName = $"{filename}.xlsx";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllBytes(saveFileDialog1.FileName, bin);
                }
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;


namespace DBsExecuter.Classes
{
    /// <summary>
    /// 2123
    /// </summary>
    static class ExcelHelper
    {
        public static void FillExcel(string pathToDirecory, List<Statistic> report)
        {
            string pathToExcl = Path.Combine(pathToDirecory, DateTime.Now.Ticks+".xls");
            using (var excl = new ExcelPackage(new FileInfo(pathToExcl)))
            {
                int i = 2;
                ExcelWorksheet ws = excl.Workbook.Worksheets.Add("1");
                ws.Cells[1, 1].Value = "DbName";
                ws.Cells[1, 2].Value = "CmdName";
                ws.Cells[1, 3].Value = "CmdType";
                ws.Cells[1, 4].Value = "Cnt";
                foreach (var rep in report)
                {
                    ws.Cells[i, 1].Value = rep.DbName;
                    ws.Cells[i, 2].Value = rep.CmdName;
                    ws.Cells[i, 3].Value = rep.CmdType;
                    ws.Cells[i, 4].Value = rep.Cnt;
                    i++;
                }
                excl.Save();
            }
        }
    }
}

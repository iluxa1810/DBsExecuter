using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;


namespace DBsExecuter.Classes
{
    static class ExcelHelper
    {
        public static void FillExcel(string pathToDirecory, List<Statistic> report)
        {
            string pathToExcl = Path.Combine(pathToDirecory, DateTime.Now+".xls");
            using (var excl = new ExcelPackage(new FileInfo(pathToExcl)))
            {
                ExcelWorksheet ws = excl.Workbook.Worksheets.Add("1");
                for (int i = 1; i <= 3; i++)
                {
                    ws.Cells[i, i].Value = 1;
                }
                excl.Save();
            }
        }
    }
}

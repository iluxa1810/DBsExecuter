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
        static void FillExcel(string pathToExcel, List<Statistic> report)
        {
            using (var excl=new ExcelPackage(new FileInfo(pathToExcel)))
            {
                ExcelWorksheet ws = excl.Workbook.Worksheets.Add("1");
            }
        }
    }
}

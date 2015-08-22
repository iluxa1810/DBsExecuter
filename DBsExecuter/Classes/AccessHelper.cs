using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBsExecuter.Classes
{
    class Statistic
    {
        public string DbName { get; set; }
        public string CmdName { get; set; }
        public string CmdType { get; set; }
        public int Cnt { get; set; }

    }
    static class AccessHelper
    {
        public static List<OleDbConnectionStringBuilder> GetConnsStrings(string pathToDb)
        {
            var oleDb = new List<OleDbConnectionStringBuilder>();
            foreach (var file in Directory.GetFiles(pathToDb,"*.mdb"))
            {
                oleDb.Add(
                    new OleDbConnectionStringBuilder()
                    {
                        Provider = "Microsoft.ACE.OLEDB.12.0",
                        DataSource = file,
                        PersistSecurityInfo = false
                    });
            }
            return oleDb;
        }
        public static void ExecuteCommands(List<OleDbConnectionStringBuilder> connsStr, Package commands, List<Statistic> report)
        {
            foreach (var connStr in connsStr)
            {
                using (var conn = new OleDbConnection(connStr.ConnectionString))
                {
                    conn.Open();
                    foreach (var command in commands.Tasks)
                    {
                        report.Add(new Statistic()
                        {
                            CmdName = command.QueryName,
                            CmdType = command.QueryType,
                            DbName = Path.GetFileName(conn.DataSource)
                        });
                        OleDbCommand cmd = new OleDbCommand(command.Query, conn);
                        if (command.QueryType == "UPDATE")
                        {
                            report[report.Count - 1].Cnt = cmd.ExecuteNonQuery();
                            continue;
                        }
                        if (command.QueryType == "SCALAR")
                        {
                            report[report.Count - 1].Cnt = (int)cmd.ExecuteScalar();
                        }
                    }
                }
            }
        }
    }
}

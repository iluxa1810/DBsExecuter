using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using DBsExecuter.Classes;

namespace DBsExecuter
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Package pck = XmlHelper.GetXmlData(textBox1.Text);
            var conns=AccessHelper.GetConnsStrings(@"C:\Users\iluxa1810\Documents\visual studio 2015\Projects\DBsExecuter\DBsExecuter\dbs\");
            List<Statistic> stat =new List<Statistic>();
            AccessHelper.ExecuteCommands(conns, pck, stat);
            ExcelHelper.FillExcel(@"D:\Laba1.xls", stat);
        }
    }
}

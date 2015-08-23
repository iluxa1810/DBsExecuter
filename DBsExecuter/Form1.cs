using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        public static void Logger(string str, RichTextBox logBox)
        {
            logBox.AppendText("\n" + str + "\n");
            logBox.Refresh();
            logBox.ScrollToCaret();
        }

        private static void CheckPath(params string[] pathes)
        {
            if (!pathes.Any())
            {
                throw new Exception("Ошибка в путях");
            }
            foreach (var path in pathes)
            {
                if (!Path.IsPathRooted(path))
                {
                    throw new Exception("Ошибка в путях");
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Logger("Проверяю пути", logBox);
                CheckPath(textBox1.Text, textBox2.Text, textBox3.Text);
                Logger("Считываю конфигурацию", logBox);
                Package pck = XmlHelper.GetXmlData(textBox2.Text);
                Logger("Получаю строки подключения", logBox);
                var conns = AccessHelper.GetConnsStrings(textBox1.Text);
                List<Statistic> stat = new List<Statistic>();
                Logger("Выполняю запросы", logBox);
                AccessHelper.ExecuteCommands(conns, pck, stat,logBox);
                Logger("Создаю Excel", logBox);
                ExcelHelper.FillExcel(textBox3.Text, stat);
                Logger("Программа успешно завершена", logBox);
            }
            catch (Exception ex)
            {
                Logger(ex.ToString(), logBox);
                Logger("Программа завершилась с ошибками", logBox);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            var f = new FolderBrowserDialog();
            f.ShowDialog();
            textBox1.Text = f.SelectedPath;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var f = new OpenFileDialog();
            f.Filter = "File *.pck|*.pck";
            f.ShowDialog();
            textBox2.Text = f.FileName;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var f = new FolderBrowserDialog();
            f.ShowDialog();
            textBox3.Text = f.SelectedPath;
        }
    }
}

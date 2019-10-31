using System;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace UpdateExcels
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            label1.Text = "";
            button2.Enabled = false;
            folderBrowserDialog1.SelectedPath = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            label1.Text = "";
            button2.Enabled = false;
            folderBrowserDialog1.SelectedPath = "";
            if(folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                label1.Text = folderBrowserDialog1.SelectedPath;
                button2.Enabled = true;

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb;
            string[] files = System.IO.Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.xlsb", System.IO.SearchOption.AllDirectories);
            foreach (string file in files)
            {
                try
                {
                    sb.Append(file);
                    wb = excel.Workbooks.Open(file);
                    wb.SaveAs(file.Replace(".xlsb", ".xlsx"), Excel.XlFileFormat.xlWorkbookDefault);
                    wb.Close();
                    sb.Append( " -> " + file.Replace(".xlsb", ".xlsx"));

                }
                catch(Exception ex)
                {
                    sb.Append(" - error: " + ex.Message);
                }
                sb.Append(Environment.NewLine);
            }
            
            label2.Text = sb.ToString();
        }
    }
}

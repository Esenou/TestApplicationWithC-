using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace TestApplication
{
    public partial class Form1 : Form
    {

       // private Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename =| DataDirectory |\Database1.mdf; Integrated Security = True
        public Form1()
        {
            InitializeComponent();
        }

        private void tableBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.tableBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.database1DataSet);
           // Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename =| DataDirectory |\Database1.mdf; Integrated Security = True
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "database1DataSet.Table". При необходимости она может быть перемещена или удалена.
            this.tableTableAdapter.Fill(this.database1DataSet.Table);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string file, socialnumber = "12345678901234";
            string sourceFile = @"C:\Users\Acer\source\repos\TestApplication\TestApplication\Template\example.xlsx";
            string destinationFile = @"C:\Users\Acer\source\repos\TestApplication\TestApplication\Result\example.xlsx";

            try
            {
                File.Copy(sourceFile, destinationFile, true);
                file = @"C:\Users\Acer\source\repos\TestApplication\TestApplication\Result\example.xlsx";

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWb = xlApp.Workbooks.Open(file);
                Microsoft.Office.Interop.Excel.Worksheet xlSht = xlWb.Sheets[1];

                DataTable table = database1DataSet.Tables["Table"];
                DataRow[] rows = table.Select();

                // Print the value one column of each DataRow.
                for (int i = 0; i < rows.Length; i++)
                {
                    if (rows[i]["socialnumber"].Equals(socialnumber))
                    {
                        xlSht.Cells[3, "B"].Formula = rows[i]["id"];
                        xlSht.Cells[4, "B"].Formula = rows[i]["name"];
                        xlSht.Cells[5, "B"].Formula = rows[i]["birthdate"];
                        xlSht.Cells[6, "B"].Formula = rows[i]["phonenumber"];
                        xlSht.Cells[7, "B"].Formula = rows[i]["address"];
                        xlSht.Cells[8, "B"].Formula = rows[i]["socialnumber"];

                        xlSht.Cells[4, "D"].Formula = rows[i]["id"];
                        xlSht.Cells[4, "E"].Formula = rows[i]["name"];
                        xlSht.Cells[4, "F"].Formula = rows[i]["birthdate"];
                        xlSht.Cells[4, "G"].Formula = rows[i]["phonenumber"];
                        xlSht.Cells[4, "H"].Formula = rows[i]["address"];
                        xlSht.Cells[4, "I"].Formula = rows[i]["socialnumber"];
                    }
                }
                xlWb.Close(true);
                xlApp.Quit();
                System.Diagnostics.Process.Start(file);
            }
            catch (IOException iox)
            {
                Console.WriteLine(iox.Message);
            }
            

        }
    }
}

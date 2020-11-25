using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Office2010.Word;

namespace WordToExcel
{
    public partial class Form1 : Form
    {
       // private object openFileDialog1;

        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        // To browse and select files on a computer in an application  
        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = this.openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Open word document for read  
            using (var doc = WordprocessingDocument.Open(textBox1.Text.Trim(), false))
            {
                // To create a temporary table   
                DataTable dt = new DataTable();
                int rowCount = 0;

                // Find the first table in the document.   
                Table table = doc.MainDocumentPart.Document.Body.Elements<Table>().First();

                // To get all rows from table  
                IEnumerable<TableRow> rows = table.Elements<TableRow>();

                // To read data from rows and to add records to the temporary table  
                foreach (TableRow row in rows)
                {
                    if (rowCount == 0)
                    {
                        foreach (TableCell cell in row.Descendants<TableCell>())
                        {
                            dt.Columns.Add(cell.InnerText);
                        }
                        rowCount += 1;
                    }
                    else
                    {
                        dt.Rows.Add();
                        int i = 0;
                        foreach (TableCell cell in row.Descendants<TableCell>())
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.InnerText;
                            i++;
                        }
                    }
                }

                // To display the result   
                // Bind datatable(temporary table) to the datagridview   
                dataGridView1.DataSource = dt;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(dataGridView1.Rows.Count > 0)
            {

                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

                excelApp.Application.Workbooks.Add(Type.Missing);

                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    excelApp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView1.Rows.Count-1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        excelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }

                excelApp.Columns.AutoFit();
                excelApp.Visible = true;

            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
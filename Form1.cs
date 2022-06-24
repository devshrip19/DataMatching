using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;


namespace DataMatching
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
              
                String fileToOpen = openFileDialog1.FileName;
                String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                fileToOpen+
                                ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                OleDbConnection con = new OleDbConnection(constr);
                OleDbCommand oconn = new OleDbCommand("Select * From [Sheet1$]", con);
                con.Open();

                OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                DataTable data = new DataTable();
                sda.Fill(data);
                dataGridView1.DataSource = data;
                con.Close();
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    String fileToOpen = openFileDialog1.FileName;
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    fileToOpen +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select * From [Sheet1$]", con);
                    con.Open();

                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable data = new DataTable();
                    sda.Fill(data);
                    dataGridView2.DataSource = data;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }
        private void compareDatagridviews()
        {
            DataTable src1 = GetDataTableFromDGV(dataGridView1);
            DataTable src2 = GetDataTableFromDGV(dataGridView2);

            for (int i = 0; i < src1.Rows.Count; i++)
            {
                var row1 = src1.Rows[i].ItemArray;
                var row2 = src2.Rows[i].ItemArray;

                for (int j = 0; j < row1.Length; j++)
                {
                    if (!row1[j].ToString().Equals(row2[j].ToString()))
                    {
                        dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.Red;
                        dataGridView2.Rows[i].Cells[j].Style.BackColor = Color.Red;
                        MessageBox.Show(dataGridView1.Rows[i].Cells[j].Value.ToString()+"!="+ dataGridView2.Rows[i].Cells[j].Value.ToString());
             
                    }
                }
            }
        }
        private DataTable GetDataTableFromDGV(DataGridView dgv)
        {
            var dt = new DataTable();
            foreach (DataGridViewColumn column in dgv.Columns)
            {
                if (column.Visible)
                {
                    // You could potentially name the column based on the DGV column name (beware of dupes)
                    // or assign a type based on the data type of the data bound to this DGV column.
                    dt.Columns.Add();
                }
            }

            object[] cellValues = new object[dgv.Columns.Count];
            foreach (DataGridViewRow row in dgv.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    cellValues[i] = row.Cells[i].Value;
                }
                dt.Rows.Add(cellValues);
            }

            return dt;
        }

        

        private void button3_Click(object sender, EventArgs e)
        {
            compareDatagridviews();
        }
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace AccessDB
{
    public partial class Form4 : Form
    {
        static string path = Directory.GetCurrentDirectory() + @"\ОА";
        OleDbConnection dbCon = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + @"\ОА.accdb");

        public Form4()
        {
            InitializeComponent();
            textBox1.Enabled = false;
            textBox1.TextAlign = HorizontalAlignment.Center;
        }

        private void Form4_Shown(object sender, EventArgs e)
        {
            dbCon.Open();

            checkedListBox1.CheckOnClick = true;

            if (textBox1.Text != "")
            {
                OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT * FROM [" + textBox1.Text + "]", dbCon);
                DataTable dataTable = new DataTable();
                dbAdapter1.Fill(dataTable);
            }
            else
            {
                MessageBox.Show("Не выбрана таблица для запроса");
            }

            if (textBox1.Text != "")
            {
                OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT DISTINCT [НД в области стандартизации] FROM [" + textBox1.Text + "]", dbCon);
                DataTable dataTable = new DataTable();
                dbAdapter1.Fill(dataTable);
                checkedListBox1.DataSource = dataTable;
                //checkedListBox1.DisplayMember = "НД_в_области_стандартизации";
                checkedListBox1.ValueMember = "НД в области стандартизации";
                //checkedListBox1.ValueMember = "Код";
            }
            else
            {
                MessageBox.Show("Не выбрана таблица для запроса");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int i = 0;
            Close();
            textBox2.Clear();
            foreach (object item in checkedListBox1.CheckedItems)
            {
                string curItemString = ((DataRowView)item)[checkedListBox1.DisplayMember].ToString();
                // выполняем действия со строкой
                textBox2.Text = textBox2.Text + curItemString + ", ";
                //Data.curItemStringList[i] = curItemString;
                i++;
            }
        }
    }
}

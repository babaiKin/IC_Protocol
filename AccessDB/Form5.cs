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
    public partial class Form5 : Form
    {
        static string path = Directory.GetCurrentDirectory() + @"\ОА";
        OleDbConnection dbCon = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + @"\Сотрудники.accdb");

        public Form5()
        {
            InitializeComponent();
        }

        private void Form5_Shown(object sender, EventArgs e)
        {
            dbCon.Open();

            checkedListBox1.CheckOnClick = true;

            OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT DISTINCT [Фамилия ИО] FROM [Сотрудники]", dbCon);
            DataTable dataTable = new DataTable();
            dbAdapter1.Fill(dataTable);
            checkedListBox1.DataSource = dataTable;
            checkedListBox1.ValueMember = "Фамилия ИО";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            My.poveritelListLength = checkedListBox1.CheckedItems.Count; 
            int i = 0;
            Close();
            textBox2.Clear();

            foreach (object item in checkedListBox1.CheckedItems)
            {
                string curItemString = ((DataRowView)item)[checkedListBox1.DisplayMember].ToString();
                //выполняем действия со строкой
                //textBox2.Text = curItemString;

                //цикл вставки должности
                string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + @"\Сотрудники.accdb";
                string queryString = "SELECT * FROM [Сотрудники] WHERE [Фамилия ИО] = '" + curItemString + "'";

                try
                {
                    using (OleDbConnection connection = new OleDbConnection(connectionString))
                    {
                        OleDbCommand command = new OleDbCommand(queryString, connection);
                        connection.Open();
                        OleDbDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            //string Word1 = reader.GetValue(2).ToString();
                            //textBox1.Text = Word1;
                            textBox1.Text = reader.GetValue(2).ToString();
                        }
                        reader.Close();
                        connection.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("failed to connect");
                }
                
                /*
                OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT [Должность] FROM [Сотрудники] WHERE [Фамилия ИО] =" + textBox2.Text, dbCon);
                DataTable dataTable = new DataTable();
                dbAdapter1.Fill(dataTable);
                checkedListBox1.DataSource = dataTable;
                checkedListBox1.ValueMember = "Должность";
                */
                //конец цикла вставки должности

                My.poveritelList[i] = textBox1.Text + " _______________________________ " + curItemString /*textBox2.Text*/;
                i++;
            }

            dbCon.Close();
        }

        private void Form5_Load(object sender, EventArgs e)
        {
            
        }

        private void Form5_FormClosing(object sender, FormClosingEventArgs e)
        {
            dbCon.Close();
        }
    }
}

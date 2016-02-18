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
    public partial class Form6 : Form
    {
        static string path = Directory.GetCurrentDirectory() + @"\ОА";
        OleDbConnection dbCon = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + @"\Оборудование.accdb");

        public Form6()
        {
            InitializeComponent();
        }

        private void Form6_Shown(object sender, EventArgs e)
        {
            dbCon.Open();
            dataGridView1.RowHeadersVisible = false;

            OleDbDataAdapter dbAdapter2 = new OleDbDataAdapter(@"SELECT * FROM [Оборудование]", dbCon);
            DataTable dataTable = new DataTable();
            dbAdapter2.Fill(dataTable);
            dataGridView1.DataSource = dataTable;


            //видимость некоторых колонок и запрет на ввод во все, кроме "результатов"
            this.dataGridView1.Columns[0].Visible = false;
            this.dataGridView1.Columns[1].Visible = false;
            this.dataGridView1.Columns[2].Visible = true;
            this.dataGridView1.Columns[3].Visible = false;
            this.dataGridView1.Columns[4].Visible = true;
            this.dataGridView1.Columns[5].Visible = false;
            this.dataGridView1.Columns[6].Visible = false;
            this.dataGridView1.Columns[7].Visible = false;
            this.dataGridView1.Columns[8].Visible = false;
            this.dataGridView1.Columns[9].Visible = false;
            this.dataGridView1.Columns[10].Visible = false;
            this.dataGridView1.Columns[11].Visible = false;
            this.dataGridView1.Columns[12].Visible = false;


            dbCon.Close();
            /*
            dbCon.Open();

            checkedListBox1.CheckOnClick = true;

            OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT [Заводской], [Наименование] FROM [Оборудование]", dbCon);
            DataTable dataTable = new DataTable();
            dbAdapter1.Fill(dataTable);
            checkedListBox1.DataSource = dataTable;
            checkedListBox1.DisplayMember = "Наименование";
            checkedListBox1.ValueMember = "Заводской";*/
        }

        private void Form6_FormClosing(object sender, FormClosingEventArgs e)
        {
            dbCon.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int selectedCellCount = dataGridView1.GetCellCount(DataGridViewElementStates.Selected);

            if (selectedCellCount > 0)
            {
                //отсчет в массиве начинается с 0, а в dgv - с 1 
                string[,] stringArray = new string[dataGridView1.SelectedRows.Count, 2];
                for (int n = 0; n < dataGridView1.SelectedRows.Count; n++)
                {
                    //MessageBox.Show(dataGridView1[1, dataGridView1.SelectedRows[i].Index].Value.ToString() + " || " + dataGridView1[4, dataGridView1.SelectedRows[i].Index].Value.ToString());
                    stringArray[n, 0] = dataGridView1[2, dataGridView1.SelectedRows[n].Index].Value.ToString();
                    stringArray[n, 1] = dataGridView1[4, dataGridView1.SelectedRows[n].Index].Value.ToString();
                }

                //MessageBox.Show(stringArray[0, 0] + " || " + stringArray[0,1]);

                for (int n = dataGridView1.Rows.Count - 1; n >= 0; n--)
                {
                    if (dataGridView1[1, n].Selected)
                    {
                        //i--;
                    }
                    else
                        dataGridView1.Rows.Remove(dataGridView1.Rows[n]);
                }





                My.oborudovanieListLength = dataGridView1.SelectedRows.Count;
                int i = 0;
                Close();
                textBox2.Clear();

                foreach (object item in dataGridView1.SelectedRows)
                {
                    string curItemString = ((DataRowView)item)[checkedListBox1.DisplayMember].ToString();
                    //выполняем действия со строкой
                    //textBox2.Text = curItemString;

                    //цикл вставки должности
                    string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + @"\Оборудование.accdb";
                    string queryString = "SELECT * FROM [Оборудование] WHERE [Наименование] = '" + curItemString + "'";

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


                    OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT [Должность] FROM [Сотрудники] WHERE [Фамилия ИО] =" + textBox2.Text, dbCon);
                    DataTable dataTable = new DataTable();
                    dbAdapter1.Fill(dataTable);
                    checkedListBox1.DataSource = dataTable;
                    checkedListBox1.ValueMember = "Должность";

                    //конец цикла вставки должности

                    //My.poveritelList[i] = textBox1.Text + " _______________________________ " + curItemString /*textBox2.Text;
                    i++;

                }

                dbCon.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //удаление всех невыделенных ячеек
            int selectedCellCount = dataGridView1.GetCellCount(DataGridViewElementStates.Selected);
            textBox3.Text = "";
            if (selectedCellCount > 0)
            {
                My.oborudovanieListLength = dataGridView1.SelectedRows.Count;
                //отсчет в массиве начинается с 0, а в dgv - с 1 
                string[,] stringArray = new string[dataGridView1.SelectedRows.Count, 2];
                for (int n = 0; n < dataGridView1.SelectedRows.Count; n++)
                {
                    //MessageBox.Show(dataGridView1[1, dataGridView1.SelectedRows[i].Index].Value.ToString() + " || " + dataGridView1[4, dataGridView1.SelectedRows[i].Index].Value.ToString());
                    stringArray[n, 0] = dataGridView1[2, dataGridView1.SelectedRows[n].Index].Value.ToString();
                    stringArray[n, 1] = dataGridView1[4, dataGridView1.SelectedRows[n].Index].Value.ToString();
                    //textBox3.Text = textBox3.Text + "\r- " + dataGridView1[2, dataGridView1.SelectedRows[n].Index].Value.ToString() + ", " + dataGridView1[4, dataGridView1.SelectedRows[n].Index].Value.ToString() + ", " + dataGridView1[11, dataGridView1.SelectedRows[n].Index].Value.ToString();
                    //textBox3.AppendText("\r- " + dataGridView1[2, dataGridView1.SelectedRows[n].Index].Value.ToString() + ", " + dataGridView1[4, dataGridView1.SelectedRows[n].Index].Value.ToString() + ", " + dataGridView1[11, dataGridView1.SelectedRows[n].Index].Value.ToString());

                    //MessageBox.Show(dataGridView1[2, dataGridView1.SelectedRows[n].Index].Value.ToString() + " || " + dataGridView1[4, dataGridView1.SelectedRows[n].Index].Value.ToString());
                    My.oborudovanieList[n] = dataGridView1[2, dataGridView1.SelectedRows[n].Index].Value.ToString() + ", " + dataGridView1[4, dataGridView1.SelectedRows[n].Index].Value.ToString() + ", " + dataGridView1[11, dataGridView1.SelectedRows[n].Index].Value.ToString();
                    //MessageBox.Show("" + My.oborudovanieList[n]);
                }

                //MessageBox.Show(stringArray[0, 0] + " || " + stringArray[0,1]);


                for (int n = dataGridView1.Rows.Count - 1; n >= 0; n--)
                {
                    if (dataGridView1[1, n].Selected)
                    {
                        //i--;
                    }
                    else
                        dataGridView1.Rows.Remove(dataGridView1.Rows[n]);
                }
            }

            //вывод всех оставшихся ячеек в tmpl
            //My.oborudovanie = textBox3.Text;

        }
    }
}
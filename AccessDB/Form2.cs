using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace AccessDB
{
    public partial class Form2 : Form
    {
        OleDbConnection dbCon = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Directory.GetCurrentDirectory() + @"\ОА\Сотрудники.accdb");

        public Form2()
        {
            InitializeComponent();
            comboBox1.Items.Clear();
            dbCon.Open();
            OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT * FROM Сотрудники ORDER BY [Фамилия ИО] ASC", dbCon);
            DataTable dataTable = new DataTable();
            dbAdapter1.Fill(dataTable);
            comboBox1.DataSource = dataTable;
            comboBox1.DisplayMember = "Фамилия ИО";
            comboBox1.ValueMember = "Код";
            textBox1.Text = "0000"; //временный ввод пароля, чтобы не задалбливало каждый раз вводить при тесте
            dbCon.Close();
        }

        //переход дальше по кнопке
        public void button1_Click(object sender, EventArgs e)
        {
            dbCon.Open();
            OleDbCommand countSQL = new OleDbCommand(@"SELECT COUNT(*) FROM Сотрудники WHERE (((Сотрудники.[Фамилия ИО])= '" + (comboBox1.SelectedItem as DataRowView).Row["Фамилия ИО"].ToString() + "') AND ((Сотрудники.[Пароль])= '" + textBox1.Text + "'))", dbCon);
            countSQL.ExecuteNonQuery();
            int count = (int)countSQL.ExecuteScalar();
            if (count == 1)
            {
                Hide();
                Form3 form3 = new Form3();
                form3.Show();
                form3.sotrudnik = (comboBox1.SelectedItem as DataRowView).Row["Фамилия ИО"].ToString();
                form3.textBox15.Text = (comboBox1.SelectedItem as DataRowView).Row["Фамилия ИО"].ToString();
                label111.Text = (comboBox1.SelectedItem as DataRowView).Row["Должность"].ToString();
                form3.textBox14.Text = this.label111.Text;
            }
            else
            {
                MessageBox.Show("failed to LogIn :: неверный пароль");
            }
            dbCon.Close();
        }

        //переход дальше по enter'у
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) // Проверям нажата ли именно клавиша Enter
            {
                dbCon.Open();
                OleDbCommand countSQL = new OleDbCommand(@"SELECT COUNT(*) FROM Сотрудники WHERE (((Сотрудники.[Фамилия ИО])= '" + (comboBox1.SelectedItem as DataRowView).Row["Фамилия ИО"].ToString() + "') AND ((Сотрудники.[Пароль])= '" + textBox1.Text + "'))", dbCon);
                countSQL.ExecuteNonQuery();
                int count = (int)countSQL.ExecuteScalar();
                if (count == 1)
                {
                    Hide();
                    Form3 form3 = new Form3();
                    form3.Show();
                    form3.sotrudnik = (comboBox1.SelectedItem as DataRowView).Row["Фамилия ИО"].ToString();
                    label111.Text = (comboBox1.SelectedItem as DataRowView).Row["Должность"].ToString();
                    form3.textBox14.Text = this.label111.Text;
                }
                else
                {
                    MessageBox.Show("failed to LogIn :: неверный пароль");
                }
                dbCon.Close();
            }
        }

        //кнопка выхода
        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
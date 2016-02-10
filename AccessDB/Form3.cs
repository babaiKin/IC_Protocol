using System;
using System.IO;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Linq;

using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Threading.Tasks;

namespace AccessDB
{
    public partial class Form3 : Form
    {
        static string path = Directory.GetCurrentDirectory() + @"\ОА";
        OleDbConnection dbCon = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + @"\ОА.accdb");
        OleDbConnection myConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + @"\Заказчики.accdb");
        OleDbConnection sotrCon = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + @"\Сотрудники.accdb");
        OleDbDataAdapter myDataAdapter = new OleDbDataAdapter();
        DataSet myDataSet = new DataSet();
        BindingSource myBS = new BindingSource();

        public string sotrudnik
        {
            get
            {
                return label15.Text;
            }
            set
            {
                label15.Text = value;
            }
        }


        private void Form3_Load(object sender, EventArgs e)
        {
            //button1.Enabled = false;
            dateTimePicker1.CustomFormat = "dd MMMM yyyy   HH:mm";
            dateTimePicker1.Format = DateTimePickerFormat.Custom;

            myConn.Open();
            //toolStripStatusLabel1.Text = "Статус подключения: " + myConn.State.ToString();
            myDataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM Заказчики", myConn);
            myDataAdapter.Fill(myDataSet.Tables["Заказчики"]);

            myBS.DataMember = "Заказчики";
            myBS.DataSource = myDataSet;

            comboBox1.DataSource = myBS;
            comboBox1.DisplayMember = "Заказчик";

            textBox3.DataBindings.Add(new Binding("Text", myBS, "Адрес", true));
            myConn.Close();

            //берем название протокола
            //убираем все буковки
            /*
            string lastNumDir = Directory.GetCurrentDirectory() + @"\ОА\последний_номер.txt";
            string lastNum = System.IO.File.ReadAllText(lastNumDir);
            string lastNum = Convert.ToString(Convert.ToInt32(System.IO.File.ReadAllText(lastNumDir)) + 1);
            this.textBox1.Text = lastNum;

            string str = "";
            for (int i = 0; i < textBox1.Text.Length; i++)
            {
                if (char.IsDigit(textBox1.Text[i]))
                {
                    str += textBox1.Text[i];
                }
            }
            textBox1.Text = Convert.ToString(Convert.ToInt32(str) + 1);
            */

            //поиск последнего номера протокола
           /* string[] dirs = Directory.GetFiles(Directory.GetCurrentDirectory() + @"\ОА\Протоколы", "*");
            if (dirs.Count() != 0)
            {
                try
                {
                    //this.textBox1.Text = Convert.ToString(Convert.ToInt32(System.IO.Path.GetFileNameWithoutExtension(dirs[dirs.Length - 1])) + 1);
                    this.textBox1.Text = Convert.ToInt32(System.IO.Path.GetFileNameWithoutExtension(dirs[dirs.Length - 1])) + 1 + "";
                }
                catch
                {
                    MessageBox.Show("Обнаружено неверное название свидетельсва", "ошибка");
                }
            }

            else
                textBox1.Text = "файлы протоколов не найдены...";
            */
            comboBox2.Items.Clear();
            dbCon.Open();
            DataTable tbls = dbCon.GetSchema("Tables", new string[] { null, null, null, "TABLE" }); //список всех таблиц
            foreach (DataRow row in tbls.Rows)
            {
                string TableName = row["TABLE_NAME"].ToString();
                comboBox2.Items.Add(TableName);
            };
            dbCon.Close();
        }


        public Form3()
        {
            InitializeComponent();
            myDataSet.Tables.Add("Заказчики");
        }


        private void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(Data.curItemStringList[0]);
            if (comboBox2.Text == "")
            {
                MessageBox.Show("Поле 'Группа, согласно ОА' не заполнено!");
                return;
            }

            if (textBox1.Text == "")
            {
                MessageBox.Show("Поле 'Номер протокола' не заполнено!");
                return;
            }

            //if (comboBox2.Text == "" || textBox2.Text == "" || textBox6.Text == "" || textBox9.Text == "" || textBox13.Text == "")
            //{
            //    MessageBox.Show("Не все обязательные поля заполнены!");
            //    return;
            //}
            else
            {
                int count = 0;
                DirectoryInfo d = new DirectoryInfo(Directory.GetCurrentDirectory() + @"\ОА\Протоколы\");
                StringBuilder fileList = new StringBuilder();
                string pathExtension = "протокол №" + textBox1.Text + ".doc";

                foreach (FileInfo f in d.GetFiles(pathExtension))
                {
                    string u = Convert.ToString(fileList.AppendLine(f.Name));
                    count++;
                }

                if (count.ToString() == "0")
                {
                    Hide();
                    Form1 form1 = new Form1();
                    form1.Show();
                    //form1.sotrudniki = label15.Text;
                    form1.label1.Text = this.textBox1.Text;
                    //form1.label2.Text = this.textBox2.Text;
                    form1.label2.Text = this.comboBox1.Text;
                    form1.label3.Text = this.textBox3.Text;
                    //form1.label4.Text = this.textBox4.Text;
                    form1.label5.Text = this.textBox5.Text;
                    form1.label4.Text = this.textBox4.Text;
                    form1.label6.Text = this.textBox6.Text;
                    form1.label23.Text = this.comboBox2.Text;
                    form1.label7.Text = this.textBox7.Text;
                    form1.label8.Text = this.textBox8.Text;
                    form1.label9.Text = this.dateTimePicker1.Text;
                    //form1.label9.Text = this.textBox9.Text;
                    form1.label10.Text = this.dateTimePicker2.Text + " - " + this.dateTimePicker3.Text;
                    //form1.label10.Text = this.textBox10.Text;
                    form1.label11.Text = this.textBox11.Text;
                    form1.label12.Text = this.textBox12.Text;
                    //form1.label13.Text = this.textBox13.Text;  условия проведения испытаний
                    form1.label13.Text = "Температура: " + this.textBox13.Text + " °C," + " Влажность: " + this.textBox2.Text + " %," + " Давление: " + this.textBox9.Text + " мм. рт. ст.";
                    //form1.label18.Text = "\rТемпература: " + this.textBox13.Text + " °C";
                    //form1.label19.Text = "\rВлажность: " + this.textBox2.Text + " %";
                    //form1.label20.Text = "\rДавление: " + this.textBox9.Text + " мм. рт. ст.";
                    if (textBox16.Text != "")
                        form1.label13.Text = form1.label13.Text + ", Электромагнитное поле: " + this.textBox16.Text + " А/м";
                    //form1.label21.Text = "\rЭлектромагнитное поле: " + this.textBox16.Text + " А/м";
                    if (textBox10.Text != "")
                        form1.label13.Text = form1.label13.Text + ", Магнитное поле: " + this.textBox10.Text + " А/м";
                    //form1.label22.Text = "\rМагнитное поле: " + this.textBox10.Text + " А/м";

                    form1.label14.Text = this.textBox14.Text;
                    form1.label15.Text = this.label15.Text;
                    //form1.label16.Text = this.textBox15.Text;
                    //form1.label17.Text = this.textBox16.Text;
                    form1.textBox1.Text = this.comboBox2.Text;
                   
                }
                else
                    MessageBox.Show("Протокол с таким номером уже существует");
            }
        }

        //переход на сл вкладку
        //нужно допилить проверку на непустые значения!!
        private void textBox5_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                tabControl1.SelectTab(tabPage2);
            }
        }


        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;

            if (!Char.IsDigit(ch) && e.KeyChar != '-' && ch != 8) //Если символ, введенный с клавы - не цифра, не знак '-' , не backspace ,
            {
                e.Handled = true;                                 // то событие не обрабатывается. ch!=8 (8 - это Backspace)
            }
        }

        
        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.Value = dateTimePicker1.Value; 
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            //test valueChanged
        }

        private void button2_Click(object sender, EventArgs e)
        {
         //test button   
        }

        private void button3_Click(object sender, EventArgs e)
        {
            tabControl1.SelectTab(tabPage2);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            tabControl1.SelectTab(tabPage1);
        }

        private void tabControl1_MouseMove(object sender, MouseEventArgs e)
        {
            /*
            if (comboBox2.Text != "")
                button1.Enabled = true;
            if (comboBox2.Text == "")
                button1.Enabled = false;
            */
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form4 form4 = new Form4();
            form4.Show();
            form4.textBox1.Text = this.comboBox2.Text;
            form4.button1.Click += (senderSlave, eSlave) =>
            //this.textBox12.Text = form4.checkedListBox1.Text;
            this.textBox12.Text = form4.textBox2.Text;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            myConn.Open();
            //toolStripStatusLabel1.Text = "Статус подключения: " + myConn.State.ToString();

            OleDbCommand com = new OleDbCommand("INSERT INTO [Заказчики] ([Заказчик], [Адрес]) VALUES('" + comboBox1.Text + "','" + textBox3.Text + "')", myConn);
            com.ExecuteNonQuery();

            myConn.Close();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //Form5 form5 = new Form5();
            //form5.Show();
            //form5.textBox1.Text = this.comboBox2.Text;
            //form5.button1.Click += (senderSlave, eSlave) =>
            //this.textBox17.Text = form5.checkedListBox1.Text;
            //this.textBox17.Text = form5.textBox2.Text;
        }
    }
}


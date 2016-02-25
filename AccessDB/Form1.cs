using System;
using System.IO;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
//using System.Linq;
//using System.Threading.Tasks;
//using Office = Microsoft.Office.Core;
//using System.Threading;
//using System.Management;
//using System.Reflection;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Drawing;
//using System.Text;

namespace AccessDB
{
    public partial class Form1 : Form
    {
        bool exit;
        static string path = Directory.GetCurrentDirectory() + @"\ОА";
        OleDbConnection dbCon = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + @"\ОА.accdb");

        public Form1()
        {
            InitializeComponent();
        }

        
        private void Form1_Load(object sender, EventArgs e)
        {
            exit = true;
        }
        
        //выбор таблицы//НЕ ИСПОЛЬЗУЕТСЯ!!!!
        private void button1_Click(object sender, EventArgs e)
        {
                //OleDbConnection dbCon = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openFileDialog1.FileName);
                //comboBox1.Items.Clear();
                dbCon.Open();
                DataTable tbls = dbCon.GetSchema("Tables", new string[] { null, null, null, "TABLE" }); //список всех таблиц
                foreach (DataRow row in tbls.Rows)
                {
                    string TableName = row["TABLE_NAME"].ToString();
                    comboBox1.Items.Add(TableName);
                };
                //dbCon.Close();
            }

        //запись выбранной таблицы в грид//НЕ ИСПОЛЬЗУЕТСЯ!!!!
        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex != -1)
            {
               //OleDbConnection dbCon = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openFileDialog1.FileName);
               //dbCon.Open();
               OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT * FROM " + comboBox1.SelectedItem , dbCon);
               DataTable dataTable = new DataTable();
               dbAdapter1.Fill(dataTable);
               dataGridView1.DataSource = dataTable;
               //dbCon.Close();
            }
            else
            {
                MessageBox.Show("Не выбрана таблица для запроса");
            }
            foreach (int indexChecked in checkedListBox1.CheckedIndices)
            {
                checkedListBox1.SetItemChecked(indexChecked, false);
            }

            //заполнение чекбокса
            checkedListBox1.CheckOnClick = true;
            if (comboBox1.SelectedIndex != -1)
            {
                //OleDbConnection dbCon = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openFileDialog1.FileName);
                //dbCon.Open();
                OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT * FROM " + comboBox1.SelectedItem, dbCon);
                DataTable dataTable = new DataTable();
                dbAdapter1.Fill(dataTable);
                checkedListBox1.DataSource = dataTable;
                checkedListBox1.DisplayMember = "Наименование";
                checkedListBox1.ValueMember = "Код";
                //dbCon.Close();
            }
            else
            {
                MessageBox.Show("Не выбрана таблица для запроса");
            }
        }

        //заполнение чекбокса данными из таблицы//НЕ ИСПОЛЬЗУЕТСЯ!!!!
        private void button3_Click(object sender, EventArgs e)
        {
            //checkedListBox1.Items.Clear();
            foreach (int indexChecked in checkedListBox1.CheckedIndices)
            {
                checkedListBox1.SetItemChecked(indexChecked, false);
            }
            
            checkedListBox1.CheckOnClick = true;
            if (comboBox1.SelectedIndex != -1)
            {
                //OleDbConnection dbCon = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openFileDialog1.FileName);
                //dbCon.Open();
                OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT * FROM [" + comboBox1.SelectedItem + "]", dbCon);
                DataTable dataTable = new DataTable();
                dbAdapter1.Fill(dataTable);
                checkedListBox1.DataSource = dataTable;
                checkedListBox1.DisplayMember = "наименование";
                checkedListBox1.ValueMember = "Код";
                //dbCon.Close();
            }
            else
            {
                MessageBox.Show("Не выбрана таблица для запроса");
            }
        }

        //сортировка по "галочным" ГОСТам//НЕ ИСПОЛЬЗУЕТСЯ!!!!
        private void button4_Click(object sender, EventArgs e)
        {
            checkedListBox1.CheckOnClick = true;
            if (textBox1.Text != "")                  //проверка выбрана ли таблица
            {
                if (checkedListBox1.Items.Count > 0)            //проверка открыта ли таблица и перенесены ли данные в чекбокс
                {
                    if (checkedListBox1.CheckedItems.Count > 0) //проверка выбраны ли данные
                    {
                        //заполнение "галочными" данными//не корректно работает при выборке по показателю
                        
                        foreach (object item in checkedListBox1.CheckedItems)
                        {
                            string curItemString = ((DataRowView)item)[checkedListBox1.DisplayMember].ToString();
                            // выполняем действия со строкой
                            DataTable dataTable = new DataTable();
                            OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT * FROM [" + textBox1.Text + "] WHERE [НД в области стандартизации] = '" + curItemString + "'", dbCon);
                            dbAdapter1.Fill(dataTable);
                            dataGridView1.DataSource = dataTable;

                            //заполнение второго чекбокса (наименование показателя)
                            checkedListBox2.DataSource = dataTable;
                            checkedListBox2.DisplayMember = "Наименование_показателя_единицы_измерений";
                            //checkedListBox2.ValueMember = "Код";
                            checkedListBox2.ValueMember = "Наименование_показателя_единицы_измерений";
                        }
                    }
                    else
                    {
                        MessageBox.Show("Не выбраны данные для добавления");
                    }
                }
                else
                {
                    MessageBox.Show("Не открыта таблица");
                }
            }
            else
            {
                MessageBox.Show("Не выбрана таблица для запроса");
            }
        }

       
        //экспорт в шаблон Word
        private void button6_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("" + My.oborudovanie.ToString() );
            //if (textBox2.Text == "")
            //{
            //    MessageBox.Show("Не выбран(ы) испытатель(ли)!");
            //    return;
            //}

            //MessageBox.Show("" + My.poveritelListLength);
            //dbCon.Close();
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();
            object missing = Type.Missing;

            //Объявляем новый экземпляр класса Stopwatch
            //запускаем
            Stopwatch testStopwatch = new Stopwatch();
            testStopwatch.Start();

            if (dataGridView1.Rows.Count != 0)
            {
                object fileName = Directory.GetCurrentDirectory() + @"\ОА\tmpl.doc";
                object falseValue = false;
                object trueValue = true;
                
                doc = app.Documents.Open(ref fileName, ref missing, ref falseValue,
                                    ref missing, ref missing, ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing);
                //app.Visible = true;

                //exception();
                //string[] txt = { label1.Text, DateTime.Now.ToLongDateString(), label2.Text, label3.Text, /*label4.Text,*/ label5.Text, label4.Text, label6.Text, label23.Text, label7.Text, label8.Text, label9.Text, label10.Text, label11.Text, label12.Text, /*label13.Text*/ label18.Text, label19.Text, label20.Text, label21.Text, label22.Text, label14.Text, label15.Text };
                string[] txt = { label1.Text, label26.Text, /*DateTime.Now.ToLongDateString(),*/ label2.Text, label3.Text, /*label4.Text,*/ label5.Text, label4.Text, label6.Text, label23.Text, label7.Text, label8.Text, label9.Text, label10.Text, label11.Text, label12.Text, label13.Text /*, label18.Text, label19.Text, label20.Text, label21.Text, label22.Text*/ /*, label14.Text, label15.Text*/ };
                string[] FindObj = { "$num$", "$date$", /*"$date$",*/ "$zakazchik$", "$adress_zakazchika$", /*"$postavshik$",*/ "$name_izgotov$", "$adress_izgotov$", "$product_name$", "$product_group$", "$name$", "$proba$", "$date_time_postuplenia$", "$date_exe$", "$osnovanie$", "$nd$", "$conditions$" /*, "$temperature$", "$vlazhnost$","$davlenie$", "$elmagpole$", "$magpole$"*/ /*, "$dolznost$", "$sotrudnik$"*/ };

                int n = 0;
                while (n < FindObj.Length)
                {
                    //Очищаем параметры поиска
                    app.Selection.Find.ClearFormatting();
                    app.Selection.Find.Replacement.ClearFormatting();

                    //Задаём параметры замены и выполняем замену.
                    object findTextNUM = FindObj[n];
                    object replaceWithNUM = txt[n];
                    object replaceNUM = 2;

                    app.Selection.Find.Execute(ref findTextNUM, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceWithNUM, ref replaceNUM, ref missing, ref missing, ref missing, ref missing);
                    n++;
                }
                n = 0;


                //string poveritelListFind = "$poveritelList$";
                for (int nn = 0;  nn < My.poveritelListLength; nn++)
                {
                    if (My.poveritelList[nn] == "")
                        nn++;
                    else
                        textBox2.AppendText(My.poveritelList[nn] + "\r\r");
                }

                //Очищаем параметры поиска
                app.Selection.Find.ClearFormatting();
                app.Selection.Find.Replacement.ClearFormatting();

                //Задаём параметры замены и выполняем замену.
                object findTextNUM2 = "$poveritelList$";
                object replaceWithNUM2 = textBox2.Text;
                object replaceNUM2 = 2;

                app.Selection.Find.Execute(ref findTextNUM2, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref replaceWithNUM2, ref replaceNUM2, ref missing, ref missing, ref missing, ref missing);


                //string oborudovanieListFind = "$oborudovanie$";
                for (int nn = 0; nn < My.oborudovanieListLength; nn++)
                {
                    //textBox2.AppendText(My.poveritelList[nn] + "\r\r");
                    //Очищаем параметры поиска
                    app.Selection.Find.ClearFormatting();
                    app.Selection.Find.Replacement.ClearFormatting();

                    //Задаём параметры замены и выполняем замену.
                    object findTextNUM3 = "$oborudovanie$";
                    object replaceWithNUM3 = "- " + My.oborudovanieList[nn] + ";\r$oborudovanie$";
                    object replaceNUM3 = 2;

                    app.Selection.Find.Execute(ref findTextNUM3, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceWithNUM3, ref replaceNUM3, ref missing, ref missing, ref missing, ref missing);
                }

                //Задаём параметры замены и выполняем замену.
                object findTextNUM4 = "$oborudovanie$";
                object replaceWithNUM4 = "";
                object replaceNUM4 = 2;

                app.Selection.Find.Execute(ref findTextNUM4, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref replaceWithNUM4, ref replaceNUM4, ref missing, ref missing, ref missing, ref missing);

                Microsoft.Office.Interop.Word.Table table = doc.Tables[2];

                int i;

                for (i = 1; i < dataGridView1.Rows.Count + 1; i++)
                {
                    table.Rows.Add();
                    table.Cell(i + 2, 1).Range.Text = i.ToString();          //нумерация

                    for (int j = 1; j < dataGridView1.ColumnCount - 1; j++)  //цикл вставки dGV
                    {
                        //table.Cell(i + 2, j + 1).Range.Text = dataGridView1.Rows[i - 1].Cells[j].Value.ToString();
                        table.Cell(i + 3, j+1).Range.Text = dataGridView1.Rows[i - 1].Cells[j].Value.ToString();
                    }
                }

                //количество страниц/листов
                Word.WdStatistic stat = Word.WdStatistic.wdStatisticPages;
                double x = doc.ComputeStatistics(stat, ref missing); //страницы
                string y = Math.Ceiling(x / 2).ToString("G17");      //листы
                //label16.Text = x.ToString("G17");

                //поиск и вставка x&y
                string[] xy = { x.ToString("G17"), y };
                string[] Findxy = { "$x$", "$y$" };

                while (n < Findxy.Length)
                {
                    //Очищаем параметры поиска
                    app.Selection.Find.ClearFormatting();
                    app.Selection.Find.Replacement.ClearFormatting();

                    //Задаём параметры замены и выполняем замену.
                    object findTextNUM = Findxy[n];
                    object replaceWithNUM = xy[n];
                    object replaceNUM = 2;

                    app.Selection.Find.Execute(ref findTextNUM, ref missing, ref missing, ref missing,
                                               ref missing, ref missing, ref missing, ref missing, ref missing,
                                               ref replaceWithNUM, ref replaceNUM, ref missing, ref missing, ref missing, ref missing);
                    n++;
                }
                n = 0;

                DialogResult res = MessageBox.Show("Экспорт завершен. При нажатии ДА будет открыт сгенерированный файл, при нажатии НЕТ произойдет автоматическое сохранение файла и его открытие.", "Экспорт в Excel", MessageBoxButtons.YesNoCancel);

                if (res == DialogResult.Yes)    //открытие сгенерированного файла
                {
                    try
                    {
                        //Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\ОА\tmpl.doc";
                        //object newfileName = Directory.GetCurrentDirectory() + @"\ОА\протокол №" + label1.Text + ".doc";
                        app.Visible = true;

                        //doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
                        //                         ref missing, ref missing, ref missing, ref missing, ref missing,
                        //                         ref missing, ref missing, ref missing, ref missing, ref missing,
                        //                         ref missing, ref missing, ref missing);
                        //// Закрываем родительскую форму
                        //Hide();
                        MessageBox.Show("Экспорт успешно завершен, протокол открыт. При необходимости, сохраните протокол.");

                        /*
                        string lastNumDir = Directory.GetCurrentDirectory() + @"\ОА\последний_номер.txt";
                        string lastNumUPD = label1.Text;
                        System.IO.File.WriteAllText(lastNumDir, lastNumUPD);
                        */
                    }
                    catch
                    {
                        MessageBox.Show("ooooops...! что-то пошло не так. Пожалуйста обратитесь в службу поддержки");
                    }
                }

                if (res == DialogResult.No)     //автоматическое сохранение файла 
                {
                    try
                    {
                        object Target = (Directory.GetCurrentDirectory() + @"\ОА\Протоколы\" + label1.Text + ".doc");  // куда сохранить
                        object format_ = Word.WdSaveFormat.wdFormatDocumentDefault;
                        //Сохранение файла
                        doc.SaveAs(ref Target, ref format_,
                                   ref missing, ref missing, ref missing,
                                   ref missing, ref missing, ref missing,
                                   ref missing, ref missing, ref missing,
                                   ref missing, ref missing, ref missing,
                                   ref missing, ref missing);

                        //object falseValue = false;
                        //object trueValue = true;

                        doc = app.Documents.Open(ref Target, ref missing, ref falseValue,
                                            ref missing, ref missing, ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing);
                        ///
                        ///колонтитулы
                        ///

                        //верний колонтитул
                        //первая страница
                        foreach (Word.Section section in app.ActiveDocument.Sections)
                        {
                            Object oMissing = System.Reflection.Missing.Value;
                            Microsoft.Office.Interop.Word.Selection s = app.Selection;

                            doc.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekFirstPageHeader;
                            s.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            doc.ActiveWindow.Selection.TypeText("");
                            doc.ActiveWindow.Selection.Fields.Add(s.Range);

                            doc.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument; //выход из колонтитула
                        }

                        //верхний колонтитул
                        //остальные страницы
                        foreach (Word.Section section in app.ActiveDocument.Sections)
                        {
                            Object oMissing = System.Reflection.Missing.Value;
                            Microsoft.Office.Interop.Word.Selection s = app.Selection;

                            doc.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekPrimaryHeader;
                            s.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            doc.ActiveWindow.Selection.TypeText("ФБУ «Нижегородский ЦСМ» ИЦ «НИЖЕГОРОДСИСПЫТАНИЯ»          Протокол №" + label1.Text + " от " + label26.Text);
                            doc.ActiveWindow.Selection.Fields.Add(s.Range);

                            doc.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument; //выход из колонтитула
                        }

                        //нижний колонтитул
                        //первая страница
                        foreach (Word.Section section in app.ActiveDocument.Sections)
                        {
                            //нижний колонтитул
                            Object oMissing = System.Reflection.Missing.Value;
                            Microsoft.Office.Interop.Word.Selection s = app.Selection;

                            // код для номеров страницы
                            doc.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekFirstPageFooter;
                            s.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;

                            doc.ActiveWindow.Selection.TypeText("страница ");
                            object CurrentPage = Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage; //текущая страница
                            doc.ActiveWindow.Selection.Fields.Add(s.Range, ref CurrentPage, ref oMissing, ref oMissing);

                            doc.ActiveWindow.Selection.TypeText(" из ");
                            object TotalPages = Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages; //всего страниц
                            doc.ActiveWindow.Selection.Fields.Add(s.Range, ref TotalPages, ref oMissing, ref oMissing);

                            doc.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument; //выход из колонтитула
                        }

                        //нижний колонтитул
                        //первая страница
                        foreach (Word.Section section in app.ActiveDocument.Sections)
                        {
                            //нижний колонтитул
                            Object oMissing = System.Reflection.Missing.Value;
                            Microsoft.Office.Interop.Word.Selection s = app.Selection;
                            
                            // код для номеров страницы
                            doc.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekPrimaryFooter;
                            s.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;

                            doc.ActiveWindow.Selection.TypeText("страница ");
                            object CurrentPage = Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage; //текущая страница
                            doc.ActiveWindow.Selection.Fields.Add(s.Range, ref CurrentPage, ref oMissing, ref oMissing);

                            doc.ActiveWindow.Selection.TypeText(" из ");
                            object TotalPages = Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages; //всего страниц
                            doc.ActiveWindow.Selection.Fields.Add(s.Range, ref TotalPages, ref oMissing, ref oMissing);

                            doc.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument; //выход из колонтитула
                        }

                        app.Visible = true;

                        MessageBox.Show("Экспорт успешно завершен, протокол сохранен под номером " + label1.Text);
                    }
                    catch
                    {
                        MessageBox.Show("ooooops...! что-то пошло не так. Пожалуйста обратитесь в службу поддержки");
                    }
                }

                if (res == DialogResult.Cancel) //отмена 
                {
                    MessageBox.Show("Сохранение результатов экспорта отменено");
                    ((Microsoft.Office.Interop.Word._Application)app).Quit(false, ref missing, ref missing);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app);
                }
                //Останавливаем
                testStopwatch.Stop();

                //Теперь можем смотреть время выполнения операции
                TimeSpan tSpan; tSpan = testStopwatch.Elapsed;
                //MessageBox.Show("Время выполнения операции - " + tSpan.ToString()); //время выполнения операции
            }
            else
            {
                MessageBox.Show("Не открыта таблица для запроса");
            }
            exit = false;



            Hide();
            Form3 form3 = new Form3();
            form3.Show();
        }

        //если поле пустое ,то и в протоколе ему делать нечего
        /*public void exception()
        {
            if (label5.Text != "")
                label5.Text = "Наименование изготовителя: " + label5.Text;
            else if (label5.Text == "")
                label5.Text = "";
        }*/

        //закрытие формы и соединения
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        //заполнение датагрида1 данными из выбранной таблицы(группы) //НЕ ИСПОЛЬЗУЕТСЯ!!!!
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ////if (comboBox1.SelectedIndex != -1)
            ////{
            ////    OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT * FROM " + comboBox1.SelectedItem, dbCon);

            ////    DataTable dataTable = new DataTable();
            ////    dbAdapter1.Fill(dataTable);
            ////    dataGridView1.DataSource = dataTable;
            ////    //dbCon.Close();
            ////}
            ////else
            ////{
            ////    MessageBox.Show("Не выбрана таблица для запроса");
            ////}
            ////foreach (int indexChecked in checkedListBox1.CheckedIndices)
            ////{
            ////    checkedListBox1.SetItemChecked(indexChecked, false);
            ////}

            //////вывод группы показателей//для второго комбобокса
            /////*OleDbDataAdapter dbAdapter2 = new OleDbDataAdapter(@"SELECT DISTINCT Группа_показателей FROM " + comboBox1.SelectedItem, dbCon);
            ////DataTable ds = new DataTable();
            ////dbAdapter2.Fill(ds);
            ////comboBox2.DataSource = ds;
            ////comboBox2.DisplayMember = "Группа_показателей";
            //////comboBox2.ValueMember = "Код";
            ////*/

            //////заполнение чекбокса
            ////checkedListBox1.CheckOnClick = true;
            ////if (comboBox1.SelectedIndex != -1)
            ////{
            ////    //OleDbConnection dbCon = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openFileDialog1.FileName);
            ////    //dbCon.Open();
            ////    OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT * FROM " + comboBox1.SelectedItem, dbCon);
            ////    DataTable dataTable = new DataTable();
            ////    dbAdapter1.Fill(dataTable);
            ////    checkedListBox1.DataSource = dataTable;
            ////    checkedListBox1.DisplayMember = "Наименование_показателя_единицы_измерений";
            ////    //checkedListBox1.ValueMember = "Код";
            ////    //dbCon.Close();
            ////}
            ////else
            ////{
            ////    MessageBox.Show("Не выбрана таблица для запроса");
            ////}
        }

        //хер знает что, но, в любом случае, //НЕ ИСПОЛЬЗУЕТСЯ!!!!
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT * FROM " + comboBox1.SelectedItem + " WHERE Группа_показателей = '" + comboBox2.Text + "'", dbCon);
            //DataTable dataTable = new DataTable();
            //dbAdapter1.Fill(dataTable);
            //dataGridView1.DataSource = dataTable;
        }

        //тест кнопка //НЕ ИСПОЛЬЗУЕТСЯ!!!!
        private void button5_Click(object sender, EventArgs e)
        {
            OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT DISTINCT Группа_показателей FROM " + comboBox1.SelectedItem, dbCon);
            DataTable ds = new DataTable();
            dbAdapter1.Fill(ds);
            comboBox2.DataSource = ds;
            comboBox2.DisplayMember = "Группа_показателей";
            //comboBox2.ValueMember = "Код";
        }


        private void button7_Click(object sender, EventArgs e)
        {
            checkedListBox1.Items.Clear();

            int selectedCellCount = dataGridView1.GetCellCount(DataGridViewElementStates.Selected);

            if (selectedCellCount > 0)
            {
                //отсчет в массиве начинается с 0, а в dgv - с 1 
                string[,] stringArray = new string[dataGridView1.SelectedRows.Count,2];
                for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
                {
                    //MessageBox.Show(dataGridView1[1, dataGridView1.SelectedRows[i].Index].Value.ToString() + " || " + dataGridView1[4, dataGridView1.SelectedRows[i].Index].Value.ToString());
                    stringArray[i, 0] = dataGridView1[1, dataGridView1.SelectedRows[i].Index].Value.ToString();
                    stringArray[i, 1] = dataGridView1[4, dataGridView1.SelectedRows[i].Index].Value.ToString();
                }

                //MessageBox.Show(stringArray[0, 0] + " || " + stringArray[0,1]);

                for (int i = dataGridView1.Rows.Count - 1; i >= 0; i--)
                {
                    if (dataGridView1[1, i].Selected)
                    {
                        //i--;
                    }
                    else
                        dataGridView1.Rows.Remove(dataGridView1.Rows[i]);
                }




                /*
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    //прихерачить условие
                    //ежели строка не выделена - удалить
                    if (dataGridView1.Rows[dataGridView1.Rows.IndexOf(row)].Selected != false)
                    {
                        MessageBox.Show("Строка " + dataGridView1.Rows[dataGridView1.Rows.IndexOf(row)].Index + " выделена!");
                        dataGridView1.Rows.Remove(row);
                    }
                    //dataGridView1.Rows.Remove(row);
                }*/



                //удаление выделенных строк
                //нужно инвертировать и сделать удаление не выделенных
                /*
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    dataGridView1.Rows.Remove(row);
                }*/

                /* переписано //НЕ ИСПОЛЬЗУЕТСЯ!!!
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                for (int i = 0; i < selectedCellCount; i++)
                {
                    sb.Append("Cell: ");
                    sb.Append(dataGridView1.SelectedCells[i].Value.ToString());
                    sb.Append(", Column: ");
                    sb.Append(dataGridView1.SelectedCells[i].ColumnIndex.ToString());
                    sb.Append(Environment.NewLine);

                    //checkedListBox1.Items.Add(dataGridView1.SelectedCells[i].Value.ToString());
                    //MessageBox.Show(dataGridView1.SelectedCells[i].ColumnIndex + " || " + dataGridView1.SelectedCells[i].RowIndex);

                    MessageBox.Show(dataGridView1[dataGridView1.SelectedCells[i].ColumnIndex - 3, dataGridView1.SelectedCells[i].RowIndex].Value.ToString() + "  " + dataGridView1[dataGridView1.SelectedCells[i].ColumnIndex , dataGridView1.SelectedCells[i].RowIndex].Value.ToString());

                    //Наименование_показателя_единицы_измерений
                    //НД_на_метод_испытаний
                    ///DataTable dataTable = new DataTable();
                    ///OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT * FROM " + textBox1.Text + " WHERE ( Наименование_показателя_единицы_измерений = '" + dataGridView1[dataGridView1.SelectedCells[i].ColumnIndex - 3, dataGridView1.SelectedCells[i].RowIndex].Value.ToString() + "' AND НД_на_метод_испытаний = '" + dataGridView1[dataGridView1.SelectedCells[i].ColumnIndex, dataGridView1.SelectedCells[i].RowIndex].Value.ToString() + "')", dbCon);
                    ///dbAdapter1.Fill(dataTable);
                    ///dataGridView1.DataSource = dataTable;
                }
                sb.Append("Total: " + selectedCellCount.ToString());                
                */

                //MessageBox.Show(sb.ToString(), "Selected Rows");
            }


        }

        //заполнение чекбокса в зависимости от комбобокса //НЕ ИСПОЛЬЗУЕТСЯ!!!!
        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {




            ////if (comboBox1.SelectedIndex != -1)
            ////{
            ////    OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT * FROM " + comboBox1.SelectedItem, dbCon);
            ////    DataTable dataTable = new DataTable();
            ////    dbAdapter1.Fill(dataTable);
            ////    dataGridView1.DataSource = dataTable;

            ////    /*
            ////    OleDbDataAdapter dbAdapter2 = new OleDbDataAdapter(@"SELECT DISTINCT Группа_показателей FROM " + comboBox1.SelectedItem, dbCon);
            ////    DataTable ds = new DataTable();
            ////    dbAdapter2.Fill(ds);
            ////    comboBox2.DataSource = ds;
            ////    comboBox2.DisplayMember = "Группа_показателей";
            ////    */
            ////}
            ////else
            ////{
            ////    MessageBox.Show("Не выбрана таблица для запроса");
            ////}

            ////foreach (int indexChecked in checkedListBox1.CheckedIndices)
            ////{
            ////    checkedListBox1.SetItemChecked(indexChecked, false);
            ////}

            ////checkedListBox1.CheckOnClick = true;
            ////if (comboBox1.SelectedIndex != -1)
            ////{
            ////    OleDbDataAdapter dbAdapter1 = new OleDbDataAdapter(@"SELECT * FROM " + comboBox1.SelectedItem, dbCon);
            ////    DataTable dataTable = new DataTable();
            ////    dbAdapter1.Fill(dataTable);
            ////    checkedListBox1.DataSource = dataTable;
            ////    checkedListBox1.DisplayMember = "Наименование_показателя_единицы_измерений";
            ////    checkedListBox1.ValueMember = "Код";
            ////}
            ////else
            ////{
            ////    MessageBox.Show("Не выбрана таблица для запроса");
            ////}
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Hide();
                    Form3 form3 = new Form3();
                    form3.Show();
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            dataGridView1.RowHeadersVisible = false;

            if (textBox1.Text != "")
            {
                OleDbDataAdapter dbAdapter2 = new OleDbDataAdapter(@"SELECT * FROM [" + textBox1.Text + "]", dbCon);
                DataTable dataTable = new DataTable();
                dbAdapter2.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }
            else
            {
                MessageBox.Show("Не выбрана таблица для запроса");
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (Convert.ToString(row.Cells[1].Value) == "" & Convert.ToString(row.Cells[2].Value) == "" & Convert.ToString(row.Cells[4].Value) == "")
                    dataGridView1.Rows.Remove(row);
                //row.Visible = false;
                //else row.Visible = true;
            }

            /*
                        foreach (int indexChecked in checkedListBox2.CheckedIndices)
                        {
                            checkedListBox2.SetItemChecked(indexChecked, false);
                        }

                        if (textBox1.Text != "")
                        {
                            OleDbDataAdapter dbAdapter2 = new OleDbDataAdapter(@"SELECT * FROM " + textBox1.Text, dbCon);
                            DataTable dataTable = new DataTable();
                            dbAdapter2.Fill(dataTable);
                            checkedListBox2.DataSource = dataTable;
                            checkedListBox2.DisplayMember = "Наименование_показателя_единицы_измерений";
                            checkedListBox2.ValueMember = "Код";
                        }
                        else
                        {
                            MessageBox.Show("Не выбрана таблица для запроса");
                        }*/

            //видимость некоторых колонок и запрет на ввод во все, кроме "результатов"
            this.dataGridView1.Columns[0].Visible = false;
            this.dataGridView1.Columns[1].ReadOnly = true;
            this.dataGridView1.Columns[2].ReadOnly = true;
            this.dataGridView1.Columns[3].ReadOnly = false;
            this.dataGridView1.Columns[4].ReadOnly = true;
            this.dataGridView1.Columns[5].Visible = false;
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(checkedListBox1.SelectedItem.ToString() + "");
        }
        
        //сортировка по "галочным" показателям//НЕ ИСПОЛЬЗУЕТСЯ!!!!
        private void button9_Click(object sender, EventArgs e)
        {
            checkedListBox2.CheckOnClick = true;

            if (checkedListBox2.CheckedItems.Count > 0) //проверка выбраны ли данные
            {
                //заполнение "галочными" данными//не корректно работает при выборке по показателю
                DataTable dataTable = new DataTable();

                foreach (object item2 in checkedListBox2.CheckedItems)
                {
                    string curItemString2 = ((DataRowView)item2)[checkedListBox2.DisplayMember].ToString();
                    OleDbDataAdapter dbAdapter2 = new OleDbDataAdapter(@"SELECT * FROM [" + textBox1.Text + "] WHERE Наименование_показателя_единицы_измерений = '" + curItemString2 + "'", dbCon);

                    dbAdapter2.Fill(dataTable);
                    dataGridView1.DataSource = dataTable;
                }
            }
            else
            {
                MessageBox.Show("Не выбраны данные для добавления");
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Hide();
            Form3 form3 = new Form3();
            form3.Show();

            form3.textBox1.Text = this.label1.Text;
            form3.comboBox1.Text = this.label2.Text;
            form3.textBox3.Text = this.label3.Text;
            form3.textBox5.Text = this.label5.Text;
            form3.textBox4.Text = this.label4.Text;
            form3.textBox6.Text = this.label6.Text;
            //form3.comboBox2.Text = this.label23.Text;
            form3.textBox7.Text = this.label7.Text;
            form3.textBox8.Text = this.label8.Text;
            form3.dateTimePicker1.Text = this.label9.Text;
            //form3.label10.Text = this.dateTimePicker2.Text + " - " + this.dateTimePicker3.Text;
            form3.textBox11.Text = this.label11.Text;
            form3.textBox12.Text = this.label12.Text;

            //form3.label13.Text = "Температура: " + this.textBox13.Text + " °C," + " Влажность: " + this.textBox2.Text + " %," + " Давление: " + this.textBox9.Text + " мм. рт. ст.";
            //if (textBox16.Text != "")
            //    form1.label13.Text = form1.label13.Text + ", Электромагнитное поле: " + this.textBox16.Text + " А/м";
            //if (textBox10.Text != "")
            //    form1.label13.Text = form1.label13.Text + ", Магнитное поле: " + this.textBox10.Text + " А/м";

            form3.textBox14.Text = this.label14.Text;
            form3.label15.Text = this.label15.Text;
            //form1.comboBox2.Text = this.textBox1.Text;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            //this.Enabled = false;
            //MessageBox.Show("|" + My.poveritelListLength + "|");

            if (My.poveritelListLength == 0)
            {
                Form5 form5 = new Form5();
                form5.Show();
            }

            else if (My.poveritelListLength > 0)
            {                
                Form5 form5 = new Form5();
                form5.Show();
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            Form6 form6 = new Form6();
            form6.Show();
        }
    }
}

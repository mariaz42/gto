using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp7
{

    public partial class Form1 : Form
    {
        SqlConnection MyConnection;
        // System.DateTime stage_1 = new System.DateTime.Year(6);365
        TimeSpan year_6 = TimeSpan.Parse("2191");//2191 < 6 лет
        TimeSpan year_8 = TimeSpan.Parse("2922");//8 лет < 2922 < 9 лет
        TimeSpan year_10 = TimeSpan.Parse("4017");//10 лет < 3652 < 11 лет
        TimeSpan year_12 = TimeSpan.Parse("4748");//12 лет < 4382 < 13 лет
        TimeSpan year_15 = TimeSpan.Parse("5478");//15 лет < 5478 < 16 лет
        TimeSpan year_17 = TimeSpan.Parse("6209");// 17 лет < 6209 < 18 лет
        TimeSpan year_29 = TimeSpan.Parse("6209");// 17 лет < 6209 < 18 лет
        TimeSpan year_39 = TimeSpan.Parse("6209");// 17 лет < 6209 < 18 лет
        TimeSpan year_49 = TimeSpan.Parse("6209");// 17 лет < 6209 < 18 лет
        TimeSpan year_59 = TimeSpan.Parse("6209");// 17 лет < 6209 < 18 лет
        TimeSpan year_69 = TimeSpan.Parse("6209");// 17 лет < 6209 < 18 лет

        string[] list = new string[25];

        public int cat;
        public int g;
        public int ys;
        public int y;

        bool gender;

        public Control[] textboxarray;
        public Control[] labelarray;

        public List<string[]> data = new List<string[]>();
        string[] data_select = new string[25];



        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int cate = 0;

            button2.Visible = true;

            int day = 0;
            int month = 0;
            int year = 0;
            // DateTime sde = DateTime.Parse(DateTime.Now.ye)
            DateTime currentDate = DateTime.Parse(DateTime.Now.Date.ToShortDateString());
            DateTime birthdate = DateTime.Parse(dateTimePicker1.Value.ToString());
            //DateTime birthdate = DateTime.Parse( DataTime);
            day = currentDate.Day - birthdate.Day;
            month = currentDate.Month - birthdate.Month;
            year = currentDate.Year - birthdate.Year;

            System.TimeSpan diff1 = currentDate.Subtract(birthdate);


            //if ( diff1 > year_6 ) { MessageBox.Show("6"); }
            if (diff1 > year_6 && diff1 <= year_8) { cate = 1; }
            if (diff1 > year_8 && diff1 <= year_10) { cate = 2; }
            if (diff1 > year_10 && diff1 <= year_12) { cate = 3; }
            if (diff1 > year_12 && diff1 <= year_15) { cate = 4; }
            if (diff1 > year_15 && diff1 <= year_17) { cate = 5; }

            cat = cate;

            //excer(cat);

            gender_fun(cat);

            dataGridView1.Columns.Clear();

            //Мальчик
            if (comboBox1.SelectedIndex == 0)
            {

                SqlCommand command = new SqlCommand("SELECT  id_stage,name_test.name_test,norm_boys.norm_boys_gold, norm_boys.norm_boys_silver, norm_boys.norm_boys_bronze FROM stage_boys inner join name_test on name_test.id_name_test = stage_boys.id_name_test inner join norm_boys on norm_boys.id_norm_boys = stage_boys.id_stage  where norm_boys.num_stage = @cat ", MyConnection);

                command.Parameters.AddWithValue("cat", cat);

                command.ExecuteNonQuery();

                SqlDataAdapter MyDataAdapter = new SqlDataAdapter(command);
                DataTable MyDataTable = new DataTable();
                MyDataAdapter.Fill(MyDataTable);
                dataGridView1.DataSource = MyDataTable;

                SqlDataReader reader = command.ExecuteReader();

                data = new List<string[]>();

                while (reader.Read())
                {
                    data.Add(new string[5]);
                    data[data.Count - 1][0] = reader[0].ToString();
                    data[data.Count - 1][1] = reader[1].ToString();
                    data[data.Count - 1][2] = reader[2].ToString();
                    data[data.Count - 1][3] = reader[3].ToString();
                    data[data.Count - 1][4] = reader[4].ToString();

                }
                int p = 0;
                foreach (string[] item in data)
                {
                    data_select[p] = item[1];
                    p++;
                }

                //MessageBox.Show(data_select[2]);
                reader.Close();

            }
            else//Девочка
            {
                SqlCommand command = new SqlCommand("SELECT  id_stage,name_test.name_test, norm_girls.norm_girls_gold, norm_girls.norm_girls_silver, norm_girls.norm_girls_bronze FROM stage_girls inner join name_test on name_test.id_name_test = stage_girls.id_name_test inner join norm_girls on norm_girls.id_norm_girls = stage_girls.id_norm_girls  where norm_girls.num_stage = @cat ", MyConnection);

                command.Parameters.AddWithValue("cat", cat);

                command.ExecuteNonQuery();

                SqlDataAdapter MyDataAdapter = new SqlDataAdapter(command);
                DataTable MyDataTable = new DataTable();
                MyDataAdapter.Fill(MyDataTable);
                dataGridView1.DataSource = MyDataTable;

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string connetion = @"Data Source=LAPTOP-O1R0FBHL\SQLEXPRESS;Initial Catalog=GGTO;Integrated Security=true;";
            MyConnection = new SqlConnection(connetion);
            MyConnection.Open();

            string[] categories = new string[2];
            categories[0] = "Мальчик";
            categories[1] = "Девочка";
            comboBox1.Items.Add(categories[0]);
            comboBox1.Items.Add(categories[1]);
            comboBox1.SelectedItem = categories[0];

            dateTimePicker1.MaxDate = new DateTime(2014, DateTime.Now.Month, DateTime.Now.Day);
            dateTimePicker1.MinDate = new DateTime(2004, DateTime.Now.Month, DateTime.Now.Day); ;

            dataGridView1.Visible = false;
            radioButton2.Checked = true;
            button2.Visible = false;

        }


        void pain()
        {
            if (comboBox1.SelectedIndex == 0)
            {

                SqlCommand command = new SqlCommand("SELECT  id_stage,name_test.name_test,norm_boys.norm_boys_gold, norm_boys.norm_boys_silver, norm_boys.norm_boys_bronze FROM stage_boys inner join name_test on name_test.id_name_test = stage_boys.id_name_test inner join norm_boys on norm_boys.id_norm_boys = stage_boys.id_stage  where norm_boys.num_stage = @cat ", MyConnection);

                command.Parameters.AddWithValue("cat", cat);

                command.ExecuteNonQuery();

                SqlDataAdapter MyDataAdapter = new SqlDataAdapter(command);
                DataTable MyDataTable = new DataTable();
                MyDataAdapter.Fill(MyDataTable);
                dataGridView1.DataSource = MyDataTable;

                SqlDataReader reader = command.ExecuteReader();

                data = new List<string[]>();

                while (reader.Read())
                {
                    data.Add(new string[5]);
                    data[data.Count - 1][0] = reader[0].ToString();
                    data[data.Count - 1][1] = reader[1].ToString();
                    data[data.Count - 1][2] = reader[2].ToString();
                    data[data.Count - 1][3] = reader[3].ToString();
                    data[data.Count - 1][4] = reader[4].ToString();

                }
                int p = 0;
                foreach (string[] item in data)
                {
                    data_select[p] = item[1];
                    p++;
                }

                //MessageBox.Show(data_select[2]);
                reader.Close();

            }
            else//Девочка
            {
                SqlCommand command = new SqlCommand("SELECT  id_stage,name_test.name_test, norm_girls.norm_girls_gold, norm_girls.norm_girls_silver, norm_girls.norm_girls_bronze FROM stage_girls inner join name_test on name_test.id_name_test = stage_girls.id_name_test inner join norm_girls on norm_girls.id_norm_girls = stage_girls.id_norm_girls  where norm_girls.num_stage = @cat ", MyConnection);

                command.Parameters.AddWithValue("cat", cat);

                command.ExecuteNonQuery();

                SqlDataAdapter MyDataAdapter = new SqlDataAdapter(command);
                DataTable MyDataTable = new DataTable();
                MyDataAdapter.Fill(MyDataTable);
                dataGridView1.DataSource = MyDataTable;

            }
        }




        private void button2_Click(object sender, EventArgs e)
        {
            result();
        }



        void gender_fun(int cat)
        {
            //MessageBox.Show(comboBox1.SelectedIndex.ToString());

            //Мальчик
            if (comboBox1.SelectedIndex == 0)
            {
                SqlCommand command = new SqlCommand("SELECT  count(id_stage) FROM stage_boys inner join name_test on name_test.id_name_test = stage_boys.id_name_test inner join norm_boys on norm_boys.id_norm_boys = stage_boys.id_stage  where norm_boys.num_stage = @cat ", MyConnection);

                command.Parameters.AddWithValue("cat", cat);

                command.ExecuteNonQuery();

                SqlDataReader reader = command.ExecuteReader();

                reader.Read();

                y = (Convert.ToInt16(reader[0]));

                reader.Close();

                //MessageBox.Show(y.ToString());

                labeltext(y);



            }
            else//Девочка
            {
                SqlCommand command = new SqlCommand("SELECT  count(id_stage) FROM stage_girls inner join name_test on name_test.id_name_test = stage_girls.id_name_test inner join norm_girls on norm_girls.id_norm_girls = stage_girls.id_norm_girls  where norm_girls.num_stage = @cat ", MyConnection);

                command.Parameters.AddWithValue("cat", cat);

                command.ExecuteNonQuery();

                SqlDataReader reader = command.ExecuteReader();

                reader.Read();

                y = (Convert.ToInt16(reader[0]));

                reader.Close();

                //MessageBox.Show(y.ToString());

                labeltext(y);



            }
        }




        public void labeltext(int y)
        {

            pain();

            g++;

            textboxarray = new Control[y];
            labelarray = new Control[y];

            //MessageBox.Show(ys.ToString());


            if (g > 1)
            {
                for (int i = 0; i < ys; i++)
                {

                    foreach (Control tempCtrl in this.Controls)
                    {
                        if (tempCtrl.Name == "tboxi" + i.ToString())
                        {
                            this.Controls.Remove(tempCtrl);
                        }
                    }

                    foreach (Control tempCtrl1 in this.Controls)
                    {
                        if (tempCtrl1.Name == "labia" + i.ToString())
                        {
                            this.Controls.Remove(tempCtrl1);
                        }
                    }
                }


                for (int i = 0; i < y; i++)
                {

                    textboxarray[i] = new TextBox() { Name = "tboxi" + i.ToString(), Font = new Font(FontFamily.GenericMonospace, 14.0F, FontStyle.Regular, GraphicsUnit.Pixel), Location = new Point(10, 5 + (i * 25)), Text = "" };
                    this.Controls.Add(textboxarray[i]);

                    labelarray[i] = new Label() { Name = "labia" + i.ToString(), Font = new Font(FontFamily.GenericMonospace, 14.0F, FontStyle.Regular, GraphicsUnit.Pixel), Location = new Point(130, 5 + (i * 25)), Size = new Size(600, 15), Text = data_select[i] };
                    this.Controls.Add(labelarray[i]);

                }

            }
            else
            {

                for (int i = 0; i < y; i++)
                {

                    textboxarray[i] = new TextBox() { Name = "tboxi" + i.ToString(), Font = new Font(FontFamily.GenericMonospace, 14.0F, FontStyle.Regular, GraphicsUnit.Pixel), Location = new Point(10, 5 + (i * 25)), Text = "" };
                    this.Controls.Add(textboxarray[i]);

                    labelarray[i] = new Label() { Name = "labia" + i.ToString(), Font = new Font(FontFamily.GenericMonospace, 14.0F, FontStyle.Regular, GraphicsUnit.Pixel), Location = new Point(130, 5 + (i * 25)), Size = new Size(600, 15), Text = data_select[i] };
                    this.Controls.Add(labelarray[i]);
                }

            }

            ys = y;


        }

        void result()
        {

            int i = 0;

            int[] data_result = new int[y];

            foreach (Control item in textboxarray)
            {

                if ((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) > (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))))
                {
                    if (!string.IsNullOrEmpty(textboxarray[i].Text) && !string.IsNullOrWhiteSpace(textboxarray[i].Text))
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) >= Convert.ToDouble(item.Text.ToString()))
                        {
                            if (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) >= Convert.ToDouble(item.Text.ToString()))
                            {
                                if (Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value) >= Convert.ToDouble(item.Text.ToString()))
                                {
                                    // //MessageBox.Show(dataGridView1.Rows[i].Cells[4].Value.ToString() + ">" + item.Text.ToString() + "G");
                                    data_result[i] = 1;
                                }
                                else
                                {
                                    ////MessageBox.Show(dataGridView1.Rows[i].Cells[4].Value.ToString() + ">" + item.Text.ToString() + "S");
                                    data_result[i] = 2;
                                }
                            }
                            else
                            {
                                ////MessageBox.Show(dataGridView1.Rows[i].Cells[3].Value.ToString() + ">" + item.Text.ToString() + "B");
                                data_result[i] = 3;
                            }
                        }
                        else
                        {
                            ////MessageBox.Show(dataGridView1.Rows[i].Cells[2].Value.ToString() + ">" + item.Text.ToString() + "N");
                            data_result[i] = 5;
                        }
                    }

                }
                else
                {
                    if (!string.IsNullOrEmpty(textboxarray[i].Text) && !string.IsNullOrWhiteSpace(textboxarray[i].Text))
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) <= Convert.ToDouble(item.Text.ToString()))
                        {
                            if (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) <= Convert.ToDouble(item.Text.ToString()))
                            {
                                if (Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value) <= Convert.ToDouble(item.Text.ToString()))
                                {
                                    ////MessageBox.Show(dataGridView1.Rows[i].Cells[4].Value.ToString() + ">" + item.Text.ToString() + "G");
                                    data_result[i] = 1;
                                }
                                else
                                {
                                    ////MessageBox.Show(dataGridView1.Rows[i].Cells[4].Value.ToString() + ">" + item.Text.ToString() + "S");
                                    data_result[i] = 2;
                                }
                            }
                            else
                            {
                                // //MessageBox.Show(dataGridView1.Rows[i].Cells[3].Value.ToString() + ">" + item.Text.ToString() + "B");
                                data_result[i] = 3;
                            }
                        }

                        else
                        {
                            // //MessageBox.Show(dataGridView1.Rows[i].Cells[2].Value.ToString() + ">" + item.Text.ToString() + "N");
                            data_result[i] = 5;
                        }
                    }
                }
                i++;
            }

            int sum_test = 0;

            foreach (int item in data_result)
            {
                sum_test = sum_test + item;
            }

            double s_y = sum_test / y;
            int final_result = 0;

            //MessageBox.Show(s_y.ToString() + " Итог");
            if (s_y < 3.1)
            {
                label4.Text = "Участник " + textBox1.Text + " получает Бронзовую медаль ГТО!";
                //MessageBox.Show("BRONZE");
                final_result = 3;

                if (s_y < 2.1)
                {
                    label4.Text = "Участник " + textBox1.Text + " получает Сребрянную медаль ГТО!";
                    //MessageBox.Show("SILVER");
                    final_result = 2;

                    if (s_y < 1.1)
                    {
                        final_result = 1;
                        //MessageBox.Show("GOLD");

                        label4.Text = "Участник " + textBox1.Text + " получает Золотую медаль ГТО!";
                    }
                }
            }
            else
            {
                label4.Text = "Участник " + textBox1.Text + " не получает медаль ГТО!";
                //MessageBox.Show("NOT MEDAL");
                final_result = 4;
            }



            SqlCommand command = new SqlCommand("INSERT INTO member_result (member_fio, member_id, member_gender, member_data_born, member_cat, result) VALUES (@member_fio, @member_id,@member_gender, @member_data_born, @member_cat,  @result ) ", MyConnection);

            command.Parameters.AddWithValue("member_fio", textBox1.Text);
            command.Parameters.AddWithValue("member_id", textBox2.Text);
            command.Parameters.AddWithValue("member_gender", comboBox1.SelectedIndex);
            command.Parameters.Add("member_data_born", SqlDbType.Date).Value = dateTimePicker1.Value.Date;
            command.Parameters.AddWithValue("member_cat", cat);
            command.Parameters.AddWithValue("result", final_result);

            command.ExecuteNonQuery();

            // MessageBox.Show(textboxarray[5].Text);  



            //******************************

            SqlCommand command_3 = new SqlCommand("SELECT [result],[member_fio] FROM [member_result]  where [member_id] = @member_id", MyConnection);

            command_3.Parameters.AddWithValue("member_id", textBox2.Text);

            SqlDataReader reader_1 = command_3.ExecuteReader();

            reader_1.Read();

            int ss2 = (Convert.ToInt16(reader_1[0]));

            reader_1.Close();

            //******************************

            for (int k = 0; k < y; k++)
            {

                SqlCommand command_2 = new SqlCommand("SELECT  [id_name_test],[name_test] FROM [name_test] where name_test like @name_test ", MyConnection);

                command_2.Parameters.AddWithValue("name_test", data_select[k]);

                SqlDataReader reader = command_2.ExecuteReader();

                reader.Read();

                string id_name_test = (Convert.ToString(reader[0]));

                reader.Close();

                //******************************************

                SqlCommand command_1 = new SqlCommand("INSERT INTO [data_result]([data_member_result],[id_name_test],[id_member]) VALUES (@data_member_result, @id_name_test, @id_member) ", MyConnection);

                command_1.Parameters.AddWithValue("data_member_result", textboxarray[k].Text);
                command_1.Parameters.AddWithValue("id_name_test", id_name_test);
                command_1.Parameters.AddWithValue("id_member", ss2);

                command_1.ExecuteNonQuery();

            }

            button2.Visible = false;

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.Visible = true;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.Visible = false;
        }
        void filldatagrid()
        {
            SqlCommand command = new SqlCommand("SELECT member_id, [member_fio], member_gender.name_member_gender, [member_data_born], data_cat.name_cate  ,data_prize.name_prize FROM [member_result] inner join member_gender on member_result.member_gender = member_gender.id_member_gender inner join data_cat on member_result.member_cat = data_cat.id_data_cat  inner join data_prize on member_result.[result] = data_prize.id_data_prize ", MyConnection);

            // command.Parameters.AddWithValue("cat", cat);

            command.ExecuteNonQuery();

            SqlDataAdapter MyDataAdapter = new SqlDataAdapter(command);
            DataTable MyDataTable = new DataTable();
            MyDataAdapter.Fill(MyDataTable);
            dataGridView1.DataSource = MyDataTable;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            radioButton1.Checked = true;
            filldatagrid();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            filldatagrid();
            ExportToExcel();
        }

        private void ExportToExcel()
        {
            // Creating a Excel object.
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "ExportedFromDatGrid";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;

                //Loop through each row and read value from each column.
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        // Excel index starts from 1,1. As first Row would have the Column headers, adding a condition check.
                        if (cellRowIndex == 1)
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = dataGridView1.Columns[j].HeaderText;
                        }
                        else
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user.
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveDialog.FilterIndex = 2;

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Готово!");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            filldatagrid();



            try
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Word Documents (*.docx)|*.docx";
                sfd.FileName = "GTO.docx";
                if (sfd.ShowDialog() == DialogResult.OK)
                { Export_Data_To_Word(dataGridView1, sfd.FileName); }
            }
            catch { }
        }

        public void Export_Data_To_Word(DataGridView DGV, string filename)//Функция для экспорта данных из dataFridView в Word
        {
            //Заголовок таблицы в Word
            string[] headtext = new string[6];
            headtext[0] = "ID участника";
            headtext[1] = "ФИО участника";
            headtext[2] = "Пол участника";
            headtext[3] = "Дата рождения";
            headtext[4] = "Степень ГТО";
            headtext[5] = "Медаль ГТО";

            if (DGV.Rows.Count != 0)
            {
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                //add rows
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    } //end row loop
                } //end column loop

                Word.Document oDoc = new Word.Document();
                oDoc.Application.Visible = true;

                //page orintation
                oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";
                    }
                }

                //table format
                oRange.Text = oTemp;
                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;
                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();
                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                //header row style

                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;



                //add header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = headtext[c];
                }

                //table style 
                oDoc.Application.Selection.Tables[1].set_Style("Grid Table 4 - Accent 5");
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header text
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.Text = "your header text";
                    headerRange.Font.Size = 16;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
                //save the file
                oDoc.SaveAs2(filename);
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using ExcelObj = Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


using GemBox.Spreadsheet;
using GemBox.Spreadsheet.WinFormsUtilities;


namespace MainProject
{
    public partial class Form1 : Form
    {

        public void Clear(DataGridView dataGridView)
        {
            while (dataGridView.Rows.Count > 1)
                for (int i = 0; i < dataGridView.Rows.Count - 1; i++)
                    dataGridView.Rows.Remove(dataGridView.Rows[i]);
        }

        public void ClearAll()
        {
            Clear(inputStudentGridView);
            //Clear(dataGridExams);
            dataGridExams.Rows.Clear();
            radioButton3.Checked = false;
            radioButton4.Checked = false;
            radioButton5.Checked = false;
            radioButton6.Checked = false;
            radioButton7.Checked = false;
            groupBox.Text = null;
            courseBox.SelectedIndex = -1;
            radioButton8.Checked = false;
            radioButton9.Checked = false;
        }

        public bool Error()
        {
            bool flag=true;
            bool NoSurname = false;
            bool NoName = false;
            bool NoLastName = false;
            bool NoExams = false;
            bool NoNotes = false;
            bool Group = false;
            bool Course = false;
            bool check = false;

            if (inputStudentGridView[0, 0].Value != null)
            {
                if (isCorrect(inputStudentGridView[0, 0].Value.ToString()))
                    NoSurname = true;
            }
            if (inputStudentGridView[1, 0].Value != null)
            {
                if (isCorrect(inputStudentGridView[1, 0].Value.ToString()))
                    NoName = true;
            }
            if (inputStudentGridView[2, 0].Value != null)
            {
                if (isCorrect(inputStudentGridView[2, 0].Value.ToString()))
                    NoLastName = true;
            }
            if (courseBox.SelectedIndex != -1)
                Course = true;
            if ((isCorrect(groupBox.Text)))
                Group = true;

            for (int i = 0; i < dataGridExams.RowCount-1; i++)
            {
                if (dataGridExams[i, 0].Value != null)
                {
                    if (isCorrect(dataGridExams[i, 0].Value.ToString()))
                        NoExams = true;
                    else break;
                }else break;

                if (dataGridExams[i, 0].Value != null)
                {
                    if (isCorrectNote(dataGridExams[i, 1].Value.ToString()))
                        NoNotes = true;
                    else break;
                }
                break;

            }

            if (radioButton8.Checked == true || radioButton9.Checked == true)
                check = true;


            if (!NoSurname || !NoName || !NoLastName || !NoExams || !NoNotes || !Group || !Course || !check)
            {
                MessageBox.Show("Введите корректные данные: заполните ФИО, Номер группы и Список экзаменов символами, " +
                    "Оценки - 'отл', 'хор', 'уд', 'неуд', выберите Курс, поставьте Своевременность");
                flag = false;
            }

            return flag;

        }

        public bool isCorrect(string data)
        {
            bool ok = false;
            if (data != null)
            {
                if (data.ToString() != "" && data.ToString() != "  ")
                {
                    ok = true;
                }
            }
            return ok;
        }

        public bool isCorrectNote(string data)
        {
            bool ok = false;
            if (data=="отл"|| data == "хор" || data == "неуд" || data == "уд")
            {
                ok = true;
            }

            return ok;
        }

        public Form1()
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            InitializeComponent();

        }

        

        private void обАвтореToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form4 frm = new Form4();
            frm.Show();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "student.xlsx");
                var workbook = ExcelFile.Load(path);
                
            DataGridViewConverter.ExportToDataGridView(workbook.Worksheets.ActiveWorksheet, this.mainGridView, new ExportToDataGridViewOptions() { ColumnHeaders = true });

        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            saveFileDialog.FilterIndex = 3;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                var workbook = new ExcelFile();
                var worksheet = workbook.Worksheets.Add("Sheet1");

                // From DataGridView to ExcelFile.
                DataGridViewConverter.ImportFromDataGridView(worksheet, this.mainGridView, new ImportFromDataGridViewOptions() { ColumnHeaders = true });

                workbook.Save(saveFileDialog.FileName);
                MessageBox.Show("Данные сохранены");
            }
            Application.Exit();
        }



        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files| *.xls; *.xlsx; *.xlsm";
            openFileDialog.FilterIndex = 1;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var workbook = ExcelFile.Load(openFileDialog.FileName);

                DataGridViewConverter.ExportToDataGridView(workbook.Worksheets.ActiveWorksheet, this.mainGridView, new ExportToDataGridViewOptions() { ColumnHeaders = true });
            }
        }

        private void button1_Click(object sender, EventArgs e)

        {
            bool flag = Error();

           if (flag)
           {
                int countRow = mainGridView.RowCount;
                mainGridView.RowCount++;

                int column = dataGridExams.ColumnCount;
                string[] exams = new string[column];
                string[] notes = new string[column];

                for (int i = 0; i < column; i++)
                {
                    exams[i] = dataGridExams[i, 0].Value.ToString();
                    notes[i] = dataGridExams[i, 1].Value.ToString();
                }

                string ok = "0";
                if (radioButton8.Checked == true)
                {
                    ok = "вовремя";
                }
                else if (radioButton9.Checked == false)
                {
                    ok = "не вовремя";
                }

                mainGridView[0, countRow].Value = inputStudentGridView[0, 0].Value.ToString();
                mainGridView[1, countRow].Value = inputStudentGridView[1, 0].Value.ToString();
                mainGridView[2, countRow].Value = inputStudentGridView[2, 0].Value.ToString();
                mainGridView[3, countRow].Value = Convert.ToInt32(courseBox.Text);
                mainGridView[4, countRow].Value = groupBox.Text;
                mainGridView[5, countRow].Value = ok;


                for (int i = 0; i < column; i++)
                {
                    mainGridView[6 + i, countRow].Value = exams[i];
                }
                for (int i = 0; i < column; i++)
                {
                    mainGridView[11 + i, countRow].Value = notes[i];
                }


                if (radioButton2.Checked == true)
                {
                    int totalRows = mainGridView.Rows.Count;
                    int rowIndex = mainGridView.SelectedCells[0].OwningRow.Index;

                    if (rowIndex != totalRows - 1)
                    {
                        var selectedRow = mainGridView.Rows[totalRows - 1];
                        mainGridView.Rows.Remove(selectedRow);
                        mainGridView.Rows.Insert(rowIndex + 1, selectedRow);
                    }
                    mainGridView.ClearSelection();
                    mainGridView.Rows[rowIndex + 1].Cells[rowIndex + 1].Selected = true;
                }

                if (radioButton13.Checked == true)
                {
                    int totalRows = mainGridView.Rows.Count;
                    int rowIndex = mainGridView.SelectedCells[0].OwningRow.Index;
                    var selectedRow = mainGridView.Rows[totalRows - 1];
                    mainGridView.Rows.Remove(selectedRow);
                    mainGridView.Rows.Insert(rowIndex, selectedRow);
                    mainGridView.ClearSelection();
                    mainGridView.Rows[rowIndex].Selected = true;
                }

                ClearAll();
           }

        }
        

        private void radioButton7_Click(object sender, EventArgs e)
        {

            if (radioButton3.Checked == true)
            {
                dataGridExams.ColumnCount = 1;
            }
            else if (radioButton4.Checked == true)
            {
                dataGridExams.ColumnCount = 2;
            }
            else if (radioButton5.Checked == true)
            {
                dataGridExams.ColumnCount = 3;
            }
            else if (radioButton6.Checked == true)
            {
                dataGridExams.ColumnCount = 4;
            }
            else if (radioButton7.Checked == true)
            {
                dataGridExams.ColumnCount = 5;
            }

            dataGridExams.RowCount = 2;
            dataGridExams.Rows[0].HeaderCell.Value = "Экзамен";
            dataGridExams.Rows[1].HeaderCell.Value = "Оценка";
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            saveFileDialog.FilterIndex = 3;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                var workbook = new ExcelFile();
                var worksheet = workbook.Worksheets.Add("Sheet1");

                // From DataGridView to ExcelFile.
                DataGridViewConverter.ImportFromDataGridView(worksheet, this.mainGridView, new ImportFromDataGridViewOptions() { ColumnHeaders = true });

                workbook.Save(saveFileDialog.FileName);
                MessageBox.Show("Данные сохранены");
            }

            
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            saveFileDialog.FilterIndex = 3;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                var workbook = new ExcelFile();
                var worksheet = workbook.Worksheets.Add("Sheet1");
                DataGridViewConverter.ImportFromDataGridView(worksheet, this.mainGridView, new ImportFromDataGridViewOptions() { ColumnHeaders = true });
                workbook.Save(saveFileDialog.FileName);
                MessageBox.Show("Данные сохранены");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

            mainGridView.Rows[0].Selected = true;
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            int max = mainGridView.Rows.Count;

            int index = mainGridView.CurrentRow.Index;
            if (index == 0)
            {
                mainGridView.CurrentCell = mainGridView[0,max-1];
            }
            else
            {
                mainGridView.CurrentCell = mainGridView[0, index - 1];
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int max = mainGridView.Rows.Count;
            mainGridView.Rows[max-1].Selected = true;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int max = mainGridView.Rows.Count;

            int index = mainGridView.CurrentRow.Index;
            if (index == max-1)
            {
                mainGridView.CurrentCell = mainGridView[0, 0];
            }
            else
            {
                mainGridView.CurrentCell = mainGridView[0, index + 1];
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            int index = mainGridView.CurrentRow.Index;
            mainGridView.Rows.RemoveAt(index);

        }

        private void создатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "student.xlsx");
            var workbook = ExcelFile.Load(path);

            DataGridViewConverter.ExportToDataGridView(workbook.Worksheets.ActiveWorksheet, this.mainGridView, new ExportToDataGridViewOptions() { ColumnHeaders = true });
        }

        private void поОдномуПолюToolStripMenuItem_Click(object sender, EventArgs e)

        {
            string head = mainGridView.Columns[0].Name;
            mainGridView.Sort(mainGridView.Columns[head], ListSortDirection.Ascending);
        }

        private void сложнаяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int totalRows = mainGridView.Rows.Count;
            bool flag = true;
            while (flag)
            {
                flag = false;
                for (int i = 0; i < totalRows - 1; ++i)
                {
                    if (mainGridView[0, i].Value.ToString().CompareTo(mainGridView[0, i + 1].Value.ToString()) > 0)
                    {

                        var selectedRow1 = mainGridView.Rows[i];
                        var selectedRow2 = mainGridView.Rows[i + 1];
                        mainGridView.Rows.Remove(selectedRow1);
                        mainGridView.Rows.Remove(selectedRow2);
                        mainGridView.Rows.Insert(i, selectedRow2);
                        mainGridView.Rows.Insert(i + 1, selectedRow1);
                        flag = true;
                    }
                }
            }
            flag = true;

            while (flag)
            {
                flag = false;
                for (int i = 0; i < totalRows - 1; ++i)
                {                 
                    if (mainGridView[0, i].Value.ToString().CompareTo(mainGridView[0, i + 1].Value.ToString()) == 0
                        && mainGridView[1, i].Value.ToString().CompareTo(mainGridView[1, i + 1].Value.ToString()) > 0)
                    {
                        int k = i;
                        var selectedRow1 = mainGridView.Rows[i];
                        var selectedRow2 = mainGridView.Rows[i + 1];
                        mainGridView.Rows.Remove(selectedRow1);
                        mainGridView.Rows.Remove(selectedRow2);
                        mainGridView.Rows.Insert(i, selectedRow2);
                        mainGridView.Rows.Insert(i + 1, selectedRow1);
                        flag = true;
                    }
                    
                }
            }

            flag = true;
            while (flag)
            {
                flag = false;
                for (int i = 0; i < totalRows - 1; ++i)
                {
                    if ((mainGridView[0, i].Value.ToString().CompareTo(mainGridView[0, i + 1].Value.ToString()) == 0)
                        && mainGridView[1, i].Value.ToString().CompareTo(mainGridView[1, i + 1].Value.ToString()) == 0
                        && mainGridView[2, i].Value.ToString().CompareTo(mainGridView[2, i + 1].Value.ToString()) > 0)
                    {
                        var selectedRow1 = mainGridView.Rows[i];
                        var selectedRow2 = mainGridView.Rows[i + 1];
                        mainGridView.Rows.Remove(selectedRow1);
                        mainGridView.Rows.Remove(selectedRow2);
                        mainGridView.Rows.Insert(i, selectedRow2);
                        mainGridView.Rows.Insert(i + 1, selectedRow1);
                        flag = true;

                    }
                }
            }
            flag = true;
            while (flag)
            {
                flag = false;
                for (int i = 0; i < totalRows - 1; ++i)
                {
                    if (mainGridView[0, i].Value.ToString().CompareTo(mainGridView[0, i + 1].Value.ToString()) == 0
                        && mainGridView[1, i].Value.ToString().CompareTo(mainGridView[1, i + 1].Value.ToString()) == 0
                        && mainGridView[2, i].Value.ToString().CompareTo(mainGridView[2, i + 1].Value.ToString()) == 0
                        && mainGridView[3, i].Value.ToString().CompareTo(mainGridView[3, i + 1].Value.ToString()) > 0)
                    {
                        var selectedRow1 = mainGridView.Rows[i];
                        var selectedRow2 = mainGridView.Rows[i + 1];
                        mainGridView.Rows.Remove(selectedRow1);
                        mainGridView.Rows.Remove(selectedRow2);
                        mainGridView.Rows.Insert(i, selectedRow2);
                        mainGridView.Rows.Insert(i + 1, selectedRow1);
                        flag = true;

                    }
                }
            }

        }

        private void вариантЗаданияУсловияПоискафильтрацияИСортировкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Вариант 2. Заполнить таблицу данными о студенте и его успехах. Осуществить фильтрацию: простая - по фамилии, " +
               "сложная - по ФИО и группе");
        }

        private void очиститьТаблицуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainGridView.Rows.Clear();
        }


        public void reset (DataGridView mainGridView, DataGridView searchDataGridView, int row, int i)
        {
            searchDataGridView.Rows.Add();
            for (int j = 0; j < mainGridView.ColumnCount - 1; ++j)
            {
                searchDataGridView[j, row].Value = mainGridView[j, i].Value;
            }
        } 

        public bool Exist(string mainObject, string value)
        {
            bool result = (value == "") || ((mainObject == value));
            return result;
        }

        public bool check(DataGridView mainDataGrid, string[] buffer, int i)
        {
            bool result = false;
            bool result1 = Exist(mainDataGrid[0, i].Value.ToString(), buffer[0]);
            bool result2 = Exist(mainDataGrid[1, i].Value.ToString(), buffer[1]);
            bool result3 = Exist(mainDataGrid[2, i].Value.ToString(), buffer[2]);
            bool result4 = Exist(mainDataGrid[4, i].Value.ToString(), buffer[3]);
            bool result5 = Exist(mainDataGrid[3, i].Value.ToString(), buffer[4]);

            if (Exist(mainDataGrid[0,i].Value.ToString(), buffer[0]) 
                && Exist(mainDataGrid[1, i].Value.ToString(), buffer[1])
                && Exist(mainDataGrid[2, i].Value.ToString(), buffer[2])
                && Exist(mainDataGrid[4, i].Value.ToString(), buffer[3]) 
                && Exist(mainDataGrid[3,i].Value.ToString(), buffer[4]))
                
            {
                result = true;
            }
            return result;
        }

        string [] filtration( DataGridView mainGridView, string[] buffer)
        {
            int totalRows = mainGridView.Rows.Count;
            int row = 0;
            string [] result = new string [totalRows];
            
            for (int i = 0; i < totalRows; ++i)
            {
                if (check(mainGridView, buffer, i))
                {
                    result[row]=Convert.ToString(i);
                    ++row;
                }
            }
            return result;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            searchDataGridView.Rows.Clear();
            int totalRows = mainGridView.Rows.Count;
            String[] buffer = new String[5];
            String[] index = new String[totalRows];

            if (dataGridStudentSearch[0, 0].Value == null) buffer[0] = ""; else { buffer[0] = dataGridStudentSearch[0, 0].Value.ToString();};
            if (dataGridStudentSearch[1, 0].Value == null) buffer[1] = ""; else { buffer[1] = dataGridStudentSearch[1, 0].Value.ToString();};
            if (dataGridStudentSearch[2, 0].Value == null) buffer[2] = ""; else { buffer[2] = dataGridStudentSearch[2, 0].Value.ToString();};
            if (textBoxSearch.Text == null) buffer[3] = ""; else { buffer[3] = textBoxSearch.Text;};
            if (comboBoxSearch.SelectedIndex == -1) buffer[4] = ""; else { buffer[4] = comboBoxSearch.Text;  };

            index = filtration(mainGridView, buffer);
            int ind = 0;

            while (index[ind] != null)
            {
                ind++;
            }
               
            for (int i = 0; i < ind; ++i)
            {
                reset(mainGridView, searchDataGridView, i, Convert.ToInt32(index[i]));
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridStudentSearch.Rows.Clear();
            comboBoxSearch.SelectedIndex = -1;
            textBoxSearch.Clear();
        }
    }
}

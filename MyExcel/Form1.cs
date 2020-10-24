using System;
using System.Collections.Generic;
using System.Windows.Forms;


namespace MyExcel
{
    public partial class Form1 : Form
    {
        string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        int curRow, curCol;
        string last_cell_name;
        public static Dictionary<string, Cell> dict1 = new Dictionary<string, Cell>();

        public Form1()
        {
            InitializeComponent();
            
            //Создаём начальную таблицу
            DataGridViewColumn A = new DataGridViewColumn();
            DataGridViewColumn B = new DataGridViewColumn();
            DataGridViewColumn C = new DataGridViewColumn();
            DataGridViewColumn D = new DataGridViewColumn();

            Cell CellA = new Cell(); A.CellTemplate = CellA;
            Cell CellB = new Cell(); B.CellTemplate = CellB;
            Cell CellC = new Cell(); C.CellTemplate = CellC;
            Cell CellD = new Cell(); D.CellTemplate = CellD;

            A.HeaderText = "A"; A.Name = "A";
            B.HeaderText = "B"; B.Name = "B";
            C.HeaderText = "C"; C.Name = "C";
            D.HeaderText = "D"; D.Name = "D";

            dataGridView1.Columns.Add(A);
            dataGridView1.Columns.Add(B);
            dataGridView1.Columns.Add(C);
            dataGridView1.Columns.Add(D);

            for (int i = 0; i < 4; i++)
            {
                DataGridViewRow row = new DataGridViewRow();
                dataGridView1.Rows.Add(row);
                setRowNumber(dataGridView1); 
            }


        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }


        //Клетка созадётся в момент нажатия на неё, чтобы создавалось ровно столько объектов, сколько нам нужно.
        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            curCol = dataGridView1.CurrentCell.ColumnIndex;
            curRow = dataGridView1.CurrentCell.RowIndex;
            string cell_name = (char)(curCol + 65) + (curRow + 1).ToString();
            if (!dict1.ContainsKey(cell_name))
            {
                Cell cell = new Cell();
                cell.Value = "0";
                cell.Exp = "0";
                dict1[cell_name] = cell;
            }
        }

        //Заполнение клетки
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                curCol = dataGridView1.CurrentCell.ColumnIndex;
                curRow = dataGridView1.CurrentCell.RowIndex;
                string cell_name = (char)(curCol + 65) + (curRow + 1).ToString();
                double res = Calculator.Evaluate(dataGridView1[curCol, curRow].Value.ToString());
                dict1[cell_name].Exp = dataGridView1[curCol, curRow].Value.ToString();
                textBox1.Text = dict1[cell_name].Exp;
                dict1[cell_name].Value = res.ToString();
                dataGridView1[curCol, curRow].Value = res.ToString();
                RefreshCells();
            }
            catch(ArgumentException)
            {
                string message = "Будь ласка, пересвідчіться, що вираз правильний";
                MessageBox.Show(message);
            }
            catch (DivideByZeroException)
            {
                string message = "Ділити на нуль не можна";
                MessageBox.Show(message);
                dict1[last_cell_name].Exp = "0";
                dict1[last_cell_name].Value = "0";
                dataGridView1[curCol, curRow].Value = dict1[last_cell_name].Value;
                textBox1.Text = dict1[last_cell_name].Exp;
            }
            catch (NullReferenceException) { }
        }


        //Добавление строк
        private void AddRow_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = new DataGridViewRow();
            dataGridView1.Rows.Add(row);
            setRowNumber(dataGridView1);
        }


        //Удаление строк
        private void DelRow_Click(object sender, EventArgs e)
        {
            int row = dataGridView1.CurrentCell.RowIndex;
            dataGridView1.Rows.RemoveAt(row);
            setRowNumber(dataGridView1);
            for (int col = 0; col < dataGridView1.ColumnCount; col++)
            {
                string cell_name = (char)(col + 65) + (row + 1).ToString();
                if (dict1.ContainsKey(cell_name))
                {
                    dict1[cell_name].Value = "0";
                    dict1[cell_name].Exp = "0";
                }
            }
            RefreshCells();
        }


        //Добавление колонок
        private void AddCol_Click(object sender, EventArgs e)
        {
            DataGridViewColumn column = new DataGridViewColumn();
            DataGridViewCell n = new DataGridViewTextBoxCell(); column.CellTemplate = n;
            try
            {
                int a = dataGridView1.ColumnCount - 1;
                string b = dataGridView1.Columns[a].Name;
                int c = letters.IndexOf(b);
                char d = letters[c + 1];
                string f = d.ToString();
                column.HeaderText = f; column.Name = f;
                dataGridView1.Columns.Add(column);
            }
            catch (Exception)
            {
                string message = "Дуже багато колонок. Будь, ласка видаліть одну";
                MessageBox.Show(message);
            }
        }


        //Удаление колонок
        private void DelCol_Click(object sender, EventArgs e)
        {
            int col = dataGridView1.ColumnCount - 1;
            for (int row = 0; row < dataGridView1.RowCount; row++)
            {
                string cell_name = (char)(col + 65) + (row + 1).ToString();
                if (dict1.ContainsKey(cell_name))
                {
                    dict1[cell_name].Value = "0";
                    dict1[cell_name].Exp = "0";
                }
            }
            if (col > 0)
            {
                dataGridView1.Columns.RemoveAt(col);
                RefreshCells();
            }
            else if (col == 0)
            {
                string message = "Вы не можете видалити останню колонку";
                MessageBox.Show(message);
            }
            else
            {
                string message = "Щось не так, як повинно бути";
                MessageBox.Show(message);
            }
        }


        //Ищем совпадение формули в Textbox и Expr у клетки
        private void textBox1_Enter(object sender, EventArgs e)
        {
            try
            {
                foreach(KeyValuePair<string, Cell> name in dict1)
                {
                    if (name.Value.Exp == textBox1.Text)
                    {
                        last_cell_name = name.Key;
                    }
                }
            }
            catch { }
        }

        //Отправляем Expr вибраной клетки в Textbox
        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            curCol = dataGridView1.CurrentCell.ColumnIndex;
            curRow = dataGridView1.CurrentCell.RowIndex;
            string cell_name = (char)(curCol + 65) + (curRow + 1).ToString();
            if (!dict1.ContainsKey(cell_name))
            {
                Cell cell = new Cell();
                cell.Value = "0";
                cell.Exp = "0";
                dict1[cell_name] = cell;
            }
            textBox1.Text = dict1[cell_name].Exp;
        }


        //Сохраняем изменения Textbox в Expr клетки
        private void textBox1_Leave(object sender, EventArgs e)
        {
            try 
            {
                dict1[last_cell_name].Exp = textBox1.Text;
                dict1[last_cell_name].Value = Calculator.Evaluate(textBox1.Text).ToString();
                string b = last_cell_name[1].ToString();
                curCol = (char)last_cell_name[0] - 65;
                curRow = Convert.ToInt32(b) - 1;
                dataGridView1[curCol, curRow].Value = dict1[last_cell_name].Value;
                foreach (KeyValuePair<string, Cell> name in dict1)
                {
                    if (name.Key != last_cell_name)
                    {
                        RefreshCells();
                    }
                }
            }
            catch (DivideByZeroException)
            {
                string message = "Ділити на нуль не можна";
                MessageBox.Show(message);
                dict1[last_cell_name].Exp = "0";
                dict1[last_cell_name].Value = "0";
                dataGridView1[curCol, curRow].Value = dict1[last_cell_name].Value;
                textBox1.Text = dict1[last_cell_name].Exp;
            }
        }


        //Обновление значения в клетках
        private void RefreshCells()
        {
            foreach (KeyValuePair<string, Cell> name in dict1)
            {
                name.Value.Value = Calculator.Evaluate(name.Value.Exp).ToString();
                string cell = name.Key;
                string a = cell[1].ToString();
                curCol = (char)cell[0] - 65;
                curRow = Convert.ToInt32(a) - 1;
                dataGridView1[curCol, curRow].Value = dict1[name.Key].Value;
            }
        }


        //Назначение номера для каждой строки
        private void setRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = String.Format("{0}", row.Index + 1);
            }
        }
    }
}

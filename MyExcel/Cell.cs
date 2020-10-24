using System.Windows.Forms;

namespace MyExcel
{
    public class Cell : DataGridViewTextBoxCell
    {
        private string val;
        private string exp;

        public Cell()
        {
            val = "";
            exp = "";
        }

        public string Value
        {
            get { return val; }
            set { val = value; }
        }

        public string Exp
        {
            get { return exp; }
            set { exp = value; }
        }

    }
}

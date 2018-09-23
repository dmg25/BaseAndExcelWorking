using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConnectToSCADABD
{
    public partial class FormInitVal : Form
    {
        string condition;
        bool edit = false;

        public List<ProgramReadDB.InitValue> InitValues = new List<ProgramReadDB.InitValue>(); // лист объектов с начальными значениями
        
        public FormInitVal()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("Данные будут записаны в БД! Продолжить?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
            {

                foreach (ProgramReadDB.InitValue iv in InitValues)
                {
                    string SQL = "Update ISACARDS set INITIALVALUE = '" + textBox1.Text + "' where CARDSID =" + iv.ObjID;

                    ProgramConnect connect = new ProgramConnect();
                    connect.WriteToBase(SQL);
                }

                MessageBox.Show("Данные записаны");
            }
        }

        public void FormInitVal_Load(object sender, EventArgs e)
        {
           // bool firstIter = true;
            foreach (ProgramReadDB.InitValue iv in InitValues)
            {
                listBox1.Items.Add(iv.Marka);

          /*      if (firstIter)
                {
                    condition = condition + "CARDSID = " + iv.ObjID;
                    firstIter = false;
                }
                else 
                {
                    condition = condition + " or CARDSID = " + iv.ObjID;
                }*/

            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = InitValues[listBox1.SelectedIndex].InitialValue ;
        }
        
        private void textBox1_Click(object sender, EventArgs e)
        {
            if (!edit)
            {
                MessageBox.Show("Редактируйте поле внимательно");
                edit = true;
            }
        }

     }
}

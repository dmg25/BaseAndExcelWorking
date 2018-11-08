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
        public string BaseAddr;
        bool edit = false;

        public List<string> ObjTypeName = new List<string>(); //массив типов объектов
        public List<string> ObjTypeNameUnic = new List<string>(); //массив типов объектов без повторений
        public class InitValuesLists { public List<ProgramReadDB.InitValue> InitValuesList = new List<ProgramReadDB.InitValue>();}
        public List<ProgramReadDB.InitValue> InitValues = new List<ProgramReadDB.InitValue>(); // лист объектов с начальными значениями
        public List<InitValuesLists> InitValuesLists1 = new List<InitValuesLists>();
        public List<ProgramReadDB.InitValue> InitValuesListType = new List<ProgramReadDB.InitValue>();

        public FormInitVal()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("Данные будут записаны в БД! Продолжить?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
            {

                foreach (ProgramReadDB.InitValue list in InitValuesLists1[comboBox1.SelectedIndex].InitValuesList)
                {
                    string SQL = "Update ISACARDS set INITIALVALUE = '" + textBox1.Text + "' where CARDSID =" + list.ObjID;

                    ProgramConnect connect = new ProgramConnect();
                    connect.WriteToBase(SQL, BaseAddr);
                }

                MessageBox.Show("Данные записаны");
                label4.Visible = true;
            }
        }

        public void FormInitVal_Load(object sender, EventArgs e)
        {
            foreach (ProgramReadDB.InitValue iv in InitValues)
            {
                listBox1.Items.Add(iv.Marka);
                ObjTypeName.Add(iv.ObjType);      
            }

            ObjTypeNameUnic = ObjTypeName.Distinct().ToList(); //убираем повторяющиеся типы объектов
            foreach (string name in ObjTypeNameUnic)
            {
                comboBox1.Items.Add(name);
            }


            for (int i = 0; i < ObjTypeNameUnic.Count; i++)
            {
                InitValuesListType.Clear();
                foreach (ProgramReadDB.InitValue iv in InitValues)
                {
                    if (iv.ObjType == ObjTypeNameUnic[i]) { InitValuesListType.Add(iv); }
                }
                var list = new InitValuesLists() { InitValuesList = InitValuesListType.ToList() };
                InitValuesLists1.Add(list);
            }

            comboBox1.SelectedIndex = 0;
            //-------------первоначальный выбор списка объектов----------------------------------------------------
            listBox1.Items.Clear();
            foreach (ProgramReadDB.InitValue list in InitValuesLists1[comboBox1.SelectedIndex].InitValuesList)
            {
                listBox1.Items.Add(list.Marka);
            }

            listBox1.SetSelected(0, true); // выбираем первый элемент по дефолту
            //----------------------------------------------------------------------------------------------------
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           // textBox1.Text = InitValues[listBox1.SelectedIndex].InitialValue ;
            textBox1.Text = InitValuesLists1[comboBox1.SelectedIndex].InitValuesList[listBox1.SelectedIndex].InitialValue;
        }
        
        private void textBox1_Click(object sender, EventArgs e)
        {
            if (!edit)
            {
                MessageBox.Show("Редактируйте поле внимательно");
                edit = true;
            }
        }
       // private void comboBox1_SelectedIndexChanged() { }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox1.Items.Clear();

            foreach (ProgramReadDB.InitValue list in InitValuesLists1[comboBox1.SelectedIndex].InitValuesList)
            {
                listBox1.Items.Add(list.Marka);
            }       
         }
     }
}

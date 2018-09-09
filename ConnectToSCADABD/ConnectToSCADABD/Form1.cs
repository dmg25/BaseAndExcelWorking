using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FirebirdSql.Data.FirebirdClient;

namespace ConnectToSCADABD
{

    
    public partial class Form1 : Form
    {

       
        DataGridView TmpDG; // переменная для сменяемости таблиц
        string SQL; // "select * from OBJTYPE "; //Where Speed=@Speed
        string BaseAddr; //адрес базы

        string ConStr;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            BaseAddr = textBox2.Text;

            ConStr = "character set=WIN1251;initial catalog=" + BaseAddr + ";user id=SYSDBA;password=masterkey"; // наша строка подключения, сделать её изменяемой!!!
           
          // SQL = "select * from ISAOBJ";
           SQL = textBox1.Text;

           if  (radioButton1.Checked) 
           {
               TmpDG = Table1;
           }
           else if (radioButton2.Checked)
           {
               TmpDG = Table2;
           }

           ConnectToBase();
        
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ConStr = "character set=WIN1251;initial catalog=C:\\SCADABD.GDB;user id=SYSDBA;password=masterkey"; // наша строка подключения, сделать её изменяемой!!!
            TmpDG = Table2;
        //    SQL = "select * from ISAUPCHANNEL";
            ConnectToBase();
        }


        
        private void ConnectToBase(/*object sender, EventArgs e*/)
        {
            // Строка подключения
                using (FbConnection fbc = new FbConnection(ConStr))   // используем Using для последующего высвобождения ресурсов, должно быть оптимальнее
                {
                    try
                    {
                        fbc.Open();                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message,
                            System.IO.Path.GetFileName(
                            System.Reflection.Assembly.GetExecutingAssembly().Location));
                        return;
                    }

                    // Транзакция, которая тупо откуда-то скопирована
                    FbTransactionOptions fbto = new FbTransactionOptions();
                    fbto.TransactionBehavior = FbTransactionBehavior.NoWait |
                         FbTransactionBehavior.ReadCommitted |
                         FbTransactionBehavior.RecVersion;
                    FbTransaction fbt = fbc.BeginTransaction(fbto);

                    // Создаем простой запрос
                   
                    FbCommand fbcom = new FbCommand(SQL, fbc, fbt);
                    //            fbcom.Parameters.Clear();
                    //              fbcom.Parameters.AddWithValue("speed", 100);
 
                    // Создаем адаптер данных
                    FbDataAdapter fbda = new FbDataAdapter(fbcom);

                    DataSet ds = new DataSet();   //из формы не тащится, видимо там надо задавать способ подключения через стандартные способы, а они не пашут.

                    try
                    {
                        fbda.Fill(ds);  // заполняем DataSet
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message,
                            System.IO.Path.GetFileName(
                            System.Reflection.Assembly.GetExecutingAssembly().Location));                       
                        return;
                    }
                    finally
                    {
                        fbt.Rollback();
                        fbc.Close();
                    }

                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0) return;

                    

                    TmpDG.DataSource = dt;
                    
                   // MessageBox.Show("ololo");
                    label1.Text = "Имя таблицы: " + dt.TableName + " ; Кол-во столбцов:" + dt.Columns.Count + " ; Кол-во строк:" + dt.Rows.Count;
                  //  Console.WriteLine("Имя таблицы: " + dt.TableName + " ; Кол-во столбцов:" + dt.Columns.Count + " ; Кол-во строк:" + dt.Rows.Count);

                 /*   for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        Console.Write("\t{0}", dt.Columns[i].ColumnName);
                    }

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        DataRow dr = dt.Rows[i];

                        if (dr.IsNull("NAME") == false)
                        {
                            // необходимые действия
                            //  Int32 id = Convert.ToInt32(dr["ID"]);

                            var cells = dt.Rows[i].ItemArray;
                            foreach (object cell in cells)
                                Console.Write("\t{0}", cell);
                            Console.WriteLine();



                        }
                    }*/
                    
                }
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

       

    }
}

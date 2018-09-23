/* Добавить:
 * 1.+ Диагностика неправильного файла экспорта (сообщение и возврат к началу)
 * 2.+ Блокировку кнопок, чтобы не сделать один этап раньше момента
 * 3.+ Добавить нормальный путь для сохранения файла и прописывать ему автоимя
 * 4.+ Сделать раскрывающийся список с ЛИСТАМИ и под каждый (приписывая тип объекта),
 *    сделать свой выбор параметров канала
 * 5. Нам не нужно записывать каждый канал. Подумать, каким образом это осуществить.
 * 6.+ Реакция на нажатие кнопки ОТМЕНА при открытии и сохранении
 * -----------------------------------------------------------------
 * Параметры сохранения:
 * 1. Добавить перечень листов (с указанием типа) и для каждого настройку параметров канала
 * 2. Сделать галочку "Применить ко всем типам"
 * 3. Сделать галочку Сохранить без параметров каналов
 *    В таком случае игнорировать индексы и сохранять в один файл
 *    Отписывать в лог решения пользователя (как сохранен файл, с какими параметрами)
 *    
 * ! При выборе сохранения без каналов, должны блокироваться все чекбоксы и выпадающий список
 * */


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;

using FirebirdSql.Data.FirebirdClient;

using Excel = Microsoft.Office.Interop.Excel;
using ExcelLibrary.BinaryDrawingFormat;
using ExcelLibrary.BinaryFileFormat;
using ExcelLibrary.SpreadSheet;
using ExcelLibrary.CompoundDocumentFormat;


namespace ConnectToSCADABD
{
    public partial class Form1 : Form
    {
        ProgramReadDB ReadDB = new ProgramReadDB();
        ProgramLoadExcel loadfile = new ProgramLoadExcel();
        ProgramSaveExcel savefile = new ProgramSaveExcel();
        string BaseAddr; //адрес базы
        public string ConStr;

        bool FirstIter; // переменная для создания sql запроса, убирающая одну запятую*/
        string SaveXlsPath; //Путь к ПАПКЕ для сохронения файла для импорта

        int preIndexCombobox; //индекс предыдущего выбранного типа, для того чтобы его сохранить при выборе следующего.
        int ObjCounter; // счетчик для замены цифры ПЛК на его название в повторном SQL запросе
        int ParamsCounter = 9; //временная переменная, указывающая кол-во параметров в канале
        bool readID; // триггер успешно прочитанного файла
        bool readBD; // триггер успешно прочитанной БД

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
            readBD = false;
            BaseAddr = textBox2.Text;
            ConStr = "character set=WIN1251;initial catalog=" + BaseAddr + ";user id=SYSDBA;password=masterkey"; // наша строка подключения, сделать её изменяемой!!!

            ReadDB.TeconObjectChannels.Clear(); // очистка предыдущего поиска
            ReadDB.TeconObjects.Clear();
            ReadDB.ObjTypeID.Clear();
            preIndexCombobox = 0;
            textBox1.AppendText( "Начат сбор данных из БД;\n");

            // textBox1.ScrollToCaret();
            int i=0;
            foreach (int ID in loadfile.ObjID)   // для каждого распознанного ID делаем SQL запрос с последующими действиями
            {
                ReadDB.AddObjChannel(ID);     //добавляем каждый канал каждого тех объекта в список
                ReadDB.AddObjDefChannel(ID);
                ReadDB.CompareChannels();   //сравниваем два списка каналов
                ReadDB.AddObj(ID);            //делаем список тех объектов, содержащий списки каналов
                i++;
                label1.Text = "Прогресс: " + i + "/" + loadfile.ObjID.Count.ToString();
            }
            ReadDB.FullKlName(); // находим полный путь для классификатора
            textBox1.AppendText( "Прочитаны данные из БД;\n");
            ReadDB.FindTypes();           //разбираем объекты на типы, ищем типы с одинаковыми каналами, присваиваем индексы
            textBox1.AppendText("Данные обработаны;\n");
            ShowData();            //показываем выбранные данные в таблице на форме
            textBox1.AppendText("Данные готовы к сохранению;\n");
            // textBox1.ScrollToCaret();
            readBD = true;
            EnabledCheck();
            ChOptionsToAll();

            button4.Enabled = true;
        }

 

        private void ShowData()
    {

//--------------------Отображение собранной из БД информации-----------------------------------------------------------------------------------------------------------------------------
           Table2.ColumnCount = 14; //без указания количества столбцов не сработает
           Table2.RowCount = ReadDB.TeconObjects.Count + 1;
           Table2.Columns[0].Name = "Nп/п";
           Table2.Columns[1].Name = "Marka";
           Table2.Columns[2].Name = "Name";
           Table2.Columns[3].Name = "Disc";
           Table2.Columns[4].Name = "ObjTypeName";
           Table2.Columns[5].Name = "Arc_Per";
           Table2.Columns[6].Name = "ObjSign";
           Table2.Columns[7].Name = "PLC_Name";
           Table2.Columns[8].Name = "PLC_GR";
           Table2.Columns[9].Name = "EVKLASSIFIKATORNAME";
           Table2.Columns[9].Name = "KKS";
           Table2.Columns[10].Name = "POUNAME";
           Table2.Columns[11].Name = "KLASSIFIKATORNAME";
           Table2.Columns[12].Name = "PLC_varname";
           Table2.Columns[13].Name = "PLC_address";


           int TmpCounter = 0;
           foreach (ProgramReadDB.TeconObject TObj in ReadDB.TeconObjects)
            {
                Table2[ 0, TmpCounter].Value = Convert.ToString(TmpCounter + 1);
                Table2[ 1, TmpCounter].Value = TObj.Marka;
                Table2[ 2, TmpCounter].Value = TObj.Name;
                Table2[ 3, TmpCounter].Value = TObj.Disc;
                Table2[ 4, TmpCounter].Value = TObj.ObjTypeName;
                Table2[ 5, TmpCounter].Value = TObj.Arc_Per;
                Table2[ 6, TmpCounter].Value = TObj.ObjSign;
                Table2[ 7, TmpCounter].Value = TObj.PLC_Name;
                Table2[ 8, TmpCounter].Value = TObj.PLC_GR;
                Table2[ 9, TmpCounter].Value = TObj.EVKLASSIFIKATORNAME;
                Table2[ 10, TmpCounter].Value = TObj.KKS;
                Table2[ 11, TmpCounter].Value = TObj.KLASSIFIKATORNAME;
                Table2[ 12, TmpCounter].Value = TObj.PLC_varname;
                Table2[ 13, TmpCounter].Value = TObj.PLC_address;
                TmpCounter++;
            }

//блок вывода статистики по собранной информации---------------------------------------------------------------------------------------------

       //    textBox1.Clear();
           textBox1.AppendText("Кол-во тех. объектов: " + ReadDB.TeconObjects.Count.ToString() + "\n\n");
           textBox1.AppendText("Кол-во типов объектов: " + ReadDB.ObjTypeIDUnic.Count.ToString() + "\n\n");
           textBox1.AppendText("Перечень типов объектов: " + "\n");
           // textBox1.ScrollToCaret();

           foreach (int i in ReadDB.ObjTypeIDUnic)
          {
              foreach (ProgramReadDB.TeconObject to in ReadDB.TeconObjects)
           {
               if (to.ObjTypeID == i)
               {
                   textBox1.AppendText(to.ObjTypeName + "\n");
                   break;
               }
           }
           
          }
           textBox1.AppendText("\n" + "Кол-во листов в файле Excel: " + ReadDB.TablesNum.ToString() + "\n");

           for (int i = 1; i <= ReadDB.TablesNum; i++)
          {
              int p = 0;
              string s = "";
              foreach (ProgramReadDB.TeconObject to in ReadDB.TeconObjects)
              {
                  if (to.Index == i) { p++; s = to.ObjTypeName; }
              }
              textBox1.AppendText("Лист" + i.ToString() + ": " + p.ToString() + ", "+s + ";\n");
              comboBox1.Items.Add("Лист" + i.ToString() + "; " + s);
               var STP = new ProgramSaveExcel.SaveTableParams()
                  {
                       TableNum = i,
                       TypeName = s,
                  };
               savefile.SaveTablesParamsList.Add(STP);
          }
          comboBox1.SelectedIndex = 0;
          //label1.Text = SaveTablesParamsList.Count.ToString();
          // textBox1.ScrollToCaret();
        }

        
          private void button3_Click(object sender, EventArgs e)
          {
            readID = false;
            //Открываем файл Экселя
            openFileDialog1.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ReadDB.SQLParams = ""; //очищаем предыдущий поиск параметров
                loadfile.LoadExcelFile(openFileDialog1.FileName);
                loadfile.ObjID.ForEach(delegate(int ID)
                {
                    if (!FirstIter)
                    {
                        ReadDB.SQLParams = ReadDB.SQLParams + ID.ToString();
                    }
                    else { ReadDB.SQLParams = ReadDB.SQLParams + "," + ID.ToString(); }
                    FirstIter = true;
                });

                textBox1.AppendText("========Новый файл=======\n");
                textBox1.AppendText("Файл открыт;\n");
                textBox1.AppendText(openFileDialog1.FileName.ToString() + ";\n");
                textBox1.AppendText("ID объектов считаны (" + loadfile.ObjID.Count.ToString() + ")шт.;\n");
                readID = true;
                readBD = false;
                
                button1.Enabled = true;
                button4.Enabled = false;
                comboBox1.Items.Clear(); // очистка списка листов
                savefile.SaveTablesParamsList.Clear();
                EnabledCheck();            
             }
          }


          private void button4_Click(object sender, EventArgs e)
          {
              //То что выбрано из параметров на момент нажатия кнопки сохранения
              int j = comboBox1.SelectedIndex;
              savefile.SaveTablesParamsList[j].S0 = checkBox1.Checked;
              savefile.SaveTablesParamsList[j].S100 = checkBox2.Checked;
              savefile.SaveTablesParamsList[j].M = checkBox3.Checked;
              savefile.SaveTablesParamsList[j].PLC_VARNAME = checkBox4.Checked;
              savefile.SaveTablesParamsList[j].ED_IZM = checkBox5.Checked;
              savefile.SaveTablesParamsList[j].ARH_APP = checkBox6.Checked;
              savefile.SaveTablesParamsList[j].DISC = checkBox7.Checked;
              savefile.SaveTablesParamsList[j].KA = checkBox8.Checked;
              savefile.SaveTablesParamsList[j].KB = checkBox9.Checked;

              savefile.Sheets.Clear();
              bool SaveWithoutCh = false;

              //saveFileDialog1.Filter = "Excel files (*.xls;*.xlsx)|*.xlsx;*.xls";
              saveFileDialog1.Filter = "Excel files (*.xls)|*.xls";
              var culture = new CultureInfo("ru-RU");
              string name = "";
              if (checkBox11.Checked) 
              {
                  name = "ТеконИмпорт_Листов_1;_Объектов_" + ReadDB.TeconObjects.Count.ToString() + ";_" + DateTime.Now.ToString(culture);
              } else 
              {
                  name = "ТеконИмпорт_Листов_" + ReadDB.TablesNum.ToString() + ";_Объектов_" + ReadDB.TeconObjects.Count.ToString() + ";_" + DateTime.Now.ToString(culture);
              }
              name = name.Replace(":", "_");
              saveFileDialog1.FileName = name;

              if (saveFileDialog1.ShowDialog() == DialogResult.OK)
              {
                  textBox1.AppendText("Начат процесс сохранения;\n");
                  savefile.SaveFileExcel(saveFileDialog1.FileName, checkBox11.Checked, ReadDB.TablesNum, ReadDB.TeconObjects);
                  textBox1.AppendText("Сохранение завершено;\n");
              }
          }

          private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
          {
              int i = comboBox1.SelectedIndex;
              groupBox2.Text = "Лист" + savefile.SaveTablesParamsList[i].TableNum.ToString() + " Тип: " + savefile.SaveTablesParamsList[i].TypeName;

              savefile.SaveTablesParamsList[preIndexCombobox].S0 = checkBox1.Checked;
              savefile.SaveTablesParamsList[preIndexCombobox].S100 = checkBox2.Checked;
              savefile.SaveTablesParamsList[preIndexCombobox].M = checkBox3.Checked;
              savefile.SaveTablesParamsList[preIndexCombobox].PLC_VARNAME = checkBox4.Checked;
              savefile.SaveTablesParamsList[preIndexCombobox].ED_IZM = checkBox5.Checked;
              savefile.SaveTablesParamsList[preIndexCombobox].ARH_APP = checkBox6.Checked;
              savefile.SaveTablesParamsList[preIndexCombobox].DISC = checkBox7.Checked;
              savefile.SaveTablesParamsList[preIndexCombobox].KA = checkBox8.Checked;
              savefile.SaveTablesParamsList[preIndexCombobox].KB = checkBox9.Checked;
              
              checkBox1.Checked = savefile.SaveTablesParamsList[i].S0;
              checkBox2.Checked = savefile.SaveTablesParamsList[i].S100;
              checkBox3.Checked = savefile.SaveTablesParamsList[i].M;
              checkBox4.Checked = savefile.SaveTablesParamsList[i].PLC_VARNAME;
              checkBox5.Checked = savefile.SaveTablesParamsList[i].ED_IZM;
              checkBox6.Checked = savefile.SaveTablesParamsList[i].ARH_APP;
              checkBox7.Checked = savefile.SaveTablesParamsList[i].DISC;
              checkBox8.Checked = savefile.SaveTablesParamsList[i].KA;
              checkBox9.Checked = savefile.SaveTablesParamsList[i].KB;

              preIndexCombobox = i; 
          }

          private void checkBox11_CheckedChanged(object sender, EventArgs e)
          {
              EnabledCheck();
          }

          private void checkBox10_CheckedChanged(object sender, EventArgs e)
          {
              EnabledCheck();
          }

          private void EnabledCheck()
          {
              if (checkBox11.Checked || !readBD || !readID)
              {
                  checkBox1.Enabled = false;
                  checkBox2.Enabled = false;
                  checkBox3.Enabled = false;
                  checkBox4.Enabled = false;
                  checkBox5.Enabled = false;
                  checkBox6.Enabled = false;
                  checkBox7.Enabled = false;
                  checkBox8.Enabled = false;
                  checkBox9.Enabled = false;
                 // checkBox10.Enabled = false;
              }

              if (!checkBox11.Checked &&  readBD && readID)
              {
                  checkBox1.Enabled = true;
                  checkBox2.Enabled = true;
                  checkBox3.Enabled = true;
                  checkBox4.Enabled = true;
                  checkBox5.Enabled = true;
                  checkBox6.Enabled = true;
                  checkBox7.Enabled = true;
                  checkBox8.Enabled = true;
                  checkBox9.Enabled = true;
                 // checkBox10.Enabled = true;
              }

              if (readBD && readID)
              {
                  checkBox11.Enabled = true;
                  comboBox1.Enabled = true;
              }
              else { checkBox11.Enabled = false; comboBox1.Enabled = false; }
          }

          private void ChOptionsToAll()
          {
              for (int j = 0; j < comboBox1.Items.Count; j++)
              {
                  savefile.SaveTablesParamsList[j].S0 = checkBox1.Checked;
                  savefile.SaveTablesParamsList[j].S100 = checkBox2.Checked;
                  savefile.SaveTablesParamsList[j].M = checkBox3.Checked;
                  savefile.SaveTablesParamsList[j].PLC_VARNAME = checkBox4.Checked;
                  savefile.SaveTablesParamsList[j].ED_IZM = checkBox5.Checked;
                  savefile.SaveTablesParamsList[j].ARH_APP = checkBox6.Checked;
                  savefile.SaveTablesParamsList[j].DISC = checkBox7.Checked;
                  savefile.SaveTablesParamsList[j].KA = checkBox8.Checked;
                  savefile.SaveTablesParamsList[j].KB = checkBox9.Checked;
              }
          }

          private void checkBox1_CheckedChanged(object sender, EventArgs e)
          {
               }

          private void Form1_Load(object sender, EventArgs e)
          {

          }

          private void button2_Click(object sender, EventArgs e)
          {
              ChOptionsToAll();
          }

          private void button5_Click(object sender, EventArgs e)
          {
              FormInitVal f2 = new FormInitVal();
              

              foreach (int ID in loadfile.ObjID)   // для каждого распознанного ID делаем SQL запрос с последующими действиями
              {
                  ReadDB.AddInitValue(ID);
                //  MessageBox.Show("Прошел запрос " + ID.ToString());
              }

              foreach (ProgramReadDB.InitValue iv in ReadDB.InitValues)
              {
                  f2.InitValues.Add(iv);
              }
              f2.Show();   

          }
      }
}

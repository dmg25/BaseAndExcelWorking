﻿/* Добавить:
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
using System.Threading;

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
        public string BaseAddr; //адрес базы
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

            ConStr = "character set=WIN1251;initial catalog=" + BaseAddr + ";user id=SYSDBA;password=masterkey"; // наша строка подключения, в данный момент не активна

            ReadDB.TeconObjectChannels.Clear(); // очистка предыдущего поиска
            ReadDB.TeconObjects.Clear();
            ReadDB.ObjTypeID.Clear();
            ReadDB.ObjTypeChannelsMatched.Clear();
            ReadDB.SQL_Channels.Clear();
            ReadDB.SQL_DefChannels.Clear();
            ReadDB.SQL_Objects.Clear();

            preIndexCombobox = 0;
            textBox1.AppendText("Начат сбор данных из БД;\n");

            // textBox1.ScrollToCaret();
            FormProcess f = new FormProcess();
            f.progressBar1.Minimum = 0;
            f.progressBar1.Maximum = loadfile.ObjID.Count;
            f.Location = new Point(170, 400);
            f.Show();
            Thread.Sleep(10); 
            Enabled = false; // относится к форме


            //перенос ID в новый лист.
            ReadDB.Big_SQL(loadfile.ObjID);

            int i = 0;

            foreach (int ID in loadfile.ObjID)   // для каждого распознанного ID делаем SQL запрос с последующими действиями
            {
                ReadDB.AddObjChannel(ID);     //добавляем каждый канал каждого тех объекта в список
                ReadDB.AddObjDefChannel(ID);  //затем читаем дефолтные свойства каналов
                ReadDB.CompareChannels();   //сравниваем два списка каналов
                ReadDB.AddObj(ID);            //делаем список тех объектов, содержащий списки каналов
                i++;
                f.progressBar1.Value = i;
                // label1.Text = "Прогресс: " + i + "/" + loadfile.ObjID.Count.ToString();
            }
            f.Close();
            Enabled = true;
            TopMost = true;
            TopMost = false; // вытаскиваем на передний план

            ReadDB.FullKlName(); // находим полный путь для классификатора
            textBox1.AppendText("Прочитаны данные из БД;\n");
            ReadDB.FindTypes();           //разбираем объекты на типы, ищем типы с одинаковыми каналами, присваиваем индексы
            ReadDB.RepeatChNameAlarm(); //ищем повторяющиеся имена каналов у типов.
            textBox1.AppendText("Данные обработаны;\n");
            ShowData();            //показываем выбранные данные в таблице на форме
            textBox1.AppendText("Данные готовы к сохранению;\n");
            // textBox1.ScrollToCaret();
            readBD = true;
        //    EnabledCheck();
            ChOptionsToAll(); //???

            button1.BackColor = Color.LightGreen;
            //разблокировать кнопки дальше               
            button4.Enabled = true; button4.BackColor = SystemColors.Control;
            button5.Enabled = true; button5.BackColor = SystemColors.Control; 

            //блок вывода sql листов на форму
            dataGridView1.ColumnCount = 12;
            dataGridView1.Columns[0].Name = "S0";
            dataGridView1.Columns[1].Name = "S100";
            dataGridView1.Columns[2].Name = "M";
            dataGridView1.Columns[3].Name = "PLC_VARNAME";
            dataGridView1.Columns[4].Name = "ED_IZM";
            dataGridView1.Columns[5].Name = "ARH_APP";
            dataGridView1.Columns[6].Name = "DISC";
            dataGridView1.Columns[7].Name = "KA";
            dataGridView1.Columns[8].Name = "KB";
            dataGridView1.Columns[9].Name = "ID";
            dataGridView1.Columns[10].Name = "ChannelName";
            dataGridView1.Columns[11].Name = "LoadID";

            for (int p = 0; p<ReadDB.SQL_Channels.Count; p++)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[p].Cells[0].Value = ReadDB.SQL_Channels[p].S0;
                dataGridView1.Rows[p].Cells[1].Value = ReadDB.SQL_Channels[p].S100;
                dataGridView1.Rows[p].Cells[2].Value = ReadDB.SQL_Channels[p].M;
                dataGridView1.Rows[p].Cells[3].Value = ReadDB.SQL_Channels[p].PLC_VARNAME;
                dataGridView1.Rows[p].Cells[4].Value = ReadDB.SQL_Channels[p].ED_IZM;
                dataGridView1.Rows[p].Cells[5].Value = ReadDB.SQL_Channels[p].ARH_APP;
                dataGridView1.Rows[p].Cells[6].Value = ReadDB.SQL_Channels[p].DISC;
                dataGridView1.Rows[p].Cells[7].Value = ReadDB.SQL_Channels[p].KA;
                dataGridView1.Rows[p].Cells[8].Value = ReadDB.SQL_Channels[p].KB;
                dataGridView1.Rows[p].Cells[9].Value = ReadDB.SQL_Channels[p].ID;
                dataGridView1.Rows[p].Cells[10].Value = ReadDB.SQL_Channels[p].ChannelName;
                dataGridView1.Rows[p].Cells[11].Value = ReadDB.SQL_Channels[p].LoadID;                  
            }
            label1.Text = ReadDB.SQL_Channels.Count.ToString();

            dataGridView2.ColumnCount = 12;
            dataGridView2.Columns[0].Name = "S0";
            dataGridView2.Columns[1].Name = "S100";
            dataGridView2.Columns[2].Name = "M";
            dataGridView2.Columns[3].Name = "PLC_VARNAME";
            dataGridView2.Columns[4].Name = "ED_IZM";
            dataGridView2.Columns[5].Name = "ARH_APP";
            dataGridView2.Columns[6].Name = "DISC";
            dataGridView2.Columns[7].Name = "KA";
            dataGridView2.Columns[8].Name = "KB";
            dataGridView2.Columns[9].Name = "ID";
            dataGridView2.Columns[10].Name = "ChannelName";
            dataGridView2.Columns[11].Name = "LoadID";

            for (int p = 0; p < ReadDB.SQL_DefChannels.Count; p++)
            {
                dataGridView2.Rows.Add();
                dataGridView2.Rows[p].Cells[0].Value = ReadDB.SQL_DefChannels[p].S0;
                dataGridView2.Rows[p].Cells[1].Value = ReadDB.SQL_DefChannels[p].S100;
                dataGridView2.Rows[p].Cells[2].Value = ReadDB.SQL_DefChannels[p].M;
                dataGridView2.Rows[p].Cells[3].Value = ReadDB.SQL_DefChannels[p].PLC_VARNAME;
                dataGridView2.Rows[p].Cells[4].Value = ReadDB.SQL_DefChannels[p].ED_IZM;
                dataGridView2.Rows[p].Cells[5].Value = ReadDB.SQL_DefChannels[p].ARH_APP;
                dataGridView2.Rows[p].Cells[6].Value = ReadDB.SQL_DefChannels[p].DISC;
                dataGridView2.Rows[p].Cells[7].Value = ReadDB.SQL_DefChannels[p].KA;
                dataGridView2.Rows[p].Cells[8].Value = ReadDB.SQL_DefChannels[p].KB;
                dataGridView2.Rows[p].Cells[9].Value = ReadDB.SQL_DefChannels[p].ID;
                dataGridView2.Rows[p].Cells[10].Value = ReadDB.SQL_DefChannels[p].ChannelName;
                dataGridView2.Rows[p].Cells[11].Value = ReadDB.SQL_DefChannels[p].LoadID;
            }
            label2.Text = ReadDB.SQL_DefChannels.Count.ToString();

            dataGridView3.ColumnCount = 17;
            dataGridView3.Columns[0].Name = "Marka";
            dataGridView3.Columns[1].Name = "Name";
            dataGridView3.Columns[2].Name = "Disc";
            dataGridView3.Columns[3].Name = "ObjTypeName";
            dataGridView3.Columns[4].Name = "Arc_Per";
            dataGridView3.Columns[5].Name = "ObjSign";
            dataGridView3.Columns[6].Name = "PLC_Name";
            dataGridView3.Columns[7].Name = "PLC_GR";
            dataGridView3.Columns[8].Name = "EVKLASSIFIKATORNAME";
            dataGridView3.Columns[9].Name = "KKS";
            dataGridView3.Columns[10].Name = "POUNAME";
            dataGridView3.Columns[11].Name = "KLASSIFIKATORNAME";
            dataGridView3.Columns[12].Name = "PLC_varname";
            dataGridView3.Columns[13].Name = "PLC_address";
            dataGridView3.Columns[14].Name = "ObjTypeID";
            dataGridView3.Columns[15].Name = "LoadID";


            for (int p = 0; p < ReadDB.SQL_Objects.Count; p++)
            {
                dataGridView3.Rows.Add();
                dataGridView3.Rows[p].Cells[0].Value = ReadDB.SQL_Objects[p].Marka;
                dataGridView3.Rows[p].Cells[1].Value = ReadDB.SQL_Objects[p].Name;
                dataGridView3.Rows[p].Cells[2].Value = ReadDB.SQL_Objects[p].Disc;
                dataGridView3.Rows[p].Cells[3].Value = ReadDB.SQL_Objects[p].ObjTypeName;
                dataGridView3.Rows[p].Cells[4].Value = ReadDB.SQL_Objects[p].Arc_Per;
                dataGridView3.Rows[p].Cells[5].Value = ReadDB.SQL_Objects[p].ObjSign;
                dataGridView3.Rows[p].Cells[6].Value = ReadDB.SQL_Objects[p].PLC_Name;
                dataGridView3.Rows[p].Cells[7].Value = ReadDB.SQL_Objects[p].PLC_GR;
                dataGridView3.Rows[p].Cells[8].Value = ReadDB.SQL_Objects[p].EVKLASSIFIKATORNAME;
                dataGridView3.Rows[p].Cells[9].Value = ReadDB.SQL_Objects[p].KKS;
                dataGridView3.Rows[p].Cells[10].Value = ReadDB.SQL_Objects[p].POUNAME;
                dataGridView3.Rows[p].Cells[11].Value = ReadDB.SQL_Objects[p].KLASSIFIKATORNAME;
                dataGridView3.Rows[p].Cells[12].Value = ReadDB.SQL_Objects[p].PLC_varname;
                dataGridView3.Rows[p].Cells[13].Value = ReadDB.SQL_Objects[p].PLC_address;
                dataGridView3.Rows[p].Cells[14].Value = ReadDB.SQL_Objects[p].ObjTypeID;
                dataGridView3.Rows[p].Cells[15].Value = ReadDB.SQL_Objects[p].LoadID;
            }
            label3.Text = ReadDB.SQL_Objects.Count.ToString();
            //-------------------------------

          
        }

        private void ShowData()
    {
/*
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
            }*/

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
             // comboBox1.Items.Add("Лист" + i.ToString() + "; " + s);
               var STP = new ProgramSaveExcel.SaveTableParams()
                  {
                       TableNum = i,
                       TypeName = s,
                  };
               savefile.SaveTablesParamsList.Add(STP);
          }
         // comboBox1.SelectedIndex = 0;
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

                button3.BackColor = Color.LightGreen;
                //разблокировать кнопки дальше               
                button1.Enabled = true; button1.BackColor = SystemColors.Control;
                button4.Enabled = false; button4.BackColor = SystemColors.Control;
                button5.Enabled = false; button5.BackColor = SystemColors.Control; 
                

            //    comboBox1.Items.Clear(); // очистка списка листов
                savefile.SaveTablesParamsList.Clear();
           //     EnabledCheck();            
             }
          }


          private void button4_Click(object sender, EventArgs e)
          {
              //То что выбрано из параметров на момент нажатия кнопки сохранения
             /* int j = comboBox1.SelectedIndex;
              savefile.SaveTablesParamsList[j].S0 = checkBox1.Checked;
              savefile.SaveTablesParamsList[j].S100 = checkBox2.Checked;
              savefile.SaveTablesParamsList[j].M = checkBox3.Checked;
              savefile.SaveTablesParamsList[j].PLC_VARNAME = checkBox4.Checked;
              savefile.SaveTablesParamsList[j].ED_IZM = checkBox5.Checked;
              savefile.SaveTablesParamsList[j].ARH_APP = checkBox6.Checked;
              savefile.SaveTablesParamsList[j].DISC = checkBox7.Checked;
              savefile.SaveTablesParamsList[j].KA = checkBox8.Checked;
              savefile.SaveTablesParamsList[j].KB = checkBox9.Checked;*/

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
                  button4.BackColor = Color.LightGreen;
              }

              
             
          }

          private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
          {
             /* int i = comboBox1.SelectedIndex;
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

              preIndexCombobox = i; */
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
             /* if (checkBox11.Checked || !readBD || !readID)
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
              else { checkBox11.Enabled = false; comboBox1.Enabled = false; }*/
          }

          private void ChOptionsToAll()
          {
             /* for (int j = 0; j < comboBox1.Items.Count; j++)
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
              }*/
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
              f2.BaseAddr = BaseAddr;
              ReadDB.InitValues.Clear();
              f2.InitValues.Clear();

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

              button5.BackColor = Color.LightGreen;
             

          }

        private void button6_Click(object sender, EventArgs e)
        {
            openDB.Filter = "Firebird DB files (*.GDB)|*.GDB";
            if (openDB.ShowDialog() == DialogResult.OK)
            {
                BaseAddr = openDB.FileName;
                //закрасить цветом кнопку
                button6.BackColor = Color.LightGreen;
                //разблокировать кнопки дальше               
                button3.Enabled = true;  button3.BackColor = SystemColors.Control;
                button1.Enabled = false; button1.BackColor = SystemColors.Control;
                button4.Enabled = false; button4.BackColor = SystemColors.Control;
                button5.Enabled = false; button5.BackColor = SystemColors.Control;
                ReadDB.BaseAddr = BaseAddr; //кидаем адрес базы в readBase

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {

        }
      }
}

            
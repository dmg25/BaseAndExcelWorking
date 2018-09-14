/* Добавить:
 * 1.+ Диагностика неправильного файла экспорта (сообщение и возврат к началу)
 * 2.+ Блокировку кнопок, чтобы не сделать один этап раньше момента
 * 3.+ Добавить нормальный путь для сохранения файла и прописывать ему автоимя
 * 4. Сделать раскрывающийся список с ЛИСТАМИ и под каждый (приписывая тип объекта),
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


namespace ConnectToSCADABD
{
    public partial class Form1 : Form
    {
        System.Data.DataTable TmpDG = new System.Data.DataTable(); // переменная, чтобы не создавать новые таблицы при новом зарпосе
        System.Data.DataTable dt1 = new System.Data.DataTable(); // таблица с данными из БД

     //   string SQL; // "select * from OBJTYPE "; //Where Speed=@Speed
        string SQLParams; //часть запроса, содержащая перечень нужных строк
        string SQL_CARDS; //запрос основных параметров объекта
        string SQL_CARDPARAMS; //запрос параметров каналов объекта
        string SQL_OBJTYPES; //запрос параметров каналов объекта
        string BaseAddr; //адрес базы

        string ConStr;

        string CellStr; //содержимое ячейки Excel
        int ExcelCellCnt; // счетчик ячеек Excel
        int CellInt; // переведённое в int содержимое ячейки
        bool FirstIter; // переменная для создания sql запроса, убирающая одну запятую
        string SaveXlsPath; //Путь к ПАПКЕ для сохронения файла для импорта

        int ObjCounter; // счетчик для замены цифры ПЛК на его название в повторном SQL запросе
        int ParamsCounter = 9; //временная переменная, указывающая кол-во параметров в канале
        List<int> ObjID = new List<int>();  // массив с прочитанными ID из Excel
        List<int> ObjTypeID = new List<int>(); //массив типов объектов
        List<int> ObjTypeIDUnic = new List<int>(); //массив типов объектов без повторений
        List<string> ObjTypeChannelsList = new List<string>(); //список каналов типа объекта

        int TablesNum;
        bool readdenID; // триггер успешно прочитанного файла


//checklist для выбора параметров канала--------
        bool bS0;
        bool bS100;
        bool bM;
        bool bPLC_VARNAME;
        bool bED_IZM;
        bool bARH_APP;
        bool bDISC;
        bool bKA;
        bool bKB;
//--------------------------------------------------------------------------------------------------------------

        class TeconObject   // класс с описанием основных параметров техобъекта
        {
            public string Marka;
            public string Name ;
            public string Disc ;
            public string ObjTypeName ;     
            public string Arc_Per ;
            public string ObjSign ;
            public string PLC_Name ;
            public string PLC_GR ;
            public string EVKLASSIFIKATORNAME ;  
            public string KKS ;
            public string POUNAME ;   
            public string KLASSIFIKATORNAME ;  
            public string PLC_varname;  
            public string PLC_address;
            public List<TeconObjectChannel> Channels = new List<TeconObjectChannel>();

            public int ObjTypeID;
            public int Index;
        }

        class TeconObjectChannel   // класс с описанием параметров канала объекта
        {
            public string S0;
            public string S100;
            public string M;
            public string PLC_VARNAME;
            public string ED_IZM;
            public string ARH_APP;
            public string DISC;
            public string KA;
            public string KB;
            public string ID;

            public string ChannelName;
        }

        class ObjTypeChannels   // класс с перечнем каналов каждого типа
        {
            public int TypeID;
            public int Index; //индекс совпадения с другими типами
            public List<string> Channels = new List<string>();
        }

        class SaveTableParams       // класс для постраничного выбора параметров каналов
        {
            public int TableNum;
            public string TypeName;

            public bool S0;
            public bool S100;
            public bool M;
            public bool PLC_VARNAME;
            public bool ED_IZM;
            public bool ARH_APP;
            public bool DISC;
            public bool KA;
            public bool KB;
         
        }

        List<TeconObject> TeconObjects = new List<TeconObject>();
        List<TeconObjectChannel> TeconObjectChannels = new List<TeconObjectChannel>();
        List<ObjTypeChannels> ObjTypeCh = new List<ObjTypeChannels>();
        List<SaveTableParams> SaveTablesParamsList = new List<SaveTableParams>();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            BaseAddr = textBox2.Text;
            ConStr = "character set=WIN1251;initial catalog=" + BaseAddr + ";user id=SYSDBA;password=masterkey"; // наша строка подключения, сделать её изменяемой!!!
            TeconObjectChannels.Clear(); // очистка предыдущего поиска
            TeconObjects.Clear();
            textBox1.AppendText( "Начат сбор данных из БД;\n");
            // textBox1.ScrollToCaret();
            foreach (int ID in ObjID)   // для каждого распознанного ID делаем SQL запрос с последующими действиями
            {
                AddObjChannel(ID);     //добавляем каждый канал каждого тех объекта в список
                AddObj(ID);            //делаем список тех объектов, содержащий списки каналов
            }
            textBox1.AppendText( "Прочитаны данные из БД;\n");
            FindTypes();           //разбираем объекты на типы, ищем типы с одинаковыми каналами, присваиваем индексы
            textBox1.AppendText("Данные обработаны;\n");
            ShowData();            //показываем выбранные данные в таблице на форме
            textBox1.AppendText("Данные готовы к сохранению;\n");
            // textBox1.ScrollToCaret();
            button4.Enabled = true;
        }



        private void ConnectToBase(string SQL)
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

                    dt1 = ds.Tables[0];

                    if (dt1.Rows.Count == 0) return;

                    //Table3.DataSource = dt1;
                }
        }


        private void AddObjChannel(int ID)
        {
            SQL_CARDPARAMS = "Select CARDPARAMS.S0, CARDPARAMS.S100, CARDPARAMS.M, CARDPARAMS.PLC_VARNAME, CARDPARAMS.ED_IZM, CARDPARAMS.ARH_APP, CARDPARAMS.DISC, CARDPARAMS.KA, CARDPARAMS.KB, CARDPARAMS.OBJTYPEPARAMID from CARDPARAMS where CARDPARAMS.CARDID = " + ID;
            ConnectToBase(SQL_CARDPARAMS);
            TmpDG = dt1;

            for (int i = 0; i < TmpDG.Rows.Count; i++)
            {
                
                var objChnl = new TeconObjectChannel()
                {
                    S0 = TmpDG.Rows[i][0].ToString(),
                    S100 = TmpDG.Rows[i][1].ToString(),
                    M = TmpDG.Rows[i][2].ToString(),
                    PLC_VARNAME = TmpDG.Rows[i][3].ToString(),
                    ED_IZM = TmpDG.Rows[i][4].ToString(),
                    ARH_APP = TmpDG.Rows[i][5].ToString(),
                    DISC = TmpDG.Rows[i][6].ToString(),
                    KA = TmpDG.Rows[i][7].ToString(),
                    KB = TmpDG.Rows[i][8].ToString(),
                    ID = TmpDG.Rows[i][9].ToString(),
                };

                // Делаем короткий запрос для получения имени контроллера, ибо в едином запросе такое хз как сделать
                ConnectToBase("Select OBJTYPEPARAM.NAME from OBJTYPEPARAM where OBJTYPEPARAM.ID = " + objChnl.ID);
                objChnl.ChannelName = dt1.Rows[0][0].ToString();
             
                TeconObjectChannels.Add(objChnl);
            }
        }

        private void AddObj(int ID)
        {

            SQL_CARDS = "Select CARDS.MARKA, CARDS.NAME, CARDS.DISC, OBJTYPE.NAME, CARDS.ARH_PER, CARDS.OBJSIGN, CARDS.PLC_ID, CARDS.PLC_GR, EVKLASSIFIKATOR.NAME, CARDS.KKS, ISAOBJ.NAME, KLASSIFIKATOR.NAME, CARDS.PLC_VARNAME, CARDS.PLC_ADRESS, CARDS.OBJTYPEID from CARDS, OBJTYPE, KLASSIFIKATOR, EVKLASSIFIKATOR, ISAOBJ, RESOURCES where CARDS.ID = " + ID + " and CARDS.OBJTYPEID = OBJTYPE.ID and CARDS.EVKLID = EVKLASSIFIKATOR.ID and CARDS.TEMPLATEID = ISAOBJ.ID and CARDS.KLID = KLASSIFIKATOR.ID";
            ConnectToBase(SQL_CARDS);

//---------------------------Заполняем параметры объекта---------------------------------------------------------------                 
            var obj = new TeconObject()
            {
                Marka = dt1.Rows[0][0].ToString(),
                Name = dt1.Rows[0][1].ToString(),
                Disc = dt1.Rows[0][2].ToString(),
                ObjTypeName = dt1.Rows[0][3].ToString(),
                Arc_Per = dt1.Rows[0][4].ToString(),
                ObjSign = dt1.Rows[0][5].ToString(),
                PLC_Name = dt1.Rows[0][6].ToString(),
                PLC_GR = dt1.Rows[0][7].ToString(),
                EVKLASSIFIKATORNAME = dt1.Rows[0][8].ToString(),
                KKS = dt1.Rows[0][9].ToString(),
                POUNAME = dt1.Rows[0][10].ToString(),
                KLASSIFIKATORNAME = dt1.Rows[0][11].ToString(),
                PLC_varname = dt1.Rows[0][12].ToString(),
                PLC_address = dt1.Rows[0][13].ToString(),
                Channels = TeconObjectChannels.ToList(),

                ObjTypeID = Convert.ToInt16(dt1.Rows[0][14]),
            };


            ObjTypeID.Add(Convert.ToInt16(dt1.Rows[0][14])); // заполняем список типами объектов

            // Делаем короткий запрос для получения имени контроллера, ибо в едином запросе такое хз как сделать
            ConnectToBase("Select CARDS.MARKA from CARDS where CARDS.ID = " + obj.PLC_Name);
            obj.PLC_Name = dt1.Rows[0][0].ToString();

            TeconObjects.Add(obj);
            TeconObjectChannels.Clear();
        }

        private void FindTypes()
        {
        ObjTypeIDUnic = ObjTypeID.Distinct().ToList(); //убираем повторяющиеся типы объектов

       // TablesNum = 0;
        ObjTypeCh.Clear(); // для повторной активации функции

        foreach (int id in ObjTypeIDUnic)
        {
            SQL_OBJTYPES = "Select OBJTYPEPARAM.NAME from OBJTYPEPARAM where OBJTYPEPARAM.PID = " + id;
            ConnectToBase(SQL_OBJTYPES);
            

           // textBox1.Clear();

            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                ObjTypeChannelsList.Add(dt1.Rows[i][0].ToString());
            }

            var ObjTypeCh1 = new ObjTypeChannels()
            {
                TypeID = id,
                Channels = ObjTypeChannelsList.ToList(),
                Index = 0,  // 0 - значит индекс еще не заполнен
            };

            ObjTypeCh.Add(ObjTypeCh1); //заполняем список с каналами типа объекта
        }

        
        // выявление одинаковых каналов у типов и присваивание им индексов
       int k = 1;
        for (int i = 0; i <= ObjTypeCh.Count-1; i++)
        {
            for (int j = 0; j <= ObjTypeCh.Count-1; j++)
            {
                if ((ObjTypeCh[i].Channels.SequenceEqual(ObjTypeCh[j].Channels)) && (i != j))
                {
                    if (ObjTypeCh[i].Index != 0)
                    {
                       // MessageBox.Show("есть совпадение");
                        ObjTypeCh[j].Index = ObjTypeCh[i].Index;
                    }
                    else
                    {
                        ObjTypeCh[i].Index = k;
                        ObjTypeCh[j].Index = ObjTypeCh[i].Index;
                        k++;
                    }
                }
            }
        }

         

       // Присваиваем оставшимся уникальным типам (нулевые индексы) индексы больше, чем у неуникальных
        for (int i = 0; i <= ObjTypeCh.Count-1; i++)
        {
            if (ObjTypeCh[i].Index == 0)
            {
                ObjTypeCh[i].Index = k;
                k++;
            }
        }

        // Присваиваем индексы каждому объекту, согласно его типу
        for (int i = 0; i <= TeconObjects.Count-1; i++)
        {
            for (int j = 0; j <= ObjTypeCh.Count-1; j++)
            {
                if (TeconObjects[i].ObjTypeID == ObjTypeCh[j].TypeID)
                {
                    TeconObjects[i].Index = ObjTypeCh[j].Index;
                    break;
                }
                else if (j == ObjTypeCh.Count)
                {
                    MessageBox.Show("Для объекта: " + TeconObjects[i].Name + "; не найден тип!");
                }
            }
        }

        TablesNum = k-1; //присваиваем количество таблиц

    /*    textBox1.Clear();
            foreach (TeconObject i in TeconObjects)
            {
                textBox1.AppendText(i.Index.ToString() + "\n";
            }*/
           

         }

        private void ShowData()
    {
//--------------------Отображение собранной из БД информации-----------------------------------------------------------------------------------------------------------------------------
           Table2.ColumnCount = 14; //без указания количества столбцов не сработает
           Table2.RowCount = TeconObjects.Count + 1;
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
           foreach (TeconObject TObj in TeconObjects)
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
           textBox1.AppendText("Кол-во тех. объектов: " + TeconObjects.Count.ToString() + "\n\n");
           textBox1.AppendText("Кол-во типов объектов: " + ObjTypeIDUnic.Count.ToString() + "\n\n");
           textBox1.AppendText("Перечень типов объектов: " + "\n");
           // textBox1.ScrollToCaret();

          foreach (int i in ObjTypeIDUnic)
          {
            foreach (TeconObject to in TeconObjects)
           {
               if (to.ObjTypeID == i)
               {
                   textBox1.AppendText(to.ObjTypeName + "\n");
                   break;
               }
           }
           
          }
          textBox1.AppendText("\n" + "Кол-во листов в файле Excel: " + TablesNum.ToString() + "\n");

          for (int i = 1; i <= TablesNum; i++)
          {
              int p = 0;
              string s = "";
              foreach (TeconObject to in TeconObjects)
              {
                  if (to.Index == i) { p++; s = to.ObjTypeName; }
              }
              textBox1.AppendText("Лист"+i.ToString() +": " + p.ToString() + ", "+s + ";\n");
              comboBox1.Items.Add("Лист" + i.ToString() + "; " + s);
               var STP = new SaveTableParams()
                  {
                       TableNum = i,
                       TypeName = s,
                  };
               SaveTablesParamsList.Add(STP);
          }
          comboBox1.SelectedIndex = 0;
          //label1.Text = SaveTablesParamsList.Count.ToString();
          // textBox1.ScrollToCaret();
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        }


//-----------------------------------------------------------------------------------------------------------------------------------------------------------------      

        
          private void button3_Click(object sender, EventArgs e)
        {
            //Открываем файл Экселя
            openFileDialog1.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //Создаём приложение.
               Excel.Application ObjExcel = new Excel.Application();
                //Открываем книгу.                                                                                                                                                        
               Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false,Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Выбираем таблицу(лист).
               Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

                SQLParams = ""; //очищаем предыдущий поиск параметров
                ObjID.Clear();
                ExcelCellCnt = 0;
                FirstIter = false;
                CellStr = "1";
               // MessageBox.Show(CellStr);

                Excel.Range forYach = ObjWorkSheet.Cells[4, 3] as Excel.Range;
                if (forYach.Value2.ToString() != "ID")
                {
                    MessageBox.Show("Выбран некорректный файл экспорта: Найдено: "+ forYach.Value2.ToString() + " вместо ID!" );
                    return;
                }

                while (CellStr != "")  //будем читать столбец, пока не найдем пустую ячейку
                {
                    
                   Excel.Range range = ObjWorkSheet.get_Range("C" + (ExcelCellCnt + 5).ToString());
                    CellStr = range.Text.ToString();
                    if (CellStr != "") { CellInt = Convert.ToInt16(CellStr); ObjID.Add(CellInt);} //Убрать потом как-нибудь это условие, чтобы не дублировалось                
                    ExcelCellCnt++;
                }

                ObjID.ForEach(delegate(int ID)
                {
                    if (!FirstIter)
                    {
                        SQLParams = SQLParams + ID.ToString();
                    }
                    else { SQLParams = SQLParams + "," + ID.ToString(); }
                    FirstIter = true;
                });

                //Удаляем приложение (выходим из экселя) - будет висеть в процессах!
                ObjExcel.Quit();

                //Очищаем от старого текста окно вывода.
                textBox1.Clear();
                textBox1.Text = "Файл открыт;\n";
                textBox1.AppendText("ID объектов считаны (" + ObjID.Count.ToString()+ ")шт.;\n");
                // textBox1.ScrollToCaret();
               // readdenID = true;
                button1.Enabled = true;
                button4.Enabled = false;
            }
        }


          private void button4_Click(object sender, EventArgs e)
          {

              /* if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
               {
                   SaveXlsPath = folderBrowserDialog1.SelectedPath;
               }*/

              bool SaveWithoutCh = false;

              saveFileDialog1.Filter = "Excel files (*.xls;*.xlsx)|*.xlsx;*.xls";
              var culture = new CultureInfo("ru-RU");
              string name =  "ТеконИмпорт_Листов_" + TablesNum.ToString() + ";_Объектов_" + TeconObjects.Count.ToString() + ";_" + DateTime.Now.ToString(culture);
              name = name.Replace(":", "_");
              saveFileDialog1.FileName = name;

              if (saveFileDialog1.ShowDialog() == DialogResult.OK)
              {
                  SaveXlsPath = saveFileDialog1.FileName;


                  textBox1.AppendText("Начат процесс сохранения;\n");
                  // textBox1.ScrollToCaret();
                  //Объявляем приложение
                  Excel.Application ex = new Excel.Application();
                  //Отобразить Excel
                  //   ex.Visible = true;


                  //Проверяем чеклист
                  bS0 = checkBox1.Checked;
                  bS100 = checkBox2.Checked;
                  bM = checkBox3.Checked;
                  bPLC_VARNAME = checkBox4.Checked;
                  bED_IZM = checkBox5.Checked;
                  bARH_APP = checkBox6.Checked;
                  bDISC = checkBox7.Checked;
                  bKA = checkBox8.Checked;
                  bKB = checkBox9.Checked;
                  SaveWithoutCh = checkBox11.Checked;

                  if ((!bS0 && !bS100 && !bM && !bPLC_VARNAME && !bED_IZM && !bARH_APP && !bDISC && !bKA && !bKB && !checkBox11.Checked))
                  {

                      DialogResult dialogResult = MessageBox.Show("Не выбран ни один параметр канала ни на одном листе! Объединить всё в одну таблицу?", "Не выбраны параметры каналов", MessageBoxButtons.YesNo);
                      if (dialogResult == DialogResult.Yes)
                      {
                          SaveWithoutCh = true;
                          checkBox11.Checked = true;
                      }
                      else if (dialogResult == DialogResult.No)
                      {
                          MessageBox.Show("Выберите хотя-бы один параметр.");
                          return;
                      }

                  }

                  int TablesNum1 = 1;
                  if (!SaveWithoutCh)    // если выбран вариант записи без каналов, то всего одна таблица у нас
                  {
                      TablesNum1 = TablesNum;
                  }

                  //Количество листов в рабочей книге
                  ex.SheetsInNewWorkbook = TablesNum1;  // кол-во таблиц = кол-во листов
                  //Добавить рабочую книгу
                  Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);


                  for (int TblCount = 1; TblCount <= TablesNum1; TblCount++)
                  {
                      //Получаем первый лист документа (счет начинается с 1)
                      Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(TblCount);
                      //Название листа (вкладки снизу)
                      string s = "";
                      if (!SaveWithoutCh)
                      {
                          foreach (TeconObject to in TeconObjects)  //берем первое попавшееся имя объекта, у которого индекс совпадает
                          {
                              if (to.Index == TblCount) { s = to.ObjTypeName; break; }
                          }
                      }
                      sheet.Name = "Лист" + TblCount.ToString() + "; " + s;   // вставляем в название номер таблицы
                    
                      //заполнения ячеек
                      //Шапка
                      sheet.Cells[3, 1] = String.Format("№ п/п");
                      sheet.Cells[3, 2] = String.Format("Время импорта");
                      sheet.Cells[3, 3] = String.Format("Марка");
                      sheet.Cells[3, 4] = String.Format("Наименование");
                      sheet.Cells[3, 5] = String.Format("Описание");
                      sheet.Cells[3, 6] = String.Format("Тип объекта");
                      sheet.Cells[3, 7] = String.Format("Подпись");
                      sheet.Cells[3, 8] = String.Format("KKS");
                      sheet.Cells[3, 9] = String.Format("PLC_Переменная");
                      sheet.Cells[3, 10] = String.Format("Контроллер");
                      sheet.Cells[3, 11] = String.Format("Ресурс/Группа");
                      sheet.Cells[3, 12] = String.Format("Адрес");
                      sheet.Cells[3, 13] = String.Format("Шаблон");
                      sheet.Cells[3, 14] = String.Format("Пер. архивирования");
                      sheet.Cells[3, 15] = String.Format("Группа событий");
                      sheet.Cells[3, 16] = String.Format("Классификатор");
                      sheet.Cells[2, 16] = String.Format("Классификатор");

                      sheet.Cells[4, 3] = String.Format("MARKA");
                      sheet.Cells[4, 4] = String.Format("NAME");
                      sheet.Cells[4, 5] = String.Format("DISC");
                      sheet.Cells[4, 6] = String.Format("OBJTYPENAME");
                      sheet.Cells[4, 7] = String.Format("OBJSIGN");
                      sheet.Cells[4, 8] = String.Format("KKS");
                      sheet.Cells[4, 9] = String.Format("PLC_VARNAME");
                      sheet.Cells[4, 10] = String.Format("PLC_NAME");
                      sheet.Cells[4, 11] = String.Format("PLC_GR");
                      sheet.Cells[4, 12] = String.Format("PLC_ADRESS");
                      sheet.Cells[4, 13] = String.Format("POUNAME");
                      sheet.Cells[4, 14] = String.Format("ARH_PER");
                      sheet.Cells[4, 15] = String.Format("EVKLASSIFIKATORNAME");
                      sheet.Cells[4, 16] = String.Format("KLASSIFIKATORNAME");

                      //заполнение объектами
                      int TmpCounter = 1;
                      foreach (TeconObject TObj in TeconObjects)
                      {
                          if ((TObj.Index == TblCount) || (SaveWithoutCh))
                          {
                              sheet.Cells[TmpCounter + 4, 1] = Convert.ToString(TmpCounter);
                              sheet.Cells[TmpCounter + 4, 3] = TObj.Marka;
                              sheet.Cells[TmpCounter + 4, 4] = TObj.Name;
                              sheet.Cells[TmpCounter + 4, 5] = TObj.Disc;
                              sheet.Cells[TmpCounter + 4, 6] = TObj.ObjTypeName;
                              sheet.Cells[TmpCounter + 4, 7] = TObj.ObjSign;
                              sheet.Cells[TmpCounter + 4, 8] = TObj.KKS;
                              sheet.Cells[TmpCounter + 4, 9] = TObj.PLC_varname;
                              sheet.Cells[TmpCounter + 4, 10] = TObj.PLC_Name;
                              sheet.Cells[TmpCounter + 4, 11] = TObj.PLC_GR;
                              sheet.Cells[TmpCounter + 4, 12] = TObj.PLC_address;
                              sheet.Cells[TmpCounter + 4, 13] = TObj.POUNAME;
                              sheet.Cells[TmpCounter + 4, 14] = TObj.Arc_Per;
                              sheet.Cells[TmpCounter + 4, 15] = TObj.EVKLASSIFIKATORNAME;
                              sheet.Cells[TmpCounter + 4, 16] = TObj.KLASSIFIKATORNAME;

                              sheet.Cells[TmpCounter + 4, 2] = TObj.Index.ToString();//не забудь убрать потом!

                              //Цикличное заполнение каналов
                              int k = 1;
                              int t = 0;



                              if (!SaveWithoutCh)
                              {
                                  foreach (TeconObjectChannel ch in TObj.Channels)
                                  {
                                      // Excel.Range ChNameRange = (Excel.Range)sheet.get_Range(sheet.Cells[2, 16 + k], sheet.Cells[2, 16 + k + 8]).Cells; //sheet.get_Range(sheet.Cells[2, 16+k], sheet.Cells[2, 16+k+8]);
                                      // ChNameRange.Merge(Type.Missing);
                                      sheet.Cells[2, 16 + k] = ch.ChannelName;
                                      Excel.Range c1 = sheet.Cells[2, 16 + k];

                                      if (!checkBox10.Checked)  // если не выбран пункт применить ко всем, тогда для каждой читаем
                                      {
                                          bS0 = SaveTablesParamsList[TblCount-1].S0;
                                          bS100 = SaveTablesParamsList[TblCount-1].S100;
                                          bM = SaveTablesParamsList[TblCount-1].M;
                                          bPLC_VARNAME = SaveTablesParamsList[TblCount-1].PLC_VARNAME;
                                          bED_IZM = SaveTablesParamsList[TblCount-1].ED_IZM;
                                          bARH_APP = SaveTablesParamsList[TblCount-1].ARH_APP;
                                          bDISC = SaveTablesParamsList[TblCount-1].DISC;
                                          bKA = SaveTablesParamsList[TblCount-1].KA;
                                          bKB = SaveTablesParamsList[TblCount-1].KB;
                                      }


                                      if (bS0) { sheet.Cells[TmpCounter + 4, 16 + k] = ch.S0; sheet.Cells[4, 16 + k] = "S0"; sheet.Cells[3, 16 + k] = "Шкала барогр низ"; k++; }
                                      if (bS100) { sheet.Cells[TmpCounter + 4, 16 + k] = ch.S100; sheet.Cells[4, 16 + k] = "S100"; sheet.Cells[3, 16 + k] = "Шкала барогр верх"; k++; }
                                      if (bM) { sheet.Cells[TmpCounter + 4, 16 + k] = ch.M; sheet.Cells[4, 16 + k] = "M"; sheet.Cells[3, 16 + k] = "Округлить до"; k++; }
                                      if (bPLC_VARNAME) { sheet.Cells[TmpCounter + 4, 16 + k] = ch.PLC_VARNAME; sheet.Cells[4, 16 + k] = "PLC_VARNAME"; sheet.Cells[3, 16 + k] = "PLC переменная"; k++; }
                                      if (bED_IZM) { sheet.Cells[TmpCounter + 4, 16 + k] = ch.ED_IZM; sheet.Cells[4, 16 + k] = "ED_IZM"; sheet.Cells[3, 16 + k] = "Ед. изм."; k++; }
                                      if (bARH_APP) { sheet.Cells[TmpCounter + 4, 16 + k] = ch.ARH_APP; sheet.Cells[4, 16 + k] = "ARH_APP"; sheet.Cells[3, 16 + k] = "Апертура арх."; k++; }
                                      if (bDISC) { sheet.Cells[TmpCounter + 4, 16 + k] = ch.DISC; sheet.Cells[4, 16 + k] = "DISC"; sheet.Cells[3, 16 + k] = "Описание"; k++; }
                                      if (bKA) { sheet.Cells[TmpCounter + 4, 16 + k] = ch.KA; sheet.Cells[4, 16 + k] = "KA"; sheet.Cells[3, 16 + k] = "Коэф. KA"; k++; }
                                      if (bKB) { sheet.Cells[TmpCounter + 4, 16 + k] = ch.KB; sheet.Cells[4, 16 + k] = "KB"; sheet.Cells[3, 16 + k] = "Коэф. KB"; k++; }

                                      Excel.Range c2 = sheet.Cells[2, 16 + k - 1];
                                      sheet.get_Range(c1, c2).Merge();

                                      //    sheet.Cells[TmpCounter + 4, 2] = .ToString();

                                  }
                              }

                              //textBox1.Clear();
                              // textBox1.AppendText(TObj.Channels[1].S100 + "\n";

                              TmpCounter++;
                          }
                      }
                  }

                  //Сохренение файла
                  ex.Application.ActiveWorkbook.SaveAs(SaveXlsPath, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                  //Удаляем приложение (выходим из экселя) - будет висеть в процессах!
                  ex.Quit();

                  textBox1.AppendText("Сохранение завершено;\n");
                  // textBox1.ScrollToCaret();
              }
          }

          private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
          {
              //при нажатии на любой чекбокс эта процедура тоже вызывается.
              int i = comboBox1.SelectedIndex;

              SaveTablesParamsList[i].S0 = bS0;
              SaveTablesParamsList[i].S100 = bS100;
              SaveTablesParamsList[i].M = bM;
              SaveTablesParamsList[i].PLC_VARNAME = bPLC_VARNAME;
              SaveTablesParamsList[i].ED_IZM = bED_IZM;
              SaveTablesParamsList[i].ARH_APP = bARH_APP;
              SaveTablesParamsList[i].DISC = bDISC;
              SaveTablesParamsList[i].KA = bKA;
              SaveTablesParamsList[i].KB = bKB;

              groupBox2.Text = "Лист" + SaveTablesParamsList[i].TableNum.ToString() + " Тип: " + SaveTablesParamsList[i].TypeName;
                  
          }

  
          }
    
}

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
        System.Data.DataTable TmpDG = new System.Data.DataTable(); // переменная, чтобы не создавать новые таблицы при новом зарпосе
        System.Data.DataTable dt1 = new System.Data.DataTable(); // таблица с данными из БД
        System.Data.DataTable dtDef = new System.Data.DataTable(); // таблица, которая собирается из двух.

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

        int preIndexCombobox; //индекс предыдущего выбранного типа, для того чтобы его сохранить при выборе следующего.
        int ObjCounter; // счетчик для замены цифры ПЛК на его название в повторном SQL запросе
        int ParamsCounter = 9; //временная переменная, указывающая кол-во параметров в канале
        List<int> ObjID = new List<int>();  // массив с прочитанными ID из Excel
        List<int> ObjTypeID = new List<int>(); //массив типов объектов
        List<int> ObjTypeIDUnic = new List<int>(); //массив типов объектов без повторений
        List<string> ObjTypeChannelsList = new List<string>(); //список каналов типа объекта
        List<Worksheet> Sheets = new List<Worksheet>();  // список листов, перезапись не катит, так как они до сохранения висят в памяти.
        List<HeadExcelCh> HeadExcelChList = new List<HeadExcelCh>(); //лист для шапки каналов
        

        int TablesNum;
        bool readID; // триггер успешно прочитанного файла
        bool readBD; // триггер успешно прочитанной БД


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
            public List<TeconObjectChannel> OldChannels = new List<TeconObjectChannel>();

            public int ObjTypeID;
            public int Index;
        }

        class TeconObjectChannel   // класс с описанием параметров канала объекта
        {
            public string S0 ;
            public string S100 ;
            public string M ;
            public string PLC_VARNAME ;
            public string ED_IZM ;
            public string ARH_APP ;
            public string DISC ;
            public string KA ;
            public string KB ;
            public string ID ;

            public string S0name =  "S0";
            public string S100name =  "S100";
            public string Mname =  "M";
            public string PLC_VARNAMEname =  "PLC_VARNAME";
            public string ED_IZMname =  "ED_IZM";
            public string ARH_APPname =  "ARH_APP";
            public string DISCname =  "DISC";
            public string KAname =  "KA";
            public string KBname =  "KB";

            public string ChannelName ;
        }

        class HeadExcelCh
        {
            public string ChName;
            public string ChTitle;
            public string ChParam;
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
        List<TeconObjectChannel> TeconObjectDefChannels = new List<TeconObjectChannel>();
        List<TeconObjectChannel> TeconObjectOldChannels = new List<TeconObjectChannel>();

        List<ObjTypeChannels> ObjTypeCh = new List<ObjTypeChannels>();
        List<SaveTableParams> SaveTablesParamsList = new List<SaveTableParams>();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            readBD = false;
            BaseAddr = textBox2.Text;
            ConStr = "character set=WIN1251;initial catalog=" + BaseAddr + ";user id=SYSDBA;password=masterkey"; // наша строка подключения, сделать её изменяемой!!!
            TeconObjectChannels.Clear(); // очистка предыдущего поиска
            TeconObjects.Clear();
            ObjTypeID.Clear();
            preIndexCombobox = 0;
            textBox1.AppendText( "Начат сбор данных из БД;\n");

            // textBox1.ScrollToCaret();
            int i=0;
            foreach (int ID in ObjID)   // для каждого распознанного ID делаем SQL запрос с последующими действиями
            {
                AddObjChannel(ID);     //добавляем каждый канал каждого тех объекта в список
                AddObjDefChannel(ID);
                CompareChannels();   //сравниваем два списка каналов
                AddObj(ID);            //делаем список тех объектов, содержащий списки каналов
                i++;
                label1.Text = "Прогресс: " + i + "/" + ObjID.Count.ToString();
               // MessageBox.Show(dtDef.Rows.Count.ToString()); 
               // textBox1.Lines[textBox1.Lines.Length-1] = "Прогресс: " + i + "/" + ObjID.Count.ToString();
               // MessageBox.Show((textBox1.Lines.Length - 1).ToString() + ";" + ObjID.Count.ToString());
            }
            FullKlName(); // находим полный путь для классификатора
            textBox1.AppendText( "Прочитаны данные из БД;\n");
            FindTypes();           //разбираем объекты на типы, ищем типы с одинаковыми каналами, присваиваем индексы
            textBox1.AppendText("Данные обработаны;\n");
            ShowData();            //показываем выбранные данные в таблице на форме
            textBox1.AppendText("Данные готовы к сохранению;\n");
            // textBox1.ScrollToCaret();
            readBD = true;
            EnabledCheck();
            ChOptionsToAll();

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

                    //Добавим условие, при котором будет происходить цикличный SQL запрос не теряя подключения (для ускорения выбора данных по ID)


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

               
                ConnectToBase("Select OBJTYPEPARAM.NAME from OBJTYPEPARAM where OBJTYPEPARAM.ID = " + objChnl.ID);
                objChnl.ChannelName = dt1.Rows[0][0].ToString();
             
                TeconObjectChannels.Add(objChnl);
            }
        }

        private void AddObjDefChannel(int ID)
        {
            
            //читаем описание канала, получаем список описаний всех каналов в ISA типе
            string SQL_disc = "select isaobjfields.disc, isaobjfields.name from isaobjfields where isaobjfields.isaobjid = (select isacardstemplate.tid from isacardstemplate where isacardstemplate.objtypeid = (select cards.objtypeid from cards where cards.id = " + ID + "))";
            ConnectToBase(SQL_disc);
            dtDef = dt1; // скопировали первую часть таблицы |disc|name|
            
            string SQL_ArhApp = "select OBJTYPEPARAM.isev, OBJTYPEPARAM.name from objtypeparam where objtypeparam.pid = (select cards.objtypeid from cards where cards.id = " + ID + ")";
            ConnectToBase(SQL_ArhApp);
            TmpDG = dt1; // вторая часть таблицы |isev|.name|
           //имея список каналов объекта (верхних), на его основе просто создадим другой список
            foreach (TeconObjectChannel toc in TeconObjectChannels)
            {
                var objChnl1 = new TeconObjectChannel();

                for (int i = 0; i < TmpDG.Rows.Count; i++)
                {
                    if (toc.ChannelName == TmpDG.Rows[i][1].ToString())
                    {
                         if ((Convert.ToInt16(TmpDG.Rows[i][0].ToString()) >= 1) && (Convert.ToInt16(TmpDG.Rows[i][0].ToString()) <= 6))
                            {
                                objChnl1.ARH_APP = "0";
                            }
                            else
                            {
                                objChnl1.ARH_APP = "-1";
                            }
                        break;
                    }
                }

                string newName = toc.ChannelName;
                if (newName.Substring(0,1) == ".")
                    {
                        newName = toc.ChannelName.Remove(0, 1);
                      //  MessageBox.Show("Точка обнаружена: " + newName);
                    }

                for (int i = 0; i < dtDef.Rows.Count; i++)
                {
                    if (newName == dtDef.Rows[i][1].ToString())
                    {
                        objChnl1.DISC = dtDef.Rows[i][0].ToString();
                        break;
                    }                               
                }

                    objChnl1.S0 = "0";
                    objChnl1.S100 = "100";
                    objChnl1.M = "1";
                    objChnl1.PLC_VARNAME = "";
                    objChnl1.ED_IZM = "";                    
                    objChnl1.KA = "1";
                    objChnl1.KB = "0";
                    objChnl1.ChannelName = toc.ChannelName;


                TeconObjectDefChannels.Add(objChnl1);

             //   MessageBox.Show("Записано Имя:" + objChnl1.ChannelName + "; Апертура:" + objChnl1.ARH_APP + "; Описание:" + objChnl1.DISC);
            }

        }

        private void CompareChannels()
        {
            TeconObjectOldChannels = TeconObjectChannels.ToList();

            for (int i = 0; i < TeconObjectChannels.Count; i++)
            {
                int j = 0;  // Если ни одного условия не выполнится, скипнем весь канал сразу

                if (TeconObjectDefChannels[i].DISC != "") { TeconObjectChannels[i].DISC = TeconObjectDefChannels[i].DISC; j++; } else {TeconObjectChannels[i].DISC  = "skipskipskip"; }
                if (TeconObjectDefChannels[i].ARH_APP != TeconObjectChannels[i].ARH_APP) { TeconObjectChannels[i].ARH_APP = TeconObjectDefChannels[i].ARH_APP; j++; } else { TeconObjectChannels[i].ARH_APP = "skipskipskip"; }
                if (TeconObjectDefChannels[i].S0 != TeconObjectChannels[i].S0) { TeconObjectChannels[i].S0 = TeconObjectDefChannels[i].S0; j++; } else { TeconObjectChannels[i].S0 = "skipskipskip"; }
                if (TeconObjectDefChannels[i].S100 != TeconObjectChannels[i].S100) { TeconObjectChannels[i].S100 = TeconObjectDefChannels[i].S100; j++; } else { TeconObjectChannels[i].S100 = "skipskipskip"; }
                if (TeconObjectDefChannels[i].M != TeconObjectChannels[i].M) { TeconObjectChannels[i].M = TeconObjectDefChannels[i].M; j++; } else { TeconObjectChannels[i].M = "skipskipskip"; }
                if (TeconObjectDefChannels[i].PLC_VARNAME != TeconObjectChannels[i].PLC_VARNAME) { TeconObjectChannels[i].PLC_VARNAME = TeconObjectDefChannels[i].PLC_VARNAME; j++; } else { TeconObjectChannels[i].PLC_VARNAME = "skipskipskip"; }
                if (TeconObjectDefChannels[i].ED_IZM != TeconObjectChannels[i].ED_IZM) { TeconObjectChannels[i].ED_IZM = TeconObjectDefChannels[i].ED_IZM; j++; } else { TeconObjectChannels[i].ED_IZM = "skipskipskip"; }
                if (TeconObjectDefChannels[i].KA != TeconObjectChannels[i].KA) { TeconObjectChannels[i].KA = TeconObjectDefChannels[i].KA; j++; } else { TeconObjectChannels[i].KA = "skipskipskip"; }
                if (TeconObjectDefChannels[i].KB != TeconObjectChannels[i].KB) { TeconObjectChannels[i].KB = TeconObjectDefChannels[i].KB; j++; } else { TeconObjectChannels[i].KB = "skipskipskip"; }

                //if (j == 0) { TeconObjectChannels[i].ChannelName = "skipskipskip"; }
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
                OldChannels = TeconObjectDefChannels.ToList(),
                ObjTypeID = Convert.ToInt16(dt1.Rows[0][14]),
            };


            ObjTypeID.Add(Convert.ToInt16(dt1.Rows[0][14])); // заполняем список типами объектов

           // MessageBox.Show("ObjTypeID: " + ObjTypeID.Count.ToString());

            // Делаем короткий запрос для получения имени контроллера, ибо в едином запросе такое хз как сделать
            ConnectToBase("Select CARDS.MARKA from CARDS where CARDS.ID = " + obj.PLC_Name);
            obj.PLC_Name = dt1.Rows[0][0].ToString();

                TeconObjects.Add(obj);
            TeconObjectChannels.Clear();
            TeconObjectDefChannels.Clear();
        }

        private void FullKlName() 
        {
            //SQL_CARDS = "Select CARDS.MARKA, CARDS.NAME, CARDS.DISC, OBJTYPE.NAME, CARDS.ARH_PER, CARDS.OBJSIGN, CARDS.PLC_ID, CARDS.PLC_GR, EVKLASSIFIKATOR.NAME, CARDS.KKS, ISAOBJ.NAME, KLASSIFIKATOR.NAME, CARDS.PLC_VARNAME, CARDS.PLC_ADRESS, CARDS.OBJTYPEID from CARDS, OBJTYPE, KLASSIFIKATOR, EVKLASSIFIKATOR, ISAOBJ, RESOURCES where CARDS.ID = " + ID + " and CARDS.OBJTYPEID = OBJTYPE.ID and CARDS.EVKLID = EVKLASSIFIKATOR.ID and CARDS.TEMPLATEID = ISAOBJ.ID and CARDS.KLID = KLASSIFIKATOR.ID";
            string SQL_KLASS = "Select * from KLASSIFIKATOR";
            ConnectToBase(SQL_KLASS);

            foreach (TeconObject obj in TeconObjects)
            {
                string KlassPath = obj.KLASSIFIKATORNAME;
                int TmpID = 0;

                string st = "";
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    st = dt1.Rows[i][2].ToString();   // выбираем НАЗВАНИЕ
                    if (st == obj.KLASSIFIKATORNAME) { TmpID = Convert.ToInt16(dt1.Rows[i][1]); /*выбираем PID*/ break; }
                }


                while (TmpID != 0)
                {
                    //foundRows = dt1.Select(TmpID.ToString(), "PID"); //не надо ли oчищать массив?

                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        if ((Convert.ToInt16(dt1.Rows[i][0]) == TmpID)) 
                        {
                            TmpID = Convert.ToInt16(dt1.Rows[i][1]);
                            if (TmpID == 0) { break; }
                            KlassPath = Convert.ToString(dt1.Rows[i][2]) + @"\" + KlassPath;
                            break; 
                        }
                    }                  
                }
               // MessageBox.Show("Результат: " + KlassPath);
                obj.KLASSIFIKATORNAME = KlassPath;
            }

        }



        private void FindTypes()
        {
            ObjTypeIDUnic.Clear();
            ObjTypeCh.Clear(); // для повторной активации функции
            ObjTypeChannelsList.Clear();

        ObjTypeIDUnic = ObjTypeID.Distinct().ToList(); //убираем повторяющиеся типы объектов

       // TablesNum = 0;


        foreach (int id in ObjTypeIDUnic)
        {
            SQL_OBJTYPES = "Select OBJTYPEPARAM.NAME from OBJTYPEPARAM where OBJTYPEPARAM.PID = " + id;
            ConnectToBase(SQL_OBJTYPES);

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
                       // MessageBox.Show("Новый тип: №" + k.ToString());
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
                //MessageBox.Show("Новый тип: №" + k.ToString());
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
              textBox1.AppendText("Лист" + i.ToString() + ": " + p.ToString() + ", "+s + ";\n");
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
            readID = false;

            //Открываем файл Экселя
            openFileDialog1.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //Создаём приложение.
           //    Excel.Application ObjExcel = new Excel.Application();
                Workbook book = Workbook.Load(openFileDialog1.FileName);

                //Открываем книгу.                                                                                                                                                        
             //  Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false,Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Выбираем таблицу(лист).
            //   Excel.Worksheet ObjWorkSheet;
            //    ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
                Worksheet sheet = book.Worksheets[0];

                SQLParams = ""; //очищаем предыдущий поиск параметров
                ObjID.Clear();
                ExcelCellCnt = 0;
                FirstIter = false;
                CellStr = "1";
               // MessageBox.Show(CellStr);

                //Excel.Range forYach = ObjWorkSheet.Cells[4, 3] as Excel.Range;
                if (Convert.ToString(sheet.Cells[3, 2]) != "ID")
                {
                    MessageBox.Show("Выбран некорректный файл экспорта: Найдено: "+ Convert.ToString(sheet.Cells[3, 2]) + " вместо ID!" );
                    return;
                }
                string CellContent = "";
                while (CellStr.Length > 0)  //будем читать столбец, пока не найдем пустую ячейку
                {
                    //MessageBox.Show();
                //   Excel.Range range = ObjWorkSheet.get_Range("C" + (ExcelCellCnt + 5).ToString());
                    CellStr = Convert.ToString(sheet.Cells[4 + ExcelCellCnt, 2]);  //[строка/столбец]
                   // CellStr = range.Text.ToString();
                    if (CellStr.Length > 0) { CellInt = Convert.ToInt16(CellStr); ObjID.Add(CellInt); } //Убрать потом как-нибудь это условие, чтобы не дублировалось                
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
               // ObjExcel.Quit();

                //Очищаем от старого текста окно вывода.
              //  textBox1.Clear();
                textBox1.AppendText("========Новый файл=======\n");
                textBox1.AppendText("Файл открыт;\n");
                textBox1.AppendText(openFileDialog1.FileName.ToString() + ";\n");
                textBox1.AppendText("ID объектов считаны (" + ObjID.Count.ToString()+ ")шт.;\n");
                // textBox1.ScrollToCaret();
                readID = true;
                readBD = false;
                
                button1.Enabled = true;
                button4.Enabled = false;
                comboBox1.Items.Clear(); // очистка списка листов
                SaveTablesParamsList.Clear();
                EnabledCheck();            

            }
        }


          private void button4_Click(object sender, EventArgs e)
          {

              //То что выбрано из параметров на момент нажатия кнопки сохранения
              int j = comboBox1.SelectedIndex;
              SaveTablesParamsList[j].S0 = checkBox1.Checked;
              SaveTablesParamsList[j].S100 = checkBox2.Checked;
              SaveTablesParamsList[j].M = checkBox3.Checked;
              SaveTablesParamsList[j].PLC_VARNAME = checkBox4.Checked;
              SaveTablesParamsList[j].ED_IZM = checkBox5.Checked;
              SaveTablesParamsList[j].ARH_APP = checkBox6.Checked;
              SaveTablesParamsList[j].DISC = checkBox7.Checked;
              SaveTablesParamsList[j].KA = checkBox8.Checked;
              SaveTablesParamsList[j].KB = checkBox9.Checked;

              Sheets.Clear();
              bool SaveWithoutCh = false;


              //saveFileDialog1.Filter = "Excel files (*.xls;*.xlsx)|*.xlsx;*.xls";
              saveFileDialog1.Filter = "Excel files (*.xls)|*.xls";
              var culture = new CultureInfo("ru-RU");
              string name = "";
              if (checkBox11.Checked) 
              {
                   name = "ТеконИмпорт_Листов_1;_Объектов_" + TeconObjects.Count.ToString() + ";_" + DateTime.Now.ToString(culture);
              } else 
              {
                   name = "ТеконИмпорт_Листов_" + TablesNum.ToString() + ";_Объектов_" + TeconObjects.Count.ToString() + ";_" + DateTime.Now.ToString(culture);
              }
              name = name.Replace(":", "_");
              saveFileDialog1.FileName = name;

              if (saveFileDialog1.ShowDialog() == DialogResult.OK)
              {
                  SaveXlsPath = saveFileDialog1.FileName;

                  Workbook workbook = new Workbook();

                  textBox1.AppendText("Начат процесс сохранения;\n");

                  SaveWithoutCh = checkBox11.Checked;

                  foreach (SaveTableParams st in SaveTablesParamsList)
                  {
                      bS0 = st.S0;
                      bS100 = st.S100;
                      bM = st.M;
                      bPLC_VARNAME = st.PLC_VARNAME;
                      bED_IZM = st.ED_IZM;
                      bARH_APP = st.ARH_APP;
                      bDISC = st.DISC;
                      bKA = st.KA;
                      bKB = st.KB;
                  }

                   /* if ((!bS0 && !bS100 && !bM && !bPLC_VARNAME && !bED_IZM && !bARH_APP && !bDISC && !bKA && !bKB && !checkBox11.Checked))
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

                    }*/
                  

                  int TablesNum1 = 1;
                  if (!SaveWithoutCh)    // если выбран вариант записи без каналов, то всего одна таблица у нас
                  {
                      TablesNum1 = TablesNum;
                  }

                  for (int TblCount = 1; TblCount <= TablesNum1; TblCount++)
                  {
                      Worksheet sheet = new Worksheet("Name");

                      string s = "";
                      if (!SaveWithoutCh)
                      {
                          foreach (TeconObject to in TeconObjects)  //берем первое попавшееся имя объекта, у которого индекс совпадает
                          {
                              if (to.Index == TblCount) { s = to.ObjTypeName; break; }
                          }
                      }
                    //  sheet.Name = "Лист" + TblCount.ToString() + "; " + s;   // вставляем в название номер таблицы
                      sheet.Name = "Лист" + TblCount.ToString() + "; " + s;

                      for (int i = 0; i <= 100; i++)  // Если будет менее 100 строк подряд, Офис 2010 откажется открыть файл. Фича библиотеки
                      {
                          sheet.Cells[i, 1] = new Cell("");
                      }
                      //заполнения ячеек
                      //Шапка
                      sheet.Cells[2, 0] = new Cell("№ п/п");
                      sheet.Cells[2, 1] = new Cell("Время импорта");
                      sheet.Cells[2, 2] = new Cell("Марка");
                      sheet.Cells[2, 3] = new Cell("Наименование");
                      sheet.Cells[2, 4] = new Cell("Описание");
                      sheet.Cells[2, 5] = new Cell("Тип объекта");
                      sheet.Cells[2, 6] = new Cell("Подпись");
                      sheet.Cells[2, 7] = new Cell("KKS");
                      sheet.Cells[2, 8] = new Cell("PLC_Переменная");
                      sheet.Cells[2, 9] = new Cell("Контроллер");
                      sheet.Cells[2, 10] = new Cell("Ресурс/Группа");
                      sheet.Cells[2, 11] = new Cell("Адрес");
                      sheet.Cells[2, 12] = new Cell("Шаблон");
                      sheet.Cells[2, 13] = new Cell("Пер. архивирования");
                      sheet.Cells[2, 14] = new Cell("Группа событий");
                      sheet.Cells[2, 15] = new Cell("Классификатор");
                      sheet.Cells[1, 15] = new Cell("Классификатор");

                      sheet.Cells[3, 2] = new Cell("MARKA");
                      sheet.Cells[3, 3] = new Cell("NAME");
                      sheet.Cells[3, 4] = new Cell("DISC");
                      sheet.Cells[3, 5] = new Cell("OBJTYPENAME");
                      sheet.Cells[3, 6] = new Cell("OBJSIGN");
                      sheet.Cells[3, 7] = new Cell("KKS");
                      sheet.Cells[3, 8] = new Cell("PLC_VARNAME");
                      sheet.Cells[3, 9] = new Cell("PLC_NAME");
                      sheet.Cells[3, 10] = new Cell("PLC_GR");
                      sheet.Cells[3, 11] = new Cell("PLC_ADRESS");
                      sheet.Cells[3, 12] = new Cell("POUNAME");
                      sheet.Cells[3, 13] = new Cell("ARH_PER");
                      sheet.Cells[3, 14] = new Cell("EVKLASSIFIKATORNAME");
                      sheet.Cells[3, 15] = new Cell("KLASSIFIKATORNAME");

                      bool b = true; // первый параметр запишем без условий
                      //Создадим шапку таблицы и перечень используемых каналов
                      foreach (TeconObject TObj in TeconObjects) 
                      {
                          foreach (TeconObjectChannel Tch in TObj.Channels)
                          {
                              if (Tch.S0 != "skipskipskip")
                              {
                                 // if (b) { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "S0", ChTitle = "Шкала барогр. низ" }; HeadExcelChList.Add(obj); b = false; }
                                  var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "S0", ChTitle = "Шкала барогр. низ" }; HeadExcelChList.Add(obj); b = false; 

                                  /*   foreach (HeadExcelCh hec in HeadExcelChList.ToArray())
                                  {
                                      if ((Tch.ChannelName == hec.ChName) && (hec.ChParam == "S0"))
                                      {
                                          break;
                                      }
                                      else 
                                      {
                                          var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "S0", ChTitle = "Шкала барогр. низ" }; HeadExcelChList.Add(obj);
                                      }
                                  }*/
                              }

                              if (Tch.S100 != "skipskipskip")
                              {
                              //    if (b) { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "S100", ChTitle = "Шкала барогр. верх" }; HeadExcelChList.Add(obj); b = false; }
                                   var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "S100", ChTitle = "Шкала барогр. верх" }; HeadExcelChList.Add(obj); b = false; 

                                  /*    foreach (HeadExcelCh hec in HeadExcelChList.ToArray())
                                  {
                                      if ((Tch.ChannelName == hec.ChName) && (hec.ChParam == "S100"))
                                      {
                                          break;
                                      }
                                      else
                                      {
                                          var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "S100", ChTitle = "Шкала барогр. верх" }; HeadExcelChList.Add(obj);
                                      }
                                  }*/
                              }

                              if (Tch.M != "skipskipskip")
                              {
                                //  if (b) { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "M", ChTitle = "Округлить до" }; HeadExcelChList.Add(obj); b = false; }
                                  var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "M", ChTitle = "Округлить до" }; HeadExcelChList.Add(obj); b = false; 

                                  /*foreach (HeadExcelCh hec in HeadExcelChList.ToArray())
                                 {
                                      if ((Tch.ChannelName == hec.ChName) && (hec.ChParam == "M"))
                                      {
                                          break;
                                      }
                                      else
                                      {
                                          var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "M", ChTitle = "Округлить до" }; HeadExcelChList.Add(obj);
                                      }
                                  }*/
                              }

                              if (Tch.PLC_VARNAME != "skipskipskip")
                              {
                                  //if (b) { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "PLC_VARNAME", ChTitle = "PLC_переменная" }; HeadExcelChList.Add(obj); b = false; }
                                   var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "PLC_VARNAME", ChTitle = "PLC_переменная" }; HeadExcelChList.Add(obj); b = false; 

                                  /*  foreach (HeadExcelCh hec in HeadExcelChList.ToArray())
                                  {
                                      if ((Tch.ChannelName == hec.ChName) && (hec.ChParam == "PLC_VARNAME"))
                                      {
                                          break;
                                      }
                                      else
                                      {
                                          var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "PLC_VARNAME", ChTitle = "PLC_переменная" }; HeadExcelChList.Add(obj);
                                      }
                                  }*/
                              }

                              if (Tch.ED_IZM != "skipskipskip")
                              {
                                  //if (b) { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "ED_IZM", ChTitle = "Ед. изм." }; HeadExcelChList.Add(obj); b = false; }
                                  var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "ED_IZM", ChTitle = "Ед. изм." }; HeadExcelChList.Add(obj); b = false; 

                                  /*  foreach (HeadExcelCh hec in HeadExcelChList.ToArray())
                                  {
                                      if ((Tch.ChannelName == hec.ChName) && (hec.ChParam == "ED_IZM"))
                                      {
                                          break;
                                      }
                                      else
                                      {
                                          var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "ED_IZM", ChTitle = "Ед. изм." }; HeadExcelChList.Add(obj);
                                      }
                                  }*/
                              }

                              if (Tch.DISC != "skipskipskip")
                              {
                                 // if (b) { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "DISC", ChTitle = "Описание" }; HeadExcelChList.Add(obj); b = false; }
                                   var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "DISC", ChTitle = "Описание" }; HeadExcelChList.Add(obj); b = false; 

                                  /*   foreach (HeadExcelCh hec in HeadExcelChList.ToArray())
                                /*  {
                                      if ((Tch.ChannelName == hec.ChName) && (hec.ChParam == "DISC"))
                                      {
                                          break;
                                      }
                                      else
                                      {
                                          var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "DISC", ChTitle = "Описание" }; HeadExcelChList.Add(obj);
                                      }
                                  }*/
                              }

                              if (Tch.KA != "skipskipskip")
                              {
                                  //if (b) { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "KA", ChTitle = "Коэф. КА" }; HeadExcelChList.Add(obj); b = false; }
                                   var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "KA", ChTitle = "Коэф. КА" }; HeadExcelChList.Add(obj); b = false; 

                                  /*  foreach (HeadExcelCh hec in HeadExcelChList.ToArray())
                               /*   {
                                      if ((Tch.ChannelName == hec.ChName) && (hec.ChParam == "KA"))
                                      {
                                          break;
                                      }
                                      else
                                      {
                                          var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "KA", ChTitle = "Коэф. КА" }; HeadExcelChList.Add(obj);
                                      }
                                  }*/
                              }

                              if (Tch.KB != "skipskipskip")
                              {
                                 // if (b) { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "KB", ChTitle = "Коэф. КВ" }; HeadExcelChList.Add(obj); b = false; }
                                   var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "KB", ChTitle = "Коэф. КВ" }; HeadExcelChList.Add(obj); b = false; 

                                  /*  foreach (HeadExcelCh hec in HeadExcelChList.ToArray())
                                 {
                                      if ((Tch.ChannelName == hec.ChName) && (hec.ChParam == "KB"))
                                      {
                                          break;
                                      }
                                      else
                                      {
                                          var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "KB", ChTitle = "Коэф. КВ" }; HeadExcelChList.Add(obj);
                                      }
                                  }*/
                              }
                          }
                      }

                      //Запишем шапку в Ecxel
                      //предварительно уберем из списка повторяющиеся элементы
                      var distinct = from item in HeadExcelChList
                                     group item by new { item.ChName, item.ChTitle, item.ChParam } into matches
                                     select matches.First();

                      HeadExcelChList = new List<HeadExcelCh>(distinct);

                   //   MessageBox.Show(HeadExcelChList.Count.ToString());

                      int k1 = 1;
                      foreach (HeadExcelCh hec in HeadExcelChList)
                      {
                          sheet.Cells[1, 15 + k1] = new Cell(hec.ChName);
                          sheet.Cells[2, 15 + k1] = new Cell(hec.ChTitle);
                          sheet.Cells[3, 15 + k1] = new Cell(hec.ChParam);
                          k1++;
                      }


                     // if (!checkBox10.Checked)  // если не выбран пункт применить ко всем, тогда для каждой читаем
                     // {
                          bS0 = SaveTablesParamsList[TblCount - 1].S0;
                          bS100 = SaveTablesParamsList[TblCount - 1].S100;
                          bM = SaveTablesParamsList[TblCount - 1].M;
                          bPLC_VARNAME = SaveTablesParamsList[TblCount - 1].PLC_VARNAME;
                          bED_IZM = SaveTablesParamsList[TblCount - 1].ED_IZM;
                          bARH_APP = SaveTablesParamsList[TblCount - 1].ARH_APP;
                          bDISC = SaveTablesParamsList[TblCount - 1].DISC;
                          bKA = SaveTablesParamsList[TblCount - 1].KA;
                          bKB = SaveTablesParamsList[TblCount - 1].KB;

                        /*  textBox1.AppendText("\nTblCount = " + (TblCount-1).ToString() + ";\n");
                          textBox1.AppendText("bS0 = " + bS0.ToString() + ";\n");
                          textBox1.AppendText("bS0 = " + bS0.ToString() + ";\n");
                          textBox1.AppendText("bS100 = " + bS100.ToString() + ";\n");
                          textBox1.AppendText("bM = " + bM.ToString() + ";\n");
                          textBox1.AppendText("bPLC_VARNAME = " + bPLC_VARNAME.ToString() + ";\n");
                          textBox1.AppendText("bED_IZM = " + bED_IZM.ToString() + ";\n");
                          textBox1.AppendText("bARH_APP = " + bARH_APP.ToString() + ";\n");
                          textBox1.AppendText("bDISC = " + bDISC.ToString() + ";\n");
                          textBox1.AppendText("bKA = " + bKA.ToString() + ";\n");
                          textBox1.AppendText("bKB = " + bKB.ToString() + ";\n");*/
                    //  }

                      //заполнение объектами
                      int TmpCounter = 0;
                      foreach (TeconObject TObj in TeconObjects)
                      {
                          if ((TObj.Index == TblCount) || (SaveWithoutCh))
                          {
                              sheet.Cells[TmpCounter + 4, 0] = new Cell(Convert.ToString(TmpCounter));
                              sheet.Cells[TmpCounter + 4, 2] = new Cell(TObj.Marka);
                              sheet.Cells[TmpCounter + 4, 3] = new Cell(TObj.Name);
                              sheet.Cells[TmpCounter + 4, 4] = new Cell(TObj.Disc);
                              sheet.Cells[TmpCounter + 4, 5] = new Cell(TObj.ObjTypeName);
                              sheet.Cells[TmpCounter + 4, 6] = new Cell(TObj.ObjSign);
                              sheet.Cells[TmpCounter + 4, 7] = new Cell(TObj.KKS);
                              sheet.Cells[TmpCounter + 4, 8] = new Cell(TObj.PLC_varname);
                              sheet.Cells[TmpCounter + 4, 9] = new Cell(TObj.PLC_Name);
                              sheet.Cells[TmpCounter + 4, 10] = new Cell(TObj.PLC_GR);
                              sheet.Cells[TmpCounter + 4, 11] = new Cell(TObj.PLC_address);
                              sheet.Cells[TmpCounter + 4, 12] = new Cell(TObj.POUNAME);
                              sheet.Cells[TmpCounter + 4, 13] = new Cell(TObj.Arc_Per);
                              sheet.Cells[TmpCounter + 4, 14] = new Cell(TObj.EVKLASSIFIKATORNAME);
                              sheet.Cells[TmpCounter + 4, 15] = new Cell(TObj.KLASSIFIKATORNAME);

                              //Цикличное заполнение каналов
                              int k = 1;
                              int preK = 1;
                              int t = 0;
                              string Channel = "";
                              string ChannelParam = "";
                              if (!SaveWithoutCh)
                              {
                                  while (Convert.ToString(sheet.Cells[1, 15 + k]).Length > 0)
                                  {
                                      Channel = Convert.ToString(sheet.Cells[1, 15 + k]);
                                      ChannelParam = Convert.ToString(sheet.Cells[3, 15 + k]);
                                     // MessageBox.Show(Convert.ToString(sheet.Cells[1, 15 + k]).Length.ToString() + "; k+15 = " + (k + 15).ToString());
                                      preK = k;
                                      for (int i = 0; i < TObj.Channels.Count; i++)
                                      {
                                          /*TObj.Channels[i][ChannelParam]*/ /*TObj.Channels[i].S0 == ChannelParam*/
                                          if ((TObj.Channels[i].ChannelName == Channel) && (TObj.Channels[i].S0name == ChannelParam)) { if (TObj.Channels[i].S0 == "skipskipskip") { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.OldChannels[i].S0); k++; break; } else { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.Channels[i].S0); k++; break; } }
                                          if ((TObj.Channels[i].ChannelName == Channel) && (TObj.Channels[i].S100name == ChannelParam)) { if (TObj.Channels[i].S100 == "skipskipskip") { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.OldChannels[i].S100); k++; break; } else { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.Channels[i].S100); k++; break; } }
                                          if ((TObj.Channels[i].ChannelName == Channel) && (TObj.Channels[i].Mname == ChannelParam)) { if (TObj.Channels[i].M == "skipskipskip") { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.OldChannels[i].M); k++; break; } else { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.Channels[i].M); k++; break; } }
                                          if ((TObj.Channels[i].ChannelName == Channel) && (TObj.Channels[i].PLC_VARNAMEname == ChannelParam)) { if (TObj.Channels[i].PLC_VARNAME == "skipskipskip") { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.OldChannels[i].PLC_VARNAME); k++; break; } else { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.Channels[i].PLC_VARNAME); k++; break; } }
                                          if ((TObj.Channels[i].ChannelName == Channel) && (TObj.Channels[i].ED_IZMname == ChannelParam)) { if (TObj.Channels[i].ED_IZM == "skipskipskip") { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.OldChannels[i].ED_IZM); k++; break; } else { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.Channels[i].ED_IZM); k++; break; } }
                                          if ((TObj.Channels[i].ChannelName == Channel) && (TObj.Channels[i].ARH_APPname == ChannelParam)) { if (TObj.Channels[i].ARH_APP == "skipskipskip") { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.OldChannels[i].ARH_APP); k++; break; } else { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.Channels[i].ARH_APP); k++; break; } }
                                          if ((TObj.Channels[i].ChannelName == Channel) && (TObj.Channels[i].DISCname == ChannelParam)) { if (TObj.Channels[i].DISC == "skipskipskip") { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.OldChannels[i].DISC); k++; break; } else { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.Channels[i].DISC); k++; break; } }
                                          if ((TObj.Channels[i].ChannelName == Channel) && (TObj.Channels[i].KAname == ChannelParam)) { if (TObj.Channels[i].KA == "skipskipskip") { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.OldChannels[i].KA); k++; break; } else { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.Channels[i].KA); k++; break; } }
                                          if ((TObj.Channels[i].ChannelName == Channel) && (TObj.Channels[i].KBname == ChannelParam)) { if (TObj.Channels[i].KB == "skipskipskip") { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.OldChannels[i].KB); k++; break; } else { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(TObj.Channels[i].KB); k++; break; } }

                                      }
                                      if (preK == k) { MessageBox.Show("Не найден канал! Параметр" + ChannelParam + "; Канал: " + /*TObj.Channels[i].ChannelName */ "==" + Channel); /*break;*/ }

                                  }
                              }






                             /* if (!SaveWithoutCh)
                              {
                                  foreach (TeconObjectChannel ch in TObj.Channels)
                                  {
                                      if ((bS0) && (ch.S0 != "skipskipskip")) { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(ch.S0); sheet.Cells[3, 15 + k] = new Cell("S0"); sheet.Cells[2, 15 + k] = new Cell("Шкала барогр низ"); sheet.Cells[1, 15 + k] = new Cell(ch.ChannelName); k++; }
                                      if ((bS100) && (ch.S100 != "skipskipskip")) { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(ch.S100); sheet.Cells[3, 15 + k] = new Cell("S100"); sheet.Cells[2, 15 + k] = new Cell("Шкала барогр верх"); sheet.Cells[1, 15 + k] = new Cell(ch.ChannelName); k++; }
                                      if ((bM) && (ch.M != "skipskipskip")) { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(ch.M); sheet.Cells[3, 15 + k] = new Cell("M"); sheet.Cells[2, 15 + k] = new Cell("Округлить до"); sheet.Cells[1, 15 + k] = new Cell(ch.ChannelName); k++; }
                                      if ((bPLC_VARNAME) && (ch.PLC_VARNAME != "skipskipskip")) { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(ch.PLC_VARNAME); sheet.Cells[3, 15 + k] = new Cell("PLC_VARNAME"); sheet.Cells[2, 15 + k] = new Cell("PLC переменная"); sheet.Cells[1, 15 + k] = new Cell(ch.ChannelName); k++; }
                                      if ((bED_IZM) && (ch.ED_IZM != "skipskipskip")) { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(ch.ED_IZM); sheet.Cells[3, 15 + k] = new Cell("ED_IZM"); sheet.Cells[2, 15 + k] = new Cell("Ед. изм."); sheet.Cells[1, 15 + k] = new Cell(ch.ChannelName); k++; }
                                      if ((bARH_APP) && (ch.ARH_APP != "skipskipskip")) { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(ch.ARH_APP); sheet.Cells[3, 15 + k] = new Cell("ARH_APP"); sheet.Cells[2, 15 + k] = new Cell("Апертура арх."); sheet.Cells[1, 15 + k] = new Cell(ch.ChannelName); k++; }
                                      if ((bDISC) && (ch.DISC != "skipskipskip")) { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(ch.DISC); sheet.Cells[3, 15 + k] = new Cell("DISC"); sheet.Cells[2, 15 + k] = new Cell("Описание"); sheet.Cells[1, 15 + k] = new Cell(ch.ChannelName); k++; }
                                      if ((bKA) && (ch.KA != "skipskipskip")) { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(ch.KA); sheet.Cells[3, 15 + k] = new Cell("KA"); sheet.Cells[2, 15 + k] = new Cell("Коэф. KA"); sheet.Cells[1, 15 + k] = new Cell(ch.ChannelName); k++; }
                                      if ((bKB) && (ch.KB != "skipskipskip")) { sheet.Cells[TmpCounter + 4, 15 + k] = new Cell(ch.KB); sheet.Cells[3, 15 + k] = new Cell("KB"); sheet.Cells[2, 15 + k] = new Cell("Коэф. KB"); sheet.Cells[1, 15 + k] = new Cell(ch.ChannelName); k++; }
                                  }
                              }*/

                              TmpCounter++;
                          }
                      }
                      Sheets.Add(sheet);
                  }

                  foreach (Worksheet sh in Sheets)
                  {
                      workbook.Worksheets.Add(sh);
                  }

                  workbook.Save(SaveXlsPath);

                  textBox1.AppendText("Сохранение завершено;\n");
              }
          }

          private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
          {
              int i = comboBox1.SelectedIndex;
              groupBox2.Text = "Лист" + SaveTablesParamsList[i].TableNum.ToString() + " Тип: " + SaveTablesParamsList[i].TypeName;

              SaveTablesParamsList[preIndexCombobox].S0 = checkBox1.Checked;
              SaveTablesParamsList[preIndexCombobox].S100 = checkBox2.Checked;
              SaveTablesParamsList[preIndexCombobox].M = checkBox3.Checked;
              SaveTablesParamsList[preIndexCombobox].PLC_VARNAME = checkBox4.Checked;
              SaveTablesParamsList[preIndexCombobox].ED_IZM = checkBox5.Checked;
              SaveTablesParamsList[preIndexCombobox].ARH_APP = checkBox6.Checked;
              SaveTablesParamsList[preIndexCombobox].DISC = checkBox7.Checked;
              SaveTablesParamsList[preIndexCombobox].KA = checkBox8.Checked;
              SaveTablesParamsList[preIndexCombobox].KB = checkBox9.Checked;
              
              checkBox1.Checked = SaveTablesParamsList[i].S0;
              checkBox2.Checked = SaveTablesParamsList[i].S100;
              checkBox3.Checked = SaveTablesParamsList[i].M;
              checkBox4.Checked = SaveTablesParamsList[i].PLC_VARNAME;
              checkBox5.Checked = SaveTablesParamsList[i].ED_IZM;
              checkBox6.Checked = SaveTablesParamsList[i].ARH_APP;
              checkBox7.Checked = SaveTablesParamsList[i].DISC;
              checkBox8.Checked = SaveTablesParamsList[i].KA;
              checkBox9.Checked = SaveTablesParamsList[i].KB;

              //label1.Text = comboBox1.SelectedIndex.ToString();
              preIndexCombobox = i; 
          }

          private void checkBox11_CheckedChanged(object sender, EventArgs e)
          {
              EnabledCheck();
          }

          private void checkBox10_CheckedChanged(object sender, EventArgs e)
          {
            //  if (checkBox10.Checked) { ChOptionsToAll(); }
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
                  SaveTablesParamsList[j].S0 = checkBox1.Checked;
                  SaveTablesParamsList[j].S100 = checkBox2.Checked;
                  SaveTablesParamsList[j].M = checkBox3.Checked;
                  SaveTablesParamsList[j].PLC_VARNAME = checkBox4.Checked;
                  SaveTablesParamsList[j].ED_IZM = checkBox5.Checked;
                  SaveTablesParamsList[j].ARH_APP = checkBox6.Checked;
                  SaveTablesParamsList[j].DISC = checkBox7.Checked;
                  SaveTablesParamsList[j].KA = checkBox8.Checked;
                  SaveTablesParamsList[j].KB = checkBox9.Checked;
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

      }
    
}

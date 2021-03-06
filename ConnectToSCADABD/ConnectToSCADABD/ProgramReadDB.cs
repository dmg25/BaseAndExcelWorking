﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;

namespace ConnectToSCADABD
{
    public class ProgramReadDB
    {
         System.Data.DataTable TmpDG = new System.Data.DataTable(); // переменная, чтобы не создавать новые таблицы при новом зарпосе
         System.Data.DataTable dtDef = new System.Data.DataTable(); // таблица, которая собирается из двух.

        public string SQLParams; //часть запроса, содержащая перечень нужных строк
        public string SQL_CARDS; //запрос основных параметров объекта
        public string SQL_CARDPARAMS; //запрос параметров каналов объекта
        public string SQL_OBJTYPES; //запрос параметров каналов объекта
        public int TablesNum;  // хранит количество будущих таблиц
        public string BaseAddr; // адрес базы, пишем его сюда из главной формы

//---------------------Коллекции классов, которые хранят все данные из больших запросов-----------------------------------------------

        public class SQL_TeconObjectChannel   // класс с описанием параметров канала объекта
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
            public int LoadID; // ID из экспортного файла
        }

        public class SQL_TeconObject   // класс с описанием основных параметров техобъекта
        {
            public string Marka;
            public string Name;
            public string Disc;
            public string ObjTypeName;
            public string Arc_Per;
            public string ObjSign;
            public string PLC_Name;
            public string PLC_GR;
            public string EVKLASSIFIKATORNAME;
            public string KKS;
            public string POUNAME;
            public string KLASSIFIKATORNAME;
            public string PLC_varname;
            public string PLC_address;
            public int ObjTypeID;
            public int LoadID; // ID из экспортного файла
           // public int Index;
        }

       public List<SQL_TeconObjectChannel> SQL_Channels = new List<SQL_TeconObjectChannel>();
       public List<SQL_TeconObjectChannel> SQL_DefChannels = new List<SQL_TeconObjectChannel>();
       public List<SQL_TeconObject> SQL_Objects = new List<SQL_TeconObject>();

//------------------------------------------------------------------------------------------------------------------------------------

        public class InitValue  // класс объекта с начальным значением его иса объекта
        {
            public string ObjID;
            public string Marka;
            public string InitialValue;
            public string ObjType;

        }

        public class TeconObject   // класс с описанием основных параметров техобъекта
        {
            public string Marka;
            public string Name;
            public string Disc;
            public string ObjTypeName;
            public string Arc_Per;
            public string ObjSign;
            public string PLC_Name;
            public string PLC_GR;
            public string EVKLASSIFIKATORNAME;
            public string KKS;
            public string POUNAME;
            public string KLASSIFIKATORNAME;
            public string PLC_varname;
            public string PLC_address;
            public List<TeconObjectChannel> Channels = new List<TeconObjectChannel>();  // два сравниваемых между собой листа. В первый заносятся все каналы, параметры которых если совпадают с умолчаниями, изменяются на "skipskipskip"
            public List<TeconObjectChannel> OldChannels = new List<TeconObjectChannel>();  // резервный лист с каналами до сравнивания с умолчаниями. Если на одном тех объекте добавился параметр канала, а на другом он записываться не должен, то пустая ячейка заполнится существующим значением, чтобы оно не удалилось

            public int ObjTypeID;
            public int Index;  // Индекс, согласно которому будет происходить разделение объектов по разным листам.
        }

        public class TeconObjectChannel   // класс с описанием параметров канала объекта
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

            public string S0name = "S0";
            public string S100name = "S100";
            public string Mname = "M";
            public string PLC_VARNAMEname = "PLC_VARNAME";
            public string ED_IZMname = "ED_IZM";
            public string ARH_APPname = "ARH_APP";
            public string DISCname = "DISC";
            public string KAname = "KA";
            public string KBname = "KB";

            public string ChannelName;
        }

        public class ObjTypeChannels   // класс с перечнем каналов каждого типа
            {
                public int TypeID;
                public int Index; //индекс совпадения с другими типами
                public List<string> Channels = new List<string>();
            }

        public List<ProgramReadDB.TeconObject> TeconObjects = new List<ProgramReadDB.TeconObject>();
        public List<ProgramReadDB.TeconObjectChannel> TeconObjectChannels = new List<ProgramReadDB.TeconObjectChannel>();
        public List<ProgramReadDB.TeconObjectChannel> TeconObjectDefChannels = new List<ProgramReadDB.TeconObjectChannel>();
        public List<ProgramReadDB.TeconObjectChannel> TeconObjectOldChannels = new List<ProgramReadDB.TeconObjectChannel>();

        public List<ProgramReadDB.ObjTypeChannels> ObjTypeCh = new List<ProgramReadDB.ObjTypeChannels>();
        public List<InitValue> InitValues = new List<InitValue>(); // лист объектов с начальными значениями
        public List<int> ObjTypeID = new List<int>(); //массив типов объектов
        public List<int> ObjTypeIDUnic = new List<int>(); //массив типов объектов без повторений
        public List<string> ObjTypeChannelsList = new List<string>(); //список каналов типа объекта
        public List<string> ObjTypeChannelsMatched = new List<string>(); //список каналов типа объекта


        public void Big_SQL(List<int> ObjID) //ОНО СРАБОТАЕТ ТАК? //Запрос всех каналов
        {
            
            // Делаем часть строки с ID для запроса Каналов
            string Ch = "";
            bool firstIter = true;
            for (int i = 0; i < ObjID.Count; i++)   // объединили все id в один запрос
            {
                if (firstIter) { Ch = Ch + " (CARDPARAMS.CARDID =  " + ObjID[i].ToString() + " and OBJTYPEPARAM.ID = CARDPARAMS.OBJTYPEPARAMID) "; firstIter = false; }  // запросы проверены, таким образом мы не получим бесконечного цикла в ответе и ID каналов подменятся на имена
                else { Ch = Ch + " or (CARDPARAMS.CARDID = " + ObjID[i].ToString() + " and OBJTYPEPARAM.ID = CARDPARAMS.OBJTYPEPARAMID) "; }
            }
            
            //1. Запрос параметров каналов
            SQL_CARDPARAMS = "Select CARDPARAMS.S0, CARDPARAMS.S100, CARDPARAMS.M, CARDPARAMS.PLC_VARNAME, CARDPARAMS.ED_IZM, CARDPARAMS.ARH_APP, CARDPARAMS.DISC, CARDPARAMS.KA, CARDPARAMS.KB, CARDPARAMS.OBJTYPEPARAMID, OBJTYPEPARAM.NAME, CARDPARAMS.CARDID from CARDPARAMS, OBJTYPEPARAM where " + Ch;
            ProgramConnect connect = new ProgramConnect();
            connect.ConnectToBase(SQL_CARDPARAMS, BaseAddr);
            TmpDG = connect.dt1;

            for (int i = 0; i < TmpDG.Rows.Count; i++)
            {
                var objChnl = new SQL_TeconObjectChannel()
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
                    ID = TmpDG.Rows[i][9].ToString(),  //OBJPARAMID
                    ChannelName = TmpDG.Rows[i][10].ToString(),
                    LoadID = Convert.ToInt32(TmpDG.Rows[i][11]),  //ID исходный     
                };

                //connect.ConnectToBase("Select OBJTYPEPARAM.NAME from OBJTYPEPARAM where OBJTYPEPARAM.ID = " + objChnl.ID, BaseAddr);
                //objChnl.ChannelName = connect.dt1.Rows[0][0].ToString();
                SQL_Channels.Add(objChnl);
            }
            TmpDG.Clear();

            //2. Запрос деф параметров каналов
            //На удивление, имеем аналогичную часть запроса как и в предыдущем шаге, поэтому ничего менять в ней не будем.

            SQL_CARDPARAMS = "select OBJTYPEPARAM.disc, OBJTYPEPARAM.isev, OBJTYPEPARAM.NAME, cardparams.cardid, OBJTYPEPARAM.id  from OBJTYPEPARAM, cardparams where  " + Ch;
            ProgramConnect connect1 = new ProgramConnect();  // не до экспериментов, создал новую переменную, можно проверить потом, прокатит с той же или нет
            connect1.ConnectToBase(SQL_CARDPARAMS, BaseAddr);
            TmpDG = connect1.dt1;

            for (int i = 0; i < TmpDG.Rows.Count; i++)
            {
                var DefobjChnl = new SQL_TeconObjectChannel()
                {
                    S0 = "0",
                    S100 = "100",
                    M = "1",
                    PLC_VARNAME = "",
                    ED_IZM = "",
                    ARH_APP = TmpDG.Rows[i][1].ToString(),
                    DISC = TmpDG.Rows[i][0].ToString(),
                    KA = "1",
                    KB = "0",
                    ID = TmpDG.Rows[i][4].ToString(),  //OBJPARAMID
                    ChannelName = TmpDG.Rows[i][2].ToString(),
                    LoadID = Convert.ToInt32(TmpDG.Rows[i][3]),  //ID исходный                   
                };

                //connect.ConnectToBase("Select OBJTYPEPARAM.NAME from OBJTYPEPARAM where OBJTYPEPARAM.ID = " + objChnl.ID, BaseAddr);
                //objChnl.ChannelName = connect.dt1.Rows[0][0].ToString();
                SQL_DefChannels.Add(DefobjChnl);
                
            }

            TmpDG.Clear();

            //3. Запрос параметров тех объекта
            Ch = "";
            firstIter = true;
            for (int i = 0; i < ObjID.Count; i++)   // объединили все id в один запрос
            {
                if (firstIter) { Ch = Ch + " (CARDS.ID = " + ObjID[i].ToString() + " and CARDS.OBJTYPEID = OBJTYPE.ID and CARDS.EVKLID = EVKLASSIFIKATOR.ID and CARDS.TEMPLATEID = ISAOBJ.ID and CARDS.KLID = KLASSIFIKATOR.ID) "; firstIter = false; }  // запросы проверены, таким образом мы не получим бесконечного цикла в ответе и ID каналов подменятся на имена
                else { Ch = Ch + " or (CARDS.ID = " + ObjID[i].ToString() + " and CARDS.OBJTYPEID = OBJTYPE.ID and CARDS.EVKLID = EVKLASSIFIKATOR.ID and CARDS.TEMPLATEID = ISAOBJ.ID and CARDS.KLID = KLASSIFIKATOR.ID) "; }
            }

            SQL_CARDS = "Select CARDS.MARKA, CARDS.NAME, CARDS.DISC, OBJTYPE.NAME, CARDS.ARH_PER, CARDS.OBJSIGN, CARDS.PLC_ID, CARDS.PLC_GR, EVKLASSIFIKATOR.NAME, CARDS.KKS, ISAOBJ.NAME, KLASSIFIKATOR.NAME, CARDS.PLC_VARNAME, CARDS.PLC_ADRESS, CARDS.OBJTYPEID, CARDS.ID from CARDS, OBJTYPE, KLASSIFIKATOR, EVKLASSIFIKATOR, ISAOBJ where " + Ch; //, RESOURCES
            ProgramConnect connect2 = new ProgramConnect();  // не до экспериментов, создал новую переменную, можно проверить потом, прокатит с той же или нет
            connect2.ConnectToBase(SQL_CARDS, BaseAddr);
            TmpDG = connect2.dt1;

            //---------------------------Заполняем параметры объекта---------------------------------------------------------------  
            for (int i = 0; i < TmpDG.Rows.Count; i++)
            {
                var obj = new SQL_TeconObject()
                {
                    Marka = TmpDG.Rows[i][0].ToString(),
                    Name = TmpDG.Rows[i][1].ToString(),
                    Disc = TmpDG.Rows[i][2].ToString(),
                    ObjTypeName = TmpDG.Rows[i][3].ToString(),
                    Arc_Per = TmpDG.Rows[i][4].ToString(),
                    ObjSign = TmpDG.Rows[i][5].ToString(),
                    PLC_Name = TmpDG.Rows[i][6].ToString(),
                    PLC_GR = TmpDG.Rows[i][7].ToString(),
                    EVKLASSIFIKATORNAME = TmpDG.Rows[i][8].ToString(),
                    KKS = TmpDG.Rows[i][9].ToString(),
                    POUNAME = TmpDG.Rows[i][10].ToString(),
                    KLASSIFIKATORNAME = TmpDG.Rows[i][11].ToString(),
                    PLC_varname = TmpDG.Rows[i][12].ToString(),
                    PLC_address = TmpDG.Rows[i][13].ToString(),
                    ObjTypeID = Convert.ToInt16(TmpDG.Rows[i][14]),
                    LoadID = Convert.ToInt16(TmpDG.Rows[i][15])

                };
                //MessageBox.Show(i.ToString());
                SQL_Objects.Add(obj);
            }
            TmpDG.Clear();
            //4. Запрос одного параметра с названием ПЛК(ОТДЕЛЬНО ПОДУМОЙ!)
            //ПОКА БЕЗ ЭТОГО ПРОБУЕМ, ПОТОМ ПРИКРУТИМ


        }

        
            public void AddInitValue(int ID)
            {
                string SQL = "Select ISACARDS.MARKA, ISACARDS.INITIALVALUE from ISACARDS where ISACARDS.CARDSID = " + ID;
                ProgramConnect connect = new ProgramConnect();
                connect.ConnectToBase(SQL, BaseAddr);

                var obj = new InitValue()
                {
                    Marka = connect.dt1.Rows[0][0].ToString(),
                    InitialValue = connect.dt1.Rows[0][1].ToString(), 
                    ObjID = ID.ToString()
                };

                SQL = "Select OBJTYPE.NAME from CARDS, OBJTYPE where CARDS.ID = " + ID + " and CARDS.OBJTYPEID = OBJTYPE.ID";
                connect.ConnectToBase(SQL, BaseAddr);

                obj.ObjType = connect.dt1.Rows[0][0].ToString();

                InitValues.Add(obj);
            }

            /*public void WriteInitValue(string SQL)
            {
                //string SQL = "Select ISACARDS.MARKA, ISACARDS.INITIALVALUE from ISACARDS where ISACARDS.CARDSID = " + ID;
                ProgramConnect connect = new ProgramConnect();
                connect.ConnectToBase(SQL);
            }*/



            public void AddObjChannel(int ID)
            {
                //часть, которая читает лст вместо sql запроса
                //для каждого параметра пробегаемся по списку и выделяем тот, где ID совпадает - записываем. Добавляем в лист. Далее как обычно.
                //создадим кучку локальных переменных, чтобы потом скопом их внести в лист

               /* string tmp_S0;
                string tmp_S100;
                string tmp_M;
                string tmp_PLC_VARNAME;
                string tmp_ED_IZM;
                string tmp_ARH_APP;
                string tmp_DISC;
                string tmp_KA;
                string tmp_KB;
                string tmp_ID;
                string tmp_ChannelName;*/

                //выбираем каждый объект листа, ID которого совпадает с искомым, и добавляем в наш новый лист. 
                foreach (SQL_TeconObjectChannel sql in SQL_Channels)
                {
                    if (sql.LoadID == ID)
                    {
                        var objChnl = new TeconObjectChannel()
                        {
                            S0 = sql.S0,
                            S100 = sql.S100,
                            M = sql.M,
                            PLC_VARNAME = sql.PLC_VARNAME,
                            ED_IZM = sql.ED_IZM,
                            ARH_APP = sql.ARH_APP,
                            DISC = sql.DISC,
                            KA = sql.KA,
                            KB = sql.KB,
                            ID = sql.ID,  //OBJPARAMID
                            ChannelName = sql.ChannelName,
                        };
                        TeconObjectChannels.Add(objChnl);  // один лист на один объект. потом он очистится
                    }
                }


               /* SQL_CARDPARAMS = "Select CARDPARAMS.S0, CARDPARAMS.S100, CARDPARAMS.M, CARDPARAMS.PLC_VARNAME, CARDPARAMS.ED_IZM, CARDPARAMS.ARH_APP, CARDPARAMS.DISC, CARDPARAMS.KA, CARDPARAMS.KB, CARDPARAMS.OBJTYPEPARAMID, OBJTYPEPARAM.NAME from CARDPARAMS, OBJTYPEPARAM where CARDPARAMS.CARDID = " + ID + " and OBJTYPEPARAM.ID = CARDPARAMS.OBJTYPEPARAMID";
                ProgramConnect connect = new ProgramConnect();
                connect.ConnectToBase(SQL_CARDPARAMS, BaseAddr); 

                TmpDG = connect.dt1;

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
                        ID = TmpDG.Rows[i][9].ToString(),  //OBJPARAMID
                        ChannelName = TmpDG.Rows[i][10].ToString(),


                    };

                    //connect.ConnectToBase("Select OBJTYPEPARAM.NAME from OBJTYPEPARAM where OBJTYPEPARAM.ID = " + objChnl.ID, BaseAddr);
                    //objChnl.ChannelName = connect.dt1.Rows[0][0].ToString();
                    TeconObjectChannels.Add(objChnl);
                }*/

            }

            public void AddObjDefChannel(int ID)
            {
                 
              /* 1. пробегаем по листу с текущими каналами и составляем новый лист с objparamid 
               * 2. параллельно составляем строку с айдишниками и делаем ОР
               * 3. делаем запрос со всеми айдишниками и вываливаем в таблицу
               * 4. пробегаем по таблице и при совпадении айдишника КОТОРЫЙ ТОЖЕ В НЕЙ ЕСТЬ записываем данные в дефлист
               * 5. получили дефлист с теми же индексами, что и лист текущих каналов
               * 
               * */

               /* foreach (SQL_TeconObjectChannel sql in SQL_Channels)
                {
                    if (sql.LoadID == ID)
                    {
                        var objChnl = new TeconObjectChannel()
                        {
                            S0 = sql.S0,
                            S100 = sql.S100,
                            M = sql.M,
                            PLC_VARNAME = sql.PLC_VARNAME,
                            ED_IZM = sql.ED_IZM,
                            ARH_APP = sql.ARH_APP,
                            DISC = sql.DISC,
                            KA = sql.KA,
                            KB = sql.KB,
                            ID = sql.ID,  //OBJPARAMID
                            ChannelName = sql.ChannelName,
                        };
                        TeconObjectChannels.Add(objChnl);
                    }
                }*/


              //  MessageBox.Show(TeconObjectChannels.Count.ToString());
              //  MessageBox.Show(SQL_DefChannels.Count.ToString());

                foreach (TeconObjectChannel toc in TeconObjectChannels)
                {
                    foreach (SQL_TeconObjectChannel sql in SQL_DefChannels)
                    {
                        var objChnl1 = new TeconObjectChannel();
                       // MessageBox.Show("sql.LoadID: "+sql.LoadID.ToString() + " == ID: " + ID.ToString());
                      //  MessageBox.Show("toc.ID: " + toc.ID.ToString() + " == sql.ID: " + sql.ID);
                        if ((sql.LoadID == ID) && (toc.ID == sql.ID))
                        {
                            if ((Convert.ToInt16(sql.ARH_APP) >= 1) && (Convert.ToInt16(sql.ARH_APP) <= 6))
                            {
                                objChnl1.ARH_APP = "0";
                            }
                            else
                            {
                                objChnl1.ARH_APP = "-1";
                            }

                            objChnl1.DISC = sql.DISC;                                                  


                            objChnl1.S0 = "0";
                            objChnl1.S100 = "100";
                            objChnl1.M = "1";
                            objChnl1.PLC_VARNAME = "";
                            objChnl1.ED_IZM = "";
                            objChnl1.KA = "1";
                            objChnl1.KB = "0";
                            objChnl1.ChannelName = toc.ChannelName;

                            TeconObjectDefChannels.Add(objChnl1);
                            break;
                        }
                    }                   
                }

        /*        string SQL_ArhApp = "select OBJTYPEPARAM.isev, OBJTYPEPARAM.name from objtypeparam where objtypeparam.pid = (select cards.objtypeid from cards where cards.id = " + ID + ")";
                connect.ConnectToBase(SQL_ArhApp);
                TmpDG = connect.dt1; // вторая часть таблицы |isev|.name|*/

                //имея список каналов объекта (верхних), на его основе просто создадим другой список

       /*         foreach (TeconObjectChannel toc in TeconObjectChannels)
                {
                    var objChnl1 = new TeconObjectChannel();

                    for (int i = 0; i < dtDef.Rows.Count; i++)
                    {
                        if (toc.ID == dtDef.Rows[i][2].ToString())
                        {
                            if ((Convert.ToInt16(dtDef.Rows[i][1].ToString()) >= 1) && (Convert.ToInt16(dtDef.Rows[i][1].ToString()) <= 6))
                            {
                                objChnl1.ARH_APP = "0";
                            }
                            else
                            {
                                objChnl1.ARH_APP = "-1";
                            }

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

                    TeconObjectDefChannels.Add(objChnl1);*/
                





              /*  foreach (TeconObjectChannel toc in TeconObjectChannels)
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
                    if (newName.Substring(0, 1) == ".")
                    {
                        newName = toc.ChannelName.Remove(0, 1);
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
                }*/
            }

            public void RepeatChNameAlarm()
            {
                if (ObjTypeChannelsMatched.Count > 0)
                {
                    List<string> ObjTypeChannelsMatchedDistinct = new List<string>();
                    ObjTypeChannelsMatchedDistinct = ObjTypeChannelsMatched.Distinct().ToList();
                    string str = "";
                    for (int i = 0; i < ObjTypeChannelsMatchedDistinct.Count; i++)
                    {
                        str = str + "'"+ObjTypeChannelsMatchedDistinct[i]+"' ; " ;                   
                    }

                    MessageBox.Show("Обнаружены одинаковые каналы у следующих типов: '" + str + "\n Повторяющиеся каналы будут записаны в файл некорректно(только 1 экз.)! \n Переименование канала в библиотеке типов решит проблему.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        

            public void CompareChannels()
            {
                TeconObjectOldChannels = TeconObjectChannels.ToList();

                for (int i = 0; i < TeconObjectChannels.Count; i++)
                {
                    int j = 0;  // Если ни одного условия не выполнится, скипнем весь канал сразу

                    if (TeconObjectDefChannels[i].DISC != "") { TeconObjectChannels[i].DISC = TeconObjectDefChannels[i].DISC; j++; } else { TeconObjectChannels[i].DISC = "skipskipskip"; }
                    if (TeconObjectDefChannels[i].ARH_APP != TeconObjectChannels[i].ARH_APP) { TeconObjectChannels[i].ARH_APP = TeconObjectDefChannels[i].ARH_APP; j++; } else { TeconObjectChannels[i].ARH_APP = "skipskipskip"; }
                    if (TeconObjectDefChannels[i].S0 != TeconObjectChannels[i].S0) { TeconObjectChannels[i].S0 = TeconObjectDefChannels[i].S0; j++; } else { TeconObjectChannels[i].S0 = "skipskipskip"; }
                    if (TeconObjectDefChannels[i].S100 != TeconObjectChannels[i].S100) { TeconObjectChannels[i].S100 = TeconObjectDefChannels[i].S100; j++; } else { TeconObjectChannels[i].S100 = "skipskipskip"; }
                    if (TeconObjectDefChannels[i].M != TeconObjectChannels[i].M) { TeconObjectChannels[i].M = TeconObjectDefChannels[i].M; j++; } else { TeconObjectChannels[i].M = "skipskipskip"; }
                    if (TeconObjectDefChannels[i].PLC_VARNAME != TeconObjectChannels[i].PLC_VARNAME) { TeconObjectChannels[i].PLC_VARNAME = TeconObjectDefChannels[i].PLC_VARNAME; j++; } else { TeconObjectChannels[i].PLC_VARNAME = "skipskipskip"; }                   
                    if (TeconObjectDefChannels[i].KA != TeconObjectChannels[i].KA) { TeconObjectChannels[i].KA = TeconObjectDefChannels[i].KA; j++; } else { TeconObjectChannels[i].KA = "skipskipskip"; }
                    if (TeconObjectDefChannels[i].KB != TeconObjectChannels[i].KB) { TeconObjectChannels[i].KB = TeconObjectDefChannels[i].KB; j++; } else { TeconObjectChannels[i].KB = "skipskipskip"; }
                    if (TeconObjectChannels[i].ED_IZM.Length != 0)
                    { //исключаем случай, когда в поле откуда-то берутся пробелы, именно в единицах измерения.
                        string s = TeconObjectChannels[i].ED_IZM;
                        for (int l = 0; l < s.Length; l++)
                        {
                            if (s[l].ToString() != " ") { j++; break; }
                            if (s.Length - 1 == l) { TeconObjectChannels[i].ED_IZM = "skipskipskip"; }
                        }
                    }
                    else { TeconObjectChannels[i].ED_IZM = "skipskipskip"; }
                    
                }
            }

            public void AddObj(int ID)
            {
                foreach (SQL_TeconObject sql in SQL_Objects)
                {
                    if (sql.LoadID == ID)
                    {
                        var obj = new TeconObject()
                        {
                            Marka = sql.Marka,
                            Name = sql.Name,
                            Disc = sql.Disc,
                            ObjTypeName = sql.ObjTypeName,
                            Arc_Per = sql.Arc_Per,
                            ObjSign = sql.ObjSign,
                            PLC_Name = sql.PLC_Name,
                            PLC_GR = sql.PLC_GR,
                            EVKLASSIFIKATORNAME = sql.EVKLASSIFIKATORNAME,
                            KKS = sql.KKS,
                            POUNAME = sql.POUNAME,
                            KLASSIFIKATORNAME = sql.KLASSIFIKATORNAME,
                            PLC_varname = sql.PLC_varname,
                            PLC_address = sql.PLC_address,
                            Channels = TeconObjectChannels.ToList(),
                            OldChannels = TeconObjectDefChannels.ToList(),
                            ObjTypeID = sql.ObjTypeID,
                        };
                        ObjTypeID.Add(sql.ObjTypeID); // заполняем список типами объектов


                        /*   SQL_CARDS = "Select CARDS.MARKA, CARDS.NAME, CARDS.DISC, OBJTYPE.NAME, CARDS.ARH_PER, CARDS.OBJSIGN, CARDS.PLC_ID, CARDS.PLC_GR, EVKLASSIFIKATOR.NAME, CARDS.KKS, ISAOBJ.NAME, KLASSIFIKATOR.NAME, CARDS.PLC_VARNAME, CARDS.PLC_ADRESS, CARDS.OBJTYPEID, CARDS.ID from CARDS, OBJTYPE, KLASSIFIKATOR, EVKLASSIFIKATOR, ISAOBJ, RESOURCES where CARDS.ID = " + ID + " and CARDS.OBJTYPEID = OBJTYPE.ID and CARDS.EVKLID = EVKLASSIFIKATOR.ID and CARDS.TEMPLATEID = ISAOBJ.ID and CARDS.KLID = KLASSIFIKATOR.ID";
                           ProgramConnect connect = new ProgramConnect();
                           connect.ConnectToBase(SQL_CARDS, BaseAddr);

                           //---------------------------Заполняем параметры объекта---------------------------------------------------------------                 
                           var obj = new TeconObject()
                           {
                               Marka = connect.dt1.Rows[0][0].ToString(),
                               Name = connect.dt1.Rows[0][1].ToString(),
                               Disc = connect.dt1.Rows[0][2].ToString(),
                               ObjTypeName = connect.dt1.Rows[0][3].ToString(),
                               Arc_Per = connect.dt1.Rows[0][4].ToString(),
                               ObjSign = connect.dt1.Rows[0][5].ToString(),
                               PLC_Name = connect.dt1.Rows[0][6].ToString(),
                               PLC_GR = connect.dt1.Rows[0][7].ToString(),
                               EVKLASSIFIKATORNAME = connect.dt1.Rows[0][8].ToString(),
                               KKS = connect.dt1.Rows[0][9].ToString(),
                               POUNAME = connect.dt1.Rows[0][10].ToString(),
                               KLASSIFIKATORNAME = connect.dt1.Rows[0][11].ToString(),
                               PLC_varname = connect.dt1.Rows[0][12].ToString(),
                               PLC_address = connect.dt1.Rows[0][13].ToString(),
                               Channels = TeconObjectChannels.ToList(),
                               OldChannels = TeconObjectDefChannels.ToList(),
                               ObjTypeID = Convert.ToInt16(connect.dt1.Rows[0][14]),
                           };


                           ObjTypeID.Add(Convert.ToInt16(connect.dt1.Rows[0][14])); // заполняем список типами объектов*/

                        //!!!!!!!!                // Делаем короткий запрос для получения имени контроллера, ибо в едином запросе такое хз как сделать
                        //!!!!!!!!                connect.ConnectToBase("Select CARDS.MARKA from CARDS where CARDS.ID = " + obj.PLC_Name, BaseAddr);
                        //!!!!!!!!                obj.PLC_Name = connect.dt1.Rows[0][0].ToString();

                        //---------------------проверяем, нет ли повторяющихся имен каналов, создаем список--------------------------------------------------------------------------------


                        List<string> ChannelNames = new List<string>();
                        List<string> ChannelNamesDst = new List<string>();
                        foreach (ProgramReadDB.TeconObjectChannel to in TeconObjectChannels)
                        {
                            ChannelNames.Add(to.ChannelName);
                        }

                        ChannelNamesDst = ChannelNames.Distinct().ToList();
                        if (ChannelNamesDst.Count != ChannelNames.Count)
                        {
                            ObjTypeChannelsMatched.Add(obj.ObjTypeName);
                            //   MessageBox.Show("Обнаружены повторяющиеся имена каналов у типа: '");
                        }

                        /*
                        List<ProgramReadDB.TeconObjectChannel> TeconObjectChannelsDst = new List<ProgramReadDB.TeconObjectChannel>();
                        TeconObjectChannelsDst = TeconObjectChannels.Distinct().ToList();
                        if (TeconObjectChannelsDst.Count != obj.Channels.Count) 
                        {
                            ObjTypeChannelsMatched.Add(obj.ObjTypeName);
                            MessageBox.Show("Обнаружены повторяющиеся имена каналов у типа: '"  );
                        }*/
                        //-------------------------------------------------------------------------------------------------------------------------------------------------
                        TeconObjects.Add(obj);
                        TeconObjectChannels.Clear();
                        TeconObjectDefChannels.Clear();
                    }
                }
            }


            public void FindTypes()
        {
            ObjTypeIDUnic.Clear();
            ObjTypeCh.Clear(); // для повторной активации функции
            ObjTypeChannelsList.Clear();

        ObjTypeIDUnic = ObjTypeID.Distinct().ToList(); //убираем повторяющиеся типы объектов

        foreach (int id in ObjTypeIDUnic)
        {
            SQL_OBJTYPES = "Select OBJTYPEPARAM.NAME from OBJTYPEPARAM where OBJTYPEPARAM.PID = " + id;
            ProgramConnect connect = new ProgramConnect();
            connect.ConnectToBase(SQL_OBJTYPES, BaseAddr);

            for (int i = 0; i < connect.dt1.Rows.Count; i++)
            {
                ObjTypeChannelsList.Add(connect.dt1.Rows[i][0].ToString());
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

         }

            public void FullKlName()
            {
                string SQL_KLASS = "Select * from KLASSIFIKATOR";
                ProgramConnect connect = new ProgramConnect();
                connect.ConnectToBase(SQL_KLASS, BaseAddr);

                foreach (TeconObject obj in TeconObjects)
                {
                    string KlassPath = obj.KLASSIFIKATORNAME;
                    int TmpID = 0;

                    string st = "";
                    for (int i = 0; i < connect.dt1.Rows.Count; i++)
                    {
                        st = connect.dt1.Rows[i][2].ToString();   // выбираем НАЗВАНИЕ
                        if (st == obj.KLASSIFIKATORNAME) { TmpID = Convert.ToInt16(connect.dt1.Rows[i][1]); /*выбираем PID*/ break; }
                    }


                    while (TmpID != 0)
                    {
                        for (int i = 0; i < connect.dt1.Rows.Count; i++)
                        {
                            if ((Convert.ToInt16(connect.dt1.Rows[i][0]) == TmpID))
                            {
                                TmpID = Convert.ToInt16(connect.dt1.Rows[i][1]);
                                if (TmpID == 0) { break; }
                                KlassPath = Convert.ToString(connect.dt1.Rows[i][2]) + @"\" + KlassPath;
                                break;
                            }
                        }
                    }
                    obj.KLASSIFIKATORNAME = KlassPath;
                }

            }

        }

    
}

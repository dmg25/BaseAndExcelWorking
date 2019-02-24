using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;

using ExcelLibrary.BinaryDrawingFormat;
using ExcelLibrary.BinaryFileFormat;
using ExcelLibrary.SpreadSheet;
using ExcelLibrary.CompoundDocumentFormat;

namespace ConnectToSCADABD
{
    public class ProgramSaveExcel
    {
        public class HeadExcelCh
        {
            public string ChName;
            public string ChTitle;
            public string ChParam;
        }

        public class SaveTableParams       // класс для постраничного выбора параметров каналов
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
        public List<Worksheet> Sheets = new List<Worksheet>();  // список листов, перезапись не катит, так как они до сохранения висят в памяти.
        public List<HeadExcelCh> HeadExcelChList = new List<HeadExcelCh>(); //лист для шапки каналов
        public List<SaveTableParams> SaveTablesParamsList = new List<SaveTableParams>();

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

      //  ProgramReadDB ReadDB = new ProgramReadDB();
       
        public void SaveFileExcel(string FileName, bool WithoutChannel, int TablesNum, List<ProgramReadDB.TeconObject> TeconObjects)
        {
           
             string SaveXlsPath = FileName;

            Workbook workbook = new Workbook();

           // textBox1.AppendText("Начат процесс сохранения;\n");

           bool SaveWithoutCh = WithoutChannel;

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
                    foreach (ProgramReadDB.TeconObject to in TeconObjects)  //берем первое попавшееся имя объекта, у которого индекс совпадает
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


//-----------------------------блок создания шапки для каналов--------------------------------------------------------------------------------------

                HeadExcelChList.Clear();  //на всякий случай очистим шапку из предыдущего листа
                if (!SaveWithoutCh)
                {
                    foreach (ProgramReadDB.TeconObject TObj in TeconObjects)
                    {
                        //Создадим шапку таблицы и перечень используемых каналов

                        bool b = true; // первый параметр запишем без условий

                        if (TObj.Index == TblCount)
                        {

                            foreach (ProgramReadDB.TeconObjectChannel Tch in TObj.Channels)
                            {
                                if (Tch.S0 != "skipskipskip") { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "S0", ChTitle = "Шкала барогр. низ" }; HeadExcelChList.Add(obj); b = false; }
                                if (Tch.S100 != "skipskipskip") { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "S100", ChTitle = "Шкала барогр. верх" }; HeadExcelChList.Add(obj); b = false; }
                                if (Tch.M != "skipskipskip") { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "M", ChTitle = "Округлить до" }; HeadExcelChList.Add(obj); b = false; }
                                if (Tch.PLC_VARNAME != "skipskipskip") { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "PLC_VARNAME", ChTitle = "PLC_переменная" }; HeadExcelChList.Add(obj); b = false; }
                                if (Tch.ED_IZM != "skipskipskip") { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "ED_IZM", ChTitle = "Ед. изм." }; HeadExcelChList.Add(obj); b = false; }
                                if (Tch.DISC != "skipskipskip") { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "DISC", ChTitle = "Описание" }; HeadExcelChList.Add(obj); b = false; }
                                if (Tch.KA != "skipskipskip") { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "KA", ChTitle = "Коэф. КА" }; HeadExcelChList.Add(obj); b = false; }
                                if (Tch.KB != "skipskipskip") { var obj = new HeadExcelCh() { ChName = Tch.ChannelName, ChParam = "KB", ChTitle = "Коэф. КВ" }; HeadExcelChList.Add(obj); b = false; }
                            }
                        }
                    }

                    //Запишем шапку в Ecxel
                    //предварительно уберем из списка повторяющиеся элементы
                    var distinct = from item in HeadExcelChList
                                   group item by new { item.ChName, item.ChTitle, item.ChParam } into matches
                                   select matches.First();

                    HeadExcelChList = new List<HeadExcelCh>(distinct);

                    int k1 = 1;
                    foreach (HeadExcelCh hec in HeadExcelChList)
                    {
                        sheet.Cells[1, 15 + k1] = new Cell(hec.ChName);
                        sheet.Cells[2, 15 + k1] = new Cell(hec.ChTitle);
                        sheet.Cells[3, 15 + k1] = new Cell(hec.ChParam);
                        k1++;
                    }
                }
//-----------------------------конец блока создания шапки------------------------------------------------------------------------------------------




                bS0 = SaveTablesParamsList[TblCount - 1].S0;
                bS100 = SaveTablesParamsList[TblCount - 1].S100;
                bM = SaveTablesParamsList[TblCount - 1].M;
                bPLC_VARNAME = SaveTablesParamsList[TblCount - 1].PLC_VARNAME;
                bED_IZM = SaveTablesParamsList[TblCount - 1].ED_IZM;
                bARH_APP = SaveTablesParamsList[TblCount - 1].ARH_APP;
                bDISC = SaveTablesParamsList[TblCount - 1].DISC;
                bKA = SaveTablesParamsList[TblCount - 1].KA;
                bKB = SaveTablesParamsList[TblCount - 1].KB;

                //заполнение объектами
                int TmpCounter = 0;
                foreach (ProgramReadDB.TeconObject TObj in TeconObjects)
                {
                    if ((TObj.Index == TblCount) || (SaveWithoutCh))
                    {
                        sheet.Cells[TmpCounter + 4, 0] = new Cell(Convert.ToString(TmpCounter + 1));  // счетчик объектов
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
                            while (Convert.ToString(sheet.Cells[1, 15 + k]).Length > 0)  // если k не прибавится, будет бесконечный цикл!
                            {
                                Channel = Convert.ToString(sheet.Cells[1, 15 + k]);
                                ChannelParam = Convert.ToString(sheet.Cells[3, 15 + k]);
                                // MessageBox.Show(Convert.ToString(sheet.Cells[1, 15 + k]).Length.ToString() + "; k+15 = " + (k + 15).ToString());
                                preK = k;
                                for (int i = 0; i < TObj.Channels.Count; i++)
                                {
                                    /*TObj.Channels[i][ChannelParam]*/
                                    /*TObj.Channels[i].S0 == ChannelParam*/
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
                                if (preK == k) { MessageBox.Show("Не найден канал! Параметр" + ChannelParam + "; Канал: " + /*TObj.Channels[i].ChannelName */ "==" + Channel + "тип: " +TObj.ObjTypeName); /*break;*/ k++; }

                            }
                        }

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
        }
    }
}
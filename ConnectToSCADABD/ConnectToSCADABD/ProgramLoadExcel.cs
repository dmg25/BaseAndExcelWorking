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
    public class ProgramLoadExcel
    {
        string CellStr; //содержимое ячейки Excel
        int ExcelCellCnt; // счетчик ячеек Excel
        int CellInt; // переведённое в int содержимое ячейки
        bool FirstIter; // переменная для создания sql запроса, убирающая одну запятую
        public List<int> ObjID = new List<int>();  // массив с прочитанными ID из Excel

        public void LoadExcelFile(string FileName)
        {
            Workbook book = Workbook.Load(FileName);
            Worksheet sheet = book.Worksheets[0];

            ObjID.Clear();
            ExcelCellCnt = 0;
            FirstIter = false;
            CellStr = "1";
           
            if (Convert.ToString(sheet.Cells[3, 2]) != "ID")
            {
                MessageBox.Show("Выбран некорректный файл экспорта: Найдено: " + Convert.ToString(sheet.Cells[3, 2]) + " вместо ID!");
                return;
            }

            //!!!МОЖНО ДОБАВИТЬ ПОИСК ID ПО ВСЕЙ ДЛИНЕ ШАПКИ, ЕСЛИ НЕ ННАЙДЕНО В НУЖНОМ МЕСТЕ

            CellStr = Convert.ToString(sheet.Cells[4 + ExcelCellCnt, 2]); //[строка/столбец] 
            while (CellStr.Length > 0)  //будем читать столбец, пока не найдем пустую ячейку
            {                             
                CellInt = Convert.ToInt32(CellStr);
                ObjID.Add(CellInt);                
                ExcelCellCnt++;
                CellStr = Convert.ToString(sheet.Cells[4 + ExcelCellCnt, 2]);
            }
        }
    }
}
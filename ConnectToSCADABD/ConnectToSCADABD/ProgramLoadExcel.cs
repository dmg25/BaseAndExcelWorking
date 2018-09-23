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

            //Открываем книгу.                                                                                                                                                        
            //  Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false,Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).
            //   Excel.Worksheet ObjWorkSheet;
            //    ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            Worksheet sheet = book.Worksheets[0];

           // ReadDB.SQLParams = ""; //очищаем предыдущий поиск параметров
            ObjID.Clear();
            ExcelCellCnt = 0;
            FirstIter = false;
            CellStr = "1";
            // MessageBox.Show(CellStr);

            //Excel.Range forYach = ObjWorkSheet.Cells[4, 3] as Excel.Range;
            if (Convert.ToString(sheet.Cells[3, 2]) != "ID")
            {
                MessageBox.Show("Выбран некорректный файл экспорта: Найдено: " + Convert.ToString(sheet.Cells[3, 2]) + " вместо ID!");
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

           

        }
    }
}
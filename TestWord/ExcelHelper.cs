using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestWord
{
    class ExcelHelper
    {
        private FileInfo _fileInfo;
        string[,] list = new string[50, 5];

        public ExcelHelper(string fileName)
        {
            if (File.Exists(fileName))
            {
                _fileInfo = new FileInfo(fileName);
            }
            else
            {
                throw new ArgumentException("File not found");
            }

            
        }

        internal void Process()
        {
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(_fileInfo.FullName);
            //Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1-й лист
            Excel.Worksheet ObjWorkSheet = null;
            try
            {
                ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Worksheets["Титул"];
            }
            catch (Exception)
            {
                ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];//получить 1-й лист
            }
            if (ObjWorkSheet == null)
                return;
            
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//последнюю ячейку
                                                                                                // размеры базы
            int lastColumn = (int)lastCell.Column;
            int lastRow = (int)lastCell.Row;
            // Перенос в промежуточный массив класса Form1: string[,] list = new string[50, 5]; 
            for (int j = 0; j < 5; j++) //по всем колонкам
                for (int i = 0; i < lastRow; i++) // по всем строкам
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString(); //считываем данные

            //ВСЕ СТРОКИ С ДАННЫМИ
            int n = lastRow;
            string s;
            string st = "";
            for (int i = 0; i < n; i++) // по всем строкам
            {
                s = "";
                for (int j = 0; j < 5; j++) //по всем колонкам
                    s += " | " + list[i, j];
                st+=s+"\n";
            }

            //ПОИСК СЛОВ И ИХ ЗНАЧЕНИЙ
            string colToCheck = "A1:Z46";
            string DIRECTION = "направление";
            string PROFILE = "профиль";
            string QUALIFICATION = "квалификация";
            string FORM_STUDY = "форма обучения";
            string YEAR_START = "год начала подготовки (по учебному плану)";

            //List<string> items = new List<string>() { DIRECTION, PROFILE,  QUALIFICATION};
            Dictionary<string, string> items = new Dictionary<string, string>() {
                { DIRECTION, "" },
                { PROFILE, "" },
                { QUALIFICATION, "" },
                { FORM_STUDY, "" },
                { YEAR_START, "" },
            };


            Excel.Range resultRange;
            Excel.Range colRange = ObjWorkSheet.Range[colToCheck];//get the range object where you want to search from

            string address = "Строка не найдена";
            string value = "";

            try
            {
                foreach (var item in items)
                {
                    resultRange = colRange.Find(

                    What: item.Key,

                    LookIn: Excel.XlFindLookIn.xlValues,

                    LookAt: Excel.XlLookAt.xlPart,

                    SearchOrder: Excel.XlSearchOrder.xlByRows,

                    SearchDirection: Excel.XlSearchDirection.xlNext);


                    if (resultRange != null)
                    {
                        address = resultRange.Address.ToString();

                        string find_text = item.Key.ToString();
                        string result_text = resultRange.Value.ToString();
                        if (find_text.Length < result_text.Length)
                        {
                            if (result_text.Substring(item.Key.Length).StartsWith(": ")
                                || result_text.Substring(item.Key.Length).StartsWith("  "))
                                items[item.Key] = result_text.Substring(item.Key.Length + 2);

                            else if (result_text.Substring(item.Key.Length).StartsWith(":")
                                || result_text.Substring(item.Key.Length).StartsWith(" "))
                                items[item.Key] = result_text.Substring(item.Key.Length + 1);

                            else
                                items[item.Key] = result_text.Substring(item.Key.Length);
                        }
                        else
                        {
                            int _row = resultRange.Row;
                            int _column = (int)resultRange.Column;

                            //Проверка есть ли объединненые столбцы
                            if (resultRange.MergeCells)
                                _column = resultRange.Column + resultRange.MergeArea.Columns.Count;

                            string sValue = colRange.Cells[_row, _column].Value != null ? colRange.Cells[_row, _column].Value.ToString() : "пусто";

                            items[item.Key] = sValue;
                        }

                    }
                    MessageBox.Show(address + "\n_" + items[item.Key] + "_\n\n" + st);
                }
            }
            catch (Exception ex)
            {
                ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
                ObjWorkExcel.Quit(); // выйти из Excel
                GC.Collect(); // убрать за собой
                MessageBox.Show("Ошибка: " + ex.Message);
            }
            
            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из Excel
            GC.Collect(); // убрать за собой
        }
    }
}

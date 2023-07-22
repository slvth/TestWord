using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.ConstrainedExecution;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Shapes;
using DataTable = System.Data.DataTable;
using Word = Microsoft.Office.Interop.Word;

namespace TestWord
{
    class WordHelper
    {
        private FileInfo _fileInfo; 
        private Word.Application app;
        private Word._Document wordDocument;

        public WordHelper(string fileName)
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

        internal bool Process(Dictionary<string, string> items)
        {
            
            


            //СОЗДАНИЕ И ЗАПОЛНЕНИЕ ТАБЛИЦЫ
            var dt = new DataTable();
            dt.Columns.Add(new DataColumn("Picture ID", typeof(int)));
            dt.Columns.Add(new DataColumn("Title", typeof(string)));
            dt.Columns.Add(new DataColumn("Date Added", typeof(DateTime)));
            
            DataRow dr = dt.NewRow();
            dr["Picture ID"] = 1;
            dr["Title"] = "Earth";
            dr["Date Added"] = new DateTime();
            dt.Rows.Add(dr);

            try
            {
                Object file = _fileInfo.FullName;

                app = new Word.Application();
                wordDocument = app.Documents.Open(file, ReadOnly: false);

                replaceText(items);
                createTable1();
                createTable2();

                saveWord();

                app.ActiveDocument.Close();

                return true;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Ошибка! "+ex.Message+"\n"+
                    ex.StackTrace + "\n" + 
                    ex.TargetSite+"\n"+
                    ex.HelpLink);
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                }

                /*
                //открытие сохраненного файла
                if(pathMain!="")
                    System.Diagnostics.Process.Start(new ProcessStartInfo(@pathMain) { UseShellExecute = true });*/
            }
            return false;
        }

        private void saveWord()
        {   //путь и название будущего ФАЙЛА
            String name = DateTime.Now.ToString("dd-MM-yyyy HHmmss ") + _fileInfo.Name;
            String pathMain = "";

            //выбор пути и сохранение
            using (var path_dialog = new FolderBrowserDialog())
                if (path_dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    //Путь к директории
                    Object path = path_dialog.SelectedPath;

                    //Если нужно, через FileInfo можно получить другие данные
                    FileInfo targetDir = new FileInfo((string)path);

                    string pathToFolder = targetDir.FullName + "";
                    string name_folder = targetDir.Name;

                    pathMain = pathToFolder.ToString() + "\\" + name;

                    Object newFileName = System.IO.Path.Combine(@pathToFolder.ToString(), name);
                    app.ActiveDocument.SaveAs2(newFileName);

                    System.Windows.MessageBox.Show("Успешное сохранение!");
                };

            //открытие сохраненного файла
            if (pathMain != "")
                System.Diagnostics.Process.Start(new ProcessStartInfo(@pathMain) { UseShellExecute = true });
        }

        private void replaceText(Dictionary<string, string> items) 
        {
            Object missing = Type.Missing;

            //замена простого текст
            foreach (var item in items)
            {
                Word.Find find = app.Selection.Find;
                find.Text = item.Key;
                find.Replacement.Text = item.Value;

                Object wrap = Word.WdFindWrap.wdFindContinue;
                Object replace = Word.WdReplace.wdReplaceAll;

                find.Execute(FindText: Type.Missing,
                    MatchCase: false,
                    MatchWholeWord: false,
                MatchWildcards: false,
                    MatchSoundsLike: missing,
                    MatchAllWordForms: false,
                    Forward: true,
                    Wrap: wrap,
                Format: false,
                    ReplaceWith: missing, Replace: replace);
            }
        }
        private void createTable1()
        {
            //СОЗДАНИЕ И ЗАПОЛНЕНИЕ ТАБЛИЦЫ
            var dt = new DataTable();
            dt.Columns.Add(new DataColumn("Picture ID", typeof(int)));
            dt.Columns.Add(new DataColumn("Title", typeof(string)));
            dt.Columns.Add(new DataColumn("Date Added", typeof(DateTime)));

            DataRow dr = dt.NewRow();
            dr["Picture ID"] = 1;
            dr["Title"] = "Earth";
            dr["Date Added"] = new DateTime();
            dt.Rows.Add(dr);

            //вставка таблицы
            app.Selection.Find.Execute("<TABLE1>");
            Word.Range wordRange = app.Selection.Range;

            var wordTable = wordDocument.Tables.Add(wordRange,
                dt.Rows.Count, dt.Columns.Count);
            wordTable.Borders.Enable = 1;
            wordTable.Columns.Width = 100;
            //wordTable.Columns[1].Cells[1].Column.Cells.Borders.OutsideColor = WdColor.wdColorDarkRed;
            //wordTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            wordTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            wordTable.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth225pt;
            //..Borders.InsideColor = Word.WdColor.wdColorAqua;

            for (var j = 0; j < dt.Rows.Count; j++)
            {
                for (var k = 0; k < dt.Columns.Count; k++)
                {
                    wordTable.Cell(j + 1, k + 1).Range.Text = dt.Rows[j][k].ToString();
                }
            }
        }

        private void createTable2()
        {
            ///!!!!!!!!!!!!!!
            //вставка таблицы2
            //СОЗДАНИЕ И ЗАПОЛНЕНИЕ ТАБЛИЦЫ
            var dt2 = new DataTable();
            dt2.Columns.Add(new DataColumn("Номер", typeof(string)));
            dt2.Columns.Add(new DataColumn("Тема", typeof(string)));
            dt2.Columns.Add(new DataColumn("Семестр", typeof(string)));
            dt2.Columns.Add(new DataColumn("Лекции", typeof(string)));
            dt2.Columns.Add(new DataColumn("Практические", typeof(string)));
            dt2.Columns.Add(new DataColumn("Лабораторные", typeof(string)));
            dt2.Columns.Add(new DataColumn("СРС", typeof(string)));
            DataRow dr2 = dt2.NewRow();

            /*
            dr2["Номер"] = "№п/п";
            dr2["Тема"] = "Тема дисциплины";
            dr2["Семестр"] =  "семестр";
            dr2["ВРЕМЯ"] = "2";
            dr2["СРС"] = "4";*/

            dt2.Rows.Add(dr2);
            dr2 = dt2.NewRow();

            /*
            dr2["Номер"] = "f";
            dr2["Тема"] = "Тема 11. Основы метрологии";
            dr2["Семестр"] = "4";
            dr2["Лекции"] = "2";
            dr2["Практические"] = "4";
            dr2["Лабораторные"] = "4";
            dr2["СРС"] = "4";*/

            dt2.Rows.Add(dr2);
            dr2 = dt2.NewRow();

            /*
             ....
            */

            dt2.Rows.Add(dr2);
            dr2 = dt2.NewRow();

            app.Selection.Find.Execute("<TABLE2>");
            Word.Range wordRange2 = app.Selection.Range;

            var wordTable2 = wordDocument.Tables.Add(wordRange2,
                dt2.Rows.Count, dt2.Columns.Count);
            wordTable2.Borders.Enable = 1;



            //At this point, rng is at the start of the first (left-most) cell of the two
            //using new objects for the split cells

            //wordTable.Columns[1].Cells[1].Column.Cells.Borders.OutsideColor = WdColor.wdColorDarkRed;
            //wordTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            //wordTable2.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            //wordTable2.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth225pt;
            //..Borders.InsideColor = Word.WdColor.wdColorAqua;

            for (var j = 0; j < dt2.Rows.Count; j++)
            {
                for (var k = 0; k < dt2.Columns.Count; k++)
                {
                    //wordTable2.Cell(j + 1, k + 1).Range.Text = dt2.Rows[j][k].ToString();
                }
            }

            /*
            Cell cell = wordTable2.Cell(1, 4);
            Word.Range rng = cell.Range;
            cell.Merge(wordTable2.Cell(1, 5));
            cell.Merge(wordTable2.Cell(1, 5));
            Word.Range rng2 = cell.Range;
            Word.Cell newCel1 = rng2.Cells[1];
            Word.Cell newCel2 = rng2.Next(1, 1).Cells[1];
            newCel1.Range.Text = "Первый";
            newCel2.Range.Text = "ВТОРОЙ";*/

            wordTable2.Cell(1, 4).Merge(wordTable2.Cell(1, 5));
            wordTable2.Cell(1, 4).Merge(wordTable2.Cell(1, 5));

            wordTable2.Cell(1, 1).Range.Text = "№ п/п";
            wordTable2.Cell(1, 2).Range.Text = "Темы дисциплины";
            wordTable2.Cell(1, 3).Range.Text = "семестр";
            wordTable2.Cell(1, 4).Range.Text = "Виды и часы " +
                "контактной \nработы, \nих трудоемкость \n(в часах)";
            wordTable2.Cell(1, 5).Range.Text = "СРС";

            //направление текста
            wordTable2.Cell(1, 1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable2.Cell(1, 2).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable2.Cell(1, 3).Range.Orientation = WdTextOrientation.wdTextOrientationUpward;
            wordTable2.Cell(1, 3).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable2.Cell(1, 5).Range.Orientation = WdTextOrientation.wdTextOrientationUpward;
            wordTable2.Cell(1, 5).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            wordTable2.Cell(1, 1).Merge(wordTable2.Cell(2, 1));
            wordTable2.Cell(1, 2).Merge(wordTable2.Cell(2, 2));
            wordTable2.Cell(1, 3).Merge(wordTable2.Cell(2, 3));
            wordTable2.Cell(1, 5).Merge(wordTable2.Cell(2, 7));

            //wordTable2.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed);


            float width_column1, width_column2, width_column3,
                width_column4, width_column5,
                width_column6, width_column7, point;

            point = 28.35f;
            width_column1 = 1.13f * point;
            width_column2 = 7.83f * point;
            width_column3 = 1.8f * point;
            width_column4 = 1.51f * point;
            width_column5 = 1.67f * point;
            width_column6 = 1.66f * point;
            width_column7 = 1.18f * point;


            wordTable2.Cell(1, 1).Width = width_column1;
            wordTable2.Cell(1, 2).Width = width_column2;
            wordTable2.Cell(1, 3).Width = width_column3;
            wordTable2.Cell(1, 4).Width = 4.84f * 28.35f;
            wordTable2.Cell(1, 5).Width = width_column7;
            wordTable2.Cell(1, 4).Height = 1.21f * 28.35f;
            /*
            wordTable2.Cell(1, 1).Width = 1.13f * 28.35f;
            wordTable2.Cell(1, 2).Width = 7.83f * 28.35f;
            wordTable2.Cell(1, 3).Width = 1.8f * 28.35f;
            wordTable2.Cell(1, 4).Width = 4.84f * 28.35f;
            wordTable2.Cell(1, 5).Width = 1.18f * 28.35f;
            wordTable2.Cell(1, 4).Height = 1.21f * 28.35f;
            */


            /*

            wordTable2.Columns[0].Width = 1.13f * 28.35f;
            wordTable2.Columns[1].Width = 7.83f * 28.35f;
            wordTable2.Columns[2].Width = 1.8f * 28.35f;
            wordTable2.Columns[3].Width = 4.84f * 28.35f;
            wordTable2.Columns[4].Width = 1.19f * 28.35f;
            wordTable2.Rows[1].Height = 1.21f * 28.35f;*/


            wordTable2.Cell(2, 4).Range.Text = "Лекции";
            wordTable2.Cell(2, 5).Range.Text = "Практические занятия";
            wordTable2.Cell(2, 6).Range.Text = "Лабораторные занятия";

            //направление текста
            wordTable2.Cell(2, 4).Range.Orientation = WdTextOrientation.wdTextOrientationUpward;
            wordTable2.Cell(2, 4).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable2.Cell(2, 5).Range.Orientation = WdTextOrientation.wdTextOrientationUpward;
            wordTable2.Cell(2, 5).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable2.Cell(2, 6).Range.Orientation = WdTextOrientation.wdTextOrientationUpward;
            wordTable2.Cell(2, 6).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;


            wordTable2.Cell(2, 4).Width = width_column4;
            wordTable2.Cell(2, 5).Width = width_column5;
            wordTable2.Cell(2, 6).Width = width_column6;
            wordTable2.Cell(2, 7).Width = width_column7;
            /*
            wordTable2.Cell(2, 4).Width = 1.51f * 28.35f;
            wordTable2.Cell(2, 5).Width = 1.67f * 28.35f;
            wordTable2.Cell(2, 6).Width = 1.66f * 28.35f;
            wordTable2.Cell(2, 7).Width = 1.18f * 28.35f;
            */



            wordTable2.Cell(2, 5).Height = 3.31f * 28.35f;


            wordTable2.Cell(3, 1).Range.Text = "3.1";
            wordTable2.Cell(3, 2).Range.Text = "3.2";
            wordTable2.Cell(3, 3).Range.Text = "3.3";
            wordTable2.Cell(3, 4).Range.Text = "3.4";
            wordTable2.Cell(3, 5).Range.Text = "3.5";
            wordTable2.Cell(3, 6).Range.Text = "3.6";
            wordTable2.Cell(3, 7).Range.Text = "3.7";



            wordTable2.Cell(3, 1).Width = width_column1;
            wordTable2.Cell(3, 2).Width = width_column2;
            wordTable2.Cell(3, 3).Width = width_column3;
            wordTable2.Cell(3, 4).Width = width_column4;
            wordTable2.Cell(3, 5).Width = width_column5;
            wordTable2.Cell(3, 6).Width = width_column6;
            wordTable2.Cell(3, 7).Width = width_column7;


            wordTable2.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            wordTable2.Range.ParagraphFormat.SpaceBefore = 0;
            wordTable2.Range.ParagraphFormat.SpaceAfter = 0;
            wordTable2.Cell(1, 5).Merge(wordTable2.Cell(2, 7));
        }
    }


}

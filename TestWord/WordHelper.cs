using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Windows.Documents;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;
using Word = Microsoft.Office.Interop.Word;

namespace TestWord
{
    class WordHelper
    {
        private FileInfo _fileInfo; 
        private Word.Application app;
        private Word._Document wordDocument;

        //данные
        DisciplineModel discipline;
        List<ThemeModel> themes;

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
            try
            {
                Object file = _fileInfo.FullName;

                app = new Word.Application();
                wordDocument = app.Documents.Open(file, ReadOnly: false);

                getData();
                replaceText(items);
                createTable1();
                createTable2();
                createTable3();
                createTable4();

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
            }
            return false;
        }

        private void getData()
        {
            //данные - код, имя, знать, уметь, владеть, индикаторы
            List<CompetenceModel> competences = new List<CompetenceModel>(){
                new CompetenceModel("ОПК-11",
                    "Способен проводить научные эксперименты с использованием современного исследовательского оборудования и приборов, оценивать результаты исследований",
                    "Фундаментальные физические законы, константы и эффекты, используемые при измерениях, физические ограничения точности измерений, международную систему единиц величин и основные теории размерностей",
                    "Применять методы и средства измерений для решения измерительных задач",
                    "Способами расчёта погрешностей измерений",
                    new List<CompetenceModel>(){
                        new CompetenceModel("ОПК-11.1", "Знает фундаментальные физические законы, константы и эффекты, используемые при измерениях, физические ограничения точности измерений, международную систему единиц величин и основные теории размерностей"),
                        new CompetenceModel("ОПК-11.3", "Умеет применять методы и средства измерений для решения измерительных задач"),
                        new CompetenceModel("ОПК-11.4", "Владеет навыками работы  используемых средств измерения и контроля технологических процессов и   способами расчёта погрешностей измерений"),
                }),
                new CompetenceModel("ОПК-12",
                    "Способен проводить научные эксперименты с использованием современного исследовательского оборудования и приборов, оценивать результаты исследований",
                    "Фундаментальные физические законы, константы и эффекты, используемые при измерениях, физические ограничения точности измерений, международную систему единиц величин и основные теории размерностей",
                    "Применять методы и средства измерений для решения измерительных задач",
                    "Способами расчёта погрешностей измерений",
                    new List<CompetenceModel>(){
                        new CompetenceModel("ОПК-12.1", "Знает фундаментальные физические законы, константы и эффекты, используемые при измерениях, физические ограничения точности измерений, международную систему единиц величин и основные теории размерностей"),
                        new CompetenceModel("ОПК-12.3", "Умеет применять методы и средства измерений для решения измерительных задач"),
                        new CompetenceModel("ОПК-12.4", "Владеет навыками работы  используемых средств измерения и контроля технологических процессов и   способами расчёта погрешностей измерений"),
                }),
            };

            themes = new List<ThemeModel>() {
                new ThemeModel("Тема 1. Основы метрологии", 8, 1, 1, 2, 3, 4,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("История развития метрологии", 2, ""),
                    },
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Измерение линейных размеров с помощью штангенинструментов", 2, ""),
                        new ChildThemeModel("Электрические измерения напряжения", 2, ""),
                    },
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Системы физических единиц", 2, ""),
                        new ChildThemeModel("Размерность физических единиц", 2, "")
                    }
                ),
                new ThemeModel("Тема 2. Средства и методы измерения", 8, 1, 5, 6, 7, 8,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Виды и методы измерений", 2, "Проблемная лекция"),
                        new ChildThemeModel("Средства измерений", 2, "Лекция с запланированными ошибками"),
                    },
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Поверка СИ температуры", 2, ""),
                        new ChildThemeModel("Проверка средств измерения давления", 2, ""),
                        new ChildThemeModel("Аттестация средств измерения давления", 2, ""),
                    },
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Температурные шкалы", 2, ""),
                        new ChildThemeModel("Метрологические характеристики средств измерения", 2, "работа в малых группах")
                    }
                ),
                new ThemeModel("Тема 3. Погрешности измерения", 8, 2, 9, 10, 11, 12,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Основы метрологического обеспечения производства", 2, ""),
                        new ChildThemeModel("Понятие о погрешности измерений", 2, ""),
                    },
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Определение метрологических характеристик средств измерения", 2, ""),
                        new ChildThemeModel("Влияние газового фактора на точность измерений", 2, ""),
                        new ChildThemeModel("Определение погрешностей СИ при изменении характеристики среды", 2, ""),
                        new ChildThemeModel("Влияние не стабильности потока на точность измерения", 2, ""),
                    },
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Определение погрешностей измерения", 2, "групповое обсуждение"),
                        new ChildThemeModel("Погрешности косвенных измерений", 2, ""),
                        new ChildThemeModel("Определение доверительных границ и доверительных интервалов", 2, "работа в малых группах"),
                    }
                ),
                new ThemeModel("Тема 4. Основы стандартизации", 8, 2, 13, 14, 15, 16,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("История развития стандартизации", 2, "Лекция-визуализация"),
                        new ChildThemeModel("Методы и средства стандартизации", 2, ""),
                    },
                    null,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Нормативно-правовые документы по стандартизации", 2, ""),
                    }
                ),
                new ThemeModel("Тема 5. Основы сертификации", 8, 2, 17, 18, 19, 20,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Основные понятия сертификации", 2, ""),
                    },
                    null,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Сходства и отличия «Сертификация соответствия» и «Декларирование соответствия»", 2, ""),
                    }
                ),
            };

            List<int> semesters = new List<int>() { 8 };
            int total_lecture_hour = 16;
            int total_practical_hour = 18;
            int total_laboratory_hour = 18;
            int total_independent_hour = 20;

            discipline = new DisciplineModel(
                semesters,
                total_lecture_hour,
                total_laboratory_hour,
                total_practical_hour,
                total_independent_hour,
                themes,
                competences
                );
        }

        private class DisciplineModel
        {
            public List<int> semesters { get; set; }

            public int total_lecture_hour { get; set; }
            public int total_practical_hour { get; set; }
            public int total_laboratory_hour { get; set; }
            public int total_independent_hour { get; set; }
            public List<ThemeModel> themes { get; set; }
            public List<CompetenceModel> competences { get; set; }

            public DisciplineModel(
                List<int> semesters, 
                int total_lecture_hour,
                int total_laboratory_hour,
                int total_practical_hour, 
                int total_independent_hour, 
                List<ThemeModel> themes, List<CompetenceModel> competences)
            {
                this.semesters = semesters;
                this.total_lecture_hour = total_lecture_hour;
                this.total_laboratory_hour = total_laboratory_hour;
                this.total_practical_hour = total_practical_hour;
                this.total_independent_hour = total_independent_hour;
                this.themes = themes;
                this.competences = competences;
            }
        }

        private class ThemeModel
        {
            public string theme { get; set; }
            public int semester { get; set; }
            public int module { get; set; }
            public int lecture_hour { get; set; }
            public int practical_hour { get; set; }
            public int laboratory_hour { get; set; }
            public int independent_hour { get; set; }



            public List<ChildThemeModel>? lectures { get; set; }
            public List<ChildThemeModel>? practicals { get; set; }
            public List<ChildThemeModel>? laboratories { get; set; }

            public ThemeModel(string theme, int semester, int module,
                int lecture_hour, int practical_hour,
                int laboratory_hour, int independent_hour)
            {
                this.theme = theme;
                this.semester = semester;
                this.module = module;
                this.lecture_hour = lecture_hour;
                this.practical_hour = practical_hour;
                this.laboratory_hour = laboratory_hour;
                this.independent_hour = independent_hour;
            }

            public ThemeModel(string theme, int semester, int module,
                int lecture_hour, int practical_hour,
                int laboratory_hour, int independent_hour, 
                List<ChildThemeModel>? lectures, 
                List<ChildThemeModel>? laboratories, 
                List<ChildThemeModel>? practicals)
            {
                this.theme = theme;
                this.semester = semester;
                this.module = module;
                this.lecture_hour = lecture_hour;
                this.practical_hour = practical_hour;
                this.laboratory_hour = laboratory_hour;
                this.independent_hour = independent_hour;
                this.lectures = lectures;
                this.laboratories = laboratories;
                this.practicals = practicals;
            }
        }

        private class ChildThemeModel
        {
            public string name { get; set; }
            public int hour { get; set; }
            public string method { get; set; }

            public ChildThemeModel(string name, int hour, string method)
            {
                this.name = name;
                this.hour = hour;
                this.method = method;
            }
        }

        private class CompetenceModel
        {
            public string kod { get; set; }
            public string name { get; set; }
            public string? know { get; set; }
            public string? able { get; set; }
            public string? own { get; set; }

            public List<CompetenceModel>? childs { get; set; }

            public CompetenceModel(string kod, string name)
            {
                this.kod = kod;
                this.name = name;
            }

            public CompetenceModel(string kod, string name, string know, string able, string own, List<CompetenceModel> childs)
            {
                this.kod = kod;
                this.name = name;
                this.know = know;
                this.able = able;
                this.own = own;
                this.childs = childs;
            }
        }


        private void saveWord()
        {   
            //путь и название будущего ФАЙЛА
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
                    ReplaceWith: missing, Replace: replace
                    );
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

            for (int i = 0; i < themes.Count+3; i++)
            {
                DataRow dr2 = dt2.NewRow();
                dt2.Rows.Add(dr2);
            }

            app.Selection.Find.Execute("<TABLE2>");
            Word.Range wordRange2 = app.Selection.Range;

            var wordTable2 = wordDocument.Tables.Add(wordRange2,
                dt2.Rows.Count, dt2.Columns.Count);
            wordTable2.Borders.Enable = 1;


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

            //объединение ячеек
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


            //ширина, высоты столбцов
            wordTable2.Cell(1, 1).Width = width_column1;
            wordTable2.Cell(1, 2).Width = width_column2;
            wordTable2.Cell(1, 3).Width = width_column3;
            wordTable2.Cell(1, 4).Width = 4.84f * 28.35f;
            wordTable2.Cell(1, 5).Width = width_column7;
            wordTable2.Cell(1, 4).Height = 1.21f * 28.35f;


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

            //ширина, высоты столбцов
            wordTable2.Cell(2, 4).Width = width_column4;
            wordTable2.Cell(2, 5).Width = width_column5;
            wordTable2.Cell(2, 6).Width = width_column6;
            wordTable2.Cell(2, 7).Width = width_column7;
            wordTable2.Cell(2, 5).Height = 3.31f * 28.35f;


            int countItems = themes.Count;

            for (int i = 0; i < countItems; i++)
            {
                wordTable2.Cell(3 + i, 1).Range.Text = (i+1).ToString();
                wordTable2.Cell(3 + i, 2).Range.Text = themes[i].theme;
                wordTable2.Cell(3 + i, 3).Range.Text = themes[i].semester.ToString();
                wordTable2.Cell(3 + i, 4).Range.Text = themes[i].lecture_hour.ToString();
                wordTable2.Cell(3 + i, 5).Range.Text = themes[i].practical_hour.ToString();
                wordTable2.Cell(3 + i, 6).Range.Text = themes[i].laboratory_hour.ToString();
                wordTable2.Cell(3 + i, 7).Range.Text = themes[i].independent_hour.ToString();

                wordTable2.Cell(3 + i, 1).Width = width_column1;
                wordTable2.Cell(3 + i, 2).Width = width_column2;
                wordTable2.Cell(3 + i, 3).Width = width_column3;
                wordTable2.Cell(3 + i, 4).Width = width_column4;
                wordTable2.Cell(3 + i, 5).Width = width_column5;
                wordTable2.Cell(3 + i, 6).Width = width_column6;
                wordTable2.Cell(3 + i, 7).Width = width_column7;

                //выравнивание=слева
                wordTable2.Cell(3 + i, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            }

            //последняя строка
            wordTable2.Cell(3 + countItems, 1).Range.Text = "";
            wordTable2.Cell(3 + countItems, 2).Range.Text = "Итого по дисциплине";
            wordTable2.Cell(3 + countItems, 3).Range.Text = "";
            wordTable2.Cell(3 + countItems, 4).Range.Text = "16";
            wordTable2.Cell(3 + countItems, 5).Range.Text = "18";
            wordTable2.Cell(3 + countItems, 6).Range.Text = "18";
            wordTable2.Cell(3 + countItems, 7).Range.Text = "20";

            wordTable2.Cell(3 + countItems, 1).Width = width_column1;
            wordTable2.Cell(3 + countItems, 2).Width = width_column2;
            wordTable2.Cell(3 + countItems, 3).Width = width_column3;
            wordTable2.Cell(3 + countItems, 4).Width = width_column4;
            wordTable2.Cell(3 + countItems, 5).Width = width_column5;
            wordTable2.Cell(3 + countItems, 6).Width = width_column6;
            wordTable2.Cell(3 + countItems, 7).Width = width_column7;

            //форматирование текста
            wordTable2.Cell(3 + countItems, 2).Range.Bold = Convert.ToInt32(true);


            //форматирование таблицы
            wordTable2.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            wordTable2.Range.ParagraphFormat.SpaceBefore = 0;
            wordTable2.Range.ParagraphFormat.SpaceAfter = 0;
            wordTable2.Cell(1, 5).Merge(wordTable2.Cell(2, 7));
        }



        private void createTable3()
        {
            var dt = new DataTable();
            dt.Columns.Add(new DataColumn("Оцениваемые компетенции", typeof(string)));
            dt.Columns.Add(new DataColumn("Код и наименование индикатора", typeof(string)));
            dt.Columns.Add(new DataColumn("Результаты освоения", typeof(string)));
            dt.Columns.Add(new DataColumn("Оценочные средства", typeof(string)));

            //данные - код, имя, знать, уметь, владеть, индикаторы
            List<CompetenceModel> competences = discipline.competences;

            for (int i = 0; i < competences.Count + 1; i++)
            {
                DataRow dr = dt.NewRow();
                dt.Rows.Add(dr);
            }

            app.Selection.Find.Execute("<TABLE3>");
            Word.Range wordRange = app.Selection.Range;

            var wordTable = wordDocument.Tables.Add(wordRange,
                dt.Rows.Count, dt.Columns.Count);
            

            wordTable.Cell(1, 1).Range.Text = "Оцениваемые компетенции (код, наименование)";
            wordTable.Cell(1, 2).Range.Text = "Код и наименование индикатора (индикаторов) достижения компетенции";
            wordTable.Cell(1, 3).Range.Text = "Результаты освоения компетенции";
            wordTable.Cell(1, 4).Range.Text = "Оценочные средства текущего контроля и промежуточной аттестации";

            //заполнение данными
            for (int i = 0;i < competences.Count;i++)
            {
                CompetenceModel item = competences[i];
                string kod = item.kod; 
                string name = item.name;
                string know = item.know;
                string able = item.able;
                string own = item.own;
                List<CompetenceModel> childs = item.childs;

                //Столбец1
                Word.Range range = wordTable.Cell(2+i, 1).Range;
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range.InsertAfter(kod);
                range.Font.Bold = Convert.ToInt32(true);
                range.InsertParagraphAfter();
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range.InsertAfter(name);
                range.Font.Bold = Convert.ToInt32(false);
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                //Столбец2
                Word.Range range2 = wordTable.Cell(2 + i, 2).Range;
                for (int j = 0; j < childs.Count; j++)
                {
                    CompetenceModel child = childs[j];
                    string kod_child = child.kod;
                    string name_child = child.name;

                    range2.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range2.InsertAfter(kod_child+".");
                    range2.Font.Bold = Convert.ToInt32(true);
                    range2.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    range2.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                    if (j == childs.Count - 1)
                    {
                        range2.InsertAfter(" "+name_child + ".");
                    }
                    else
                    {
                        range2.InsertAfter(" "+name_child + ";");
                        range2.InsertParagraphAfter();
                    }
                    range2.Font.Bold = Convert.ToInt32(false);
                    range2.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }

                //Столбец3
                Word.Range range3 = wordTable.Cell(2 + i, 3).Range;
                //знать
                range3.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range3.InsertAfter("Знать:");
                range3.Font.Bold = Convert.ToInt32(true);
                range3.InsertParagraphAfter();
                range3.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                range3.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range3.InsertAfter(know.ToLower() + ";");
                range3.Font.Bold = Convert.ToInt32(false);
                range3.InsertParagraphAfter();
                range3.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                //уметь
                range3.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range3.InsertAfter("Уметь:");
                range3.Font.Bold = Convert.ToInt32(true);
                range3.InsertParagraphAfter();
                range3.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                
                range3.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range3.InsertAfter(able.ToLower() + ";");
                range3.Font.Bold = Convert.ToInt32(false);
                range3.InsertParagraphAfter();
                range3.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                //владеть
                range3.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range3.InsertAfter("Владеть:");
                range3.Font.Bold = Convert.ToInt32(true);
                range3.InsertParagraphAfter();
                range3.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                range3.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range3.InsertAfter(own.ToLower() + ".");
                range3.Font.Bold = Convert.ToInt32(false);
                range3.InsertParagraphAfter();
                range3.Collapse(Word.WdCollapseDirection.wdCollapseEnd);


                //!!!!!!!!!!!!!!!!ДОРАБОТАТЬ - данные брать из excel или программы
                //Столбец4
                Word.Range range4 = wordTable.Cell(2 + i, 4).Range;
                //текущий контроль
                range4.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range4.InsertAfter("Текущий контроль:");
                range4.Font.Bold = Convert.ToInt32(true);
                range4.InsertParagraphAfter();
                range4.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                range4.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range4.InsertAfter("Компьютерное тестирование по теме 1-5\nПрактические задачи по темам 1-5\nЛабораторные работыпо темам 1-3"); //!!!!!
                range4.Font.Bold = Convert.ToInt32(false);
                range4.InsertParagraphAfter();
                range4.InsertParagraphAfter();
                range4.InsertParagraphAfter();
                range4.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                //промежуточная аттестация
                range4.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range4.InsertAfter("Промежуточная аттестация:");
                range4.Font.Bold = Convert.ToInt32(true);
                range4.InsertParagraphAfter();
                range4.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                range4.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range4.InsertAfter("Экзамен"); //!!!!!
                range4.Font.Bold = Convert.ToInt32(false);
                range4.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            }

            float width_column1, width_column2, width_column3,
                width_column4, point;


            //форматирование таблицы
            wordTable.Borders.Enable = 1;
            wordTable.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            wordTable.Range.ParagraphFormat.SpaceBefore = 0;
            wordTable.Range.ParagraphFormat.SpaceAfter = 0;
            wordTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

            wordTable.Cell(1, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordTable.Cell(1, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordTable.Cell(1, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordTable.Cell(1, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            wordTable.Cell(1, 1).Range.Bold = Convert.ToInt32(true);
            wordTable.Cell(1, 2).Range.Bold = Convert.ToInt32(true);
            wordTable.Cell(1, 3).Range.Bold = Convert.ToInt32(true);
            wordTable.Cell(1, 4).Range.Bold = Convert.ToInt32(true);
        }

        private void createTable4()
        {
            string main_key = "<TABLE4>";
            app.Selection.Find.Execute(main_key);
            Word.Range wordRange = app.Selection.Range;
            Dictionary<string, string> text_keys = new Dictionary<string, string>();
            Dictionary<string, int> table_keys = new Dictionary<string, int>();

            //вставка тэгов в ворд
            foreach (var semester in discipline.semesters)
            {
                string semester_text = "<SEMESTER" + semester + ">";
                string semester_table = "<SEMESTER_TABLE" + semester + ">";
                text_keys.Add(semester_text, semester + " cеместр");
                table_keys.Add(semester_table, semester);

                wordRange.InsertAfter(semester_text);
                wordRange.InsertParagraphAfter();
                wordRange.InsertAfter(semester_table);
            }
            //удаление main_key из word
            app.Selection.Find.Execute(main_key);
            Word.Range wordRangeDelete = app.Selection.Range;
            wordRangeDelete.Delete();

            //замена тэгов на слова семестров
            replaceText(text_keys);

            string[] columns = {
                "Тема",
                "Кол-во часов",
                "Используемый метод",
                "Формируемые компетенции" };

            /*
            int count_theme = discipline.themes.Where(a=>a.semester.Equals(semester));
            int count_module = 2 * discipline.semesters.Count;
            for (int i = 0; i < count_theme + count_module; i++)
                dt.Rows.Add();
            */

            //замена тэгов на таблицы семестров
            foreach (var semester in table_keys)
            {
                var dt = new DataTable();
                foreach (string column in columns)
                    dt.Columns.Add(column);

                int count_theme = discipline.themes.Where(a => a.semester.Equals(semester.Value)).Count();
                int count_theme_lecture = discipline.themes.Where(a => a.semester.Equals(semester.Value) && a.lectures is not null).SelectMany(a=>a.lectures).ToList().Count;
                int count_theme_laboratory = discipline.themes.Where(a => a.semester.Equals(semester.Value) && a.laboratories is not null).SelectMany(a => a.laboratories).ToList().Count;
                int count_theme_practical = discipline.themes.Where(a => a.semester.Equals(semester.Value) && a.practicals is not null).SelectMany(a => a.practicals).ToList().Count;
                int total_row = count_theme + count_theme_lecture + count_theme_laboratory + count_theme_practical;
                
                for (int i = 0; i < total_row+3; i++)
                    dt.Rows.Add();

                app.Selection.Find.Execute(semester.Key);
                Word.Range wordRangeTable = app.Selection.Range;
                var wordTable = wordDocument.Tables.Add(wordRangeTable,
                    dt.Rows.Count, dt.Columns.Count);
                wordTable.Cell(1, 1).Range.Text = columns[0];
                wordTable.Cell(1, 2).Range.Text = columns[1];
                wordTable.Cell(1, 3).Range.Text = columns[2];
                wordTable.Cell(1, 4).Range.Text = columns[3];

                /*
                //форматирование
                float point, width_column1;
                point = 28.35f;
                width_column1 = 8.73f * point;
                //wordTable.Columns[1].Width = width_column1;*/
                
                


                bool first_write1 = true;
                bool first_write2 = true;
                int current_row = 2;
                int current_lecture = 1;
                int current_laboratories = 1;
                int current_practical = 1;

                for (int i = 0; i < discipline.themes.Count; i++) 
                {
                    if (semester.Value.Equals(discipline.themes[i].semester)) 
                    {
                        if (discipline.themes[i].module == 1)
                        {
                            if (first_write1)
                            {
                                wordTable.Cell(current_row, 1).Range.Text = "Дисциплинарный модуль " + semester.Value + ".1";
                                wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row, 4));
                                //форматирование
                                wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);

                                current_row++;
                                first_write1 = false;
                            }
                            List<ChildThemeModel>? lectures = discipline.themes[i].lectures;
                            List<ChildThemeModel>? laboratories = discipline.themes[i].laboratories;
                            List<ChildThemeModel>? practicals = discipline.themes[i].practicals;

                            wordTable.Cell(current_row, 1).Range.Text = "Тема " + (i + 1) + " " + discipline.themes[i].theme;
                            wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row, 4));
                            //форматирование
                            wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
                            current_row++;

                            if (lectures is not null)
                                foreach (var lecture in lectures)
                                {
                                    wordTable.Cell(current_row, 1).Range.Text = "Лекция " + current_lecture + ". " + lecture.name;
                                    wordTable.Cell(current_row, 2).Range.Text = lecture.hour.ToString();
                                    wordTable.Cell(current_row, 3).Range.Text = lecture.method;
                                    wordTable.Cell(current_row, 4).Range.Text = string.Join(", ", discipline.competences.Select(a => a.kod));
                                    
                                    //форматирование
                                    wordTable.Cell(current_row, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                                   
                                    //переход на новую строку
                                    current_row++;
                                    current_lecture++;
                                }

                            if (laboratories is not null)
                                foreach (var lab in laboratories)
                                {
                                    wordTable.Cell(current_row, 1).Range.Text = "Лабораторная работа  " + current_laboratories + ". " + lab.name;
                                    wordTable.Cell(current_row, 2).Range.Text = lab.hour.ToString();
                                    wordTable.Cell(current_row, 3).Range.Text = lab.method;
                                    wordTable.Cell(current_row, 4).Range.Text = string.Join(", ", discipline.competences.Select(a => a.kod));
                                    //форматирование
                                    wordTable.Cell(current_row, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                                    current_row++;
                                    current_laboratories++;
                                }

                            if (practicals is not null)
                                foreach (var practical in practicals)
                                {
                                    wordTable.Cell(current_row, 1).Range.Text = "Практическое занятие " + current_practical + ". " + practical.name;
                                    wordTable.Cell(current_row, 2).Range.Text = practical.hour.ToString();
                                    wordTable.Cell(current_row, 3).Range.Text = practical.method;
                                    wordTable.Cell(current_row, 4).Range.Text = string.Join(", ", discipline.competences.Select(a => a.kod));
                                    //форматирование
                                    wordTable.Cell(current_row, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                                    current_row++;
                                    current_practical++;
                                }




                        }
                        else if(discipline.themes[i].module == 2)
                        {
                            if (first_write2)
                            {
                                wordTable.Cell(current_row, 1).Range.Text = "Дисциплинарный модуль " + semester.Value + ".2";
                                wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row, 4));
                                //форматирование
                                wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
                                current_row++;
                                first_write2 = false;
                            }
                            List<ChildThemeModel>? lectures = discipline.themes[i].lectures;
                            List<ChildThemeModel>? laboratories = discipline.themes[i].laboratories;
                            List<ChildThemeModel>? practicals = discipline.themes[i].practicals;

                            wordTable.Cell(current_row, 1).Range.Text = "Тема " + (i + 1) + " " + discipline.themes[i].theme;
                            wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row, 4));
                            //форматирование
                            wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
                            current_row++;
                            
                            if (lectures is not null)
                                foreach (var lecture in lectures)
                                {
                                    wordTable.Cell(current_row, 1).Range.Text = "Лекция " + current_lecture + ". " + lecture.name;
                                    wordTable.Cell(current_row, 2).Range.Text = lecture.hour.ToString();
                                    wordTable.Cell(current_row, 3).Range.Text = lecture.method;
                                    wordTable.Cell(current_row, 4).Range.Text = string.Join(", ", discipline.competences.Select(a => a.kod));
                                    //форматирование
                                    wordTable.Cell(current_row, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                                    current_row++;
                                    current_lecture++;
                                }

                            if(laboratories is not null)
                                foreach (var lab in laboratories)
                                {
                                    wordTable.Cell(current_row, 1).Range.Text = "Лабораторная работа  " + current_laboratories + ". " + lab.name;
                                    wordTable.Cell(current_row, 2).Range.Text = lab.hour.ToString();
                                    wordTable.Cell(current_row, 3).Range.Text = lab.method;
                                    wordTable.Cell(current_row, 4).Range.Text = string.Join(", ", discipline.competences.Select(a => a.kod));
                                    //форматирование
                                    wordTable.Cell(current_row, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                                    current_row++;
                                    current_laboratories++;
                                }

                            if (practicals is not null)
                                foreach (var practical in practicals)
                                {
                                    wordTable.Cell(current_row, 1).Range.Text = "Практическое занятие " + current_practical + ". " + practical.name;
                                    wordTable.Cell(current_row, 2).Range.Text = practical.hour.ToString();
                                    wordTable.Cell(current_row, 3).Range.Text = practical.method;
                                    wordTable.Cell(current_row, 4).Range.Text = string.Join(", ", discipline.competences.Select(a => a.kod));
                                    //форматирование
                                    wordTable.Cell(current_row, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                                    current_row++;
                                    current_practical++;
                                }
                            
                        }
                    }
                }

                //форматирование таблицы
                wordTable.Borders.Enable = Convert.ToInt32(true);
                wordTable.Range.ParagraphFormat.SpaceBefore = 0;
                wordTable.Range.ParagraphFormat.SpaceAfter = 0;
                wordTable.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);

                wordTable.Cell(1, 1).Range.Bold = Convert.ToInt32(true);
                wordTable.Cell(1, 2).Range.Bold = Convert.ToInt32(true);
                wordTable.Cell(1, 3).Range.Bold = Convert.ToInt32(true);
                wordTable.Cell(1, 4).Range.Bold = Convert.ToInt32(true);
            }


            


            /*
            //заполнение данными
            for (int i = 0; i < competences.Count; i++)
            {
                CompetenceModel item = competences[i];
                string kod = item.kod;
                string name = item.name;
                string know = item.know;
                string able = item.able;
                string own = item.own;
                List<CompetenceModel> childs = item.childs;

                //Столбец1
                Word.Range range = wordTable.Cell(2 + i, 1).Range;
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range.InsertAfter(kod);
                range.Font.Bold = Convert.ToInt32(true);
                range.InsertParagraphAfter();
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range.InsertAfter(name);
                range.Font.Bold = Convert.ToInt32(false);
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            }

            wordTable.Borders.Enable = Convert.ToInt32(true);*/
        }

    }
}

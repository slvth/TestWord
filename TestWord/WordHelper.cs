using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
                //wordDocument = app.Documents.Open(file, ReadOnly: false);

                getData();
                replaceText(items);
                createTable1();
                createTable2();
                createTable3();
                createTable4();
                createTable5();
                createTable6();
                createTable7();

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
                if (app is not null)
                {
                    app.Quit();
                }
            }
            return false;
        }

        private void getData()
        {
            //6.1 Перечень оценочных средств - текущий контроль
            List<EvaluationToolModel> controls = new List<EvaluationToolModel>()
            {
                new EvaluationToolModel("лабораторная работа",
                    "Темы, задания для выполнения лабораторных работ; вопросы к их защите",
                    "Может выполняться в индивидуальном порядке или группой обучающихся. Задания в лабораторных работах должны включать элемент командной работы. Позволяет оценить умения, обучающихся самостоятельно конструировать свои знания в процессе решения практических задач и оценить уровень сформированности аналитических, исследовательских навыков, а также навыков практического мышления. Позволяет оценить способность к профессиональным трудовым действиям"
                ),
                new EvaluationToolModel("Практическая задача",
                    "Комплект задач и заданий",
                    "Средство оценки умения применять полученные теоретические знания в практической ситуации. Задача должна быть направлена на оценивание тех компетенций, которые подлежат освоению в данной дисциплине, должна содержать четкую инструкцию по выполнению или алгоритм действий"
                ),
                new EvaluationToolModel("Тестирование компьютерное",
                    "Фонд тестовых заданий",
                    "Система стандартизированных заданий, позволяющая автоматизировать процедуру измерения уровня знаний и умений, обучающегося по соответствующим компетенциям. Обработка результатов тестирования на компьютере обеспечивается специальными программами. Позволяет проводить самоконтроль (репетиционное тестирование), может выступать в роли тренажера при подготовке к зачету или экзамену"
                ),
            };
            //6.1 Перечень оценочных средств - промежуточная аттестация
            List<EvaluationToolModel> attestations = new List<EvaluationToolModel>()
            {
                new EvaluationToolModel("Экзамен",
                    "Перечень вопросов, фонд тестовых заданий",
                    "Итоговая форма определения степени достижения запланированных результатов обучения (оценивания уровня освоения компетенций). Экзамен нацелен на комплексную проверку освоения дисциплины. Экзамен проводится в форме тестирования по всем темам дисциплины"
                ),
            };

            //данные - код, имя, знать, уметь, владеть, индикаторы
            //1 Перечень планируемых результатов обучения по дисциплине
            //6.2 Уровень освоения компетенций и критерии оценивания результатов обучения
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
                    },
                    new List<ResultDescModel>(){
                        //знать 5,4,3,2
                        new ResultDescModel(
                            "Сформированные систематические представления об фундаментальных физических законах, константах и эффектах, используемых при измерениях, физических ограничениях точности измерений, международной системе единиц величин и основных теориях размерностей",
                            "Сформированные, но содержащие отдельные пробелы представления об фундаментальных физических законах, константах и эффектах, используемых при измерениях, физических ограничениях точности измерений, международной системе единиц величин и основных теориях размерностей",
                            "Неполные представления об фундаментальных физических законах, константах и эффектах, используемых при измерениях, физических ограничениях точности измерений, международной системе единиц величин и основных теориях размерностей",
                            "Фрагментарные представления об фундаментальных физических законах, константах и эффектах, используемых при измерениях, физических ограничениях точности измерений, международной системе единиц величин и основных теориях размерностей"),
                        //уметь 5,4,3,2
                        new ResultDescModel(
                            "Сформированное умение применять методы и средства измерений для решения измерительных задач",
                            "В целом успешное, но содержащее отдельные пробелы умение применять методы и средства измерений для решения измерительных задач",
                            "В целом успешное, но не систематическое умение применять методы и средства измерений для решения измерительных задач",
                            "Фрагментарное умение применять методы и средства измерений для решения измерительных задач"),
                        //владеть 5,4,3,2
                        new ResultDescModel(
                            "Успешное и систематическое владение способами расчёта погрешностей измерений",
                            "В целом успешное, но содержащее отдельные пробелы владение способами расчёта погрешностей измерений",
                            "В целом успешное, но не систематическое владение способами расчёта погрешностей измерений",
                            "Фрагментарное владение способами расчёта погрешностей измерений")
                    }
                ),
                new CompetenceModel("ОПК-12",
                    "Способен проводить научные эксперименты с использованием современного исследовательского оборудования и приборов, оценивать результаты исследований",
                    "Фундаментальные физические законы, константы и эффекты, используемые при измерениях, физические ограничения точности измерений, международную систему единиц величин и основные теории размерностей",
                    "Применять методы и средства измерений для решения измерительных задач",
                    "Способами расчёта погрешностей измерений",
                    new List<CompetenceModel>(){
                        new CompetenceModel("ОПК-12.1", "Знает фундаментальные физические законы, константы и эффекты, используемые при измерениях, физические ограничения точности измерений, международную систему единиц величин и основные теории размерностей"),
                        new CompetenceModel("ОПК-12.3", "Умеет применять методы и средства измерений для решения измерительных задач"),
                        new CompetenceModel("ОПК-12.4", "Владеет навыками работы  используемых средств измерения и контроля технологических процессов и   способами расчёта погрешностей измерений"),
                    },
                    new List<ResultDescModel>(){
                        //знать 5,4,3,2
                        new ResultDescModel(
                            "Сформированные систематические представления об фундаментальных физических законах, константах и эффектах, используемых при измерениях, физических ограничениях точности измерений, международной системе единиц величин и основных теориях размерностей",
                            "Сформированные, но содержащие отдельные пробелы представления об фундаментальных физических законах, константах и эффектах, используемых при измерениях, физических ограничениях точности измерений, международной системе единиц величин и основных теориях размерностей",
                            "Неполные представления об фундаментальных физических законах, константах и эффектах, используемых при измерениях, физических ограничениях точности измерений, международной системе единиц величин и основных теориях размерностей",
                            "Фрагментарные представления об фундаментальных физических законах, константах и эффектах, используемых при измерениях, физических ограничениях точности измерений, международной системе единиц величин и основных теориях размерностей"),
                        //уметь 5,4,3,2
                        new ResultDescModel(
                            "Сформированное умение применять методы и средства измерений для решения измерительных задач",
                            "В целом успешное, но содержащее отдельные пробелы умение применять методы и средства измерений для решения измерительных задач",
                            "В целом успешное, но не систематическое умение применять методы и средства измерений для решения измерительных задач",
                            "Фрагментарное умение применять методы и средства измерений для решения измерительных задач"),
                        //владеть 5,4,3,2
                        new ResultDescModel(
                            "Успешное и систематическое владение способами расчёта погрешностей измерений",
                            "В целом успешное, но содержащее отдельные пробелы владение способами расчёта погрешностей измерений",
                            "В целом успешное, но не систематическое владение способами расчёта погрешностей измерений",
                            "Фрагментарное владение способами расчёта погрешностей измерений")
                    }
                ),
            };

            themes = new List<ThemeModel>() {
                new ThemeModel("Тема 1. Основы метрологии", 8, 1, 2, 4, 4, 4,
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
                new ThemeModel("Тема 2. Средства и методы измерения", 8, 1, 4, 4, 6, 4,
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
                new ThemeModel("Тема 3. Погрешности измерения", 8, 2, 4, 6, 8, 4,
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
                new ThemeModel("Тема 4. Основы стандартизации", 8, 2, 4, 2, 0, 4,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("История развития стандартизации", 2, "Лекция-визуализация"),
                        new ChildThemeModel("Методы и средства стандартизации", 2, ""),
                    },
                    null,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Нормативно-правовые документы по стандартизации", 2, ""),
                    }
                ),
                new ThemeModel("Тема 5. Основы сертификации", 8, 2, 2, 2, 0, 4,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Основные понятия сертификации", 2, ""),
                    },
                    null,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Сходства и отличия «Сертификация соответствия» и «Декларирование соответствия»", 2, ""),
                    }
                ),
            };

            List<int> semesters = new List<int>() { 4 };
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

            discipline.controls = controls;
            discipline.attestations = attestations;
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

            public List<EvaluationToolModel> controls { get; set; }
            public List<EvaluationToolModel> attestations { get; set; }

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

        private class EvaluationToolModel
        {
            public string name { get; set; }
            public string description { get; set; }
            public string path { get; set; }

            public EvaluationToolModel(string name, string path, string description)
            {
                this.name = name;
                this.path = path;
                this.description = description;
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
            public string kod { get; set; } //компетенция, например ОПК11
            public string name { get; set; }

            public string? know { get; set; }
            public string? able { get; set; }
            public string? own { get; set; }

            public List<CompetenceModel>? childs { get; set; } //индикаторы, например ОПК11.1, ОПК11.2
            public List<ResultDescModel>? results { get; set; } //6.2 знать, уметь, владеть описание по оценкам 

            public CompetenceModel(string kod, string name) //для индикаторов
            {
                this.kod = kod;
                this.name = name;
            }

            public CompetenceModel(string kod, string name, string know, string able, string own, List<CompetenceModel> childs, List<ResultDescModel> results)
            {
                this.kod = kod;
                this.name = name;
                this.know = know;
                this.able = able;
                this.own = own;
                this.childs = childs;
                this.results = results;
            }
        }

        private class ResultDescModel
        {
            public string result5 { get; set; }
            public string result4 { get; set; }
            public string result3 { get; set; }
            public string result2 { get; set; }

            public ResultDescModel(string result5, string result4, string result3, string result2)
            {
                this.result5 = result5;
                this.result4 = result4;
                this.result3 = result3;
                this.result2 = result2;
            }
        }

        //6.3.1.2 Тестирование компьютерное
        private class QuestionModel
        {
            public string name { get; set; }
            public List<QuestionModel> answers { get; set; }
            public List<string>? competence { get; set; }
            public int? semester { get; set; }
            public int? module { get; set; }

            public QuestionModel(string name, int semester, int module, 
                List<QuestionModel> answers)
            {
                this.name = name;
                this.semester = semester;
                this.module = module;
                this.answers = answers;
            }

            public QuestionModel(string name, List<string> competence)
            {
                this.name = name;
                this.competence = competence;
            }

            public QuestionModel(string name)
            {
                this.name = name;
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
            var dt2 = new DataTable();
            dt2.Columns.Add(new DataColumn("Номер", typeof(string)));
            dt2.Columns.Add(new DataColumn("Тема", typeof(string)));
            dt2.Columns.Add(new DataColumn("Семестр", typeof(string)));
            dt2.Columns.Add(new DataColumn("Лекции", typeof(string)));
            dt2.Columns.Add(new DataColumn("Практические", typeof(string)));
            dt2.Columns.Add(new DataColumn("Лабораторные", typeof(string)));
            dt2.Columns.Add(new DataColumn("СРС", typeof(string)));

            for (int i = 0; i < themes.Count+3; i++)
                dt2.Rows.Add();

            app.Selection.Find.Execute("<TABLE2>");
            Word.Range wordRange2 = app.Selection.Range;
            var wordTable2 = wordDocument.Tables.Add(wordRange2,
                dt2.Rows.Count, dt2.Columns.Count);

            //форматирование
            for (int i = 1; i <= 2; i++)
                for (int j = 1; j <= dt2.Columns.Count; j++)
                    wordTable2.Cell(i, j).Range.Bold = Convert.ToInt32(true);
            wordTable2.Cell(1, 4).Merge(wordTable2.Cell(1, 5));
            wordTable2.Cell(1, 4).Merge(wordTable2.Cell(1, 5));

            //заполнение шаблона
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

            //заполнение шаблона
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
                wordTable2.Cell(3 + i, 4).Range.Text = themes[i].lecture_hour !=0 ? themes[i].lecture_hour.ToString() : "-";
                wordTable2.Cell(3 + i, 5).Range.Text = themes[i].practical_hour != 0 ? themes[i].practical_hour.ToString() : "-";
                wordTable2.Cell(3 + i, 6).Range.Text = themes[i].laboratory_hour != 0 ? themes[i].laboratory_hour.ToString() : "-";
                wordTable2.Cell(3 + i, 7).Range.Text = themes[i].independent_hour != 0 ? themes[i].independent_hour.ToString() : "-";

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
            wordTable2.Cell(3 + countItems, 4).Range.Text = discipline.total_lecture_hour.ToString();
            wordTable2.Cell(3 + countItems, 5).Range.Text = discipline.total_practical_hour.ToString();
            wordTable2.Cell(3 + countItems, 6).Range.Text = discipline.total_laboratory_hour.ToString();
            wordTable2.Cell(3 + countItems, 7).Range.Text = discipline.total_independent_hour.ToString();

            wordTable2.Cell(3 + countItems, 1).Width = width_column1;
            wordTable2.Cell(3 + countItems, 2).Width = width_column2;
            wordTable2.Cell(3 + countItems, 3).Width = width_column3;
            wordTable2.Cell(3 + countItems, 4).Width = width_column4;
            wordTable2.Cell(3 + countItems, 5).Width = width_column5;
            wordTable2.Cell(3 + countItems, 6).Width = width_column6;
            wordTable2.Cell(3 + countItems, 7).Width = width_column7;
            
            for (int i = 1; i <= 7; i++)
                wordTable2.Cell(3 + countItems, i).Range.Bold = Convert.ToInt32(true);

            //форматирование таблицы
            wordTable2.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            wordTable2.Range.ParagraphFormat.SpaceBefore = 0;
            wordTable2.Range.ParagraphFormat.SpaceAfter = 0;
            wordTable2.Borders.Enable = 1;
            //Столбец СРС
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
                text_keys.Add(semester_text, $"Семестр {semester}");
                table_keys.Add(semester_table, semester);

                wordRange.InsertAfter(semester_text);
                wordRange.InsertParagraphAfter();
                wordRange.InsertAfter(semester_table);
            }
            //удаление main_key из word
            app.Selection.Find.Execute(main_key);
            Word.Range wordRangeDelete = app.Selection.Range;
            wordRangeDelete.Delete();
            
            //форматирование слова Семестр {0}
            foreach (var item in text_keys)
            {
                app.Selection.Find.Execute(item.Key);
                Word.Range range = app.Selection.Range;
                range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                range.Bold = Convert.ToInt32(true);
                range.Font.Size = 14;
            }

            //замена тэгов на слова семестров
            replaceText(text_keys);

            string[] columns = {
                "Тема",
                "Кол-во часов",
                "Используемый метод",
                "Формируемые компетенции" };

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

                            //загрузка данных в таблицу
                            int[] currents = loadDataTable4(wordTable, current_lecture, current_laboratories, current_practical, current_row, i);
                            current_lecture = currents[0];
                            current_laboratories = currents[1];
                            current_practical = currents[2];
                            current_row = currents[3];
                            
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

                            //загрузка данных в таблицу
                            int[] currents = loadDataTable4(wordTable, current_lecture, current_laboratories, current_practical, current_row, i);
                            current_lecture = currents[0];
                            current_laboratories = currents[1];
                            current_practical = currents[2];
                            current_row = currents[3];
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
        }
        //загрузка данных в таблицу4
        //чтобы убрать повтор кода
        private int[] loadDataTable4(Word.Table wordTable,
            int current_lecture, int current_laboratories, 
            int current_practical, int current_row, int i)
        {
            List<ChildThemeModel>? lectures = discipline.themes[i].lectures;
            List<ChildThemeModel>? laboratories = discipline.themes[i].laboratories;
            List<ChildThemeModel>? practicals = discipline.themes[i].practicals;

            int hour_lecture = lectures is not null ? lectures.Select(a => a.hour).Sum() : 0;
            int hour_lab = laboratories is not null ? laboratories.Select(a => a.hour).Sum() : 0;
            int hour_practical = practicals is not null ? practicals.Select(a => a.hour).Sum() : 0;
            int total_hour = hour_lecture + hour_lab + hour_practical;

            wordTable.Cell(current_row, 1).Range.Text = $"Тема {i + 1}. {discipline.themes[i].theme} ({total_hour} ч.)";
            wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row, 4));
            //форматирование
            wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
            current_row++;

            if (lectures is not null)
                foreach (var lecture in lectures)
                {
                    //Столбец 1 - Тема
                    Word.Range rangeColumn1 = wordTable.Cell(current_row, 1).Range;
                    rangeColumn1.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    rangeColumn1.InsertAfter("Лекция " + current_lecture + ".");
                    rangeColumn1.Font.Italic = Convert.ToInt32(true);
                    rangeColumn1.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    rangeColumn1.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    rangeColumn1.InsertAfter(" " + lecture.name);
                    rangeColumn1.Font.Italic = Convert.ToInt32(false);
                    rangeColumn1.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    
                    //Столбец 2 - Кол-во часов
                    wordTable.Cell(current_row, 2).Range.Text = lecture.hour.ToString();
                    
                    //Столбец 3 - Используемый метод
                    wordTable.Cell(current_row, 3).Range.Text = lecture.method;

                    //Столбец 4 - Формируемые компетенции
                    wordTable.Cell(current_row, 4).Range.Text = string.Join(", ", discipline.competences.Select(a => a.kod));

                    //форматирование
                    wordTable.Cell(current_row, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    wordTable.Cell(current_row, 3).Range.Italic = Convert.ToInt32(true);

                    //переход на новую строку
                    current_row++;
                    current_lecture++;
                }

            if (laboratories is not null)
                foreach (var lab in laboratories)
                {
                    Word.Range rangeColumn1 = wordTable.Cell(current_row, 1).Range;
                    rangeColumn1.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    rangeColumn1.InsertAfter("Лабораторная работа  " + current_laboratories + ".");
                    rangeColumn1.Font.Italic = Convert.ToInt32(true);
                    rangeColumn1.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    rangeColumn1.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    rangeColumn1.InsertAfter(" " + lab.name);
                    rangeColumn1.Font.Italic = Convert.ToInt32(false);
                    rangeColumn1.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    wordTable.Cell(current_row, 2).Range.Text = lab.hour.ToString();
                    wordTable.Cell(current_row, 3).Range.Text = lab.method;
                    wordTable.Cell(current_row, 4).Range.Text = string.Join(", ", discipline.competences.Select(a => a.kod));
                    
                    //форматирование
                    wordTable.Cell(current_row, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    wordTable.Cell(current_row, 3).Range.Italic = Convert.ToInt32(true);

                    current_row++;
                    current_laboratories++;
                }

            if (practicals is not null)
                foreach (var practical in practicals)
                {
                    Word.Range rangeColumn1 = wordTable.Cell(current_row, 1).Range;
                    rangeColumn1.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    rangeColumn1.InsertAfter("Практическое занятие " + current_practical + ".");
                    rangeColumn1.Font.Italic = Convert.ToInt32(true);
                    rangeColumn1.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    rangeColumn1.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    rangeColumn1.InsertAfter(" " + practical.name);
                    rangeColumn1.Font.Italic = Convert.ToInt32(false);
                    rangeColumn1.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    wordTable.Cell(current_row, 2).Range.Text = practical.hour.ToString();
                    wordTable.Cell(current_row, 3).Range.Text = practical.method;
                    wordTable.Cell(current_row, 4).Range.Text = string.Join(", ", discipline.competences.Select(a => a.kod));
                    
                    //форматирование
                    wordTable.Cell(current_row, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    wordTable.Cell(current_row, 3).Range.Italic = Convert.ToInt32(true);

                    current_row++;
                    current_practical++;
                }
            int[] currents = {current_lecture, current_laboratories, current_practical, current_row};
            return currents;
        }

        private void createTable5()
        {
            var dt = new DataTable();
            dt.Columns.Add(new DataColumn("Этап", typeof(string)));
            dt.Columns.Add(new DataColumn("Название", typeof(string)));
            dt.Columns.Add(new DataColumn("Описание", typeof(string)));
            dt.Columns.Add(new DataColumn("Представление", typeof(string)));

            //данные - код, имя, знать, уметь, владеть, индикаторы
            List<EvaluationToolModel> controls = discipline.controls;
            List<EvaluationToolModel> attestations = discipline.attestations;

            int total_row = controls.Count + attestations.Count;
            for (int i = 0; i < total_row+3; i++)
                dt.Rows.Add();

            app.Selection.Find.Execute("<TABLE5>");
            Word.Range wordRange = app.Selection.Range;
            var wordTable = wordDocument.Tables.Add(wordRange,
                dt.Rows.Count, dt.Columns.Count);

            wordTable.Cell(1, 1).Range.Text = "Этапы формирования компетенции";
            wordTable.Cell(1, 2).Range.Text = "Вид оценочного средства";
            wordTable.Cell(1, 3).Range.Text = "Краткая характеристика оценочного средства";
            wordTable.Cell(1, 4).Range.Text = "Представление оценочного средства в фонде";
            wordTable.Cell(2, 1).Range.Text = "Текущий контроль";

            //форматирование
            wordTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordTable.Cell(2, 1).Range.Bold = Convert.ToInt32(true);
            wordTable.Cell(2, 1).Merge(wordTable.Cell(2, 4));

            int current_row = 3;
            int stage = 1;
            foreach (var control in controls)
            {
                wordTable.Cell(current_row, 1).Range.Text = stage.ToString();
                wordTable.Cell(current_row, 2).Range.Text = control.name;
                wordTable.Cell(current_row, 3).Range.Text = control.description;
                wordTable.Cell(current_row, 4).Range.Text = control.path;
                //форматирование
                wordTable.Cell(current_row, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordTable.Cell(current_row, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                wordTable.Cell(current_row, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                //переход на новую строку
                stage++;
                current_row++;
            }

            wordTable.Cell(current_row, 1).Range.Text = "Промежуточная аттестация";
            wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
            wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row, 4));
            current_row++;

            foreach (var attestation in attestations)
            {
                wordTable.Cell(current_row, 1).Range.Text = stage.ToString();
                wordTable.Cell(current_row, 2).Range.Text = attestation.name;
                wordTable.Cell(current_row, 3).Range.Text = attestation.description;
                wordTable.Cell(current_row, 4).Range.Text = attestation.path;
                //форматирование
                wordTable.Cell(current_row, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordTable.Cell(current_row, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                wordTable.Cell(current_row, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                //переход на новую строку
                stage++;
                current_row++;
            }

            //форматирование
            wordTable.Borders.Enable = Convert.ToInt32(true);
            wordTable.Range.ParagraphFormat.SpaceBefore = 0;
            wordTable.Range.ParagraphFormat.SpaceAfter = 0;
            wordTable.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
        }

        //6.2 Уровень освоения компетенции и критерии оценивания результатов обучения
        private void createTable6()
        {
            List<CompetenceModel> competences = discipline.competences;


            var dt = new DataTable();
            dt.Columns.Add("номер");
            dt.Columns.Add("код");
            dt.Columns.Add("индикатор");
            dt.Columns.Add("результаты");
            dt.Columns.Add("продвинутый");
            dt.Columns.Add("средний");
            dt.Columns.Add("базовый");
            dt.Columns.Add("компетенции не освоены");

            for (int i = 1; i <= (competences.Count * 3)+4; i++)
                dt.Rows.Add();

            app.Selection.Find.Execute("<TABLE6>");
            Word.Range wordRange = app.Selection.Range;
            var wordTable = wordDocument.Tables.Add(wordRange, dt.Rows.Count, dt.Columns.Count);
            //форматирование
            wordTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            for (int i = 1; i <= 4; i++)
                for (int j = 1; j <= wordTable.Columns.Count; j++)
                {
                    wordTable.Cell(i, j).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    wordTable.Cell(i, j).Range.Bold = Convert.ToInt32(true);
                }

            //шапка таблицы
            //строка1
            wordTable.Cell(1, 1).Range.Text = "№ п/п";
            wordTable.Cell(1, 2).Range.Text = "Оцениваемые компетенции (код, наименование)";
            wordTable.Cell(1, 3).Range.Text = "Код и наименование индикатора (индикаторов) достижения компетенции";
            wordTable.Cell(1, 4).Range.Text = "Планируемые результаты обучения";
            wordTable.Cell(1, 5).Range.Text = "Уровень освоения компетенций";
            //строка2
            wordTable.Cell(2, 5).Range.Text = "Продвинутый уровень";
            wordTable.Cell(2, 6).Range.Text = "Средний уровень";
            wordTable.Cell(2, 7).Range.Text = "Базовый уровень";
            wordTable.Cell(2, 8).Range.Text = "Компетенции не освоены";
            //строка3
            wordTable.Cell(3, 5).Range.Text = "Критерии оценивания результатов обучения";
            //строка4
            wordTable.Cell(4, 5).Range.Text = "«отлично»\n(от 86 до 100 баллов)";
            wordTable.Cell(4, 6).Range.Text = "«хорошо»\n(от 71 до 85 баллов)";
            wordTable.Cell(4, 7).Range.Text = "«удовлетворительно»\n(от 55 до 70 баллов)\n";
            wordTable.Cell(4, 8).Range.Text = "«неудовлетв.»\n(менее 55 баллов)";
            //объединение строк, столбцов
            wordTable.Cell(1, 1).Merge(wordTable.Cell(4, 1));
            wordTable.Cell(1, 2).Merge(wordTable.Cell(4, 2));
            wordTable.Cell(1, 3).Merge(wordTable.Cell(4, 3));
            wordTable.Cell(1, 4).Merge(wordTable.Cell(4, 4));
            wordTable.Cell(1, 5).Merge(wordTable.Cell(1, 8));
            wordTable.Cell(3, 5).Merge(wordTable.Cell(3, 8));

            int current_row = 5;
            for (int i = 0; i < competences.Count; i++)
            {
                CompetenceModel current_competence = competences[i];
                string kod = current_competence.kod;
                string name = current_competence.name;
                string know = current_competence.know;
                string able = current_competence.able;
                string own = current_competence.own;
                List<CompetenceModel> childs = current_competence.childs;

                //Столбец1
                wordTable.Cell(current_row, 1).Range.Text = (i + 1).ToString();

                //Столбец2
                Word.Range range2 = wordTable.Cell(current_row, 2).Range;
                range2.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range2.InsertAfter(kod);
                range2.Font.Bold = Convert.ToInt32(true);
                range2.InsertParagraphAfter();
                range2.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range2.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range2.InsertAfter(name);
                range2.Font.Bold = Convert.ToInt32(false);
                range2.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                //Столбец3
                Word.Range range3 = wordTable.Cell(current_row, 3).Range;
                for (int j = 0; j < childs.Count; j++)
                {
                    CompetenceModel child = childs[j];
                    string kod_child = child.kod;
                    string name_child = child.name;

                    range3.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range3.InsertAfter(kod_child + ".");
                    range3.Font.Bold = Convert.ToInt32(true);
                    range3.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    range3.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                    if (j == childs.Count - 1)
                    {
                        range3.InsertAfter(" " + name_child + ".");
                    }
                    else
                    {
                        range3.InsertAfter(" " + name_child + ";");
                        range3.InsertParagraphAfter();
                    }
                    range3.Font.Bold = Convert.ToInt32(false);
                    range3.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }

                //Столбец4
                //знать
                Word.Range range4_1 = wordTable.Cell(current_row, 4).Range;
                range4_1.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range4_1.InsertAfter("Знать:");
                range4_1.Font.Bold = Convert.ToInt32(true);
                range4_1.InsertParagraphAfter();
                range4_1.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                range4_1.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range4_1.InsertAfter(know.ToLower() + ";");
                range4_1.Font.Bold = Convert.ToInt32(false);
                range4_1.InsertParagraphAfter();
                range4_1.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                //уметь
                Word.Range range4_2 = wordTable.Cell(current_row + 1, 4).Range;
                range4_2.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range4_2.InsertAfter("Уметь:");
                range4_2.Font.Bold = Convert.ToInt32(true);
                range4_2.InsertParagraphAfter();
                range4_2.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                range4_2.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range4_2.InsertAfter(able.ToLower() + ";");
                range4_2.Font.Bold = Convert.ToInt32(false);
                range4_2.InsertParagraphAfter();
                range4_2.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                //владеть
                Word.Range range4_3 = wordTable.Cell(current_row + 2, 4).Range;
                range4_3.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range4_3.InsertAfter("Владеть:");
                range4_3.Font.Bold = Convert.ToInt32(true);
                range4_3.InsertParagraphAfter();
                range4_3.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                range4_3.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range4_3.InsertAfter(own.ToLower() + ".");
                range4_3.Font.Bold = Convert.ToInt32(false);
                range4_3.InsertParagraphAfter();
                range4_3.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                //Столбец5-8
                List<ResultDescModel> results = current_competence.results;
                ResultDescModel knowResult = results[0];
                ResultDescModel ableResult = results[1];
                ResultDescModel ownResult = results[2];
                //знать
                wordTable.Cell(current_row, 5).Range.Text = knowResult.result5;
                wordTable.Cell(current_row, 6).Range.Text = knowResult.result4;
                wordTable.Cell(current_row, 7).Range.Text = knowResult.result3;
                wordTable.Cell(current_row, 8).Range.Text = knowResult.result2;
                //уметь
                wordTable.Cell(current_row + 1, 5).Range.Text = ableResult.result5;
                wordTable.Cell(current_row + 1, 6).Range.Text = ableResult.result4;
                wordTable.Cell(current_row + 1, 7).Range.Text = ableResult.result3;
                wordTable.Cell(current_row + 1, 8).Range.Text = ableResult.result2;
                //владеть
                wordTable.Cell(current_row + 2, 5).Range.Text = ownResult.result5;
                wordTable.Cell(current_row + 2, 6).Range.Text = ownResult.result4;
                wordTable.Cell(current_row + 2, 7).Range.Text = ownResult.result3;
                wordTable.Cell(current_row + 2, 8).Range.Text = ownResult.result2;

                //форматирование - объединение строк, столбцов
                wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row + 2, 1));
                wordTable.Cell(current_row, 2).Merge(wordTable.Cell(current_row + 2, 2));
                wordTable.Cell(current_row, 3).Merge(wordTable.Cell(current_row + 2, 3));

                //переход новую компенцию
                current_row = current_row + 3;
            }
            //форматирование
            wordTable.Borders.Enable = Convert.ToInt32(true);
            wordTable.Range.ParagraphFormat.SpaceBefore = 0;
            wordTable.Range.ParagraphFormat.SpaceAfter = 0;
            wordTable.Range.Font.ColorIndex = WdColorIndex.wdAuto;
            wordTable.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
        }

        //6.3.1.2 Содержание оценочного средства
        private void createTable7()
        {
            List<QuestionModel> test_computer = new List<QuestionModel>()
            {
                new QuestionModel("Укажите цель метрологии:", 4,1,
                    new List<QuestionModel>(){
                        new QuestionModel("обеспечение единства измерений с необходимой и требуемой, точностью"),
                        new QuestionModel("разработка и совершенствование средств и методов измерений повышения их точности"),
                        new QuestionModel("разработка новой и совершенствование, действующей правовой и нормативной базы"),
                        new QuestionModel("совершенствование эталонов единиц измерения для повышения их точности"),
                        new QuestionModel("обеспечение единства измерений с необходимой и требуемой, точностью"),
                    }
                ),
                new QuestionModel("Метрология -..", 4, 1,
                    new List<QuestionModel>(){
                        new QuestionModel("ответ1"),
                        new QuestionModel("ответ2"),
                        new QuestionModel("ответ3"),
                        new QuestionModel("ответ4"),
                    }
                ),
                new QuestionModel("Косвенные измерения - это такие измерения, при которых", 4, 1,
                    new List<QuestionModel>(){
                        new QuestionModel("ответ1"),
                        new QuestionModel("ответ2"),
                        new QuestionModel("ответ3"),
                        new QuestionModel("ответ4"),
                        new QuestionModel("ответ5"),
                    }
                ),
                new QuestionModel("Прямые измерения — это такие измерения, при которых:", 4, 1,
                    new List<QuestionModel>(){
                        new QuestionModel("ответ1"),
                        new QuestionModel("ответ2"),
                        new QuestionModel("ответ3"),
                        new QuestionModel("ответ4"),
                    }
                ),
                new QuestionModel("Значение любой ФВ Q, представленное в виде Q=q[Q] называется…", 4, 2,
                    new List<QuestionModel>(){
                        new QuestionModel("ответ1"),
                        new QuestionModel("ответ2"),
                        new QuestionModel("ответ3"),
                        new QuestionModel("ответ4"),
                    }
                ),
                new QuestionModel("Определяющим уравнением ускорения является: a=v/t. Размерность", 4, 2,
                    new List<QuestionModel>(){
                        new QuestionModel("ответ1"),
                        new QuestionModel("ответ2"),
                        new QuestionModel("ответ3"),
                        new QuestionModel("ответ4"),
                        new QuestionModel("ответ5"),
                    }
                ),
            };

            DataTable dt = new DataTable();
            dt.Columns.Add("Код");
            dt.Columns.Add("Вопрос");
            int max_answers = test_computer.Max(a => a.answers.Count);
            for (int i = 1; i <= max_answers; i++)
                dt.Columns.Add(i+"");
            int row_count = 2+test_computer.Count+test_computer.Select(a=>a.semester).Distinct().Count()*2;
            for (int i = 1; i <= row_count; i++)
                dt.Rows.Add();

            app.Selection.Find.Execute("<TABLE7>");
            Word.Range wordRange = app.Selection.Range;
            Word.Table wordTable = wordDocument.Tables.Add(wordRange, dt.Rows.Count, dt.Columns.Count);
            //шапка
            wordTable.Cell(1, 1).Range.Text = "Код компетенции";
            wordTable.Cell(1, 2).Range.Text = "Тестовые вопросы";
            wordTable.Cell(1, 3).Range.Text = "Варианты ответов";
            for (int i = 1; i <= max_answers; i++)
                wordTable.Cell(2, 2 + i).Range.Text = i.ToString();

            //форматирование
            wordTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordTable.Borders.Enable = Convert.ToInt32(true);
            wordTable.Cell(1, 1).Merge(wordTable.Cell(2, 1));
            wordTable.Cell(1, 2).Merge(wordTable.Cell(2, 2));
            wordTable.Cell(1, 3).Merge(wordTable.Cell(1, max_answers + 2));

            int current_row = 3;
            foreach (var semester in discipline.semesters)
            {
                var list1 = test_computer.Where(a => a.semester == semester && a.module == 1).ToList();
                var list2 = test_computer.Where(a => a.semester == semester && a.module == 2).ToList();

                //МОДУЛЬ_1
                wordTable.Cell(current_row, 1).Range.Text = "Дисциплинарный модуль " + semester + ".1";
                wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row, max_answers + 2));
                wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
                current_row++;

                //данные
                wordTable.Cell(current_row, 1).Range.Text = string.Join(", ", discipline.competences.Select(a => a.kod));
                wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
                wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row + list1.Count - 1, 1));
                foreach (var item in list1)
                {
                    wordTable.Cell(current_row, 2).Range.Text = item.name;
                    wordTable.Cell(current_row, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    for (int i = 0; i < item.answers.Count; i++)
                    {
                        wordTable.Cell(current_row, 3 + i).Range.Text = item.answers[i].name;
                        wordTable.Cell(current_row, 3 + i).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    }
                    current_row++;
                }

                //МОДУЛЬ_2
                wordTable.Cell(current_row, 1).Range.Text = "Дисциплинарный модуль " + semester + ".2";
                wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row, max_answers + 2));
                wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
                current_row++;

                //данные
                wordTable.Cell(current_row, 1).Range.Text = string.Join(", ", discipline.competences.Select(a => a.kod));
                wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
                wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row + list2.Count - 1, 1));
                foreach (var item in list2)
                {
                    wordTable.Cell(current_row, 2).Range.Text = item.name;
                    wordTable.Cell(current_row, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    for (int i = 0; i < item.answers.Count; i++)
                    {
                        wordTable.Cell(current_row, 3 + i).Range.Text = item.answers[i].name;
                        wordTable.Cell(current_row, 3 + i).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    }
                    current_row++;
                }

            }
            wordTable.Range.ParagraphFormat.SpaceBefore = 0;
            wordTable.Range.ParagraphFormat.SpaceAfter = 0;
            wordTable.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed);
        }


    }
}

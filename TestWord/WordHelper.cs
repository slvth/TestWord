using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Data;
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
                createTable8();

                createTable9();
                createTable10(); ////
                createTable11();
                createTable12();
                createTable13();
                createTable14();
                createTable15();
                createTable16();

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
                new ThemeModel("Тема 1. Основы метрологии", 4, 1, 2, 4, 4, 4,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("История развития метрологии", 2, "", 0),
                    },
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Измерение линейных размеров с помощью штангенинструментов", 2, "", 1),
                        new ChildThemeModel("Электрические измерения напряжения", 2, "", 1),
                    },
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Системы физических единиц", 2, "", 1),
                        new ChildThemeModel("Размерность физических единиц", 2, "", 1)
                    }
                ),
                new ThemeModel("Тема 2. Средства и методы измерения", 4, 1, 4, 4, 6, 4,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Виды и методы измерений", 2, "Проблемная лекция", 0),
                        new ChildThemeModel("Средства измерений", 2, "Лекция с запланированными ошибками", 0),
                    },
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Поверка СИ температуры", 2, "", 1),
                        new ChildThemeModel("Проверка средств измерения давления", 2, "", 1),
                        new ChildThemeModel("Аттестация средств измерения давления", 2, "", 1),
                    },
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Температурные шкалы", 2, "", 1),
                        new ChildThemeModel("Метрологические характеристики средств измерения", 2, "работа в малых группах", 1),
                    }
                ),
                new ThemeModel("Тема 3. Погрешности измерения", 4, 2, 4, 6, 8, 4,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Основы метрологического обеспечения производства", 2, "", 0),
                        new ChildThemeModel("Понятие о погрешности измерений", 2, "", 0),
                    },
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Определение метрологических характеристик средств измерения", 2, "", 1),
                        new ChildThemeModel("Влияние газового фактора на точность измерений", 2, "", 1),
                        new ChildThemeModel("Определение погрешностей СИ при изменении характеристики среды", 2, "", 1),
                        new ChildThemeModel("Влияние не стабильности потока на точность измерения", 2, "", 1),
                    },
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Определение погрешностей измерения", 2, "групповое обсуждение", 1),
                        new ChildThemeModel("Погрешности косвенных измерений", 2, "", 1),
                        new ChildThemeModel("Определение доверительных границ и доверительных интервалов", 2, "работа в малых группах", 1),
                    }
                ),
                new ThemeModel("Тема 4. Основы стандартизации", 4, 2, 4, 2, 0, 4,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("История развития стандартизации", 2, "Лекция-визуализация", 0),
                        new ChildThemeModel("Методы и средства стандартизации", 2, "", 0),
                    },
                    null,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Нормативно-правовые документы по стандартизации", 2, "", 1),
                    }
                ),
                new ThemeModel("Тема 5. Основы сертификации", 4, 2, 2, 2, 0, 4,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Основные понятия сертификации", 2, "", 0),
                    },
                    null,
                    new List<ChildThemeModel> {
                        new ChildThemeModel("Сходства и отличия «Сертификация соответствия» и «Декларирование соответствия»", 2, "", 1),
                    }
                ),
            };

            List<int> semesters = new List<int>() { 4, 8 };
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

            //баллы
            public float score { get; set; }

            public ChildThemeModel(string name, int hour, string method, float score)
            {
                this.name = name;
                this.hour = hour;
                this.method = method;
                this.score = score;
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

            //вопросы к экзамену 6.3.4.3. 
            public QuestionModel(string name, List<string> competence)
            {
                this.name = name;
                this.competence = competence;
            }


            public QuestionModel(string name)
            {
                this.name = name;
            }

            //тест экзамен
            public QuestionModel(string name, List<QuestionModel> answers)
            {
                this.name = name;
                this.semester = semester;
                this.module = module;
                this.answers = answers;
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

                wordRange.Collapse(WdCollapseDirection.wdCollapseStart);
                wordRange.InsertAfter(semester_text);
                wordRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordRange.Bold = Convert.ToInt32(true);
                wordRange.Font.Size = 14;
                wordRange.InsertParagraphAfter();
                wordRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                wordRange.Collapse(WdCollapseDirection.wdCollapseStart);
                wordRange.InsertAfter(semester_table);
                wordRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordRange.Bold = Convert.ToInt32(false);
                wordRange.Font.Size = 12;
                wordRange.InsertParagraphAfter();
                wordRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            }
            //удаление main_key из word
            app.Selection.Find.Execute(main_key);
            Word.Range wordRangeDelete = app.Selection.Range;
            wordRangeDelete.Delete();

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

                app.Selection.Find.Wrap = WdFindWrap.wdFindContinue;
                bool g = app.Selection.Find.Execute(semester.Key);
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
            //замена тэгов на слова семестров
            replaceText(text_keys);
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

                if(list1.Count is 0 && list2.Count is 0)
                    continue;
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

        //Экзамен 6.3.4.3. Содержание оценочного средства
        private void createTable8()
        {
            //вопросы к экзамену
            List<QuestionModel> questions = new List<QuestionModel>()
            {
                new QuestionModel("Основные, дополнительные, производные, внесистемные единицы, допускаемые к применению наравне с единицами СИ", new List<string>(){"ОПК-11", "ОПК-12"}),
                new QuestionModel("Кратные и дольные единицы измерений", new List<string>(){"ОПК-11", "ОПК-12"}),
                new QuestionModel("Международная система единиц физических величин (СИ), ее применение в России", new List<string>(){"ОПК-11"}),
                new QuestionModel("Эталоны физических величин", new List<string>(){"ОПК-12"}),
            };

            app.Selection.Find.Execute("<TABLE8>");
            var wordRange = app.Selection.Range;

            int colomn_count = 2+discipline.competences.Count;
            int row_count = 1+questions.Count;

            Word.Table wordTable = wordDocument.Tables.Add(wordRange, row_count, colomn_count);
            wordTable.Borders.Enable = Convert.ToInt32(true);

            wordTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            //заполнение
            wordTable.Cell(1, 1).Range.Text = "№ п/п";
            wordTable.Cell(1, 2).Range.Text = "Примерные вопросы к экзамену";
            for (int i = 0; i < discipline.competences.Count; i++)
                wordTable.Cell(1, 3 + i).Range.Text = discipline.competences[i].kod;

            int current_row = 2;
            for (int i = 0; i < questions.Count; i++)
            {
                var question = questions[i];
                wordTable.Cell(current_row, 1).Range.Text = (i + 1).ToString();
                wordTable.Cell(current_row, 2).Range.Text = question.name;
                for (int j = 0; j < discipline.competences.Count; j++)
                {
                    var current_competence = discipline.competences[j];
                    for (int k = 0; k < question.competence.Count; k++)
                    {
                        if (question.competence[k] == current_competence.kod)
                            wordTable.Cell(current_row, 3 + j).Range.Text = "+";
                    }
                }
                wordTable.Cell(current_row, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                current_row++;
            }

            wordTable.Range.ParagraphFormat.SpaceBefore = 0;
            wordTable.Range.ParagraphFormat.SpaceAfter = 0;
            wordTable.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
        }
        
        //образей вариантов тестовых заданий на экзамен
        private void createTable9()
        {
            List<QuestionModel> test = new List<QuestionModel>()
            {
                new QuestionModel(
                    "Значение любой физической величины Q, представленное в виде Q=q[Q] произведения ее числового значения q на единицу измерения [Q], называется…",
                    new List<QuestionModel>()
                    {
                        new QuestionModel("Математической моделью измерений"),
                        new QuestionModel("Функцией преобразования средства измерений"),
                        new QuestionModel("Основным уравнением измерений"),
                        new QuestionModel("Шкалой физической величины"),
                    }
                ),
                new QuestionModel(
                    "Уравнение ускорения a=v/t. Размерность ускорения запишется следующим образом",
                    new List<QuestionModel>()
                    {
                        new QuestionModel("LT^(-2)"),
                        new QuestionModel("MLT^(-1)"),
                        new QuestionModel("LT^(-1)"),
                        new QuestionModel("L^(-2) T"),
                    }
                ),
                new QuestionModel(
                    "Определение искомого значения ФВ на основании результатов прямых измерений других ФВ, функционально связанных с искомой величиной, называется ________ измерением",
                    new List<QuestionModel>()
                    {
                        new QuestionModel("Косвенным"),
                        new QuestionModel("Совместным"),
                        new QuestionModel("Совокупным"),
                        new QuestionModel("Прямым"),
                    }
                ),
            };

            DataTable dt = new DataTable();
            dt.Columns.Add("Код");
            dt.Columns.Add("Вопрос");
            int max_answers = test.Max(a => a.answers.Count);
            for (int i = 1; i <= max_answers; i++)
                dt.Columns.Add(i + "");
            int row_count = 2 + test.Count;
            for (int i = 1; i <= row_count; i++)
                dt.Rows.Add();

            app.Selection.Find.Execute("<TABLE9>");
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

            //данные
            wordTable.Cell(current_row, 1).Range.Text = string.Join(", ", discipline.competences.Select(a => a.kod));
            wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
            wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row + test.Count - 1, 1));
            foreach (var item in test)
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
            
            wordTable.Range.ParagraphFormat.SpaceBefore = 0;
            wordTable.Range.ParagraphFormat.SpaceAfter = 0;
            wordTable.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed);
        }

        //ДОРАБОТАТЬ
        private void createTable10()
        {
            string main_key = "<TABLE10>";
            app.Selection.Find.Execute(main_key);
            Word.Range wordRange = app.Selection.Range;

            Dictionary<int, Dictionary<string, string>> text_keys = new Dictionary<int, Dictionary<string, string>>();
            Dictionary<int, List<string>> table_keys = new Dictionary<int, List<string>>();

            //вставка тэгов в ворд
            foreach (var semester in discipline.semesters)
            {
                List<string> list_key_text = new List<string>()
                {
                    $"<TEXT10_{semester}.0>", 
                    $"<TEXT10_{semester}.1>", 
                    $"<TEXT10_{semester}.2>"
                };
                List<string> list_key_table = new List<string>()
                {
                    $"<TABLE10_{semester}.0>",
                    $"<TABLE10_{semester}.1>",
                    $"<TABLE10_{semester}.2>",
                };

                text_keys.Add(semester, new Dictionary<string, string>()
                {
                    [list_key_text[0]] = $"Семестр {semester}",
                    [list_key_text[1]] = $"Дисциплинарный модуль {semester}.1",
                    [list_key_text[2]] = $"Дисциплинарный модуль {semester}.2",
                });

                table_keys.Add(semester, list_key_table);

                for (int i = 0; i < 3; i++)
                {
                    wordRange.Collapse(WdCollapseDirection.wdCollapseStart);
                    wordRange.InsertAfter(list_key_text[i]);
                    wordRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    wordRange.Bold = Convert.ToInt32(true);
                    wordRange.Font.Size = 14;
                    wordRange.InsertParagraphAfter();
                    wordRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                    wordRange.Collapse(WdCollapseDirection.wdCollapseStart);
                    wordRange.InsertAfter(list_key_table[i]);
                    wordRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    wordRange.Bold = Convert.ToInt32(false);
                    wordRange.Font.Size = 12;
                    wordRange.InsertParagraphAfter();
                    wordRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                }
            }
            //удаление main_key из word
            app.Selection.Find.Wrap = WdFindWrap.wdFindContinue;
            app.Selection.Find.Execute(main_key);
            Word.Range wordRangeDelete = app.Selection.Range;
            wordRangeDelete.Delete();

            foreach (var item in text_keys)
                replaceText(item.Value);

            int current_lecture = 1;
            int current_laboratories = 1;
            int current_practical = 1;
            //замена тэгов на таблицы семестров
            foreach (var semester in table_keys)
            {
                List<string> current_keys = semester.Value;
                
                //ГЛАВНАЯ ТАБЛИЦА
                var dt_main = new DataTable();
                for (int i = 0; i < 5; i++)
                    dt_main.Rows.Add();
                for (int i = 0; i < 3; i++)
                    dt_main.Columns.Add();
                app.Selection.Find.Wrap = WdFindWrap.wdFindContinue;
                app.Selection.Find.Execute(current_keys[0]);
                Word.Range wordRangeTable = app.Selection.Range;
                var wordTable = wordDocument.Tables.Add(wordRangeTable,
                    dt_main.Rows.Count, dt_main.Columns.Count);
                wordTable.Cell(1, 1).Range.Text = "Дисциплинарный модуль";
                //!!!! лабораторные устные опросы ДОРАБОТАТЬ!!!!!!!!!
                wordTable.Cell(2, 1).Range.Text = "Текущий контроль (лабораторные работы, практические задачи)";
                wordTable.Cell(3, 1).Range.Text = "Текущий контроль (тестирование)";
                wordTable.Cell(4, 1).Range.Text = "Общее количество баллов";
                wordTable.Cell(5, 1).Range.Text = "Итоговый балл:";
                wordTable.Cell(1, 2).Range.Text = $"ДМ {semester.Key}.1";
                wordTable.Cell(1, 3).Range.Text = $"ДМ {semester.Key}.2";
                //!!!! Доработать, чтобы баллы брались $$откуда-то$$
                //баллы за практические, лабы, устные
                wordTable.Cell(2, 2).Range.Text = "9-15";
                wordTable.Cell(2, 3).Range.Text = "9-15";
                //баллы за тестирование
                wordTable.Cell(3, 2).Range.Text = "8-15";
                wordTable.Cell(3, 3).Range.Text = "9-15";
                //общее кол-во баллов
                wordTable.Cell(4, 2).Range.Text = "17-30";
                wordTable.Cell(4, 3).Range.Text = "18-30";
                //итоговый балл
                /* 
                string score = "";
                if (экзамен || зачет)
                    score = "35-60";
                else if(зачет с оценкой)
                    score = "35-60";
                */
                wordTable.Cell(5, 2).Range.Text = "35-60";
                //форматирование
                wordTable.Borders.Enable = Convert.ToInt32(true);
                wordTable.Range.ParagraphFormat.SpaceBefore = 0;
                wordTable.Range.ParagraphFormat.SpaceAfter = 0;
                wordTable.Cell(1, 1).Range.Bold = Convert.ToInt32(true);
                wordTable.Cell(1, 2).Range.Bold = Convert.ToInt32(true);
                wordTable.Cell(1, 3).Range.Bold = Convert.ToInt32(true);
                wordTable.Cell(4, 1).Range.Bold = Convert.ToInt32(true);
                wordTable.Cell(4, 2).Range.Bold = Convert.ToInt32(true);
                wordTable.Cell(4, 3).Range.Bold = Convert.ToInt32(true);
                wordTable.Cell(5, 1).Range.Bold = Convert.ToInt32(true);
                wordTable.Cell(5, 2).Range.Bold = Convert.ToInt32(true);
                for (int i = 1; i <= 5; i++)
                    wordTable.Cell(i, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                for (int i = 1; i <= 5; i++)
                    for (int j = 2; j <= 3; j++)
                        wordTable.Cell(i, j).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordTable.Cell(5, 2).Merge(wordTable.Cell(5, 3));
                wordTable.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);

                //ТАБЛИЦЫ - дисциплинарный модуль1,2
                for (int module = 1; module <= 2; module++)
                {
                    List<ChildThemeModel> lecture_list = discipline.themes.Where(a => a.semester.Equals(semester.Key) && a.module.Equals(module) && a.lectures is not null).SelectMany(a => a.lectures.Where(b => b.score is not 0)).ToList();
                    List<ChildThemeModel> laboratory_list = discipline.themes.Where(a => a.semester.Equals(semester.Key) && a.module.Equals(module) && a.laboratories is not null).SelectMany(a => a.laboratories.Where(b => b.score is not 0)).ToList();
                    List<ChildThemeModel> practical_list = discipline.themes.Where(a => a.semester.Equals(semester.Key) && a.module.Equals(module) && a.practicals is not null).SelectMany(a => a.practicals.Where(b => b.score is not 0)).ToList();
                    int row = 2 + lecture_list.Count + laboratory_list.Count + practical_list.Count + 4;
                    int column = 3;
                    app.Selection.Find.Wrap = WdFindWrap.wdFindContinue;
                    app.Selection.Find.Execute(current_keys[module]);
                    wordRangeTable = app.Selection.Range;
                    wordTable = wordDocument.Tables.Add(wordRangeTable, row, column);
                    wordTable.Cell(1, 1).Range.Text = "№п/п";
                    wordTable.Cell(1, 2).Range.Text = "Виды работ";
                    wordTable.Cell(1, 3).Range.Text = "Максимальный балл";
                    wordTable.Cell(2, 1).Range.Text = "Текущий контроль";
                    //форматирование
                    wordTable.Borders.Enable = Convert.ToInt32(true);
                    wordTable.Range.ParagraphFormat.SpaceBefore = 0;
                    wordTable.Range.ParagraphFormat.SpaceAfter = 0;
                    wordTable.Range.Font.Size = 12;
                    wordTable.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                    wordTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    wordTable.Cell(1, 1).Range.Bold = Convert.ToInt32(true);
                    wordTable.Cell(1, 2).Range.Bold = Convert.ToInt32(true);
                    wordTable.Cell(1, 3).Range.Bold = Convert.ToInt32(true);
                    wordTable.Cell(2, 1).Range.Bold = Convert.ToInt32(true);
                    wordTable.Cell(2, 1).Merge(wordTable.Cell(2, 3));
                    int current_row = 3;
                    if (lecture_list is not null)
                        for (int index = 0; index < lecture_list.Count; index++)
                        {
                            wordTable.Cell(current_row, 1).Range.Text = (index + 1).ToString();
                            wordTable.Cell(current_row, 2).Range.Text = $"Лекция-{current_lecture} {lecture_list[index].name}";
                            wordTable.Cell(current_row, 3).Range.Text = lecture_list[index].score.ToString();
                            //форматирование
                            wordTable.Cell(current_row, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            current_lecture++;
                            current_row++;
                        }
                    if (laboratory_list is not null)
                        for (int index = 0; index < laboratory_list.Count; index++)
                        {
                            wordTable.Cell(current_row, 1).Range.Text = (index + 1).ToString();
                            wordTable.Cell(current_row, 2).Range.Text = $"Л.Р.-{current_laboratories} {laboratory_list[index].name}";
                            wordTable.Cell(current_row, 3).Range.Text = laboratory_list[index].score.ToString();
                            //форматирование
                            wordTable.Cell(current_row, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            current_laboratories++;
                            current_row++;
                        }
                    if (practical_list is not null)
                        for (int index = 0; index < practical_list.Count; index++)
                        {
                            wordTable.Cell(current_row, 1).Range.Text = (index + 1).ToString();
                            wordTable.Cell(current_row, 2).Range.Text = $"П.З.-{current_practical} {practical_list[index].name}";
                            wordTable.Cell(current_row, 3).Range.Text = practical_list[index].score.ToString();
                            //форматирование
                            wordTable.Cell(current_row, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            current_practical++;
                            current_row++;
                        }
                    wordTable.Cell(current_row, 1).Range.Text = "Итого:";
                    wordTable.Cell(current_row, 3).Range.Text = "15";
                    //форматирование
                    wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
                    wordTable.Cell(current_row, 3).Range.Bold = Convert.ToInt32(true);
                    wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row, 2));
                    wordTable.Cell(current_row, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    current_row++;
                    wordTable.Cell(current_row, 1).Range.Text = "Текущий контроль";
                    //форматирование
                    wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
                    wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row, 3));
                    current_row++;
                    wordTable.Cell(current_row, 1).Range.Text = "1";
                    wordTable.Cell(current_row, 2).Range.Text = "Тестирование";
                    wordTable.Cell(current_row, 3).Range.Text = "15";
                    //форматирование
                    wordTable.Cell(current_row, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    wordTable.Cell(current_row, 3).Range.Bold = Convert.ToInt32(true);
                    current_row++;
                    wordTable.Cell(current_row, 1).Range.Text = $"Итого по ДМ {semester.Key}.{module}";
                    wordTable.Cell(current_row, 3).Range.Text = "30";
                    //форматирование
                    wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
                    wordTable.Cell(current_row, 3).Range.Bold = Convert.ToInt32(true);
                    wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row, 2));
                    wordTable.Cell(current_row, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    //финальное форматирование
                    wordTable.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);

                }
            }
        }

        private void createTable11()
        {
            //основная литература
            List<LiteratureModel> main = new List<LiteratureModel>()
            {
                new LiteratureModel("Мирный, В. И. Законодательная метрология : учебное пособие / В. И. Мирный", 101, 1),
                new LiteratureModel("Воробьева, Г. Н. Метрология, стандартизация и сертификация", "http://www.iprbookshop.ru/57097.html", 0),
                new LiteratureModel("Семенов, И. В. Метрология, стандартизация и сертификация", "https://www.iprbookshop.ru/115857.html", 1),
            };
            //дополнительная литература
            List<LiteratureModel> additional = new List<LiteratureModel>()
            {
                new LiteratureModel("Лепявко, А. П. Измерительные преобразователи давления", "https://www.iprbookshop.ru/118037.html", 1),
                new LiteratureModel("Мирный, В. И. Прикладная метрология : учебное пособие / В. И. Мирный", "https://www.iprbookshop.ru/118037.html", 1),
                new LiteratureModel("Многофазные испытательные стенды газожидкостных смесей: метрология, моделирование, подобие : монография", "https://www.iprbookshop.ru/118037.html", 1),
            };
            //учебно-методические издания
            List<LiteratureModel> methodical = new List<LiteratureModel>()
            {
                new LiteratureModel("Ситдикова И.П., Ахметзянов Р.Р. Метрология, стандартизация и сертификация: методические указания для выполнения лабораторных работ", "http://elibrary.agni-rt.ru", 1),
                new LiteratureModel("Ситдикова И.П., Ахметзянов Р.Р. Метрология, стандартизация и сертификация: методические указания по выполнению курсовой работы", "http://elibrary.agni-rt.ru", 1),
                new LiteratureModel("Ситдикова И.П., Ахметзянов Р.Р. Метрология, стандартизация и сертификация: методические указания для проведения практических занятий", "http://elibrary.agni-rt.ru", 1),
            };
            int row = 2 + main.Count+1+additional.Count+1+methodical.Count;
            int column = 4;
            app.Selection.Find.Execute("<TABLE11>");
            Word.Range wordTableRange = app.Selection.Range;
            Word.Table wordTable = app.ActiveDocument.Tables.Add(wordTableRange, row, column);
            wordTable.Cell(1, 1).Range.Text = "№ п/п";
            wordTable.Cell(1, 2).Range.Text = "Библиографическое описание";
            wordTable.Cell(1, 3).Range.Text = "Количество печатных экземпляров или адрес электронного ресурса";
            wordTable.Cell(1, 4).Range.Text = "Коэффициент обеспеченности";
            //форматирование
            wordTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordTable.Cell(1, 4).Range.Orientation = WdTextOrientation.wdTextOrientationUpward;
            wordTable.Cell(1, 1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable.Cell(1, 2).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable.Cell(1, 3).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable.Cell(1, 4).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable.Rows[1].Height = app.CentimetersToPoints(4f);
            wordTable.Range.Columns[1].Width = app.CentimetersToPoints(0.99f);
            wordTable.Range.Columns[2].Width = app.CentimetersToPoints(7.84f);
            wordTable.Range.Columns[3].Width = app.CentimetersToPoints(7.30f);
            wordTable.Range.Columns[4].Width = app.CentimetersToPoints(1.22f);
            int current_row = 2;
            //основная литература
            wordTable.Cell(current_row, 1).Range.Text = "Основная литература";
            wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
            wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row, 4));
            current_row++;
            for (int i = 0; i < main.Count; i++)
            {
                wordTable.Cell(current_row, 1).Range.Text = $"{i + 1}";
                wordTable.Cell(current_row, 2).Range.Text = $"{main[i].name}";
                wordTable.Cell(current_row, 3).Range.Text = main[i].link is not null ? $"Режим доступа:\n{main[i].link}" : $"{main[i].count} экз.";
                wordTable.Cell(current_row, 4).Range.Text = main[i].coefficient is 0 ? "" : $"{main[i].coefficient}";
                //форматирование
                wordTable.Cell(current_row, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordTable.Cell(current_row, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                current_row++;
            }
            //дополнительная литература
            wordTable.Cell(current_row, 1).Range.Text = "Дополнительная литература";
            wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
            wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row, 4));
            current_row++;
            for (int i = 0; i < additional.Count; i++)
            {
                wordTable.Cell(current_row, 1).Range.Text = $"{i + 1}";
                wordTable.Cell(current_row, 2).Range.Text = $"{additional[i].name}";
                wordTable.Cell(current_row, 3).Range.Text = additional[i].link is not null ? $"Режим доступа:\n{additional[i].link}" : $"{additional[i].count} экз.";
                wordTable.Cell(current_row, 4).Range.Text = additional[i].coefficient is 0 ? "" : $"{additional[i].coefficient}";
                //форматирование
                wordTable.Cell(current_row, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordTable.Cell(current_row, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                current_row++;
            }
            //Учебно-методические издания
            wordTable.Cell(current_row, 1).Range.Text = "Учебно-методические издания";
            wordTable.Cell(current_row, 1).Range.Bold = Convert.ToInt32(true);
            wordTable.Cell(current_row, 1).Merge(wordTable.Cell(current_row, 4));
            current_row++;
            for (int i = 0; i < methodical.Count; i++)
            {
                wordTable.Cell(current_row, 1).Range.Text = $"{i + 1}";
                wordTable.Cell(current_row, 2).Range.Text = $"{methodical[i].name}";
                wordTable.Cell(current_row, 3).Range.Text = methodical[i].link is not null ? $"Режим доступа:\n{methodical[i].link}" : $"{methodical[i].count} экз.";
                wordTable.Cell(current_row, 4).Range.Text = methodical[i].coefficient is 0 ? "" : $"{methodical[i].coefficient}";
                //форматирование
                wordTable.Cell(current_row, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordTable.Cell(current_row, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                current_row++;
            }
            
            //wordTable.Range.Rows[1].Height = app.CentimetersToPoints(3.22f);
            //wordTable.Range.Columns[4].Width = app.CentimetersToPoints(1.28f);

            
            wordTable.Borders.Enable = Convert.ToInt32(true);
            wordTable.Range.ParagraphFormat.SpaceBefore = 0;
            wordTable.Range.ParagraphFormat.SpaceAfter = 0;
            wordTable.Range.Font.Size = 12;
            wordTable.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
        }

        private class LiteratureModel
        {
            //6 и 8
            public string name { get; set; }
            public int? count { get; set; }
            public string? link { get; set; }
            public float? coefficient { get; set; }
            //6
            public LiteratureModel(string name, int count, float coefficient)
            {
                this.name = name;
                this.count = count;
                this.coefficient = coefficient;
            }
            public LiteratureModel(string name, string link, float coefficient)
            {
                this.name = name;
                this.link = link;
                this.coefficient = coefficient;
            }
            //8
            public LiteratureModel(string name, string link)
            {
                this.name = name;
                this.link = link;
            }
        }

        private void createTable12()
        {
            List<LiteratureModel> site_list = new List<LiteratureModel>()
            {
                new LiteratureModel("Учебно-методическая литература для учащихся и студентов, размещенная на сайте «Studmed.ru»", "http://www.studmed.ru "),
                new LiteratureModel("Единое окно доступа к информационным ресурсам", "http://window.edu.ru/ "),
                new LiteratureModel("Российская государственная библиотека", "http://www.rsl.ru "),
                new LiteratureModel("Электронная библиотека Elibrary", "http://elibrary.ru "),
                new LiteratureModel("Электронно-библиотечная система IPRbooks", "http://elibrary.ru "),
                new LiteratureModel("Электронная библиотека АГНИ", "http://elibrary.agni-rt.ru "),
                new LiteratureModel("Энциклопедия России «Библиотекарь»", "http://bibliotekar.ru "),
            };
            int row = 1 + site_list.Count;
            int column = 3;
            app.Selection.Find.Execute("<TABLE12>");
            Word.Range wordTableRange = app.Selection.Range;
            Word.Table wordTable = wordDocument.Tables.Add(wordTableRange, row, column);
            //форматирование
            wordTable.Borders.Enable = Convert.ToInt32(true);
            wordTable.Range.ParagraphFormat.SpaceBefore = 0;
            wordTable.Range.ParagraphFormat.SpaceAfter = 0;
            wordTable.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            wordTable.Rows[1].Range.Bold = Convert.ToInt32(true);
            wordTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordTable.Cell(1, 1).Range.Text = "№ п/п";
            wordTable.Cell(1, 2).Range.Text = "Наименование";
            wordTable.Cell(1, 3).Range.Text = "Алрес в Интернете";
            int current_row = 2;
            for (int i = 0; i < site_list.Count; i++)
            {
                wordTable.Cell(current_row, 1).Range.Text = $"{i + 1}";
                wordTable.Cell(current_row, 2).Range.Text = $"{site_list[i].name}";
                wordTable.Cell(current_row, 3).Range.Text = $"{site_list[i].link}";
                //форматирование
                wordTable.Cell(current_row, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordTable.Cell(current_row, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                current_row++;
            }
            wordTable.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
        }

        private void createTable13() //10
        {
            List<SoftwareModel> software_list = new List<SoftwareModel>()
            {
                new SoftwareModel("Microsoft Office Professional Plus 2016 Rus Academic OLP (Word, Excel, PowerPoint, Access)", "№67892163 от 26.12.2016г.", "№0297/136 от 23.12.2016г."),
                new SoftwareModel("Microsoft Office Standard 2016 Rus Academic OLP (Word, Excel, PowerPoint)", "№67892163 от 26.12.2016г.", "№0297/136 от 23.12.2016г."),
                new SoftwareModel("Microsoft Windows Professional 10 Rus Upgrade Academic OLP", "№67892163 от 26.12.2016г.", "№0297/136 от 23.12.2016г."),
                new SoftwareModel("ABBYY Fine Reader 12 Professional", "№197059 от 26.12.2016г.", "№0297/136 от 23.12.2016г."),
                new SoftwareModel("Kaspersky Endpoint Security для бизнеса – Стандартный Russian Edition", "№ 24С4-221222-121357-913-1225", "№691447/581-2022 от 16.12.2022г."), //Доработать текущий год поставить
                new SoftwareModel("Электронно-библиотечная система IPRbooks", "", "Лицензионный договор №409-2022 от 03.11.2022г."),
                new SoftwareModel("Образовательная платформа для подготовки кадров в цифровой экономике DATALIB.RU", "", "Лицензионный договор №428-2022/22d/B от 09.11.2022г."),
                new SoftwareModel("ПО «Автоматизированная тестирующая система", "Свидетельство государственной регистрации программ для ЭВМ №2014614238 от 01.04.2014г.", ""),
            };
            int row = 1 + software_list.Count;
            int column = 4;
            app.Selection.Find.Execute("<TABLE13>");
            Word.Range wordTableRange = app.Selection.Range;
            Word.Table wordTable = wordDocument.Tables.Add(wordTableRange, row, column);
            //форматирование
            wordTable.Borders.Enable = Convert.ToInt32(true);
            wordTable.Range.ParagraphFormat.SpaceBefore = 0;
            wordTable.Range.ParagraphFormat.SpaceAfter = 0;
            wordTable.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            wordTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordTable.Rows[1].Range.Bold = Convert.ToInt32(true);
            wordTable.Range.Font.Size = 12;
            wordTable.Cell(1, 1).Range.Text = "№ п/п";
            wordTable.Cell(1, 2).Range.Text = "Наименование программного обеспечения";
            wordTable.Cell(1, 3).Range.Text = "Лицензия";
            wordTable.Cell(1, 4).Range.Text = "Договор";
            int current_row = 2;
            for (int i = 0; i < software_list.Count; i++)
            {
                wordTable.Cell(current_row, 1).Range.Text = $"{i + 1}";
                wordTable.Cell(current_row, 2).Range.Text = $"{software_list[i].name}";
                wordTable.Cell(current_row, 3).Range.Text = $"{software_list[i].license}";
                wordTable.Cell(current_row, 4).Range.Text = $"{software_list[i].contract}";
                //форматирование
                wordTable.Cell(current_row, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordTable.Cell(current_row, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordTable.Cell(current_row, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                current_row++;
            }
            wordTable.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
        }

        private class SoftwareModel //10
        {
            public string name { get; set; }
            public string license { get; set; }
            public string contract { get; set; }

            public SoftwareModel(string name, string license, string contract)
            {
                this.name = name;
                this.license = license;
                this.contract = contract;
            }
        }

        private void createTable14()
        {


        }

        private void createTable15()
        {


        }

        private void createTable16()
        {


        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Linq;

namespace TestWord
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            var helper = new WordHelper("shablon.docx");

            var items = new Dictionary<string, string>
            {
                //EXCEL ИЛИ ПРОГРАММНО РАССЧИТАТЬ
                //0-1
                {"<YEAR>", "2023" },
                {"<INDEX>", "Б1.0.23" },
                {"<DISCIPLINE>", "Метрология, стандартизация и сертификация" }, //6, 6.3.1.1 (ЛАБЫ), 6.4, 11, аннотация
                {"<DIRECTION>", "15.03.04 - Автоматизация технологических процессов и производств" }, //2, 6.4, 12, аннотация
                {"<PROFILE>", "Автоматизация технологических процессов и производств" }, //2, 12, аннотация
                {"<QUALIFICATION>", "бакалавр" },
                {"<FORM_STUDY>", "очная" },
                {"<LANGUAGE_STUDY>", "русский" },
                {"<YEAR_START>", "2023" },
                //2
                {"<BLOCK_1>", "Блока 1 \"Дисциплины (модули)\""},
                {"<BLOCK_2>", "обязательной части"},
                {"<COURSE_SEMESTER>", " 2 курсе в 4 семестре"},
                //3
                {"<TOILSOMENESS>", "4 зачетных единиц, 144 часов"},
                {"<WORK>", "Контактная работа обучающихся с преподавателем - 58 часов:\r\n- лекции 16 ч.;\r\n- практические занятия 18 ч.;\r\n- лабораторные работы 18 ч.\r\nСамостоятельная работа 20ч.\r\nКонтроль (экзамен) 36 ч."},
                {"<ATTESTATION>", "экзамен в 4 семестре"}, //6.4 //зачет с оценкой в 1, 2, 3 семестрах, экзамен в 4 семестре
                //6
                {"<ATTESTATION_2>", "экзамена"}, //зачета с оценкой (1, 2, 3 семестры) и экзамена (4 семестр)
                //

                //ВВОДИМЫЕ ДАННЫЕ
                {"<AUTHOR>", "И.П.Ситдикова" },
                {"<REVIEWER>", "К.Л.Горшкова" },
                {"<DEPARTMENT_CHAIR>", "И.П.Ситдикова" },
                //5
                //{"<METHOD_BOOK>", "Ситдикова И.П., Ахметзянов Р.Р. Метрология, стандартизация и сертификация: методические указания для выполнения лабораторных работ и организации самостоятельной работы по дисциплине «Метрология, стандартизация и сертификация» для бакалавров направления подготовки 15.03.04 «Автоматизация технологических процессов и производств» очной формы обучения. – Альметьевск: АГНИ, 2021г." }

            };

            helper.Process(items);
        }

        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            // Выбрать путь и имя файла в диалоговом окне
            OpenFileDialog ofd = new OpenFileDialog();
            // Задаем расширение имени файла по умолчанию (открывается папка с программой)
            ofd.DefaultExt = "*.xls;*.xlsx";
            // Задаем строку фильтра имен файлов, которая определяет варианты
            ofd.Filter = "(*.xlsx)|*.xls";
            // Задаем заголовок диалогового окна
            ofd.Title = "Выберите файл EXCEL";
            if (!(ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)) // если файл не выбран -> Выход
                return;

            var helper = new ExcelHelper(ofd.FileName);
            helper.Process();
        }
    }
}

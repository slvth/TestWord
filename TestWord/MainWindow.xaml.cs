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
                {"<YEAR>", "2023" },
                {"<INDEX>", "Б1.0.23" },
                {"<DISCIPLINE>", "Метрология, стандартизация и сертификация" },
                {"<DIRECTION>", "15.03.04 - Автоматизация технологических процессов и производств" },
                {"<PROFILE>", "Автоматизация технологических процессов и производств" },
                {"<QUALIFICATION>", "бакалавр" },
                {"<FORM_STUDY>", "очная" },
                {"<LANGUAGE_STUDY>", "русский" },
                {"<YEAR_START>", "2023" },
                
                //ВВОДИМЫЕ ДАННЫЕ
                {"<AUTHOR>", "И.П.Ситдикова" },
                {"<REVIEWER>", "К.Л.Горшкова" },
                {"<DEPARTMENT_CHAIR>", "И.П.Ситдикова" },
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

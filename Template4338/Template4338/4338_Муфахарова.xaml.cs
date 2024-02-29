using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;


namespace Template4338
{
    /// <summary>
    /// Логика взаимодействия для _4338_Муфахарова.xaml
    /// </summary>
    public partial class _4338_Муфахарова : Window
    {
        public _4338_Муфахарова()
        {
            InitializeComponent();
        }
        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = lastCell.Column;
            int _rows = lastCell.Row;
            string[,] list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 1; i < _rows; i++) // Начинаем считывание с первой строки (вторая строка в Excel)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text; // Обратите внимание на индексы, они сдвинуты на 1
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (usersEntities2 usersEntities = new usersEntities2())
            {
                for (int i = 1; i < _rows; i++) // Начинаем считывание с первой строки (вторая строка в Excel)
                {
                    usersEntities.Employees.Add(new Employee()
                    {
                        Role = list[i, 0], // В списке (list) первый столбец содержит роль
                        FullName = list[i, 1], // В списке второй столбец содержит полное имя
                        Login = list[i, 2], // В списке третий столбец содержит логин
                        Password = list[i, 3] // В списке четвертый столбец содержит пароль
                    });
                }
                usersEntities.SaveChanges();
            }
        }
        private void BnImport1_Click(object sender, RoutedEventArgs e)
        {

            List<Employee> allEmployees;
            using (usersEntities2 usersEntities = new usersEntities2())
            {
                allEmployees = usersEntities.Employees.ToList();
            }

            // Создаем новую книгу Excel
            var app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add();

            // Создаем словарь для группировки по ролям
            Dictionary<string, List<Employee>> groupedEmployees = new Dictionary<string, List<Employee>>();

            // Группируем данные по ролям
            foreach (var employee in allEmployees)
            {
                if (!groupedEmployees.ContainsKey(employee.Role))
                {
                    groupedEmployees[employee.Role] = new List<Employee>();
                }
                groupedEmployees[employee.Role].Add(employee);
            }

            // Проверяем количество категорий и выводим информацию о них
            Console.WriteLine($"Найдено {groupedEmployees.Count} категории(й):");
            foreach (var category in groupedEmployees.Keys)
            {
                Console.WriteLine(category);
            }

            // Создаем листы для каждой категории и заполняем данными
            foreach (var kvp in groupedEmployees)
            {
                Excel.Worksheet roleSheet = workbook.Worksheets.Add();
                roleSheet.Name = kvp.Key;

                // Добавляем заголовки
                roleSheet.Cells[1, 1] = "Login";
                roleSheet.Cells[1, 2] = "Password";

                // Заполняем данными
                int row = 2;
                foreach (var employee in kvp.Value)
                {
                    roleSheet.Cells[row, 1] = employee.Login;
                    roleSheet.Cells[row, 2] = ComputeHash(employee.Password); // Хэшируем пароль
                    row++;
                }
            }

            // Отображаем приложение Excel
            app.Visible = true;
        }

        // Метод для хэширования пароля (можно заменить на ваш алгоритм хэширования)
        private string ComputeHash(string input)
        {
            using (SHA256 sha256Hash = SHA256.Create())
            {
                byte[] bytes = sha256Hash.ComputeHash(Encoding.UTF8.GetBytes(input));
                StringBuilder builder = new StringBuilder();
                foreach (byte b in bytes)
                {
                    builder.Append(b.ToString("x2"));
                }
                return builder.ToString();
            }
        }
    }
}





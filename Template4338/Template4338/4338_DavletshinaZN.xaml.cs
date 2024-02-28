using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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
using Excel = Microsoft.Office.Interop.Excel;
namespace Template4338
{
    /// <summary>
    /// Логика взаимодействия для _4338_DavletshinaZN.xaml
    /// </summary>
    public partial class _4338_DavletshinaZN : Window
    {
        Encoding encoding = Encoding.UTF8;

        public _4338_DavletshinaZN()
        {
            InitializeComponent();
        }

        public void BnExport_Click(object sender, EventArgs e)
        {
            List<Dant> allOrders;
            using (labEntities3 usersEntities = new labEntities3())
            {
                allOrders = usersEntities.Dant
                                    .OrderBy(o => o.CreationDate)
                                    .ToList();
            }

           
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allOrders.Select(o => o.CreationDate).Distinct().Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            // группировка заказов по дате создания
            var ordersByDate = allOrders.GroupBy(o => o.CreationDate);

            // добавление данных на каждый лист
            int sheetIndex = 1;
            foreach (var orderGroup in ordersByDate)
            {
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[sheetIndex];
                worksheet.Name = orderGroup.Key; // Название листа - дата создания заказа

                // Добавление названий колонок
                worksheet.Cells[1, 1] = "Id";
                worksheet.Cells[1, 2] = "Код заказа";
                worksheet.Cells[1, 3] = "Код клиента";
                worksheet.Cells[1, 4] = "Услуги";

                // Добавление данных
                int row = 2;
                foreach (var order in orderGroup)
                {
                    string orderCode = order.ClientCode.ToString() + "/" + order.CreationDate;

                    worksheet.Cells[row, 1] = order.ID;
                    worksheet.Cells[row, 2] = orderCode;
                    worksheet.Cells[row, 3] = order.ClientCode;
                    worksheet.Cells[row, 4] = order.Servicee;

                    row++;
                }

                Excel.Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[row - 1, 4]];
                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                range.Columns.AutoFit();

                sheetIndex++;
            }

           
            app.Visible = true;
        }







        public void BnImport_Click(object sender, EventArgs e)
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
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (labEntities3 usersEntities = new labEntities3())
            {
                
                for (int i = 0; i < _rows; i++)
                {
                    Dant user = new Dant();
                    //  обрабатываем возможные исключения
                    try
                    {
                        user.CreationDate = list[i, 1];
                        user.OrderTime = list[i, 2];
                        user.ClientCode = int.Parse(list[i, 3]);
                        user.Servicee = list[i, 4];
                        string statuss = encoding.GetString(Encoding.Default.GetBytes(list[i, 5]));
                        user.ClosingDate = list[i, 6];
                        user.RentalTime = list[i, 7];

                        // Добавление в контекст сущностей
                        usersEntities.Dant.Add(user);
                    }
                    catch (FormatException ex)
                    {
                        //  ошибка при преобразовании данных
                        Console.WriteLine($"Ошибка формата данных в строке {i + 1}: {ex.Message}");
                    }
                 
                    usersEntities.SaveChanges();
                }
            }
        }
    }
}

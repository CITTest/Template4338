using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Microsoft.Win32;
using System.ComponentModel;

namespace Template4338
{
    public partial class Window1 : Window
    {
        public Window1()
        {
            InitializeComponent();
        }



        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                List<Model> importedData = LoadDataFromExcel(filePath);
                SaveDataToDatabase(importedData);
                Window2 dataWindow = new Window2();
                dataWindow.Show();
            }
        }



        private List<Model> LoadDataFromExcel(string filePath)
        {
            List<Model> data = new List<Model>();

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; 

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    Model rowData = new Model();

                    try
                    {
                        rowData.ФИО = worksheet.Cells[row, 1].Text;
                        rowData.Код_клиента = Convert.ToInt64(worksheet.Cells[row, 2].Text);
                        rowData.Дата_рождения = Convert.ToDateTime(worksheet.Cells[row, 3].Text);
                        rowData.Индекс = Convert.ToInt32(worksheet.Cells[row, 4].Text);
                        rowData.Город = worksheet.Cells[row, 5].Text;
                        rowData.Улица = worksheet.Cells[row, 6].Text;
                        rowData.Дом = Convert.ToInt32(worksheet.Cells[row, 7].Text);
                        rowData.Квартира = Convert.ToInt32(worksheet.Cells[row, 8].Text);
                        rowData.Mail = worksheet.Cells[row, 9].Text;

                        data.Add(rowData);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка при обработке строки {row}: {ex.Message}");
                    }
                }
            }

            return data;
        }


        private void SaveDataToDatabase(List<Model> data)
        {
            using (var context = new DBcontext())
            {
                context.EnsureDatabaseCreated();

                foreach (var item in data)
                {
                    context.Users.Add(item);
                }

                context.SaveChanges();

                MessageBox.Show("Данные успешно импортированы в базу данных.");
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<string, List<Model>> groupedData = GroupData();

            SaveGroupedDataToExcel(groupedData);
        }

        private Dictionary<string, List<Model>> GroupData()
        {
            Dictionary<string, List<Model>> groupedData = new Dictionary<string, List<Model>>();


            using (var context = new DBcontext())
            {
                var distinctStreets = context.Users.Select(u => u.Улица).Distinct().ToList();

                foreach (var street in distinctStreets)
                {
                    var usersOnStreet = context.Users.Where(u => u.Улица == street).ToList();
                    groupedData.Add(street, usersOnStreet);
                }
            }

            return groupedData;
        }


        private void SaveGroupedDataToExcel(Dictionary<string, List<Model>> groupedData)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                foreach (var group in groupedData)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(group.Key);

                    var properties = typeof(Model).GetProperties();

                    for (int col = 1; col <= properties.Length; col++)
                    {
                        worksheet.Cells[1, col].Value = properties[col - 1].Name;
                    }

                    for (int row = 2; row <= group.Value.Count + 1; row++)
                    {
                        var item = group.Value[row - 2];

                        for (int col = 1; col <= properties.Length; col++)
                        {
                            worksheet.Cells[row, col].Value = properties[col - 1].GetValue(item);
                        }
                    }
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Files|*.xlsx";
                if (saveFileDialog.ShowDialog() == true)
                {
                    package.SaveAs(new FileInfo(saveFileDialog.FileName));
                    MessageBox.Show("Данные успешно экспортированы в Excel.");
                }
            }
        }

        private void DataToDatabaseButton_Click(object sender, RoutedEventArgs e)
        {
            Window2 dataWindow = new Window2();
            dataWindow.Show();
            MessageBox.Show("Вы перешли к просмотру данных");
        }
    }
}

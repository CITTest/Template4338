using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Word;
using Microsoft.EntityFrameworkCore;

namespace Template4338
{
    public partial class Window1 : System.Windows.Window
    {
        public Window1()
        {
            InitializeComponent();
        }

        private void ImportJSONButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "JSON Files|*.json";

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                List<Model> importedData = LoadDataFromJSON(filePath);
                SaveDataToDatabase(importedData);
                MessageBox.Show("Данные успешно импортированы из JSON в базу данных.");
            }
        }

        private void ExportToWordButton_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<string, List<Model>> groupedData = GroupData();
            SaveGroupedDataToWord(groupedData);
        }

        private void SaveGroupedDataToWord(Dictionary<string, List<Model>> groupedData)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Add();

            foreach (var group in groupedData)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph = doc.Paragraphs.Add();
                paragraph.Range.Text = group.Key;

                Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(paragraph.Range, group.Value.Count + 1, 10);
                table.Cell(1, 2).Range.Text = "FullName";
                table.Cell(1, 3).Range.Text = "CodeClient";
                table.Cell(1, 4).Range.Text = "BirthDate";
                table.Cell(1, 5).Range.Text = "Index";
                table.Cell(1, 6).Range.Text = "City";
                table.Cell(1, 7).Range.Text = "Street";
                table.Cell(1, 8).Range.Text = "Home";
                table.Cell(1, 9).Range.Text = "Kvartira";
                table.Cell(1, 10).Range.Text = "E_mail";

                for (int i = 0; i < group.Value.Count; i++)
                {
                    var item = group.Value[i];
                    table.Cell(i + 2, 2).Range.Text = item.FullName;
                    table.Cell(i + 2, 3).Range.Text = item.CodeClient;
                    table.Cell(i + 2, 4).Range.Text = item.BirthDate;
                    table.Cell(i + 2, 5).Range.Text = item.Index;
                    table.Cell(i + 2, 6).Range.Text = item.City;
                    table.Cell(i + 2, 7).Range.Text = item.Street;
                    table.Cell(i + 2, 8).Range.Text = item.Home.ToString();
                    table.Cell(i + 2, 9).Range.Text = item.Kvartira.ToString();
                    table.Cell(i + 2, 10).Range.Text = item.E_mail;
                }

                doc.Paragraphs.Add();

                paragraph.Range.InsertBreak(WdBreakType.wdPageBreak);

            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word Files|*.docx";
            if (saveFileDialog.ShowDialog() == true)
            {
                doc.SaveAs2(saveFileDialog.FileName);
                doc.Close();
                wordApp.Quit();
                MessageBox.Show("Данные успешно экспортированы в Word.");
            }
        }

        private List<Model> LoadDataFromJSON(string filePath)
        {
            try
            {
                string jsonContent = File.ReadAllText(filePath);
                List<Model> data = JsonConvert.DeserializeObject<List<Model>>(jsonContent);
                return data;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных из JSON: {ex.Message}");
                return new List<Model>();
            }
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

                try
                {
                    context.SaveChanges();
                }
                catch (DbUpdateException ex)
                {
                    Exception innerException = ex.InnerException;
                    while (innerException != null)
                    {
                        Console.WriteLine(innerException.Message);
                        innerException = innerException.InnerException;
                    }

                }
                MessageBox.Show("Данные успешно импортированы в базу данных.");
            }
        }

        private Dictionary<string, List<Model>> GroupData()
        {
            using (var context = new DBcontext())
            {
                Dictionary<string, List<Model>> groupedData = new Dictionary<string, List<Model>>();

                var sortedData = context.Users.OrderBy(u => u.Street).ToList();

                var distinctStreets = sortedData.Select(u => u.Street).Distinct().ToList();

                foreach (var street in distinctStreets)
                {
                    var usersOnStreet = sortedData.Where(u => u.Street == street).ToList();
                    groupedData.Add(street, usersOnStreet);
                }

                return groupedData;
            }
        }

    }
}

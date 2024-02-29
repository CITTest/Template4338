using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Template4338
{
    /// <summary>
    /// Логика взаимодействия для Window2.xaml
    /// </summary>
    public partial class Window2 : Window
    {
            public Window2()
            {
                InitializeComponent();
                LoadData();
                dataGrid.ItemsSource = LoadDataFromDatabase();

            }

        private void LoadData()
            {
                using (var context = new DBcontext())
                {
                    dataGrid.ItemsSource = context.Users.ToList();
                }
            }

        private void AddAndSaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Model newData = new Model
                {
                    ФИО = fioTextBox.Text,
                    Код_клиента = Convert.ToInt64(codeTextBox.Text),
                    Дата_рождения = birthdateDatePicker.SelectedDate.GetValueOrDefault(),
                    Индекс = Convert.ToInt32(indexTextBox.Text),
                    Город = cityTextBox.Text,
                    Улица = streetTextBox.Text,
                    Дом = Convert.ToInt32(houseTextBox.Text),
                    Квартира = Convert.ToInt32(apartmentTextBox.Text),
                    Mail = emailTextBox.Text
                };

                SaveDataToDatabase(newData);

                dataGrid.ItemsSource = LoadDataFromDatabase();

                MessageBox.Show("Данные успешно добавлены и сохранены в базу данных.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void SaveDataToDatabase(Model newData)
        {
            using (var context = new DBcontext())
            {
                context.EnsureDatabaseCreated();
                context.Users.Add(newData);
                context.SaveChanges();
            }
        }

        private List<Model> LoadDataFromDatabase()
        {
            using (var context = new DBcontext())
            {
                context.EnsureDatabaseCreated();
                return context.Users.ToList();
            }
        }
    }
}

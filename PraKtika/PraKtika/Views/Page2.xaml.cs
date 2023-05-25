using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace PraKtika.Views
{
    /// <summary>
    /// Логика взаимодействия для Page2.xaml
    /// </summary>
    public partial class Page2 : Page
    {
        public Page2()
        {
            InitializeComponent();
            
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Image files|*.bmp;*.jpg;*.png";
            openFileDialog.FilterIndex = 1;
            using (var helper = new ExcelHelper())
            {
                try
                {
                    if (helper.Open(filePath: System.IO.Path.Combine(Environment.CurrentDirectory, "exhibits.xlsx")))
                    {
                        int a = helper.Heigthq() + 1;
                        if (Name.Text == "" || Description.Text == "" || Age.Text == "" || Price.Text == "" || Location.Text == "") 
                        {
                            if (Name.Text == ""){Name.BorderBrush = new SolidColorBrush(Color.FromRgb(238, 32, 77));}
                            if (Description.Text == ""){Description.BorderBrush = new SolidColorBrush(Color.FromRgb(238, 32, 77));}
                            if (Age.Text == ""){Age.BorderBrush = new SolidColorBrush(Color.FromRgb(238, 32, 77));}
                            if (Price.Text == ""){Price.BorderBrush = new SolidColorBrush(Color.FromRgb(238, 32, 77));}
                            if (Location.Text == ""){Location.BorderBrush = new SolidColorBrush(Color.FromRgb(238, 32, 77));}
                            MessageBox.Show("вы не заполнили все ячейки");
                        }
                        else 
                        {
                            if (openFileDialog.ShowDialog()==true) 
                            {
                                Name.BorderBrush = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                                Description.BorderBrush = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                                Age.BorderBrush = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                                Price.BorderBrush = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                                Location.BorderBrush = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                                helper.Set("A", a, Name.Text);
                                helper.Set("B", a, Description.Text);
                                helper.Set("C", a, Age.Text);
                                helper.Set("D", a, Price.Text);
                                helper.Set("E", a, Location.Text);
                                
                               
                                FileInfo fileInf = new FileInfo(openFileDialog.FileName);
                                if (fileInf.Exists)
                                {
                                     fileInf.CopyTo(System.IO.Path.Combine(Environment.CurrentDirectory, "Image", fileInf.Name), true);
                                }
                                helper.Set("F", a, System.IO.Path.Combine(Environment.CurrentDirectory, "Image",  fileInf.Name));
                                helper.Save();
                                helper.Dispose();
                            
                           

                                MessageBox.Show("Вы успешно добавили экспонат");

                                NavigationService.Navigate(new Page1());
                               
                            }
                            
                        }
                        
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
        }
    }
}

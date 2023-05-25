using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.ConstrainedExecution;
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
    /// Логика взаимодействия для Page3.xaml
    /// </summary>
    public partial class Page3 : Page
    {



        
        public Page3()
        {
            string[,] list = new string[632,632];
            InitializeComponent();
            using (var helper = new ExcelHelper())
            {
                try
                {
                    if (helper.Open(filePath: System.IO.Path.Combine(Environment.CurrentDirectory, "exhibits.xlsx")))
                    {
                        int index = (int)App.Current.Resources["Index"];
                        helper.CreateInfo(list,index,listboxInfo,helper);
                        int height = helper.Heigthq();
                        int length = helper.Length();
                        list = helper.Export(list, length, height);
                        string filename = list[index, 5];
                        BitmapImage img = new BitmapImage();
                        img.BeginInit();
                        img.UriSource = new Uri(filename, UriKind.RelativeOrAbsolute);
                        img.EndInit();
                        image.Source = img;
                        helper.Dispose();
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }

        }

        private void listboxInfo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int index2 = listboxInfo.SelectedIndex;

            App.Current.Resources["Index2"] = index2;
            NavigationService.Navigate(new Page4());
        }

        private void ButtonNavigate_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new Page1());
        }
    }
}

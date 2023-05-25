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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PraKtika.Views
{
    /// <summary>
    /// Логика взаимодействия для Page5.xaml
    /// </summary>
    public partial class Page5 : Page
    {

        public Page5()
        {
            InitializeComponent();
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            using (var helper = new ExcelHelper())
            {
                try
                {
                    if (helper.Open(filePath: System.IO.Path.Combine(Environment.CurrentDirectory, "exhibits.xlsx")))
                    {
                        int index = (int)App.Current.Resources["Index"]; 
                        helper.DeleteInfo(helper,index);
                        helper.Save();
                        helper.Dispose();
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
            
            NavigationService.Navigate(new Page1());
        }
        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new Page3());
        }
    }
}

using PraKtika.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Navigation;
using System.Windows.Shapes;



namespace PraKtika.Views
{
    /// <summary>
    /// Логика взаимодействия для Page1.xaml
    /// </summary>
    public partial class Page1 : Page
    {
        public string[,] list = new string[632, 632];

        public Page1()
        {
            InitializeComponent();
            text.Text = "разверните список и выберите\nодин экспонат для просмотра\n подробной информации";
            using (var helper = new ExcelHelper())
            {
                try
                {
                    if (helper.Open(filePath: System.IO.Path.Combine(Environment.CurrentDirectory, "exhibits.xlsx")))
                    {
                        helper.CreateSpisok(list, Combo, helper);
                        helper.Dispose();
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }   
            }    
        }
        private void Click2(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new Page2());
        }

        private void Combo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int index = Combo.SelectedIndex;
            App.Current.Resources["Index"] = index;
            NavigationService.Navigate(new Page5());    
        }
    }
}

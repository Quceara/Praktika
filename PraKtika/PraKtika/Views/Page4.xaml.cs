using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using static System.Net.Mime.MediaTypeNames;

namespace PraKtika.Views
{
    /// <summary>
    /// Логика взаимодействия для Page4.xaml
    /// </summary>
    public partial class Page4 : Page
    {
        public Page4()
        {
            InitializeComponent();

            int index2 = (int)App.Current.Resources["Index2"];

            if (index2 == 0) 
            {
                Text.Visibility = Visibility.Hidden;
                Textblock.Text = "Нажмите далее и загрузите новую картинку";
            }
            

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
            using (var helper = new ExcelHelper())
            {
                try
                {
                    if (helper.Open(filePath: System.IO.Path.Combine(Environment.CurrentDirectory, "exhibits.xlsx")))
                    {
                        
                        int index = (int)App.Current.Resources["Index"];
                        int index2 = (int)App.Current.Resources["Index2"];
                        
                        if (index2 == 0)
                        {
                            OpenFileDialog openFileDialog = new OpenFileDialog();
                            openFileDialog.Filter = "Image files|*.bmp;*.jpg;*.png";
                            openFileDialog.FilterIndex = 1;
                            if (openFileDialog.ShowDialog() == true)
                            {

                                FileInfo fileInf = new FileInfo(openFileDialog.FileName);
                                if (fileInf.Exists)
                                {

                                    fileInf.CopyTo(System.IO.Path.Combine(Environment.CurrentDirectory, "Image", fileInf.Name), true);
                                }

                                helper.Set("F", index + 1, System.IO.Path.Combine(Environment.CurrentDirectory, "Image", fileInf.Name));
                            }
                        }
                        if (index2 == 1) 
                        {
                            helper.Set("A",index + 1,Text.Text);
                        }
                        else if (index2 == 2)
                        {
                            helper.Set("B", index + 1, Text.Text);
                        }
                        else if (index2 == 3)
                        {
                            helper.Set("C", index + 1, Text.Text);
                        }
                        else if (index2 == 4)
                        {
                            helper.Set("D", index + 1, Text.Text);
                        }
                        else if(index2 == 5)
                        {
                            helper.Set("E", index + 1, Text.Text);
                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
                helper.Save();
                helper.Dispose();
            }
            NavigationService.Navigate(new Page3());
        }
    }
}

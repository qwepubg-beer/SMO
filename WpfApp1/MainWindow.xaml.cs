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
using WpfApp1.Pages;
namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }


        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            MenuItem item = (MenuItem)sender;

            switch (item.Tag)
            {
                case "1.1":
                    MainFrame.NavigationService.Navigate(new Page1(1));
                    break;
                case "1.2":
                    MainFrame.NavigationService.Navigate(new Page1(2));
                    break;
                case "1.3":
                    MainFrame.NavigationService.Navigate(new Page1(3));
                    break;
                case "2.1":
                    MainFrame.NavigationService.Navigate(new Page2(1));
                    break;
                case "2.2":
                    MainFrame.NavigationService.Navigate(new Page2(2));
                    break;
                case "2.3":
                    MainFrame.NavigationService.Navigate(new Page2(3));
                    break;
            }
                
        }
    }
}

using System;
using System.Windows;
using System.Windows.Input;
using ExcelDiffToolView.Pages;

namespace UIKitTutorials
{
    public partial class MainWindow : Window
    {
        private static MainWindow _instance;

        public MainWindow()
        {
            InitializeComponent();
            _instance = this;
            PagesNavigation.Navigate(new System.Uri("Pages/DiffPage.xaml", UriKind.RelativeOrAbsolute));
        }

        public static MainWindow GetInstance()
        {
            return _instance ??= new MainWindow();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void btnRestore_Click(object sender, RoutedEventArgs e)
        {
            if (WindowState == WindowState.Normal)
                WindowState = WindowState.Maximized;
            else
                WindowState = WindowState.Normal;
        }

        private void btnMinimize_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void diff_Click(object sender, RoutedEventArgs e)
        {
            PagesNavigation.Navigate(new System.Uri("Pages/DiffPage.xaml", UriKind.RelativeOrAbsolute));
        }

        private void merge_Click(object sender, RoutedEventArgs e)
        {
            PagesNavigation.Navigate(new System.Uri("Pages/MergePage.xaml", UriKind.RelativeOrAbsolute));
        }

        private void about_Click(object sender, RoutedEventArgs e)
        {
            PagesNavigation.Navigate(new System.Uri("Pages/AboutPage.xaml", UriKind.RelativeOrAbsolute));
        }

        private void DragWindow(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void OpenPopup_Click(object sender, RoutedEventArgs e)
        {
            var popup = new PopupWindow
            {
                Owner = this
            };
            popup.ShowDialog();
        }
    }
}
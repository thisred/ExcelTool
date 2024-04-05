using System.Windows;

namespace ExcelDiffToolView.Pages
{
    public partial class PopupWindow : Window
    {
        public PopupWindow()
        {
            InitializeComponent();
        }

        private void ClosePopup_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
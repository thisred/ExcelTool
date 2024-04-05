using System.Windows;
using System.Windows.Controls;
using ExcelDiffToolView.Model;
using Microsoft.Win32;

namespace ExcelDiffToolView.Pages
{
    public partial class MergePage : Page
    {
        public ExcelMergeManager ExcelMerge { get; set; } = new();

        public MergePage()
        {
            InitializeComponent();
        }

        private void SelectSourceFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                SourceFile.Text = openFileDialog.FileName;
            }

            string path = openFileDialog.FileName;
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            ExcelMerge.SelectSourceExcelFile(path);
        }

        private void SelectMergeFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                MergeFile.Text = openFileDialog.FileName;
            }

            string path = openFileDialog.FileName;
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            ExcelMerge.MergeTable = GameManager.LoadExcelFile(path);
        }

        private async void MergeExcelFile_Click(object sender, RoutedEventArgs e)
        {
            bool success = false;
            await Task.Run(async () =>
            {
                try
                {
                    success = await ExcelMerge.SaveResult();
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.ToString(), "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            });
            if (success)
            {
                MessageBox.Show($"已经将结果保存到桌面，请打开桌面查看", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("是不是有啥东西填错了，还是说表格不对应？一定不是bug吧", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Clear(object sender, RoutedEventArgs e)
        {
            SourceFile.Clear();
            MergeFile.Clear();
            ExcelMerge.Reset();
        }
    }
}
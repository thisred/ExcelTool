using System.Windows;
using System.Windows.Controls;
using ExcelDiffToolView.Model;
using Microsoft.Win32;

namespace ExcelDiffToolView.Pages
{
    public partial class DiffPage : Page
    {
        public ExcelDiffManager ExcelData { get; set; } = new();

        public DiffPage()
        {
            InitializeComponent();
        }

        private void SelectNewExcel_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                NewFilePath.Text = openFileDialog.FileName;
            }

            string path = openFileDialog.FileName;
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            ExcelData.MyExcelTableNewVersion = GameManager.LoadExcelFile(path);
        }

        private void SelectOldExcel_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                OldFilePath.Text = openFileDialog.FileName;
            }

            string path = openFileDialog.FileName;
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            ExcelData.MyExcelTableOldVersion = GameManager.LoadExcelFile(path);
        }

        private void Reset(object sender, RoutedEventArgs e)
        {
            NewFilePath.Clear();
            OldFilePath.Clear();
            ExcelData.Reset();
        }

        private void CompareExcelFile_Click(object sender, RoutedEventArgs e)
        {
            if (ExcelData.CompareExcelFile())
            {
                string boxText =
                    $"共有{ExcelData.ExcelRowDiffs.Count}个差异，删除{ExcelData.RemovedRows.Count}个，新增{ExcelData.AddedRows.Count}个";
                MessageBox.Show(boxText, "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("是不是有啥东西填错了，还是说表格不对应？一定不是bug吧", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void SaveResult(object sender, RoutedEventArgs e)
        {
            ExcelData.CompareExcelFile();
            await Task.Run(async () =>
            {
                bool success = await ExcelData.SaveResult();
                if (success)
                {
                    string messageBoxText =
                        $"共有{ExcelData.ExcelRowDiffs.Count}个差异，删除{ExcelData.RemovedRows.Count}个，新增{ExcelData.AddedRows.Count}个,已经将结果保存到桌面，请打开桌面查看";
                    MessageBox.Show(messageBoxText, "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("是不是有啥东西填错了，还是说表格不对应？一定不是bug吧", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            });
        }
    }
}
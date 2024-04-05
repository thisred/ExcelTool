using System.Diagnostics;
using System.IO;
using System.Windows;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelDiffToolView.Model;

public class ExcelMergeManager
{
    public MyExcelTable MergeTable { get; set; } = new();

    public string TargetFilePath { get; set; }

    public void Reset()
    {
        MergeTable = new MyExcelTable();
        TargetFilePath = string.Empty;
    }

    public void SelectSourceExcelFile(string filePath)
    {
        TargetFilePath = filePath;
    }

    public async Task<bool> SaveResult()
    {
        if (string.IsNullOrWhiteSpace(MergeTable.FilePath) || string.IsNullOrWhiteSpace(TargetFilePath))
        {
            return false;
        }

        var fileInfo = new FileInfo(TargetFilePath);
        var sourcePackage = new ExcelPackage(fileInfo);
        var sourceWs = sourcePackage.Workbook.Worksheets[0];
        int dimensionRows = sourceWs.Dimension.Rows;
        int dimensionColumns = sourceWs.GetRealColumnCount();
        int max = MergeTable.ColumnIdxDict.Keys.Max() + 1;
        dimensionColumns = Math.Max(dimensionColumns, max);
        object[,] firstRow = ((object[,])sourceWs.Cells[1, 1, 1, dimensionColumns].Value);
        int idIndex = 0;
        for (int i = 0; i < firstRow.Length; i++)
        {
            if (firstRow[0, i] == null)
            {
                continue;
            }

            string item = (string)firstRow[0, i];
            if (item.StartsWith("#") || string.IsNullOrWhiteSpace(item))
            {
                continue;
            }

            string lower = item.ToLower();
            if (lower.Equals("id", StringComparison.CurrentCultureIgnoreCase))
            {
                idIndex = i;
            }
        }

        for (int i = 1; i <= dimensionRows; i++)
        {
            var rowRange = sourceWs.Cells[i, 1, i, dimensionColumns];
            object[,] row = (object[,])rowRange.Value;
            object o = row[0, 0];
            if (o != null)
            {
                string s = (string)o;
                if (s.StartsWith("#"))
                {
                    continue;
                }
            }

            object idObj = row[0, idIndex];
            if (idObj == null) continue;
            string cellValue = Convert.ToString(idObj);
            if (!int.TryParse(cellValue, out int id))
            {
                continue;
            }

            rowRange.Style.Fill.PatternType = ExcelFillStyle.None;
            if (MergeTable.Rows.TryGetValue(id, out var excelRow))
            {
                foreach ((int key, string value) in excelRow.Columns)
                {
                    object cell = row[0, key];
                    if (cell != null && Convert.ToString(cell) != value)
                    {
                        string columnName = MergeTable.ColumnIdxDict.GetValueOrDefault(key);
                        rowRange.SetCellValue(0, key, value);
                    }

                    if (cell == null)
                    {
                        rowRange.SetCellValue(0, key, value);
                    }
                }
            }
        }

        string fileName =
            $"{Path.GetFileNameWithoutExtension(TargetFilePath)}_{DateTime.Now:yyyy-MM-dd-HH-mm-ss}.xlsx";
        try
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string newFilePath = Path.Combine(desktopPath, fileName);
            await sourcePackage.SaveAsAsync(newFilePath);
            Process.Start(new ProcessStartInfo
            {
                FileName = newFilePath,
                UseShellExecute = true
            });
        }
        catch (Exception e)
        {
            MessageBox.Show($"{e}", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }

        return true;
    }
}
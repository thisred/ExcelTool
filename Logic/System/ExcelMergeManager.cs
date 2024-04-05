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

        var sourcePackage = new ExcelPackage(TargetFilePath);
        var sourceWs = sourcePackage.Workbook.Worksheets[0];
        int dimensionRows = sourceWs.Dimension.Rows;
        var sourceExcelColum = sourceWs.GetExcelColum();
        int max = MergeTable.ExcelColum.RealColumnCount;
        int dimensionColumns = Math.Max(sourceExcelColum.RealColumnCount, max);
        int idIndex = sourceExcelColum.NameToColumnIdx["id"];

        for (int i = 1; i <= dimensionRows; i++)
        {
            var sourceRowRange = sourceWs.Cells[i, 1, i, dimensionColumns];
            object[,] row = (object[,])sourceRowRange.Value;
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

            sourceRowRange.Style.Fill.PatternType = ExcelFillStyle.None;
            if (MergeTable.Rows.TryGetValue(id, out var excelRow))
            {
                foreach ((int key, string value) in excelRow.Columns)
                {
                    if (!MergeTable.ExcelColum.IdxToColumnName.TryGetValue(key, out string mergeColumn))
                    {
                        continue;
                    }

                    if (!sourceExcelColum.NameToColumnIdx.TryGetValue(mergeColumn, out int idx))
                    {
                        // 源文件中不存在被合并的文件的列
                        continue;
                    }

                    object cell = row[0, idx];
                    if (cell != null && Convert.ToString(cell) != value)
                    {
                        sourceRowRange.SetCellValue(0, idx, value);
                    }

                    if (cell == null)
                    {
                        sourceRowRange.SetCellValue(0, idx, value);
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
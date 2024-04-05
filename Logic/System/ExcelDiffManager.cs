using System.Diagnostics;
using System.IO;
using System.Windows;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Color = System.Drawing.Color;

namespace ExcelDiffToolView.Model;

public class ExcelDiffManager
{
    public MyExcelTable MyExcelTableNewVersion { get; set; } = new();
    public MyExcelTable MyExcelTableOldVersion { get; set; } = new();

    public Dictionary<int, ExcelRowDiff> ExcelRowDiffs = new();
    public Dictionary<int, MyExcelRow> RemovedRows = new();
    public Dictionary<int, MyExcelRow> AddedRows = new();


    /// <summary>
    /// 比较两个Excel文件的不同
    /// </summary>
    public bool CompareExcelFile()
    {
        if (string.IsNullOrWhiteSpace(MyExcelTableNewVersion.FilePath) ||
            string.IsNullOrWhiteSpace(MyExcelTableOldVersion.FilePath))
        {
            return false;
        }

        foreach ((int rowId, var row) in MyExcelTableNewVersion.Rows)
        {
            if (MyExcelTableOldVersion.Rows.TryGetValue(row.Id, out var value))
            {
                foreach (var column in row.Columns)
                {
                    if (value.Columns.TryGetValue(column.Key, out string oldColumn))
                    {
                        // 变化
                        if (oldColumn != column.Value)
                        {
                            ExcelRowDiffs.TryAdd(row.Id, new ExcelRowDiff());
                            ExcelRowDiffs[row.Id].DiffColumns.Add(new DiffColumn()
                            {
                                ColumnIdx = column.Key,
                                OldValue = oldColumn,
                                NewValue = column.Value
                            });
                        }
                    }
                    else
                    {
                        // 新增了一列
                        ExcelRowDiffs.TryAdd(row.Id, new ExcelRowDiff());
                        ExcelRowDiffs[row.Id].DiffColumns.Add(new DiffColumn()
                        {
                            ColumnIdx = column.Key,
                            OldValue = string.Empty,
                            NewValue = column.Value
                        });
                    }
                }
            }
            else
            {
                // 新增
                AddedRows[row.Id] = row;
            }
        }

        foreach ((int rowId, var row) in MyExcelTableOldVersion.Rows)
        {
            if (!MyExcelTableNewVersion.Rows.ContainsKey(row.Id))
            {
                // 删除
                RemovedRows[row.Id] = row;
            }
        }

        return true;
    }

    /// <summary>
    /// 重置
    /// </summary>
    public void Reset()
    {
        MyExcelTableNewVersion = new MyExcelTable();
        MyExcelTableOldVersion = new MyExcelTable();
        ExcelRowDiffs = new Dictionary<int, ExcelRowDiff>();
        RemovedRows = new Dictionary<int, MyExcelRow>();
        AddedRows = new Dictionary<int, MyExcelRow>();
    }


    /// <summary>
    /// 保存结果到新生成的Excel到桌面，文件名使用时间做后缀，并打开新生成的Excel
    /// </summary>
    public async Task<bool> SaveResult()
    {
        string newTableFilePath = MyExcelTableNewVersion.FilePath;
        if (string.IsNullOrWhiteSpace(newTableFilePath) ||
            string.IsNullOrWhiteSpace(MyExcelTableOldVersion.FilePath))
        {
            return false;
        }

        using var package = new ExcelPackage(newTableFilePath);
        var worksheet = package.Workbook.Worksheets[0];
        int dimensionRows = worksheet.Dimension.Rows;
        int dimensionColumns = MyExcelTableNewVersion.ExcelColum.RealColumnCount;

        int idIndex = MyExcelTableNewVersion.ExcelColum.NameToColumnIdx["id"];
        for (int i = 1; i <= dimensionRows; i++)
        {
            var rowRange = worksheet.Cells[i, 1, i, dimensionColumns];
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

            rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rowRange.Style.Fill.BackgroundColor.SetAuto();
            rowRange.Style.Fill.PatternType = ExcelFillStyle.None;
            if (ExcelRowDiffs.TryGetValue(id, out var diff))
            {
                foreach (var diffColumn in diff.DiffColumns)
                {
                    var worksheetCell = worksheet.Cells[i, diffColumn.ColumnIdx + 1];
                    worksheetCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheetCell.Style.Fill.BackgroundColor.SetColor(Color.Blue);
                }
            }

            if (AddedRows.ContainsKey(id))
            {
                rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rowRange.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            }
        }

        string fileName =
            $"{Path.GetFileNameWithoutExtension(newTableFilePath)}_{DateTime.Now:yyyy-MM-dd-HH-mm-ss}.xlsx";
        try
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string newFilePath = Path.Combine(desktopPath, fileName);
            await package.SaveAsAsync(newFilePath);
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
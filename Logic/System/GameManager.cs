using System.Diagnostics;
using System.IO;
using OfficeOpenXml;

namespace ExcelDiffToolView.Model;

public class GameManager
{
    public static GameManager Instance { get; } = new();

    static GameManager()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public static MyExcelTable LoadExcelFile(string filePath)
    {
        var excelTable = new MyExcelTable() { FilePath = filePath };
        using var p = new ExcelPackage();
        var sourceFile = new FileInfo(filePath);
        var sourcePackage = new ExcelPackage(sourceFile);
        var sourceWs = sourcePackage.Workbook.Worksheets[0];
        int dimensionRows = sourceWs.Dimension.Rows;
        int dimensionColumns = sourceWs.GetRealColumnCount();

        object[,] firstRowRange = ((object[,])sourceWs.Cells[1, 1, 1, dimensionColumns].Value);
        int idIndex = 0;
        for (int i = 0; i < firstRowRange.Length; i++)
        {
            if (firstRowRange[0, i] == null)
            {
                continue;
            }

            string item = Convert.ToString(firstRowRange[0, i]);
            if (item == null || item.StartsWith("#") || string.IsNullOrWhiteSpace(item))
            {
                continue;
            }

            string lower = item.ToLower();
            if (lower.Equals("id", StringComparison.CurrentCultureIgnoreCase))
            {
                idIndex = i;
            }

            excelTable.Columns.Add(lower);
            excelTable.ColumnIdxDict[i] = lower;
            excelTable.IdxColumnDict[lower] = i;
        }

        for (int i = 1; i <= dimensionRows; i++)
        {
            object[,] row = (object[,])sourceWs.Cells[i, 1, i, dimensionColumns].Value;
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

            var excelRow = new MyExcelRow
            {
                Id = id
            };
            for (int key = 0; key < row.Length; key++)
            {
                object o1 = row[0, key];
                string itemValue = o1 == null ? string.Empty : Convert.ToString(o1);
                excelRow.Columns[key] = itemValue;
            }

            if (!excelTable.Rows.TryAdd(id, excelRow))
            {
                // todo 添加日志
            }
        }

        return excelTable;
    }


    public static void OpenExcel(string filePath)
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = filePath,
            UseShellExecute = true
        });
    }
}
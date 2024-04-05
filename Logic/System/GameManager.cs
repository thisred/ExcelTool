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
        using var p = new ExcelPackage();
        var sourcePackage = new ExcelPackage(filePath);
        var sourceWs = sourcePackage.Workbook.Worksheets[0];
        var excelTable = new MyExcelTable
        {
            FilePath = filePath,
            ExcelColum = sourceWs.GetExcelColum()
        };
        int dimensionRows = sourceWs.Dimension.Rows;
        int dimensionColumns = excelTable.ExcelColum.RealColumnCount;
        int idIndex = excelTable.ExcelColum.NameToColumnIdx["id"];
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
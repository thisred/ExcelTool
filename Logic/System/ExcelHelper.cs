using OfficeOpenXml;

namespace ExcelDiffToolView.Model;

public static class ExcelHelper
{
    /// <summary>
    /// 获取Excel列
    /// </summary>
    public static MyExcelColum GetExcelColum(this ExcelWorksheet worksheet)
    {
        var excelColum = new MyExcelColum();
        int dimensionColumns = worksheet.GetRealColumnCount();
        excelColum.RealColumnCount = dimensionColumns;
        object[,] firstRowRange = ((object[,])worksheet.Cells[1, 1, 1, dimensionColumns].Value);
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
            excelColum.Columns.Add(lower);
            excelColum.IdxToColumnName[i] = lower;
            excelColum.NameToColumnIdx[lower] = i;
        }

        return excelColum;
    }

    /// <summary>
    /// 空列会导致程序异常缓慢
    /// 连续读到null列10列则会认为表格列数量不准确
    /// </summary>
    public static int GetRealColumnCount(this ExcelWorksheet worksheet)
    {
        int dimensionColumns = worksheet.Dimension.Columns;
        object[,] firstRow = ((object[,])worksheet.Cells[1, 1, 1, dimensionColumns].Value);
        int tempNullColumn = 0;
        int realColumn = 0;
        for (int i = 0; i < firstRow.Length; i++)
        {
            realColumn = i + 1;
            if (firstRow[0, i] == null)
            {
                tempNullColumn++;
                if (tempNullColumn > 10)
                {
                    realColumn -= tempNullColumn;
                    break;
                }
            }
            else
            {
                tempNullColumn = 0;
            }
        }

        return realColumn;
    }

    /// <summary>
    /// todo 获取准确的行数
    /// </summary>
    public static int GetRealRowCount(this ExcelWorksheet worksheet)
    {
        int realColumnCount = worksheet.GetRealColumnCount();
        int dimensionRows = worksheet.Dimension.Rows;
        object[,] firstColumns = ((object[,])worksheet.Cells[1, 1, dimensionRows, realColumnCount].Value);
        int tempNullRow = 0;
        int realRow = 0;
        for (int row = 0; row < firstColumns.Length; row++)
        {
            for (int column = 0; column < realColumnCount; column++)
            {
            }
        }

        return realRow;
    }
}
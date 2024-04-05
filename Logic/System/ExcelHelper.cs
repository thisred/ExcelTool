using OfficeOpenXml;

namespace ExcelDiffToolView.Model;

public static class ExcelHelper
{
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
}
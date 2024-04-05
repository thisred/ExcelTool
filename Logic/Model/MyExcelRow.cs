using System.Collections.Generic;

namespace ExcelDiffToolView.Model;

public class MyExcelRow
{
    public int Id { get; set; }

    /// <summary>
    /// 下标 => 值，从0开始
    /// </summary>
    public Dictionary<int, string> Columns { get; set; } = new();
}
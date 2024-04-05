using System.Collections.Generic;

namespace ExcelDiffToolView.Model;

public class MyExcelRow
{
    public int Id { get; set; }
    public Dictionary<int, string> Columns { get; set; } = new();
}
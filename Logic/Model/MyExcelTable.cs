using System.Collections.Generic;

namespace ExcelDiffToolView.Model;

public class MyExcelTable
{
    public string FilePath { get; set; }
    public Dictionary<int, MyExcelRow> Rows { get; set; } = new();
    public List<string> Columns { get; set; } = new();
    public Dictionary<int, string> ColumnIdxDict { get; set; } = new();
    public Dictionary<string, int> IdxColumnDict { get; set; } = new();
}
using System.Collections.Generic;

namespace ExcelDiffToolView.Model;

public class MyExcelTable
{
    public string FilePath { get; set; }
    public Dictionary<int, MyExcelRow> Rows { get; set; } = new();
    public MyExcelColum ExcelColum { get; set; } = new();
}
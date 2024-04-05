namespace ExcelDiffToolView.Model;

public class MyExcelColum
{
    public List<string> Columns { get; set; } = new();

    /// <summary>
    /// 列下标 => 小写列名
    /// </summary>
    public Dictionary<int, string> IdxToColumnName { get; set; } = new();

    /// <summary>
    /// 小写列名 => 列下标
    /// </summary>
    public Dictionary<string, int> NameToColumnIdx { get; set; } = new();

    /// <summary>
    /// 真正有数据的列数
    /// </summary>
    public int RealColumnCount { get; set; }
}
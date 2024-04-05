using System.Collections.Generic;

namespace ExcelDiffToolView.Model;

public class DiffColumn
{
    public int ColumnIdx { get; set; }
    public string OldValue { get; set; }
    public string NewValue { get; set; }
}

public class ExcelRowDiff
{
    public List<DiffColumn> DiffColumns { get; set; } = new();
}
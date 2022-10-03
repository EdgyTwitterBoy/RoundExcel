using RoundExcel.CellManagement;

namespace RoundExcel;

public class SheetInfo
{
    public List<string> ExcludeColumns { get; private set; }
    public List<int> ExcludeRows { get; private set; }

    public CellRange Range { get; private set; }
    public SheetInfo(string excludeColumns, string excludeRows, string range)
    {
        ExcludeColumns = new List<string>();
        ExcludeRows = new List<int>();
        Range = new CellRange(new Cell("A1"), new Cell(range));

        if (!string.IsNullOrEmpty(excludeColumns))
        {
            ExcludeColumns = excludeColumns.Split(',').ToList();
        }

        if (!string.IsNullOrEmpty(excludeRows))
        {
            ExcludeRows = excludeRows.Split(',').Select(int.Parse).ToList();
        }
    }

    public SheetInfo(){}
    
    public void SetExcludeColumns(string excludeColumns)
    {
        if (!string.IsNullOrEmpty(excludeColumns))
        {
            ExcludeColumns = excludeColumns.Split(',').Select(x => x.Trim()).ToList();
        }
    }
    
    public void SetExcludeRows(string excludeRows)
    {
        if (!string.IsNullOrEmpty(excludeRows))
        {
            ExcludeRows = excludeRows.Split(',').Select(x => x.Trim()).Select(int.Parse).ToList();
        }
    }
    
    public void SetRange(string range)
    {
        Range = new CellRange(new Cell("A1"), new Cell(range));
    }
}
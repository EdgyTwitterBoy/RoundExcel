namespace RoundExcel;

public class Cell
{
    public string Column;
    public int Row;
    
    public Cell(string column, int row)
    {
        Column = column;
        Row = row;
    }
}
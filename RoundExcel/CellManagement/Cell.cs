namespace RoundExcel.CellManagement;

public class Cell
{
    public string Column;
    public int Row;
    
    public Cell(string column, int row)
    {
        Column = column;
        Row = row;
    }

    public Cell(string cell)
    {
        Column = "";
        int index = 0;
        
        while (index < cell.Length && char.IsLetter(cell[index]))
        {
            Column += cell[index];
            index++;
        }
        
        Row = int.Parse(cell.Substring(index));
    }

    public override string ToString()
    {
        return $"{Column}{Row}";
    }
}
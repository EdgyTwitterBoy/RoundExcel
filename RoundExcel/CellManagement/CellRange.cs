namespace RoundExcel.CellManagement;

public class CellRange
{
    public Cell StartCell { get; set; }
    public Cell EndCell { get; set; }

    public CellRange(Cell startCell, Cell endCell)
    {
        StartCell = startCell;
        EndCell = endCell;
    }
    
    public List<Cell> GetCells()
    {   
        List<Cell> list = new();
        int rowStart = StartCell.Row;
        int rowEnd = EndCell.Row;
        int colStart = CellConverter.ConvertColumnToInt(StartCell.Column);
        int colEnd = CellConverter.ConvertColumnToInt(EndCell.Column);
        
        for (int row = rowStart; row <= rowEnd; row++)
        {
            for (int col = colStart; col <= colEnd; col++)
            {
                list.Add(new Cell(CellConverter.ConvertIntToColumn(col), row));
            }
        }
        
        return list;
    }
        
}
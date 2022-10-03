namespace RoundExcel;

public class CellRange
{
    public Cell StartCell { get; set; }
    public Cell EndCell { get; set; }

    public CellRange(Cell startCell, Cell endCell)
    {
        StartCell = startCell;
        EndCell = endCell;
    }
    
    // public List<Cell>()
    // {
    //     
    // }
        
}
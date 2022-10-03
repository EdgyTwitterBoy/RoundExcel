namespace RoundExcel.CellManagement;

public static class CellConverter
{
    private static string columnOrder = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

    public static int ConvertColumnToInt(string column)
    {
        return (column.Length - 1) * columnOrder.Length + columnOrder.IndexOf(column[^1]);
    }

    public static string ConvertIntToColumn(int index)
    {
        string column = "";
        int multiplier = index / columnOrder.Length;

        if (multiplier > 0) column += columnOrder[multiplier - 1];
        column += columnOrder[index - multiplier * columnOrder.Length];

        return column;
    }
}
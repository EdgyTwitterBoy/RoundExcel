using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using RoundExcel.CellManagement;

namespace RoundExcel;

public static class Program
{
    public static void Main(string[] args) => new App().Main(args);
}

public class App
{
    public void Main(string[] args)
    {
        string path = $@"./{GetFileName()}.xlsx";
        string newPath = path.Replace(".xlsx", "_rounded.xlsx");
        File.Copy(path, newPath, true);
        SheetInfo sheetInfo = new SheetInfo();
        sheetInfo.SetExcludeRows(GetExcludeRows());
        sheetInfo.SetExcludeColumns(GetExcludeColumns());
        sheetInfo.SetRange(GetSecondRangeCell());

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var package = new ExcelPackage(new FileInfo(newPath));
        var sheet = package.Workbook.Worksheets.First();
        Console.WriteLine($"{sheet.Cells["A1"].Text}");
        foreach (var cell in sheetInfo.Range.GetCells())
        {
            if (double.TryParse(sheet.Cells[cell.ToString()].Text, out var newValue))
            {
                if(sheetInfo.ExcludeRows.Contains(cell.Row) || sheetInfo.ExcludeColumns.Contains(cell.Column)) continue;
                Console.WriteLine("Changing value of cell {0} from {1} to {2}", cell, sheet.Cells[cell.ToString()].Text, RoundToSignificantDigits(newValue).ToString().Replace('.', ','));
                sheet.Cells[cell.ToString()].Value = RoundToSignificantDigits(newValue).ToString().Replace('.', ',');
            }
        }
        package.Save();
    }
    
    private double RoundToSignificantDigits(double d){
        if(d == 0)
            return 0;

        double scale = Math.Pow(10, Math.Floor(Math.Log10(Math.Abs(d))) + 1);
        return scale * Math.Round(d / scale, 3);
    }


    private string GetSecondRangeCell()
    {
        Console.Clear();
        Console.WriteLine("|A1|--|--|--|");
        Console.WriteLine("|--|--|--|--|");
        Console.WriteLine("|--|--|--|--|");
        Console.WriteLine("|--|--|--|??|\n");
        Console.WriteLine("Enter the bottom right cell of the range you want to round: ");
        return Console.ReadLine() ?? "";
    }

    private string GetExcludeRows()
    {
        Console.Clear();
        Console.WriteLine("Enter the rows to exclude (e.g. 1,2,3): ");
        return Console.ReadLine() ?? "";
    }

    private string GetExcludeColumns()
    {
        Console.Clear();
        Console.WriteLine("Enter the columns you want to exclude (e.g. A,B,C): ");
        return Console.ReadLine() ?? "";
    }

    private string GetFileName()
    {
        bool fileOk = false;
        string? fileName = "";
        
        while (!fileOk)
        {
            Console.WriteLine("Enter .xlsx file name: ");
            fileName = Console.ReadLine();
            fileOk = IsFileNameOk(fileName);
        }
        
        return fileName;
    }

    private bool IsFileNameOk(string? fileName)
    {
        Console.Clear();
        if (fileName == null)
        {
            Console.WriteLine("You need to enter a file name!");
            return false;
        }

        if (!File.Exists(@$"./{fileName}.xlsx"))
        {
            Console.WriteLine($"Couldn't find this file \"{fileName}.xlsx\". Did you enter the right name? Is it in the same folder as this program?");
            return false;
        }

        return true;
    }
}
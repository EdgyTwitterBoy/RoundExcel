using OfficeOpenXml;
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
        
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var package = new ExcelPackage(new FileInfo(path));
        var sheet = package.Workbook.Worksheets["List"];
        Console.WriteLine($"{sheet.Cells["A1"].Text}");
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
        if (fileName == null)
        {
            Console.WriteLine("You need to enter a file name!");
            return false;
        }

        if (!File.Exists(@$"./{fileName}.xlsx"))
        {
            Console.WriteLine("Couldn't find this file. Did you enter the right name? Is it in the same folder as this program?");
            return false;
        }

        return true;
    }
}
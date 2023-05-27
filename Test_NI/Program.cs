using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Data.Common;


var ex1 = @"C:\Users\roman\source\repos\Test_NI\Test_NI\Задание 1__Файл 1.xlsx";
var ex2 = @"C:\Users\roman\source\repos\Test_NI\Test_NI\Задание 1__Файл 2.xlsx";
string ex3 = "C:\\Users\\roman\\source\\repos\\Test_NI\\Test_NI\\Res.xlsx";
Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

try
{
    Workbook wb1 = app.Workbooks.Open(ex1);
    Workbook wb2 = app.Workbooks.Open(ex2);
    Workbook wb3 = app.Workbooks.Open(ex3);
    Worksheet sh1 = wb1.Sheets[1];
    Worksheet sh2 = wb2.Sheets[1];
    Worksheet sh3 = wb3.Sheets[1];
    Range rng1 = sh1.Range["B11", "B30"];
    Range rng2 = sh2.Range["C11", "O30"];
    Application excelApp = new Excel.Application();
    int resi = 1;
    int resj = 1;
    string text;
    string vivod;
    int pr;
    for (int i = 1; i <= rng1.Rows.Count; i++)
    {
        for (int j = 1; j <= rng1.Columns.Count; j++)
        {
            var val1 = rng1.Cells[i, j].Value;
            for (int l= 1; l <= rng1.Rows.Count; l++)
            { 
            var val2 = rng2.Cells[l, j].Value;

                if (val1 == val2)
                {
                  //  Console.WriteLine("Начало нового объекта");
                    for (int j1 = 1; j1 <= rng2.Columns.Count; j1++)
                    {
                        var val3 = rng1.Cells[i, j1].Value;
                        var val4 = rng2.Cells[l, j1].Value;
                        text = rng2.Cells[i, j1].Value?.ToString();
                        vivod ="";
                        if (val3 != val4)
                        {
                            //  rng2.Cells[l, j1].Interior.Color = 254;
                            sh3.Cells[resi, resj].Value = val4;
                            pr = j1 + 2;
                            vivod = "Строка:" + l.ToString() + " " + "Столбец:" + pr.ToString();
                            sh3.Cells[resi, resj + 1].Value = vivod;
                            resi++;
                        }
                       // Console.WriteLine(text)
                       // Console.WriteLine($"I = {i}");
                       // Console.WriteLine($"J1 = {j1}");
                    }
                }
            }
        }
    }

    app.DisplayAlerts = false;

    wb1.Close(false);
    wb2.Close(true, ex2);
    wb3.Close(true, ex3);

    Console.WriteLine("Успешно");
}
catch (Exception exp)
{
    Console.WriteLine(exp.Message);
}
finally
{
    app.Quit();
}
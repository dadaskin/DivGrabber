using QuoteGrabber5;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

// Top-level Statements 
// See https://aka.ms/new-console-template for more information

Console.WriteLine("DivGrabber v0.1 - Copyright(c) 2026 David R. Adaskin, all rights reserved");

Excel.Application _oXl;
Excel._Workbook _oWb;

List<StockInformation> mInvestments =  ReadSpreadsheet("FifthTestPortfolio.xls");
DoWebRequestAndParse(mInvestments);
UpdateSpreadsheet(mInvestments);

Console.WriteLine("DivGrabber done.  Review and then Save Spreadsheet manually");

// End of Top-level Statements

List<StockInformation> ReadSpreadsheet(string ssName)
{
    var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
    var ssPathName = documentsFolder + @"\" + ssName;

    Console.WriteLine($"Reading the Spreadsheet: {ssPathName}");

    try
    {
        // Open the spreadsheet
         _oXl = new Excel.Application { Visible = true };
        var z = Missing.Value;
        _oWb = _oXl.Workbooks.Open(ssPathName, z, z, z, z, z, z, z, z, z, z, z, z, z, z);

        foreach (var obj in _oWb.Sheets)
        {
            ReadSymbolsFromSheet((Excel._Worksheet)obj);
        }
    }
    catch (Exception err)
    {
        Console.WriteLine(err.Message);
    }


    StockInformation issue1 = new StockInformation("ABC", false, "Joint", "57");
    StockInformation issue2 = new StockInformation("DEF", true, "Joint", "58");

    return [issue1, issue2];
}

void ReadSymbolsFromSheet(Excel._Worksheet sheet)
{
    Console.WriteLine($"Reading from sheet: {sheet.Name}");
}

void DoWebRequestAndParse(List<StockInformation> mInvestments)
{
    Console.WriteLine("Doing Web Request TBD");
}

void UpdateSpreadsheet(List<StockInformation> mInvestments)
{
    Console.WriteLine("Updating Spreadsheet TBD");
}







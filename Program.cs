using QuoteGrabber5;
using System.Globalization;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

#region  Top-level Statements 
// See https://aka.ms/new-console-template for more information

Console.WriteLine("DivGrabber v0.1 - Copyright(c) 2026 David R. Adaskin, all rights reserved");

IList<StockInformation> _mInvestments = [];
string _mDateString;
int _mLastSymbolRow;

// Excel Spreadsheet landmarks
const string MSymbolColStr = "A";
const string MPricePerShareCol = "C";
const string MAnnualDividendCol = "N";
const int MFirstSymbolRow = 5;
const string MSymbolListTerminationString = "Cash";
const string MTimeStampCell = "C1";

Excel.Application _oXl;
Excel._Workbook _oWb;

ReadSpreadsheet("FifthTestPortfolio.xls");
DisplayResults();


DoWebRequestAndParse(_mInvestments);
UpdateSpreadsheet(_mInvestments);

Console.WriteLine("DivGrabber done.  Review and then Save Spreadsheet manually");

#endregion End of Top-level Statements

#region General Utils

void DisplayResults()
{
    foreach (var info in _mInvestments)
    {
        Console.WriteLine($"{info.Symbol,-14} {info.SheetName,-15} {info.RowStr,4} {info.PricePerShareStr,8} {info.AnnualDividend,5} {info.YearRange}");
    }
}

#endregion General Utils

#region Read Spreadsheet methods
void ReadSpreadsheet(string ssName)
{
    var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
    var ssPathName = documentsFolder + @"\" + ssName;

    Console.WriteLine($"---Reading the Spreadsheet: {ssPathName}");
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
        Console.WriteLine(err.ToString());
    }
}

void ReadSymbolsFromSheet(Excel._Worksheet sheet)
{
    Console.WriteLine($"-Reading from sheet: {sheet.Name}");

    var row = MFirstSymbolRow;

    var symbol = string.Empty;
    while (symbol != MSymbolListTerminationString)
    {
        var cell = MSymbolColStr + row.ToString(CultureInfo.InvariantCulture);
        var symbolRng = sheet.Range[cell, cell];
        symbol = (string)symbolRng.Value2;

        // Do we need to "clean" the symbols?

        if ((symbol != null) && (symbol != MSymbolListTerminationString))
        {
            var parseAsFund = (symbol.EndsWith("X") &&
                              (symbol != "CVX") &&
                              (symbol != "FAX"));
            var issue = new StockInformation(symbol, parseAsFund, sheet.Name, row.ToString(CultureInfo.InvariantCulture));

            var alreadyInList = _mInvestments.Any(item => item.Symbol == issue.Symbol);
            if (!alreadyInList)
            {
                _mInvestments.Add(issue);
            }
        }
        else if ((symbol != null) && (symbol == MSymbolListTerminationString) && sheet.Name.Contains("Joint"))
        {
            _mLastSymbolRow = row - 1;
        }

        row++;
    }
}

#endregion Read Spreadsheet methods


void DoWebRequestAndParse(IList<StockInformation> mInvestments)
{
    Console.WriteLine("---Doing Web Request TBD");
}

void UpdateSpreadsheet(IList<StockInformation> mInvestments)
{
    Console.WriteLine("---Updating Spreadsheet TBD");
}







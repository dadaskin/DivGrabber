using DivGrabber;
using System.Globalization;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

internal class Program
{
    private readonly IList<StockInformation> _mInvestments = [];

    // Excel Spreadsheet landmarks
    const int MFirstSymbolRow = 5;
    const string MSymbolColStr = "A";
    const string MNumSharesStr = "B";
    const string MAcqDateColStr = "F";
    const string MCumDivColStr = "I";
    const string MSymbolListTerminationString = "Cash";

    Excel.Application? _oXl;
    Excel._Workbook? _oWb;

    private static void Main(string[] args)
    {
        _ = new Program();
    }

    public Program()
    {
        Console.WriteLine("DivGrabber v0.1 - Copyright(c) 2026 David R. Adaskin, all rights reserved");
        ReadSpreadsheet("FifthTestPortfolio.xls");
        DisplayResults();

        DoWebRequestAndParse(_mInvestments);
     //   UpdateSpreadsheet(_mInvestments);

        Console.WriteLine("DivGrabber done.  Review and then Save Spreadsheet manually");
    }

    #region General Utils
    void DisplayResults()
    {
        foreach (var info in _mInvestments)
        {
                string msg = $"{info.Symbol,-14} {info.BlockList.Count} block(s)\n";
                foreach (var block in info.BlockList)
                {
                    msg += $"    {block.SheetName,-15} {block.RowStr,4} {block.AcquistionDate.ToString("dd-MMM-yyyy")}  {block.NumShares, 6} --- /* {block.CumulativeDividend}*/\n";
                }
            Console.Write(msg);
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
            // Read Symbol
            var cell = MSymbolColStr + row.ToString(CultureInfo.InvariantCulture);
            var symbolRng = sheet.Range[cell, cell];
            symbol = (string)symbolRng.Value2;


            if (symbol != null && symbol != MSymbolListTerminationString)
            {
                symbol = CleanSymbol(symbol);

                var hasNoDividend = (symbol.EndsWith("X") && symbol != "CVX" && symbol != "FAX")||
                                    (symbol == "AMZN") ||
                                    (symbol == "BRKB") ||
                                    (symbol == "QYLD") ||
                                    (symbol == "GLD") ||
                                    (symbol.Length == 9);

                if (hasNoDividend)
                {
                    row++;
                    continue;
                }

                // Read Number of Shares
                cell = MNumSharesStr + row.ToString(CultureInfo.InvariantCulture);
                var numSharesRng = sheet.Range[cell, cell];
                var numShares = (float)numSharesRng.Value2;

                // Read Acquisition Date
                cell = MAcqDateColStr + row.ToString(CultureInfo.InvariantCulture);
                var acqDateRng = sheet.Range[cell, cell];
                var acqDate = DateTime.FromOADate(acqDateRng.Value2);

                var block = new Block(sheet.Name, row.ToString(CultureInfo.InvariantCulture), acqDate, numShares);
                var alreadyInList = _mInvestments.Any(item => item.Symbol == symbol);

                if (alreadyInList)
                {
                    var issue = _mInvestments.First(a => a.Symbol == symbol);
                    issue.BlockList.Add(block);
                }
                else
                {
                    var issue = new StockInformation(symbol, block);
                    _mInvestments.Add(issue);
                }
            }
            row++;
        }
    }

    static string CleanSymbol(string symbol )
    {
        // Remove any extra information from symbol string
        var firstUnnecessaryCharacter = symbol.IndexOf(' ');
        if (firstUnnecessaryCharacter > 0)
            symbol = symbol.Remove(firstUnnecessaryCharacter);

        return symbol;
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
}


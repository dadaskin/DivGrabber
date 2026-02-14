
using System.Security.Cryptography.X509Certificates;

namespace DivGrabber
{
    public class StockInformation(string symbol, Block firstBlock)
    {
        public string Symbol { get; private set; } = symbol;
        public string HtmlDivHistory { get; set; } = "";
        public List<Block> BlockList { get; set; } = [firstBlock];
    }

    public class Block(string sheetName, string rowStr, DateTime acqDate, float numShares)
    {
        public string SheetName { get; set; } = sheetName;
        public string RowStr { get; set; } = rowStr;
        public DateTime AcquistionDate { get; set; } = acqDate;
        public float NumShares { get; set; } = numShares;
        public float CumulativeDividend { get; set; } = 0.0f;
    }
}

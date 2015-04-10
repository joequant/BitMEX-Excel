using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;

namespace BitMex
{
    //All static functions here are exposed as Excel functions (see example spreadsheet)
    public static class ExposedFunctions
    {
        public static object BitMexBid(string product, int level = 0)
        {
            return XlCall.RTD("BitMex.DataServer", null, product, DataPoint.Bid.ToString(), level.ToString());
        }

        public static object BitMexBidVol(string product, int level = 0)
        {
            return XlCall.RTD("BitMex.DataServer", null, product, DataPoint.BidVol.ToString(), level.ToString());
        }

        public static object BitMexAsk(string product, int level = 0)
        {
            return XlCall.RTD("BitMex.DataServer", null, product, DataPoint.Ask.ToString(), level.ToString());
        }

        public static object BitMexAskVol(string product, int level = 0)
        {
            return XlCall.RTD("BitMex.DataServer", null, product, DataPoint.AskVol.ToString(), level.ToString());
        }

        //download instruments. Default to downloading all, or set e.g. state="Open" to download open contracts only
        public static object BitMexInstruments(string state = null)
        {
            BitMexAPI _api = new BitMexAPI();
            List<BitMexInstrument> instruments = _api.DownloadInstrumentList(state);

            int rows = instruments.Count + 1; //add 1 for header row

            object[,] result = new string[rows, 12];

            result[0, 0] = "Symbol";
            result[0, 1] = "RootSymbol";
            result[0, 2] = "State";
            result[0, 3] = "Typ";
            result[0, 4] = "Expiry";
            result[0, 5] = "TickSize";
            result[0, 6] = "Multiplier";
            result[0, 7] = "PrevClosePrice";
            result[0, 8] = "TotalVolume";
            result[0, 9] = "Volume";
            result[0, 10] = "Vwap";
            result[0, 11] = "OpenInterest";

            for (int i = 0; i < instruments.Count; i++)
            {
                result[i + 1, 0] = instruments[i].symbol;
                result[i + 1, 1] = instruments[i].rootSymbol;
                result[i + 1, 2] = instruments[i].state;
                result[i + 1, 3] = instruments[i].typ;
                result[i + 1, 4] = instruments[i].expiry;
                result[i + 1, 5] = instruments[i].tickSize.ToString();
                result[i + 1, 6] = instruments[i].multiplier.ToString();
                result[i + 1, 7] = instruments[i].prevClosePrice.ToString();
                result[i + 1, 8] = instruments[i].totalVolume.ToString();
                result[i + 1, 9] = instruments[i].volume.ToString();
                result[i + 1, 10] = instruments[i].vwap.ToString();
                result[i + 1, 11] = instruments[i].openInterest.ToString();
            }

            // Call Resize via Excel - so if the Resize add-in is not part of this code, it should still work.
            return XlCall.Excel(XlCall.xlUDF, "Resize", result);
        }
    }

}

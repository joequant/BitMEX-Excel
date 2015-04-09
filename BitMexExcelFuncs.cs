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

    }

}

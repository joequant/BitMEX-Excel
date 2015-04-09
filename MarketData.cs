using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BitMex
{
    public class MarketDataSnapshot
    {
        public string Product;

        //BBO
        public decimal Bid { get { return BidDepth.Count > 0 ? BidDepth[0].Price : 0M; } }
        public decimal BidVol { get { return BidDepth.Count > 0 ? BidDepth[0].Qty : 0M; } }
        public decimal Ask { get { return AskDepth.Count > 0 ? AskDepth[0].Price : 0M; } }
        public decimal AskVol { get { return AskDepth.Count > 0 ? AskDepth[0].Qty : 0M; } }

        //Full depth
        public List<MarketDepth> BidDepth = new List<MarketDepth>();
        public List<MarketDepth> AskDepth = new List<MarketDepth>();

        public MarketDataSnapshot(string product)
        {
            Product = product;
        }

        public override string ToString()
        {
            return Product + "{ " + BidVol + "@" + Bid + " / " + Ask + "@" + AskVol + "} ";
        }
    }

    public class MarketDepth
    {
        public decimal Price;
        public decimal Qty;

        public MarketDepth() { }

        public MarketDepth(decimal price, decimal qty)
        {
            Price = price;
            Qty = qty;
        }
    }
}

using fastJSON;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using WebSocket4Net;

namespace BitMex
{
    internal class BitMexWebsocket
    {
        private WebSocket ws;
        public event EventHandler<MarketDataSnapshot> OnDataUpdate; 
        private List<string> SnapshotQueue = new List<string>();

        public void Reconnect()
        {
            //connect to public websocket
            ws = new WebSocket("wss://www.bitmex.com/realtime/websocket?heartbeat=true");
            ws.AutoSendPingInterval = 30;  //30 seconds ping
            ws.EnableAutoSendPing = true;

            //callbacks
            ws.MessageReceived += ws_MessageReceived;
            ws.Error += ws_Error;
            ws.Closed += ws_Closed;
            ws.Opened += ws_Opened;
            ws.Open();            
        }

        internal void Close()
        {
            if(ws != null)
                ws.Close();
        }

        private void ws_Opened(object sender, EventArgs e)
        {
            Logging.Log("Websocket opened");
        }

        private void ws_Closed(object sender, EventArgs e)
        {
            Logging.Log("Websocket closed");
        }

        private void ws_Error(object sender, SuperSocket.ClientEngine.ErrorEventArgs e)
        {
            Logging.Log("Websocket error {0}", e.Exception.Message);
            Reconnect();
        }

        private void ws_MessageReceived(object sender, MessageReceivedEventArgs e)
        {
            //pass to message cracker
            HandleData(e.Message);
        }

        public void GetSnapshot(string product)
        {
            //get snapshot
            if(ws != null && ws.State == WebSocketState.Open)
                SendSnapshotRequest(product);
            else
                SnapshotQueue.Add(product);
        }

        private void SendSnapshotRequest(string product)
        {
            if (ws == null || ws.State != WebSocketState.Open)
                Logging.Log("Cannot send snapshot request - websocket closed");
            else
            {
                string request = "{\"op\":\"getSymbol\", \"args\": [\"" + product + "\"]}";
                Logging.Log("Send snapshot request: {0}", request);
                ws.Send(request);
            }
        }


        public void HandleData(string data)
        {
            try
            {
                //Logging.Log("Incoming {0}", data);

                //subscriptions
                if (data.Contains("Welcome"))
                {
                    //subscribe streaming depth changes (top 10 levels)
                    ws.Send("{\"op\": \"subscribe\", \"args\": [\"trade\", \"orderBook10\"]}");

                    //send request for any products
                    foreach (string product in SnapshotQueue)
                        SendSnapshotRequest(product);
                    SnapshotQueue.Clear();
                }
                else if (data.Contains("data"))
                {
                    //try to deserialise data
                    BitMexStream streamType = JSON.ToObject<BitMexStream>(data);
                    if (streamType.table.Equals("orderBook10"))
                    {
                        ProcessDepthSnapshot(data);
                    }
                    else if (streamType.table.Equals("orderBook25"))
                    {
                        //returned by getSymbol snapshot call
                        ProcessDepth(data);
                    }
                }
                else if (data.Contains("error"))
                {
                    Logging.Log("BitMexHandler data error {0}", data);
                    Reconnect();
                }
            }
            catch (Exception ex)
            {
                Logging.Log("BitMexHandler couldnt parse {0} {1}", data, ex.Message);
            }
        }

        private void ProcessDepthSnapshot(string data)
        {
            try
            {
                BitMexStreamDepthSnap depth = JSON.ToObject<BitMexStreamDepthSnap>(data);

                if (depth.data.Count != 1)
                {
                    Logging.Log("BitMexHandler missing depth snap data");
                    return;
                }
                BitMexDepthCache info = depth.data[0];
                string symbol = info.symbol;

                //parse into snapshot
                MarketDataSnapshot snap = new MarketDataSnapshot(symbol);

                foreach (object[] bid in info.bids)
                {
                    decimal price = Convert.ToDecimal(bid[0]);
                    int rawQty = Convert.ToInt32(bid[1]);
                    snap.BidDepth.Add(new MarketDepth(price, rawQty));
                }

                foreach (object[] ask in info.asks)
                {
                    decimal price = Convert.ToDecimal(ask[0]);
                    int rawQty = Convert.ToInt32(ask[1]);
                    snap.AskDepth.Add(new MarketDepth(price, rawQty));
                }

                //notify listeners
                if(OnDataUpdate != null)
                    OnDataUpdate(null, snap);

            }
            catch (Exception ex)
            {
                Logging.Log("BitMexHandler Depth {0}", ex.Message);
            }
        }

        private void ProcessDepth(string data)
        {
            try
            {
                BitMexStreamDepth depth = JSON.ToObject<BitMexStreamDepth>(data);

                if (depth.data.Count <= 0)
                    return;

                string symbol = depth.data[0].symbol;

                //create cache container if not already available
                MarketDataSnapshot snap = new MarketDataSnapshot(symbol);

                //normally returns "partial" on response to getSymbol
                if( depth.action.Equals("partial"))
                {
                    foreach (BitMexDepth info in depth.data)
                    {
                        int level = info.level;

                        //extend depth if new levels added
                        if ((info.bidPrice.HasValue || info.bidSize.HasValue) && snap.BidDepth.Count < level + 1)
                            for (int i = snap.BidDepth.Count; i < level + 1; i++)
                                snap.BidDepth.Add(new MarketDepth());

                        if ((info.askPrice.HasValue || info.askSize.HasValue) && snap.AskDepth.Count < level + 1)
                            for (int i = snap.AskDepth.Count; i < level + 1; i++)
                                snap.AskDepth.Add(new MarketDepth());

                        //update values, or blank out if values are null
                        if (info.bidPrice.HasValue || info.bidSize.HasValue)
                        {
                            if (info.bidPrice.HasValue)
                                snap.BidDepth[level].Price = info.bidPrice.Value;

                            if (info.bidSize.HasValue)
                                snap.BidDepth[level].Qty = info.bidSize.Value;
                        }

                        if (info.askPrice.HasValue || info.askSize.HasValue)
                        {
                            if (info.askPrice.HasValue)
                                snap.AskDepth[level].Price = info.askPrice.Value;

                            if (info.askSize.HasValue)
                                snap.AskDepth[level].Qty = info.askSize.Value;
                        }
                    }

                    //notify listeners
                    if (OnDataUpdate != null)
                        OnDataUpdate(null, snap);
                }

            }
            catch (Exception ex)
            {
                Logging.Log("BitMexHandler Depth", ex);
            }

        }

    }

    [Serializable]
    public class BitMexStream
    {
        public string table;
        public string action;
        public List<string> keys;
    }

    [Serializable]
    public class BitMexStreamDepthSnap
    {
        public string table;
        public string action;
        public List<string> keys;
        public List<BitMexDepthCache> data;
    }

    [Serializable]
    public class BitMexDepthCache
    {
        public string symbol;
        public long timestamp;

        public List<object[]> bids = new List<object[]>();
        public List<object[]> asks = new List<object[]>();
    }

    [Serializable]
    public class BitMexStreamDepth
    {
        public string table;
        public string action;
        //public List<string> keys;
        public List<BitMexDepth> data;
    }

    [Serializable]
    public class BitMexDepth
    {
        public string symbol;
        public int level;
        public int? bidSize;
        public decimal? bidPrice;
        public int? askSize;
        public decimal? askPrice;
        public long timestamp;

        public override string ToString()
        {
            return symbol + " " + level + " " + bidSize + " @ " + bidPrice + " / " + askSize + " @ " + askPrice;
        }
    }
}

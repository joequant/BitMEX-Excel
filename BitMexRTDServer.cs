using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Runtime.InteropServices;

using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using ExcelDna.Logging;

public class AutoRegisterAddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        ExcelIntegration.RegisterUnhandledExceptionHandler(ex => "!!! EXCEPTION: " + ex.ToString());
    }

    public void AutoClose()
    {
    }
}

namespace BitMex
{
    // Implement the Excel RTD COM interface that allows Excel to communicate with external data sources
    // Uses the Excel-DNA library (https://exceldna.codeplex.com/)

    // When the server is started, we start a background thread that connects a websocket to BitMex
    // and subscribes for changes. When new data comes in, we check the list of subscribed data from
    // Excel and update any topics that are subscribed. This causes a callback to Excel that triggers
    // a real-time update

    [ComVisible(true)]
    public class DataServer : ExcelRtdServer
    {
        private BitMexAPI _api;
        private Dictionary<string, TopicCollection> _topics;  //one topic corresponds to one unique item of data in Excel (product, datapoint, depth)
        private Dictionary<int, TopicSubscriptionDetails> _topicDetails;
        private DataCache _cache;
        
        public DataServer()
        {
            Logging.Log("DataServer created");
            _topics = new Dictionary<string, TopicCollection>();
            _topicDetails = new Dictionary<int, TopicSubscriptionDetails>();
            _cache = new DataCache();
        }

        //initialisation method called when data is first accessed from Excel
        protected override bool ServerStart()
        {
            Logging.Log("ServerStart");
            _api = new BitMexAPI();

            //download instruments to get last price snapshot
            List<BitMexInstrument> instruments = _api.DownloadInstrumentList(null);
            foreach (BitMexInstrument instrument in instruments)
                _cache.Instruments[instrument.symbol] = instrument;

            //then connect to websocket to get updates
            _api.OnDepthUpdate += _api_OnDataUpdate;  //subscribe to events so that we'll be notified when prices change
            _api.OnTradeUpdate += _api_OnTradeUpdate; //and when products trade
            _api.Reconnect();

            return true;
        }

        //called when all subscriptions have been removed from Excel
        protected override void ServerTerminate()
        {
            Logging.Log("ServerTerminate");
            _api.Close();
        }

        void _api_OnTradeUpdate(object sender, BitMexTrade trade)
        {
            //update data cache
            if(!_cache.Instruments.ContainsKey(trade.symbol))
            {
                Logging.Log("Market trade, but instrument not yet downloaded", trade.symbol, trade.price);
            }
            else
            {
                Logging.Log("Updating last price {0} {1}", trade.symbol, trade.price);
                _cache.Instruments[trade.symbol].lastPrice = trade.price;
            }

            //no current subscription for this product
            if (!_topics.ContainsKey(trade.symbol))
                return;

            Dictionary<Tuple<DataPoint, int>, Topic> productTopics = _topics[trade.symbol].TopicItems;

            //update any last price/size subscriptions
            Tuple<DataPoint, int> key = Tuple.Create(DataPoint.Last, 0);
            if (productTopics.ContainsKey(key))
                productTopics[key].UpdateValue(trade.price);

        }


        //callback from the background websocket thread. Checks and updates any matching topics.
        void _api_OnDataUpdate(object sender, MarketDataSnapshot snap)
        {
            //update data cache
            _cache.MarketData[snap.Product] = snap;

            //no current subscription for this product
            if (!_topics.ContainsKey(snap.Product))
                return;

            Dictionary<Tuple<DataPoint, int>, Topic> productTopics = _topics[snap.Product].TopicItems;

            //iterate down the depth on the bid side, updating any currently subscribed topics
            for (int level = 0; level < snap.BidDepth.Count; level++)
            {
                //update bid
                Tuple<DataPoint, int> key = Tuple.Create(DataPoint.Bid, level);
                if (productTopics.ContainsKey(key))
                    productTopics[key].UpdateValue(snap.BidDepth[level].Price);

                //update bidvol
                key = Tuple.Create(DataPoint.BidVol, level);
                if (productTopics.ContainsKey(key))
                    productTopics[key].UpdateValue(snap.BidDepth[level].Qty);
            }

            //iterate down the depth on the ask side, updating any currently subscribed topics
            for (int level = 0; level < snap.AskDepth.Count; level++)
            {
                //update ask
                Tuple<DataPoint, int> key = Tuple.Create(DataPoint.Ask, level);
                if (productTopics.ContainsKey(key))
                    productTopics[key].UpdateValue(snap.AskDepth[level].Price);

                //update askvol
                key = Tuple.Create(DataPoint.AskVol, level);
                if (productTopics.ContainsKey(key))
                    productTopics[key].UpdateValue(snap.AskDepth[level].Qty);
            }
        }

        int GetTopicId(Topic topic)
        {
            return (int)typeof(Topic)
                        .GetField("TopicId", BindingFlags.GetField | BindingFlags.Instance | BindingFlags.NonPublic)
                        .GetValue(topic);
        }

        //New RTD subscription callback
        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            object result;
            int topicId = GetTopicId(topic);

            Logging.Log("ConnectData: {0} - {{{1}}}", topicId, string.Join(", ", topicInfo));

            //parse and validate request
            //---

            //count and default parameters
            if (topicInfo.Count < 2)
                return ExcelErrorUtil.ToComError(ExcelError.ExcelErrorNA); //need at least a product and a datapoint
            else if (topicInfo.Count == 2)
                topicInfo.Add("0"); //default to top level of book

            //parse parameters
            string product = topicInfo[0];

            DataPoint dataPoint;
            bool dataPointOk = Enum.TryParse<DataPoint>(topicInfo[1], out dataPoint);

            int level;
            bool levelOk = Int32.TryParse(topicInfo[2], out level);

            //return error if parameters are invalid
            if(!dataPointOk || !levelOk)
                return ExcelErrorUtil.ToComError(ExcelError.ExcelErrorNA); //incorrect level or datapoint request

            if(level > 9 || level < 0)
                return ExcelErrorUtil.ToComError(ExcelError.ExcelErrorNA); //level out of range

            //store subscription request
            //---

            //if product has not yet been subscribed, create subscription container
            if (!_topics.ContainsKey(product))
                _topics[product] = new TopicCollection();

            //store subscription request
            _topics[product].TopicItems[Tuple.Create(dataPoint, level)] = topic;
            if(_topicDetails.ContainsKey(topicId))
                Logging.Log("ERROR: duplicate topicId: {0} {1}", topicId, _topicDetails[topicId]);
            _topicDetails.Add(topicId, new TopicSubscriptionDetails(product, dataPoint, level));


            //return data
            //--
            if(dataPoint == DataPoint.Last)
            {
                if (!_cache.Instruments.ContainsKey(product))
                {
                    //may be empty when sheet is first loaded
                    result = ExcelErrorUtil.ToComError(ExcelError.ExcelErrorGettingData);             //return "pending data" to Excel
                }
                else
                {
                    //get data from cache
                    BitMexInstrument inst = _cache.Instruments[product];
                    switch(dataPoint)
                    {
                        case DataPoint.Last:
                            result = inst.lastPrice;
                            break;

                        default:
                            result = ExcelErrorUtil.ToComError(ExcelError.ExcelErrorNA);
                            break;
                    }
                }
            }
            else //bid/ask etc
            {
                if (!_cache.MarketData.ContainsKey(product))
                {
                    //if product is not yet in cache, request snapshot
                    //this can happen if a product doesnt update very frequently
                    _api.GetSnapshot(product);
                    result = ExcelErrorUtil.ToComError(ExcelError.ExcelErrorGettingData);             //return "pending data" to Excel
                }
                else
                {
                    //get data from cache
                    MarketDataSnapshot snap = _cache.MarketData[product];
                    switch(dataPoint)
                    {
                        case DataPoint.Bid:
                            result = snap.BidDepth[level].Price;
                            break;

                        case DataPoint.BidVol:
                            result = snap.BidDepth[level].Qty;
                            break;

                        case DataPoint.Ask:
                            result = snap.AskDepth[level].Price;
                            break;

                        case DataPoint.AskVol:
                            result = snap.AskDepth[level].Qty;
                            break;

                        default:
                            result = ExcelErrorUtil.ToComError(ExcelError.ExcelErrorNA);
                            break;
                    }
                }
            }

            return result;
        }

        //Called when an RTD subscription is no longer needed
        protected override void DisconnectData(Topic topic)
        {
            int topicId = GetTopicId(topic);
            Logging.Log("DisconnectData: {0}", topicId);

            //remove topic on unsubscribe
            if(!_topicDetails.ContainsKey(topicId))
            {
                Logging.Log("ERROR: no subscription for disconnect request");
            }
            else
            {
                //get corresponding details for this subscription
                TopicSubscriptionDetails subsDetails = _topicDetails[topicId];

                //remove from topics list
                if (_topics.ContainsKey(subsDetails.Product))
                {
                    Tuple<DataPoint, int> key = Tuple.Create<DataPoint, int>(subsDetails.DataPoint, subsDetails.Level);
                    _topics[subsDetails.Product].TopicItems.Remove(key);
                }

                //remove from details list
                _topicDetails.Remove(topicId);
            }

        }
    }


    //list of the different data points that we can provide
    internal enum DataPoint
    {
        Bid,
        Ask,
        BidVol,
        AskVol,
        Last
    }

    internal class TopicCollection
    {
        internal Dictionary<Tuple<DataPoint, int>, ExcelDna.Integration.Rtd.ExcelRtdServer.Topic> TopicItems = new Dictionary<Tuple<DataPoint, int>, ExcelRtdServer.Topic>();
    }

    internal class TopicSubscriptionDetails
    {
        internal string Product { get; private set; }
        internal DataPoint DataPoint { get; private set; }
        internal int Level { get; private set; }

        internal TopicSubscriptionDetails(string product, DataPoint dataPoint, int level)
        {
            Product = product;
            DataPoint = dataPoint;
            Level = level;
        }

        public override string ToString()
        {
            return Product + "|" + DataPoint + "|" + Level;
        }
        
    }
}

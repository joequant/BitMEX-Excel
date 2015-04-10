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
    [ComVisible(true)]
    public class DataServer : ExcelRtdServer
    {
        private BitMexAPI _api;
        private Dictionary<string, TopicCollection> _topics;  //one topic corresponds to one unique item of data in Excel (product, datapoint, depth)
        
        public DataServer()
        {
            Logging.Log("DataServer created");
            _topics = new Dictionary<string, TopicCollection>();
        }

        //called when data is first accessed from Excel
        protected override bool ServerStart()
        {
            Logging.Log("ServerStart");
            _api = new BitMexAPI();
            _api.OnDataUpdate += ws_OnDataUpdate; //subscribe to events when prices change
            _api.Reconnect();

            return true;
        }

        //called when all subscriptions have been removed from Excel
        protected override void ServerTerminate()
        {
            Logging.Log("ServerTerminate");
            _api.Close();
        }

        void ws_OnDataUpdate(object sender, MarketDataSnapshot snap)
        {
            Logging.Log("Data Update: {0}", snap);

            //no subscription for this product
            if (!_topics.ContainsKey(snap.Product))
                return;

            for (int level = 0; level < snap.BidDepth.Count; level++)
            {
                //update bid
                Tuple<DataPoint, int> key = Tuple.Create(DataPoint.Bid, level);
                if (_topics[snap.Product].TopicItems.ContainsKey(key))
                    _topics[snap.Product].TopicItems[key].UpdateValue(snap.BidDepth[level].Price);

                //update bidvol
                key = Tuple.Create(DataPoint.BidVol, level);
                if (_topics[snap.Product].TopicItems.ContainsKey(key))
                    _topics[snap.Product].TopicItems[key].UpdateValue(snap.BidDepth[level].Qty);
            }

            for (int level = 0; level < snap.AskDepth.Count; level++)
            {
                //update ask
                Tuple<DataPoint, int> key = Tuple.Create(DataPoint.Ask, level);
                if (_topics[snap.Product].TopicItems.ContainsKey(key))
                    _topics[snap.Product].TopicItems[key].UpdateValue(snap.AskDepth[level].Price);

                //update askvol
                key = Tuple.Create(DataPoint.AskVol, level);
                if (_topics[snap.Product].TopicItems.ContainsKey(key))
                    _topics[snap.Product].TopicItems[key].UpdateValue(snap.AskDepth[level].Qty);
            }
        }

        int GetTopicId(Topic topic)
        {
            return (int)typeof(Topic)
                        .GetField("TopicId", BindingFlags.GetField | BindingFlags.Instance | BindingFlags.NonPublic)
                        .GetValue(topic);
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            Logging.Log("ConnectData: {0} - {{{1}}}", GetTopicId(topic), string.Join(", ", topicInfo));

            //validate subscription request
            if (topicInfo.Count < 2)
                return ExcelErrorUtil.ToComError(ExcelError.ExcelErrorNA); //need at least a product and a datapoint
            else if (topicInfo.Count == 2)
                topicInfo.Add("0"); //default to top level

            //parse request
            string product = topicInfo[0];
            DataPoint dataPoint = (DataPoint)Enum.Parse(typeof(DataPoint), topicInfo[1]);
            int level = Int32.Parse(topicInfo[2]);

            //new product - request snapshot
            if (!_topics.ContainsKey(product))
            {
                _topics[product] = new TopicCollection();
                _api.GetSnapshot(product);
            }

            _topics[product].TopicItems[Tuple.Create(dataPoint, level)] = topic;

            //return "pending data"
            return ExcelErrorUtil.ToComError(ExcelError.ExcelErrorGettingData);
        }

        protected override void DisconnectData(Topic topic)
        {
            //TODO remove topic on unsubscribe
            Logging.Log("DisconnectData: {0}", GetTopicId(topic));
        }
    }


    //list of the different data points that we can provide
    internal enum DataPoint
    {
        Bid,
        Ask,
        BidVol,
        AskVol
    }

    internal class TopicCollection
    {
        internal Dictionary<Tuple<DataPoint, int>, ExcelDna.Integration.Rtd.ExcelRtdServer.Topic> TopicItems = new Dictionary<Tuple<DataPoint, int>, ExcelRtdServer.Topic>();
    }

}

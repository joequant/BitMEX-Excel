using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BitMex
{
    internal class DataCache
    {
        //use thread-safe dictionaries as this will be accessed by both the websocket thread and the excel thread
        public ConcurrentDictionary<string, MarketDataSnapshot> MarketData { get { return _marketData; } }
        public ConcurrentDictionary<string, BitMexInstrument> Instruments { get { return _instruments; } }

        private ConcurrentDictionary<string, MarketDataSnapshot> _marketData = new ConcurrentDictionary<string, MarketDataSnapshot>();
        private ConcurrentDictionary<string, BitMexInstrument> _instruments = new ConcurrentDictionary<string, BitMexInstrument>();
    }
}

# BitMexExcel
Excel plugin for the BitMex exchange. Provides real-time market data from BitMex into Excel.

Compile with Visual Studio Express 2012 for Windows Desktop, or just grab the XLL file.

## Functions
Currently provides bid and ask prices and sizes for the top 10 levels of each product (Open the included .xlx for examples)

<pre>
=BitMexBid(product, depthlevel)
=BitMexBidVol(product, depthlevel)
=BitMexAsk(product, depthlevel)
=BitMexAskVol(product, depthlevel)
</pre>

## Compiling
Install NuGet libraries:
* Install-Package Excel-DNA
* Install-Package WebSocket4Net
* Install-Package fastJSON


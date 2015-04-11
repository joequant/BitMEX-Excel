# BitMexExcel
Excel plugin for the BitMex exchange. Provides real-time market data from BitMex into Excel.

Released under Apache 2.0 licence, so free to use for commercial as well as non-commercial purposes.

Compile with Visual Studio Express 2012 for Windows Desktop, or just grab the XLL file from the Releases.

## Installation
Download the .xll file from the selected release (ensure you have the correct 32-bit or 64-bit release)
Open Excel 2013. Select File.. Options.. Add-Ins, and click Manage "Excel Add-ins" Go...
Click "Browse" and select the relevant .xll file to install.

## Functions
Currently provides bid and ask prices and sizes for the top 10 levels of each product (Open the included .xlsx for examples)

<pre>
//get bid/ask info. Omit depthlevel to get top of book
=BitMexBid(product, {depthlevel})
=BitMexBidVol(product, {depthlevel})
=BitMexAsk(product, {depthlevel})
=BitMexAskVol(product, {depthlevel})
=BitMexLastPrice(product)

//download all instruments as an array into excel. pass state to download particular instruments in a particular state (e.g. "Open") only, or omit to get all instruments.
=BitMexInstruments({state})
=BitMexInstrumentsActive()   //equivalent to =BitMexInstruments("Open")
</pre>

## Compiling
Download and install Visual Studio Express 2012 for Windows Desktop.

Install NuGet libraries:
* Install-Package Excel-DNA
* Install-Package WebSocket4Net
* Install-Package fastJSON

Compile the project. Install the resulting .xll file into Excel.

## Issues
Check the log file (notepad %temp%\BitMexRTD.log) for any issues.

# Overview

The Intrinio Excel Add-in extends the functionality of Microsoft Excel by enabling you to access the Intrinio API without any programming experience.  This Excel Add-in works on both Mac OS X and Microsoft Windows versions of Excel.

With the Intrinio Excel Add-in, you can access your Intrinio Data Feeds through the various Excel custom functions.  

# Download Intrinio Excel Add-in

Below are the links for the latest versions of the Intrinio Excel Add-in.  Select the appropriate version (Windows or Mac) to download.  Please see the install instructions below to begin using the Intrinio Excel add-in

[Microsoft Windows Intrinio Excel Add-in](https://s3.amazonaws.com/intrinio-production/intrinio-excel-addin/Intrinio_Excel_Addin.exe)

[Mac OS X Intrinio Excel Add-in](https://s3.amazonaws.com/intrinio-production/intrinio-excel-addin/Intrinio_Excel_Addin.zip)

# GitHub

The Intrinio Microsoft Excel add-in is released open source on [github](https://github.com/intrinio/intrinio-excel).  If you have any questions, find bugs, or wish to improve the functionality, please feel free to contribute in the [intrinio-excel](https://github.com/intrinio/intrinio-excel) repository.

# System Requirements

## Supported Operating Systems

The Intrinio Excel Add-in for Microsoft Excel is supported on both Windows and Mac, including the following operating systems:

**Microsoft Windows**

*   Microsoft Windows 10 - (both 32-bit and 64-bit versions)
*   Microsoft Windows 8.1 - (both 32-bit and 64-bit versions)
*   Microsoft Windows 8 - (both 32-bit and 64-bit versions)
*   Microsoft Windows 7 -  (both 32-bit and 64-bit versions)
*   Microsoft Windows Vista - (both 32-bit and 64-bit versions)
*   Microsoft Windows XP - (both 32-bit and 64-bit versions)
*   Microsoft Windows Server 2012
*   Microsoft Windows Server 2008

**Mac OS X**

*   Mac OS X version 10.5.8 or a later version

## Additional Software Requirements

One of the following versions of Microsoft Office is required:

**Microsoft Windows**

*   Microsoft Excel 2016 (both 32-bit and 64-bit versions)
*   Microsoft Excel 2013 (both 32-bit and 64-bit versions)
*   Microsoft Excel 2010 (both 32-bit and 64-bit versions)

**Mac OS X**

*   Microsoft Excel 2016 (both 32-bit and 64-bit versions)
*   Microsoft Excel 2011 (both 32-bit and 64-bit versions)

The Intrinio Excel Add-in has functionality built in it for both the 32-bit and 64-bit versions of Excel. The Add-in is not currently supported on Microsoft Office web applications.

# Install Instructions

Below are quick install instructions and a Youtube video that walk you through the whole process one step at a time. It's helpful to follow along. If you still have trouble - contact us by clicking the Help button.  You'll only have to do this once, and we promise it's worth it.

## Windows

<iframe allowfullscreen="allowfullscreen" src="https://www.youtube.com/embed/LiyVf4wC0mg" frameborder="0" height="315" width="420"></iframe>

1.  Download the latest version of the Intrinio Excel add-in by clicking on the link above.
2.  The Intrinio Excel add-in will be saved as an Application in your Downloads folder. Find it there, and double click on it.
3.  You'll see a pop-up titled "Intrinio Excel Add-in". Check the box to agree to the terms and conditions, then click "Install".
4.  After a few moments the add-in will be installed, and you can click "Finish".
5.  Open Microsoft Excel and go to the Intrinio tab on the top ribbon. Select API Keys and a prompt will pop up asking for your API Username and API Password.
6.  Copy the API Username and paste it in the User API Username field into the prompt in Excel.
7.  Copy the API Password and paste it in the API Password field into the prompt in Excel.
8.  Click the "Start" button to begin using the Intrinio API via the Intrinio Excel Add-in!
9.  You can begin by building your own spreadsheet or by opening a template.

## Mac OS

<iframe allowfullscreen="allowfullscreen" src="https://www.youtube.com/embed/tptSG9epqjA" frameborder="0" height="315" width="420"></iframe>

1.  Download the latest version of the Intrinio Excel add-in by clicking on the link above.
2.  Extract the Intrinio Excel add-in zip file to a folder on your hard drive. (On Mac OS X it may automatically extract it for you.)
3.  NOTE: If you've already done this once OR if you are updating the add-in, be sure to overwrite the old file.
4.  DO NOT click on Intrinio_Excel_Addin.xlam - you'll select this later.
5.  Open a BLANK Microsoft Excel worksheet on a Windows or Mac computer.
6.  Open the Intrinio Excel Add-in through the Excel Add-in manager.  Click on the "Tools" top menu, Add-Ins >> Select, and navigate to the folder where you extracted the Intrinio Excel Add-In and select it.
7.  Close out the blank Excel spreadsheet
8.  In the next step: if an alert pops up asking you about macros, be sure to click "Enable Macros" (macros are required for the add-in to work - on Windows you can permanently enable them - check our [Youtube](http://intrin.io/1OPsvMC) for instructions).
9.  In the next step: if an alert pops up asking you about links, be sure to click "Ignore Links" or "Do Not Update Links" (this only pops up in the template).
10.  Navigate to your Intrinio_Excel_Addin folder. Open the "Templates" folder, then the "Industrials" folder. Open: IntrinioFinancialData-Industrials.xlsm (YOU MUST OPEN THIS FILE FIRST).
11.  A prompt will pop up and ask for your API Username and API Password - which you can find on the credentials tab at the top of the [Data Feed page](https://www.intrinio.com/datafeed).
12.  Copy the API Username and paste it in the User API Username field into the prompt in Excel.
13.  Copy the API Password and paste it in the API Password field into the prompt in Excel.
14.  Click the "Start" button to begin using the Intrinio API via the Intrinio Excel Add-in! In this template you can enter a ticker to get started.
15.  You can use any of the templates provided in the Intrinio Excel Add-In folder that you downloaded, or you can build your own models in Excel.

A whole data buffet is being pulled into Excel. It may take longer than you expect to populate. On Windows computers you will see text at the very bottom of Excel that says "Calculating". This means the data is still downloading. On Mac computers you will see text at the very bottom that says "Ready" when the data has finished downloading. If you click on anything while the data is downloading it will delay the process, so please be patient.

Note: The template is just an illustration of the breadth of data we provide. Click in any cell to see the formula being used to pull that data in. Feel free to use the template if it is helpful to you, but there are no limits to the spreadsheets and models you can create on your own. Once you've installed the add-in, these formulas will work in any workbook in Excel. You can create, save, re-open and edit - and the data will be updated continuously.

# Intrinio Excel Functions

Below are all of the Excel custom functions for accessing the Intrinio API through the Excel add-in.

## IntrinioDataPoint

**=`IntrinioDataPoint(identifier,item)`**  
Returns that most recent data point for a selected identifier (ticker symbol, CIK ID, Federal Reserve Economic Data Series ID, etc.) for a selected tag. The complete list of tags available through this function are available <a href="/tags.html#data-point" target="_blank">here</a>. Income statement, cash flow statement, and ratios are returned as trailing twelve months values. All other data points are returned as their most recent value, either as of the last release financial statement or the most recent reported value.

### Parameters

```
=IntrinioDataPoint("`AAPL","name")

Apple Inc.

=IntrinioDataPoint("0000320193","ticker")

AAPL

`=IntrinioDataPoint("AAPL","pricetoearnings")

17.8763

`=IntrinioDataPoint("AAPL","totalrevenue")

199800000000.0


=IntrinioDataPoint("FRED.GDP","value")

18,034.8

=IntrinioDataPoint("DMD.ERP","ttm_erp")

0.0612
```

*   **identifier** - an identifier for the company or data point, including the SEC CIK ID, FRED Series ID, or Damodaran ERP: **<a href="http://www.sec.gov/edgar/searchedgar/cik.htm" target="_blank">CENTRAL INDEX KEY</a> | <a href="/master/economic-indices.html" target="_blank">ECONOMIC INDICES</a> | <a href="/tags.html#dmd-erp" target="_blank">DAMODARAN ERP</a> | <a href="/master/us-securities.html#home" target="_blank">TICKER SYMBOL</a> | <a href="/master/stock-indices.html" target="_blank">INDEX SYMBOL</a>**
*   **item** - the specified standardized tag or series ID requested: **<a href="/tags.html#data-point" target="_blank">INTRINIO DATA POINT TAGS</a> | <a href="/tags.html#economic-data" target="_blank">ECONOMIC TAGS</a> | <a href="/tags.html#dmd-erp" target="_blank">DAMODARAN ERP</a>**

## IntrinioHistoricalData

**`=IntrinioHistoricalData(ticker,item,sequence,start_date,end_date,frequency,data_type)`**  
Returns that historical data for for a selected identifier (ticker symbol or index symbol) for a selected tag.  The complete list of tags available through this function are available <a href="/tags.html#historical_data" target="_blank">here</a>.  Income statement, cash flow statement, and ratios are returned as trailing twelve months values by default, but can be changed with the type parametrer.  All other historical data points are returned as their value on a certain day based on filings reported as of that date.

### Parameters

```
=IntrinioHistoricalData("AAPL","open_price",0)

121.85


=IntrinioHistoricalData("AAPL","adj_close_price",0,"2012-01-01","2012-12-31")

71.43

=IntrinioHistoricalData("AAPL","close_price",0,,,"yearly")

110.38
```

*   **ticker** - the stock market ticker symbol associated with the company's common stock or index.  If the company is foreign, use the stock exchange code, followed by a colon, then the ticker.  You may request up to 150 tickers at once by separating them by a coma (each ticker and item combination requested will count as 1 query of the API): **<a href="/master/us-securities.html" target="_blank">TICKER SYMBOL</a> | <a href="/master/stock-indices.html" target="_blank">INDEX SYMBOL</a> | <a href="/master/economic-indices.html" target="_blank">FRED SYMBOL</a> | <a href="/master/sic-indices.html" target="_blank">SIC SYMBOL</a> | <a href="https://www.ffiec.gov/nicpubweb/nicweb/SearchForm.aspx" target="_blank">RSSD ID</a>**
*   **item** - the specified standardized tag requested:  **<a href="/tags.html#historical-data" target="_blank">INTRINIO TAGS</a> | <a href="/tags.html#economic-data" target="_blank">ECONOMIC TAGS</a>**
*   **sequence** - an integer 0 or greater for calling a single historical data point from the first entry, based on sort order: **`0..last available`**
*   **start_date** (optional) - the first date in which historical stock prices are delivered - historical daily prices go back to 1996 for most companies, but some go back further to the 1970s or to the date of the IPO: **`YYYY-MM-DD`**
*   **end_date** (optional, default=today) - the last date in which historical stock prices are delivered - end of day prices are available around 5 p.m. EST and 15 minute delayed prices are updated every minute throughout the trading day: **`YYYY-MM-DD`**
*   **frequency** (optional, returns daily historical price data otherwise) - the frequency of the historical prices & valuation data: **`daily | weekly | monthly | quarterly | yearly`**
*   **data_type** (optional, returns trailing twelve months (TTM) for the income statement, cash flow statement and calculations, and quarterly (QTR) for balance sheet) - the type of periods requested - includes fiscal years for annual data, quarters for quarterly data and trailing twelve months for annual data on a quarterly basis OR the type of statistic requested when querying using the SIC Indices: **`( FY | QTR | TTM | YTD )`** OR **`( count | sum | max | 75thpctl | mean | median | 25thpctl | min )`**
*   **show_date** (optional, false by default, hence returning the value) if true, the function will return the date value instead of the data point value for the given query: **`true | false`**

## IntrinioHistoricalPrices

**`=IntrinioHistoricalPrices(ticker,item,sequence,start_date,end_date,frequency)`**  
Returns professional-grade historical stock prices for a company. New EOD prices are available at 5p.m. EST and intraday IEX real-time prices are updated every minute during the trading day. Historical prices are available back to 1996 or the IPO data in most cases, with some companies with data back to the 1970s. Data from Quandl and QuoteMedia.

### Parameters

```
=IntrinioHistoricalPrices("AAPL","open",0)

121.85

=IntrinioHistoricalPrices("AAPL","date",0,"2012-01-01","2012-12-31")

2012-12-31

=IntrinioHistoricalPrices("AAPL","adj_close",0,"2012-01-01","2012-12-31")

71.43

=IntrinioHistoricalPrices("AAPL","date",0,,,"yearly")

2014-12-31

=IntrinioHistoricalPrices("AAPL","close",0,,,"yearly")

110.38
```

*   **ticker** - the stock market ticker symbol associated with the companies common stock securities:**<a href="/master/us-securities.html#home" target="_blank">TICKER SYMBOL</a>**
*   **item** - the selected observation of the historical prices:**`date | open | high | low | close | volume | ex_dividend | split_ratio | adj_open | adj_high | adj_low | adj_close | adj_volume`**
*   **sequence** - an integer 0 or greater for calling a single stock historical stock price data point from the first entry, based on sort order:**`0..last available`**
*   **start_date** (optional, all historical prices for the security will be queried memory, which will result in a slower loading time) - the first date in which historical stock prices are delivered - historical daily prices go back to 1996 for most companies, but some go back further to the 1970s or to the date of the IPO:**`YYYY-MM-DD`**
*   **end_date** (optional, all historical prices for the security will be queried, which will result in a slower loading time) - the last date in which historical stock prices are delivered - end of day prices are available around 5 p.m. EST and 15 minute delayed prices are updated every minute throughout the trading day:**`YYYY-MM-DD`**
*   **frequency** (optional, daily data will be pulled in by default) - the last date in which historical stock prices are delivered - end of day prices are available around 5 p.m. EST and 15 minute delayed prices are updated every minute throughout the trading day: **`daily | weekly | monthly | quarterly | yearly`**

## IntrinioFundamentals

**`=IntrinioFundamentals(ticker,statement,period_type,sequence,item)`**  
Returns a list of available standardized fundamentals (fiscal year and fiscal period) for a given ticker and statement. Also, you may add a date and type parameter to specify the fundamentals you wish to be returned in the response.

### Parameters

```
=IntrinioFundamentals("AAPL","income_statement","FY",0,"end_date")

2014-09-27

=IntrinioFundamentals("AAPL","balance_sheet","QTR",0,"fiscal_period")

Q3

=IntrinioFundamentals("AAPL","balance_sheet","QTR",0,"fiscal_year")

2015
```

*   **ticker** - the stock market ticker symbol associated with the companies common stock securities: **<a href="/master/us-securities.html#home" target="_blank">TICKER SYMBOL</a>**
*   **statement** - the financial statement requested, options include the income statement, balance sheet, statement of cash flows and calculated metrics and ratios :**`income_statement | balance_sheet | cash_flow_statement | calculations`**
*   **period_type** - the type of periods requested - includes fiscal years for annual data, quarters for quarterly data and trailing twelve months for annual data on a quarterly basis: **`FY | QTR | TTM | YTD`**
*   **sequence** - an integer 0 or greater for calling a single fundamental from the first entry: **`0..last available`**
*   **item** - the return value for the fundamental: **`fiscal_year | fiscal_period | end_date | start_date`**

## IntrinioTags

**`=IntrinioTags(ticker,statement,sequence,item)`**  
Returns the As Reported XBRL tags and labels for a given ticker, statement, and date or fiscal year/fiscal quarter.

A basic list of all industrial standardized tags can be found <a href="/tags.html#industrial" target="_blank">here</a>.
A basic list of all financial standardized tags can be found <a href="/tags.html#financial" target="_blank">here</a>.

### Parameters

```
=IntrinioTags("AAPL","income_statement",0,"tag")

operatingrevenue

=IntrinioTags("AAPL","balance_sheet",3,"name")

Short-Term Investments
```

*   **ticker** - the stock market ticker symbol associated with the companies common stock securities: **<a href="/master/us-securities.html#home" target="_blank">TICKER SYMBOL</a>**
*   **statement** - the financial statement requested, options include the income statement, balance sheet, statement of cash flows, calculated metrics and ratios, and current data points :**`income_statement | balance_sheet | cash_flow_statement | calculations`**
*   **sequence** - an integer 0 or greater for calling a single tag from the first entry, based on order: **`0..last available`**
*   **item**  - the returned value for the data tag: **`name | tag | balance | unit`**

## IntrinioFinancials

**`=IntrinioFinancials(ticker,statement,fiscal_year,fiscal_period,tag,rounding)`**  
Returns professional-grade historical financial data. This data is standardized, cleansed and verified to ensure the highest quality data sourced directly from the XBRL financial statements. The primary purpose of standardized financials are to facilitate comparability across a single company's fundamentals and across all companies fundamentals.

For example, it is possible to compare total revenues between two companies as of a certain point in time, or within a single company across multiple time periods. This is not possible using the as reported financial statements because of the inherent complexity of reporting standards.

### Parameters

```
=IntrinioFinancials("AAPL","income_statement",2014,"FY","operatingrevenue","A")

182,795,000,000

=IntrinioFinancials("AAPL","balance_sheet",2,"QTR","totalequity","B")

123.328

=IntrinioFinancials("AAPL","income_statement",7,"TTM","netincometocommon","M")

37,037
```

*   **ticker** - the stock market ticker symbol associated with the companies common stock securities: **<a href="/master/us-securities.html#home" target="_blank">TICKER SYMBOL</a>**
*   **statement** - the financial statement requested, options include the income statement, balance sheet, statement of cash flows and calculated metrics and ratios : **`income_statement | balance_sheet | cash_flow_statement | calculations`**
*   **fiscal_year** - the fiscal year associated with the fundamental OR the sequence of the requested fundamental (i.e. 0 is the first available fundamental associated with the fiscal period type): **`YYYY`** OR **`0..last available`**
*   **fiscal_period** - the fiscal period associated with the fundamental, or the fiscal period type in association with the sequence selected in the fiscal year parameter: **`FY | Q1 | Q2 | Q3 | Q4 | Q1TTM | Q2TTM | Q3TTM | Q2YTD | Q3YTD `** OR **`FY | QTR | YTD | TTM`**
*   **tag** - the specified standardized tag: **<a href="/tags.html#industrial" target="_blank">STANDARDIZED INDUSTRIAL TAGS</a> | <a href="/tags.html#financial" target="_blank">STANDARDIZED FINANCIAL TAGS</a>**
*   **rounding** (optional, actuals by default) - round the returned value (actuals, thousands, millions, billions):**`A | K | M | B`**

## IntrinioReportedFundamentals

**`=IntrinioReportedFundamentals(ticker,statement,period_type,sequence,item)`**  
Returns an as reported fundamental (fiscal year, fiscal period, start date, and end date) for a given ticker and statement. Also, you may add a period type parameter to specify the fundamentals you wish to be returned in the response.

### Parameters

```
=IntrinioReportedFundamentals("AAPL","income_statement","FY",0,"fiscal_year")

2014

=IntrinioReportedFundamentals("AAPL","income_statement","QTR",2,"fiscal_period")

Q1

=IntrinioReportedFundamentals("AAPL","balance_sheet","QTR",5,"end_date")

2013-12-28
```

*   **ticker** - the stock market ticker symbol associated with the companies common stock securities:**<a href="/master/us-securities.html#home" target="_blank">TICKER SYMBOL</a>**
*   **statement** - the financial statement requested, options include the income statement, balance sheet and statement of cash flows: **`income_statement | balance_sheet | cash_flow_statement`**
*   **period_type** - the type of periods requested - includes fiscal years for annual data, quarters for quarterly data: **`FY | QTR`**
*   **sequence** - an integer 0 or greater for calling a single fundamental from the first entry: **`0..last available`**
*   **item** - the selected return value from the fundamental: **`fiscal_year | fiscal_period | end_date | start_date`**

## IntrinioReportedTags

**`=IntrinioReportedTags(ticker,statement,fiscal_year,fiscal_period,sequence,item)`**  
Returns the As Reported XBRL tags and labels for a given ticker, statement, and date or fiscal year/fiscal quarter.

### Parameters

```
=IntrinioReportedTags("AAPL","income_statement",2014,"FY",0,"name")

Net sales

=IntrinioReportedTags("AAPL","income_statement",2014,"FY",0,"tag")

SalesRevenueNet

=IntrinioReportedTags("AAPL","balance_sheet",7,"QTR",28,"name")

Retained earnings

=IntrinioReportedTags("AAPL","balance_sheet",7,"QTR",28,"tag")

RetainedEarningsAccumulatedDeficit
```

*   **ticker** - the stock market ticker symbol associated with the companies common stock securities:**<a href="/master/us-securities.html#home" target="_blank">TICKER SYMBOL</a>**
*   **statement** - the financial statement requested: **`income_statement | balance_sheet | cash_flow_statement`**
*   **fiscal_year** - the fiscal year associated with the fundamental OR the sequence of the requested fundamental (i.e. 0 is the first available fundamental associated with the fiscal period type): **`YYYY`** OR **`**`0..last available`**`**
*   **fiscal_period** - the fiscal period associated with the fundamental, or the fiscal period type in association with the sequence selected in the fiscal year parameter: **`FY | Q1 | Q2 | Q3 | Q4 | Q1TTM | Q2TTM | Q3TTM | Q2YTD | Q3YTD` **OR **`FY | QTR | YTD | TTM`**
*   **item** - the selected return value for the reported tags: **`name | tag | domain_tag | balance | unit`**

## IntrinioReportedFinancials

**`=IntrinioReportedFinancials(ticker,statement,fiscal_year,fiscal_period,xbrl_tag,domain_tag)`**  
Returns the financial data directly from the xbrl filing of the company's financial statements.

### Parameters

```
=IntrinioReportedFinancials("AAPL","income_statement",2014,"FY","SalesRevenueNet")

182795000000

=IntrinioReportedFinancials("AAPL","income_statement",2,"QTR","EarningsPerShareBasic")

3.08

=IntrinioReportedFinancials("AAPL","balance_sheet",2014,"FY","PropertyPlantAndEquipmentNet")

20,624,000,000

=IntrinioReportedFinancials("AAPL","balance_sheet",1,"QTR","LongTermDebt")

40,072,000,000
```

*   **ticker** - the stock market ticker symbol associated with the companies common stock securities: **<a href="/master/us-securities.html#home" target="_blank">TICKER SYMBOL</a>**
*   **statement** - the financial statement requested, options include the income statement, balance sheet, statement of cash flows and calculated metrics and ratios: **`income_statement | balance_sheet | cash_flow_statement`**
*   **fiscal_year** - the fiscal year associated with the fundamental OR the sequence of the requested fundamental (i.e. 0 is the first available fundamental associated with the fiscal period type): **`YYYY`** OR **`0..last available`**
*   **fiscal_period** - the fiscal period associated with the fundamental, or the fiscal period type in association with the sequence selected in the fiscal year parameter: **`FY | Q1 | Q2 | Q3 | Q4 | Q1TTM | Q2TTM | Q3TTM | Q2YTD | Q3YTD` **OR **`**`FY | QTR | YTD | TTM`**`**
*   **xbrl_tag** - the specified XBRL tag: **`All Available XBRL Tags`**
*   **domain_tag** (optional) - the specified domain XBRL tag, associated with certain data points on the financial statements that have a dimension associated with the data point: **`All Available Domain XBRL Tags`**

## IntrinioBankFundamentals

**`=IntrinioBankFundamentals(identifier,statement,period_type,sequence,item)`**  
Returns a list of available standardized fundamentals (fiscal year and fiscal period) for a given ticker and statement. Also, you may add a date and type parameter to specify the fundamentals you wish to be returned in the response.

### Parameters

```
=IntrinioBankFundamentals("STT","RI","FY",0,"end_date")

2015-12-31

=IntrinioBankFundamentals("STT","RC","YTD",1,"fiscal_period")

Q3

=IntrinioBankFundamentals("STT","RI","QTR",0,"fiscal_year")

2015
```

*   **identifier** - the stock market ticker symbol associated with the companies common stock securities or the Federal Reserve RSSD ID unique identifier for the bank: **<a href="/master/us-securities.html#home" target="_blank">TICKER SYMBOL</a> | <a href="https://www.ffiec.gov/nicpubweb/nicweb/SearchForm.aspx" target="_blank">RSSD ID LOOKUP</a>**
*   **statement** - (optional, returns all available statements) - the call report/UBPR financial statement requested
*   **period_type** - the type of periods requested - includes fiscal years for annual data, quarters for quarterly data and trailing twelve months for annual data on a quarterly basis: **`FY | YTD`**
*   **sequence** - an integer 0 or greater for calling a single fundamental from the first entry: **`0..last available`**
*   **item** - the return value for the fundamental: **`fiscal_year | fiscal_period | end_date | start_date`**

## IntrinioBankTags

**`=IntrinioBankTags(identifier,statement,sequence,item)`**  
Returns the Bank Call Report or UBPR Report XBRL tags and labels for a given identifier, statement, and date or fiscal year/fiscal quarter.

### Parameters

```
=IntrinioBankTags("STT","RI",10,"tag")

RIADB489

=IntrinioBankTags("STT","RC",20,"name")

Goodwill
```

*   **identifier** - the stock market ticker symbol associated with the companies common stock securities or the Federal Reserve RSSD ID unique identifier for the bank: **<a href="/master/us-securities.html#home" target="_blank">TICKER SYMBOL</a> | <a href="https://www.ffiec.gov/nicpubweb/nicweb/SearchForm.aspx" target="_blank">RSSD ID LOOKUP</a>**
*   **statement** - (optional, returns all available statements) - the call report/UBPR financial statement requested
*   **sequence** - an integer 0 or greater for calling a single tag from the first entry, based on order: **`0..last available`**
*   **item**  - the returned value for the data tag: **`name | tag | balance | unit`**


## IntrinioBankFinancials

**`=IntrinioBankFinancials(identifier,statement,fiscal_year,fiscal_period,tag,rounding)`**  
Returns professional-grade historical financial data for bank and bank holding companies. This data is directly from the Call Reports and UBPR Reports filed with the FDIC & Federal Reserve.

### Parameters

```
=IntrinioBankFinancials("STT","RI",2015,"FY","RIAD4107","K")

2,489,609

=IntrinioBankFinancials("STT","RC",2,"QTR","RCFDB529","M")

18,313.937
```

*   **identifier** - the stock market ticker symbol associated with the companies common stock securities or the Federal Reserve RSSD ID unique identifier for the bank: **<a href="/master/us-securities.html#home" target="_blank">TICKER SYMBOL</a> | <a href="https://www.ffiec.gov/nicpubweb/nicweb/SearchForm.aspx" target="_blank">RSSD ID LOOKUP</a>**
*   **statement** - (optional, returns all available statements) - the call report/UBPR financial statement requested
*   **fiscal_year** - the fiscal year associated with the fundamental OR the sequence of the requested fundamental (i.e. 0 is the first available fundamental associated with the fiscal period type): **`YYYY`** OR **`0..last available`**
*   **fiscal_period** - the fiscal period associated with the fundamental, or the fiscal period type in association with the sequence selected in the fiscal year parameter: **`FY | Q1 | Q2 | Q3 | Q4 | Q2YTD | Q3YTD `** OR **`FY | QTR | YTD | TTM`**
*   **tag** - the specified Call Report or UBPR Report XBRL Tag requested
*   **rounding** (optional, actuals by default) - round the returned value (actuals, thousands, millions, billions):**`A | K | M | B`**

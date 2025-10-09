# Dividend Data Spreadsheet Add-on Documentation
This add-on provides custom functions for fetching stock data from Dividend Data. It includes tools for dividends, statements, metrics, ratios, growth, quotes, profiles, funds, segments, KPIs, commodities, and batch quotes/dividends.
To use, install the add-on in Google Sheets and use the functions in cells.

_Below, I'll explain each of the functions available:_

## 1) DIVIDENDDATA
### Description
Retrieves various dividend-related data for a single stock symbol. This function is useful for analyzing a company's dividend payments, yields, growth trends, and sustainability through payout ratios. It can return single values for quick metrics or tables for historical data.

### Parameters

**symbol**: Stock ticker symbol (e.g., `"MSFT"`). Required.

**metric**: The dividend metric to retrieve (default: `"fwd_payout"`). 

_Available metrics:_ 
- `"fwd_payout"` (forward annual payout)
- `"ttm_payout"` (TTM payout)
- `"fwd_yield"` (forward yield)
- `"ttm_yield"` (TTM yield)
- `"frequency"` (payment frequency)
- `"history"` (historical table)
- `"growth"` (growth rates table)
- `"1y_cagr"` (1-year CAGR)
- `"3y_cagr"` (3-year CAGR)
- `"5y_cagr"` (5-year CAGR)
- `"10y_cagr"` (10-year CAGR)
- `"payout_ratio"` (EPS payout ratio)
- `"fcf_payout_ratio"` (FCF payout ratio).

**showHeaders**: Boolean to include headers in history or growth tables (default: false).

### Examples:

**Input**: `=DIVIDENDDATA("MSFT", "fwd_yield")`

**Output**: Forward dividend yield as decimal (e.g., 0.008).

**Input**: `=DIVIDENDDATA("MSFT", "history", TRUE)`

**Output**: Table with headers and full historical dividend data.

| Declaration Date  | Record Date | Payment Date  | Adjusted Dividend | Dividend  | Yield | Frequency |
| ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- |
| 9/14/2025 | 11/19/2025  | 12/10/2025  | $0.91  | $0.91  | 0.65%  | Quarterly |
| 6/9/2025 | 8/20/2025 | 9/10/2025  | $0.83  | $0.83 | 0.65%  | Quarterly |
| 3/10/2025  | 5/14/2025  | 6/11/2025  | $0.83  | $0.83  | 0.72%  | Quarterly |
| 12/2/2024 | 2/19/2025  | 3/12/2025  | $0.83  | $0.83  | 0.76%  | Quarterly  |

_The data above is formatted. In reality, it will return the raw numbers. You can choose how to format within the returned cells._


## 2) DIVIDENDDATA_BATCH
### Description
Retrieves batch dividend data for multiple stock symbols. This function is efficient for fetching dividend information across several tickers at once, such as latest payouts or historical tables. It supports both latest values and full history.

### Parameters

**symbols**: Comma-separated stock ticker symbols (e.g., `"MSFT,KMB,O"`). Required.

**metric**: Select a metric (default: `"fwd_payout"`). Or select multiple with comma-separated metrics _(Aside from `"all"` and `"history"`)_. 

- `"all"` Get all the latest dividend metrics for the provided tickers.
- `"history"` Get the dividend history of all provided tickers.

_Available metrics:_ 
- `"fwd_payout"`
- `"adjdividend"`
- `"dividend"`
- `"recorddate"`
- `"declarationdate"`
- `"paymentdate"`
- `"yield"`
- `"frequency"`

**showHeaders**: Boolean to include header row and symbol column (defaults to true for `"all"` or `"history"`).

### Examples

**Input**: `=DIVIDENDDATA_BATCH("MSFT,AAPL", "fwd_payout,yield", TRUE)`

**Output**: Table with symbols, forward payouts, and yields.

**Input**: `=DIVIDENDDATA_BATCH("MSFT,KMB", "history")`

**Output**: Flat historical dividend table for all symbols.

## 3) DIVIDENDDATA_STATEMENT
### Description
Retrieves full financial statements for a stock. This function is useful for in-depth financial analysis, providing complete income statements, balance sheets, or cash flow statements over time. It supports filtering by period and year.

### Parameters

**symbol**: Stock ticker symbol (e.g., `"MSFT"`). Required.

**metric**: The statement type: `"income"`, `"balance"`, or `"cash_flow"`. Required.

**showHeaders**: Boolean to include header row (default: `false`).

**period**: Period filter: `"FY"`, `"Q1"`, `"Q2"`, `"Q3"`, `"Q4"`, `"annual"`, `"quarter"`, `"ttm" (default: `''`).

**year**: Specific year to filter (e.g., `2025`, default: `''`).

### Examples

**Input**: `=DIVIDENDDATA_STATEMENT("MSFT", "income", TRUE, "annual", "2024")`

**Output**: Income statement table for 2024 with headers.

**Input**: `=DIVIDENDDATA_STATEMENT("AAPL", "cash_flow", FALSE, "ttm")`

**Output**: TTM cash flow statement without headers.

## 4) DIVIDENDDATA_METRICS
### Description
Retrieves a specific metric from financial statements. This function allows targeted extraction of key financial figures like revenue or free cash flow, either as the latest value or historical table.

### Parameters

**symbo**l: Stock ticker symbol (e.g., MSFT). Required.
**metric**: The specific financial metric (e.g., `"revenue"`, `"freeCashFlow"`). 

_All available metrics_: 
* `"Revenue"` (Revenue)
* `"CostOfRevenue"` (Cost of Revenue)
* `"GrossProfit"` (Gross Profit)
* `"ResearchAndDevelopmentExpenses"` (Research and Development Expenses)
* `"GeneralAndAdministrativeExpenses"` (General and Administrative Expenses)
* `"SellingAndMarketingExpenses"` (Selling and Marketing Expenses)
* `"SellingGeneralAndAdministrativeExpenses"` (Selling General and Administrative Expenses)
* `"OtherExpenses"` (Other Expenses)
* `"OperatingExpenses"` (Operating Expenses)
* `"CostAndExpenses"` (Cost and Expenses)
* `"NetInterestIncome"` (Net Interest Income)
* `"InterestIncome"` (Interest Income)
* `"InterestExpense"` (Interest Expense)
* `"DepreciationAndAmortization"` (Depreciation and Amortization)
* `"Ebitda"` (Ebitda)
* `"Ebit"` (Ebit)
* `"NonOperatingIncomeExcludingInterest"` (Non Operating Income Excluding Interest)
* `"OperatingIncome"` (Operating Income)
* `"TotalOtherIncomeExpensesNet"` (Total Other Income Expenses Net)
* `"IncomeBeforeTax"` (Income Before Tax)
* `"IncomeTaxExpense"` (Income Tax Expense)
* `"NetIncomeFromContinuingOperations"` (Net Income From Continuing Operations)
* `"NetIncomeFromDiscontinuedOperations"` (Net Income From Discontinued Operations)
* `"OtherAdjustmentsToNetIncome"` (Other Adjustments To Net Income)
* `"NetIncome"` (Net Income)
* `"NetIncomeDeductions"` (Net Income Deductions)
* `"BottomLineNetIncome"` (Bottom Line Net Income)
* `"Eps"` (Eps)
* `"EpsDiluted"` (Eps Diluted)
* `"WeightedAverageShsOut"` (Weighted Average Shs Out)
* `"WeightedAverageShsOutDil"` (Weighted Average Shs Out Dil)
* `"CashAndCashEquivalents"` (Cash and Cash Equivalents)
* `"ShortTermInvestments"` (Short Term Investments)
* `"CashAndShortTermInvestments"` (Cash and Short Term Investments)
* `"NetReceivables"` (Net Receivables)
* `"AccountsReceivables"` (Accounts Receivables)
* `"OtherReceivables"` (Other Receivables)
* `"Inventory"` (Inventory)
* `"Prepaids"` (Prepaids)
* `"OtherCurrentAssets"` (Other Current Assets)
* `"TotalCurrentAssets"` (Total Current Assets)
* `"PropertyPlantEquipmentNet"` (Property Plant Equipment Net)
* `"Goodwill"` (Goodwill)
* `"IntangibleAssets"` (Intangible Assets)
* `"GoodwillAndIntangibleAssets"` (Goodwill and Intangible Assets)
* `"LongTermInvestments"` (Long Term Investments)
* `"TaxAssets"` (Tax Assets)
* `"OtherNonCurrentAssets"` (Other Non Current Assets)
* `"TotalNonCurrentAssets"` (Total Non Current Assets)
* `"OtherAssets"` (Other Assets)
* `"TotalAssets"` (Total Assets)
* `"TotalPayables"` (Total Payables)
* `"AccountPayables"` (Account Payables)
* `"OtherPayables"` (Other Payables)
* `"AccruedExpenses"` (Accrued Expenses)
* `"ShortTermDebt"` (Short Term Debt)
* `"CapitalLeaseObligationsCurrent"` (Capital Lease Obligations Current)
* `"TaxPayables"` (Tax Payables)
* `"DeferredRevenue"` (Deferred Revenue)
* `"OtherCurrentLiabilities"` (Other Current Liabilities)
* `"TotalCurrentLiabilities"` (Total Current Liabilities)
* `"LongTermDebt"` (Long Term Debt)
* `"CapitalLeaseObligationsNonCurrent"` (Capital Lease Obligations Non Current)
* `"DeferredRevenueNonCurrent"` (Deferred Revenue Non Current)
* `"DeferredTaxLiabilitiesNonCurrent"` (Deferred Tax Liabilities Non Current)
* `"OtherNonCurrentLiabilities"` (Other Non Current Liabilities)
* `"TotalNonCurrentLiabilities"` (Total Non Current Liabilities)
* `"OtherLiabilities"` (Other Liabilities)
* `"CapitalLeaseObligations"` (Capital Lease Obligations)
* `"TotalLiabilities"` (Total Liabilities)
* `"TreasuryStock"` (Treasury Stock)
* `"PreferredStock"` (Preferred Stock)
* `"CommonStock"` (Common Stock)
* `"RetainedEarnings"` (Retained Earnings)
* `"AdditionalPaidInCapital"` (Additional Paid In Capital)
* `"AccumulatedOtherComprehensiveIncomeLoss"` (Accumulated Other Comprehensive Income Loss)
* `"OtherTotalStockholdersEquity"` (Other Total Stockholders Equity)
* `"TotalStockholdersEquity"` (Total Stockholders Equity)
* `"TotalEquity"` (Total Equity)
* `"MinorityInterest"` (Minority Interest)
* `"TotalLiabilitiesAndTotalEquity"` (Total Liabilities and Total Equity)
* `"TotalInvestments"` (Total Investments)
* `"TotalDebt"` (Total Debt)
* `"NetDebt"` (Net Debt)
* `"NetIncome"` (Net Income)
* `"DepreciationAndAmortization"` (Depreciation and Amortization)
* `"DeferredIncomeTax"` (Deferred Income Tax)
* `"StockBasedCompensation"` (Stock Based Compensation)
* `"ChangeInWorkingCapital"` (Change In Working Capital)
* `"AccountsReceivables"` (Accounts Receivables)
* `"Inventory"` (Inventory)
* `"AccountsPayables"` (Accounts Payables)
* `"OtherWorkingCapital"` (Other Working Capital)
* `"OtherNonCashItems"` (Other Non Cash Items)
* `"NetCashProvidedByOperatingActivities"` (Net Cash Provided By Operating Activities)
* `"InvestmentsInPropertyPlantAndEquipment"` (Investments In Property Plant and Equipment)
* `"AcquisitionsNet"` (Acquisitions Net)
* `"PurchasesOfInvestments"` (Purchases of Investments)
* `"SalesMaturitiesOfInvestments"` (Sales Maturities of Investments)
* `"OtherInvestingActivities"` (Other Investing Activities)
* `"NetCashProvidedByInvestingActivities"` (Net Cash Provided By Investing Activities)
* `"NetDebtIssuance"` (Net Debt Issuance)
* `"LongTermNetDebtIssuance"` (Long Term Net Debt Issuance)
* `"ShortTermNetDebtIssuance"` (Short Term Net Debt Issuance)
* `"NetStockIssuance"` (Net Stock Issuance)
* `"NetCommonStockIssuance"` (Net Common Stock Issuance)
* `"CommonStockIssuance"` (Common Stock Issuance)
* `"CommonStockRepurchased"` (Common Stock Repurchased)
* `"NetPreferredStockIssuance"` (Net Preferred Stock Issuance)
* `"NetDividendsPaid"` (Net Dividends Paid)
* `"CommonDividendsPaid"` (Common Dividends Paid)
* `"PreferredDividendsPaid"` (Preferred Dividends Paid)
* `"OtherFinancingActivities"` (Other Financing Activities)
* `"NetCashProvidedByFinancingActivities"` (Net Cash Provided By Financing Activities)
* `"EffectOfForexChangesOnCash"` (Effect of Forex Changes On Cash)
* `"NetChangeInCash"` (Net Change In Cash)
* `"CashAtEndOfPeriod"` (Cash At End of Period)
* `"CashAtBeginningOfPeriod"` (Cash At Beginning of Period)
* `"OperatingCashFlow"` (Operating Cash Flow)
* `"CapitalExpenditure"` (Capital Expenditure)
* `"FreeCashFlow"` (Free Cash Flow)
* `"IncomeTaxesPaid"` (Income Taxes Paid)
* `"InterestPaid"` (Interest Paid)

**showHeaders**: If `true`, returns historical table; else, latest value (default: `false`).
**period**: Period: `"annual"`, `"quarter"`, `"ttm"` (default: `''`).
**year**: Specific year filter (default: `''`).

### Examples

**Input**: `=DIVIDENDDATA_METRICS("MSFT", "revenue")`

**Output**: Latest revenue value for Microsoft stock. Annual by default.

**Input**: `=DIVIDENDDATA_METRICS("AAPL", "freeCashFlow", TRUE, "quarter", "2024")`

**Output**: Quarterly free cash flow history table for 2024 of Apple stock.

## 5) DIVIDENDDATA_RATIOS
### Description
Retrieves a specific financial ratio or key metric for a stock. Useful for valuation, liquidity, and efficiency analysis, either latest value or historical.

### Parameters

**symbol**: Stock ticker symbol (e.g., MSFT). Required.

**metric**: The ratio or key metric. (e.g., currentRatio, peRatio). Available: currentRatio, peRatio, payoutRatio, roic, debtToEquity, etc.

_All available metrics:_
* `"RevenuePerShare"` (Revenue Per Share)
* `"NetIncomePerShare"` (Net Income Per Share)
* `"OperatingCashFlowPerShare"` (Operating Cash Flow Per Share)
* `"FreeCashFlowPerShare"` (Free Cash Flow Per Share)
* `"CashPerShare"` (Cash Per Share)
* `"BookValuePerShare"` (Book Value Per Share)
* `"TangibleBookValuePerShare"` (Tangible Book Value Per Share)
* `"ShareholdersEquityPerShare"` (Shareholders Equity Per Share)
* `"InterestDebtPerShare"` (Interest Debt Per Share)
* `"MarketCap"` (Market Cap)
* `"EnterpriseValue"` (Enterprise Value)
* `"PeRatio"` (Pe Ratio)
* `"PriceToSalesRatio"` (Price To Sales Ratio)
* `"Pocfratio"` (Pocfratio)
* `"PfcfRatio"` (Pfcf Ratio)
* `"PbRatio"` (Pb Ratio)
* `"PtbRatio"` (Ptb Ratio)
* `"EvToSales"` (Ev To Sales)
* `"EnterpriseValueOverEBITDA"` (Enterprise Value Over EBITDA)
* `"EvToOperatingCashFlow"` (Ev To Operating Cash Flow)
* `"EvToFreeCashFlow"` (Ev To Free Cash Flow)
* `"EarningsYield"` (Earnings Yield)
* `"FreeCashFlowYield"` (Free Cash Flow Yield)
* `"DebtToEquity"` (Debt To Equity)
* `"DebtToAssets"` (Debt To Assets)
* `"NetDebtToEBITDA"` (Net Debt To EBITDA)
* `"CurrentRatio"` (Current Ratio)
* `"InterestCoverage"` (Interest Coverage)
* `"IncomeQuality"` (Income Quality)
* `"DividendYield"` (Dividend Yield)
* `"PayoutRatio"` (Payout Ratio)
* `"SalesGeneralAndAdministrativeToRevenue"` (Sales General And Administrative To Revenue)
* `"ResearchAndDevelopmentToRevenue"` (Research And Development To Revenue)
* `"IntangiblesToTotalAssets"` (Intangibles To Total Assets)
* `"CapexToOperatingCashFlow"` (Capex To Operating Cash Flow)
* `"CapexToRevenue"` (Capex To Revenue)
* `"CapexToDepreciation"` (Capex To Depreciation)
* `"StockBasedCompensationToRevenue"` (Stock Based Compensation To Revenue)
* `"GrahamNumber"` (Graham Number)
* `"Roic"` (Roic)
* `"ReturnOnTangibleAssets"` (Return On Tangible Assets)
* `"GrahamNetNet"` (Graham Net Net)
* `"WorkingCapital"` (Working Capital)
* `"TangibleAssetValue"` (Tangible Asset Value)
* `"NetCurrentAssetValue"` (Net Current Asset Value)
* `"InvestedCapital"` (Invested Capital)
* `"AverageReceivables"` (Average Receivables)
* `"AveragePayables"` (Average Payables)
* `"AverageInventory"` (Average Inventory)
* `"DaysSalesOutstanding"` (Days Sales Outstanding)
* `"DaysPayablesOutstanding"` (Days Payables Outstanding)
* `"DaysOfInventoryOnHand"` (Days Of Inventory On Hand)
* `"ReceivablesTurnover"` (Receivables Turnover)
* `"PayablesTurnover"` (Payables Turnover)
* `"InventoryTurnover"` (Inventory Turnover)
* `"Roe"` (Roe)
* `"CapexPerShare"` (Capex Per Share)
* `"CurrentRatio"` (Current Ratio)
* `"QuickRatio"` (Quick Ratio)
* `"CashRatio"` (Cash Ratio)
* `"DaysOfSalesOutstanding"` (Days Of Sales Outstanding)
* `"DaysOfInventoryOutstanding"` (Days Of Inventory Outstanding)
* `"OperatingCycle"` (Operating Cycle)
* `"DaysOfPayablesOutstanding"` (Days Of Payables Outstanding)
* `"CashConversionCycle"` (Cash Conversion Cycle)
* `"GrossProfitMargin"` (Gross Profit Margin)
* `"OperatingProfitMargin"` (Operating Profit Margin)
* `"PretaxProfitMargin"` (Pretax Profit Margin)
* `"NetProfitMargin"` (Net Profit Margin)
* `"EffectiveTaxRate"` (Effective Tax Rate)
* `"ReturnOnAssets"` (Return On Assets)
* `"ReturnOnEquity"` (Return On Equity)
* `"ReturnOnCapitalEmployed"` (Return On Capital Employed)
* `"NetIncomePerEBT"` (Net Income Per EBT)
* `"EbtPerEbit"` (Ebt Per Ebit)
* `"EbitPerRevenue"` (Ebit Per Revenue)
* `"DebtRatio"` (Debt Ratio)
* `"DebtEquityRatio"` (Debt Equity Ratio)
* `"LongTermDebtToCapitalization"` (Long Term Debt To Capitalization)
* `"TotalDebtToCapitalization"` (Total Debt To Capitalization)
* `"InterestCoverage"` (Interest Coverage)
* `"CashFlowCoverageRatios"` (Cash Flow Coverage Ratios)
* `"ShortTermCoverageRatios"` (Short Term Coverage Ratios)
* `"CapitalExpenditureCoverageRatio"` (Capital Expenditure Coverage Ratio)
* `"DividendPaidAndCapexCoverageRatio"` (Dividend Paid And Capex Coverage Ratio)
* `"DividendPayoutRatio"` (Dividend Payout Ratio)
* `"PriceBookValueRatio"` (Price Book Value Ratio)
* `"PriceToBookRatio"` (Price To Book Ratio)
* `"PriceToSalesRatio"` (Price To Sales Ratio)
* `"PriceEarningsRatio"` (Price Earnings Ratio)
* `"PriceToFreeCashFlowsRatio"` (Price To Free Cash Flows Ratio)
* `"PriceToOperatingCashFlowsRatio"` (Price To Operating Cash Flows Ratio)
* `"PriceCashFlowRatio"` (Price Cash Flow Ratio)
* `"PriceEarningsToGrowthRatio"` (Price Earnings To Growth Ratio)
* `"PriceSalesRatio"` (Price Sales Ratio)
* `"DividendYield"` (Dividend Yield)
* `"EnterpriseValueMultiple"` (Enterprise Value Multiple)
* `"PriceFairValue"` (Price Fair Value)

**showHeaders**: If `true`, returns historical table (default: `false`).

**period**: Period: `"annual"`, `"quarter"`, `"ttm"` (default: `''`).

**year**: Specific year filter like `2023` (default: `''`).

### Examples

**Input**: `=DIVIDENDDATA_RATIOS("MSFT", "peRatio")`

**Output**: Latest P/E ratio for Microsoft.

**Input**: `=DIVIDENDDATA_RATIOS("AAPL", "currentRatio", TRUE, "annual")`

**Output**: Annual current ratio history table for Apple.

## 6) DIVIDENDDATA_GROWTH
Description
Retrieves a specific growth metric for financial figures. Useful for trend analysis, like revenue or EPS growth rates over time.
Parameters

symbol: Stock ticker symbol (e.g., MSFT). Required.
metric: The growth metric (e.g., revenueGrowth, epsGrowth). Available: revenueGrowth, epsGrowth, dividendsPerShareGrowth, etc. (see code for full list).
showHeaders: If true, returns historical table (default: false).
period: Period: annual, quarter, q1, q2, q3, q4, fy (default: '').
year: Specific year filter (default: '').

Examples

Input: =DIVIDENDDATA_GROWTH("MSFT", "revenueGrowth")
Output: Latest revenue growth rate.
Input: =DIVIDENDDATA_GROWTH("AAPL", "epsGrowth", TRUE, "quarter")
Output: Quarterly EPS growth history table.

## DIVIDENDDATA_QUOTE
Description
Retrieves stock quote data, including current price, changes, or historical prices. Useful for real-time monitoring or historical analysis of stock performance.
Parameters

symbol: Stock ticker symbol (e.g., AAPL). Required.
metric: The quote metric: price, change, volume, full, history (default: "price").
fromDate: Start date for history (YYYY-MM-DD, for history only).
toDate: End date for history (YYYY-MM-DD, for history only).
showHeaders: Include headers for full or history (default: true).

Examples

Input: =DIVIDENDDATA_QUOTE("AAPL", "price")
Output: Current price.
Input: =DIVIDENDDATA_QUOTE("MSFT", "history", "2024-01-01", "2024-12-31", TRUE)
Output: Historical price table with headers.


## DIVIDENDDATA_QUOTE_BATCH
Description
Retrieves batch quote data for multiple stocks. Efficient for monitoring prices or volumes across tickers.
Parameters

symbols: Comma-separated tickers (e.g., AAPL,MSFT). Required.
metrics: Comma-separated metrics or "all": price, change, volume (default: "all").
showHeaders: Include headers and symbol column (defaults based on metrics).

Examples

Input: =DIVIDENDDATA_QUOTE_BATCH("AAPL,MSFT", "price,change")
Output: Table with prices and changes.
Input: =DIVIDENDDATA_QUOTE_BATCH("MSFT,KMB", "all", TRUE)
Output: Full quotes table with headers.


## DIVIDENDDATA_PROFILE
Description
Retrieves company profile information. Useful for overview details like market cap, sector, or description.
Parameters

symbol: Stock ticker symbol (e.g., MSFT). Required.
metric: Specific profile metric or "full". Available: marketcap, beta, lastdividend, companyname, sector, etc. (see code for full list).
showHeaders: Include headers for "full" (default: false).

Examples

Input: =DIVIDENDDATA_PROFILE("MSFT", "marketcap")
Output: Market capitalization.
Input: =DIVIDENDDATA_PROFILE("AAPL", "full", TRUE)
Output: Full profile table with headers.

## DIVIDENDDATA_FUND
Description
Retrieves data for ETFs or mutual funds. Useful for fund analysis, including holdings, expense ratios, or sector exposures.
Parameters

symbol: Fund ticker symbol (e.g., SPY). Required.
metric: Fund metric: holdings, countryweighting, symbol, name, description, isin, assetclass, securitycusip, domicile, website, etfcompany, expenseratio, assetsundermanagement, avgvolume, inceptiondate, nav, navcurrency, holdingscount, updatedat, sectorslist.
showHeaders: Include headers for tables (default: false).

Examples

Input: =DIVIDENDDATA_FUND("SPY", "expenseratio")
Output: Expense ratio.
Input: =DIVIDENDDATA_FUND("SPY", "holdings", TRUE)
Output: Holdings table with headers.

## DIVIDENDDATA_SEGMENTS
Description
Retrieves revenue segmentation data by product or geography. Useful for understanding revenue sources and diversification.
Parameters

symbol: Stock ticker symbol (e.g., AAPL). Required.
metric: Segmentation type: products or geographic. Required.
period: Period: annual or quarter (default: annual).
year: Specific year filter (default: '').
showHeaders: Include header row (default: false).

Examples

Input: =DIVIDENDDATA_SEGMENTS("AAPL", "products", "annual", "2024", TRUE)
Output: Product segments table for 2024 with headers.
Input: =DIVIDENDDATA_SEGMENTS("MSFT", "geographic", "quarter")
Output: Quarterly geographic segments table.

## DIVIDENDDATA_KPIS
Description
Retrieves key performance indicators (KPIs) and segments from Fiscal.ai. Useful for advanced metrics like customer acquisition cost or churn rates.
Parameters

ticker: Stock ticker symbol (e.g., MSFT). Required.
period: Period: annual or quarterly (default: annual).
year: Optional year filter (default: '').
showheaders: Include header row (default: TRUE).

Examples

Input: =DIVIDENDDATA_KPIS("MSFT", "annual", "2024", TRUE)
Output: Annual KPIs table for 2024 with headers.
Input: =DIVIDENDDATA_KPIS("AAPL", "quarterly")
Output: Quarterly KPIs table without headers.

## DIVIDENDDATA_COMMODITIES
Description
Retrieves commodities data, such as prices or history. Useful for tracking commodity markets like oil or gold.
Parameters

symbol: Commodity symbol (e.g., CLUSD). Required except for "list".
metric: Metric: list, price, fullquote, history. Required.
fromDate: Start date for history.
toDate: End date for history.
showHeaders: Include headers for tables.

Examples

Input: =DIVIDENDDATA_COMMODITIES("CLUSD", "price")
Output: Current oil price.
Input: =DIVIDENDDATA_COMMODITIES(, "list", , , TRUE)
Output: Commodities list table with headers.


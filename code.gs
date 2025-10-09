// Set API Key for Fmp

function setApiKey() {
  // Run this ONCE to set your key (replace with your real key).
  // In production, prompt users or use a settings sidebar.
  const properties = PropertiesService.getUserProperties();
  properties.setProperty('FMP_API_KEY', 'f5b9d1bd0535959d2309799bf894cfaa');  // Your key here
  console.log('API key set!');
}

// Helper to get the key
function getApiKey() {
  return PropertiesService.getUserProperties().getProperty('FMP_API_KEY');
}


// Helper function for cached API fetch
function fetchWithCache(url, ttl = 300) {
  const cache = CacheService.getScriptCache();
  let content = cache.get(url);
  if (content) {
    return content;
  }
  const response = UrlFetchApp.fetch(url);
  content = response.getContentText();
  if (content.length < 90000) {
    cache.put(url, content, ttl);
  }
  return content;
}

// Helper function to calculate dividend CAGR for a given period
function getDividendCAGR(symbol, years) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('API key not set. Please run setApiKey() to configure.');

  const url = `https://financialmodelingprep.com/stable/dividends?symbol=${symbol}&apikey=${apiKey}`;
  const content = fetchWithCache(url, 3600);  // 1 hour for dividends
  const data = JSON.parse(content);

  if (data.length === 0) return 0;

  // Data is sorted latest first (descending paymentDate)
  const latest = data[0];
  const latestDate = new Date(latest.paymentDate);

  let daysBack;
  switch (years) {
    case 1: daysBack = 335; break;  // Improved from 335 to 365 for consistency
    case 3: daysBack = 1095; break;
    case 5: daysBack = 1825; break;
    case 10: daysBack = 3650; break;
    default: throw new Error('Invalid years parameter for CAGR. Use 1, 3, 5, or 10.');
  }

  const pastThreshold = new Date(latestDate.getTime() - daysBack * 24 * 60 * 60 * 1000);

  // Find the newest entry with paymentDate <= pastThreshold
  let pastEntry = null;
  for (let entry of data) {
    const entryDate = new Date(entry.paymentDate);
    if (entryDate <= pastThreshold) {
      pastEntry = entry;
      break;  // First match is the newest due to descending sort
    }
  }

  if (!pastEntry) return 0;

  const latestAdj = parseFloat(latest.adjDividend);
  const pastAdj = parseFloat(pastEntry.adjDividend);

  if (isNaN(latestAdj) || isNaN(pastAdj) || pastAdj === 0) return 0;

  const cagr = Math.pow(latestAdj / pastAdj, 1 / years) - 1;
  return isNaN(cagr) ? 0 : cagr;
}


// FUNCTIONS

/**
 * Retrieves dividend data for a stock. Returns values or tables for metrics like payouts, yields, history, growth, and ratios.
 * @param {string} symbol - Stock ticker symbol (e.g., MSFT).
 * @param {string} [metric="fwd_payout"] - Metric to retrieve: fwd_payout, ttm_payout, fwd_yield, ttm_yield, frequency, history, growth, 1y_cagr, 3y_cagr, 5y_cagr, 10y_cagr, payout_ratio, fcf_payout_ratio.
 * @param {boolean} [showHeaders=false] - Include headers for history or growth tables.
 * @return {string|number|array} - Dividend data.
 * @customfunction
 */
function DIVIDENDDATA(symbol, metric = "fwd_payout", showHeaders = false) {
  const apiKey = getApiKey();
  if (!apiKey) return 'API key not set. Please run setApiKey() to configure.';

  try {
    symbol = symbol.toUpperCase();
    let url, content, data;

    switch (metric.toLowerCase()) {
      case 'fwd_payout':
        // Use original stable API - returns table [date, dividend, adjDividend], latest first
        url = `https://financialmodelingprep.com/stable/dividends?symbol=${symbol}&apikey=${apiKey}`;
        content = fetchWithCache(url, 3600);
        data = JSON.parse(content);
        if (data.length === 0) return 0;
        const latestDiv = data[0].dividend;  // Latest first
        const frequency = data[0].frequency || 'quarterly';
        const periodsPerYear = { 'weekly': 52, 'monthly': 12, 'quarterly': 4, 'semiAnnual': 2, 'annual': 1, 'special': 1 }[frequency] || 4;
        const annualDiv = parseFloat(latestDiv) * periodsPerYear;
        return annualDiv;

      case 'ttm_payout':
        url = `https://financialmodelingprep.com/stable/profile?symbol=${symbol}&apikey=${apiKey}`;
        content = fetchWithCache(url, 300);
        data = JSON.parse(content);
        if (data.length === 0) return 0;
        return data[0].lastDividend;

      case 'fwd_yield':
        // Replicate your dividend summary (annual_dividend / price)
        url = `https://financialmodelingprep.com/api/v3/quote/${symbol}?apikey=${apiKey}`;
        content = fetchWithCache(url, 60);  // Short for quotes
        data = JSON.parse(content);

        if (data.length === 0) return 'No quote data available for symbol ' + symbol;
        const price = data[0].price;
        const fwd_dividend = DIVIDENDDATA(symbol, 'fwd_payout');
        if (fwd_dividend.length === 0) return 0;
        return fwd_dividend / price;  // Yield as decimal (format in sheet as %)

      case 'ttm_yield':
        url = `https://financialmodelingprep.com/stable/profile?symbol=${symbol}&apikey=${apiKey}`;
        content = fetchWithCache(url, 300);
        data = JSON.parse(content);
        if (data.length === 0) return 0;
        return data[0].lastDividend / data[0].price;

      case 'frequency':
        url = `https://financialmodelingprep.com/stable/dividends?symbol=${symbol}&apikey=${apiKey}`;
        content = fetchWithCache(url, 3600);
        data = JSON.parse(content);
        if (data.length === 0) return 'No dividend data available for symbol ' + symbol;
        return data[0].frequency;

      case 'history':
        // Use original stable API - returns table [date, dividend, adjDividend], latest first
        url = `https://financialmodelingprep.com/stable/dividends?symbol=${symbol}&apikey=${apiKey}`;
        content = fetchWithCache(url, 3600);
        data = JSON.parse(content);

        if (data.length === 0) return [['No dividend history available for symbol ' + symbol]];

        // Clean and map to 2D array (fix for Sheets spilling)
        let cleaned = data.map(row => [
          row.declarationDate ? new Date(row.declarationDate) : '',  // Declaration Date (handle empty)
          row.recordDate ? new Date(row.recordDate) : '',  // Record Date / Ex-Dividend Date
          row.paymentDate ? new Date(row.paymentDate) : '',  // Payment Date
          parseFloat(row.adjDividend.toString().replace(/^0+/, '')) || 0,  // Adjusted Dividend
          parseFloat(row.dividend.toString().replace(/^0+/, '')) || 0,  // Dividend
          parseFloat(row.yield.toString().replace(/^0+/, '')) / 100 || 0,  // Yield
          row.frequency || ''  // Frequency
        ]);

        // Optional headers
        if (showHeaders) {
          cleaned.unshift(['Declaration Date', 'Record Date', 'Payment Date', 'Adjusted Dividend', 'Dividend', 'Yield', 'Frequency']);
        }

        return cleaned;  // Latest first (no reverse)

      case 'growth':
        // Replicate your dividend_growth_calc: % change YoY from history
        const histData = DIVIDENDDATA(symbol, 'history');  // Get without headers
        if (histData.length < 2) return [['Insufficient dividend history for growth calculation for symbol ' + symbol]];

        // Data is latest first (descending dates)
        // For growth, calculate (current - previous) / abs(previous), where previous is older (next row)
        const growth = [];
        for (let i = 0; i < histData.length - 1; i++) {
          const currentAdj = histData[i][4];  // Adjusted Dividend index
          const previousAdj = histData[i + 1][4];  // Older
          const rate = previousAdj !== 0 ? (currentAdj - previousAdj) / Math.abs(previousAdj) : 0;
          growth.push([histData[i][1], rate]);  // Use Ex-Div Date (index 1), raw rate
        }
        growth.push([histData[histData.length - 1][1], 0]);  // Last has 0

        // Optional headers
        if (showHeaders) {
          growth.unshift(['Date', 'Growth Rate']);
        }

        return growth;

      case '1y_cagr':
        return getDividendCAGR(symbol, 1);

      case '3y_cagr':
        return getDividendCAGR(symbol, 3);

      case '5y_cagr':
        return getDividendCAGR(symbol, 5);

      case '10y_cagr':
        return getDividendCAGR(symbol, 10);

      case 'payout_ratio':
        // EPS payout ratio from stable ratios endpoint
        url = `https://financialmodelingprep.com/stable/ratios?symbol=${symbol}&apikey=${apiKey}`;
        content = fetchWithCache(url, 3600);
        data = JSON.parse(content);

        if (data.length === 0) return 'No ratio data available for symbol ' + symbol;
        const payoutRatio = data[0].dividendPayoutRatio || 0;
        return payoutRatio;  // Raw number (user formats as % in Sheets)

      case 'fcf_payout_ratio':
        // FCF payout ratio from stable ratios endpoint: (dividendPerShare / freeCashFlowPerShare) * 100
        url = `https://financialmodelingprep.com/stable/ratios?symbol=${symbol}&apikey=${apiKey}`;
        content = fetchWithCache(url, 3600);
        data = JSON.parse(content);

        if (data.length === 0) return 'No ratio data available for symbol ' + symbol;
        const divPerShare = data[0].dividendPerShare || 0;
        const fcfPerShare = data[0].freeCashFlowPerShare || 0;
        return fcfPerShare > 0 ? (divPerShare / fcfPerShare) : 0;  // Raw number or 0

      default:
        return 'Invalid metric: ' + metric + '. Valid metrics are: fwd_payout, ttm_payout, fwd_yield, ttm_yield, frequency, history, growth, 1y_cagr, 3y_cagr, 5y_cagr, 10y_cagr, payout_ratio, fcf_payout_ratio.';
    }
  } catch (error) {
    return 'An error occurred while fetching data for symbol ' + symbol + ' and metric ' + metric + '. Please check the symbol and parameters.';
  }
}


/**
 * Retrieves batch dividend data for multiple stocks. Returns table with latest or historical data for metrics like payouts or yields.
 * @param {string} symbols - Comma-separated tickers (e.g., MSFT,KMB,O).
 * @param {string} [metric="fwd_payout"] - Metrics: adjdividend, dividend, recorddate, paymentdate, declarationdate, yield, frequency, fwd_payout, "all", or "history".
 * @param {boolean} [showHeaders] - Include headers (defaults based on metric).
 * @return {array} - Dividend table.
 * @customfunction
 */
function DIVIDENDDATA_BATCH(symbols, metric = "fwd_payout", showHeaders) {
  const apiKey = getApiKey();
  if (!apiKey) return 'API key not set. Please run setApiKey() to configure.';

  try {
    if (!symbols) return 'Symbols parameter is required.';

    // Process symbols
    const symbolArray = symbols.split(',').map(s => s.trim().toUpperCase());
    const symbolList = symbolArray.join(',');

    const url = `https://financialmodelingprep.com/api/v4/private/dividend_data/dividend-yield?symbol=${symbolList}&apikey=${apiKey}`;
    const content = fetchWithCache(url, 3600);
    const data = JSON.parse(content);

    if (data.length === 0) return [['No dividend data available for the provided symbols']];

    // Group by symbol and sort each group by date desc
    const dataMap = new Map();
    symbolArray.forEach(sym => dataMap.set(sym, []));
    data.forEach(item => {
      const sym = item.symbol.toUpperCase();
      if (dataMap.has(sym)) {
        dataMap.get(sym).push(item);
      }
    });
    dataMap.forEach((group, sym) => {
      group.sort((a, b) => new Date(b.date) - new Date(a.date));
    });

    // Metric to API key mapping (API uses camelCase)
    const metricToApiKey = {
      "date": "date",
      "adjdividend": "adjDividend",
      "dividend": "dividend",
      "recorddate": "recordDate",
      "paymentdate": "paymentDate",
      "declarationdate": "declarationDate",
      "yield": "yield",
      "frequency": "frequency"
    };

    const allMetrics = Object.keys(metricToApiKey);
    const numericMetrics = ["adjdividend", "dividend", "yield"];
    const dateMetrics = ["date", "recorddate", "paymentdate", "declarationdate"];

    // Process metric
    const loweredMetric = metric.toLowerCase();
    let requested;
    const isHistory = loweredMetric === "history";
    if (isHistory) {
      // Special handling below
    } else if (loweredMetric === "all") {
      requested = allMetrics;
    } else {
      requested = loweredMetric.split(',').map(m => m.trim());
      for (let m of requested) {
        if (m !== "fwd_payout" && !allMetrics.includes(m)) {
          return 'Invalid metric: ' + m + '. Valid metrics are: adjdividend, dividend, recorddate, paymentdate, declarationdate, yield, frequency, fwd_payout, "all", or "history".';
        }
      }
    }

    // Determine effective showHeaders
    if (showHeaders === undefined) {
      showHeaders = (loweredMetric === "all" || isHistory);
    }

    // Function to format headers
    function formatHeader(key) {
      return key
        .replace(/([a-z])([A-Z])/g, '$1 $2')
        .replace(/_/g, ' ')
        .replace(/\b\w/g, char => char.toUpperCase());
    }

    if (isHistory) {
      // For history: flat table with all rows, sorted by date desc overall
      let allRows = [];
      data.forEach(item => allRows.push(item));
      allRows.sort((a, b) => new Date(b.date) - new Date(a.date));

      const fields = ["symbol", ...allMetrics];
      const formattedHeaders = fields.map(f => formatHeader(f));

      let table = allRows.map(item => fields.map(field => {
        const loweredField = field.toLowerCase();
        let value;
        if (loweredField === "symbol") {
          value = item.symbol;
        } else {
          value = item[metricToApiKey[loweredField]];
        }
        if (dateMetrics.includes(loweredField)) {
          return value ? new Date(value) : '';
        }
        return value ?? (numericMetrics.includes(loweredField) ? 0 : '');
      }));

      if (showHeaders) {
        table.unshift(formattedHeaders);
      }

      return table;
    } else {
      // For non-history: latest per symbol
      const fields = showHeaders ? ["symbol", ...requested] : requested;
      const formattedHeaders = fields.map(f => f === "fwd_payout" ? "Fwd Payout" : formatHeader(f));

      let table = symbolArray.map(sym => {
        const group = dataMap.get(sym);
        const latest = group && group.length > 0 ? group[0] : {};
        return fields.map(field => {
          const loweredField = field.toLowerCase();
          if (loweredField === "symbol") return sym;
          if (loweredField === "fwd_payout") {
            const div = latest[metricToApiKey["dividend"]] || 0;
            const freq = (latest[metricToApiKey["frequency"]] || 'quarterly').toLowerCase();
            const periods = { 'weekly': 52, 'monthly': 12, 'quarterly': 4, 'semiannual': 2, 'annual': 1, 'special': 1 }[freq] || 4;
            return div * periods;
          }
          let value = latest[metricToApiKey[loweredField]];
          if (dateMetrics.includes(loweredField)) {
            return value ? new Date(value) : '';
          }
          return value ?? (numericMetrics.includes(loweredField) ? 0 : '');
        });
      });

      if (showHeaders) {
        table.unshift(formattedHeaders);
      }

      return table;
    }

  } catch (error) {
    return 'An error occurred while fetching batch dividend data for symbols ' + symbols + ' and metric ' + metric + '. Please check the symbols and parameters.';
  }
}


/**
 * Retrieves full financial statements for a stock. Returns table with income, balance, or cash flow data, filtered by period and year.
 * @param {string} symbol - Stock ticker symbol (e.g., MSFT).
 * @param {string} metric - Statement type: income, balance, cash_flow.
 * @param {boolean} [showHeaders=false] - Include headers.
 * @param {string} [period=''] - Period: FY, Q1, Q2, Q3, Q4, annual, quarter, ttm.
 * @param {string} [year=''] - Year filter (e.g., 2025).
 * @return {array} - Statement table.
 * @customfunction
 */
function DIVIDENDDATA_STATEMENT(symbol, metric, showHeaders = false, period = '', year = '') {
  const apiKey = getApiKey();
  if (!apiKey) return 'API key not set. Please run setApiKey() to configure.';

  try {
    symbol = symbol.toUpperCase();
    let endpoint;
    switch (metric.toLowerCase()) {
      case 'income':
        endpoint = 'income-statement';
        break;
      case 'balance':
        endpoint = 'balance-sheet-statement';
        break;
      case 'cash_flow':
        endpoint = 'cash-flow-statement';
        break;
      default:
        return 'Invalid metric: ' + metric + '. Valid metrics are: income, balance, cash_flow.';
    }

    let url;
    const loweredPeriod = period.toLowerCase();
    if (loweredPeriod === 'ttm') {
      // Special handling for TTM: use -ttm endpoint, no period param, ignore year for fetch but filter later if provided
      url = `https://financialmodelingprep.com/stable/${endpoint}-ttm?symbol=${symbol}&limit=500&apikey=${apiKey}`;
    } else {
      url = `https://financialmodelingprep.com/stable/${endpoint}?symbol=${symbol}&limit=500`;
      if (period) url += `&period=${period}`;
      url += `&apikey=${apiKey}`;
    }

    const content = fetchWithCache(url, 3600);
    let data = JSON.parse(content);

    if (data.length === 0) return [['No statement data available for symbol ' + symbol + ' and metric ' + metric]];

    // Filter by year if provided (works for both regular and TTM)
    if (year) {
      data = data.filter(row => row.fiscalYear === year);
      if (data.length === 0) return [['No data available for the specified year ' + year]];
    }

    // Sort by date descending (latest first)
    data.sort((a, b) => new Date(b.date) - new Date(a.date));

    // Get raw headers dynamically from first object
    const rawHeaders = Object.keys(data[0]);

    // Function to format headers to human-readable (camelCase/snake_case to Title Case with spaces)
    function formatHeader(key) {
      return key
        .replace(/([a-z])([A-Z])/g, '$1 $2')  // camelCase to space-separated
        .replace(/_/g, ' ')  // snake_case to space
        .replace(/\b\w/g, char => char.toUpperCase());  // Capitalize words
    }

    const formattedHeaders = rawHeaders.map(formatHeader);

    // Map to 2D array
    let table = data.map(row => rawHeaders.map(key => {
      const value = row[key];
      // Handle dates and numbers
      if (key.toLowerCase().includes('date')) return new Date(value);
      if (typeof value === 'number') return value;
      return value || '';  // Handle null/undefined
    }));

    // Optional formatted headers
    if (showHeaders) {
      table.unshift(formattedHeaders);
    }

    return table;

  } catch (error) {
    return 'An error occurred while fetching statement data for symbol ' + symbol + ', metric ' + metric + ', period ' + period + ', year ' + year + '. Please check the parameters.';
  }

}


/**
 * Retrieves specific metric from statements. Returns latest value or history table for figures like revenue or freeCashFlow.
 * @param {string} symbol - Stock ticker symbol (e.g., MSFT).
 * @param {string} metric - Metric: revenue, netIncome, freeCashFlow, eps, totalAssets, totalDebt, etc. (see code for full list).
 * @param {boolean} [showHeaders=false] - Return history table if true.
 * @param {string} [period=''] - Period: annual, quarter, ttm.
 * @param {string} [year=''] - Year filter.
 * @return {number|array} - Metric value or table.
 * @customfunction
 */
function DIVIDENDDATA_METRICS(symbol, metric, showHeaders = false, period = '', year = '') {
  const apiKey = getApiKey();
  if (!apiKey) return 'API key not set. Please run setApiKey() to configure.';

  const metricToStatement = {
  // Income statement metrics
  'revenue': 'income', 'costofrevenue': 'income', 'grossprofit': 'income', 'researchanddevelopmentexpenses': 'income', 'generalandadministrativeexpenses': 'income',
  'sellingandmarketingexpenses': 'income', 'sellinggeneralandadministrativeexpenses': 'income', 'otherexpenses': 'income', 'operatingexpenses': 'income',
  'costandexpenses': 'income', 'netinterestincome': 'income', 'interestincome': 'income', 'interestexpense': 'income', 'depreciationandamortization': 'income',
  'ebitda': 'income', 'ebit': 'income', 'nonoperatingincomeexcludinginterest': 'income', 'operatingincome': 'income', 'totalotherincomeexpensesnet': 'income',
  'incomebeforetax': 'income', 'incometaxexpense': 'income', 'netincomefromcontinuingoperations': 'income', 'netincomefromdiscontinuedoperations': 'income',
  'otheradjustmentstonetincome': 'income', 'netincome': 'income', 'netincomedeductions': 'income', 'bottomlinenetincome': 'income', 'eps': 'income',
  'epsdiluted': 'income', 'weightedaverageshsout': 'income', 'weightedaverageshsoutdil': 'income',
  
  // Balance sheet metrics
  'cashandcashequivalents': 'balance', 'shortterminvestments': 'balance', 'cashandshortterminvestments': 'balance', 'netreceivables': 'balance',
  'accountsreceivables': 'balance', 'otherreceivables': 'balance', 'inventory': 'balance', 'prepaids': 'balance', 'othercurrentassets': 'balance',
  'totalcurrentassets': 'balance', 'propertyplantequipmentnet': 'balance', 'goodwill': 'balance', 'intangibleassets': 'balance', 'goodwillandintangibleassets': 'balance',
  'longterminvestments': 'balance', 'taxassets': 'balance', 'othernoncurrentassets': 'balance', 'totalnoncurrentassets': 'balance', 'otherassets': 'balance',
  'totalassets': 'balance', 'totalpayables': 'balance', 'accountpayables': 'balance', 'otherpayables': 'balance', 'accruedexpenses': 'balance',
  'shorttermdebt': 'balance', 'capitalleaseobligationscurrent': 'balance', 'taxpayables': 'balance', 'deferredrevenue': 'balance', 'othercurrentliabilities': 'balance',
  'totalcurrentliabilities': 'balance', 'longtermdebt': 'balance', 'capitalleaseobligationsnoncurrent': 'balance', 'deferredrevenuenoncurrent': 'balance',
  'deferredtaxliabilitiesnoncurrent': 'balance', 'othernoncurrentliabilities': 'balance', 'totalnoncurrentliabilities': 'balance', 'otherliabilities': 'balance',
  'capitalleaseobligations': 'balance', 'totalliabilities': 'balance', 'treasurystock': 'balance', 'preferredstock': 'balance', 'commonstock': 'balance',
  'retainedearnings': 'balance', 'additionalpaidincapital': 'balance', 'accumulatedothercomprehensiveincomeloss': 'balance', 'othertotalstockholdersequity': 'balance',
  'totalstockholdersequity': 'balance', 'totalequity': 'balance', 'minorityinterest': 'balance', 'totalliabilitiesandtotalequity': 'balance',
  'totalinvestments': 'balance', 'totaldebt': 'balance', 'netdebt': 'balance',
  
  // Cash flow statement metrics
  'netincome': 'cash_flow', 'depreciationandamortization': 'cash_flow', 'deferredincometax': 'cash_flow', 'stockbasedcompensation': 'cash_flow',
  'changeinworkingcapital': 'cash_flow', 'accountsreceivables': 'cash_flow', 'inventory': 'cash_flow', 'accountspayables': 'cash_flow',
  'otherworkingcapital': 'cash_flow', 'othernoncashitems': 'cash_flow', 'netcashprovidedbyoperatingactivities': 'cash_flow',
  'investmentsinpropertyplantandequipment': 'cash_flow', 'acquisitionsnet': 'cash_flow', 'purchasesofinvestments': 'cash_flow',
  'salesmaturitiesofinvestments': 'cash_flow', 'otherinvestingactivities': 'cash_flow', 'netcashprovidedbyinvestingactivities': 'cash_flow',
  'netdebtissuance': 'cash_flow', 'longtermnetdebtissuance': 'cash_flow', 'shorttermnetdebtissuance': 'cash_flow', 'netstockissuance': 'cash_flow',
  'netcommonstockissuance': 'cash_flow', 'commonstockissuance': 'cash_flow', 'commonstockrepurchased': 'cash_flow', 'netpreferredstockissuance': 'cash_flow',
  'netdividendspaid': 'cash_flow', 'commondividendspaid': 'cash_flow', 'preferreddividendspaid': 'cash_flow', 'otherfinancingactivities': 'cash_flow',
  'netcashprovidedbyfinancingactivities': 'cash_flow', 'effectofforexchangesoncash': 'cash_flow', 'netchangeincash': 'cash_flow',
  'cashatendofperiod': 'cash_flow', 'cashatbeginningofperiod': 'cash_flow', 'operatingcashflow': 'cash_flow', 'capitalexpenditure': 'cash_flow',
  'freecashflow': 'cash_flow', 'incometaxespaid': 'cash_flow', 'interestpaid': 'cash_flow'
  };

  const metricToRawKey = {
  // Income statement metrics
  'revenue': 'revenue', 'costofrevenue': 'costOfRevenue', 'grossprofit': 'grossProfit', 'researchanddevelopmentexpenses': 'researchAndDevelopmentExpenses', 'generalandadministrativeexpenses': 'generalAndAdministrativeExpenses',
  'sellingandmarketingexpenses': 'sellingAndMarketingExpenses', 'sellinggeneralandadministrativeexpenses': 'sellingGeneralAndAdministrativeExpenses', 'otherexpenses': 'otherExpenses', 'operatingexpenses': 'operatingExpenses',
  'costandexpenses': 'costAndExpenses', 'netinterestincome': 'netInterestIncome', 'interestincome': 'interestIncome', 'interestexpense': 'interestExpense', 'depreciationandamortization': 'depreciationAndAmortization',
  'ebitda': 'ebitda', 'ebit': 'ebit', 'nonoperatingincomeexcludinginterest': 'nonOperatingIncomeExcludingInterest', 'operatingincome': 'operatingIncome', 'totalotherincomeexpensesnet': 'totalOtherIncomeExpensesNet',
  'incomebeforetax': 'incomeBeforeTax', 'incometaxexpense': 'incomeTaxExpense', 'netincomefromcontinuingoperations': 'netIncomeFromContinuingOperations', 'netincomefromdiscontinuedoperations': 'netIncomeFromDiscontinuedOperations',
  'otheradjustmentstonetincome': 'otherAdjustmentsToNetIncome', 'netincome': 'netIncome', 'netincomedeductions': 'netIncomeDeductions', 'bottomlinenetincome': 'bottomLineNetIncome', 'eps': 'eps',
  'epsdiluted': 'epsDiluted', 'weightedaverageshsout': 'weightedAverageShsOut', 'weightedaverageshsoutdil': 'weightedAverageShsOutDil',
  
  // Balance sheet metrics
  'cashandcashequivalents': 'cashAndCashEquivalents', 'shortterminvestments': 'shortTermInvestments', 'cashandshortterminvestments': 'cashAndShortTermInvestments', 'netreceivables': 'netReceivables',
  'accountsreceivables': 'accountsReceivables', 'otherreceivables': 'otherReceivables', 'inventory': 'inventory', 'prepaids': 'prepaids', 'othercurrentassets': 'otherCurrentAssets',
  'totalcurrentassets': 'totalCurrentAssets', 'propertyplantequipmentnet': 'propertyPlantEquipmentNet', 'goodwill': 'goodwill', 'intangibleassets': 'intangibleAssets', 'goodwillandintangibleassets': 'goodwillAndIntangibleAssets',
  'longterminvestments': 'longTermInvestments', 'taxassets': 'taxAssets', 'othernoncurrentassets': 'otherNonCurrentAssets', 'totalnoncurrentassets': 'totalNonCurrentAssets', 'otherassets': 'otherAssets',
  'totalassets': 'totalAssets', 'totalpayables': 'totalPayables', 'accountpayables': 'accountPayables', 'otherpayables': 'otherPayables', 'accruedexpenses': 'accruedExpenses',
  'shorttermdebt': 'shortTermDebt', 'capitalleaseobligationscurrent': 'capitalLeaseObligationsCurrent', 'taxpayables': 'taxPayables', 'deferredrevenue': 'deferredRevenue', 'othercurrentliabilities': 'otherCurrentLiabilities',
  'totalcurrentliabilities': 'totalCurrentLiabilities', 'longtermdebt': 'longTermDebt', 'capitalleaseobligationsnoncurrent': 'capitalLeaseObligationsNonCurrent', 'deferredrevenuenoncurrent': 'deferredRevenueNonCurrent',
  'deferredtaxliabilitiesnoncurrent': 'deferredTaxLiabilitiesNonCurrent', 'othernoncurrentliabilities': 'otherNonCurrentLiabilities', 'totalnoncurrentliabilities': 'totalNonCurrentLiabilities', 'otherliabilities': 'otherLiabilities',
  'capitalleaseobligations': 'capitalLeaseObligations', 'totalliabilities': 'totalLiabilities', 'treasurystock': 'treasuryStock', 'preferredstock': 'preferredStock', 'commonstock': 'commonStock',
  'retainedearnings': 'retainedEarnings', 'additionalpaidincapital': 'additionalPaidInCapital', 'accumulatedothercomprehensiveincomeloss': 'accumulatedOtherComprehensiveIncomeLoss', 'othertotalstockholdersequity': 'otherTotalStockholdersEquity',
  'totalstockholdersequity': 'totalStockholdersEquity', 'totalequity': 'totalEquity', 'minorityinterest': 'minorityInterest', 'totalliabilitiesandtotalequity': 'totalLiabilitiesAndTotalEquity',
  'totalinvestments': 'totalInvestments', 'totaldebt': 'totalDebt', 'netdebt': 'netDebt',
  
  // Cash flow statement metrics
  'netincome': 'netIncome', 'depreciationandamortization': 'depreciationAndAmortization', 'deferredincometax': 'deferredIncomeTax', 'stockbasedcompensation': 'stockBasedCompensation',
  'changeinworkingcapital': 'changeInWorkingCapital', 'accountsreceivables': 'accountsReceivables', 'inventory': 'inventory', 'accountspayables': 'accountsPayables',
  'otherworkingcapital': 'otherWorkingCapital', 'othernoncashitems': 'otherNonCashItems', 'netcashprovidedbyoperatingactivities': 'netCashProvidedByOperatingActivities',
  'investmentsinpropertyplantandequipment': 'investmentsInPropertyPlantAndEquipment', 'acquisitionsnet': 'acquisitionsNet', 'purchasesofinvestments': 'purchasesOfInvestments',
  'salesmaturitiesofinvestments': 'salesMaturitiesOfInvestments', 'otherinvestingactivities': 'otherInvestingActivities', 'netcashprovidedbyinvestingactivities': 'netCashProvidedByInvestingActivities',
  'netdebtissuance': 'netDebtIssuance', 'longtermnetdebtissuance': 'longTermNetDebtIssuance', 'shorttermnetdebtissuance': 'shortTermNetDebtIssuance', 'netstockissuance': 'netStockIssuance',
  'netcommonstockissuance': 'netCommonStockIssuance', 'commonstockissuance': 'commonStockIssuance', 'commonstockrepurchased': 'commonStockRepurchased', 'netpreferredstockissuance': 'netPreferredStockIssuance',
  'netdividendspaid': 'netDividendsPaid', 'commondividendspaid': 'commonDividendsPaid', 'preferreddividendspaid': 'preferredDividendsPaid', 'otherfinancingactivities': 'otherFinancingActivities',
  'netcashprovidedbyfinancingactivities': 'netCashProvidedByFinancingActivities', 'effectofforexchangesoncash': 'effectOfForexChangesOnCash', 'netchangeincash': 'netChangeInCash',
  'cashatendofperiod': 'cashAtEndOfPeriod', 'cashatbeginningofperiod': 'cashAtBeginningOfPeriod', 'operatingcashflow': 'operatingCashFlow', 'capitalexpenditure': 'capitalExpenditure',
  'freecashflow': 'freeCashFlow', 'incometaxespaid': 'incomeTaxesPaid', 'interestpaid': 'interestPaid'
  };

  try {
    symbol = symbol.toUpperCase();
    const loweredMetric = metric.toLowerCase();
    const statement = metricToStatement[loweredMetric];
    if (!statement) return 'Invalid metric: ' + metric + '. Please check available metrics in function description.';

    // Function to format headers to human-readable (camelCase/snake_case to Title Case with spaces)
    function formatHeader(key) {
      return key
        .replace(/([a-z])([A-Z])/g, '$1 $2')  // camelCase to space-separated
        .replace(/_/g, ' ')  // snake_case to space
        .replace(/\b\w/g, char => char.toUpperCase());  // Capitalize words
    }

    const properRawKey = metricToRawKey[loweredMetric];
    if (!properRawKey) return 'Invalid metric: ' + metric + '. Please check available metrics in function description.';

    const formattedMetric = formatHeader(properRawKey);

    const fullData = DIVIDENDDATA_STATEMENT(symbol, statement, true, period, year);
    if (typeof fullData === 'string') return fullData; // Error from statement function
    if (fullData.length === 0 || fullData[0].length === 0) return 0;

    const headers = fullData[0];
    const dataRows = fullData.slice(1);
    if (dataRows.length === 0) return 0;

    let colIndex = headers.indexOf(formattedMetric);
    if (colIndex === -1) return 'Metric ' + metric + ' not found in the statement data for symbol ' + symbol;

    if (showHeaders) {
      // Return history table: Date and metric
      let dateCol = headers.indexOf('Date');
      if (dateCol === -1) dateCol = 0;  // Assume first column if not found

      let table = dataRows.map(row => [row[dateCol], row[colIndex]]);
      table.unshift(['Date', formattedMetric]);
      return table;
    } else {
      // Return single latest value
      return dataRows[0][colIndex];
    }
  } catch (error) {
    return 'An error occurred while fetching metric ' + metric + ' for symbol ' + symbol + ', period ' + period + ', year ' + year + '. Please check the parameters.';
  }
}



/**
 * Retrieves financial ratio or key metric. Returns latest value or history for ratios like currentRatio or peRatio.
 * @param {string} symbol - Stock ticker symbol (e.g., MSFT).
 * @param {string} metric - Metric: currentRatio, peRatio, payoutRatio, roic, debtToEquity, etc. (see code for full list).
 * @param {boolean} [showHeaders=false] - Return history table if true.
 * @param {string} [period=''] - Period: annual, quarter, ttm.
 * @param {string} [year=''] - Year filter.
 * @return {number|array} - Ratio value or table.
 * @customfunction
 */
function DIVIDENDDATA_RATIOS(symbol, metric, showHeaders = false, period = '', year = '') {
  const apiKey = getApiKey();
  if (!apiKey) return 'API key not set. Please run setApiKey() to configure.';

  const metricToEndpoint = {
    // Key Metrics
    'revenuepershare': 'key-metrics', 'netincomepershare': 'key-metrics', 'operatingcashflowpershare': 'key-metrics', 'freecashflowpershare': 'key-metrics', 'cashpershare': 'key-metrics','bookvaluepershare': 'key-metrics', 'tangiblebookvaluepershare': 'key-metrics', 'shareholdersequitypershare': 'key-metrics', 'interestdebtpershare': 'key-metrics', 'marketcap': 'key-metrics','enterprisevalue': 'key-metrics', 'peratio': 'key-metrics', 'pricetosalesratio': 'key-metrics', 'pocfratio': 'key-metrics', 'pfcfratio': 'key-metrics', 'pbratio': 'key-metrics', 'ptbratio': 'key-metrics', 'evtosales': 'key-metrics', 'enterprisevalueoverebitda': 'key-metrics', 'evtooperatingcashflow': 'key-metrics', 'evtofreecashflow': 'key-metrics', 'earningsyield': 'key-metrics', 'freecashflowyield': 'key-metrics', 'debttoequity': 'key-metrics', 'debttoassets': 'key-metrics','netdebttoebitda': 'key-metrics', 'currentratio': 'key-metrics', 'interestcoverage': 'key-metrics', 'incomequality': 'key-metrics', 'dividendyield': 'key-metrics','payoutratio': 'key-metrics', 'salesgeneralandadministrativetorevenue': 'key-metrics', 'researchanddevelopementtorevenue': 'key-metrics', 'intangiblestototalassets': 'key-metrics', 'capextooperatingcashflow': 'key-metrics','capextorevenue': 'key-metrics', 'capextodepreciation': 'key-metrics', 'stockbasedcompensationtorevenue': 'key-metrics', 'grahamnumber': 'key-metrics', 'roic': 'key-metrics','returnontangibleassets': 'key-metrics', 'grahamnetnet': 'key-metrics', 'workingcapital': 'key-metrics', 'tangibleassetvalue': 'key-metrics', 'netcurrentassetvalue': 'key-metrics','investedcapital': 'key-metrics', 'averagereceivables': 'key-metrics', 'averagepayables': 'key-metrics', 'averageinventory': 'key-metrics', 'dayssalesoutstanding': 'key-metrics','dayspayablesoutstanding': 'key-metrics', 'daysofinventoryonhand': 'key-metrics', 'receivablesturnover': 'key-metrics', 'payablesturnover': 'key-metrics', 'inventoryturnover': 'key-metrics',
    'roe': 'key-metrics', 'capexpershare': 'key-metrics',

    // Ratios
    'currentratio': 'ratios', 'quickratio': 'ratios', 'cashratio': 'ratios', 'daysofsalesoutstanding': 'ratios', 'daysofinventoryoutstanding': 'ratios',
    'operatingcycle': 'ratios', 'daysofpayablesoutstanding': 'ratios', 'cashconversioncycle': 'ratios', 'grossprofitmargin': 'ratios', 'operatingprofitmargin': 'ratios','pretaxprofitmargin': 'ratios', 'netprofitmargin': 'ratios', 'effectivetaxrate': 'ratios', 'returnonassets': 'ratios', 'returnonequity': 'ratios',
    'returnoncapitalemployed': 'ratios', 'netincomeperebt': 'ratios', 'ebtperebit': 'ratios', 'ebitperrevenue': 'ratios', 'debtratio': 'ratios',
    'debtequityratio': 'ratios', 'longtermdebttocapitalization': 'ratios', 'totaldebttocapitalization': 'ratios', 'interestcoverage': 'ratios', 'cashflowcoverageratios': 'ratios','shorttermcoverageratios': 'ratios', 'capitalexpenditurecoverageratio': 'ratios', 'dividendpaidandcapexcoverageratio': 'ratios', 'dividendpayoutratio': 'ratios', 'pricebookvalueratio': 'ratios','pricetobookratio': 'ratios', 'pricetosalesratio': 'ratios', 'priceearningsratio': 'ratios', 'pricetofreecashflowsratio': 'ratios', 'pricetooperatingcashflowsratio': 'ratios','pricecashflowratio': 'ratios', 'priceearningstogrowthratio': 'ratios', 'pricesalesratio': 'ratios', 'dividendyield': 'ratios', 'enterprisevaluemultiple': 'ratios','pricefairvalue': 'ratios'
  };

  const metricToRawKey = {
    // Key Metrics
    'revenuepershare': 'revenuePerShare',
    'netincomepershare': 'netIncomePerShare',
    'operatingcashflowpershare': 'operatingCashFlowPerShare',
    'freecashflowpershare': 'freeCashFlowPerShare',
    'cashpershare': 'cashPerShare',
    'bookvaluepershare': 'bookValuePerShare',
    'tangiblebookvaluepershare': 'tangibleBookValuePerShare',
    'shareholdersequitypershare': 'shareholdersEquityPerShare',
    'interestdebtpershare': 'interestDebtPerShare',
    'marketcap': 'marketCap',
    'enterprisevalue': 'enterpriseValue',
    'peratio': 'peRatio',
    'pricetosalesratio': 'priceToSalesRatio',
    'pocfratio': 'pocfratio',
    'pfcfratio': 'pfcfRatio',
    'pbratio': 'pbRatio',
    'ptbratio': 'ptbRatio',
    'evtosales': 'evToSales',
    'enterprisevalueoverebitda': 'enterpriseValueOverEBITDA',
    'evtooperatingcashflow': 'evToOperatingCashFlow',
    'evtofreecashflow': 'evToFreeCashFlow',
    'earningsyield': 'earningsYield',
    'freecashflowyield': 'freeCashFlowYield',
    'debttoequity': 'debtToEquity',
    'debttoassets': 'debtToAssets',
    'netdebttoebitda': 'netDebtToEBITDA',
    'currentratio': 'currentRatio',
    'interestcoverage': 'interestCoverage',
    'incomequality': 'incomeQuality',
    'dividendyield': 'dividendYield',
    'payoutratio': 'payoutRatio',
    'salesgeneralandadministrativetorevenue': 'salesGeneralAndAdministrativeToRevenue',
    'researchanddevelopementtorevenue': 'researchAndDevelopmentToRevenue',
    'intangiblestototalassets': 'intangiblesToTotalAssets',
    'capextooperatingcashflow': 'capexToOperatingCashFlow',
    'capextorevenue': 'capexToRevenue',
    'capextodepreciation': 'capexToDepreciation',
    'stockbasedcompensationtorevenue': 'stockBasedCompensationToRevenue',
    'grahamnumber': 'grahamNumber',
    'roic': 'roic',
    'returnontangibleassets': 'returnOnTangibleAssets',
    'grahamnetnet': 'grahamNetNet',
    'workingcapital': 'workingCapital',
    'tangibleassetvalue': 'tangibleAssetValue',
    'netcurrentassetvalue': 'netCurrentAssetValue',
    'investedcapital': 'investedCapital',
    'averagereceivables': 'averageReceivables',
    'averagepayables': 'averagePayables',
    'averageinventory': 'averageInventory',
    'dayssalesoutstanding': 'daysSalesOutstanding',
    'dayspayablesoutstanding': 'daysPayablesOutstanding',
    'daysofinventoryonhand': 'daysOfInventoryOnHand',
    'receivablesturnover': 'receivablesTurnover',
    'payablesturnover': 'payablesTurnover',
    'inventoryturnover': 'inventoryTurnover',
    'roe': 'roe',
    'capexpershare': 'capexPerShare',

    // Ratios
    'currentratio': 'currentRatio',
    'quickratio': 'quickRatio',
    'cashratio': 'cashRatio',
    'daysofsalesoutstanding': 'daysOfSalesOutstanding',
    'daysofinventoryoutstanding': 'daysOfInventoryOutstanding',
    'operatingcycle': 'operatingCycle',
    'daysofpayablesoutstanding': 'daysOfPayablesOutstanding',
    'cashconversioncycle': 'cashConversionCycle',
    'grossprofitmargin': 'grossProfitMargin',
    'operatingprofitmargin': 'operatingProfitMargin',
    'pretaxprofitmargin': 'pretaxProfitMargin',
    'netprofitmargin': 'netProfitMargin',
    'effectivetaxrate': 'effectiveTaxRate',
    'returnonassets': 'returnOnAssets',
    'returnonequity': 'returnOnEquity',
    'returnoncapitalemployed': 'returnOnCapitalEmployed',
    'netincomeperebt': 'netIncomePerEBT',
    'ebtperebit': 'ebtPerEbit',
    'ebitperrevenue': 'ebitPerRevenue',
    'debtratio': 'debtRatio',
    'debtequityratio': 'debtEquityRatio',
    'longtermdebttocapitalization': 'longTermDebtToCapitalization',
    'totaldebttocapitalization': 'totalDebtToCapitalization',
    'interestcoverage': 'interestCoverage',
    'cashflowcoverageratios': 'cashFlowCoverageRatios',
    'shorttermcoverageratios': 'shortTermCoverageRatios',
    'capitalexpenditurecoverageratio': 'capitalExpenditureCoverageRatio',
    'dividendpaidandcapexcoverageratio': 'dividendPaidAndCapexCoverageRatio',
    'dividendpayoutratio': 'dividendPayoutRatio',
    'pricebookvalueratio': 'priceBookValueRatio',
    'pricetobookratio': 'priceToBookRatio',
    'pricetosalesratio': 'priceToSalesRatio',
    'priceearningsratio': 'priceEarningsRatio',
    'pricetofreecashflowsratio': 'priceToFreeCashFlowsRatio',
    'pricetooperatingcashflowsratio': 'priceToOperatingCashFlowsRatio',
    'pricecashflowratio': 'priceCashFlowRatio',
    'priceearningstogrowthratio': 'priceEarningsToGrowthRatio',
    'pricesalesratio': 'priceSalesRatio',
    'dividendyield': 'dividendYield',
    'enterprisevaluemultiple': 'enterpriseValueMultiple',
    'pricefairvalue': 'priceFairValue'
  };

  try {
    symbol = symbol.toUpperCase();
    const loweredMetric = metric.toLowerCase();
    const endpoint = metricToEndpoint[loweredMetric];
    if (!endpoint) return 'Invalid metric: ' + metric + '. Please check available metrics in function description.';

    let properRawKey = metricToRawKey[loweredMetric];
    if (!properRawKey) return 'Invalid metric: ' + metric + '. Please check available metrics in function description.';

    function formatHeader(key) {
      return key
        .replace(/([a-z])([A-Z])/g, '$1 $2')
        .replace(/_/g, ' ')
        .replace(/\b\w/g, char => char.toUpperCase());
    }

    const loweredPeriod = period.toLowerCase();
    if (loweredPeriod === 'ttm') {
      properRawKey += 'TTM';
    }

    let url;
    if (loweredPeriod === 'ttm') {
      url = `https://financialmodelingprep.com/api/v3/${endpoint}-ttm/${symbol}?apikey=${apiKey}`;
    } else {
      url = `https://financialmodelingprep.com/api/v3/${endpoint}/${symbol}?limit=500`;
      if (period) url += `&period=${loweredPeriod}`;
      url += `&apikey=${apiKey}`;
    }

    const content = fetchWithCache(url, 3600);
    let data = JSON.parse(content);

    if (data.length === 0) return 'No ratio data available for symbol ' + symbol + ' and metric ' + metric;

    // Filter by year if provided
    if (year) {
      data = data.filter(row => row.calendarYear === year);
      if (data.length === 0) return 'No data available for the specified year ' + year;
    }

    // Sort by date descending (latest first)
    data.sort((a, b) => new Date(b.date) - new Date(a.date));

    const rawHeaders = Object.keys(data[0]);

    if (!rawHeaders.includes(properRawKey)) return 'Metric ' + metric + ' not found in the data for symbol ' + symbol;

    const formattedMetric = formatHeader(properRawKey.replace('TTM', ''));

    // Handle TTM special case (no 'date')
    if (loweredPeriod === 'ttm' && !rawHeaders.includes('date')) {
      if (data.length !== 1) return 'Unexpected data structure for TTM metric ' + metric;
      const value = data[0][properRawKey] || 0;
      if (showHeaders) {
        return [['Period', formattedMetric], ['TTM', value]];
      } else {
        return value;
      }
    }

    // Normal case with date
    if (!rawHeaders.includes('date')) return 'No date field found in data for metric ' + metric;

    if (showHeaders) {
      let table = data.map(row => [new Date(row.date), row[properRawKey] || 0]);
      table.unshift(['Date', formattedMetric]);
      return table;
    } else {
      return data[0][properRawKey] || 0;
    }

  } catch (error) {
    return 'An error occurred while fetching ratio ' + metric + ' for symbol ' + symbol + ', period ' + period + ', year ' + year + '. Please check the parameters.';
  }

}





/**
 * Retrieves growth metric. Returns latest rate or history for growth like revenueGrowth or epsGrowth.
 * @param {string} symbol - Stock ticker symbol (e.g., MSFT).
 * @param {string} metric - Metric: revenueGrowth, epsGrowth, dividendsPerShareGrowth, etc. (see code for full list).
 * @param {boolean} [showHeaders=false] - Return history table if true.
 * @param {string} [period=''] - Period: annual, quarter, q1, q2, q3, q4, fy.
 * @param {string} [year=''] - Year filter.
 * @return {number|array} - Growth value or table.
 * @customfunction
 */
function DIVIDENDDATA_GROWTH(symbol, metric, showHeaders = false, period = '', year = '') {
  const apiKey = getApiKey();
  if (!apiKey) return 'API key not set. Please run setApiKey() to configure.';

  const metricToRawKey = {
    'revenuegrowth': 'revenueGrowth',
    'grossprofitgrowth': 'grossProfitGrowth',
    'ebitgrowth': 'ebitgrowth',
    'operatingincomegrowth': 'operatingIncomeGrowth',
    'netincomegrowth': 'netIncomeGrowth',
    'epsgrowth': 'epsgrowth',
    'epsdilutedgrowth': 'epsdilutedGrowth',
    'weightedaveragesharesgrowth': 'weightedAverageSharesGrowth',
    'weightedaveragesharesdilutedgrowth': 'weightedAverageSharesDilutedGrowth',
    'dividendspersharegrowth': 'dividendsPerShareGrowth',
    'operatingcashflowgrowth': 'operatingCashFlowGrowth',
    'receivablesgrowth': 'receivablesGrowth',
    'inventorygrowth': 'inventoryGrowth',
    'assetgrowth': 'assetGrowth',
    'bookvaluepersharegrowth': 'bookValueperShareGrowth',
    'debtgrowth': 'debtGrowth',
    'rdexpensegrowth': 'rdexpenseGrowth',
    'sgaexpensesgrowth': 'sgaexpensesGrowth',
    'freecashflowgrowth': 'freeCashFlowGrowth',
    'tenyrevenuegrowthpershare': 'tenYRevenueGrowthPerShare',
    'fiveyrevenuegrowthpershare': 'fiveYRevenueGrowthPerShare',
    'threeyrevenuegrowthpershare': 'threeYRevenueGrowthPerShare',
    'tenyoperatingcfgrowthpershare': 'tenYOperatingCFGrowthPerShare',
    'fiveyoperatingcfgrowthpershare': 'fiveYOperatingCFGrowthPerShare',
    'threeyoperatingcfgrowthpershare': 'threeYOperatingCFGrowthPerShare',
    'tenynetincomegrowthpershare': 'tenYNetIncomeGrowthPerShare',
    'fiveynetincomegrowthpershare': 'fiveYNetIncomeGrowthPerShare',
    'threeynetincomegrowthpershare': 'threeYNetIncomeGrowthPerShare',
    'tenyshareholdersequitygrowthpershare': 'tenYShareholdersEquityGrowthPerShare',
    'fiveyshareholdersequitygrowthpershare': 'fiveYShareholdersEquityGrowthPerShare',
    'threeyshareholdersequitygrowthpershare': 'threeYShareholdersEquityGrowthPerShare',
    'tenydividendpersharegrowthpershare': 'tenYDividendperShareGrowthPerShare',
    'fiveydividendpersharegrowthpershare': 'fiveYDividendperShareGrowthPerShare',
    'threeydividendpersharegrowthpershare': 'threeYDividendperShareGrowthPerShare',
    'ebitdagrowth': 'ebitdaGrowth',
    'growthcapitalexpenditure': 'growthCapitalExpenditure',
    'tenybottomlinenetincomegrowthpershare': 'tenYBottomLineNetIncomeGrowthPerShare',
    'fiveybottomlinenetincomegrowthpershare': 'fiveYBottomLineNetIncomeGrowthPerShare',
    'threeybottomlinenetincomegrowthpershare': 'threeYBottomLineNetIncomeGrowthPerShare'
  };

  try {
    symbol = symbol.toUpperCase();
    const loweredMetric = metric.toLowerCase();
    const properRawKey = metricToRawKey[loweredMetric];
    if (!properRawKey) return 'Invalid metric: ' + metric + '. Please check available metrics in function description.';

    function formatHeader(key) {
      return key
        .replace(/([a-z])([A-Z])/g, '$1 $2')
        .replace(/_/g, ' ')
        .replace(/\b\w/g, char => char.toUpperCase());
    }

    const formattedMetric = formatHeader(properRawKey);

    let apiPeriod = '';
    const loweredPeriod = period.toLowerCase();
    if (loweredPeriod === 'annual' || loweredPeriod === 'fy') {
      apiPeriod = 'annual';
    } else if (loweredPeriod === 'quarter') {
      apiPeriod = 'quarter';
    } else if (['q1', 'q2', 'q3', 'q4'].includes(loweredPeriod)) {
      apiPeriod = loweredPeriod.toUpperCase();
    }

    let url = `https://financialmodelingprep.com/stable/financial-growth?symbol=${symbol}`;
    if (apiPeriod) url += `&period=${apiPeriod}`;
    url += `&limit=500&apikey=${apiKey}`;

    const content = fetchWithCache(url, 3600);
    let data = JSON.parse(content);

    if (!Array.isArray(data)) data = [data];

    if (data.length === 0) return 'No growth data available for symbol ' + symbol + ' and metric ' + metric;

    // Filter by year if provided
    if (year) {
      data = data.filter(row => row.fiscalYear === year);
      if (data.length === 0) return 'No data available for the specified year ' + year;
    }

    // Sort by date descending (latest first)
    data.sort((a, b) => new Date(b.date) - new Date(a.date));

    const rawHeaders = Object.keys(data[0]);

    if (!rawHeaders.includes(properRawKey)) return 'Metric ' + metric + ' not found in the data for symbol ' + symbol;

    if (showHeaders) {
      let table = data.map(row => [new Date(row.date), row[properRawKey] || 0]);
      table.unshift(['Date', formattedMetric]);
      return table;
    } else {
      return data[0][properRawKey] || 0;
    }

  } catch (error) {
    return 'An error occurred while fetching growth metric ' + metric + ' for symbol ' + symbol + ', period ' + period + ', year ' + year + '. Please check the parameters.';
  }
}


/**
 * Retrieves stock quote data. Returns price, change, volume, full details, or history.
 * @param {string} symbol - Stock ticker symbol (e.g., AAPL).
 * @param {string} [metric="price"] - Metric: price, change, volume, full, history.
 * @param {string} [fromDate] - Start date for history (YYYY-MM-DD).
 * @param {string} [toDate] - End date for history (YYYY-MM-DD).
 * @param {boolean} [showHeaders=true] - Include headers for full or history.
 * @return {string|number|array} - Quote data.
 * @customfunction
 */
function DIVIDENDDATA_QUOTE(symbol, metric = "price", fromDate, toDate, showHeaders = true) {
  const apiKey = getApiKey();
  if (!apiKey) return 'API key not set. Please run setApiKey() to configure.';

  try {
    symbol = symbol.toUpperCase();
    const loweredMetric = metric.toLowerCase();

    let actualFromDate = fromDate;
    let actualToDate = toDate;
    let actualShowHeaders = showHeaders;

    // Adjust parameters if showHeaders is misplaced
    if (typeof toDate === 'boolean') {
      actualShowHeaders = toDate;
      actualToDate = undefined;
    }
    if (typeof fromDate === 'boolean') {
      actualShowHeaders = fromDate;
      actualFromDate = undefined;
      actualToDate = undefined;
    }

    // Handle year-only dates for history
    if (typeof actualFromDate === 'string' && actualFromDate.match(/^\d{4}$/)) {
      actualFromDate += '-01-01';
    }
    if (typeof actualToDate === 'string' && actualToDate.match(/^\d{4}$/)) {
      actualToDate += '-12-31';
    }

    // Default to last 365 days if no dates provided for history
    if (loweredMetric === 'history' && !actualFromDate && !actualToDate) {
      const today = new Date();
      const oneYearAgo = new Date(today.getTime() - 365 * 24 * 60 * 60 * 1000);
      actualFromDate = oneYearAgo.toISOString().split('T')[0];
      actualToDate = today.toISOString().split('T')[0];
    }

    // Function to format headers to human-readable
    function formatHeader(key) {
      return key
        .replace(/([a-z])([A-Z])/g, '$1 $2')
        .replace(/_/g, ' ')
        .replace(/\b\w/g, char => char.toUpperCase());
    }

    if (["price", "change", "volume"].includes(loweredMetric)) {
      const url = `https://financialmodelingprep.com/stable/quote-short?symbol=${symbol}&apikey=${apiKey}`;
      const content = fetchWithCache(url, 60);
      const data = JSON.parse(content);

      if (data.length === 0) return 'No quote data available for symbol ' + symbol;

      const item = data[0];
      switch (loweredMetric) {
        case 'price': return item.price || 0;
        case 'change': return item.change || 0;
        case 'volume': return item.volume || 0;
      }
    } else if (loweredMetric === 'full') {
      const url = `https://financialmodelingprep.com/stable/quote?symbol=${symbol}&apikey=${apiKey}`;
      const content = fetchWithCache(url, 60);
      const data = JSON.parse(content);

      if (data.length === 0) return 'No full quote data available for symbol ' + symbol;

      const item = data[0];
      const rawHeaders = Object.keys(item);
      const formattedHeaders = rawHeaders.map(formatHeader);

      const row = rawHeaders.map(key => {
        let value = item[key];
        if (key === 'timestamp') {
          return new Date(value * 1000);
        }
        return value;
      });

      let table = [row];

      if (actualShowHeaders) {
        table.unshift(formattedHeaders);
      }

      return table;
    } else if (loweredMetric === 'history') {
      let url = `https://financialmodelingprep.com/stable/historical-price-eod/light?symbol=${symbol}`;
      if (actualFromDate) url += `&from=${actualFromDate}`;
      if (actualToDate) url += `&to=${actualToDate}`;
      url += `&apikey=${apiKey}`;
      
      const content = fetchWithCache(url, 300);  // 5 min for history
      const data = JSON.parse(content);

      if (data.length === 0) return [['No historical quote data available for symbol ' + symbol]];

      // Sort by date descending (latest first)
      data.sort((a, b) => new Date(b.date) - new Date(a.date));

      const desiredKeys = ['date', 'price', 'volume'];
      const formattedHeaders = desiredKeys.map(formatHeader);

      let table = data.map(row => desiredKeys.map(key => {
        let value = row[key];
        if (key.toLowerCase() === 'date') {
          return new Date(value);
        }
        return value;
      }));

      if (actualShowHeaders) {
        table.unshift(formattedHeaders);
      }

      return table;
    } else {
      return 'Invalid metric: ' + metric + '. Valid metrics are: price, change, volume, full, history.';
    }
  } catch (error) {
    return 'An error occurred while fetching quote for symbol ' + symbol + ' and metric ' + metric + '. Please check the symbol and parameters.';
  }
}


/**
 * Retrieves company profile. Returns specific detail or full table for info like marketcap or sector.
 * @param {string} symbol - Stock ticker symbol (e.g., MSFT).
 * @param {string} metric - Metric: marketcap, beta, lastdividend, companyname, sector, etc., or "full".
 * @param {boolean} [showHeaders=false] - Include headers for full.
 * @return {string|number|array} - Profile data.
 * @customfunction
 */
function DIVIDENDDATA_PROFILE(symbol, metric, showHeaders = false) {
  const apiKey = getApiKey();
  if (!apiKey) return 'API key not set. Please run setApiKey() to configure.';

  try {
    symbol = symbol.toUpperCase();
    const url = `https://financialmodelingprep.com/stable/profile?symbol=${symbol}&apikey=${apiKey}`;
    const content = fetchWithCache(url, 3600);
    const data = JSON.parse(content);

    if (data.length === 0) return 'No profile data available for symbol ' + symbol;

    const loweredMetric = metric.toLowerCase();

    if (loweredMetric === 'full') {
      const item = data[0];
      const rawHeaders = Object.keys(item);

      function formatHeader(key) {
        return key
          .replace(/([a-z])([A-Z])/g, '$1 $2')
          .replace(/_/g, ' ')
          .replace(/\b\w/g, char => char.toUpperCase());
      }

      const formattedHeaders = rawHeaders.map(formatHeader);

      const row = rawHeaders.map(key => {
        let value = item[key];
        if (key.toLowerCase() === 'ipodate') {
          return value ? new Date(value) : '';
        }
        return value !== null ? value : '';
      });

      let table = [row];

      if (showHeaders) {
        table.unshift(formattedHeaders);
      }

      return table;
    } else {
      const metricToRawKey = {
        'symbol': 'symbol',
        'price': 'price',
        'marketcap': 'marketCap',
        'beta': 'beta',
        'lastdividend': 'lastDividend',
        'range': 'range',
        'change': 'change',
        'changepercentage': 'changePercentage',
        'volume': 'volume',
        'averagevolume': 'averageVolume',
        'companyname': 'companyName',
        'currency': 'currency',
        'cik': 'cik',
        'isin': 'isin',
        'cusip': 'cusip',
        'exchangefullname': 'exchangeFullName',
        'exchange': 'exchange',
        'industry': 'industry',
        'website': 'website',
        'description': 'description',
        'ceo': 'ceo',
        'sector': 'sector',
        'country': 'country',
        'fulltimeemployees': 'fullTimeEmployees',
        'phone': 'phone',
        'address': 'address',
        'city': 'city',
        'state': 'state',
        'zip': 'zip',
        'image': 'image',
        'ipodate': 'ipoDate',
        'defaultimage': 'defaultImage',
        'isetf': 'isEtf',
        'isactivelytrading': 'isActivelyTrading',
        'isadr': 'isAdr',
        'isfund': 'isFund'
      };

      const rawKey = metricToRawKey[loweredMetric];
      if (!rawKey) return 'Invalid metric: ' + metric + '. Valid metrics are: symbol, price, marketcap, beta, lastdividend, range, change, changepercentage, volume, averagevolume, companyname, currency, cik, isin, cusip, exchangefullname, exchange, industry, website, description, ceo, sector, country, fulltimeemployees, phone, address, city, state, zip, image, ipodate, defaultimage, isetf, isactivelytrading, isadr, isfund, or "full".';

      let value = data[0][rawKey];
      if (rawKey.toLowerCase() === 'ipodate') {
        return value ? new Date(value) : '';
      }
      return value !== null ? value : '';
    }
  } catch (error) {
    return 'An error occurred while fetching profile for symbol ' + symbol + ' and metric ' + metric + '. Please check the symbol and parameters.';
  }
}



/**
 * Retrieves ETF/fund data. Returns details like expenseRatio or tables for holdings.
 * @param {string} symbol - Fund ticker (e.g., SPY).
 * @param {string} metric - Metric: holdings, countryweighting, symbol, name, description, isin, assetclass, securitycusip, domicile, website, etfcompany, expenseratio, assetsundermanagement, avgvolume, inceptiondate, nav, navcurrency, holdingscount, updatedat, sectorslist.
 * @param {boolean} [showHeaders=true] - Include headers for tables.
 * @return {string|number|array} - Fund data.
 * @customfunction
 */
function DIVIDENDDATA_FUND(symbol, metric, showHeaders = true) {
  const apiKey = getApiKey();
  if (!apiKey) return 'API key not set. Please run setApiKey() to configure.';

  try {
    symbol = symbol.toUpperCase();
    const loweredMetric = metric.toLowerCase();

    let infoData;
    if (loweredMetric !== 'holdings' && loweredMetric !== 'countryweighting') {
      const infoUrl = `https://financialmodelingprep.com/stable/etf/info?symbol=${symbol}&apikey=${apiKey}`;
      const infoContent = fetchWithCache(infoUrl, 3600);
      infoData = JSON.parse(infoContent);

      if (infoData.length === 0) return 'No fund data available for symbol ' + symbol;
    }

    switch (loweredMetric) {
      case 'holdings':
        const holdingsUrl = `https://financialmodelingprep.com/stable/etf/holdings?symbol=${symbol}&apikey=${apiKey}`;
        const holdingsContent = fetchWithCache(holdingsUrl, 3600);
        const holdingsData = JSON.parse(holdingsContent);

        if (holdingsData.length === 0) return [['No holdings data available for symbol ' + symbol]];

        const holdingsHeaders = Object.keys(holdingsData[0]);

        function formatHeader(key) {
          return key
            .replace(/([a-z])([A-Z])/g, '$1 $2')
            .replace(/_/g, ' ')
            .replace(/\b\w/g, char => char.toUpperCase());
        }

        const formattedHoldingsHeaders = holdingsHeaders.map(formatHeader);

        const holdingsTable = holdingsData.map(row => holdingsHeaders.map(key => {
          const value = row[key];
          if (key.toLowerCase() === 'updatedat') return new Date(value);
          return value;
        }));

        if (showHeaders) {
          holdingsTable.unshift(formattedHoldingsHeaders);
        }

        return holdingsTable;

      case 'countryweighting':
        const cwUrl = `https://financialmodelingprep.com/stable/etf/country-weightings?symbol=${symbol}&apikey=${apiKey}`;
        const cwContent = fetchWithCache(cwUrl, 3600);
        const cwData = JSON.parse(cwContent);

        if (cwData.length === 0) return [['No country weighting data available for symbol ' + symbol]];

        const cwTable = cwData.map(row => [row.country, row.weightPercentage]);

        if (showHeaders) {
          cwTable.unshift(['Country', 'Weight Percentage']);
        }

        return cwTable;

      case 'sectorslist':
        const sectors = infoData[0].sectorsList || [];
        if (sectors.length === 0) return [['No sectors data available for symbol ' + symbol]];

        const sectorsTable = sectors.map(row => [row.industry, row.exposure]);

        if (showHeaders) {
          sectorsTable.unshift(['Industry', 'Exposure']);
        }

        return sectorsTable;

      default:
        const metricToRawKey = {
          'symbol': 'symbol',
          'name': 'name',
          'description': 'description',
          'isin': 'isin',
          'assetclass': 'assetClass',
          'securitycusip': 'securityCusip',
          'domicile': 'domicile',
          'website': 'website',
          'etfcompany': 'etfCompany',
          'expenseratio': 'expenseRatio',
          'assetsundermanagement': 'assetsUnderManagement',
          'avgvolume': 'avgVolume',
          'inceptiondate': 'inceptionDate',
          'nav': 'nav',
          'navcurrency': 'navCurrency',
          'holdingscount': 'holdingsCount',
          'updatedat': 'updatedAt'
        };

        const rawKey = metricToRawKey[loweredMetric];
        if (!rawKey) return 'Invalid metric: ' + metric + '. Valid metrics are: holdings, countryweighting, symbol, name, description, isin, assetclass, securitycusip, domicile, website, etfcompany, expenseratio, assetsundermanagement, avgvolume, inceptiondate, nav, navcurrency, holdingscount, updatedat, sectorslist.';

        let value = infoData[0][rawKey];

        if (loweredMetric === 'inceptiondate' || loweredMetric === 'updatedat') {
          return value ? new Date(value) : '';
        }

        return value !== undefined ? value : '';
    }
  } catch (error) {
    return 'An error occurred while fetching fund data for symbol ' + symbol + ' and metric ' + metric + '. Please check the symbol and parameters.';
  }
}

/**
 * Retrieves revenue segments. Returns table by product or geography, filtered by period and year.
 * @param {string} symbol - Stock ticker symbol (e.g., AAPL).
 * @param {string} metric - Type: products or geographic.
 * @param {string} [period=annual] - Period: annual or quarter.
 * @param {string} [year] - Year filter.
 * @param {boolean} [showHeaders=true] - Include headers.
 * @return {array} - Segments table.
 * @customfunction
 */
function DIVIDENDDATA_SEGMENTS(symbol, metric, period = 'annual', year = '', showHeaders = true) {
  const apiKey = getApiKey();
  if (!apiKey) return 'API key not set. Please run setApiKey() to configure.';

  try {
    symbol = symbol.toUpperCase();
    const loweredMetric = String(metric).toLowerCase();
    let loweredPeriod = String(period).toLowerCase();
    let filterYear = String(year).trim();
    let headersFlag = showHeaders;

    // Handle if year is boolean (user put showHeaders as fourth arg)
    if (typeof year === 'boolean') {
      headersFlag = year;
      filterYear = '';
    }

    // Handle if period is boolean (user put showHeaders as third arg)
    if (typeof period === 'boolean') {
      headersFlag = period;
      loweredPeriod = 'annual';
      filterYear = '';
    }

    let endpoint;
    if (loweredMetric === 'products') {
      endpoint = 'revenue-product-segmentation';
    } else if (loweredMetric === 'geographic') {
      endpoint = 'revenue-geographic-segmentation';
    } else {
      return 'Invalid metric: ' + metric + '. Valid metrics are: products or geographic.';
    }

    if (loweredPeriod !== 'annual' && loweredPeriod !== 'quarter') {
      return 'Invalid period: ' + period + '. Valid periods are: annual or quarter.';
    }

    const url = `https://financialmodelingprep.com/stable/${endpoint}?symbol=${symbol}&period=${loweredPeriod}&apikey=${apiKey}`;
    const content = fetchWithCache(url, 3600);
    let data = JSON.parse(content);

    if (data.length === 0) return [['No segments data available for symbol ' + symbol + ' and metric ' + metric]];

    // Filter by year if provided
    if (filterYear) {
      data = data.filter(item => String(item.fiscalYear) === filterYear);
      if (data.length === 0) return [['No data available for the specified year ' + filterYear]];
    }

    // Sort by date descending (latest first)
    data.sort((a, b) => new Date(b.date) - new Date(a.date));

    // Collect all unique segments across all years
    const allSegments = new Set();
    data.forEach(item => {
      Object.keys(item.data).forEach(seg => allSegments.add(seg));
    });

    let segments = Array.from(allSegments).sort();

    // Function to format headers
    function formatHeader(key) {
      return key
        .replace(/([a-z])([A-Z])/g, '$1 $2')
        .replace(/_/g, ' ')
        .replace(/\b\w/g, char => char.toUpperCase());
    }

    const formattedSegments = segments.map(formatHeader);

    // Build headers
    let headers = ['Fiscal Year', 'Date', ...formattedSegments];

    // Build table
    let table = data.map(item => {
      let row = [item.fiscalYear, new Date(item.date)];
      segments.forEach(seg => {
        row.push(item.data[seg] || 0);
      });
      return row;
    });

    // Optional headers
    if (headersFlag) {
      table.unshift(headers);
    }

    return table;

  } catch (error) {
    return 'An error occurred while fetching segments for symbol ' + symbol + ', metric ' + metric + ', period ' + period + ', year ' + year + '. Please check the parameters.';
  }
}

/**
 * Retrieves KPIs and segments from Fiscal.ai. Returns table with year/quarter and metrics.
 * @param {string} ticker - Stock ticker symbol (e.g., MSFT).
 * @param {string} [period=annual] - Period: annual or quarterly.
 * @param {string|number} [year] - Year filter.
 * @param {boolean} [showheaders=TRUE] - Include headers.
 * @return {array} - KPIs table.
 * @customfunction
 */
function DIVIDENDDATA_KPIS(ticker, period = 'annual', year = '', showheaders = true) {
  const fiscalApiKey = '025bb017-5cdd-4aef-a8fb-b36fecb3ea27';

  try {
    ticker = ticker.toUpperCase();
    const loweredPeriod = period.toLowerCase();
    if (loweredPeriod !== 'annual' && loweredPeriod !== 'quarterly') {
      return 'Invalid period: ' + period + '. Valid periods are: annual or quarterly.';
    }

    const periodType = 'annual%2Cquarterly';
    const url = `https://api.fiscal.ai/v1/company/segments-and-kpis?ticker=${ticker}&periodType=${periodType}&apiKey=${fiscalApiKey}`;
    const content = fetchWithCache(url, 3600);
    const json = JSON.parse(content);

    const { metrics, data } = json;

    const periodKey = loweredPeriod.charAt(0).toUpperCase() + loweredPeriod.slice(1);

    // Filter active metrics (not discontinued for this period)
    const activeMetrics = metrics.filter(m => !m.isDiscontinued[periodKey]);

    const metricIds = activeMetrics.map(m => m.metricId);
    const metricNames = activeMetrics.map(m => m.metricName);

    // Filter data to this periodType
    let filteredData = data.filter(d => d.periodType === periodKey);

    // Filter by year if provided
    if (year) {
      filteredData = filteredData.filter(d => d.calendarYear == year);
    }

    if (filteredData.length === 0) {
      return [['No KPI data available for ticker ' + ticker + ' and period ' + period], ['Data Powered By Fiscal.ai']];
    }

    // Sort by reportDate descending (latest first)
    filteredData.sort((a, b) => new Date(b.reportDate) - new Date(a.reportDate));

    // Build headers
    let headers;
    if (loweredPeriod === 'annual') {
      headers = ['Calendar Year', 'Report Date', ...metricNames];
    } else {
      headers = ['Calendar Year', 'Calendar Quarter', 'Report Date', ...metricNames];
    }

    // Build table
    const table = filteredData.map(row => {
      let base;
      if (loweredPeriod === 'annual') {
        base = [row.calendarYear, new Date(row.reportDate)];
      } else {
        base = [row.calendarYear, row.calendarQuarter, new Date(row.reportDate)];
      }
      metricIds.forEach(id => {
        base.push(row.metricsValues[id] || 0);
      });
      return base;
    });

    if (showheaders) {
      table.unshift(headers);
    }

    // Add attribution footer
    const footer = ['Data Powered By Fiscal.ai'];
    for (let i = 1; i < headers.length; i++) {
      footer.push('');
    }
    table.push(footer);

    return table;

  } catch (error) {
    return 'An error occurred while fetching KPIs for ticker ' + ticker + ', period ' + period + ', year ' + year + '. Please check the parameters.\nData Powered By Fiscal.ai';
  }
}

/**
 * Retrieves commodities data. Returns list, price, full quote, or history.
 * @param {string} symbol - Commodity symbol (e.g., CLUSD).
 * @param {string} metric - Metric: list, price, fullquote, history.
 * @param {string} [fromDate] - Start date for history.
 * @param {string} [toDate] - End date for history.
 * @param {boolean} [showHeaders=true] - Include headers for tables.
 * @return {number|string|array} - Commodities data.
 * @customfunction
 */
function DIVIDENDDATA_COMMODITIES(symbol, metric, fromDate, toDate, showHeaders = true) {
  const apiKey = getApiKey();
  if (!apiKey) return 'API key not set. Please run setApiKey() to configure.';

  try {
    let actualFromDate = fromDate;
    let actualToDate = toDate;
    let actualShowHeaders = showHeaders;

    // Adjust parameters if showHeaders is misplaced
    if (typeof toDate === 'boolean') {
      actualShowHeaders = toDate;
      actualToDate = undefined;
    }
    if (typeof fromDate === 'boolean') {
      actualShowHeaders = fromDate;
      actualFromDate = undefined;
      actualToDate = undefined;
    }

    // Handle year-only dates for history
    if (typeof actualFromDate === 'string' && actualFromDate.match(/^\d{4}$/)) {
      actualFromDate += '-01-01';
    }
    if (typeof actualToDate === 'string' && actualToDate.match(/^\d{4}$/)) {
      actualToDate += '-12-31';
    }

    // Default to last 365 days if no dates provided for history
    if (metric.toLowerCase() === 'history' && !actualFromDate && !actualToDate) {
      const today = new Date();
      const oneYearAgo = new Date(today.getTime() - 365 * 24 * 60 * 60 * 1000);
      actualFromDate = oneYearAgo.toISOString().split('T')[0];
      actualToDate = today.toISOString().split('T')[0];
    }

    let url, content, data;

    // Function to format headers to human-readable
    function formatHeader(key) {
      return key
        .replace(/([a-z])([A-Z])/g, '$1 $2')
        .replace(/_/g, ' ')
        .replace(/\b\w/g, char => char.toUpperCase());
    }

    const loweredMetric = metric.toLowerCase();

    switch (loweredMetric) {
      case 'list':
        url = `https://financialmodelingprep.com/stable/commodities-list?apikey=${apiKey}`;
        content = fetchWithCache(url, 86400);  // 24 hours
        data = JSON.parse(content);

        if (data.length === 0) return [['No commodities list data available']];

        const rawHeaders = Object.keys(data[0]);
        const formattedHeaders = rawHeaders.map(formatHeader);

        let table = data.map(row => rawHeaders.map(key => row[key] || ''));

        if (actualShowHeaders) {
          table.unshift(formattedHeaders);
        }

        return table;

      case 'price':
      case 'fullquote':
        if (!symbol) return 'Symbol parameter is required for price or fullquote.';
        symbol = symbol.toUpperCase();
        url = `https://financialmodelingprep.com/stable/quote?symbol=${symbol}&apikey=${apiKey}`;
        content = fetchWithCache(url, 60);
        data = JSON.parse(content);

        if (data.length === 0) return 'No quote data available for symbol ' + symbol;

        if (loweredMetric === 'price') {
          return data[0].price || 0;
        } else { // fullquote
          const item = data[0];
          const rawHeaders = Object.keys(item);
          const formattedHeaders = rawHeaders.map(formatHeader);

          const row = rawHeaders.map(key => {
            let value = item[key];
            if (key.toLowerCase().includes('timestamp')) {
              return new Date(value * 1000);
            }
            return value !== null ? value : '';
          });

          let fullTable = [row];

          if (actualShowHeaders) {
            fullTable.unshift(formattedHeaders);
          }

          return fullTable;
        }

      case 'history':
        if (!symbol) return 'Symbol parameter is required for history.';
        symbol = symbol.toUpperCase();
        url = `https://financialmodelingprep.com/stable/historical-price-eod/light?symbol=${symbol}`;
        if (actualFromDate) url += `&from=${actualFromDate}`;
        if (actualToDate) url += `&to=${actualToDate}`;
        url += `&apikey=${apiKey}`;
        
        content = fetchWithCache(url, 300);
        data = JSON.parse(content);

        if (data.length === 0) return [['No historical data available for symbol ' + symbol]];

        // Sort by date descending (latest first)
        data.sort((a, b) => new Date(b.date) - new Date(a.date));

        const desiredKeys = ['date', 'price', 'volume'];
        const formattedHeaders2 = desiredKeys.map(formatHeader);

        let historyTable = data.map(row => desiredKeys.map(key => {
          let value = row[key];
          if (key.toLowerCase() === 'date') {
            return new Date(value);
          }
          return typeof value === 'number' ? value : value || '';
        }));

        if (actualShowHeaders) {
          historyTable.unshift(formattedHeaders2);
        }

        return historyTable;

      default:
        return 'Invalid metric: ' + metric + '. Valid metrics are: list, price, fullquote, history.';
    }
  } catch (error) {
    return 'An error occurred while fetching commodities data for symbol ' + symbol + ' and metric ' + metric + '. Please check the parameters.';
  }
}

/**
 * Retrieves batch quotes for stocks. Returns table with prices, changes, or volumes.
 * @param {string} symbols - Comma-separated tickers (e.g., AAPL,MSFT).
 * @param {string} [metrics="all"] - Metrics: price, change, volume, or "all".
 * @param {boolean} [showHeaders] - Include headers (defaults based on metrics).
 * @return {array} - Quotes table.
 * @customfunction
 */
function DIVIDENDDATA_QUOTE_BATCH(symbols, metrics = "all", showHeaders) {
  const apiKey = getApiKey();
  if (!apiKey) return 'API key not set. Please run setApiKey() to configure.';

  try {
    if (!symbols) return 'Symbols parameter is required.';

    // Process symbols
    const symbolArray = symbols.split(',').map(s => s.trim().toUpperCase());
    const symbolList = symbolArray.join(',');

    const url = `https://financialmodelingprep.com/stable/batch-quote-short?symbols=${symbolList}&apikey=${apiKey}`;
    const content = fetchWithCache(url, 60);
    const data = JSON.parse(content);

    if (data.length === 0) return [['No quote data available for the provided symbols']];

    // Create map for preserving order and handling missing
    const dataMap = new Map(data.map(item => [item.symbol, item]));

    // Process metrics
    const allMetrics = ["price", "change", "volume"];
    let requested;
    if (metrics.toLowerCase() === "all") {
      requested = allMetrics;
    } else {
      requested = metrics.split(',').map(m => m.trim().toLowerCase());
      for (let m of requested) {
        if (!allMetrics.includes(m)) {
          return 'Invalid metric: ' + m + '. Valid metrics are: price, change, volume or "all".';
        }
      }
    }

    // Determine effective showHeaders (default based on metrics if not provided)
    if (showHeaders === undefined) {
      showHeaders = (metrics.toLowerCase() === "all");
    }

    // Determine fields
    const fields = showHeaders ? ["symbol", ...requested] : requested;

    // Format headers
    function formatHeader(key) {
      return key
        .replace(/([a-z])([A-Z])/g, '$1 $2')
        .replace(/_/g, ' ')
        .replace(/\b\w/g, char => char.toUpperCase());
    }

    const formattedHeaders = fields.map(formatHeader);

    // Build table
    let table = symbolArray.map(sym => {
      const item = dataMap.get(sym) || {};
      return fields.map(field => item[field] || 0);
    });

    if (showHeaders) {
      table.unshift(formattedHeaders);
    }

    return table;

  } catch (error) {
    return 'An error occurred while fetching batch quotes for symbols ' + symbols + ' and metrics ' + metrics + '. Please check the parameters.';
  }
}


// Add-on menu
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Initialize User', 'initializeUser')  // New menu item to store email
    .addItem('Open Dividend Data Sidebar', 'showSidebar')
    .addToUi();
}

// New function to store user email in UserProperties (runs with permissions)
function initializeUser() {
  const ui = SpreadsheetApp.getUi();
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    ui.alert('Error', 'Unable to retrieve your email. Ensure permissions are granted.', ui.ButtonSet.OK);
    return;
  }
  
  const properties = PropertiesService.getUserProperties();
  properties.setProperty('USER_EMAIL', email);
  
  ui.alert('Success', 'User initialized! Your email (' + email + ') is now stored for AI features.', ui.ButtonSet.OK);
}


// Dummy sidebar function (can be expanded later)
function showSidebar() {
  const html = HtmlService.createHtmlOutput('<p>Welcome to Dividend Data Add-on! Custom functions are now active.</p>')
    .setTitle('Dividend Data')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Homepage trigger (from your manifest; can be the same as sidebar)
function onHomepage() {
  showSidebar();
}

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


function testGetUserEmail() {
  const email = Session.getActiveUser().getEmail();
  Logger.log('User Email: ' + email);  // This will log to the Executions view
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
  if (!apiKey) throw new Error('Set your FMP API key via setApiKey()');

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
    default: throw new Error('Invalid years for CAGR');
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
 * Main dividend function: =DIVIDENDDATA("MSFT", "fwd_payout") or =DIVIDENDDATA("MSFT", "history", TRUE)
 * @param {string} symbol - Ticker (e.g., MSFT)
 * @param {string} metric - fwd_payout | ttm_payout | fwd_yield | ttm_yield | frequency | history | growth | 1Y_CAGR | 3Y_CAGR | 5Y_CAGR | 10Y_CAGR | payout_ratio | fcf_payout_ratio
 * @param {boolean} [showHeaders=false] - Include header row for tables (e.g., history)
 * @return {string|array} - Value or table
 * @customfunction
 */
function DIVIDENDDATA(symbol, metric = "fwd_payout", showHeaders = false) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('Set your FMP API key via setApiKey()');

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
        console.log(url);
        content = fetchWithCache(url, 60);  // Short for quotes
        data = JSON.parse(content);

        if (data.length === 0) throw new Error(`No data for ${symbol}`);
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
        if (data.length === 0) return 0;
        return data[0].frequency;

      case 'history':
        // Use original stable API - returns table [date, dividend, adjDividend], latest first
        url = `https://financialmodelingprep.com/stable/dividends?symbol=${symbol}&apikey=${apiKey}`;
        content = fetchWithCache(url, 3600);
        data = JSON.parse(content);

        if (data.length === 0) return [['No data for the provided ticker symbol']];

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
          cleaned.unshift(['Declartion Date', 'Record Date', 'Payment Date', 'Adjusted Dividend', 'Dividend', 'Yield', 'Frequency']);
        }

        return cleaned;  // Latest first (no reverse)

      case 'growth':
        // Replicate your dividend_growth_calc: % change YoY from history
        const histData = DIVIDENDDATA(symbol, 'history');  // Get without headers
        if (histData.length < 2) return [['Insufficient data']];

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

        if (data.length === 0) return NaN;
        const payoutRatio = data[0].dividendPayoutRatio || 0;
        return payoutRatio;  // Raw number (user formats as % in Sheets)

      case 'fcf_payout_ratio':
        // FCF payout ratio from stable ratios endpoint: (dividendPerShare / freeCashFlowPerShare) * 100
        url = `https://financialmodelingprep.com/stable/ratios?symbol=${symbol}&apikey=${apiKey}`;
        content = fetchWithCache(url, 3600);
        data = JSON.parse(content);

        if (data.length === 0) return NaN;
        const divPerShare = data[0].dividendPerShare || 0;
        const fcfPerShare = data[0].freeCashFlowPerShare || 0;
        return fcfPerShare > 0 ? (divPerShare / fcfPerShare) : NaN;  // Raw number or NaN

      default:
        throw new Error(`Invalid metric: ${metric}. Use: yield, history, growth, payout, fcfPayout`);
    }
  } catch (error) {
    return `Error: ${error.message}`;
  }
}


/**
 * Fetch batch dividend data: =DIVIDENDDATA_BATCH("MSFT,KMB,O", "fwd_payout,yield", TRUE)
 * @param {string} symbols - Comma-separated tickers (e.g., MSFT,KMB)
 * @param {string} [metric="fwd_payout"] - Comma-separated metrics: adjdividend,dividend,recorddate,paymentdate,declarationdate,yield,frequency or "all" or "fwd_payout" or "history"
 * @param {boolean} [showHeaders] - Include header row and symbol column (defaults to TRUE if metric="all" or "history", else FALSE)
 * @return {array} - Table or column(s) of dividend data
 * @customfunction
 */
function DIVIDENDDATA_BATCH(symbols, metric = "fwd_payout", showHeaders) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('Set your FMP API key via setApiKey()');

  try {
    if (!symbols) throw new Error('Symbols required');

    // Process symbols
    const symbolArray = symbols.split(',').map(s => s.trim().toUpperCase());
    const symbolList = symbolArray.join(',');

    const url = `https://financialmodelingprep.com/api/v4/private/dividend_data/dividend-yield?symbol=${symbolList}&apikey=${apiKey}`;
    const content = fetchWithCache(url, 3600);
    const data = JSON.parse(content);

    if (data.length === 0) return [['No data for the provided symbols']];

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
          throw new Error(`Invalid metric: ${m}. Use: adjdividend, dividend, recorddate, paymentdate, declarationdate, yield, frequency, fwd_payout or "all" or "history"`);
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
    return `Error: ${error.message}`;
  }
}


/**
 * Main dividend function: =DIVIDENDDATA_STATEMENT("MSFT", "income") or =DIVIDENDDATA_STATEMENT("AAPL", "cash_flow",TRUE, "Q1", "2025")
 * @param {string} symbol - Ticker (e.g., MSFT)
 * @param {string} metric - income | balance | cash_flow
 * @param {boolean} [showHeaders=false] - Include header row for tables
 * @param {string} period - FY | Q1 | Q2 | Q3 | Q4 | annual | quarter | ttm
 * @param {string} year - (e.g., 2025)
 * @return {string|array} - Value or table
 * @customfunction
 */
function DIVIDENDDATA_STATEMENT(symbol, metric, showHeaders = false, period = '', year = '') {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('Set your FMP API key via setApiKey()');

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
        throw new Error(`Invalid metric: ${metric}. Use: income, balance, cash_flow`);
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

    if (data.length === 0) return [['No data for the provided ticker symbol']];

    // Filter by year if provided (works for both regular and TTM)
    if (year) {
      data = data.filter(row => row.fiscalYear === year);
      if (data.length === 0) return [['No data for specified year']];
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
    return `Error: ${error.message}`;
  }

}


/**
 * Fetch a specific metric from financial statements: =DIVIDENDDATA_METRICS("MSFT", "revenue") or =DIVIDENDDATA_METRICS("MSFT", "freeCashFlow", true)
 * @param {string} symbol - Ticker (e.g., MSFT)
 * @param {string} metric - The specific metric (e.g., revenue, netIncome, freeCashFlow)
 * @param {boolean} [showHeaders=false] - If true, return history table with headers
 * @param {string} period - annual | quarter | ttm
 * @param {string} year - (e.g., 2025)
 * @return {number|array} - Single value or table of date and metric
 * @customfunction
 */
function DIVIDENDDATA_METRICS(symbol, metric, showHeaders = false, period = '', year = '') {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('Set your FMP API key via setApiKey()');

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
    if (!statement) throw new Error(`Invalid metric: ${metric}`);

    // Function to format headers to human-readable (camelCase/snake_case to Title Case with spaces)
    function formatHeader(key) {
      return key
        .replace(/([a-z])([A-Z])/g, '$1 $2')  // camelCase to space-separated
        .replace(/_/g, ' ')  // snake_case to space
        .replace(/\b\w/g, char => char.toUpperCase());  // Capitalize words
    }

    const properRawKey = metricToRawKey[loweredMetric];
    if (!properRawKey) throw new Error(`Invalid metric: ${metric}`);

    const formattedMetric = formatHeader(properRawKey);

    const fullData = DIVIDENDDATA_STATEMENT(symbol, statement, true, period, year);
    if (typeof fullData === 'string' && fullData.startsWith('Error')) return fullData;
    if (fullData.length === 0 || fullData[0].length === 0) return NaN;

    const headers = fullData[0];
    const dataRows = fullData.slice(1);
    if (dataRows.length === 0) return NaN;

    let colIndex = headers.indexOf(formattedMetric);
    if (colIndex === -1) throw new Error(`Metric ${metric} not found in statement`);

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
    return `Error: ${error.message}`;
  }
}



/**
 * Fetch a specific ratio or key metric: =DIVIDENDDATA_RATIOS("MSFT", "currentRatio") or =DIVIDENDDATA_RATIOS("MSFT", "peRatio", true)
 * @param {string} symbol - Ticker (e.g., MSFT)
 * @param {string} metric - The specific ratio or key metric (e.g., currentRatio, quickRatio, revenuePerShare)
 * @param {boolean} [showHeaders=false] - If true, return history table with headers
 * @param {string} period - annual | quarter | ttm
 * @param {string} year - (e.g., 2025)
 * @return {number|array} - Single value or table of date and metric
 * @customfunction
 */
function DIVIDENDDATA_RATIOS(symbol, metric, showHeaders = false, period = '', year = '') {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('Set your FMP API key via setApiKey()');

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
    if (!endpoint) throw new Error(`Invalid metric: ${metric}`);

    let properRawKey = metricToRawKey[loweredMetric];
    if (!properRawKey) throw new Error(`Invalid metric: ${metric}`);

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

    if (data.length === 0) return NaN;

    // Filter by year if provided
    if (year) {
      data = data.filter(row => row.calendarYear === year);
      if (data.length === 0) return NaN;
    }

    // Sort by date descending (latest first)
    data.sort((a, b) => new Date(b.date) - new Date(a.date));

    const rawHeaders = Object.keys(data[0]);

    if (!rawHeaders.includes(properRawKey)) throw new Error(`Metric ${metric} not found in data`);

    const formattedMetric = formatHeader(properRawKey.replace('TTM', ''));

    // Handle TTM special case (no 'date')
    if (loweredPeriod === 'ttm' && !rawHeaders.includes('date')) {
      if (data.length !== 1) throw new Error('Unexpected data structure for TTM');
      const value = data[0][properRawKey] || NaN;
      if (showHeaders) {
        return [['Period', formattedMetric], ['TTM', value]];
      } else {
        return value;
      }
    }

    // Normal case with date
    if (!rawHeaders.includes('date')) throw new Error('No date field in data');

    if (showHeaders) {
      let table = data.map(row => [new Date(row.date), row[properRawKey] || NaN]);
      table.unshift(['Date', formattedMetric]);
      return table;
    } else {
      return data[0][properRawKey] || NaN;
    }

  } catch (error) {
    return `Error: ${error.message}`;
  }

}





/**
 * Fetch a specific growth metric: =DIVIDENDDATA_GROWTH("MSFT", "revenuegrowth") or =DIVIDENDDATA_GROWTH("MSFT", "epsgrowth", true)
 * @param {string} symbol - Ticker (e.g., MSFT)
 * @param {string} metric - The specific growth metric (e.g., revenuegrowth, netincomegrowth)
 * @param {boolean} [showHeaders=false] - If true, return history table with headers
 * @param {string} period - annual | quarter | q1 | q2 | q3 | q4 | fy
 * @param {string} year - (e.g., 2025)
 * @return {number|array} - Single value or table of date and metric
 * @customfunction
 */
function DIVIDENDDATA_GROWTH(symbol, metric, showHeaders = false, period = '', year = '') {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('Set your FMP API key via setApiKey()');

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
    if (!properRawKey) throw new Error(`Invalid metric: ${metric}`);

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

    if (data.length === 0) return NaN;

    // Filter by year if provided
    if (year) {
      data = data.filter(row => row.fiscalYear === year);
      if (data.length === 0) return NaN;
    }

    // Sort by date descending (latest first)
    data.sort((a, b) => new Date(b.date) - new Date(a.date));

    const rawHeaders = Object.keys(data[0]);

    if (!rawHeaders.includes(properRawKey)) throw new Error(`Metric ${metric} not found in data`);

    if (showHeaders) {
      let table = data.map(row => [new Date(row.date), row[properRawKey] || NaN]);
      table.unshift(['Date', formattedMetric]);
      return table;
    } else {
      return data[0][properRawKey] || NaN;
    }

  } catch (error) {
    return `Error: ${error.message}`;
  }
}


/**
 * Fetch stock quote data: =DIVIDENDDATA_QUOTE("AAPL", "price") or =DIVIDENDDATA_QUOTE("AAPL", "full", , , TRUE) or =DIVIDENDDATA_QUOTE("AAPL", "history", "2025-06-10", "2025-09-10", TRUE)
 * @param {string} symbol - Ticker (e.g., AAPL)
 * @param {string} [metric="price"] - price | change | volume | full | history
 * @param {string} [fromDate] - Start date for history (YYYY-MM-DD, optional)
 * @param {string} [toDate] - End date for history (YYYY-MM-DD, optional)
 * @param {boolean} [showHeaders=true] - Include header row for 'full' or 'history'
 * @return {string|number|array} - Value or table
 * @customfunction
 */
function DIVIDENDDATA_QUOTE(symbol, metric = "price", fromDate, toDate, showHeaders = true) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('Set your FMP API key via setApiKey()');

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

      if (data.length === 0) return 'No data for the provided ticker symbol';

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

      if (data.length === 0) return 'No data for the provided ticker symbol';

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

      if (data.length === 0) return [['No historical data']];

      // Sort by date descending (latest first)
      data.sort((a, b) => new Date(b.date) - new Date(a.date));

      const rawHeaders = Object.keys(data[0]);
      const formattedHeaders = rawHeaders.map(formatHeader);

      let table = data.map(row => rawHeaders.map(key => {
        let value = row[key];
        if (key === 'date') {
          return new Date(value);
        }
        return value;
      }));

      if (actualShowHeaders) {
        table.unshift(formattedHeaders);
      }

      return table;
    } else {
      throw new Error(`Invalid metric: ${metric}. Use: price, change, volume, full, history`);
    }
  } catch (error) {
    return `Error: ${error.message}`;
  }
}


/**
 * Fetch company profile data: =DIVIDENDDATA_PROFILE("MSFT", "marketcap") or =DIVIDENDDATA_PROFILE("MSFT", "full", true)
 * @param {string} symbol - Ticker (e.g., SCHD)
 * @param {string} metric - symbol | price | marketcap | beta | lastdividend | range | change | changepercentage | volume | averagevolume | companyname | currency | cik | isin | cusip | exchangefullname | exchange | industry | website | description | ceo | sector | country | fulltimeemployees | phone | address | city | state | zip | image | ipodate | defaultimage | isetf | isactivelytrading | isadr | isfund | full
 * @param {boolean} [showHeaders=false] - Include header row for 'full'
 * @return {string|number|array} - Value or table
 * @customfunction
 */
function DIVIDENDDATA_PROFILE(symbol, metric, showHeaders = false) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('Set your FMP API key via setApiKey()');

  try {
    symbol = symbol.toUpperCase();
    const url = `https://financialmodelingprep.com/stable/profile?symbol=${symbol}&apikey=${apiKey}`;
    const content = fetchWithCache(url, 3600);
    const data = JSON.parse(content);

    if (data.length === 0) return 'No data for the provided ticker symbol';

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
      if (!rawKey) throw new Error(`Invalid metric: ${metric}`);

      let value = data[0][rawKey];
      if (rawKey.toLowerCase() === 'ipodate') {
        return value ? new Date(value) : '';
      }
      return value !== null ? value : '';
    }
  } catch (error) {
    return `Error: ${error.message}`;
  }
}



/**
 * Fetch ETF/Mutual Fund data: =DIVIDENDDATA_FUND("SPY", "holdings", true) or =DIVIDENDDATA_FUND("SPY", "expenseratio")
 * @param {string} symbol - Ticker (e.g., SPY)
 * @param {string} metric - holdings | countryweighting | symbol | name | description | isin | assetclass | securitycusip | domicile | website | etfcompany | expenseratio | assetsundermanagement | avgvolume | inceptiondate | nav | navcurrency | holdingscount | updatedat | sectorslist
 * @param {boolean} [showHeaders=false] - Include header row for tables (holdings, countryweighting, sectorslist)
 * @return {string|number|array} - Value or table
 * @customfunction
 */
function DIVIDENDDATA_FUND(symbol, metric, showHeaders = false) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('Set your FMP API key via setApiKey()');

  try {
    symbol = symbol.toUpperCase();
    const loweredMetric = metric.toLowerCase();

    let infoData;
    if (loweredMetric !== 'holdings' && loweredMetric !== 'countryweighting') {
      const infoUrl = `https://financialmodelingprep.com/stable/etf/info?symbol=${symbol}&apikey=${apiKey}`;
      const infoContent = fetchWithCache(infoUrl, 3600);
      infoData = JSON.parse(infoContent);

      if (infoData.length === 0) return 'No data for the provided ticker symbol';
    }

    switch (loweredMetric) {
      case 'holdings':
        const holdingsUrl = `https://financialmodelingprep.com/stable/etf/holdings?symbol=${symbol}&apikey=${apiKey}`;
        const holdingsContent = fetchWithCache(holdingsUrl, 3600);
        const holdingsData = JSON.parse(holdingsContent);

        if (holdingsData.length === 0) return [['No holdings data']];

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

        if (cwData.length === 0) return [['No country weighting data']];

        const cwTable = cwData.map(row => [row.country, row.weightPercentage]);

        if (showHeaders) {
          cwTable.unshift(['Country', 'Weight Percentage']);
        }

        return cwTable;

      case 'sectorslist':
        const sectors = infoData[0].sectorsList || [];
        if (sectors.length === 0) return [['No sectors data']];

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
        if (!rawKey) throw new Error(`Invalid metric: ${metric}`);

        let value = infoData[0][rawKey];

        if (loweredMetric === 'inceptiondate' || loweredMetric === 'updatedat') {
          return value ? new Date(value) : '';
        }

        return value !== undefined ? value : '';
    }
  } catch (error) {
    return `Error: ${error.message}`;
  }
}

/**
 * Fetch revenue segments data: =DIVIDENDDATA_SEGMENTS("AAPL", "products", "annual", "2024", true)
 * @param {string} symbol - Ticker (e.g., AAPL)
 * @param {string} metric - products | geographic
 * @param {string} [period=annual] - annual | quarter
 * @param {string} [year] - Filter to specific year (e.g., 2024)
 * @param {boolean} [showHeaders=false] - Include header row for table
 * @return {array} - Table of segments data
 * @customfunction
 */
function DIVIDENDDATA_SEGMENTS(symbol, metric, period = 'annual', year = '', showHeaders = false) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('Set your FMP API key via setApiKey()');

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
      throw new Error(`Invalid metric: ${metric}. Use: products or geographic`);
    }

    if (loweredPeriod !== 'annual' && loweredPeriod !== 'quarter') {
      throw new Error(`Invalid period: ${period}. Use: annual or quarter`);
    }

    const url = `https://financialmodelingprep.com/stable/${endpoint}?symbol=${symbol}&period=${loweredPeriod}&apikey=${apiKey}`;
    const content = fetchWithCache(url, 3600);
    let data = JSON.parse(content);

    if (data.length === 0) return [['No data for the provided ticker symbol']];

    // Filter by year if provided
    if (filterYear) {
      data = data.filter(item => String(item.fiscalYear) === filterYear);
      if (data.length === 0) return [['No data for specified year']];
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
    return `Error: ${error.message}`;
  }
}


/**
 * Fetch KPIs and segments data: =DIVIDENDDATA_KPIS("MSFT", "annual", "2025", TRUE)
 * @param {string} ticker - Ticker symbol (e.g., MSFT)
 * @param {string} [period=annual] - annual | quarterly
 * @param {string|number} [year] - Optional calendar year to filter
 * @param {boolean} [showheaders=TRUE] - Include header row
 * @return {array} - Table of KPIs and segments
 * @customfunction
 */
function DIVIDENDDATA_KPIS(ticker, period = 'annual', year = '', showheaders = true) {
  const fiscalApiKey = '025bb017-5cdd-4aef-a8fb-b36fecb3ea27';

  try {
    ticker = ticker.toUpperCase();
    const loweredPeriod = period.toLowerCase();
    if (loweredPeriod !== 'annual' && loweredPeriod !== 'quarterly') {
      throw new Error('Period must be "annual" or "quarterly"');
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
      return [['No data available'], ['Data Powered By Fiscal.ai']];
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
    return `Error: ${error.message}\nData Powered By Fiscal.ai`;
  }
}

/**
 * Fetch commodities data: =DIVIDENDDATA_COMMODITIES("CLUSD", "price") or =DIVIDENDDATA_COMMODITIES("GCUSD", "history", "2025-06-10", "2025-09-10", TRUE)
 * @param {string} symbol - Commodity symbol (e.g., CLUSD for Crude Oil)
 * @param {string} metric - list | price | fullquote | history
 * @param {string} [fromDate] - Start date for history (YYYY-MM-DD)
 * @param {string} [toDate] - End date for history (YYYY-MM-DD)
 * @param {boolean} [showHeaders] - Include header row for tables (list, fullquote, history)
 * @return {number|string|array} - Value or table
 * @customfunction
 */
function DIVIDENDDATA_COMMODITIES(symbol, metric, fromDate, toDate, showHeaders) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('Set your FMP API key via setApiKey()');

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

    // Set default showHeaders based on metric if not provided
    if (actualShowHeaders === undefined) {
      actualShowHeaders = (loweredMetric === 'list');
    }

    switch (loweredMetric) {
      case 'list':
        url = `https://financialmodelingprep.com/stable/commodities-list?apikey=${apiKey}`;
        content = fetchWithCache(url, 86400);  // 24 hours
        data = JSON.parse(content);

        if (data.length === 0) return [['No commodities list data']];

        const rawHeaders = Object.keys(data[0]);
        const formattedHeaders = rawHeaders.map(formatHeader);

        let table = data.map(row => rawHeaders.map(key => row[key] || ''));

        if (actualShowHeaders) {
          table.unshift(formattedHeaders);
        }

        return table;

      case 'price':
      case 'fullquote':
        if (!symbol) throw new Error('Symbol required for price or fullquote');
        symbol = symbol.toUpperCase();
        url = `https://financialmodelingprep.com/stable/quote?symbol=${symbol}&apikey=${apiKey}`;
        content = fetchWithCache(url, 60);
        data = JSON.parse(content);

        if (data.length === 0) return 'No data for the provided symbol';

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
        if (!symbol) throw new Error('Symbol required for history');
        symbol = symbol.toUpperCase();
        url = `https://financialmodelingprep.com/stable/historical-price-eod/light?symbol=${symbol}&apikey=${apiKey}`;
        if (actualFromDate) url += `&from=${actualFromDate}`;
        if (actualToDate) url += `&to=${actualToDate}`;
        content = fetchWithCache(url, 300);
        data = JSON.parse(content);

        if (data.length === 0) return [['No historical data']];

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
        throw new Error(`Invalid metric: ${metric}. Use: list, price, fullquote, history`);
    }
  } catch (error) {
    return `Error: ${error.message}`;
  }
}

/**
 * Fetch batch stock quote data: =DIVIDENDDATA_QUOTE_BATCH("AAPL,MSFT,KMB,WBA", "price,change", TRUE)
 * @param {string} symbols - Comma-separated tickers (e.g., AAPL,MSFT)
 * @param {string} [metrics="all"] - Comma-separated metrics: price,change,volume or "all"
 * @param {boolean} [showHeaders] - Include header row and symbol column (defaults to TRUE if metrics="all", else FALSE)
 * @return {array} - Table or column(s) of quote data
 * @customfunction
 */
function DIVIDENDDATA_QUOTE_BATCH(symbols, metrics = "all", showHeaders) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('Set your FMP API key via setApiKey()');

  try {
    if (!symbols) throw new Error('Symbols required');

    // Process symbols
    const symbolArray = symbols.split(',').map(s => s.trim().toUpperCase());
    const symbolList = symbolArray.join(',');

    const url = `https://financialmodelingprep.com/stable/batch-quote-short?symbols=${symbolList}&apikey=${apiKey}`;
    const content = fetchWithCache(url, 60);
    const data = JSON.parse(content);

    if (data.length === 0) return [['No data for the provided symbols']];

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
          throw new Error(`Invalid metric: ${m}. Use: price, change, volume or "all"`);
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
    return `Error: ${error.message}`;
  }
}



function getFinchatToken(uid, userTier) {
  const endpoint = 'https://api.finchat.io/auth/generate-token';
  const payload = { payload: { uid: uid, userTier: userTier } };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + PropertiesService.getScriptProperties().getProperty('FINCHAT_MASTER_API_KEY') },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true  // Add this to get full error response
  };
  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    const code = response.getResponseCode();
    const text = response.getContentText();
    if (code !== 200) {
      throw new Error('Token generation failed: HTTP ' + code + ' - ' + text);
    }
    Logger.log('Generated Token: ' + text);  // Log the token for debugging
    return text;
  } catch (e) {
    Logger.log('Token Error: ' + e.message);  // Log for debugging
    throw e;  // Propagate error to DIVIDENDDATA_AI
  }
}

// Preset prompt mappings (customize as needed)
const PROMPT_MAP = {
  'Earnings Summary': 'Provide a detailed earnings summary for {ticker}, including recent quarterly results, YoY growth, and analyst expectations.',
  'Bull/Bear': 'Give a balanced bull and bear case analysis for {ticker}, including key risks and opportunities.',
  'News': 'Summarize the latest news and events for {ticker}, focusing on the past month.',
  // Add more presets here, e.g., 'Dividend Analysis': 'Analyze the dividend history and yield for {ticker}.'
};

/**
 * Custom function for Sheets: =DIVIDENDDATA_AI(ticker, promptType)
 * @param {string} ticker - Stock ticker symbol (e.g., "AAPL")
 * @param {string} promptType - Preset prompt type (e.g., "Earnings Summary")
 * @return {string} AI-generated response
 * @customfunction
 */
function DIVIDENDDATA_AI(ticker, promptType) {
  if (!ticker || !promptType) {
    return 'Error: Provide ticker and prompt type.';
  }

  const promptTemplate = PROMPT_MAP[promptType];
  if (!promptTemplate) {
    return 'Error: Invalid prompt type. Options: ' + Object.keys(PROMPT_MAP).join(', ');
  }

  const query = promptTemplate.replace('{ticker}', ticker.toUpperCase());

  // Dynamically get the user's email as UID (like in app.R)
  const userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    return 'Error: Unable to retrieve user email. Grant permission or check scopes.';
  }
  Logger.log('Using UID (email): ' + userEmail);  // Log for debugging

  // Get token using the email as UID
  const apiKey = getFinchatToken(userEmail, "paid");  // Assume "paid" tier; make dynamic if needed (e.g., check user properties)

  const url = 'https://api.finchat.io/v1/query';
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
    },
    payload: JSON.stringify({ query: query }),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const text = response.getContentText();
    Logger.log('Query Response Code: ' + code + ' - Text: ' + text);  // Log for debugging

    if (code !== 200) {
      return 'API Error: HTTP ' + code + ' - ' + text;  // Return plain-text errors directly
    }

    const json = JSON.parse(text);
    if (json.answer) {
      return json.answer;  // Or json.answer.text if nested
    } else {
      return 'Error: Unexpected JSON structure - ' + JSON.stringify(json);
    }
  } catch (e) {
    return 'Fetch Error: ' + e.message;
  }
}






// Add-on menu
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Open AI Copilot Sidebar', 'showCopilotSidebar')
    .addToUi();
}

// Show sidebar with embedded Finchat
function showCopilotSidebar() {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          #finchat {
            border: none;
            width: 100%;
            height: 100vh;
            border-radius: 15px;
            box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.2), 0 6px 20px 0 rgba(0, 0, 0, 0.19);
          }
        </style>
      </head>
      <body>
        <iframe id="finchat" name="finchat" src="about:blank"></iframe>
        <script>
          const embedKey = '81400b13912e4d75b3cff58028e4f8cc';  // From app.R
          const iframe = document.getElementById('finchat');
          iframe.src = 'https://enterprise.finchat.io/' + embedKey;

          // Generate token (adapt from app.R; call a server-side function if needed)
          const token = google.script.run.withSuccessHandler(function(token) {
            function handleIframeReady(event) {
              if (event.data !== 'READY') return;
              const targetWindow = iframe.contentWindow;
              if (!targetWindow) return;
              targetWindow.postMessage({ token }, 'https://enterprise.finchat.io/');
            }
            window.addEventListener('message', handleIframeReady);
          }).getFinchatToken('support@dividenddata.com', 'paid');  // Replace with actual uid/tier
        </script>
      </body>
    </html>
  `).setTitle('Dividend Data AI Copilot').setWidth(800);

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
Dummy sidebar function (can be expanded later)
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

*/


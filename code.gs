const IndicatorsPositions = {
  "CDI" :  {"Row" : 5, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4},
  "MayerMultiple" :  {"Row" : 8, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4} 
}

const ConsolidatedExchangeTradedAssetsPositions = {
  "RowStart" : 3,
  "Ticker" : 1,
  "NumberOfShares" : 2,
  "AveragePrice" : 3,
  "TotalAmount" : 5,
  "Type" :11,
  "Subtype" :12,
  "Currency" : 13
}

const FisPositions = {
  "Earnings" : {"RowStart" : 3, "Ticker" : 1, "NumberOfShares" : 4},
}

const AcoesPositions = {
  "Earnings" : {"RowStart" : 3, "Ticker" : 1, "NumberOfShares" : 2},
}

const CryptocurrencyPositions = {
  "Earnings" : {"RowStart" : 3, "Ticker" : 1, "NumberOfShares" : 2},
}

const REITsPositions = {
  "Earnings" : {"RowStart" : 3, "Ticker" : 1, "NumberOfShares" : 2},
}

const StocksPositions = {
  "Earnings" : {"RowStart" : 3, "Ticker" : 1, "NumberOfShares" : 2},
}

const DividendRanges = {
  "BR" : {"Date": "Dividends(R$)!A3:A", "Tickers" : "Dividends(R$)!B3:B", "Dividends" : "Dividends(R$)!F3:F"},
  "US" : {"Date": "Dividends($)!A3:A","Tickers" : "Dividends($)!B3:B", "Dividends" : "Dividends($)!F3:F"},
}

const ExchangeTradedAssetsRanges = {
  
  "ExchangeTradedAssets" : {
    "Tickers" : "Exchange Traded Assets!A3:A",
    "NumbersOfShares" : "Exchange Traded Assets!B3:B",
    "AveragePrices" : "Exchange Traded Assets!C3:C",
    "FinancialInstitutions" : "Exchange Traded Assets!D3:D",
    "Currencys" : "Exchange Traded Assets!E3:E"
  },

  "ConsolidatedExchangeTradedAssets" : {
    "Tickers" : "Consolidated Exchange Traded Assets!A3:A",
    "Types" : "Consolidated Exchange Traded Assets!K3:K",
    "Subtypes" : "Consolidated Exchange Traded Assets!L3:L"
  },

  "FI" : {
    "Tickers" : "FI!A3:A",
    "Activitys" : "FI!B3:B",
    "Amounts" : "FI!F3:F"
  }

}

const B3apiBaseUrl = 'https://sistemaswebb3-balcao.b3.com.br'

var MyAssets = {}

function getCDI(input) {

  let Url = B3apiBaseUrl + '/featuresDIProxy/DICall/GetRateDI/eyJsYW5ndWFnZSI6InB0LWJyIn0='
  let Response = UrlFetchApp.fetch(Url)
  let Sheet = SpreadsheetApp.getActive().getSheetByName("Investments")

  if(Response.getResponseCode() == 200 && Response.getContentText() != ""){

    let Data = JSON.parse(Response.getContentText())

    if(typeof Data.rate !== 'undefined' && typeof Data.date !== 'undefined'){
      Sheet.getRange(IndicatorsPositions["CDI"]["Row"], IndicatorsPositions["CDI"]["RateCol"]).setValue(Data.rate + '%')
      Sheet.getRange(IndicatorsPositions["CDI"]["Row"], IndicatorsPositions["CDI"]["DataCol"]).setValue(Data.date)
      Sheet.getRange(IndicatorsPositions["CDI"]["Row"], IndicatorsPositions["CDI"]["UpdateCol"]).setValue(new Date())
      return
    }

  }

  Sheet.getRange(IndicatorsPositions["CDI"]["Row"], IndicatorsPositions["CDI"]["RateCol"]).setValue(0 + '%')
  Sheet.getRange(IndicatorsPositions["CDI"]["Row"], IndicatorsPositions["CDI"]["DataCol"]).setValue(new Date())
  Sheet.getRange(IndicatorsPositions["CDI"]["Row"], IndicatorsPositions["CDI"]["UpdateCol"]).setValue(new Date())

}


function calculateMayerMultiple(Data) {

  let Sum = 0

  for (let I = 0; I < Data.length; I++) {
    Sum = Sum + Data[I]['price']
  }

  return Data[199]['price']/(Sum/200) //current value divided by the average of the last 200 days
  
}

function getMayerMultipleColor(MayerMultiple) {

  if(MayerMultiple < 1){
    return 'green'
  } if(MayerMultiple < 2.4){
    return 'yellow'
  } else {
    return 'red'
  }
  
}

function getMayerMultiple(input) {
  
  let Url = 'https://firebasestorage.googleapis.com/v0/b/buy-bitcoin-worldwide.appspot.com/o/prices%2Fhistoric%2Fbtc%2Fusd.json?alt=media'
  let Response = UrlFetchApp.fetch(Url)
  let Sheet = SpreadsheetApp.getActive().getSheetByName("Investments")

  if(Response.getResponseCode() == 200 && Response.getContentText() != ""){

    let Data = JSON.parse(Response.getContentText())

    if(Data.length > 0 && typeof Data[0]['date'] !== 'undefined' && typeof Data[0]['price'] !== 'undefined'){

      Data = Data.slice(-200)
      let UpdatedAt = new Date(Data[199]['date']).toLocaleDateString('pt-BR')
      let MayerMultiple = calculateMayerMultiple(Data)
      let MayerMultipleColor = getMayerMultipleColor(MayerMultiple)

      if(typeof MayerMultiple !== 'undefined' && typeof UpdatedAt !== 'undefined'){
        Sheet.getRange(IndicatorsPositions["MayerMultiple"]["Row"], IndicatorsPositions["MayerMultiple"]["RateCol"]).setValue(MayerMultiple)
        Sheet.getRange(IndicatorsPositions["MayerMultiple"]["Row"], IndicatorsPositions["MayerMultiple"]["RateCol"]).setBackground(MayerMultipleColor)
        Sheet.getRange(IndicatorsPositions["MayerMultiple"]["Row"], IndicatorsPositions["MayerMultiple"]["DataCol"]).setValue(UpdatedAt)
        Sheet.getRange(IndicatorsPositions["MayerMultiple"]["Row"], IndicatorsPositions["MayerMultiple"]["UpdateCol"]).setValue(new Date())
        return
      }

    }

  }

  Sheet.getRange(IndicatorsPositions["MayerMultiple"]["Row"], IndicatorsPositions["MayerMultiple"]["RateCol"]).setValue(0)
  Sheet.getRange(IndicatorsPositions["MayerMultiple"]["Row"], IndicatorsPositions["MayerMultiple"]["DataCol"]).setValue(new Date())
  Sheet.getRange(IndicatorsPositions["MayerMultiple"]["Row"], IndicatorsPositions["MayerMultiple"]["UpdateCol"]).setValue(new Date())

}

function updateIndicators(input) {
  getCDI()
  getMayerMultiple()

  /*

    The indicators below have been replaced by alternative ways of obtaining information

    IPCA
    =CONCATENATE(REGEXEXTRACT(INDEX(IMPORTHTML("https://www.melhorcambio.com/ipca";"table";1);2;2);"[0-9]+[,.]+[0-9]+");"%")

    SELIC
    =CDI+0,1%

    Dollar-Real
    =GOOGLEFINANCE("CURRENCY:USDBRL")

    Bitcoin-Dollar
    =GOOGLEFINANCE("CURRENCY:BTCUSD")
  
  */
}

function reloadFormula(Sheet, Row, Col) {

  let CellarRange = Sheet.getRange(Row, Col);
  let Formulas = CellarRange.getFormulas();
  CellarRange.clearContent();
  CellarRange.setFormulas(Formulas);

}


function updateFinancialDevelopment(input){

  updateIndicators()
  SpreadsheetApp.flush();

  let Sheet = SpreadsheetApp.getActive().getSheetByName("Development")
  let NetWorth = SpreadsheetApp.getActive().getRange("Retirement!A11").getValue()

  Sheet.insertRowBefore(1)

  Sheet.getRange(1, 1).setValue(new Date())
  Sheet.getRange(1, 2).setValue(NetWorth)

}

function parseColumns(Data){

  let Result = []

  for (let I = 0; I < Data.length; I++) {

    if(Data[I][0] == ''){
      break
    }

    Result.push(Data[I][0]);

  }

  return Result

}

function addDividendsByTicker(Ticker, Country){

  Ticker = Ticker.trim()

  let TickersData = SpreadsheetApp.getActive().getRange(DividendRanges[Country]["Tickers"]).getValues();
  let Tickers = parseColumns(TickersData);

  let DividendsData = SpreadsheetApp.getActive().getRange(DividendRanges[Country]["Dividends"]).getValues();
  let Dividends = parseColumns(DividendsData);

  let Sum = 0;

  for (let I = 0; I < Tickers.length; I++) {

    if(Tickers[I].trim() != Ticker){
      continue;
    }

    Sum = Sum + Dividends[I];

  }

  return Sum;

}

function addDividendsByMonth(LineDate, Country){

  let CurrentDate = new Date(LineDate);

  let DatesData = SpreadsheetApp.getActive().getRange(DividendRanges[Country]["Date"]).getValues();
  let Dates = parseColumns(DatesData);

  let DividendsData = SpreadsheetApp.getActive().getRange(DividendRanges[Country]["Dividends"]).getValues();
  let Dividends = parseColumns(DividendsData);

  let Sum = 0;

  for (let I = 0; I < Dates.length; I++) {

    DividendDate = new Date(Dates[I])

    if(DividendDate.getFullYear() != CurrentDate.getFullYear()){
      continue;
    }

    if(DividendDate.getMonth() != CurrentDate.getMonth()){
      continue;
    }

    Sum = Sum + Dividends[I];

  }

  return Sum;

}

function consolidateExchangeTradedAssets(){

  let TickersData = SpreadsheetApp.getActive().getRange(ExchangeTradedAssetsRanges["ExchangeTradedAssets"]["Tickers"]).getValues();
  let Tickers = parseColumns(TickersData);

  let NumbersOfSharesData = SpreadsheetApp.getActive().getRange(ExchangeTradedAssetsRanges["ExchangeTradedAssets"]["NumbersOfShares"]).getValues();
  let NumbersOfShares = parseColumns(NumbersOfSharesData);

  let AveragePricesData = SpreadsheetApp.getActive().getRange(ExchangeTradedAssetsRanges["ExchangeTradedAssets"]["AveragePrices"]).getValues();
  let AveragePrices = parseColumns(AveragePricesData);

  let CurrencysData = SpreadsheetApp.getActive().getRange(ExchangeTradedAssetsRanges["ExchangeTradedAssets"]["Currencys"]).getValues();
  let Currencys = parseColumns(CurrencysData);

  let ConsolidateExchangeTradedAssets = {};

  for (let I = 0; I < Tickers.length; I++) {

    if (typeof ConsolidateExchangeTradedAssets[Tickers[I]] == 'undefined'){

      ConsolidateExchangeTradedAssets[Tickers[I]] = {};
    
    }

    if (typeof ConsolidateExchangeTradedAssets[Tickers[I]]["NumberOfShares"] == 'undefined'){

      ConsolidateExchangeTradedAssets[Tickers[I]]["NumberOfShares"] = [];

    }
    
    ConsolidateExchangeTradedAssets[Tickers[I]]["NumberOfShares"].push(NumbersOfShares[I]);

    if (typeof ConsolidateExchangeTradedAssets[Tickers[I]]["AveragePrice"] == 'undefined'){

      ConsolidateExchangeTradedAssets[Tickers[I]]["AveragePrice"] = [];

    }
    
    ConsolidateExchangeTradedAssets[Tickers[I]]["AveragePrice"].push(AveragePrices[I]);

    if (typeof ConsolidateExchangeTradedAssets[Tickers[I]]["Currency"] == 'undefined'){

      ConsolidateExchangeTradedAssets[Tickers[I]]["Currency"] = Currencys[I];

    }

  }

  for (const [Key, Value] of Object.entries(ConsolidateExchangeTradedAssets)) {

    if(ConsolidateExchangeTradedAssets[Key]["NumberOfShares"].length > 1){

      let Sum = 0;

      for (let I = 0; I < ConsolidateExchangeTradedAssets[Key]["NumberOfShares"].length; I++ ) {

        Sum += ConsolidateExchangeTradedAssets[Key]["NumberOfShares"][I];

      }

      let Multiplier = 0;
      let TotalAveragePrice = 0;

      for (let I = 0; I < ConsolidateExchangeTradedAssets[Key]["NumberOfShares"].length; I++ ) {

        Multiplier = ConsolidateExchangeTradedAssets[Key]["NumberOfShares"][I]/Sum;
        TotalAveragePrice = TotalAveragePrice + (ConsolidateExchangeTradedAssets[Key]["AveragePrice"][I] * Multiplier);

      }

      ConsolidateExchangeTradedAssets[Key]["NumberOfShares"] = Sum;
      ConsolidateExchangeTradedAssets[Key]["AveragePrice"] = TotalAveragePrice;

    } else {

      ConsolidateExchangeTradedAssets[Key]["NumberOfShares"] = ConsolidateExchangeTradedAssets[Key]["NumberOfShares"][0];
      ConsolidateExchangeTradedAssets[Key]["AveragePrice"] = ConsolidateExchangeTradedAssets[Key]["AveragePrice"][0];
    
    }

  }

  return ConsolidateExchangeTradedAssets;

}

function printConsolidateExchangeTradedAssets(ConsolidateExchangeTradedAssets){

  let Sheet = SpreadsheetApp.getActive().getSheetByName("Consolidated Exchange Traded Assets");
  let RowIndex = ConsolidatedExchangeTradedAssetsPositions["RowStart"];

  for (const [Key, Value] of Object.entries(ConsolidateExchangeTradedAssets)) {

    Sheet.getRange(RowIndex, ConsolidatedExchangeTradedAssetsPositions["Ticker"]).setValue(Key);

    for (const [Index, Info] of Object.entries(Value)) {

      Sheet.getRange(RowIndex, ConsolidatedExchangeTradedAssetsPositions[Index]).setValue(Info);

    }

    RowIndex++;

  }

}

function updateConsolidateExchangeTradedAssets(){

  let ConsolidateExchangeTradedAssets = consolidateExchangeTradedAssets();
  printConsolidateExchangeTradedAssets(ConsolidateExchangeTradedAssets);
  ConsolidateExchangeTradedAssets = updateassetAssetDetails(ConsolidateExchangeTradedAssets);
  printByType(ConsolidateExchangeTradedAssets, FisPositions, 'FI')
  printByType(ConsolidateExchangeTradedAssets, AcoesPositions, 'Ações')
  printByType(ConsolidateExchangeTradedAssets, CryptocurrencyPositions, 'Cryptocurrency')
  printByType(ConsolidateExchangeTradedAssets, REITsPositions, 'REIT')
  printByType(ConsolidateExchangeTradedAssets, StocksPositions, 'Stocks')

}

function updateassetAssetDetails(ConsolidateExchangeTradedAssets){

  let Data = SpreadsheetApp.getActive().getRange(ExchangeTradedAssetsRanges["ConsolidatedExchangeTradedAssets"]["Tickers"]).getValues();
  let Tickers = parseColumns(Data);

  Data = SpreadsheetApp.getActive().getRange(ExchangeTradedAssetsRanges["ConsolidatedExchangeTradedAssets"]["Types"]).getValues();
  let Types = parseColumns(Data);

  Data = SpreadsheetApp.getActive().getRange(ExchangeTradedAssetsRanges["ConsolidatedExchangeTradedAssets"]["Subtypes"]).getValues();
  let Subtypes = parseColumns(Data);

  let I=0;

  for (let I = 0; I < Tickers.length; I++) {

    if(Types[I] == ''){
      continue;
    }

    if(Types[I] == '-'){
      continue;
    }

    ConsolidateExchangeTradedAssets[Tickers[I]]["Type"] = Types[I].trim();

    if(Subtypes[I] == ''){
      continue;
    }

    if(Subtypes[I] == '-'){
      continue;
    }

    ConsolidateExchangeTradedAssets[Tickers[I]]["Subtype"] = Subtypes[I].trim();

  }

  return ConsolidateExchangeTradedAssets;

}

function printByType(ConsolidateExchangeTradedAssets, Positions, Type){

  let Sheet = SpreadsheetApp.getActive().getSheetByName(Type);
  let RowIndex = Positions['Earnings']["RowStart"];

  for (const [Key, Value] of Object.entries(ConsolidateExchangeTradedAssets)) {

    if((ConsolidateExchangeTradedAssets[Key]['Type'] && ConsolidateExchangeTradedAssets[Key]['Type'] == Type) || (ConsolidateExchangeTradedAssets[Key]['Subtype'] && ConsolidateExchangeTradedAssets[Key]['Subtype'] == Type)){
      
      Sheet.getRange(RowIndex, Positions['Earnings']["Ticker"]).setValue(Key);
      Sheet.getRange(RowIndex, Positions['Earnings']["NumberOfShares"]).setValue(ConsolidateExchangeTradedAssets[Key]["NumberOfShares"]);
      RowIndex++;

    }

  }

}

function onEdit(){  //running with each change in the spreadsheet

  let ActiveCell = SpreadsheetApp.getActive().getActiveCell();
  let Reference = ActiveCell.getA1Notation();
  let SheetName = ActiveCell.getSheet().getName();
  let CellValue = ActiveCell.getValue();

  if(SheetName == 'Exchange Traded Assets' && Reference == 'F3' && CellValue == true){

    updateConsolidateExchangeTradedAssets()

  }

}
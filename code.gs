const IndicatorsPositions = {
  "CDI" :  {"Row" : 5, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4},
  "MayerMultiple" :  {"Row" : 8, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4} 
}

const DividendRanges = {
  "BR" : {"Date": "Dividends(R$)!A3:A", "Tickers" : "Dividends(R$)!B3:B", "Dividends" : "Dividends(R$)!D3:D"},
  "US" : {"Date": "Dividends($)!A3:A","Tickers" : "Dividends($)!B3:B", "Dividends" : "Dividends($)!F3:F"},
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

        Sheet.getRange(IndicatorsPositions["Bitcoin-Dollar"]["Row"], IndicatorsPositions["MayerMultiple"]["RateCol"]).setValue(Data[199]['price'])
        Sheet.getRange(IndicatorsPositions["Bitcoin-Dollar"]["Row"], IndicatorsPositions["MayerMultiple"]["DataCol"]).setValue(UpdatedAt)
        Sheet.getRange(IndicatorsPositions["Bitcoin-Dollar"]["Row"], IndicatorsPositions["MayerMultiple"]["UpdateCol"]).setValue(new Date())
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

    Dolar-Real
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
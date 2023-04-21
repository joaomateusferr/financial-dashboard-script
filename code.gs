const IndicatorsPositions = {
  "Dolar-Real" : {"Row" : 3, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4},
  "SLIC" : {"Row" : 4, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4},
  "CDI" :  {"Row" : 5, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4},
  "Inflation" :  {"Row" : 6, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4},
  "Bitcoin-Dollars" :  {"Row" : 7, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4},
  "MayerMultiple" :  {"Row" : 8, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4} 
}

const B3apiBaseUrl = 'https://sistemaswebb3-balcao.b3.com.br'
const BrapiBaseUrl = 'https://brapi.dev/api/v2'
const BrapiBaseV1Url = 'https://brapi.dev/api'

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

function getSLIC(input) {
  
  let CurrentDate = new Date().toLocaleDateString('pt-BR');
  let Url = BrapiBaseUrl + '/prime-rate?country=brazil&historical=true&start=' + CurrentDate + '&end=' + CurrentDate + '&sortBy=date&sortOrder=desc'
  let Response = UrlFetchApp.fetch(Url)
  let Sheet = SpreadsheetApp.getActive().getSheetByName("Investments")

  if(Response.getResponseCode() == 200 && Response.getContentText() != ""){

    let Data = JSON.parse(Response.getContentText())

    if(typeof Data['prime-rate'][0] !== 'undefined'){
      
      Data = Data['prime-rate'][0]
      Data.value = Data.value.replace('.', ',')

      if(typeof Data.value !== 'undefined' && typeof Data.date !== 'undefined'){
        Sheet.getRange(IndicatorsPositions["SLIC"]["Row"], IndicatorsPositions["SLIC"]["RateCol"]).setValue(Data.value + '%')
        Sheet.getRange(IndicatorsPositions["SLIC"]["Row"], IndicatorsPositions["SLIC"]["DataCol"]).setValue(Data.date)
        Sheet.getRange(IndicatorsPositions["SLIC"]["Row"], IndicatorsPositions["SLIC"]["UpdateCol"]).setValue(new Date())
        return
      }

    }

  }

  Sheet.getRange(IndicatorsPositions["SLIC"]["Row"], IndicatorsPositions["SLIC"]["RateCol"]).setValue(0 + '%')
  Sheet.getRange(IndicatorsPositions["SLIC"]["Row"], IndicatorsPositions["SLIC"]["DataCol"]).setValue(new Date())
  Sheet.getRange(IndicatorsPositions["SLIC"]["Row"], IndicatorsPositions["SLIC"]["UpdateCol"]).setValue(new Date())

}

function getInflation(input) {
  
  let CurrentDate = new Date().toLocaleDateString('pt-BR');
  let Url = BrapiBaseUrl + '/inflation?country=brazil&historical=false&start=' + CurrentDate + '&end=' + CurrentDate + '&sortBy=date&sortOrder=desc'
  let Response = UrlFetchApp.fetch(Url)
  let Sheet = SpreadsheetApp.getActive().getSheetByName("Investments")

  if(Response.getResponseCode() == 200 && Response.getContentText() != ""){

    let Data = JSON.parse(Response.getContentText())

    if(typeof Data['inflation'][0] !== 'undefined'){
      
      Data = Data['inflation'][0]
      Data.value = Data.value.replace('.', ',')

      if(typeof Data.value !== 'undefined' && typeof Data.date !== 'undefined'){
        Sheet.getRange(IndicatorsPositions["Inflation"]["Row"], IndicatorsPositions["Inflation"]["RateCol"]).setValue(Data.value + '%')
        Sheet.getRange(IndicatorsPositions["Inflation"]["Row"], IndicatorsPositions["Inflation"]["DataCol"]).setValue(Data.date)
        Sheet.getRange(IndicatorsPositions["Inflation"]["Row"], IndicatorsPositions["Inflation"]["UpdateCol"]).setValue(new Date())
        return
      }

    }

  }

  Sheet.getRange(IndicatorsPositions["Inflation"]["Row"], IndicatorsPositions["Inflation"]["RateCol"]).setValue(0 + '%')
  Sheet.getRange(IndicatorsPositions["Inflation"]["Row"], IndicatorsPositions["Inflation"]["DataCol"]).setValue(new Date())
  Sheet.getRange(IndicatorsPositions["Inflation"]["Row"], IndicatorsPositions["Inflation"]["UpdateCol"]).setValue(new Date())

}

function getDolarRealRate(input) {
  
  let Url = BrapiBaseUrl + '/currency?currency=USD-BRL'
  let Response = UrlFetchApp.fetch(Url)
  let Sheet = SpreadsheetApp.getActive().getSheetByName("Investments")

  if(Response.getResponseCode() == 200 && Response.getContentText() != ""){

    let Data = JSON.parse(Response.getContentText())

    if(typeof Data['currency'][0] !== 'undefined'){

      Data = Data['currency'][0]    
      Data.askPrice = parseFloat(Data.askPrice).toFixed(2)

      Data.askPrice = Data.askPrice.toString().replace('.', ',')
      let UpdateDate = new Date(Number(Data.updatedAtTimestamp)*1000)
      let UpdatedAt = UpdateDate.getDate() + "/" + (UpdateDate.getMonth()+1) + "/" + UpdateDate.getFullYear()

      if(typeof Data.askPrice !== 'undefined' && typeof UpdatedAt !== 'undefined'){
        Sheet.getRange(IndicatorsPositions["Dolar-Real"]["Row"], IndicatorsPositions["Dolar-Real"]["RateCol"]).setValue(Data.askPrice)
        Sheet.getRange(IndicatorsPositions["Dolar-Real"]["Row"], IndicatorsPositions["Dolar-Real"]["DataCol"]).setValue(UpdatedAt)
        Sheet.getRange(IndicatorsPositions["Dolar-Real"]["Row"], IndicatorsPositions["Dolar-Real"]["UpdateCol"]).setValue(new Date())
        return
      }

    }

  }

  Sheet.getRange(IndicatorsPositions["Dolar-Real"]["Row"], IndicatorsPositions["Dolar-Real"]["RateCol"]).setValue(0)
  Sheet.getRange(IndicatorsPositions["Dolar-Real"]["Row"], IndicatorsPositions["Dolar-Real"]["DataCol"]).setValue(new Date())
  Sheet.getRange(IndicatorsPositions["Dolar-Real"]["Row"], IndicatorsPositions["Dolar-Real"]["UpdateCol"]).setValue(new Date())

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

        Sheet.getRange(IndicatorsPositions["Bitcoin-Dollars"]["Row"], IndicatorsPositions["MayerMultiple"]["RateCol"]).setValue(Data[199]['price'])
        Sheet.getRange(IndicatorsPositions["Bitcoin-Dollars"]["Row"], IndicatorsPositions["MayerMultiple"]["DataCol"]).setValue(UpdatedAt)
        Sheet.getRange(IndicatorsPositions["Bitcoin-Dollars"]["Row"], IndicatorsPositions["MayerMultiple"]["UpdateCol"]).setValue(new Date())
        return
      }

    }

  }

  Sheet.getRange(IndicatorsPositions["MayerMultiple"]["Row"], IndicatorsPositions["MayerMultiple"]["RateCol"]).setValue(0)
  Sheet.getRange(IndicatorsPositions["MayerMultiple"]["Row"], IndicatorsPositions["MayerMultiple"]["DataCol"]).setValue(new Date())
  Sheet.getRange(IndicatorsPositions["MayerMultiple"]["Row"], IndicatorsPositions["MayerMultiple"]["UpdateCol"]).setValue(new Date())

}

function updateIndicators(input) {
  getDolarRealRate()
  getSLIC()
  getCDI()
  getInflation()
}

function parseAssets(Data){

  let Assets = []

  for (let I = 0; I < Data.length; I++) {

    if(Data[I][0] == ''){
      break
    }

    Assets.push(Data[I][0]);

  }

  return Assets

}

function getAssets(input){

  let FisData = SpreadsheetApp.getActive().getRange("Investments!Y3:Y").getValues()
  let Fis = parseAssets(FisData)

  let J = 3

  for (let I = 0; I < Fis.length; I++) {

    MyAssets[Fis[I]] = {}
    MyAssets[Fis[I]]['Line'] = J
    MyAssets[Fis[I]]['Value'] = 27
    MyAssets[Fis[I]]['Earnings'] = 29

    J = J + 1

  }

  let StocksData = SpreadsheetApp.getActive().getRange("Investments!O3:O").getValues()
  let Stocks = parseAssets(StocksData)

  J = 3

  for (let I = 0; I < Stocks.length; I++) {

    MyAssets[Stocks[I]] = {}
    MyAssets[Stocks[I]]['Line'] = J
    MyAssets[Stocks[I]]['Value'] = 17
    MyAssets[Stocks[I]]['Earnings'] = 19

    J = J + 1

  }

  let Assets = Fis.concat(Stocks)
  return Assets

}

function getAssetsString(Assets){

  let AssetsString = ''

  for (let I = 0; I < Assets.length; I++) {

    if(AssetsString != ''){

      AssetsString = AssetsString + '%2C'

    }

    AssetsString = AssetsString + Assets[I]

  }

  return AssetsString

}

function getAssetsInfo(input){
  
  let Assets = getAssets()
  let AssetsString = getAssetsString(Assets)
  let Url = BrapiBaseV1Url + '/quote/'+ AssetsString +'?range=1d&interval=1d&fundamental=false&dividends=true'
  let AssetDetails = {}
  let AssetResult = {}
  let Response = UrlFetchApp.fetch(Url)
  let Sheet = SpreadsheetApp.getActive().getSheetByName("Investments")

  if(Response.getResponseCode() == 200 && Response.getContentText() != ''){

    let Data = JSON.parse(Response.getContentText())

    if(typeof Data['results'] !== 'undefined'){

      Data = Data['results']

      for (let I = 0; I < Data.length; I++) {

        AssetDetails = MyAssets[Data[I]['symbol']]

        AssetResult['Value'] = 0

        if(Data[I]['regularMarketPrice']!== 'undefined'){

          AssetResult['Value'] = Data[I]['regularMarketPrice']

        }

        AssetResult['Earnings'] = 0

        if(Data[I]['dividendsData'] !== 'undefined' && Object.keys(Data[I]['dividendsData']).length > 0 ){

          if(Data[I]['dividendsData']['cashDividends'] !== 'undefined' && Data[I]['dividendsData']['cashDividends'].length > 0){

            AssetResult['Earnings'] = Data[I]['dividendsData']['cashDividends'][0]['rate']

          }

        }
          
        Sheet.getRange(AssetDetails['Line'], AssetDetails['Value']).setValue(AssetResult['Value'])
        Sheet.getRange(AssetDetails['Line'], AssetDetails['Earnings']).setValue(AssetResult['Earnings'])

      }

    }

  }

}

function updateFinancialDevelopment(input){

  let Sheet = SpreadsheetApp.getActive().getSheetByName("Development")
  let NetWorth = SpreadsheetApp.getActive().getRange("Retirement!A11").getValue()

  Sheet.insertRowBefore(1)

  Sheet.getRange(1, 1).setValue(new Date())
  Sheet.getRange(1, 2).setValue(NetWorth)

}


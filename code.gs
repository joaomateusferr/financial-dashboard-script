const IndicatorsPositions = {
  "Dolar-Real" : {"Row" : 3, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4},
  "SLIC" : {"Row" : 4, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4},
  "CDI" :  {"Row" : 5, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4},
  "Inflation" :  {"Row" : 6, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4} 
}

const B3apiBaseUrl = 'https://sistemaswebb3-balcao.b3.com.br'
const BrapiBaseUrl = 'https://brapi.dev/api/v2'

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

  let FisData = SpreadsheetApp.getActive().getRange("Investments!T3:T").getValues()
  let Fis = parseAssets(FisData)
  let StocksData = SpreadsheetApp.getActive().getRange("Investments!N3:N").getValues()
  let Stocks = parseAssets(StocksData)
  let Assets = Fis.concat(Stocks)
  return Assets

}

function getAssetsInfo(input){
  
  let Assets = getAssets()
  
  //https://brapi.dev/api/quote/IVVB11%2CHGLG11?range=1d&interval=1d&fundamental=false&dividends=true

}
const IndicatorsPositions = {
  "Dolar-Real" : {"Row" : 3, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4},
  "SLIC" : {"Row" : 4, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4},
  "CDI" :  {"Row" : 5, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4},
  "Inflation" :  {"Row" : 6, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4} 
};

const B3apiBaseUrl = 'https://sistemaswebb3-balcao.b3.com.br'
const BrapiBaseUrl = 'https://brapi.dev/api/v2'

function getCDI(input) {

  var Url = B3apiBaseUrl + '/featuresDIProxy/DICall/GetRateDI/eyJsYW5ndWFnZSI6InB0LWJyIn0='
  var Response = UrlFetchApp.fetch(Url)
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Investments")

  if(Response.getResponseCode() == 200 && Response.getContentText() != ""){

    var Data = JSON.parse(Response.getContentText())

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
  
  var CurrentDate = new Date().toLocaleDateString('pt-BR');
  var Url = BrapiBaseUrl + '/prime-rate?country=brazil&historical=true&start=' + CurrentDate + '&end=' + CurrentDate + '&sortBy=date&sortOrder=desc'
  var Response = UrlFetchApp.fetch(Url)
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Investments")

  if(Response.getResponseCode() == 200 && Response.getContentText() != ""){

    var Data = JSON.parse(Response.getContentText())

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
  
  var CurrentDate = new Date().toLocaleDateString('pt-BR');
  var Url = BrapiBaseUrl + '/inflation?country=brazil&historical=false&start=' + CurrentDate + '&end=' + CurrentDate + '&sortBy=date&sortOrder=desc'
  var Response = UrlFetchApp.fetch(Url)
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Investments")

  if(Response.getResponseCode() == 200 && Response.getContentText() != ""){

    var Data = JSON.parse(Response.getContentText())

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
  
  var Url = BrapiBaseUrl + '/currency?currency=USD-BRL'
  var Response = UrlFetchApp.fetch(Url)
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Investments")

  if(Response.getResponseCode() == 200 && Response.getContentText() != ""){

    var Data = JSON.parse(Response.getContentText())

    if(typeof Data['currency'][0] !== 'undefined'){

      Data = Data['currency'][0]    
      Data.askPrice = parseFloat(Data.askPrice).toFixed(2)

      Data.askPrice = Data.askPrice.toString().replace('.', ',')
      var UpdateDate = new Date(Number(Data.updatedAtTimestamp)*1000)
      var UpdatedAt = UpdateDate.getDate() + "/" + (UpdateDate.getMonth()+1) + "/" + UpdateDate.getFullYear()

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
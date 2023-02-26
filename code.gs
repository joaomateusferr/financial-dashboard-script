const IndicatorsPositions = {
  "CDI" : {"Row" : 5, "RateCol" : 2, "DataCol" : 3, "UpdateCol" : 4} 
};

function getCDI(input) {

  var Url = "https://sistemaswebb3-balcao.b3.com.br/featuresDIProxy/DICall/GetRateDI/eyJsYW5ndWFnZSI6InB0LWJyIn0="  
  var Response = UrlFetchApp.fetch(Url)

  if(Response.getResponseCode() == 200 && Response.getContentText() != ""){

    var Data = JSON.parse(Response.getContentText())

    if(typeof Data.rate !== 'undefined' && typeof Data.date){
      Sheet = SpreadsheetApp.getActive().getSheetByName("Investments")
      Sheet.getRange(IndicatorsPositions["CDI"]["Row"], IndicatorsPositions["CDI"]["RateCol"]).setValue(Data.rate)
      Sheet.getRange(IndicatorsPositions["CDI"]["Row"], IndicatorsPositions["CDI"]["DataCol"]).setValue(Data.date)
      Sheet.getRange(IndicatorsPositions["CDI"]["Row"], IndicatorsPositions["CDI"]["UpdateCol"]).setValue(new Date())
      return
    }

  }

  Sheet = SpreadsheetApp.getActive().getSheetByName("Investments")
  Sheet.getRange(IndicatorsPositions["CDI"]["Row"], IndicatorsPositions["CDI"]["RateCol"]).setValue(0)
  Sheet.getRange(IndicatorsPositions["CDI"]["Row"], IndicatorsPositions["CDI"]["DataCol"]).setValue(new Date())
  Sheet.getRange(IndicatorsPositions["CDI"]["Row"], IndicatorsPositions["CDI"]["UpdateCol"]).setValue(new Date())

}

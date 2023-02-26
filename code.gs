function getCDI(input) {

  var Url = "https://sistemaswebb3-balcao.b3.com.br/featuresDIProxy/DICall/GetRateDI/eyJsYW5ndWFnZSI6InB0LWJyIn0="
  var Row = 5
  var RateCol = 2
  var DataCol = 3
  var UpdateCol = 4  
  
  var Response = UrlFetchApp.fetch(Url)

  if(Response.getResponseCode() == 200 && Response.getContentText() != ""){

    var Data = JSON.parse(Response.getContentText())

    if(typeof Data.rate !== 'undefined' && typeof Data.date){
      Sheet = SpreadsheetApp.getActive().getSheetByName("Investimentos")
      Sheet.getRange(Row, RateCol).setValue(Data.rate)
      Sheet.getRange(Row, DataCol).setValue(Data.date)
      Sheet.getRange(Row, UpdateCol).setValue(new Date())
      return
    }

  }

  Sheet = SpreadsheetApp.getActive().getSheetByName("Investimentos")
  Sheet.getRange(Row, RateCol).setValue(0)
  Sheet.getRange(Row, DataCol).setValue(new Date())
  Sheet.getRange(Row, UpdateCol).setValue(new Date())

}

# vba-json-api

You need to install vba-json script here:

```
https://github.com/VBA-tools/VBA-JSON
```

Then from your sheet, press `ctrl` + `f11` and insert new module:

```
Public Sub XmlHttpTutorial()
  Dim xmlhttp As Object
  Set xmlhttp = CreateObject("MSXML2.serverXMLHTTP")
  myurl = "https://finance.vietstock.vn/data/getmarketprice?type=2"
  xmlhttp.Open "GET", myurl, False
  xmlhttp.send
  Dim json As Object
  Set json = JsonConverter.ParseJson(xmlhttp.responseText)

  MsgBox (JsonConverter.ConvertToJson(json))
  'Sheets(1).Cells(1, 1).Value = "Movie.Title"
  Dim V As Dictionary
  Dim i As Integer
  i = 1
  
  For Each V In json
    Sheets(1).Cells(i + 1, 1).Value = V("Code")
    Sheets(1).Cells(i + 1, 2).Value = V("Name")
    Sheets(1).Cells(i + 1, 3).Value = V("Price")
    Sheets(1).Cells(i + 1, 4).Value = V("Change")
    Sheets(1).Cells(i + 1, 5).Value = V("TradingDate")
    i = i + 1
  Next V
End Sub
```

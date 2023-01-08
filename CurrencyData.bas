Attribute VB_Name = "CurrencyData"
Function CurrencyRate(Currency1 As String, Currency2 As String) As Double

Dim objRequest As Object
Dim strUrl As String
Dim binAsync As Boolean
Dim strResponse As String
Dim json As Object

Set objRequest = CreateObject("MSXML2.XMLHTTP")
strUrl = "https://query1.finance.yahoo.com/v8/finance/chart/" + Currency1 + Currency2 + "=X"
binAsync = True

With objRequest
    .Open "GET", strUrl, binAsync
    .SetRequestHeader "Content-Type", "application/json"
    .send
    
    While objRequest.readyState <> 4
        DoEvents
    Wend
    
    strResponse = .responseText
    
End With

Set json = JsonConverter.ParseJson(strResponse)

CurrencyRate = CDbl(json("chart")("result")(1)("meta")("regularMarketPrice"))
    
End Function

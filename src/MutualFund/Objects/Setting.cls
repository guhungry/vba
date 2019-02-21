VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Function ValueFromName(field As String)
    ValueFromName = WCRange.ValueFromName(Setting.Names(field))
End Function

Public Function DateFromWeb()
    DateFromWeb = DataFromWeb.Range(Setting.ValueFromName("DATE_WEB")).value
End Function

Public Function DateFromCache()
    DateFromCache = StockPrice.Range(Setting.ValueFromName("DATE_STOCKPRICE")).value
End Function

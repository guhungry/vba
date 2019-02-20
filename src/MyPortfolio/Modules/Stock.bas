Attribute VB_Name = "Stock"
'Global Variable
Public LastDataFromWeb As Long
Public LastStockPrice As Long
Public FundSource As String
Private IsDataFromWebNewer As Boolean

'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Initialize Global Variables - Must be call before doing anything
'---------------------------------------------------------------------------------------
Public Sub Initialize()
    LastDataFromWeb = WCRange.LastRow(DataFromWeb.Cells)
    LastStockPrice = WCRange.LastRow(StockPrice.Cells)
    IsDataFromWebNewer = (DateDiff("d", GetUpdateDate("STOCK"), GetUpdateDate("WEB")) >= 0)
    FundSource = "=" & WCRange.ToAddress(Fund.Range("F10"))
End Sub

'---------------------------------------------------------------------------------------
' Function  : FindStock
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return Range of matched Stock in DataFromWeb, used by GetStockValue and GetStockCell
'---------------------------------------------------------------------------------------
Private Function FindStock(StockName As String) As Range
    Dim Found As Range
    Set Found = DataFromWeb.Range("A27", "A" & LastDataFromWeb)
    Set Found = Found.Find(What:=StockName & " ", After:=Found.Cells(1, 1), LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)

    Set FindStock = Found
End Function

'---------------------------------------------------------------------------------------
' Function  : GetStockValue
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return price of matched Stock in DataFromWeb
' @param StockName  the stock name
' @return           the price of stock
' Example           GetStockValue("AIT") => 29.5
'---------------------------------------------------------------------------------------
Public Function GetStockValue(StockName As String) As Double
    Dim Found As Range
    Set Found = FindStock(StockName)
    
    If Found Is Nothing Then
        GetStockValue = -1
        Exit Function
    End If
    If Not (WCValidate.IsNumber(Found.Cells(1, 6))) Then
        GetStockValue = -1
        Exit Function
    End If
    
    GetStockValue = Found.Cells(1, 6).value
End Function

'---------------------------------------------------------------------------------------
' Function  : GetStockCell
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return address of matched Stock in DataFromWeb
' @param StockName  the stock name
' @return           the address of stock price cell
' Example           GetStockCell("AIT") => 'DataFromWeb!F6'
'---------------------------------------------------------------------------------------
Public Function GetStockCell(StockName As String) As String
    Dim Found As Range
    Set Found = FindStock(StockName)
    
    If Found Is Nothing Then
        GetStockCell = ""
        Exit Function
    End If
    If Not (WCValidate.IsNumber(Found.Cells(1, 6))) Then
        GetStockCell = ""
        Exit Function
    End If
    GetStockCell = "=" & WCRange.ToAddress(Found.Cells(1, 6))
End Function

'---------------------------------------------------------------------------------------
' Function  : GetUpdateDate
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return update date of DataFromWeb or StockPrice
' @param SheetName  the Sheet Name ('WEB', 'STOCKPRICE')
' @return           the address of stock price cell
' Example           GetUpdateDate("WEB") => '30 June 2010'
'---------------------------------------------------------------------------------------
Public Function GetUpdateDate(Optional SheetName As String = "WEB") As Date
    If SheetName = "WEB" Then
        GetUpdateDate = WCDate.ExtractDate(Setting.DateFromWeb())
    Else
        GetUpdateDate = WCDate.ExtractDate(Setting.DateFromCache())
    End If
End Function

'---------------------------------------------------------------------------------------
' Function  : GetLastUpdate
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return latest update date from DataFromWeb or StockPrice
' @return           the address of stock price cell
' Example           GetLastUpdate() => 'Last Update 30 Jun 2010 16:59:45'
'---------------------------------------------------------------------------------------
Public Function GetLastUpdate() As String
    If IsDataFromWebNewer Then
        GetLastUpdate = Setting.DateFromWeb()
    Else
        GetLastUpdate = Setting.DateFromCache()
    End If
End Function

'---------------------------------------------------------------------------------------
' Function  : IsStockPriceExist
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Check weather stock exist in StockPrice
' @param StockName  the stock name
' @return           Is stock exists in StockPrice
' Example           IsStockPriceExist("AIT") => True
'---------------------------------------------------------------------------------------
Public Function IsStockPriceExist(StockName As String) As Boolean
    Dim Found As Range
    Set Found = FindStockPriceCell(StockName)
    
    IsStockPriceExist = Not (Found Is Nothing)
End Function

'---------------------------------------------------------------------------------------
' Function  : GetStockPriceCell
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return address of matched Stock in StockPrice
' @param StockName  the stock name
' @return           the address of stock price cell
' Example           GetStockCell("AIT") => 'StockPrice!C35'
'---------------------------------------------------------------------------------------
Public Function GetStockPriceCell(StockName As String) As String
    Dim Found As Range
    Set Found = FindStockPriceCell(StockName)
    
    If Found Is Nothing Then
        GetStockPriceCell = ""
        Exit Function
    End If
    If Not WCValidate.IsNumber(Found.Cells(1, 3)) Then
        GetStockPriceCell = ""
        Exit Function
    End If
    
    GetStockPriceCell = "=" & WCRange.ToAddress(Found.Cells(1, 3))
End Function

Public Function FindStockPriceCell(StockName As String) As Range
    Dim Found As Range
    Set Found = StockPrice.Range("A5:A" & LastStockPrice)
    Set Found = Found.Find(What:=Trim(StockName), After:=Found.Cells(1, 1), LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
    FindStockPriceCell = Found
End Function

'---------------------------------------------------------------------------------------
' Procedure : InsertStockPrice
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Insert Stock into StockPrice
' @param StockName  the stock name
' @param StockDate  the update date
' @param price      the stock price
'---------------------------------------------------------------------------------------
Public Sub InsertStockPrice(StockName As String, StockDate As Date, price As Double)
    Dim BlankRow, source As Range
    Set source = StockPrice.Range("A3:C3")
    Set BlankRow = StockPrice.Range("A" & (LastStockPrice + 1) & ":C" & (LastStockPrice + 1))
    source.Copy
    BlankRow.EntireRow.Insert xlDown, xlFormatFromRightOrBelow
    Application.CutCopyMode = False
    BlankRow.Cells(0, 1).value = Trim(StockName)
    BlankRow.Cells(0, 2).value = StockDate
    BlankRow.Cells(0, 3).value = price
End Sub

'---------------------------------------------------------------------------------------
' Function  : GetUpdateDate
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return reference to the latest stock price in DataFromWeb or StockPrice
' @param StockName  the stock name
' @param Comment    the output comment (Source : Last update date)
' @return           the reference to latest stock price
' Example           GetLatestStockPrice("AIT") => "=DataFromWeb!F2525"
'---------------------------------------------------------------------------------------
'Return Reference to StockPrice or DataFromWeb
Public Function GetLatestStockPrice(StockName As String, ByRef comment As String) As String
    'Compare StockPrice & DataFromWeb Date
    If IsDataFromWebNewer Then
        comment = "Web : " & Setting.DateFromWeb()
        GetLatestStockPrice = GetStockCell(StockName)
    End If
    
    'If price is null get from stockprice
    If GetLatestStockPrice = "" Then
        comment = "StockPrice : " & Setting.DateFromCache()
        GetLatestStockPrice = GetStockPriceCell(StockName)
    End If
End Function

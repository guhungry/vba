VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataFromWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Procedure : btnResetWebData
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Clear DataFromWeb and regenerate DataFromWeb
' Change Log
'  yyyy-mm-dd  - Fix Fucked Up Web Data - Cause of Listed or Delisted stock(s).
' 2013-05-25 - Add Warrant Table
'---------------------------------------------------------------------------------------
'
Sub btnResetWebData()
    'Define Variables
    Dim LastRow As Integer

    LastRow = WCRange.LastRow(DataFromWeb.Cells) + 2

    'Clear Data
    Application.ScreenUpdating = False
    Range("A1:L" & LastRow).Delete Shift:=xlUp
    LastRow = 1
    'Delete all connections
    For Each cn In ThisWorkbook.Connections
        cn.Delete
    Next cn
    'Delete data from web query
    While DataFromWeb.QueryTables.Count
        DataFromWeb.QueryTables(1).Delete
    Wend
    'SET Main
    SetDataFromWeb "SetMain", Setting.ValueFromName("URL_SETMAIN"), Setting.ValueFromName("TAB_SETMAIN"), "A" & LastRow
    LastRow = WCRange.LastRow(DataFromWeb.Cells) + 2
    'SET Foreign
    SetDataFromWeb "SetForeign", Setting.ValueFromName("URL_SETFOREIGN"), Setting.ValueFromName("TAB_SETFOREIGN"), "A" & LastRow
    LastRow = WCRange.LastRow(DataFromWeb.Cells) + 2
    'SET Warrant
    SetDataFromWeb "SetWarrant", Setting.ValueFromName("URL_SETWARRANT"), Setting.ValueFromName("TAB_SETWARRANT"), "A" & LastRow
    LastRow = WCRange.LastRow(DataFromWeb.Cells) + 2
    'MAI Main
    SetDataFromWeb "MaiMain", Setting.ValueFromName("URL_MAIMAIN"), Setting.ValueFromName("TAB_MAIMAIN"), "A" & LastRow
    LastRow = WCRange.LastRow(DataFromWeb.Cells) + 2
    'MAI Foreign
    SetDataFromWeb "MaiForeign", Setting.ValueFromName("URL_MAIFOREIGN"), Setting.ValueFromName("TAB_MAIFOREIGN"), "A" & LastRow
    LastRow = WCRange.LastRow(DataFromWeb.Cells) + 2
    'MAI Warrant
    SetDataFromWeb "MaiWarrant", Setting.ValueFromName("URL_MAIWARRANT"), Setting.ValueFromName("TAB_MAIWARRANT"), "A" & LastRow
    LastRow = WCRange.LastRow(DataFromWeb.Cells) + 2

    Transaction.btnTransactionSortByTicker
    DataFromWeb.Activate
    Application.ScreenUpdating = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnRefreshFromWeb
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Update Data From Web Now!!!!
'---------------------------------------------------------------------------------------
'
Sub btnRefreshFromWeb()
    ActiveWorkbook.RefreshAll
End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnUpdateStockPrice
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Update Web Data to StockPrice
'---------------------------------------------------------------------------------------
'
Sub btnUpdateStockPrice()
    'Define Variables
    Dim NumStockPrice As Long
    Dim list As Range
    Dim price As Double
    Dim today As Date

    'Default Value
    Stock.Initialize
    StockPrice.Range(Setting.ValueFromName("DATE_STOCKPRICE")).value = Setting.DateFromWeb()
    Set list = StockPrice.Range("A6:A" & Application.WorksheetFunction.Max(WCRange.LastRow(StockPrice.Cells), 6))
    today = GetUpdateDate()

    ' Update Stock Price
    For Each S In list
        If DateDiff("d", S.Cells(1, 2).value, today) > 0 Then
            price = Stock.GetStockValue(S.value)
            If price > 0 Then
                S.Cells(1, 2).value = today
                S.Cells(1, 3).value = price
            End If
        End If
    Next
    
    'Insert Unknown StockPrice
    Set list = DataFromWeb.Range("A28:A" & Stock.LastDataFromWeb)
    If (Stock.LastStockPrice > 5) Then
        For Each S In list
            If WCValidate.IsNumber(S.Cells(1, 6).value) And Not (Stock.IsStockPriceExist(S.value)) Then
                price = S.Cells(1, 6).value
                Stock.InsertStockPrice S.value, today, price
            End If
        Next
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetDataFromWeb
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Set up DataFromWeb connection to given target cell
' @param Name   the connection name
' @param Table  the data table
' @param target the target cell
'---------------------------------------------------------------------------------------
'
Private Sub SetDataFromWeb(name As String, URL As String, Table As String, target As String)
    With DataFromWeb.QueryTables.Add(Connection:="URL;" & URL, Destination:=DataFromWeb.Range(target))
        .name = name
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingAll
        .WebTables = Table
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
End Sub


Attribute VB_Name = "Util"
Option Explicit


Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TimeZoneInfo) As Long


Private Type SYSTEMTIME
        intYear As Integer
        intMonth As Integer
        intwDayOfWeek As Integer
        intDay As Integer
        intHour As Integer
        intMinute As Integer
        intSecond As Integer
        intMilliseconds As Integer
End Type


Private Type TimeZoneInfo
        lngBias As Long
        intStandardName(0 To 31) As Integer
        intStandardDate As SYSTEMTIME
        intStandardBias As Long
        intDaylightName(0 To 31) As Integer
        intDaylightDate As SYSTEMTIME
        intDaylightBias As Long
End Type
    
    
  Private Enum TIME_ZONE
        TIME_ZONE_ID_INVALID = 0        ' Cannot determine DST
        TIME_ZONE_STANDARD = 1          ' Standard Time, not Daylight
        TIME_ZONE_DAYLIGHT = 2          ' Daylight Time, not Standard
    End Enum
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' NOTE: If you are using the Windows WinAPI Viewer Add-In to get
    ' function declarations, not that there is an error in the
    ' TIME_ZONE_INFORMATION structure. It defines StandardName and
    ' DaylightName As 32. This is fine if you have an Option Base
    ' directive to set the lower bound of arrays to 1. However, if
    ' your Option Base directive is set to 0 or you have no
    ' Option Base diretive, the code won't work. Instead,
    ' change the (32) to (0 To 31).
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(0 To 31) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(0 To 31) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
    End Type


Public RunWhen As Double
Public Const cRunIntervalMinutes = 2 ' in minutes
Public Const cRunWhat = "UpdateAllGamesUsingTimer"  ' the name of the procedure to run

Public Const SoccerLength As Integer = 64
Public Const AFLLength As Integer = 50
Public Const NFLLength As Integer = 48
Public Const MaxSheetlength As Integer = 72

Public Const LogFile As String = "log.txt"
Public Const OutputColumn As Integer = 2
Public Const OutputRow As Integer = 5
Public Const ListEventsMethod As String = "listEvents"
Public Const ListCompetitionsMethod As String = "listCompetitions"
Public Const ListEventTypesMethod As String = "listEventTypes"
Public Const ListMarketCatalogueMethod As String = "listMarketCatalogue"
Public Const ListMatchCatalogueMethod As String = "listMarketCatalogue"
Public Const ListSoccerMatchCatalogueMethod As String = "listMarketCatalogue"
Public Const ListMarketTypesMethod As String = "listMarketTypes"
Public Const ListCurrentOrdersMethod As String = "listCurrentOrders"
Public Const GetAccountFundsMethod As String = "getAccountFunds"

Public Const ListMarketBookMethod As String = "listMarketBook"
Public Const PlaceOrdersMethod As String = "placeOrders"
Public Const ReplaceOrdersMethod As String = "replaceOrders"
Public Const CancelOrdersMethod As String = "cancelOrders"




'
' Make a request to the Betfair API.
'
Function SendRequest(url, appkey, Session, Optional Data) As String
    On Error GoTo ErrorHandler:
    Dim xhr: Set xhr = CreateObject("MSXML2.XMLHTTP")
  
    With xhr
        .Open "POST", url & "/", False
        .setRequestHeader "X-Application", appkey
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
    End With
    
    If Session <> "" Then
        xhr.setRequestHeader "X-Authentication", Session
    End If
    
    xhr.Send Data
    SendRequest = xhr.responseText
    
    If xhr.Status <> 200 Then
        Err.Raise vbObjectError + 1000, "Util.SendRequest", "The call to API-NG was unsuccessful. Status code: " & xhr.Status & " " & xhr.statusText & ". Response was: " & xhr.responseText
    End If
    
    
    'xhr.Status = 12007 = no internet connection
    
    Set xhr = Nothing
    
    On Error GoTo 0
    Exit Function
        
ErrorHandler:
    HandleError "SendRequest"
    'Resume Next
    
    
End Function

Function SendLoginRequest(url, appkey, username, password) As String
    On Error GoTo ErrorHandler:
    
    Dim xhr: Set xhr = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim myToken As String
  
    With xhr
        .Open "POST", url & "/", False
        .setRequestHeader "X-Application", appkey
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .setRequestHeader "Accept", "application/json"
        '.setClientCertificate "Common Name"
        '.setClientCertificate "LOCAL_MACHINE\My\Michael"
         
    End With
    
    
    xhr.Send "username=" & username & "&password=" & password
    SendLoginRequest = xhr.responseText
    
    If xhr.Status <> 200 Then
        Err.Raise vbObjectError + 1000, "Util.SendRequest", "The call to API-NG was unsuccessful. Status code: " & xhr.Status & " " & xhr.statusText & ". Response was: " & xhr.responseText
    End If
    
     
    Set xhr = Nothing
    
    On Error GoTo 0
    Exit Function
        
ErrorHandler:
    HandleError "SendRequest"
    'Resume Next
    
    
End Function


Function ParseJsonRpcResponseToCollection(ByVal Response As String, Optional CallingProc As String) As Object
    On Error GoTo ErrorHandler:
    
    Dim Lib As New jsonlib
    Set ParseJsonRpcResponseToCollection = Lib.parse(Response).Item("result")
    Exit Function
    
    
ErrorHandler:
    HandleError "ParseJsonRpcResponseToCollection" & CallingProc
End Function

Function ParseRestResponseToCollection(ByVal Response As String) As Object
    On Error GoTo ErrorHandler:
    Dim Lib As New jsonlib
    Set ParseRestResponseToCollection = Lib.parse(Response)
    Exit Function
    
ErrorHandler:
    HandleError "ParseRestResponseToCollection"
End Function

Sub HandleError(CallingProc As String)

    If Err.Number <> 0 Then
        Dim Msg As String: Msg = "Error occurred: " & Err.Number & " - " & Err.Description & " "
        Call AppendToLogFile(Msg, CallingProc)
        Sheets("Example").Cells(GetNamedRngRow("ErrorLocation", "Example"), GetNamedRngColumn("ErrorLocation", "Example")).Value = Msg
    End If
    
    'End ' Exit the macro entirely

End Sub
Function GetKeepAliveUrl(Optional Country As String) As String

On Error GoTo ErrorHandler:

If Country = "" Then Country = "UK"

    If Country = "AUS" Then
        GetKeepAliveUrl = Sheets("Example").Cells(6, 2).Value 'use the AUS endpoint
    Else
        GetKeepAliveUrl = Sheets("Example").Cells(6, 2).Value 'else use the UK endpoint by default
    End If
    
Exit Function

ErrorHandler:
    HandleError "ParseRestResponseToCollection"
    Resume Next
    
End Function
Function GetLoginUrl(Optional Country As String) As String

On Error GoTo ErrorHandler:

If Country = "" Then Country = "UK"

    If Country = "AUS" Then
        GetLoginUrl = Sheets("Example").Cells(6, 14).Value 'use the AUS endpoint
    Else
        GetLoginUrl = Sheets("Example").Cells(6, 14).Value 'else use the UK endpoint by default
    End If
    
Exit Function

ErrorHandler:
    HandleError "GetLoginUrl"
    Resume Next
    
End Function
Function GetJsonRpcUrl(Optional Country As String) As String

On Error GoTo ErrorHandler:

If Country = "" Then Country = "UK"

    If Country = "AUS" Then
        GetJsonRpcUrl = Sheet4.Cells(1, 3).Value 'use the AUS endpoint
    Else
        GetJsonRpcUrl = Sheet4.Cells(1, 4).Value 'else use the UK endpoint by default
    End If
    
Exit Function

ErrorHandler:
    HandleError "GetJsonRpcUrl"
    Resume Next
    
End Function
Function GetJsonRpcAccountUrl(Optional Country As String) As String

On Error GoTo ErrorHandler:

If Country = "" Then Country = "UK"

    If Country = "AUS" Then
        GetJsonRpcAccountUrl = Sheet4.Cells(1, 5).Value 'use the AUS endpoint
    Else
        GetJsonRpcAccountUrl = Sheet4.Cells(1, 6).Value 'else use the UK endpoint by default
    End If
    
Exit Function

ErrorHandler:
    HandleError "GetJsonRpcAccountUrl"
    Resume Next
    
End Function
Function GetRestUrl() As String
    
On Error GoTo ErrorHandler:
' read from sheet
    GetRestUrl = Sheet4.Cells(2, 2).Value
    
Exit Function

ErrorHandler:
    HandleError "GetRestUrl"
    Resume Next
    
End Function

Function GetAppKey() As String
    
On Error GoTo ErrorHandler:
' read from sheet
    GetAppKey = Sheet4.Cells(3, 2).Value

Exit Function

ErrorHandler:
    HandleError "GetAppKey"
    Resume Next

End Function

Function GetSession() As String
    
On Error GoTo ErrorHandler:
' read from sheet
    GetSession = Sheet4.Cells(4, 2).Value

Exit Function

ErrorHandler:
    HandleError "GetSession"
    Resume Next
    
End Function

Function MakeJsonRpcRequestString(ByVal Method As String, ByVal RequestString As String) As String
    MakeJsonRpcRequestString = "{""jsonrpc"": ""2.0"", ""method"": ""SportsAPING/v1.0/" & Method & """, ""params"": " & RequestString & ", ""id"": 1}"
End Function
Function MakeJsonRpcAccountsRequestString(ByVal Method As String, ByVal RequestString As String) As String
    MakeJsonRpcAccountsRequestString = "{""jsonrpc"": ""2.0"", ""method"": ""AccountAPING/v1.0/" & Method & """, ""params"": " & RequestString & ", ""id"": 1}"
End Function
Function GetAccountFundsRequestString(wallet As String) As String
    'GetAccountFundsRequestString = "{""wallet"":[""" & wallet & """]}"
    GetAccountFundsRequestString = "{""wallet"":""" & wallet & """}"
End Function
Function GetListCompetitionsRequestString() As String
    GetListCompetitionsRequestString = "{""filter"":{}}"
End Function
Function GetListEventTypesRequestString() As String
    GetListEventTypesRequestString = "{""filter"":{}}"
End Function
Function GetListSingleEventRequestString(eventid As String) As String
    'GetListCurrentOrdersRequestString = "{""filter"":{""marketIds"":[""" & marketid & """]}}"
    GetListSingleEventRequestString = "{""filter"":{""eventIds"":[""" & eventid & """]}}"
End Function
Function GetListCurrentOrdersRequestString(marketid As String) As String
    'GetListCurrentOrdersRequestString = "{""filter"":{""marketIds"":[""" & marketid & """]}}"
    GetListCurrentOrdersRequestString = "{""marketIds"":[""" & marketid & """]}"
End Function


Function GetListMarketTypesRequestString(marketid As String, competitionID As String) As String
    GetListMarketTypesRequestString = "{""filter"":{""marketIds"":[""" & marketid & """],""competitionIds"":[""" & competitionID & """]}}"
End Function
'Function GetListEventsRequestString() As String
'    GetListEventsRequestString = "{""filter"":{}}"
'End Function
Function GetListEventsRequestString(ByVal EventTypeId As String, ByVal competitionID As String) As String
    GetListEventsRequestString = "{""filter"":{""competitionIds"":[""" & competitionID & """]}}"
End Function

Function GetListMarketCatalogueRequestStringNF(ByVal EventTypeId As String) As String
'Get with No Filter set = NF
    Dim dateNow As Date: dateNow = Format(Now, "yyyy-mm-dd hh:mm:ss")
    GetListMarketCatalogueRequestStringNF = "{""filter"":{""eventTypeIds"":[""" & EventTypeId & """]}}"
End Function
Function GetListMarketCatalogueRequestString(ByVal EventTypeId As String) As String
    Dim dateNow As Date: dateNow = Format(Now, "yyyy-mm-dd hh:mm:ss")
    GetListMarketCatalogueRequestString = "{""filter"":{""eventType"":[""" & EventTypeId & """],""marketCountries"":[""AUS""],""marketTypeCodes"":[""WIN""]},""marketStartTime"":{""from"":""" & dateNow & """},""sort"":""FIRST_TO_START"",""maxResults"":""1"",""marketProjection"":[""RUNNER_DESCRIPTION""]}"
End Function
Function GetListMarketCatalogueRequestStringUsingMarketID(ByVal marketid As String) As String
    Dim dateNow As Date: dateNow = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    GetListMarketCatalogueRequestStringUsingMarketID = "{""filter"":{""marketIds"":[""" & marketid & """]},""maxResults"":""100"",""marketProjection"":[""EVENT"",""EVENT_TYPE"",""RUNNER_METADATA""]}"
    'GetListMarketCatalogueRequestStringUsingMarketID = "{""filter"":{""marketTypeCodes"":[""MATCH_ODDS"",""WINNING_MARGIN""]},""maxResults"":""100"",""marketProjection"":[""EVENT"",""EVENT_TYPE"",""RUNNER_METADATA""]}}"
End Function
Function GetListMarketCatalogueRequestStringMATCH_ODDS(ByVal eventTypeIds As String) As String
    GetListMarketCatalogueRequestStringMATCH_ODDS = "{""filter"":{""eventTypeIds"":[""" & eventTypeIds & """],""marketTypeCodes"":[""MATCH_ODDS""]},""maxResults"":""450"",""sort"":""MAXIMUM_TRADED"",""marketProjection"":[""COMPETITION"",""EVENT"",""RUNNER_DESCRIPTION""]}"
    'GetListMarketCatalogueRequestStringMATCH_ODDS = "{""filter"":{""eventTypeIds"":[""1""],""marketTypeCodes"":[""MATCH_ODDS""]},""maxResults"":""350"",""sort"":""MAXIMUM_TRADED"",""marketProjection"":[""COMPETITION"",""EVENT"",""RUNNER_DESCRIPTION""]}"
    'Working
    'GetListMarketCatalogueRequestStringMATCH_ODDS = "{""filter"":{""eventTypeIds"":[""1""],""marketTypeCodes"":[""MATCH_ODDS""]},""maxResults"":""999"",""marketProjection"":[""COMPETITION"",""EVENT"",""RUNNER_DESCRIPTION""]}"
    'GetListMarketCatalogueRequestStringUsingMarketID = "{""filter"":{""marketTypeCodes"":[""MATCH_ODDS"",""WINNING_MARGIN""]},""maxResults"":""100"",""marketProjection"":[""EVENT"",""EVENT_TYPE"",""RUNNER_METADATA""]}}"
End Function
Function GetListMarketCatalogueRequestStringUsingEventID4MATCH_ODDS(ByVal eventid As String) As String
    GetListMarketCatalogueRequestStringUsingEventID4MATCH_ODDS = "{""filter"":{""eventIds"":[""" & eventid & """],""marketTypeCodes"":[""MATCH_ODDS""]},""maxResults"":""3"",""marketProjection"":[""COMPETITION"",""EVENT"",""RUNNER_DESCRIPTION""]}"
    'GetListMarketCatalogueRequestStringUsingMarketID = "{""filter"":{""marketTypeCodes"":[""MATCH_ODDS"",""WINNING_MARGIN""]},""maxResults"":""100"",""marketProjection"":[""EVENT"",""EVENT_TYPE"",""RUNNER_METADATA""]}}"
End Function

Function GetListSoccerMatchCatalogueRequestString(ByVal eventid As String) As String
    GetListSoccerMatchCatalogueRequestString = "{""filter"":{""eventType"":[""" & eventid & """],""marketName"":""Match Odds""},""maxResults"":""100"",""marketProjection"":[""RUNNER_DESCRIPTION""]}"
End Function
Function GetListMatchCatalogueRequestString(ByVal eventid As String) As String
    GetListMatchCatalogueRequestString = "{""filter"":{""eventIds"":[""" & eventid & """],""marketName"":""Match Odds""},""maxResults"":""100"",""marketProjection"":[""RUNNER_DESCRIPTION""]}"
End Function
Function GetListMatchCatalogueRequestString4MyEventTypes(ByVal eventid As String) As String
    GetListMatchCatalogueRequestString4MyEventTypes = "{""filter"":{""eventIds"":[""" & eventid & """],""marketTypeCodes"":[""MATCH_ODDS"",""OVER_UNDER_05""]},""maxResults"":""100"",""marketProjection"":[""RUNNER_DESCRIPTION""]}"
End Function
Function GetListMarketBookRequestString(ByVal marketid As String) As String
    GetListMarketBookRequestString = "{""marketIds"":[""" & marketid & """],""priceProjection"":{""priceData"":[""EX_BEST_OFFERS""]}}"
End Function
'// SP addition 20 Feb - following support email to Michael Lean re the delay of prices or prices missing...
Function GetListMarketBookRequestStringV(ByVal marketid As String) As String
    GetListMarketBookRequestStringV = "{""marketIds"":[""" & marketid & """],""priceProjection"":{""priceData"":[""EX_BEST_OFFERS""],""virtualise"":""true""}}"
End Function
Function GetPlaceOrdersRequestString(ByVal marketid As String, ByVal selectionID As String, ByVal Side As String, ByVal Price As Double, ByVal Size As Double) As String
    GetPlaceOrdersRequestString = "{""marketId"":""" & marketid & """,""instructions"":[{""selectionId"":""" & selectionID & """,""handicap"":""0"",""side"":""" & Side & """,""orderType"":""LIMIT"",""limitOrder"":{""size"":""" & Size & """,""price"":""" & Price & """,""persistenceType"":""LAPSE""}}]}"
End Function
'// ML added Handicap function due to new NFL line options
Function GetPlaceOrdersRequestStringHandicap(ByVal marketid As String, ByVal selectionID As String, ByVal Side As String, ByVal Price As Double, ByVal Size As Double, ByVal Handicap As Double) As String
    GetPlaceOrdersRequestStringHandicap = "{""marketId"":""" & marketid & """,""instructions"":[{""selectionId"":""" & selectionID & """,""handicap"":""" & Handicap & """,""side"":""" & Side & """,""orderType"":""LIMIT"",""limitOrder"":{""size"":""" & Size & """,""price"":""" & Price & """,""persistenceType"":""LAPSE""}}]}"
End Function
Function GetReplaceOrdersRequestString(ByVal marketid As String, ByVal BetID As String, ByVal newPrice As Double) As String
    GetReplaceOrdersRequestString = "{""marketId"":""" & marketid & """,""instructions"":[{""betId"":""" & BetID & """,""newPrice"":""" & newPrice & """}]}"
End Function

Function GetCancelOrdersRequestString(ByVal marketid As String, ByVal BetID As String, ByVal newPrice As Double) As String
    'newPrice may be used in the future for a price reduction
        GetCancelOrdersRequestString = "{""marketId"":""" & marketid & """,""instructions"":[{""betId"":""" & BetID & """,""sizeReduction"":""" & newPrice & """}]}"
End Function
Function GetCancelOrdersMarketIDBetIDRequestString(ByVal marketid As String, ByVal BetID As String) As String
    'newPrice may be used in the future for a price reduction
        GetCancelOrdersMarketIDBetIDRequestString = "{""marketId"":""" & marketid & """,""instructions"":[{""betId"":""" & BetID & """}]}"
End Function

Function GetEventTypeIdFromEventTypes(ByVal EventTypes As Object) As String
    
On Error GoTo ErrorHandler:

    GetEventTypeIdFromEventTypes = "0"

    Dim Index As Integer
    For Index = 1 To EventTypes.Count Step 1
        Dim EventType: Set EventType = EventTypes.Item(Index).Item("eventType")
        If EventType.Item("name") = "Australian Rules" Then
            GetEventTypeIdFromEventTypes = EventType.Item("id")
            Exit For
        End If
    Next

Exit Function

ErrorHandler:
    HandleError "GetEventTypeIdFromEventTypes"
    Resume Next
    
End Function

Function GetMarketIdFromMarketCatalogue(ByVal Response As Object) As String
    GetMarketIdFromMarketCatalogue = Response.Item(1).Item("marketId")
End Function

Function GetSelectionIdFromMarketBook(ByVal Response As Object) As String
    Dim Runners As Object: Set Runners = Response.Item(1).Item("runners")
        GetSelectionIdFromMarketBook = Runners.Item(1).Item("selectionId")
    Set Runners = Nothing
End Function

Function GetAvailableToBackForSelection(ByVal selectionID As String, ByVal Response As Object) As Collection
    Dim Runners As Object: Set Runners = Response.Item(1).Item("runners")
    
    Dim Index As Integer
    For Index = 1 To Runners.Count Step 1
        Dim ID: ID = Runners.Item(Index).Item("selectionId")
        If ID = selectionID Then
            Set GetAvailableToBackForSelection = Runners.Item(Index).Item("ex").Item("availableToBack")
            Exit For
        End If
    Next
    
    Set Runners = Nothing
End Function

Function Get3PricesAvailableToBackForMarket(ByVal Response As Object) As Collection
    Dim Runners As Object: Set Runners = Response.Item(1).Item("runners")
    'Dim Prices As Object: Set Prices = Response.Item(1).Item("runners").Item("ex")
 
 Dim Index, myPrice As Integer
 For Index = 1 To Runners.Count Step 1
 For myPrice = 1 To 3
        Set Get3PricesAvailableToBackForMarket = Runners.Item(Index).Item("ex").Item("availableToBack")
 Next
 Next
    
    Set Runners = Nothing
End Function
Function GetMarketIdFromMatchCatalogue(ByVal MarketType As String, ByVal Response As Object) As String
    
On Error GoTo ErrorHandler:
'Dim Runners As Object: Set Runners = Response.Item(1).Item("runners")
    Dim marketindex As Integer
    Dim Index As Integer
    For marketindex = 1 To Response.Count Step 1
        
        Dim marketid: marketid = Response.Item(marketindex).Item("marketId")
        Dim marketName: marketName = Response.Item(marketindex).Item("marketName")
    
        If marketName = MarketType Then
            GetMarketIdFromMatchCatalogue = marketid
            Exit For
        End If
    Next

On Error GoTo 0
Exit Function

ErrorHandler:
    HandleError "GetMarketIdFromMatchCatalogue"
    Resume Next
    'Set Runners = Nothing
End Function
Function GetSelectionIdFromMatchCatalogue(ByVal MarketType As String, ByVal Response As Object) As Collection
    
On Error GoTo ErrorHandler:

Dim Runners As Object: Set Runners = Response.Item().Item().Item("runners")
    Dim marketindex As Integer
    Dim Index As Integer
    For marketindex = 1 To Response.Count Step 1
        
        Dim marketid: marketid = Response.Item(marketindex).Item("marketId")
        Dim marketName: marketName = Response.Item(marketindex).Item("marketName")
    
        If marketName = MarketType Then 'then
            Set GetSelectionIdFromMatchCatalogue = Response.Item(marketindex).Item(3)
            Exit For
        End If
    Next

Exit Function

ErrorHandler:
    HandleError "GetMarketIdFromMatchCatalogue"
    Resume Next
    
    'Set Runners = Nothing
End Function
Sub OutputAvailableToBack(ByRef AvailableToBack As Object, ByRef OutputRow As Integer, ByRef OutputColumn As Integer)
    Dim Index As Integer
    For Index = 1 To AvailableToBack.Count Step 1
        Sheet4.Cells(OutputRow, OutputColumn + Index).Value = AvailableToBack.Item(Index).Item("price")
    Next
End Sub
Sub OutputlistCompetitions(ByRef ListCompetitionsResult As Object, ByRef OutputRow As Integer, ByRef OutputColumn As Integer, ByVal VertHoriz As String)
    Dim Index As Integer
    
    If VertHoriz = "Vert" Then
    
    For Index = 1 To ListCompetitionsResult.Count Step 1
        Sheet4.Cells(OutputRow + Index, OutputColumn).Value = ListCompetitionsResult.Item(Index).Item("competition").Item("id")
        Sheet4.Cells(OutputRow + Index, OutputColumn + 1).Value = ListCompetitionsResult.Item(Index).Item("competition").Item("name")
        Sheet4.Cells(OutputRow + Index, OutputColumn + 2).Value = ListCompetitionsResult.Item(Index).Item("marketCount")
        Sheet4.Cells(OutputRow + Index, OutputColumn + 3).Value = ListCompetitionsResult.Item(Index).Item("competitionRegion")
        
    Next
    
    ElseIf VertHoriz = "Horiz" Then
    
    For Index = 1 To ListCompetitionsResult.Count Step 1
        Sheet4.Cells(OutputRow, OutputColumn + Index).Value = ListCompetitionsResult.Item(Index).Item("competition").Item("id")
        Sheet4.Cells(OutputRow + 1, OutputColumn + Index).Value = ListCompetitionsResult.Item(Index).Item("competition").Item("name")
        Sheet4.Cells(OutputRow + 2, OutputColumn + Index).Value = ListCompetitionsResult.Item(Index).Item("marketCount")
        Sheet4.Cells(OutputRow + 3, OutputColumn + Index).Value = ListCompetitionsResult.Item(Index).Item("competitionRegion")
        
    Next
    
    End If
    
End Sub
Public Function OutputGetAccountFunds(ByVal GetAccountFundsCollecUK As Object, ByVal GetAccountFundsCollecAUS As Object, ByVal MailOrScreen As String)

On Error GoTo ErrorHandler:

Dim MyAccountFunds As Double
Dim MyExposure As Double
Dim BoxReply As Boolean
Dim SendMailBoolean As Boolean

If (GetAccountFundsCollecUK Is Nothing Or GetAccountFundsCollecUK Is Nothing) Then

MyAccountFunds = 99
MyExposure = 99

Else

MyAccountFunds = GetAccountFundsCollecUK.Item("availableToBetBalance")
    '+ GetAccountFundsCollecAUS.Item("availableToBetBalance") - this had to be removed since AUS wallet no longer exists
MyExposure = GetAccountFundsCollecUK.Item("exposure")
    '+ GetAccountFundsCollecAUS.Item("exposure")

End If


Exit_Err_Handler:

If MailOrScreen = "Screen" Then BoxReply = MsgBox("Your Funds Are " & MyAccountFunds - MyExposure & vbCr & MyAccountFunds & " " & MyExposure, vbOKOnly, "Funds")

If MailOrScreen = "Mail" Then
    SendMailBoolean = Sheets("Example").Cells(GetNamedRngRow("SendEmail", "Example"), GetNamedRngColumn("SendEmail", "Example")).Value
    If SendMailBoolean Then Call SendEmail(MyAccountFunds, MyExposure)
End If

Call AddAccountFundsToTrackingSpreadsheet(MyAccountFunds, MyExposure)

Exit Function
 
ErrorHandler:
    
    MyAccountFunds = 9090
    MyExposure = 9090
    HandleError OutputGetAccountFunds & " " & MyAccountFunds
    'Call AppendToLogFile("&MyAccountFunds&", "OutputGetAccountFunds")
    GoTo Exit_Err_Handler


End Function
Public Function OutputListMarketBook(ByVal MarketsCollec As Object, OPRow As Long, OPColumn As Long, Optional ThisMarketID As String) As Integer
'// Simply takes in a collection object of the H-A-D soccer market types and displays markets on activesheet
'// on the passed in output row (OPRow variable)
Dim TotalRunnersInThisMarket As Integer
Dim Backprice As Double
Dim LayPrice As Double
Dim test As Double
Dim Handicap As Integer
Dim TotalMatchedThisHandicap As Double: TotalMatchedThisHandicap = 0
Dim MaxMatched As Double
Dim MaxSelectionID As String
Dim MaxHandicap As Double
Dim MaxIndex, StartLoop, EndLoop As Integer

On Error Resume Next
Dim columncounter, Index As Integer

TotalRunnersInThisMarket = MarketsCollec.Item(1).Item("runners").Count
OutputListMarketBook = TotalRunnersInThisMarket

If TotalRunnersInThisMarket = 0 Then
    Exit Function
End If

StartLoop = 1
EndLoop = TotalRunnersInThisMarket
MaxIndex = 1

        If TotalRunnersInThisMarket > 20 Then 'then its a scrolling market

                For Index = 1 To TotalRunnersInThisMarket
                
                    If MarketsCollec.Item(1).Item("runners").Item(Index).Item("totalMatched") > MaxMatched Then
                        MaxMatched = MarketsCollec.Item(1).Item("runners").Item(Index).Item("totalMatched")
                        MaxSelectionID = MarketsCollec.Item(1).Item("runners").Item(Index).Item("selectionId")
                        MaxHandicap = MarketsCollec.Item(1).Item("runners").Item(Index).Item("handicap")
                        MaxIndex = Index
                    End If
                    'TotalMatchedThisHandicap = TotalMatchedThisHandicap + MarketsCollec.Item(1).Item("runners").Item(Index).Item("totalMatched")
                    MarketsCollec.Item(1).Item("runners").Item(Index).Item ("selectionId")
                    
                    MarketsCollec.Item(1).Item("runners").Item(Index).Item ("handicap")
                    MarketsCollec.Item(1).Item("runners").Item(Index).Item ("totalMatched")
                Next
        
                If WorksheetFunction.IsEven(MaxIndex) Then MaxIndex = MaxIndex - 1
                    StartLoop = MaxIndex
                    EndLoop = MaxIndex + 3

        End If

'// Can now populate market cells
'//for each runner in this market
For Index = StartLoop To EndLoop Step 1

With MarketsCollec.Item(1).Item("runners")

'output the Market id
'ActiveSheet.Cells(OPRow + index - 1, columncounter).Value = MarketsCollec.Item(1).Item("marketid")
'then the runner name
'ActiveSheet.Cells(OPRow + index - 1, OPColumn + 1).Value = .Item(index).Item("selectionId")
Dim Dummy: Dummy = MarketsCollec.Item(1).Item("runners").Item(Index).Item(1)

    '// Straight top of book back prices
    ActiveSheet.Cells(OPRow + Index - MaxIndex, OPColumn + 2).ClearContents
    ActiveSheet.Cells(OPRow + Index - MaxIndex, OPColumn + 3).ClearContents
    ActiveSheet.Cells(OPRow + Index - MaxIndex, OPColumn - 1).ClearContents
    
    
'    If .Item(Index).Item("ex").Item("availableToLay").Item(1).Item("price") > 0 Then
'        ActiveSheet.Cells(OPRow + Index - 1, OPColumn + 2).Value = GetBetfairIncrement(.Item(Index).Item("ex").Item("availableToLay").Item(1).Item("price"), "Down")
'    End If
'
'    If .Item(Index).Item("ex").Item("availableToBack").Item(1).Item("price") > 0 Then
'        ActiveSheet.Cells(OPRow + Index - 1, OPColumn + 3).Value = GetBetfairIncrement(.Item(Index).Item("ex").Item("availableToBack").Item(1).Item("price"), "Up")
'    End If
    
    ActiveSheet.Cells(OPRow + Index - MaxIndex, OPColumn + 2).Value = .Item(Index).Item("ex").Item("availableToBack").Item(1).Item("price")
    ActiveSheet.Cells(OPRow + Index - MaxIndex, OPColumn + 3).Value = .Item(Index).Item("ex").Item("availableToLay").Item(1).Item("price")
    ActiveSheet.Cells(OPRow + Index - MaxIndex, OPColumn - 1).Value = .Item(Index).Item("selectionId")
    If TotalRunnersInThisMarket > 20 Then
        ActiveSheet.Cells(OPRow + Index - MaxIndex, OPColumn + 1).Value = .Item(Index).Item("handicap")
    End If


End With
Next
On Error GoTo 0

End Function

Sub OutputListCurrentOrders(ByRef ListCurrentOrdersResult As Object, ByVal ThisSelectionId As String, OutputRow As Integer, Side As String)

Dim BetsOnThisMarket As Integer
Dim BetCounter As Integer: BetCounter = 0
Dim myCount As Integer
Dim PutItInColumn As Integer: PutItInColumn = 0

Dim TotalReturnBack As Double: TotalReturnBack = 0
Dim TotalSizeBack As Double: TotalSizeBack = 0
Dim TotalAveragePriceMatchedBack As Double: TotalAveragePriceMatchedBack = 0

Dim TotalReturnLay As Double: TotalReturnLay = 0
Dim TotalSizeLay As Double: TotalSizeLay = 0
Dim TotalAveragePriceMatchedLay As Double: TotalAveragePriceMatchedLay = 0


Dim Spent As Double: Spent = 0


BetsOnThisMarket = ListCurrentOrdersResult.Item("currentOrders").Count

'need to loop through all the CurrentOrders for this market BUT in the meantime, the quick and dirty solution is that
'I assume there's only 1 BET on each market so will just output the values

For myCount = 1 To BetsOnThisMarket

With ListCurrentOrdersResult.Item("currentOrders").Item(myCount)

'for each order, check the selection and side
If .Item("selectionId") = ThisSelectionId Then

    BetCounter = BetCounter + 1
'BetID=1
Cells(OutputRow, 21 + 7 * BetCounter) = .Item("betId")
'Side=7
Cells(OutputRow, 22 + 7 * BetCounter) = .Item("side")
'PriceSize =5 OrderPrice=1
Cells(OutputRow, 23 + 7 * BetCounter) = .Item("priceSize").Item("price")
'PriceSize =5 OrderSize=2
Cells(OutputRow, 24 + 7 * BetCounter) = .Item("priceSize").Item("size")
'PriceMatched =13
Cells(OutputRow, 25 + 7 * BetCounter) = .Item("averagePriceMatched")
'SizeMatched=14
Cells(OutputRow, 26 + 7 * BetCounter) = .Item("sizeMatched")
'SizeRemaining=15
Cells(OutputRow, 27 + 7 * BetCounter) = .Item("sizeRemaining")

    If .Item("side") = "BACK" Then
        PutItInColumn = 8
        TotalReturnBack = TotalReturnBack + .Item("averagePriceMatched") * .Item("sizeMatched")
        TotalSizeBack = TotalSizeBack + .Item("sizeMatched")
    
        If TotalSizeBack > 0 Then
            TotalAveragePriceMatchedBack = TotalReturnBack / TotalSizeBack
        Else
            TotalAveragePriceMatchedBack = 0
        End If
    
                If PutItInColumn > 0 Then
                    Cells(OutputRow, PutItInColumn) = TotalAveragePriceMatchedBack '.Item("averagePriceMatched")
                    Cells(OutputRow, PutItInColumn + 1) = TotalSizeBack + 0.01 '.Item("sizeMatched") + 0.01 'do this so there is a value >0 in the cell
                End If
    
        ElseIf .Item("side") = "LAY" Then
                        PutItInColumn = 10
                        TotalReturnLay = TotalReturnLay + .Item("averagePriceMatched") * .Item("sizeMatched")
                        TotalSizeLay = TotalSizeLay + .Item("sizeMatched")
                    
                        If TotalSizeLay > 0 Then
                            TotalAveragePriceMatchedLay = TotalReturnLay / TotalSizeLay
                        Else
                            TotalAveragePriceMatchedLay = 0
                        End If
                    
                                If PutItInColumn > 0 Then
                                    Cells(OutputRow, PutItInColumn) = TotalAveragePriceMatchedLay '.Item("averagePriceMatched")
                                    Cells(OutputRow, PutItInColumn + 1) = TotalSizeLay + 0.01 '.Item("sizeMatched") + 0.01 'do this so there is a value >0 in the cell
                                End If
                    
    
    End If

End If

End With

Next
'End If


End Sub

Sub OutputListCurrentOrdersNew(ByRef ListCurrentOrdersResult As Object, ByVal MyOrderAction As String, ByRef MarketIDString)

Dim BetsOnThisMarket As Integer
Dim OutputRow As Integer
Dim TruncMarketId As String
Dim startRow As Integer
Dim BetCounter As Integer: BetCounter = 0
Dim myCount, Index, Selections As Integer
Dim PutItInColumn As Integer: PutItInColumn = 0
Dim Backprice, LayPrice As Double

Dim CurrentSelection As String
Dim CurrentPrice As Double
Dim CurrentSizeRemaining As Double
Dim CurrentSide As String
Dim CurrentBetID As String
Dim CurrentMarketID As String

Dim ThisMarketName As String

Dim TotalReturnBack As Double: TotalReturnBack = 0
Dim TotalSizeBack As Double: TotalSizeBack = 0
Dim TotalAveragePriceMatchedBack As Double: TotalAveragePriceMatchedBack = 0
Dim EndPoint As String: EndPoint = Cells(4, 1).Value
Dim MyBackLimit As Double
Dim MyLayLimit As Double

Dim ListMarketBookResult As Object

Dim TotalReturnLay As Double: TotalReturnLay = 0
Dim TotalSizeLay As Double: TotalSizeLay = 0
Dim TotalAveragePriceMatchedLay As Double: TotalAveragePriceMatchedLay = 0

Dim Request
Dim Success

On Error GoTo ErrorHandler

Dim Spent As Double: Spent = 0

BetsOnThisMarket = ListCurrentOrdersResult.Item("currentOrders").Count
    MyBackLimit = Sheets("Example").Cells(GetNamedRngRow("BackLimit", "Example"), GetNamedRngColumn("BackLimit", "Example")).Value
        MyLayLimit = Sheets("Example").Cells(GetNamedRngRow("LayLimit", "Example"), GetNamedRngColumn("LayLimit", "Example")).Value


                    If ((MyOrderAction = "Cleanse") Or (MyOrderAction = "TopOfferOnly")) Then 'if we want to CLEANSE then we need to get te current MarketBook to use later
                        
                                'Use the MarketIDList to get the MarketBook
                                Request = MakeJsonRpcRequestString(ListMarketBookMethod, GetListMarketBookRequestStringV(MarketIDString))
                                Dim ListMarketBookResponse As String: ListMarketBookResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
                                AppendToLogFile "List Market Book " & ListMarketBookResponse
    
                                Set ListMarketBookResult = ParseJsonRpcResponseToCollection(ListMarketBookResponse)
                                'GetListMarketBookRequestStringV

                    End If


For myCount = 1 To BetsOnThisMarket

    With ListCurrentOrdersResult.Item("currentOrders").Item(myCount)
    
            'for each order, check the selection and side
            TruncMarketId = .Item("marketId")
            'Do While Right(TruncMarketId, 1) = "0"
           '
           '     TruncMarketId = Left(TruncMarketId, Len(TruncMarketId) - 1)
           ' Loop
            
            'If Right(.Item("marketId"), 2) = "00" Then
             '   TruncMarketId = Left(.Item("marketId"), Len(.Item("marketId")) - 2)
            'ElseIf Right(.Item("marketId"), 1) = "0" Then
             '   TruncMarketId = Left(.Item("marketId"), Len(.Item("marketId")) - 1)
            'Else
             '   TruncMarketId = .Item("marketId")
            'End If
            
                startRow = FindTheValue(ActiveSheet, 1, TruncMarketId, 1)
                'startRow = 1
                                OutputRow = FindTheValue(ActiveSheet, 3, .Item("selectionId"), startRow - 1)
                                
            
            ThisMarketName = GetMarketNameForMarketID(TruncMarketId) 'needed for NFL matches
 
                If (ThisMarketName = "Handicap") Or (ThisMarketName = "Total Points") Or (ThisMarketName = "Total Goals") Then
                    OutputRow = FindTheValue(ActiveSheet, 5, .Item("handicap"), OutputRow - 1)
                End If
            
            
            If OutputRow <> 0 Then 'it's a valid SelectionID for this Market/Event
                    
                  If MyOrderAction = "List" Then 'do all the following to LIST orders
                    
                        BetCounter = Cells(OutputRow, 49) + 1 'BetCounternow counts thenumber of BETS on THIS selection
                        Cells(OutputRow, 49) = BetCounter
                            If BetCounter > 3 Then BetCounter = 3 'need to capture this error in the future
                                    
                                Cells(OutputRow, 21 + 7 * BetCounter) = .Item("betId")
                                    Cells(OutputRow, 22 + 7 * BetCounter) = .Item("side")
                                        Cells(OutputRow, 23 + 7 * BetCounter) = .Item("priceSize").Item("price")
                                            Cells(OutputRow, 24 + 7 * BetCounter) = .Item("priceSize").Item("size")
                                                Cells(OutputRow, 25 + 7 * BetCounter) = .Item("averagePriceMatched")
                                                    Cells(OutputRow, 26 + 7 * BetCounter) = .Item("sizeMatched")
                                                        Cells(OutputRow, 27 + 7 * BetCounter) = .Item("sizeRemaining")
                                            
                                    If .Item("side") = "BACK" Then
                                                PutItInColumn = 8
                                                TotalReturnBack = Cells(OutputRow, 8) * Cells(OutputRow, 9) + .Item("averagePriceMatched") * .Item("sizeMatched")
                                                TotalSizeBack = Cells(OutputRow, 9) + .Item("sizeMatched")
                                    
                                        If TotalSizeBack > 0 Then
                                            TotalAveragePriceMatchedBack = TotalReturnBack / TotalSizeBack
                                        Else
                                            TotalAveragePriceMatchedBack = .Item("priceSize").Item("price") 'we will put in the most recent amount requested
                                        End If
                                    
                                                If PutItInColumn > 0 Then
                                                    Cells(OutputRow, PutItInColumn) = TotalAveragePriceMatchedBack '.Item("averagePriceMatched")
                                                    Cells(OutputRow, PutItInColumn + 1) = TotalSizeBack + 0 '.Item("sizeMatched") + 0.01 'do this so there is a value >0 in the cell
                                                End If
                                    
                                    ElseIf .Item("side") = "LAY" Then
                                                        PutItInColumn = 10
                                                        TotalReturnLay = Cells(OutputRow, 10) * Cells(OutputRow, 11) + .Item("averagePriceMatched") * .Item("sizeMatched")
                                                        TotalSizeLay = Cells(OutputRow, 11) + .Item("sizeMatched")
                                                    
                                                        If TotalSizeLay > 0 Then
                                                            TotalAveragePriceMatchedLay = TotalReturnLay / TotalSizeLay
                                                        Else
                                                            TotalAveragePriceMatchedLay = .Item("priceSize").Item("price") 'we will put in the most recent amount requested
                                                        End If
                                                    
                                                                If PutItInColumn > 0 Then
                                                                    Cells(OutputRow, PutItInColumn) = TotalAveragePriceMatchedLay '.Item("averagePriceMatched")
                                                                    Cells(OutputRow, PutItInColumn + 1) = TotalSizeLay + 0   '.Item("sizeMatched") + 0.01 'do this so there is a value >0 in the cell
                                                                End If
                                    End If
                    
                        ElseIf MyOrderAction = "Cleanse" Then
                        
                                'Use the MarketBookResult and find the Order within this to determine if it is MINE that is the next in line
                                CurrentSelection = .Item("selectionId")
                                CurrentPrice = .Item("priceSize").Item("price")
                                CurrentSide = .Item("side")
                                CurrentBetID = .Item("betId")
                                CurrentMarketID = .Item("marketId")
                                CurrentSizeRemaining = .Item("sizeRemaining")
                                'OutputRow = FindTheValue(ActiveSheet, 3, .Item("selectionId")) ' get the ROW that is associated with this particular order
                                
                                'TruncMarketId = Left(.Item("marketId"), 9)
                                    'startRow = FindTheValue(ActiveSheet, 1, TruncMarketId, 1)
                                    startRow = 1
                                        OutputRow = FindTheValue(ActiveSheet, 3, .Item("selectionId"), startRow)
                                
                                For Index = 1 To ListMarketBookResult.Count Step 1
                                        For Selections = 1 To ListMarketBookResult.Item(Index).Item("runners").Count Step 1
                                            If ListMarketBookResult.Item(Index).Item("runners").Item(Selections).Item("selectionId") = CurrentSelection Then
                                                'CurrentMarketID = ListMarketBookResult.Item(Index).Item("marketID")
                                                    If (ListMarketBookResult.Item(Index).Item("runners").Item(Selections).Item("ex").Item("availableToBack").Count) > 0 Then
                                                        Backprice = ListMarketBookResult.Item(Index).Item("runners").Item(Selections).Item("ex").Item("availableToBack").Item(1).Item("price")
                                                    End If
                                                    If (ListMarketBookResult.Item(Index).Item("runners").Item(Selections).Item("ex").Item("availableToLay").Count) > 0 Then
                                                        LayPrice = ListMarketBookResult.Item(Index).Item("runners").Item(Selections).Item("ex").Item("availableToLay").Item(1).Item("price")
                                                    End If
                                            End If
                                        Next
                                Next
                        
                        'NEW 27th Sep 2015
                                    If CurrentSide = "BACK" Then
                                        If ((LayPrice < CurrentPrice And CurrentSizeRemaining > 0) Or (CurrentPrice / Cells(OutputRow, 15).Value - 1 < (MyBackLimit * 0.9))) Then 'if my Current Price is NOT the best price on the market then Kill it because I've been outbid
                                            'Kill This Order
                                            Success = CancelOrdersForBetID(CurrentMarketID, CurrentBetID)
                                        End If
                                    ElseIf CurrentSide = "LAY" Then
                                        If ((Backprice > CurrentPrice And CurrentSizeRemaining > 0) Or (1 - (CurrentPrice / Cells(OutputRow, 15).Value) < (MyLayLimit * 0.9))) Then 'if my Current Price is NOT the best price on the market then Kill it
                                            'Kill This Order
                                            Success = CancelOrdersForBetID(CurrentMarketID, CurrentBetID)
                                        End If
                                    
                                    End If
                        
                                                
                                                ElseIf MyOrderAction = "TopOfferOnly" Then
                                                
                                                        'Use the MarketBookResult and find the Order within this to determine if it is MINE that is the next in line
                                                        CurrentSelection = .Item("selectionId")
                                                        CurrentPrice = .Item("priceSize").Item("price")
                                                        CurrentSide = .Item("side")
                                                        CurrentBetID = .Item("betId")
                                                        CurrentMarketID = .Item("marketId")
                                                        CurrentSizeRemaining = .Item("sizeRemaining")
                                                        OutputRow = FindTheValue(ActiveSheet, 3, .Item("selectionId"), 1) ' get the ROW that is associated with this particular order
                                                        
                                                        For Index = 1 To ListMarketBookResult.Count Step 1
                                                        'x = ListMarketBookResult.Item().Item("runners")
                                                                For Selections = 1 To ListMarketBookResult.Item(Index).Item("runners").Count Step 1
                                                                    If ListMarketBookResult.Item(Index).Item("runners").Item(Selections).Item("selectionId") = CurrentSelection Then
                                                                        'CurrentMarketID = ListMarketBookResult.Item(Index).Item("marketID")
                                                                            If (ListMarketBookResult.Item(Index).Item("runners").Item(Selections).Item("ex").Item("availableToBack").Count) > 0 Then
                                                                                Backprice = ListMarketBookResult.Item(Index).Item("runners").Item(Selections).Item("ex").Item("availableToBack").Item(1).Item("price")
                                                                            End If
                                                                            If (ListMarketBookResult.Item(Index).Item("runners").Item(Selections).Item("ex").Item("availableToLay").Count) > 0 Then
                                                                                LayPrice = ListMarketBookResult.Item(Index).Item("runners").Item(Selections).Item("ex").Item("availableToLay").Item(1).Item("price")
                                                                            End If
                                                                    End If
                                                                Next
                                                        Next
                                                
                                                'NEW 27th Sep 2015
                                                            If CurrentSide = "BACK" Then
                                                                If ((LayPrice <> CurrentPrice) And CurrentSizeRemaining > 0) Then  'if my Current Price is NOT the best price on the market then Kill it because I've been outbid
                                                                    'Kill This Order
                                                                    Success = CancelOrdersForBetID(CurrentMarketID, CurrentBetID)
                                                                End If
                                                            ElseIf CurrentSide = "LAY" Then
                                                                If ((Backprice <> CurrentPrice) And CurrentSizeRemaining > 0) Then  'if my Current Price is NOT the best price on the market then Kill it
                                                                    'Kill This Order
                                                                    Success = CancelOrdersForBetID(CurrentMarketID, CurrentBetID)
                                                                End If
                                                            
                                                            End If
                        
                        
                        
                        ElseIf MyOrderAction = "Purge" Then
                                'do things here related to PURGING order
                                CurrentSizeRemaining = .Item("sizeRemaining")
                                If CurrentSizeRemaining > 0 Then
                                CurrentMarketID = .Item("marketId")
                                CurrentBetID = .Item("betId")
                                Success = CancelOrdersForBetID(CurrentMarketID, CurrentBetID)
                                End If
                        End If 'stuff to LIST, CLEANSE or PURGE orders
                    
            End If 'Check OutputRow <>)
            
    End With 'each BetsOnThisMarket

Next 'BetsOnThisMarket

On Error GoTo 0
Exit Sub

ErrorHandler:
    HandleError "OutputListCurrentOrdersNew"
    Resume Next


End Sub

Sub OutputlistEvents(ByRef ListEventsResult As Object, ByRef OutputRow As Integer, ByRef OutputColumn As Integer)
    Dim Index As Integer
    For Index = 1 To ListEventsResult.Count Step 1
        Sheet4.Cells(OutputRow, OutputColumn + Index).Value = ListEventsResult.Item(Index).Item("event").Item("id")
        Sheet4.Cells(OutputRow + 1, OutputColumn + Index).Value = ListEventsResult.Item(Index).Item("event").Item("name")
        
    Next
End Sub
Sub OutputCompListEvents(ByRef ListEventsResult As Object, ByVal OutputRow As Integer, ByVal OutputColumn As Integer, ByVal VertHoriz As String)

    Dim Index As Integer
    
    If VertHoriz = "Vert" Then
        For Index = 1 To ListEventsResult.Count Step 1
            Cells(OutputRow + Index, OutputColumn).Value = ListEventsResult.Item(Index).Item("event").Item("id")
            Cells(OutputRow + Index, OutputColumn + 1).Value = ListEventsResult.Item(Index).Item("event").Item("name")
        Next
    ElseIf VertHoriz = "Horiz" Then
        For Index = 1 To ListEventsResult.Count Step 1
            Cells(OutputRow, OutputColumn + Index).Value = ListEventsResult.Item(Index).Item("event").Item("id")
            Cells(OutputRow + 1, OutputColumn + Index).Value = ListEventsResult.Item(Index).Item("event").Item("name")
        Next
    End If
    
End Sub
Sub OutputListSoccerMatchResult(ByRef ListSoccerMatchResult As Object, ByRef OutputRow As Integer, ByRef OutputColumn As Integer, ByVal VertHoriz As String)
    Dim Index As Integer
    Dim runindex As Integer
    
    For Index = 1 To ListSoccerMatchResult.Count Step 1
        Cells(OutputRow, OutputColumn + Index).Value = ListSoccerMatchResult.Item(Index).Item("marketId")
        Cells(OutputRow + 1, OutputColumn + Index).Value = ListSoccerMatchResult.Item(Index).Item("marketName")
        
        For runindex = 1 To ListSoccerMatchResult.Item(Index).Item("runners").Count Step 1
            Cells(OutputRow + 2 + (runindex - 1) * 4, OutputColumn + Index).Value = ListSoccerMatchResult.Item(Index).Item("runners").Item(runindex).Item("selectionId")
            Cells(OutputRow + 3 + (runindex - 1) * 4, OutputColumn + Index).Value = ListSoccerMatchResult.Item(Index).Item("runners").Item(runindex).Item("runnerName")
            Cells(OutputRow + 4 + (runindex - 1) * 4, OutputColumn + Index).Value = ListSoccerMatchResult.Item(Index).Item("runners").Item(runindex).Item("handicap")
            Cells(OutputRow + 5 + (runindex - 1) * 4, OutputColumn + Index).Value = ListSoccerMatchResult.Item(Index).Item("runners").Item(runindex).Item("sortPriority")
        Next
    Next
End Sub

Sub OutputListMatchResult(ByRef ListMatchResult As Object, ByRef OutputRow As Integer, ByRef OutputColumn As Integer, ByVal VertHoriz As String)
    Dim Index As Integer
    Dim runindex As Integer
    Dim ThisName As String
    Dim ThisMarketType As String
    Dim OverUnderMargin As Double
    Dim TotalPointsCounter As Integer
    Dim SpreadCounter As Integer
    Dim TotalGoalsCounter As Integer
    Dim MyPoints As Integer
    Dim EndMyPoints As Integer
    Dim ThisSpread As Double
    Dim myCounter As Integer
    Dim UnderOverCounter As Integer
    Dim UnderOverStart As Integer
    
TotalPointsCounter = 0
SpreadCounter = 0
TotalGoalsCounter = 0
UnderOverCounter = 0
    
If VertHoriz = "Horiz" Then
    
    For Index = 1 To ListMatchResult.Count Step 1
        Cells(OutputRow, OutputColumn + Index).Value = ListMatchResult.Item(Index).Item("marketId")
        Cells(OutputRow + 1, OutputColumn + Index).Value = ListMatchResult.Item(Index).Item("marketName")
        
        For runindex = 1 To ListMatchResult.Item(Index).Item("runners").Count Step 1
            'Cells(OutputRow + 2 + (runindex - 1) * 2, OutputColumn + index).Value = ListMatchResult.Item(index).Item("runners").Item(runindex).Item("selectionId")
            Cells(OutputRow + 2 + (runindex - 1) * 1, OutputColumn + Index).Value = ListMatchResult.Item(Index).Item("runners").Item(runindex).Item("runnerName")
            'Cells(OutputRow + 4 + (runindex - 1) * 4, OutputColumn + index).Value = ListMatchResult.Item(index).Item("runners").Item(runindex).Item("handicap")
            'Cells(OutputRow + 5 + (runindex - 1) * 4, OutputColumn + index).Value = ListMatchResult.Item(index).Item("runners").Item(runindex).Item("sortPriority")
        Next
    Next

ElseIf VertHoriz = "Sheet" Then

'On Error Resume Next

If ListMatchResult Is Nothing Then GoTo Err_Handler

    For Index = 1 To ListMatchResult.Count Step 1
        ThisName = ListMatchResult.Item(Index).Item("marketName")
        'ThisMarketType = ListMatchResult.Item(Index).Item("marketName")
        
        If ThisName = "Match Odds" Then
            'Soccer and AFL
            'output to Match Odds
                    OutputRow = 8
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
                    
        ElseIf ThisName = "Regular Time Match Odds" Then
                    OutputRow = 18
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
                    
        ElseIf ThisName = "Moneyline" Then
                    OutputRow = 8
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
        
        ElseIf Left(ThisName, 6) = "Over/U" Then
            'Soccer ONLY
            'output to OverUnder section
            OverUnderMargin = Val(Mid(ThisName, 12, 3))
            If OverUnderMargin < 8 Then 'only interested in 0.5 up to 7.5 at this stage
                    OutputRow = 11 + (2 * (OverUnderMargin - 0.5))
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
            End If
            
        ElseIf Left(ThisName, 5) = "Under" Then
            'AFL ONLY
            'output to Under/Over section
            UnderOverStart = 14 'ÁFL
            If Cells(3, 1).Value = "NFL" Then UnderOverStart = 17 'NFL
            
            UnderOverCounter = UnderOverCounter + 2
            
            OverUnderMargin = Val(Mid(ThisName, 12, 5)) '5 char always assumes >100 - must change this
                    OutputRow = UnderOverStart + UnderOverCounter
                    OutputColumn = 1
                    
                If UnderOverCounter <= 4 Then 'make sure only 2 used
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
                        Cells(OutputRow + 1, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
                    Cells(OutputRow, OutputColumn + 4).Value = OverUnderMargin
                    Cells(OutputRow + 1, OutputColumn + 4).Value = OverUnderMargin
                End If
        
        ElseIf ThisName = "Correct Score" Then
            'Soccer ONLY
            'output to Correct Score section
                    OutputRow = 27
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
        
        ElseIf Left(ThisName, 4) = "Both" Then
            'Soccer ONLY
            'output to Both Teams to Score section
                    OutputRow = 46
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
        
        ElseIf (Left(ThisName, 4) = "Winn" And Right(ThisName, 2) = ".5") Then
            'it's a winning margin type with NUMBERS meaning
            'AFL ONLY
            'output to Both Teams to Score section
            If ThisName <> "Winning Margin" Then 'this is used in Soccer now too
                    If Right(ThisName, 4) = "24.5" Then
                        OutputRow = 20
                        OutputColumn = 1
                    ElseIf Right(ThisName, 4) = "39.5" Then
                        OutputRow = 25
                        OutputColumn = 1
                    End If
                        Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                        Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
            End If
            
        ElseIf ThisName = "Total Game Score" Then
            'output to Total Game Score section
                    OutputRow = 30
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
        
        ElseIf ThisName = "Winning Margin Spread" Then
            'output to Winning Margin Spreade section
            'AFL only
                    OutputRow = 39
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
         
        ElseIf (ThisName = "Winning Margin") And (Cells(3, 1).Value <> "Soccer") Then 'NFL or AFL
            'output to Winning Margin section
            'AFL only
                    OutputRow = 39 'AFL
                    If Cells(3, 1).Value = "NFL" Then OutputRow = 43 'NFL
                    If Cells(3, 1).Value = "Soccer" Then OutputRow = 55 'Soccer
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
                    
        ElseIf ThisName = "Winning Margin 2" Then 'NFL
            'output to Winning Margin section
            'AFL only
                    OutputRow = 25
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
                    
        ElseIf ThisName = "Total Goals" Then 'Soccer
            'output to Winning Margin section
            'AFL only
                    OutputRow = 48
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
                    
        ElseIf ThisName = "Total Points" Then 'NFL
            'output to Winning Margin section
            'AFL only
                    OutputRow = 14
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
                    
        ElseIf ThisName = "Handicap" Then 'NFL
            'output to Winning Margin section
            'AFL only
                    OutputRow = 10
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
                    
        ElseIf ThisName = "Tri Bet" Then
            'output to Tri Bet Spreade section
            'AFL only
                    OutputRow = 48
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
                    
        'the following WAS an AFL thing and was brought back just for the Grand Final 2016
        'have edited it out for now 29-09-16 by changing POINTS to POINTSXYZ
        
        ElseIf Right(ThisName, 6) = "Points" Then 'note elseif - placed AFTER Total Points in sequence
            'NFL ONLY new for 2017 and gives Total [Team Name] Points
            'output to Total Points section
            If Cells(3, 1).Value = "NFL" Then 'NFL game - can be NFL which I don't care about
                'If (ThisName <> "Half Time Total Points" And ThisName <> "First To 25 Points") Then
                 If (Left(ListMatchResult.Item(Index).Item("marketName"), 5) = "Total") And (ThisName <> "Total Points") Then
                 
                    'must identify home team
                    
                    If (Mid(ListMatchResult.Item(Index).Item("marketName"), 7, 4) = Left(Cells(7, 3), 4)) Then
                        'it's the home team
                        OutputRow = 21
                    Else
                        OutputRow = 23
                    End If
                    
                    'TotalPointsCounter = TotalPointsCounter + 6 'was 2 since was only 2 selections previously
                    'OutputRow = 33 + TotalPointsCounter
                    
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    
                    For myCounter = 0 To 1
                        Cells(OutputRow + myCounter, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
                            'Cells(OutputRow + 1, OutputColumn + 1).Value = ListMatchResult.Item(index).Item("marketName")
                        Cells(OutputRow + myCounter, OutputColumn + 3).Value = ListMatchResult.Item(Index).Item("runners").Item(myCounter + 1).Item("runnerName")
                    Next myCounter
                End If
             End If
             
        ElseIf ((Right(ThisName, 3) = "pts") Or (Mid(ThisName, Len(ThisName) - 3, 1) = "+")) Then
            'output to Spread section OR to Under/Over Section
                    SpreadCounter = SpreadCounter + 2
                    OutputRow = 8 + SpreadCounter
                If OutputRow <= 14 Then 'this is to ensure only THREE markets are used
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
                        Cells(OutputRow + 1, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
                    'cells(OutputRow,OutputColumn+3.value=cells(OutputRow,OutputColumn+1.value
                    MyPoints = InStr(Cells(OutputRow, OutputColumn + 1).Value, "+")
                    
                    If Mid(ThisName, Len(ThisName) - 3, 1) = "+" Then
                        EndMyPoints = Len(ThisName) + 1
                    Else
                        EndMyPoints = InStr(Cells(OutputRow, OutputColumn + 1).Value, "pts")
                    End If
                    
                    ThisSpread = Val(Mid(Cells(OutputRow, OutputColumn + 1).Value, MyPoints + 1, (EndMyPoints - 1) - (MyPoints)))
                    Cells(OutputRow, OutputColumn + 4).Value = ThisSpread
                End If
                
        ElseIf Left(ThisName, 8) = "Overtime" Then
            'NFL ONLY
            'output to OverUnder section
            OutputRow = 51
                    OutputColumn = 1
                    Cells(OutputRow, OutputColumn).Value = "'" & CStr(ListMatchResult.Item(Index).Item("marketId")) & ""
                    Cells(OutputRow, OutputColumn + 1).Value = ListMatchResult.Item(Index).Item("marketName")
'
        End If 'else do nothing
    Next

End If

Exit_Err_Handler:
    Exit Sub
 
Err_Handler:
'    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
'           "Error Number: " & Err.Number & vbCrLf & _
'           "Error Source: AppendTxt" & vbCrLf & _
'           "Error Description: " & Err.Description, vbCritical, "An Error has Occured!"
    GoTo Exit_Err_Handler

End Sub

Function CreateKeepAliveFile(sText As String, Optional Amount As Double)

Dim TempText, MyKeepAliveServerPath As String
Dim MyMemory As Long

MyKeepAliveServerPath = Sheets("Example").Cells(GetNamedRngRow("KeepAliveServerPath", "Example"), GetNamedRngColumn("KeepAliveServerPath", "Example")).Value

On Error GoTo Err_Handler
    Dim FileNumber As Integer
    Dim sFile As String: sFile = MyKeepAliveServerPath & " " & Amount & ".txt"
 
    FileNumber = FreeFile                   ' Get unused file number
    Open sFile For Append As #FileNumber    ' Connect to the file
    Print #FileNumber, Format(Now(), "dd-mm-yy hh:mm:ss") & " " & sText
    'Print #FileNumber, TempText                ' Append our string
    Print #FileNumber, Chr(10)              ' Chr(10) = LF, 13 = CR
    Close #FileNumber                       ' Close the file

Exit_Err_Handler:
    Exit Function
 
Err_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: AppendTxt" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, "An Error has Occured!"
    GoTo Exit_Err_Handler
    
End Function
Function AppendToLogFile(sText As String, Optional funcText As String, Optional length As Integer)

Dim TempText As String
Dim MyMemory As Long

If length = 999 Then
    TempText = sText & " " & funcText
ElseIf length > 1 Then
    TempText = Left(sText, length) & " " & funcText 'trim it to required length
Else
    TempText = Left(sText, 150) & " " & funcText 'trim it to 150 char by default
End If
'Debug.Print Application.MemoryUsed

On Error GoTo Err_Handler
    Dim FileNumber As Integer
    Dim sFile As String: sFile = ThisWorkbook.Path & "\" & LogFile
 
    FileNumber = FreeFile                   ' Get unused file number
    Open sFile For Append As #FileNumber    ' Connect to the file
    Print #FileNumber, Format(Now(), "dd-mm-yy hh:mm:ss") & " " & TempText
    'Print #FileNumber, TempText                ' Append our string
    Print #FileNumber, Chr(10)              ' Chr(10) = LF, 13 = CR
    Close #FileNumber                       ' Close the file

Exit_Err_Handler:
Exit Function
 
Err_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: AppendTxt" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, "An Error has Occured!"
    GoTo Exit_Err_Handler
    
End Function

Sub DeleteLogFile()
    DeleteFile ThisWorkbook.Path & "\" & LogFile
End Sub

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub

Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function
Function GetMarketTypeFromListMarketTypes(ByVal ThisMarketID As String, ByVal Response As Object) As String
'Takes a MarketID and the ListMarketTypes reponse and finds the MarketType associated with the MarketId

    Dim marketindex As Integer
    Dim Index As Integer
    For marketindex = 1 To Response.Count Step 1
        
        Dim marketid: marketid = Response.Item(marketindex).Item("marketId")
        Dim marketName: marketName = Response.Item(marketindex).Item("marketType")
    
        If marketid = ThisMarketID Then
            GetMarketTypeFromListMarketTypes = marketid
            Exit For
        End If
    Next
        
End Function

Sub TestMyFunction()

Dim ThisMarketID As String: ThisMarketID = "2.100993688"
Dim ThisSelectionId As String: ThisSelectionId = "39980"

Dim Request
Dim i

Dim ListMarketCatalogueResponse As String
Dim ListMarketCatalogueResult As Object

If IsNumeric(ThisMarketID) Then

        '// Get request string for our filter...
        Request = Trim(MakeJsonRpcRequestString(ListMarketCatalogueMethod, GetListMarketCatalogueRequestStringUsingMarketID(ThisMarketID)))
        
        '// Check for null string. Even though unlikely, still a valid test else we would have to deal with additional web errors later in any case
        If Request <> vbNullString Then
        
            '// Send request to betfair and receive response...
            
            ListMarketCatalogueResponse = SendRequest(GetJsonRpcUrl(), GetAppKey(), GetSession(), Request)
            
            '// Get a collection / dictionary object (SoccerMatchCatalogue) of parsed JSON arrays and items...
            
            Set ListMarketCatalogueResult = ParseJsonRpcResponseToCollection(ListMarketCatalogueResponse)
            
            '// Now test that an object was successfully returned
            If IsObject(ListMarketCatalogueResult) Then
                For i = 1 To ListMarketCatalogueResult.Item(1).Item("runners").Count
                    If ListMarketCatalogueResult.Item(1).Item("runners").Item(i).Item("selectionId") = ThisSelectionId Then
                        
                    End If
                Next i
            End If
        End If
        
End If
End Sub
Public Function FindTheValue(searchSheet As Worksheet, inColumn As Integer, thisFind, startRow As Integer)

Dim c

'With searchSheet.Columns(inColumn)
With searchSheet.Range(Cells(startRow, inColumn), Cells(MaxSheetlength, inColumn))
    Set c = .Find(thisFind, LookIn:=xlValues, LookAt:=xlWhole)
    If Not c Is Nothing Then
        FindTheValue = c.Row
    Else
        FindTheValue = 0
    End If
End With

End Function
Function PutTheResult(destSheet As Worksheet, inRow As Long, inColumn As Long, _
Optional putWhat As Variant, Optional Force As Boolean)

If IsMissing(putWhat) Then putWhat = Format(Now(), "dd-mmm")
If IsMissing(Force) Then Force = False


If Force = True Then
    destSheet.Cells(inRow, inColumn).Value = putWhat 'just write the date if that's what we want
Else
    If IsDate(putWhat) Then
            If putWhat > destSheet.Cells(inRow, inColumn).Value Then
                destSheet.Cells(inRow, inColumn).Value = putWhat
            End If 'only enter if greater than date already there
    Else
          destSheet.Cells(inRow, inColumn).Value = putWhat 'else, if it's not a date, just enter it
    End If
End If

End Function

Function GetHeaderColumn(Header)

Dim c

With ActiveSheet.rows(1)
    Set c = .Find(Header, LookIn:=xlValues)
    If Not c Is Nothing Then
        GetHeaderColumn = c.Column
    Else
        GetHeaderColumn = 0
    End If
End With

End Function

Function LoadHeaderColumns()

FromLatCol = GetHeaderColumn("FromLat")
FromLongCol = GetHeaderColumn("FromLong")
ToLatCol = GetHeaderColumn("ToLat")
ToLongCol = GetHeaderColumn("ToLong")
WPLat1Col = GetHeaderColumn("WPLat1")
WPLong1Col = GetHeaderColumn("WPLong1")
WPLat2Col = GetHeaderColumn("WPLat2")
WPLong2Col = GetHeaderColumn("WPLong2")
WPLat3Col = GetHeaderColumn("WPLat3")
WPLong3Col = GetHeaderColumn("WPLong3")
WPLat4Col = GetHeaderColumn("WPLat4")
WPLong4Col = GetHeaderColumn("WPLong4")
WPLat5Col = GetHeaderColumn("WPLat5")
WPLong5Col = GetHeaderColumn("WPLong5")
StemCol = GetHeaderColumn("Stem")

End Function
Function GetDataRow(DataToFind, ColumnToLookIn As Long)

Dim c

With ActiveSheet.Columns(ColumnToLookIn)
    Set c = .Find(DataToFind, LookIn:=xlValues)
    If Not c Is Nothing Then
        GetDataRow = c.Row
    Else
        GetDataRow = 0 'not found
    End If
End With

End Function
Function CheckSheetConditionsToBet()

Dim MaximumMarketBP As Boolean
Dim AverageMarketBP As Boolean
Dim SufficientPrices As Boolean
Dim SolverWasNotReset As Boolean
Dim Opportunities As Boolean
Dim Realistic As Boolean
Dim StDevRealistic As Boolean
Dim CountryCodeGood As Boolean


Dim ValidPrices As Integer
Dim myPrice As Range

Dim RealisticMean As Double
Dim RealisticMult As Double
Dim RealisticStDev As Double
Dim ValidPricesLimit As Double

If Cells(3, 1) = "NFL" Then
    RealisticMean = 50
    RealisticMult = 2
    ValidPricesLimit = 40
ElseIf Cells(3, 1) = "AFL" Then
    RealisticMean = 120 'less than
    RealisticMult = 1.2
    ValidPricesLimit = 55
    RealisticStDev = 34.5 '= or - 3 = 31.5 to 37.5
    
Else
    RealisticMean = 4
    RealisticMult = 0.7
    ValidPricesLimit = 55
End If


'Maximum Market Book percentage
If (Cells(5, 15) < 2) Then MaximumMarketBP = True Else MaximumMarketBP = False
'Average Market Book percentage
If (Cells(5, 14) < 1.5) Then AverageMarketBP = True Else AverageMarketBP = False
'Solver Reset
If (Cells(3, 9) <> "Reset") Then SolverWasNotReset = True Else SolverWasNotReset = False
'Too many opportunities
If (Cells(5, 16) <= 7) Then Opportunities = True Else Opportunities = False
'Realistic Condition
If ((Cells(3, 5) < RealisticMean) And Cells(3, 7) < RealisticMean And Cells(4, 9) < RealisticMult) Then Realistic = True Else Realistic = False
'Blank Country
If (Cells(5, 3) = "XYZ") Then CountryCodeGood = False Else CountryCodeGood = True 'made this XYZ instead of BLANK on 8th Mar 2016
'Realistic StDev
    If Cells(3, 1) = "AFL" Then
        If Cells(4, 14) < RealisticStDev + 3.5 And Cells(4, 14) > RealisticStDev - 3.5 Then
            RealisticStDev = True
        Else
            RealisticStDev = False
        End If
    Else
        RealisticStDev = True
    End If

ValidPrices = 0
For Each myPrice In Range(Cells(8, 6), Cells(52, 7))
    If myPrice > 0 Then ValidPrices = ValidPrices + 1
Next

'Valid Prices
If ValidPrices > ValidPricesLimit Then SufficientPrices = True Else SufficientPrices = False

'Check all SEVEN conditions are OK
If (MaximumMarketBP And AverageMarketBP And SufficientPrices And SolverWasNotReset And Opportunities And Realistic And CountryCodeGood And RealisticStDev) Then

    CheckSheetConditionsToBet = True
    Cells(2, 1) = True

Else

    CheckSheetConditionsToBet = False
    Cells(2, 1) = False

End If

End Function

Sub PlaceOrdersAutomatically(WithIncrement As Boolean, Optional BetType As String)
'sub takes either Back or Lay (default = Back) and places bets on relevant type

On Error GoTo ErrorHandler:

Dim MyCell As Range
Dim CurrentValue As Double
Dim CutOffSoccerOnSheet, MyMaxWin, MyMinBet As Double
Dim MyBackLimit, MyLayLimit, MyUpperLimit As Double
Dim BackString, LayString As String

If IsMissing(BetType) Then BetType = "Back"
If BetType = "" Then BetType = "Back"

'The definition of Back or Lay in this procedure simply defines which COLUMN is selected to act upon
'Once the cell with the correct data is selected, we just need to call the plain PlaceOrders sub

CutOffSoccerOnSheet = Cells(GetNamedRngRow("CutOffSoccer"), GetNamedRngColumn("CutOffSoccer")).Value
MyMaxWin = Sheets("Example").Cells(GetNamedRngRow("MaxWin", "Example"), GetNamedRngColumn("MaxWin", "Example")).Value
MyMinBet = Sheets("Example").Cells(GetNamedRngRow("MinBet", "Example"), GetNamedRngColumn("MinBet", "Example")).Value
MyBackLimit = Sheets("Example").Cells(GetNamedRngRow("BackLimit", "Example"), GetNamedRngColumn("BackLimit", "Example")).Value
MyLayLimit = Sheets("Example").Cells(GetNamedRngRow("LayLimit", "Example"), GetNamedRngColumn("LayLimit", "Example")).Value
MyUpperLimit = Sheets("Example").Cells(GetNamedRngRow("UpperLimit", "Example"), GetNamedRngColumn("UpperLimit", "Example")).Value

If BetType = "Back" Then
    Range(Cells(8, 20), Cells(AFLLength, 20)).Select
ElseIf BetType = "Lay" Then
    Range(Cells(8, 21), Cells(AFLLength, 21)).Select
End If
'thisMax = Cells(3, 24).Value
    
BackString = "=MAX(ROUND(" & MyMaxWin & "/(RC[-3]-1),0)," & MyMinBet & ")"
LayString = "=MAX(ROUND(" & MyMaxWin & "/(RC[-4]-1),0)," & MyMinBet & ")"
    
    For Each MyCell In Selection
        If MyCell.Value = "" Then
            CurrentValue = 0
        Else
            CurrentValue = MyCell.Value
        End If
        
  If (MyCell.Row > 37) Then
    
    End If
        
        
        If ((CurrentValue > MyBackLimit) And (BetType = "Back") And (CurrentValue < MyUpperLimit)) Then
            
            'temp fill Lay value
                If Cells(MyCell.Row, 7) = "" Then
                    Cells(MyCell.Row, 7).FormulaR1C1 = ("=RC[-1]+VLOOKUP(RC[-1],Increments!C11:C12,2)")
                End If
                
                    If WithIncrement = True Then
                        Cells(MyCell.Row, 8).FormulaR1C1 = ("=RC[-1]-VLOOKUP(RC[-1],Increments!C7:C8,2)")
                    Else
                        Cells(MyCell.Row, 8).FormulaR1C1 = ("=RC[-2]")
                    End If
            
            Cells(MyCell.Row, 9).FormulaR1C1 = BackString
            Cells(MyCell.Row, 9).Select
            Call PlaceOrders
            
        ElseIf ((CurrentValue > MyLayLimit) And (BetType = "Lay") And (CurrentValue < MyUpperLimit)) Then
            
                    If WithIncrement = True Then
                                Cells(MyCell.Row, 10).FormulaR1C1 = ("=RC[-4]+VLOOKUP(RC[-4],Increments!C11:C12,2)")
                            Else
                                Cells(MyCell.Row, 10).FormulaR1C1 = ("=RC[-3]")
                    End If
            
            Cells(MyCell.Row, 11).FormulaR1C1 = LayString
            Cells(MyCell.Row, 11).Select
            Call PlaceOrders
            
        End If
    Next

On Error GoTo 0
Exit Sub

ErrorHandler:
    HandleError "PlaceOrdersAutomatically"
    Resume Next

End Sub
Sub HideSheetsOutsideTimeWindow(Optional ByVal WindowLength As Double)

'takes Windowlength as the number of hours in the future from NOW that we want to hide
'must first unhide ALL then hide the ones that are too far into the future
Dim wsToHide As Worksheet
Dim dummyvar


If IsMissing(WindowLength) Then WindowLength = 3
If WindowLength < 1 Then WindowLength = 15

Call UnhideAll


    For Each wsToHide In Worksheets

        If Left(wsToHide.Name, 1) = "2" Then
        dummyvar = wsToHide.Cells(6, 3).Value
            If (Now() + WindowLength / 24) < wsToHide.Cells(6, 3).Value Then
            
            'ws.Visible = xlSheetVisible
                wsToHide.Visible = xlSheetHidden
            End If
            If Now() > wsToHide.Cells(6, 3).Value Then
                wsToHide.Visible = xlSheetHidden
            End If
        End If

    Next

End Sub
Sub TestCall()
Dim Dummy As Double
Dim Check As Boolean

'
'Call HideSheetsOutsideTimeWindow(15)
Check = CheckSheetConditionsToBet
'GetBetfairIncrement(2.1, "Up")
'dummy = GetTotalMatchedForEventID("275139390")
'Worksheets("27508515").Move before:=Worksheets(3)

'Range(Cells(29, 29), Cells(Dummy, 29)).Select

'ListCurrentOrdersNew "Cleanse"
'ListCurrentOrdersNew "List"
'ListCurrentOrdersNew "Purge"

'Dummy = GetUTCOffset
'Check = IsDST
Call CreateKeepAliveFile("hello", 5 - 3)
'Call PlaceOrdersAutomaticallyNEW(True, "Back")
'Call PlaceOrdersAutomaticallyNEW(True, "Lay")

'Call PlaceOrdersAutomatically("Lay")
'Dim dummy
'Dim myExcel As Object
'
'Set myExcel = CreateObject("Excel.Application")
'
'AppendToLogFile "testing"
'
'ActiveWorkbook.Save
'
'MsgBox "Microsoft Excel is currently using " & _
'    Application.MemoryUsed & " bytes"

End Sub
Function FindNextUpdateTime(JumpOffset As Integer, MinutesOfTheHour As Integer)
'Jump offset is a number of minutesd to add to NOW
'MinutesOfTheHour is a variable length array of integers of minutes that the current time MAY be rounded up to
'e.g. pass 15, 45, the function will round the current time to the NEXT instance of 15 or 45 minutes past the hour

End Function
Sub KeepAlive()

AppendToLogFile "Keep Alive"
Sheets("Example").Cells(6, 8).Value = Now() + 4 / 24

On Error GoTo ErrorHandler

Dim Request
Dim KeepAliveResponse

Request = "ABC"
        
        '// Check for null string. Even though unlikely, still a valid test else we would have to deal with additional web errors later in any case
        If Request <> vbNullString Then
        
            '// Send request to betfair and receive response...
            
            KeepAliveResponse = SendRequest(GetKeepAliveUrl(), GetAppKey(), GetSession())
                If InStr(1, KeepAliveResponse, "SUCCESS", vbTextCompare) = 0 Then
                    Login 'if the KeepAlive function returned 0 then just Login
                End If
        End If

On Error GoTo 0
Exit Sub

ErrorHandler:
    HandleError "KeepAlive"
    Resume Next

End Sub

Sub Login()

AppendToLogFile "Login"
Sheets("Example").Cells(6, 8).Value = Now() + 4 / 24

On Error GoTo ErrorHandler

Dim Request
Dim LoginResponse As String
Dim ssoid As String

Request = "ABC"
        
        '// Check for null string. Even though unlikely, still a valid test else we would have to deal with additional web errors later in any case
        If Request <> vbNullString Then
        
            '// Send request to betfair and receive response...
            
            LoginResponse = SendLoginRequest(GetLoginUrl(), GetAppKey(), "schmoopies", "trocad42")
            
            Dim start, length
                    start = 11
                    length = 44
                    ssoid = Mid(LoginResponse, start, length)
            
            Sheets("Example").Cells(4, 2).Value = ssoid
            Sheets("Example").Cells(36, 2).Value = LoginResponse
                    
        End If

On Error GoTo 0
Exit Sub

ErrorHandler:
    HandleError "Login"
    Resume Next

End Sub

Sub SendEmail(ByVal Available As Double, ByVal Exposure As Double)

Dim aOutlook As Object
Dim aEmail As Object
Dim rngeAddresses As Range, rngeCell As Range, strRecipients As String

Set aOutlook = CreateObject("Outlook.Application")
Set aEmail = aOutlook.CreateItem(0)

'set sheet to find address for e-mails as I have several people to mail to

'            Set rngeAddresses = ActiveSheet.Range("A3:A13")
'            For Each rngeCell In rngeAddresses.Cells
'            strRecipients = strRecipients & ";" & rngeCell.Value
'            Next

'just overwrite this for now to ONE address
strRecipients = "michael.lean@gmail.com"

'set Importance
aEmail.Importance = 2
    'Set Subject
    aEmail.Subject = "Betting Update " & Now()
        'Set Body for mail
        aEmail.Body = "Value = " & Available - Exposure & vbCr & _
                        "Available: " & Available & " Exposure: " & Exposure

            'Set attachment
            'aEmail.ATTACHMENTS.Add ActiveWorkbook.FullName

                'Set Recipient
                aEmail.To = strRecipients
                    'or send one off to 1 person use this static code
                    'aEmail.Recipients.Add "E-mail.address-here@ntlworld.com"
                    
                    'Send Mail
                    aEmail.Send

AppendToLogFile ("Email Sent")

End Sub
Sub PlaceOrdersAutomaticallyLay()

    Call PlaceOrdersAutomaticallyNEW(True, "Lay", False)

End Sub
Sub PlaceOrdersAutomaticallyLayNoIncrement()

    Call PlaceOrdersAutomatically(False, "Lay")

End Sub
Sub PlaceOrdersAutomaticallyBack()

    Call PlaceOrdersAutomaticallyNEW(True, "Back", False)

End Sub
Sub PlaceOrdersAutomaticallyBackNoIncrement()

    Call PlaceOrdersAutomatically(False, "Back")

End Sub
Sub PlaceOrdersAutomaticallyBoth()

    Call PlaceOrdersAutomaticallyNEW(True, "Back", False)
    Call PlaceOrdersAutomaticallyNEW(True, "Lay", False)

End Sub
Sub PlaceOrdersAutomaticallyBothNoIncrement()

    Call PlaceOrdersAutomatically(False, "Back")
    Call PlaceOrdersAutomatically(False, "Lay")

End Sub
Sub BlankSub()
    On Error GoTo ErrorHandler:
    
    
    On Error GoTo 0
    Exit Sub
    
    
ErrorHandler:
    HandleError "BlankSub"
End Sub

Sub ColourChanger()
    On Error GoTo ErrorHandler:
    
    Dim CurrentDay, NextDay As Integer
    Dim MyCell As Range
    Dim myColours(7) As Variant
    Dim ColourCycleCounter As Integer
    Dim ColourIndex As Integer
    Dim ColourCycleOffset As Integer
    
ColourCycleCounter = 0
ColourIndex = 0
ColourCycleOffset = 0


'Define colours to use
myColours(0) = 255 'red
myColours(1) = 49407 'orange
myColours(2) = 65535 'yellow
myColours(3) = 5296274 'green
myColours(4) = 5287936 'dark green
myColours(5) = 15773696 'blue
myColours(6) = 16738047 'violet

CurrentDay = Day(Selection.Cells(1, 1).Offset(0, 5).Value - 16 / 24)
    
For Each MyCell In Selection

    NextDay = Day(MyCell.Offset(0, 5).Value - 16 / 24)
    If NextDay = CurrentDay Then
        'ColourTheCell
        MyCell.Interior.Color = myColours(ColourIndex)
    Else
        CurrentDay = NextDay
        'ChangeColours
            ColourCycleCounter = ColourCycleCounter + 1
            ColourIndex = (ColourCycleCounter + ColourCycleOffset) Mod 7
        'ColourTheCell
            MyCell.Interior.Color = myColours(ColourIndex)
    End If


Next MyCell

    
    On Error GoTo 0
    Exit Sub
    
ErrorHandler:
    HandleError "ColourChanger"
End Sub

Sub LoadIncrementMatrix()
    On Error GoTo ErrorHandler:
    
Dim FirstCellRow As Long
Dim FirstCellCol As Long
    
FirstCellRow = GetNamedRngRow("TopOfArray", "Increments")
FirstCellCol = GetNamedRngColumn("TopOfArray", "Increments")




    On Error GoTo 0
    Exit Sub
        
ErrorHandler:
    HandleError "LoadIncrementMatrix"
End Sub
Sub PlaceOrdersAutomaticallyNEW(WithIncrement As Boolean, BetType As String, Contrary As Boolean)
'sub takes either Back or Lay (default = Back) and places bets on relevant type

On Error GoTo ErrorHandler:

Dim MyCell As Range
Dim CurrentOfferPrice, CurrentFairPrice, CurrentBetAmount As Double 'OfferPrice and BetAmount were generic for both, have now created Back and Lay versions below
Dim CurrentBackOfferPrice, CurrentBackBetAmount As Double
Dim CurrentLayOfferPrice, CurrentLayBetAmount As Double
Dim CutOffSoccerOnSheet, MyMaxWin, MyMinBet, MyMarginOnSheet As Double
Dim MyBackLimit, MyLayLimit, MyUpperLimit, MyMaxOdds, MyMinOdds, MyLowOdds As Double
Dim BackString, LayString As String
Dim BackOddsLimitLowAllOthers, BackOddsLimitHighAllOthers, LayOddsLimitLowAllOthers, LayOddsLimitHighAllOthers As Double
Dim BackOddsLimitLowMatchOdds, BackOddsLimitHighMatchOdds, LayOddsLimitLowMatchOdds, LayOddsLimitHighMatchOdds As Double

Dim BackOddsLimitFloorAllOthers, LayOddsLimitFloorAllOthers, BackOddsLimitFloorMatchOdds, LayOddsLimitFloorMatchOdds As Double
Dim BackOddsLimitCeilingAllOthers, LayOddsLimitCeilingAllOthers, BackOddsLimitCeilingMatchOdds, LayOddsLimitCeilingMatchOdds As Double

Dim MyBackMarketMakerLimit, MyLayMarketMakerLimit As Double 'add a premium if I'm the market maker


If IsMissing(BetType) Then BetType = "Back"
If BetType = "" Then BetType = "Back"

'The definition of Back or Lay in this procedure simply defines which COLUMN is selected to act upon
'Once the cell with the correct data is selected, we just need to call the plain PlaceOrders sub

CutOffSoccerOnSheet = Cells(GetNamedRngRow("CutOffSoccer"), GetNamedRngColumn("CutOffSoccer")).Value
    MyMarginOnSheet = Cells(GetNamedRngRow("MyMargin"), GetNamedRngColumn("MyMargin")).Value
        MyMaxWin = Sheets("Example").Cells(GetNamedRngRow("MaxWin", "Example"), GetNamedRngColumn("MaxWin", "Example")).Value
            MyMinBet = Sheets("Example").Cells(GetNamedRngRow("MinBet", "Example"), GetNamedRngColumn("MinBet", "Example")).Value
                MyBackLimit = Sheets("Example").Cells(GetNamedRngRow("BackLimit", "Example"), GetNamedRngColumn("BackLimit", "Example")).Value
                    MyLayLimit = Sheets("Example").Cells(GetNamedRngRow("LayLimit", "Example"), GetNamedRngColumn("LayLimit", "Example")).Value
                        MyUpperLimit = Sheets("Example").Cells(GetNamedRngRow("UpperLimit", "Example"), GetNamedRngColumn("UpperLimit", "Example")).Value
                            MyMaxOdds = Sheets("Example").Cells(GetNamedRngRow("MaxOdds", "Example"), GetNamedRngColumn("MaxOdds", "Example")).Value
                                MyMinOdds = Sheets("Example").Cells(GetNamedRngRow("MinOdds", "Example"), GetNamedRngColumn("MinOdds", "Example")).Value
                                    MyLowOdds = Sheets("Example").Cells(GetNamedRngRow("LowOdds", "Example"), GetNamedRngColumn("LowOdds", "Example")).Value

        MyBackMarketMakerLimit = MyBackLimit * 1.3 'add a premium if I'm the market maker
            MyLayMarketMakerLimit = MyLayLimit * 1.3 'add a premium if I'm the market maker

BackOddsLimitFloorAllOthers = Sheets("Example").Cells(GetNamedRngRow("BackOddsLimitFloorAllOthers", "Example"), GetNamedRngColumn("BackOddsLimitFloorAllOthers", "Example")).Value
    BackOddsLimitLowAllOthers = Sheets("Example").Cells(GetNamedRngRow("BackOddsLimitLowAllOthers", "Example"), GetNamedRngColumn("BackOddsLimitLowAllOthers", "Example")).Value
        BackOddsLimitHighAllOthers = Sheets("Example").Cells(GetNamedRngRow("BackOddsLimitHighAllOthers", "Example"), GetNamedRngColumn("BackOddsLimitHighAllOthers", "Example")).Value
            BackOddsLimitCeilingAllOthers = Sheets("Example").Cells(GetNamedRngRow("BackOddsLimitCeilingAllOthers", "Example"), GetNamedRngColumn("BackOddsLimitCeilingAllOthers", "Example")).Value
                BackOddsLimitFloorMatchOdds = Sheets("Example").Cells(GetNamedRngRow("BackOddsLimitFloorMatchOdds", "Example"), GetNamedRngColumn("BackOddsLimitFloorMatchOdds", "Example")).Value
                    BackOddsLimitLowMatchOdds = Sheets("Example").Cells(GetNamedRngRow("BackOddsLimitLowMatchOdds", "Example"), GetNamedRngColumn("BackOddsLimitLowMatchOdds", "Example")).Value
                        BackOddsLimitHighMatchOdds = Sheets("Example").Cells(GetNamedRngRow("BackOddsLimitHighMatchOdds", "Example"), GetNamedRngColumn("BackOddsLimitHighMatchOdds", "Example")).Value
                            BackOddsLimitCeilingMatchOdds = Sheets("Example").Cells(GetNamedRngRow("BackOddsLimitCeilingMatchOdds", "Example"), GetNamedRngColumn("BackOddsLimitCeilingMatchOdds", "Example")).Value
        
LayOddsLimitFloorAllOthers = Sheets("Example").Cells(GetNamedRngRow("LayOddsLimitFloorAllOthers", "Example"), GetNamedRngColumn("LayOddsLimitFloorAllOthers", "Example")).Value
    LayOddsLimitLowAllOthers = Sheets("Example").Cells(GetNamedRngRow("LayOddsLimitLowAllOthers", "Example"), GetNamedRngColumn("LayOddsLimitLowAllOthers", "Example")).Value
        LayOddsLimitHighAllOthers = Sheets("Example").Cells(GetNamedRngRow("LayOddsLimitHighAllOthers", "Example"), GetNamedRngColumn("LayOddsLimitHighAllOthers", "Example")).Value
            LayOddsLimitCeilingAllOthers = Sheets("Example").Cells(GetNamedRngRow("LayOddsLimitCeilingAllOthers", "Example"), GetNamedRngColumn("LayOddsLimitCeilingAllOthers", "Example")).Value
                LayOddsLimitFloorMatchOdds = Sheets("Example").Cells(GetNamedRngRow("LayOddsLimitFloorMatchOdds", "Example"), GetNamedRngColumn("LayOddsLimitFloorMatchOdds", "Example")).Value
                    LayOddsLimitLowMatchOdds = Sheets("Example").Cells(GetNamedRngRow("LayOddsLimitLowMatchOdds", "Example"), GetNamedRngColumn("LayOddsLimitLowMatchOdds", "Example")).Value
                        LayOddsLimitHighMatchOdds = Sheets("Example").Cells(GetNamedRngRow("LayOddsLimitHighMatchOdds", "Example"), GetNamedRngColumn("LayOddsLimitHighMatchOdds", "Example")).Value
                            LayOddsLimitCeilingMatchOdds = Sheets("Example").Cells(GetNamedRngRow("LayOddsLimitCeilingMatchOdds", "Example"), GetNamedRngColumn("LayOddsLimitCeilingMatchOdds", "Example")).Value


If BetType = "Back" Then
        If Cells(3, 1) = "Soccer" Then
            Range(Cells(8, 7), Cells(SoccerLength, 7)).Select 'select the current LAY prices 47 for soccer, 38 for footy
  Else
                Range(Cells(8, 7), Cells(AFLLength, 7)).Select 'select the current LAY prices 47 for soccer, 38 for footy
        End If
ElseIf BetType = "Lay" Then
        If Cells(3, 1) = "Soccer" Then
    Range(Cells(8, 6), Cells(SoccerLength, 6)).Select 'select the current BACK prices
        Else
                    Range(Cells(8, 6), Cells(AFLLength, 6)).Select 'select the current BACK prices
        End If
   
End If


MyMaxWin = WorksheetFunction.Min(ScaleMultiply(MyMaxWin) / 1.6, 1.3 * MyMaxWin)

'thisMax = Cells(3, 24).Value
    
'BackString = "=MAX(ROUND(" & MyMaxWin & "/(RC[-1]-1),0)," & MyMinBet & ")"
'LayString = "=MAX(ROUND(" & MyMaxWin & "/(RC[-1]-1),0)," & MyMinBet & ")"
    
    For Each MyCell In Selection
    
        CurrentFairPrice = Cells(MyCell.Row, 15).Value
        
        If MyCell.Value = "" Then
            'find the next BEST (highest) price that is > MyMargin*CurrentFairPrice
            'NEED a function to ROUND doubles to the nearest Betfair price
            
            MyBackLimit = MyBackMarketMakerLimit 'add a premium if I'm the market maker
            MyLayLimit = MyLayMarketMakerLimit 'add a premium if I'm the market maker
            
                    If BetType = "Back" Then
                        CurrentBackOfferPrice = GetBetfairIncrement((1 + MyBackLimit) * CurrentFairPrice, "Up")
                        'Put an If statement here to get the correct LayOfferPrice depending on whether the price is empty or not
                            If Cells(MyCell.Row, 6).Value = "" Then
                                CurrentLayOfferPrice = GetBetfairIncrement(CurrentFairPrice * (1 - MyLayLimit), "Down")
                            Else
                                CurrentLayOfferPrice = GetBetfairIncrement(Cells(MyCell.Row, 6).Value, "Up")
                            End If
                            
                    ElseIf BetType = "Lay" Then
                        CurrentLayOfferPrice = GetBetfairIncrement(CurrentFairPrice * (1 - MyLayLimit), "Down")
                        'Put an If statement here to get the correct BackOfferPrice depending on whether the price is empty or not
                            If Cells(MyCell.Row, 7).Value = "" Then
                                CurrentBackOfferPrice = GetBetfairIncrement((1 + MyBackLimit) * CurrentFairPrice, "Up")
                            Else
                                CurrentBackOfferPrice = GetBetfairIncrement(Cells(MyCell.Row, 7).Value, "Down")
                            End If
                            
                    End If
        Else
            'CurrentOfferPrice = step DOWN from lay price or UP from Back price
                    If BetType = "Back" Then
                        CurrentBackOfferPrice = GetBetfairIncrement(MyCell.Value, "Down")
                        'Put an If statement here to get the correct LayOfferPrice depending on whether the price is empty or not
                            If Cells(MyCell.Row, 6).Value = "" Then
                                CurrentLayOfferPrice = GetBetfairIncrement(CurrentFairPrice * (1 - MyLayLimit), "Down")
                            Else
                                CurrentLayOfferPrice = GetBetfairIncrement(Cells(MyCell.Row, 6).Value, "Up")
                            End If
                            
                    ElseIf BetType = "Lay" Then
                        CurrentLayOfferPrice = GetBetfairIncrement(MyCell.Value, "Up")
                        'Put an If statement here to get the correct BackOfferPrice depending on whether the price is empty or not
                            If Cells(MyCell.Row, 7).Value = "" Then
                                CurrentBackOfferPrice = GetBetfairIncrement((1 + MyBackLimit) * CurrentFairPrice, "Up")
                            Else
                                CurrentBackOfferPrice = GetBetfairIncrement(Cells(MyCell.Row, 7).Value, "Down")
                            End If
                            
                    End If
        End If
        
        
'We now have the CurrentOfferPrices for Back and Lay

CurrentBackBetAmount = WorksheetFunction.Max(Round((1 * MyMaxWin) / (CurrentBackOfferPrice - 1), 0), MyMinBet)
CurrentLayBetAmount = WorksheetFunction.Max(Round((1 * MyMaxWin) / (CurrentLayOfferPrice - 1), 0), MyMinBet)

'Added selection ID cehck for (blank) on 11th April 2016 to stop filling sheet with data for invalid selections

        If Cells(MyCell.Row, 2).Value = "Match Odds" Then

                        If (BetType = "Back") And _
                                        (((CurrentBackOfferPrice > BackOddsLimitHighMatchOdds) And (CurrentBackOfferPrice < BackOddsLimitCeilingMatchOdds)) Or ((CurrentBackOfferPrice > BackOddsLimitFloorMatchOdds) And (CurrentBackOfferPrice < BackOddsLimitLowMatchOdds))) And _
                                            (CurrentBackOfferPrice < MyMaxOdds) And _
                                                ((CurrentBackOfferPrice / CurrentFairPrice) - 1 > MyBackLimit) And _
                                                    ((CurrentBackOfferPrice / CurrentFairPrice) - 1 < MyUpperLimit) And _
                                                        (Cells(MyCell.Row, 3).Value <> "") Then 'checks selection id has something
                                              
                                                    'GOING TO TRY AND DO THIS FOR LOW MATCH ODDS SINCE I'M LOSING ALMOST 100% OF THE TIME WHEN BACKING MATCH_ODDS BELOW $10
                                                    If (CurrentBackOfferPrice > BackOddsLimitFloorMatchOdds) And (CurrentBackOfferPrice < BackOddsLimitLowMatchOdds) And (Cells(3, 1) = "Soccer") Then Contrary = True
                                                    'If CONTRARY = true then do the opposite - need 3 more lines here
                                                    If Contrary Then
                                                        Cells(MyCell.Row, 10).Value = CurrentLayOfferPrice
                                                        Cells(MyCell.Row, 11).FormulaR1C1 = CurrentLayBetAmount 'BackString
                                                        Cells(MyCell.Row, 11).Select
                                                    Else 'Contrary is NOT true so be normal
                                                        Cells(MyCell.Row, 8).Value = CurrentBackOfferPrice
                                                        Cells(MyCell.Row, 9).FormulaR1C1 = CurrentBackBetAmount 'BackString
                                                        Cells(MyCell.Row, 9).Select
                                                    End If
                                                    
                                                    Call PlaceOrders
                                    
                                ElseIf (BetType = "Lay") And _
                                            (((CurrentLayOfferPrice > LayOddsLimitHighMatchOdds) And (CurrentLayOfferPrice < LayOddsLimitCeilingMatchOdds)) Or ((CurrentLayOfferPrice > LayOddsLimitFloorMatchOdds) And (CurrentLayOfferPrice < LayOddsLimitLowMatchOdds))) And _
                                                (CurrentLayOfferPrice < MyMaxOdds) And _
                                                    ((CurrentLayOfferPrice / CurrentFairPrice) < (1 - MyLayLimit)) And _
                                                        ((CurrentLayOfferPrice / CurrentFairPrice) > (1 - MyUpperLimit)) And _
                                                            (Cells(MyCell.Row, 3).Value <> "") Then 'checks selection id has something
                                              
                                                    'GOING TO TRY AND DO THIS FOR LOW MATCH ODDS SINCE I'M LOSING ALMOST 100% OF THE TIME WHEN BACKING MATCH_ODDS BELOW $10
                                                    If (CurrentLayOfferPrice > LayOddsLimitHighMatchOdds) And (CurrentLayOfferPrice < LayOddsLimitCeilingMatchOdds) And (Cells(3, 1) = "Soccer") Then Contrary = True
                                                    'If CONTRARY = true then do the opposite - need 3 more lines here
                                                    If Contrary Then
                                                        Cells(MyCell.Row, 8).Value = CurrentBackOfferPrice
                                                        Cells(MyCell.Row, 9).FormulaR1C1 = CurrentBackBetAmount 'LayString
                                                        Cells(MyCell.Row, 9).Select
                                                    Else
                                                        Cells(MyCell.Row, 10).Value = CurrentLayOfferPrice
                                                        Cells(MyCell.Row, 11).FormulaR1C1 = CurrentLayBetAmount 'LayString
                                                        Cells(MyCell.Row, 11).Select
                                                    End If
                                                    
                                                    Call PlaceOrders
                                    
                                End If



    Else 'else it is AllOthers not Match_Odds and we proceed with this

        If ((BetType = "Back") And _
                (((CurrentBackOfferPrice > BackOddsLimitHighAllOthers) And (CurrentBackOfferPrice < BackOddsLimitCeilingAllOthers)) Or ((CurrentBackOfferPrice > BackOddsLimitFloorAllOthers) And (CurrentBackOfferPrice < BackOddsLimitLowAllOthers))) And _
                    (CurrentBackOfferPrice < MyMaxOdds) And _
                        ((CurrentBackOfferPrice / CurrentFairPrice) - 1 > MyBackLimit) And _
                            ((CurrentBackOfferPrice / CurrentFairPrice) - 1 < MyUpperLimit) And _
                                (Cells(MyCell.Row, 3).Value <> "")) Then 'checks selection id has something
                      
                            'CurrentBackBetAmount = WorksheetFunction.Max(Round(MyMaxWin / (CurrentBackOfferPrice - 1), 0), MyMinBet)
                            
                            If Contrary Then
                                                        Cells(MyCell.Row, 10).Value = CurrentLayOfferPrice
                                                        Cells(MyCell.Row, 11).FormulaR1C1 = CurrentLayBetAmount 'BackString
                                                        Cells(MyCell.Row, 11).Select
                                            Else 'Contrary is NOT true so be normal
                                                        Cells(MyCell.Row, 8).Value = CurrentBackOfferPrice
                                                        Cells(MyCell.Row, 9).FormulaR1C1 = CurrentBackBetAmount 'BackString
                                                        Cells(MyCell.Row, 9).Select
                                                    End If
                                                    
                            Call PlaceOrders
            
        ElseIf ((BetType = "Lay") And _
                    (((CurrentLayOfferPrice > LayOddsLimitHighAllOthers) And (CurrentLayOfferPrice < LayOddsLimitCeilingAllOthers)) Or ((CurrentLayOfferPrice > LayOddsLimitFloorAllOthers) And (CurrentLayOfferPrice < LayOddsLimitLowAllOthers))) And _
                        (CurrentLayOfferPrice < MyMaxOdds) And _
                            ((CurrentLayOfferPrice / CurrentFairPrice) < (1 - MyLayLimit)) And _
                                ((CurrentLayOfferPrice / CurrentFairPrice) > (1 - MyUpperLimit)) And _
                                    (Cells(MyCell.Row, 3).Value <> "")) Then 'checks selection id has something
                      
                            'CurrentLayBetAmount = WorksheetFunction.Max(Round(MyMaxWin / (CurrentLayOfferPrice - 1), 0), MyMinBet)
                            
                            'If CONTRARY = true then do the opposite - need 3 more lines here
                            'GOING TO TRY AND DO THIS FOR LOW MATCH ODDS SINCE I'M LOSING ALMOST 100% OF THE TIME WHEN BACKING MATCH_ODDS BELOW $10
                                          If (CurrentLayOfferPrice > LayOddsLimitFloorAllOthers) And (CurrentLayOfferPrice < LayOddsLimitLowAllOthers) And (Cells(3, 1) = "Soccer") Then Contrary = True
                                                    'If CONTRARY = true then do the opposite - need 3 more lines here
                                                    
                                        
                                        If Contrary Then
                                                        Cells(MyCell.Row, 8).Value = CurrentBackOfferPrice
                                                        Cells(MyCell.Row, 9).FormulaR1C1 = CurrentBackBetAmount 'LayString
                                                        Cells(MyCell.Row, 9).Select
                                            Else
                                                        Cells(MyCell.Row, 10).Value = CurrentLayOfferPrice
                                                        Cells(MyCell.Row, 11).FormulaR1C1 = CurrentLayBetAmount 'LayString
                                                        Cells(MyCell.Row, 11).Select
                                                    End If
                            Call PlaceOrders
            
            
     '#*$*#* Big RISK but trying it
              '         ElseIf ((BetType = "Back") And _
              '             (CurrentBackOfferPrice > BackOddsLimitLowAllOthers) And _
              '                 (CurrentBackOfferPrice < MyMaxOdds) And _
              '                     ((CurrentBackOfferPrice / CurrentFairPrice) - 1 > MyBackLimit) And _
              '                         ((CurrentBackOfferPrice / CurrentFairPrice) - 1 < MyUpperLimit) And _
              '                             (Cells(MyCell.Row, 3).Value <> "")) Then   'checks selection id has something
              '
              '                                        'CurrentBackBetAmount = WorksheetFunction.Max(Round(MyMaxWin / (CurrentBackOfferPrice - 1), 0), MyMinBet)
              '                                        Contrary = Not (Contrary)
              '                                        'this is the RISK - doing a CONTRARY bet based on findings from November 2017
              '                                        'if previously I would have BACKED, and seemed to lose nearly 100% - I'm going to do the contrary and LAY
              '                                                        If Contrary Then
              '                                                                    Cells(MyCell.Row, 10).Value = CurrentLayOfferPrice
              '                                                                    Cells(MyCell.Row, 11).FormulaR1C1 = CurrentLayBetAmount 'BackString
              '                                                                    Cells(MyCell.Row, 11).Select
              '                                                        Else 'Contrary is NOT true so be normal
              '                                                                    Cells(MyCell.Row, 8).Value = CurrentBackOfferPrice
              '                                                                    Cells(MyCell.Row, 9).FormulaR1C1 = CurrentBackBetAmount 'BackString
              '                                                                    Cells(MyCell.Row, 9).Select
              '                                                        End If
              '                                          Contrary = Not (Contrary) 'change it BACK
              '
              '                                         Call PlaceOrders
                                      
            
            
        End If
        
    End If 'Match Odds vs All Others

    Next 'MyCell in Selection

Cells(1, 1).Select

On Error GoTo 0
Exit Sub

ErrorHandler:
    HandleError "PlaceOrdersAutomatically"
    Resume Next

End Sub

Function GetBetfairIncrement(CurrentPrice As Double, Direction As String) As Double

    On Error GoTo ErrorHandler:
    
    Dim LookupRangeUp As Range
    Dim LookupRangeDown As Range
    
        Set LookupRangeUp = Range(Cells(1, 52), Cells(1000, 53))
        Set LookupRangeDown = Range(Cells(1, 50), Cells(1000, 51))
    
    
    If CurrentPrice < 1.02 Then
        GetBetfairIncrement = 1
        On Error GoTo 0
        Exit Function
    End If
    
    If Direction = "Up" Then
        GetBetfairIncrement = Application.WorksheetFunction.VLookup(CurrentPrice, LookupRangeUp, 2)
    Else
        GetBetfairIncrement = Application.WorksheetFunction.VLookup(CurrentPrice, LookupRangeDown, 2)
    End If
    
    On Error GoTo 0
    Exit Function
    
    
ErrorHandler:
    HandleError "GetBetfairIncrement"
    
End Function

Function GetBetIDFromListCurrentOrders(ByRef ListCurrentOrdersResult As Object, ByVal ThisSelectionId As String, ByVal ThisSide As String) As String

Dim BetsOnThisMarket As Integer
Dim BetCounter As Integer: BetCounter = 0
Dim myCount As Integer
Dim PutItInColumn As Integer: PutItInColumn = 0
Dim TotalReturn As Double: TotalReturn = 0
Dim TotalSize As Double: TotalSize = 0
Dim TotalAveragePriceMatched As Double: TotalAveragePriceMatched = 0
Dim Spent As Double: Spent = 0


BetsOnThisMarket = ListCurrentOrdersResult.Item("currentOrders").Count

'need to loop through all the CurrentOrders for this market BUT in the meantime, the quick and dirty solution is that
'I assume there's only 1 BET on each market so will just output the values

For myCount = 1 To BetsOnThisMarket

With ListCurrentOrdersResult.Item("currentOrders").Item(myCount)

If .Item("selectionId") = ThisSelectionId And _
        .Item("side") = ThisSide And _
            .Item("sizeRemaining") > 0 Then
    GetBetIDFromListCurrentOrders = .Item("betId")
    Exit Function
End If

End With

Next
'End If


End Function

Sub SortSheets()

    Dim myArray() As Variant
    Dim mySize, counter As Integer
    Dim WS As Worksheet
    Dim R As Range
    Dim N As Long
    Dim CountSortSheets As Integer
    Dim ProcessStartTime As Date
    Dim TotalMatched As Double
    Dim Ndx As Long
    Dim ThisCompetitionRegion As String
    Dim SortNumberToProcess As Integer
    'Dim eventID As String

'setup some initial variables
mySize = Worksheets.Count
ProcessStartTime = Now()
ReDim Preserve myArray(1 To mySize, 1 To 3)
CountSortSheets = 0


        For Each WS In Worksheets
        
        If (Left(WS.Name, 1) = "2") And (WS.Visible = True) Then
        
            CountSortSheets = CountSortSheets + 1
            
            ThisCompetitionRegion = WS.Cells(4, 1).Value
            
            TotalMatched = GetTotalMatchedForEventID(WS.Name, ThisCompetitionRegion)
            'NewEventTime = GetOpenDateFromEventID(ThisEventID, ThisEndPoint)
            
            myArray(CountSortSheets, 1) = WS.Name
            myArray(CountSortSheets, 2) = TotalMatched
            WS.Cells(5, 1).Value = TotalMatched
            myArray(CountSortSheets, 3) = WS.Cells(6, 3) - ProcessStartTime
            
            Application.StatusBar = "Getting Total = " & TotalMatched & " for game " & WS.Name
            
        End If
        
        Next WS


    Set WS = ThisWorkbook.Worksheets.Add
    
    ' put the array values on the worksheet
    Set R = WS.Range("A1").Resize(3, UBound(myArray) - LBound(myArray) + 1)
    R = Application.Transpose(myArray)
    
            ' sort the range
                    WS.Sort.SortFields.Clear
                    WS.Sort.SortFields.Add key:=Cells(2, 1), _
                        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                            'Descending = Price from highest to lowest
                            'Ascending = Price from lowest to highest Order:=xlAscending
                            'Have tried both, but feel low to high may be best since it means betting on games with less money and MORE likely to have wider opportunities
                    WS.Sort.SortFields.Add key:=Cells(3, 1), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'so that those with $0 dispaly in chrono order
                    With WS.Sort
                        .SetRange R
        '                .Header = xlGuess
        '                .MatchCase = False
                        .Orientation = xlLeftToRight
                        .SortMethod = xlPinYin
                        .Apply
                    End With
    
    
                                ' load the worksheet values back into the array
                                For N = 1 To R.Columns.Count
                                    myArray(N, 1) = R(1, N)
                                    myArray(N, 2) = R(2, N)
                                    myArray(N, 3) = R(3, N)
                                    
                                Next N

' delete the temporary sheet
    Application.DisplayAlerts = False
        WS.Delete
    Application.DisplayAlerts = True

Application.StatusBar = "Sorting"

                'now, do the sorting
                For Ndx = CountSortSheets To 1 Step -1
                    Worksheets("" & myArray(Ndx, 1) & "").Move before:=Worksheets(3)
                Next Ndx

'Grab some values
SortNumberToProcess = Sheets("Example").Cells(GetNamedRngRow("NumberToProcess", "Example"), GetNamedRngColumn("NumberToProcess", "Example")).Value
CountSortSheets = 0

For Each WS In Worksheets
        
        If (Left(WS.Name, 1) = "2") And (WS.Visible = True) Then 'if it's a Game AND visible
        
            CountSortSheets = CountSortSheets + 1
            If CountSortSheets > SortNumberToProcess Then WS.Visible = xlSheetHidden 'Hide those greater than NumberToProcess

        End If

Next WS


Application.OnTime Now() + TimeSerial(0, 0, 3), "ClearStatusBar"

   On Error GoTo 0
    Exit Sub
    
    
ErrorHandler:
    HandleError "SortSheets"
End Sub
Sub AddNames()
Dim RefersToString As String
Dim WS As Worksheet

For Each WS In Worksheets
        
        If (Left(WS.Name, 1) = "2") And (WS.Visible = True) And WS.Cells(3, 1).Value = "AFL" Then
        
RefersToString = "='" & WS.Name & "'!R1C16"
Sheets(WS.Name).Select
    Range("P1").Select
    ActiveWorkbook.Worksheets(WS.Name).Names.Add Name:="kurtonsheet", _
        RefersToR1C1:=RefersToString
    ActiveWorkbook.Worksheets(WS.Name).Names("kurtonsheet").Comment = ""
End If
Next WS
End Sub

Sub CleanseAllEvents()

Dim WSToCleanse As Worksheet

For Each WSToCleanse In Worksheets

If (Left(WSToCleanse.Name, 1) = "2") Then
   
    WSToCleanse.Activate
    ListCurrentOrdersNew ("Cleanse")
    
End If

Next

End Sub
Sub TopOfferOnlyAllEvents()

Dim WSToCleanse As Worksheet

For Each WSToCleanse In Worksheets

If (Left(WSToCleanse.Name, 1) = "2") Then
   
    WSToCleanse.Activate
    ListCurrentOrdersNew ("TopOfferOnly")
    
End If

Next

End Sub

Function ModEventIDRemainder(eventid As Double, modulus As Double) As Double

ModEventIDRemainder = Round(Round(modulus, 0) * ((eventid / modulus) - WorksheetFunction.RoundDown(eventid / modulus, 0)), 0)

End Function
Sub AddAccountFundsToTrackingSpreadsheet(ByVal MyAccountFunds As Double, ByVal MyExposure As Double)

Dim CurrentActiveSheet As Worksheet
Dim FundsSheet As Worksheet
Dim myOutputRow As Double
Dim mySportsbet_Balance As Double
Dim myAdjustment As Double
Dim myCurrentBackLimit As Double


Set CurrentActiveSheet = Application.ActiveSheet
Set FundsSheet = Worksheets("Tracking")

mySportsbet_Balance = Sheets("Example").Cells(GetNamedRngRow("Sportsbet_Balance", "Example"), GetNamedRngColumn("Sportsbet_Balance", "Example")).Value
myAdjustment = Sheets("Example").Cells(GetNamedRngRow("Adjustment", "Example"), GetNamedRngColumn("Adjustment", "Example")).Value
myCurrentBackLimit = Sheets("Example").Cells(GetNamedRngRow("BackLimit", "Example"), GetNamedRngColumn("BackLimit", "Example")).Value


FundsSheet.Activate

Cells(1, 1).Select

myOutputRow = ActiveSheet.Cells(1, 1).CurrentRegion.rows.Count + 1 'it wasn't found so put the LOT into a new Row

    Cells(myOutputRow, 1) = Now()
        Cells(myOutputRow, 2) = MyAccountFunds
            Cells(myOutputRow, 3) = MyExposure
                Cells(myOutputRow, 4) = MyAccountFunds + mySportsbet_Balance + myAdjustment - MyExposure
                    Cells(myOutputRow, 5) = myCurrentBackLimit
                    
CurrentActiveSheet.Activate


Call CreateKeepAliveFile("hello", MyAccountFunds + mySportsbet_Balance + myAdjustment - MyExposure)

End Sub
Sub AddTotalMatchedToTrackingSpreadsheet(ByVal eventid As Double, ByVal TotalMatched As Double)

Dim CurrentActiveSheet As Worksheet
Dim FundsSheet As Worksheet
Dim myOutputRow As Double

Set CurrentActiveSheet = Application.ActiveSheet
Set FundsSheet = Worksheets("Total Matched")

FundsSheet.Activate

Cells(1, 1).Select

myOutputRow = ActiveSheet.Cells(1, 1).CurrentRegion.rows.Count + 1 'it wasn't found so put the LOT into a new Row

    Cells(myOutputRow, 1) = Now()
        Cells(myOutputRow, 2) = eventid
            Cells(myOutputRow, 3) = TotalMatched
                'Cells(myOutputRow, 4) = MyAccountFunds - MyExposure

CurrentActiveSheet.Activate

End Sub
Sub AddSolverProgressToLogSpreadsheet(ByVal HomeMean As Double, ByVal AwayMean As Double, StdDevHA As Double, SolverError As Double, Mult As Double)

Dim CurrentActiveSheet As Worksheet
Dim FundsSheet As Worksheet
Dim myOutputRow As Double

Set CurrentActiveSheet = Application.ActiveSheet
Set FundsSheet = Worksheets("Log")

FundsSheet.Activate

Cells(1, 1).Select

myOutputRow = ActiveSheet.Cells(1, 1).CurrentRegion.rows.Count + 1 'it wasn't found so put the LOT into a new Row

    Cells(myOutputRow, 1) = Now()
        Cells(myOutputRow, 2) = HomeMean
            Cells(myOutputRow, 3) = AwayMean
                Cells(myOutputRow, 4) = StdDevHA
                    Cells(myOutputRow, 5) = SolverError
                        Cells(myOutputRow, 6) = Mult

CurrentActiveSheet.Activate

End Sub


Public Function GetUTCOffset() As Integer
    Dim lngRet As Long
    Dim udtTZI As TimeZoneInfo
    Dim CurrentDaylightBias As Long

    lngRet = GetTimeZoneInformation(udtTZI)
    GetUTCOffset = udtTZI.lngBias / 60 'returns an integer positive or negative with the current UTC offset
        CurrentDaylightBias = udtTZI.intDaylightBias
            GetUTCOffset = (-1 * GetUTCOffset) - IIf(IsDST, CurrentDaylightBias / 60, 0)
    
End Function

Public Function IsDST() As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsDST
    ' Returns TRUE if the current data is within DST, FALSE
    ' if DST is not in effect or if Windows cannot determine
    ' DST setting.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Application.Volatile True
        Dim DST As TIME_ZONE
        Dim TZI As TimeZoneInfo
        DST = GetTimeZoneInformation(TZI)
        IsDST = (DST = TIME_ZONE_DAYLIGHT)
End Function
Sub TestDST()
    Dim TZI As TimeZoneInfo
    Dim DST As TIME_ZONE
    
    DST = GetTimeZoneInformation(TZI)
    Select Case DST
        Case TIME_ZONE_ID_INVALID
            Debug.Print "Windows cannot determine DST."
        Case TIME_ZONE_STANDARD
            Debug.Print "Current date is in Standard Time"
        Case TIME_ZONE_DAYLIGHT
            Debug.Print "Current date is in Daylight Time"
        Case Else
            Debug.Print "**** ERROR: Unknown Result From GetTimeZoneInformation"
    End Select
End Sub
Public Function ScaleMultiply(ByVal MyInput As Double) As Double

'MyMaxWin = WorksheetFunction.Min(ScaleMultiply(MyMaxWin), 750)
'amplitude 0.317608165
'Period 0.246279333
'phase shift 1.058957118
'vertical shift  0.49225599
'y = A sin(Bx + C) + D

'amplitude is A
'period is 2p/B
'phase shift Is -c / B
'Vertical shift Is d

Dim myamplitude As Double
Dim myPeriod As Double
Dim myPhaseShift As Double
Dim myVerticalShift  As Double

        myamplitude = 0.317608165
        myPeriod = 0.246279333
        myPhaseShift = 1.058957118
        myVerticalShift = 0.49225599

ScaleMultiply = MyInput * 1 / (myamplitude * Sin(myPeriod * Hour(Now() + 3.5 / 24) + myPhaseShift) + myVerticalShift)



End Function

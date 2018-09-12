Attribute VB_Name = "TotalMatched"

Sub TotalMatched_Comp2Event()

Dim RowCounter As Integer: RowCounter = 1
Dim EventsInComp As Integer

myOutputCol = 1
myOutputRow = 276

For Each thiscomp In Selection

'GetEvents for the Competition
thisCompId = thiscomp
thisCompName = Cells(thiscomp.Row, 4).Value

    Request = MakeJsonRpcRequestString(ListEventsMethod, GetListEventsRequestString(99, thisCompId))
    Dim ListEventsResponse As String: ListEventsResponse = SendRequest(GetJsonRpcUrl(), GetAppKey(), GetSession(), Request)
    Dim ListEventsResult As Object: Set ListEventsResult = ParseJsonRpcResponseToCollection(ListEventsResponse)
    
    'Count the Events for the Competition
    EventsInComp = ListEventsResult.Count 'thiscomp.Row '10 + (thiscomp.Row - 7) * 2
    'RowCounter = RowCounter + EventsInComp
      
    Cells(myOutputRow + 1, myOutputCol - 1) = thisCompId
    Cells(myOutputRow + 1, myOutputCol - 0) = thisCompName
    
    OutputCompListEvents ListEventsResult, myOutputRow, myOutputCol + 1, "Vert"
    myOutputRow = myOutputRow + EventsInComp
    
Next

End Sub
Sub TotalMatched_EventID2TotalMatched()

Dim EndPoint As String
Dim MatchCatalogueString As String

'countryCode")
    countryCodeCol = GetHeaderColumn("competition:region")


For Each thisEvent In Selection

Application.StatusBar = "Processed : " & thisEvent.Row

'GetEvents for the Competition
If WorksheetFunction.IsNumber(thisEvent) Then

ThisEventID = thisEvent
EndPoint = Cells(thisEvent.Row, countryCodeCol).Value

If EndPoint = "US" Then
    MatchCatalogueString = "Moneyline"
Else
    MatchCatalogueString = "Match Odds"
End If

If (EndPoint = "AU" Or EndPoint = "NZ") Then EndPoint = "UK" 'was AUS
If (EndPoint = "" Or EndPoint = "UK") Then EndPoint = "UK"




'GetJsonRpcUrl(EndPoint)

    Request = MakeJsonRpcRequestString(ListMatchCatalogueMethod, GetListMatchCatalogueRequestString(ThisEventID))
    Dim ListMatchCatalogueResponse As String: ListMatchCatalogueResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
    'Sheet1.Cells(OutputRow + 7, OutputColumn).Value = ListSoccerMatchCatalogueResponse
    'AppendToLogFile ListSoccerMatchCatalogueResponse
        
    Dim ListMatchCatalogueResult As Object: Set ListMatchCatalogueResult = ParseJsonRpcResponseToCollection(ListMatchCatalogueResponse)
    'GetSelectionIDFromSoccerMatch
    
            
'Output the Events for the Competition
    myOutputRow = thisEvent.Row
    myOutputCol = 3 'thisEvent.Column + 2
    
    'output the plain JSON text
    'Cells(myOutputRow, myOutputCol).Value = ListSoccerMatchCatalogueResponse

    Dim ThisMarketID As String: ThisMarketID = GetMarketIdFromMatchCatalogue(MatchCatalogueString, ListMatchCatalogueResult)
    
    'Dim theseSelectionIds As Object: Set theseSelectionIds = GetSelectionIdFromMatchCatalogue("Match Odds", ListSoccerMatchCatalogueResult)
    
    Request = MakeJsonRpcRequestString(ListMarketBookMethod, GetListMarketBookRequestStringV(ThisMarketID))
    Dim ListMarketBookResponse As String: ListMarketBookResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
    'output the plain JSON text
    'Cells(myOutputRow, myOutputCol + 1).Value = ListMarketBookResponse 'AppendToLogFile ListSoccerMatchCatalogueResponse
    
    Dim ListMarketBookResult As Object: Set ListMarketBookResult = ParseJsonRpcResponseToCollection(ListMarketBookResponse)
    
    
    If ThisMarketID <> "" Then
    
        If Not ListMarketBookResult Is Nothing Then
    Dim TotalMatched As Double: TotalMatched = ListMarketBookResult.Item(1).Item("totalMatched")
        Else
    TotalMatched = 0
        End If
    'output
    'Cells(myOutputRow, myOutputCol).Value = thisMarketId
    
    
    Cells(myOutputRow, myOutputCol).Value = TotalMatched
    
    
    'Next
    End If
    'OutputCompListEvents ListEventsResult, myOutputRow + 3, myOutputCol
End If
Next
End Sub
Sub TotalMatched_AllEvents()

On Error Resume Next

Dim RowCounter As Integer: RowCounter = 1
Dim EventsInResult As Integer
Dim thisItem As Integer


myOutputCol = 1
myOutputRow = 6706


ProcessStartTime = Now()
NumberToProcess = Sheets("Example").Cells(GetNamedRngRow("NumberToProcess", "Example"), GetNamedRngColumn("NumberToProcess", "Example")).Value


    Request = MakeJsonRpcRequestString(ListMarketCatalogueMethod, GetListMarketCatalogueRequestStringMATCH_ODDS("""1", "6423", "61420"""))
    Dim ListMarketCatalogueResponse As String: ListMarketCatalogueResponse = SendRequest(GetJsonRpcUrl(), GetAppKey(), GetSession(), Request)
    Dim ListMarketCatalogueResult As Object: Set ListMarketCatalogueResult = ParseJsonRpcResponseToCollection(ListMarketCatalogueResponse)
    
    'Count the Events for the Competition
    EventsInResult = ListMarketCatalogueResult.Count 'thiscomp.Row '10 + (thiscomp.Row - 7) * 2
    
    
    'OutputCompListEvents ListMarketCatalogueResult, myOutputRow, myOutputCol + 1, "Vert"
    'myOutputRow = myOutputRow + EventsInComp
    
   For thisItem = 1 To EventsInResult
    Cells(myOutputRow, myOutputCol + 0).Value = ListMarketCatalogueResult.Item(thisItem).Item("marketId")
    Cells(myOutputRow, myOutputCol + 1).Value = ListMarketCatalogueResult.Item(thisItem).Item("marketName")
    Cells(myOutputRow, myOutputCol + 2).Value = ListMarketCatalogueResult.Item(thisItem).Item("totalMatched")
    Cells(myOutputRow, myOutputCol + 3).Value = ListMarketCatalogueResult.Item(thisItem).Item("event").Item("id")
    Cells(myOutputRow, myOutputCol + 4).Value = ListMarketCatalogueResult.Item(thisItem).Item("event").Item("name")
    Cells(myOutputRow, myOutputCol + 5).Value = ListMarketCatalogueResult.Item(thisItem).Item("event").Item("countryCode")
    Cells(myOutputRow, myOutputCol + 8).Value = ListMarketCatalogueResult.Item(thisItem).Item("event").Item("openDate")
    If Not IsEmpty(ListMarketCatalogueResult.Item(thisItem).Item("competition")) Then
        Cells(myOutputRow, myOutputCol + 7).Value = ListMarketCatalogueResult.Item(thisItem).Item("competition").Item("name")
        Cells(myOutputRow, myOutputCol + 6).Value = ListMarketCatalogueResult.Item(thisItem).Item("competition").Item("id")
    End If
    
    myOutputRow = myOutputRow + 1
    RowCounter = RowCounter + 1
    
    ElapsedTime = Now() - ProcessStartTime
    AverageProcessTime = ElapsedTime / RowCounter
    
    TimeRemaining = AverageProcessTime * (EventsInResult - RowCounter)

Application.StatusBar = "Processed : " & RowCounter & " Remaining : " & (EventsInResult - RowCounter) & " and Time Remaining = " & Format(TimeRemaining, "hh:mm:ss")
 
    
    
   Next thisItem

On Error GoTo 0

End Sub

Function GetTotalMatchedForEventID(ByVal eventid As String, ByVal EndPoint As String) As Double

If (EndPoint = "AU" Or EndPoint = "NZ") Then EndPoint = "UK"
If (EndPoint = "" Or EndPoint = "UK") Then EndPoint = "UK"

    Request = MakeJsonRpcRequestString(ListMarketCatalogueMethod, GetListMarketCatalogueRequestStringUsingEventID4MATCH_ODDS(eventid))
    Dim ListMarketCatalogueResponse As String: ListMarketCatalogueResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
    Dim ListMarketCatalogueResult As Object: Set ListMarketCatalogueResult = ParseJsonRpcResponseToCollection(ListMarketCatalogueResponse)
    
    'Count the Events for the Competition
    If Not ListMarketCatalogueResult Is Nothing Then
        EventsInResult = ListMarketCatalogueResult.Count
    Else
        EventsInResult = 0
    End If
    
    If EventsInResult > 0 Then
 
        GetTotalMatchedForEventID = ListMarketCatalogueResult.Item(1).Item("totalMatched")
    Else
        GetTotalMatchedForEventID = 0
    End If
    

End Function
Function GetMarketNameForMarketID(ByVal marketid As String) As String

Dim EndPoint As String
EndPoint = "UK"

    Request = MakeJsonRpcRequestString(ListMarketCatalogueMethod, GetListMarketCatalogueRequestStringUsingMarketID(marketid))
    Dim ListMarketCatalogueResponse As String: ListMarketCatalogueResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
    Dim ListMarketCatalogueResult As Object: Set ListMarketCatalogueResult = ParseJsonRpcResponseToCollection(ListMarketCatalogueResponse)
    
    'Count the Events for the Competition
    If Not ListMarketCatalogueResult Is Nothing Then
        EventsInResult = ListMarketCatalogueResult.Count
    Else
        EventsInResult = 0
    End If
    
    If EventsInResult > 0 Then
 
        GetMarketNameForMarketID = ListMarketCatalogueResult.Item(1).Item("marketName")
    Else
        GetMarketNameForMarketID = 0
    End If
    

End Function
Sub Build_Future_List()

On Error Resume Next

Dim RowCounter As Integer: RowCounter = 0
Dim EventsInResult As Integer
Dim thisItem As Integer
Dim EndPoint As String
Dim eventidCol As Integer
Dim endpointloop As Integer

Dim CountryCodeReject As Boolean
Dim CurrentTimeOffset As Double
Dim OlympicEvent As Integer
Dim WomensEvent As Integer
Dim UEFAEvent As Integer
Dim GameTypes, GameTypesString As String

CurrentTimeOffset = GetUTCOffset


Dim MyLocation As Worksheet
Set MyLocation = Worksheets("Future List")

MyLocation.Activate

'marketId")
marketIdCol = GetHeaderColumn("marketId")
 'marketName")
 marketNameCol = GetHeaderColumn("marketName")
  'totalMatched")
  totalMatchedCol = GetHeaderColumn("total:matched")
   'event").Item("id")
   eventidCol = GetHeaderColumn("event:id")
    'event").Item("name")
    eventnameCol = GetHeaderColumn("event:name")
    'countryCode")
    countryCodeCol = GetHeaderColumn("competition:region")
    'openDate")
    openDateCol = GetHeaderColumn("event:time")
    'competition").Item("name")
    competitionnameCol = GetHeaderColumn("competition:name")
     'competition").Item("id")
     competitionidCol = GetHeaderColumn("competition:id")
     
     'rejectcolumn
     rejectCountryCol = GetHeaderColumn("reject:Country")
     
     DateSpreadCol = GetHeaderColumn("Date Spread")
        ClockCol = GetHeaderColumn("Clock")
            LocalDateTimeCol = GetHeaderColumn("LocalDate-Time")
    
  NextFewHoursCol = GetHeaderColumn("NextFewHours")
        OffsetDateTimeCol = GetHeaderColumn("OffsetDateTime")
            CreateDateTimeCol = GetHeaderColumn("CreateDateTime")
                CheckCol = GetHeaderColumn("Check")
                  GameTypes = Sheets("Example").Cells(GetNamedRngRow("GameTypes", "Example"), GetNamedRngColumn("GameTypes", "Example")).Value
                          
    
GameTypesString = "0"
If InStr(1, GameTypes, "AFL") <> 0 Then
    GameTypesString = GameTypesString & """ ,""61420" & ""
End If
        If InStr(1, GameTypes, "NFL") <> 0 Then
            GameTypesString = GameTypesString & """ ,""6423" & ""
        End If
                If InStr(1, GameTypes, "Soccer") <> 0 Then
                    GameTypesString = GameTypesString & """ ,""1" & ""
                End If

'NextFewHours    OffsetDateTime  CreateDateTime
    
'Have to do this TWICE, once for UK endpoint and once for AUS endpoint
    
    
For endpointloop = 1 To 1
    
If endpointloop = 1 Then EndPoint = "UK"
If endpointloop = 2 Then EndPoint = "AUS"
    
myOutputCol = 1
myOutputRow = 6

ProcessStartTime = Now()
NumberToProcess = 350 'Sheets("Example").Cells(GetNamedRngRow("NumberToProcess", "Example"), GetNamedRngColumn("NumberToProcess", "Example")).Value

'Get the data, sorted by most money match on MATCH_ODDS
    Request = MakeJsonRpcRequestString(ListMarketCatalogueMethod, GetListMarketCatalogueRequestStringMATCH_ODDS(GameTypesString))
    Dim ListMarketCatalogueResponse As String: ListMarketCatalogueResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
    Dim ListMarketCatalogueResult As Object: Set ListMarketCatalogueResult = ParseJsonRpcResponseToCollection(ListMarketCatalogueResponse)
    
    Call AppendToLogFile(ListMarketCatalogueResponse, "Build_Future_List", 999)

    
    'Count the Events for the Competition
    EventsInResult = ListMarketCatalogueResult.Count 'thiscomp.Row '10 + (thiscomp.Row - 7) * 2
    
        
For thisItem = 1 To EventsInResult
   
'changed BR to BE on 22-08-17
'changed RU to BE on 26-09-17
If (ListMarketCatalogueResult.Item(thisItem).Item("event").Item("countryCode") = "BE" Or _
        ListMarketCatalogueResult.Item(thisItem).Item("event").Item("countryCode") = "BE" Or _
            ListMarketCatalogueResult.Item(thisItem).Item("event").Item("countryCode") = "BE" Or _
                ListMarketCatalogueResult.Item(thisItem).Item("event").Item("countryCode") = "AE") Then 'Or _
                    'ListMarketCatalogueResult.Item(thisItem).Item("event").Item("countryCode") = "") Then
                    'BAD PEOPLE
                    CountryCodeReject = True
Else
    CountryCodeReject = False
End If

OlympicEvent = InStr(ListMarketCatalogueResult.Item(thisItem).Item("event").Item("name"), "(Olympic)")
WomensEvent = InStr(ListMarketCatalogueResult.Item(thisItem).Item("event").Item("name"), "(W)")
UEFAEvent = InStr(ListMarketCatalogueResult.Item(thisItem).Item("event").Item("name"), "UEFA")

If (UEFAEvent > 0) Or (OlympicEvent > 0) Or (WomensEvent > 0 And ListMarketCatalogueResult.Item(thisItem).Item("event").Item("countryCode") = "BR") Then
    CountryCodeReject = False
End If
   
   
myOutputRow = FindTheValue(MyLocation, eventidCol, ListMarketCatalogueResult.Item(thisItem).Item("event").Item("id"))
If (myOutputRow = 0) And (ListMarketCatalogueResult.Item(thisItem).Item("totalMatched") >= 500) Then 'And CountryCodeReject = False Then

    myOutputRow = ActiveSheet.Cells(1, 1).CurrentRegion.rows.Count + 1 'it wasn't found so put the LOT into a new Row

    
    'Cells(myOutputRow, myOutputCol + 0).Value = ListMarketCatalogueResult.Item(thisItem).Item("marketId")
    'Cells(myOutputRow, myOutputCol + 1).Value = ListMarketCatalogueResult.Item(thisItem).Item("marketName")
    Cells(myOutputRow, totalMatchedCol).Value = ListMarketCatalogueResult.Item(thisItem).Item("totalMatched")
    Cells(myOutputRow, eventidCol).Value = ListMarketCatalogueResult.Item(thisItem).Item("event").Item("id")
    Cells(myOutputRow, eventnameCol).Value = ListMarketCatalogueResult.Item(thisItem).Item("event").Item("name")
    Cells(myOutputRow, countryCodeCol).Value = ListMarketCatalogueResult.Item(thisItem).Item("event").Item("countryCode")
    Cells(myOutputRow, openDateCol).Value = ListMarketCatalogueResult.Item(thisItem).Item("event").Item("openDate")
    
    If Not IsEmpty(ListMarketCatalogueResult.Item(thisItem).Item("competition")) Then
        Cells(myOutputRow, competitionnameCol).Value = ListMarketCatalogueResult.Item(thisItem).Item("competition").Item("name")
        Cells(myOutputRow, competitionidCol).Value = ListMarketCatalogueResult.Item(thisItem).Item("competition").Item("id")
    End If
    
    Cells(myOutputRow, DateSpreadCol).FormulaR1C1 = "=DATE(LEFT(RC[-1],4),MID(RC[-1],6,2),MID(RC[-1],9,2))"
    Cells(myOutputRow, ClockCol).FormulaR1C1 = "=TIME(MID(RC[-2],12,2),MID(RC[-2],15,2),0)"
    Cells(myOutputRow, LocalDateTimeCol).FormulaR1C1 = "=RC[-2]+RC[-1]+" & CurrentTimeOffset & "/24"
    'NextFewHours    OffsetDateTime  CreateDateTime
    Cells(myOutputRow, NextFewHoursCol).FormulaR1C1 = "=COUNTIF(RC[-1]:R[170]C[-1]," < "&RC[1]&"")"
    Cells(myOutputRow, OffsetDateTimeCol).FormulaR1C1 = "=RC[-2]+4.5/24"
    Cells(myOutputRow, CreateDateTimeCol).FormulaR1C1 = "=RC[-3]-1.5"
    Cells(myOutputRow, CheckCol).FormulaR1C1 = "Game"
    
    If CountryCodeReject = True Then
        Cells(myOutputRow, rejectCountryCol).Value = "Contrary" 'just put a label there for now TEMPORARY  - MORE ACTION NEEDED
    End If
    
Else 'just put the totalMatched into the current column to update it - why not?
    Cells(myOutputRow, totalMatchedCol).Value = ListMarketCatalogueResult.Item(thisItem).Item("totalMatched")
    Cells(myOutputRow, openDateCol).Value = ListMarketCatalogueResult.Item(thisItem).Item("event").Item("openDate")
    
        If (ListMarketCatalogueResult.Item(thisItem).Item("event").Item("countryCode") <> "") Then 'add teh country code to the spreadsheet if it's missing
            Cells(myOutputRow, countryCodeCol).Value = ListMarketCatalogueResult.Item(thisItem).Item("event").Item("countryCode")
        End If
End If


If ListMarketCatalogueResult.Item(thisItem).Item("totalMatched") > 500 Then
    Call AddTotalMatchedToTrackingSpreadsheet(ListMarketCatalogueResult.Item(thisItem).Item("event").Item("id"), ListMarketCatalogueResult.Item(thisItem).Item("totalMatched"))
End If


    'myOutputRow = myOutputRow + 1
    RowCounter = RowCounter + 1
    
    ElapsedTime = Now() - ProcessStartTime
    AverageProcessTime = ElapsedTime / thisItem
    
    TimeRemaining = AverageProcessTime * (EventsInResult - RowCounter)

        Application.StatusBar = "Processed : " & RowCounter & " Remaining : " & (EventsInResult - RowCounter) & " and Time Remaining = " & Format(TimeRemaining, "hh:mm:ss")
 
    
    
   Next thisItem

Next endpointloop


Application.OnTime Now() + TimeSerial(0, 0, 3), "ClearStatusBar"

On Error GoTo 0

End Sub

Function FindTheValue(searchSheet As Worksheet, inColumn As Integer, thisFind)

Dim c

With searchSheet.Columns(inColumn)
    Set c = .Find(thisFind, LookIn:=xlValues, LookAt:=xlWhole)
    If Not c Is Nothing Then
        FindTheValue = c.Row
    Else
        FindTheValue = 0
    End If
End With

End Function

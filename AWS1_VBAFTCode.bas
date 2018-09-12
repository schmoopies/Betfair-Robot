Attribute VB_Name = "VBAFTCode"
Option Explicit

'/////////////////////////////////////////////////////////////////////////////////////////////////
'// CODING BY Steve Pardoe of vbafasttrack.com. Contactable at: support@vbafasttrack.com//////////
'/////////////////////////////////////////////////////////////////////////////////////////////////

'// Set up types to be used for our internal storage - this way we can pass meaningful data around
'// simply and not have to rely on multiple disparate arrays

'// Simple single event (belongs to a competition) type
Public Type SingleEvent
    eventid     As Long
    eventname   As String
    opendate    As String
End Type

'// Higher level competition type which is used to store attributes of the competition itself
Public Type Competition
    CompID      As String
    CompName    As String
    Events()    As SingleEvent
End Type

'// Top level competitions type
Public Type Competitions
    Comps() As Competition
End Type

'// Error Message Constants
Private Const PassedArrayNotDimd        As String = "The input array passed in has not been dimensioned. Execution halted for procedure: "
Private Const NullStrgRet               As String = "A null (empty) string was returned from a web call in procedure: "
Private Const JsonParseFailed           As String = "JSON parsing failed in procedure: "
Private Const ZeroParsedItems           As String = "JSON parsing returned zero items in the item collection"   '// Not yet used...
Private Const RequestStrgNullSoccerCat  As String = "Request string for getting soccer catalogue was returned as NULL string in proc ListEventMarkets"
Private Const NoObjectSoccerMatchCat    As String = "Soccer match catalogue returned in proc ListEventMarkets was nothing (i.e. no object instantiated)"
Private Const MarketIDNotFOund          As String = "Market ID not found in the provided dictionary / collection in proc ListEventMarkets for: "
Private Const MarketBookRequestFail     As String = "Received a NULL string from GetListMarketBookRequestString during execution of ListEventMarkets for MarketID: "
Private Const MarketBookResponseFail    As String = "Received a NULL string back from SendRequest call in proc ListEventMarkets for Market ID: "
Private Const MarketBookResultFail      As String = "Non instantiated object (=NOTHING) returned for ListMarketBookResult in proc ListEventMarkets"
Private Const ZeroCollectionItems       As String = "Zero iutems returned in collection object ListMarketBookResult in proc ListEventMarkets"
Private Const NonNumericCellValue       As String = "Cell value was non-numeric for eventID and market could therefore not be retrieved. MarketID attempted was: "
Private Const CompListFail              As String = "Failed to get the competition list in proc CompEventsController"

'// Event column and row markers
Public Const eventid                    As String = "eventid"
Public Const eventname                  As String = "eventname"

'// Column markers for H-A-D markets
Public Const market_id_column           As String = "market_id_column"
Public Const home_back_top_price        As String = "home_back_top_price"
Public Const away_back_top_price        As String = "away_back_top_price"
Public Const draw_back_top_price        As String = "draw_back_top_price"

Public Const home_back_price            As String = "home_back_price"
Public Const home_back_size             As String = "home_back_size"
Public Const away_back_price            As String = "away_back_price"
Public Const away_back_size             As String = "away_back_size"
Public Const draw_back_price            As String = "draw_back_price"
Public Const draw_back_size             As String = "draw_back_size"
Public Const home_lay_price             As String = "home_lay_price"
Public Const home_lay_size              As String = "home_lay_size"
Public Const away_lay_price             As String = "away_lay_price"
Public Const away_lay_size              As String = "away_lay_size"
Public Const draw_lay_price             As String = "draw_lay_price"
Public Const draw_lay_size              As String = "draw_lay_size"

Public Sub GetAccountFunds(ByVal MailOrScreen As String)

On Error GoTo ErrorHandler:

'AccountFundsResponse
'availableToBetBalance double
'Exposure double
    
'make the default Mail
If (MailOrScreen <> "Mail" And MailOrScreen <> "Screen") Then MailOrScreen = "Mail"
    
Dim Request

'Get the UK account details
Request = MakeJsonRpcAccountsRequestString(GetAccountFundsMethod, GetAccountFundsRequestString("UK"))
    Dim GetAccountFundsResponseUK As String: GetAccountFundsResponseUK = SendRequest(GetJsonRpcAccountUrl("UK"), GetAppKey(), GetSession(), Request)
    
Dim GetAccountFundsResultUK: Set GetAccountFundsResultUK = ParseJsonRpcResponseToCollection(GetAccountFundsResponseUK)
  
'Get the AUS account details
Request = MakeJsonRpcAccountsRequestString(GetAccountFundsMethod, GetAccountFundsRequestString("AUSTRALIAN"))
    Dim GetAccountFundsResponseAUS As String: GetAccountFundsResponseAUS = SendRequest(GetJsonRpcAccountUrl("AUS"), GetAppKey(), GetSession(), Request)
    
Dim GetAccountFundsResultAUS: Set GetAccountFundsResultAUS = ParseJsonRpcResponseToCollection(GetAccountFundsResponseAUS)
  
'send them both to process
Call OutputGetAccountFunds(GetAccountFundsResultUK, GetAccountFundsResultAUS, MailOrScreen)


On Error GoTo 0
Exit Sub

ErrorHandler:
    HandleError "GetAccountFunds"
    Resume Next
    
End Sub
Public Sub GetAccountFundsOnScreen()

KeepAlive
GetAccountFunds ("Screen")

End Sub
Public Sub GetEventsFromCompList(Optional Country As String)
'// Initiated from button click. Firstly, gets list of the compID's that are in the current selection
'// and then retrieves their name (also name of relevent output sheet). Passes compID's (one by one for now)
'// to ListCompEvents to get all the events for the relevant competition ID. THen responsible for populating the correct spreadsheet
'// with the list of the events for that competition

On Error GoTo ErrorHandler:

If Country <> "AUS" Then Country = "UK"
'sets default endpoint country to UK if AUS isn't specified

Dim Rng             As Range
Dim CompList        As Object
Dim CompName        As String
Dim OPSheet         As Worksheet
Dim CompEvents()    As SingleEvent
Dim CompIDS(0)      As String
Dim N               As Long
Dim eventidCol              As Long
Dim eventnameCol            As Long
Dim competitionidCol        As Long
Dim competitionnameCol      As Long
Dim competitionweightCol    As Long
Dim competitionregionCol    As Long
Dim eventtimeCol            As Long
Dim RowTally        As Long
Dim CurrentRow As Long
Dim NewRow As Long
Dim Verify

KeepAlive 'run this whenever we can/should

'// Attempt to get the comp list
Set CompList = GetCompList(Country)

If Not IsObject(CompList) Then MsgBox CompListFail: Exit Sub

For Each Rng In Selection   '// Loop through each selected cell (range)

    RowTally = 0 '// Rest is important...
    
    If IsNumeric(Trim(Rng.Value)) Then   '// So far, so good..
    
        CompName = Trim(FindCompNameFromID(CompList, Trim(Rng.Value)))
        
        '// Set our output worksheet to be the right one per the competition name. Error indicates
        '// it did not exist, therefore we will simply move on to the next one...
        On Error Resume Next
        Err.Clear
        
        Application.StatusBar = "Adding " & "nothing yet" & " for Comp " & CompName & " ... patience"
        Set OPSheet = Sheets("Future List")
        If Err.Number = 0 Then  '// Can carry on as worksheet of that name does, in fact, exist...
        
            Err.Clear
            On Error GoTo 0 '// Reset normal error handling operations
            
            CompIDS(0) = Trim(Rng.Value)
            
            eventidCol = GetHeaderColumn("event:id")
            eventnameCol = GetHeaderColumn("event:name")
            competitionidCol = GetHeaderColumn("competition:id")
            competitionnameCol = GetHeaderColumn("competition:name")
            eventtimeCol = GetHeaderColumn("event:time")
            competitionregionCol = GetHeaderColumn("competition:region")
            competitionweightCol = GetHeaderColumn("competition:weight")
            
            If ListCompEvents(CompIDS(), CompEvents(), Country) Then '// Managed to get some events back for this competitoion ID
            
                '// Get column (and row) refs for the eventID and names...
                'eventidCol = GetNamedRngColumn(eventid, CompName)
                'NAMECol = GetNamedRngColumn(eventname, CompName)
                
                'IDRow = GetNamedRngRow(eventid, CompName)
                'NAMERow = GetNamedRngRow(eventname, CompName)
            
                '// Now can go through each of them and populate the relevant output spreadsheet...
                For N = LBound(CompEvents()) To UBound(CompEvents())

                    '// Need to filter out case where event name = compname...
                    If CompEvents(N).eventname <> CompName Then
                    'we look for "v" for soccer and afl but an "@" for American Football
                        If (InStr(1, CompEvents(N).eventname, " v ", vbTextCompare) > 0) Or (InStr(1, CompEvents(N).eventname, " @ ", vbTextCompare) > 0) Then 'the string was found and it's therefore a game

                            CurrentRow = GetDataRow(CompEvents(N).eventid, eventidCol)
                            If CurrentRow = 0 Then 'it wasn't found in the current list
                                Cells(1, 1).Select
                                Application.StatusBar = "Adding " & CompEvents(N).eventid & " for Comp " & CompEvents(N).eventname & " ... patience"
                                
                                NewRow = Selection.CurrentRegion.rows.Count + 1
                                Verify = PutTheResult(OPSheet, NewRow, eventidCol, CompEvents(N).eventid, True)
                                Verify = PutTheResult(OPSheet, NewRow, eventnameCol, CompEvents(N).eventname, True)
                                Verify = PutTheResult(OPSheet, NewRow, competitionidCol, Rng.Value, True)
                                Verify = PutTheResult(OPSheet, NewRow, eventtimeCol, CompEvents(N).opendate, True)
                                Verify = PutTheResult(OPSheet, NewRow, competitionnameCol, Rng.Offset(0, 1).Value, True)
                                Verify = PutTheResult(OPSheet, NewRow, competitionregionCol, Rng.Offset(0, 3).Value, True)
                                Verify = PutTheResult(OPSheet, NewRow, competitionweightCol, Rng.Offset(0, 2).Value, True)
                                
                                
                            End If
                            'OPSheet.Cells(IDRow + 1 + RowTally, IDCol).Value = CompEvents(n).eventid
                            'OPSheet.Cells(NAMERow + 1 + RowTally, NAMECol).Value = CompEvents(n).eventname
                            'OPSheet.Cells(NAMERow + 1 + RowTally, NAMECol + 1).Value = CompEvents(n).opendate
                            
                            'RowTally = RowTally + 1
                        End If 'InStr
                    End If 'check compname

                Next N
            
            End If
        
        
        End If
        
    Else
    
        MsgBox "Cell value: " & Rng.Value & " is not numeric !", vbInformation, "Input Error"
    
    End If

Next Rng

Call CleanupApplication

On Error GoTo 0
Exit Sub

ErrorHandler:
    HandleError "GetEventsFromCompList"
    Resume Next

End Sub
Public Sub CompEventsController()
 '// Initiated from button click.

Call GetEventsFromCompList

End Sub
Public Sub CompEventsControllerAUS()
'// Initiated from button click.

Call GetEventsFromCompList("AUS")

End Sub

Public Function ListCompEvents(CompIDArray() As String, EventsArr() As SingleEvent, Optional Country As String) As Boolean
'// PURPOSE: Take a list of comp ids and return an array of events for each competition. Returns TRUE to indicate success
'// FUTURE IMPROVEMENT: Pass in a variant (then could be either a string array or a range) and have the proc sort it out accordingly
'// NOTES: Multiple events can be returned by the request string having comma separated compid's in it: PROBLEM: There is no compid returned so no way to identify them !!
'// EXAMPLE: {"jsonrpc": "2.0", "method": "SportsAPING/v1.0/listEvents", "params": {"filter":{"competitionIds":["81","105"]}}, "id": 1}

'// Author: Steve Pardoe

Dim LB                  As Long
Dim UB                  As Long
Dim N                   As Long
Dim CompIDS             As String
Dim Request             As Variant
Dim ListEventsResponse  As String
Dim ListEventsResult    As Object
Dim Index               As Long

On Error GoTo errHandler

'// Check that the passed in array has been dimensioned properly...
Err.Clear
On Error Resume Next
LB = LBound(CompIDArray())
UB = UBound(CompIDArray())
If Err.Number <> 0 Then MsgBox PassedArrayNotDimd & "ListCompEvents", vbCritical, "Error": Exit Function
Err.Clear

'// If we are at this point, the array has been properly dimensioned and we can now read
'// in the values in the preparation of our events request string
For N = LB To UB

    '// Prepare the compID string...
    CompIDS = CompIDS & Chr(34) & CompIDArray(N) & Chr(34) & ","

Next N

'// Now just check if last character is a comma, we need to remove it...
If Right(CompIDS, 1) = "," Then CompIDS = Left(CompIDS, Len(CompIDS) - 1)

'// Now formulate the request string in the normal manner
Request = MakeJsonRpcRequestString(ListEventsMethod, MULTListEventsRequestString(CompIDS))

ListEventsResponse = SendRequest(GetJsonRpcUrl(Country), GetAppKey(), GetSession(), Request)

If ListEventsResponse <> vbNullString Then

    Set ListEventsResult = ParseJsonRpcResponseToCollection(ListEventsResponse)
    
    If ListEventsResult Is Nothing Then MsgBox JsonParseFailed & "ListCompEvents", vbCritical, "System Error": Exit Function
    
    If ListEventsResult.Count <= 0 Then Exit Function   '// Could use the ZeroParsedItems here and give a warning but in some cases, this may be ok, so don't want to stop execution abruptly
    
    '// Can now resize the event return array...(zero based)
    ReDim EventsArr(0 To (ListEventsResult.Count - 1))
    
    '// At this point, we have got a JSON parsed success with a number of items >= 1 that we can now iterate through and start populating our passed in return array...
    For Index = 1 To ListEventsResult.Count Step 1
    
        EventsArr(Index - 1).eventid = ListEventsResult.Item(Index).Item("event").Item("id")
        EventsArr(Index - 1).eventname = ListEventsResult.Item(Index).Item("event").Item("name")
        EventsArr(Index - 1).opendate = ListEventsResult.Item(Index).Item("event").Item("openDate")
    Next Index
    
    '// Finally set the boolean return value to indicate success here
    ListCompEvents = True

Else

    MsgBox NullStrgRet & "ListCompEvents", vbCritical, "System / Web Error": Exit Function

End If

Exit Function

errHandler:

    MsgBox "An error occurred in procedure ListCompEvents:" & vbCrLf & vbCrLf & "Error #: " & Err.Number & vbCrLf & "Description: " & Err.Description

End Function
Public Sub GetMarketBookFromMarketId()
'as it says, use a MarketID to get MarketBook for the whole marketid
Dim Request
Dim marketid: marketid = Trim(Selection.Cells.Value)

Request = MakeJsonRpcRequestString(ListMarketBookMethod, GetListMarketBookRequestString(marketid))
    Dim ListMarketBookResponse As String: ListMarketBookResponse = SendRequest(GetJsonRpcUrl(), GetAppKey(), GetSession(), Request)
    
Dim MarketBookResult: Set MarketBookResult = ParseJsonRpcResponseToCollection(ListMarketBookResponse)
    'Dim EventTypeId: EventTypeId = GetEventTypeIdFromEventTypes(EventTypeResult)
    'Cells(OutputRow + 12, OutputColumn).Value = EventTypeId
    '
    ' Call listMarketCatalogue to find the next horse race about to start and extract the market id from the response
    '
'    Request = GetListMarketCatalogueRequestStringNF(EventTypeId)
'    Dim ListMarketCatalogueResponse As String: ListMarketCatalogueResponse = SendRequest(GetJsonRpcUrl(), GetAppKey(), GetSession(), Request)
'    Cells(OutputRow + 14, OutputColumn).Value = ListMarketCatalogueResponse
'    AppendToLogFile ListMarketCatalogueResponse

Call OutputListMarketBook(MarketBookResult, 13, 1)

End Sub

Public Function GetEventTypeIdFromMarketId()
Dim Request

Request = MakeJsonRpcRequestString(ListEventTypesMethod, GetListEventTypesRequestString())
    Dim ListEventTypesResponse As String: ListEventTypesResponse = SendRequest(GetJsonRpcUrl(), GetAppKey(), GetSession(), Request)
    
Dim EventTypeResult: Set EventTypeResult = ParseJsonRpcResponseToCollection(ListEventTypesResponse)
    Dim EventTypeId: EventTypeId = GetEventTypeIdFromEventTypes(EventTypeResult)
    Cells(OutputRow + 12, OutputColumn).Value = EventTypeId
    '
    ' Call listMarketCatalogue to find the next horse race about to start and extract the market id from the response
    '
    Request = GetListMarketCatalogueRequestStringNF(EventTypeId)
    Dim ListMarketCatalogueResponse As String: ListMarketCatalogueResponse = SendRequest(GetJsonRpcUrl(), GetAppKey(), GetSession(), Request)
    Cells(OutputRow + 14, OutputColumn).Value = ListMarketCatalogueResponse
    Call AppendToLogFile(ListMarketCatalogueResponse, "GetEventTypeIdFromMarketId", 150)


End Function
Public Function GetMatchCatalogueFromEventID(ByRef eventid As String, ByRef OutputType As String, Optional ByVal EndPoint As String)

On Error GoTo ErrorHandler:
'Dim eventid: eventid = Selection.Cells.Value
Dim Request
    
    Request = MakeJsonRpcRequestString(ListMatchCatalogueMethod, GetListMatchCatalogueRequestString(eventid))
    
    'Request = MakeJsonRpcRequestString(ListMatchCatalogueMethod, GetListMatchCatalogueRequestString4MyEventTypes(eventid))
    Dim ListMatchCatalogueResponse As String: ListMatchCatalogueResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
    'Cells(OutputRow + 7, OutputColumn).Value = ListSoccerMatchCatalogueResponse
    AppendToLogFile ListMatchCatalogueResponse
    
    Dim ListMatchResult As Object: Set ListMatchResult = ParseJsonRpcResponseToCollection(ListMatchCatalogueResponse)
    OutputListMatchResult ListMatchResult, OutputRow + 30, OutputColumn - 1, OutputType
    
On Error GoTo 0
Exit Function

ErrorHandler:
    HandleError "GetMatchCatalogueFromEventID"
    Resume Next
    
    
End Function

Public Function GetOpenDateFromEventID(ByRef eventid As String, Optional ByVal EndPoint As String) As String

On Error GoTo ErrorHandler:
'Dim eventid: eventid = Selection.Cells.Value
Dim Request
    
    Request = MakeJsonRpcRequestString(ListEventsMethod, GetListSingleEventRequestString(eventid))
    Dim ListEventsResponse As String: ListEventsResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
    
    AppendToLogFile ListEventsResponse
    
    Dim ListEventsResult As Object: Set ListEventsResult = ParseJsonRpcResponseToCollection(ListEventsResponse)
    'OutputListMatchResult ListMatchResult, OutputRow + 30, OutputColumn - 1, OutputType
    
    GetOpenDateFromEventID = ListEventsResult.Item(1).Item("event").Item("openDate")
    
On Error GoTo 0
Exit Function

ErrorHandler:
    HandleError "GetOpenDateFromEventID"
    Resume Next
        
End Function

Public Sub ListEventMarkets()
'// PURPOSE: Uses user selection of ranges to determine which events markets are required for and fetch the following markets for each event (WHERE AVAILABLE):
'// BACK PRICE (H-A-D)
'// BACK PRICES AND VOLUME (SIZE) FOR TOP OF MARKET H-A-D
'// LAY PRICES AND VOLUME (SIZE) FOR TOP OF MARKET H-A-D
'// Relevant sheet (active sheet) is then populated in accordance with the defined names (markers or anchors) on the sheet
'// Makes use of AppendToLogFile procedure

'// FUTURE IMPROVEMENT: May be able to make multiple event calls at once and parse the returned JSON as long as there is an event ID there (cut down time and bandwidth)

Dim N                                   As Long
Dim Rng                                 As Range
Dim Request                             As Variant
Dim ThisEventID                         As String
Dim ListSoccerMatchCatalogueResponse    As String
Dim ListSoccerMatchCatalogueResult      As Object
Dim ListMarketBookResponse              As String
Dim ThisMarketID                        As String
Dim ListMarketBookResult                As Object
Dim OPRow                               As Long

'On Error GoTo errHandler
On Error GoTo ErrorHandler:

For Each Rng In Selection

    '// Get Event ID and check it is numeric, else go to next range in current selection
    ThisEventID = Trim(Rng.Value)
    
    If IsNumeric(ThisEventID) Then

        '// Get request string for our filter...
        Request = Trim(MakeJsonRpcRequestString(ListSoccerMatchCatalogueMethod, GetListSoccerMatchCatalogueRequestString(ThisEventID)))
        
        '// Check for null string. Even though unlikely, still a valid test else we would have to deal with additional web errors later in any case
        If Request <> vbNullString Then
        
            '// Send request to betfair and receive response...
            ListSoccerMatchCatalogueResponse = SendRequest(GetJsonRpcUrl(), GetAppKey(), GetSession(), Request)
            
            '// Get a collection / dictionary object (SoccerMatchCatalogue) of parsed JSON arrays and items...
            Set ListSoccerMatchCatalogueResult = ParseJsonRpcResponseToCollection(ListSoccerMatchCatalogueResponse)
            
            '// Now test that an object was successfully returned
            If IsObject(ListSoccerMatchCatalogueResult) Then
            
                '// Record current (event) row number for later use in markets population
                OPRow = Rng.Row
            
                ThisMarketID = Trim(GetMarketIdFromMatchCatalogue("Match Odds", ListSoccerMatchCatalogueResult))
                
                If ThisMarketID <> vbNullString Then    '// Best to make sure all good...
                
                    Request = vbNullString  '// Just in case as we're using it above...
                    
                    '// Formulate request to get markets for thsi request ID...
                    Request = Trim(MakeJsonRpcRequestString(ListMarketBookMethod, GetListMarketBookRequestStringV(ThisMarketID)))
                    
                    If Request <> vbNullString Then
                    
                        ListMarketBookResponse = Trim(SendRequest(GetJsonRpcUrl(), GetAppKey(), GetSession(), Request))
                        
                        If ListMarketBookResponse <> vbNullString Then
                            
                            Set ListMarketBookResult = ParseJsonRpcResponseToCollection(ListMarketBookResponse)
                            
                            If IsObject(ListMarketBookResult) Then  '// Instantiated object returned or not ?
                            
                                If ListMarketBookResult.Count > 0 Then
                            
                                    '// Populate marketID on sheet at appropriate place...
                                    Cells(OPRow, GetNamedRngColumn(market_id_column)).Value = ThisMarketID
                                    
                                    '// Finally we get to display the markets on the worksheet (activesheet)
                                    Call DisplaySoccerMarkets(ListMarketBookResult, OPRow)
                                
                                Else
                                
                                    AppendToLogFile ZeroCollectionItems: NotifyUserOfError
                                
                                End If
                            
                            Else    '// MarketBookResultFail
                            
                                AppendToLogFile MarketBookResultFail: NotifyUserOfError
                            
                            End If
                        
                        Else    '// MarketBookResponseFail
                        
                            AppendToLogFile MarketBookResponseFail & ThisMarketID: NotifyUserOfError
                        
                        End If
                    
                    Else    '// MarketBookRequestFail
                    
                        AppendToLogFile MarketBookRequestFail & ThisMarketID: NotifyUserOfError
                    
                    End If
                
                Else    '// Required marketID not found in JSON collection / dictionary
                
                    AppendToLogFile MarketIDNotFOund & "Match Odds": NotifyUserOfError
                
                End If
            
            Else    '// Not an object returned
            
                AppendToLogFile NoObjectSoccerMatchCat: NotifyUserOfError
            
            End If
        
        Else    '// (Request = Null)
        
            AppendToLogFile RequestStrgNullSoccerCat: NotifyUserOfError
        
        End If
        
    Else    '// Selected cell value is not numeric - tell the user this and let them know of the address of range that was used
    
        AppendToLogFile NonNumericCellValue & ThisEventID & " in cell: " & Rng.Address: NotifyUserOfError
    
    End If

Next Rng

On Error GoTo 0
Exit Sub

ErrorHandler:
    HandleError "ListEventMarkets"
    Resume Next
    
End Sub
Public Sub PlaceOrders()

On Error GoTo ErrorHandler:

'Fields are:
'MarketId
'SelectionID
'handicap = 0
'ordertype = LIMIT
'side = BACK/LAY
'size = size
'price = price

Dim PlaceOrdersResult As Object
Dim ListCurrentOrdersResult As Object
Dim EndPoint As String: EndPoint = Cells(4, 1).Value
Dim N As Integer
Dim Request As Variant
Dim tempHandicap As String


Dim marketid As String
Dim selectionID As String
Dim BetID As String
Dim Handicap As Double: Handicap = 0
Dim orderType As String: orderType = "LIMIT"
Dim orderSide As String
Dim orderSize: orderSize = Selection.Value
Dim orderPrice As Double
Dim MyMaxOdds As Integer
Dim MyMinOdds As Integer
Dim MyLowOdds As Double
Dim BadList As Boolean: BadList = False

Dim ThisResult As String
Dim ThisMarketName As String

If Selection.Column = 9 Then
    orderType = "BACK"
    'orderPrice = Cells(Selection.Row, 6).Value
    orderPrice = Cells(Selection.Row, 8).Value
ElseIf Selection.Column = 11 Then
    orderType = "LAY"
    'orderPrice = Cells(Selection.Row, 7).Value
    orderPrice = Cells(Selection.Row, 10).Value
Else: Exit Sub
End If

'Match Odds
'        If Cells(Selection.Row, 2).Value = "Match Odds" Then
'            'we are not going to place bets on Match Odds type bets for a while
'            Exit Sub
'        End If

MyMaxOdds = Sheets("Example").Cells(GetNamedRngRow("MaxOdds", "Example"), GetNamedRngColumn("MaxOdds", "Example")).Value
MyMinOdds = Sheets("Example").Cells(GetNamedRngRow("MinOdds", "Example"), GetNamedRngColumn("MinOdds", "Example")).Value
MyLowOdds = Sheets("Example").Cells(GetNamedRngRow("LowOdds", "Example"), GetNamedRngColumn("LowOdds", "Example")).Value

If orderPrice > MyMaxOdds Then
    AppendToLogFile "Place Orders - Odds Too Big:" & orderPrice & " Greater than " & MyMaxOdds
    Exit Sub                            'don't want to bet or lay if too big
End If

'If ((orderPrice < MyMinOdds) And (orderPrice > MyLowOdds)) Then
'    AppendToLogFile "Place Orders - Odds Too Small:" & orderPrice & " Less than " & MyMinOdds
'    Exit Sub                            'don't want to bet or lay if too small
'End If

If Selection.Value > 0 Then
'do something since there's a value there
    selectionID = Cells(Selection.Row, 3).Value
    marketid = Cells(Selection.Row, 1).Value
    tempHandicap = Cells(Selection.Row, 5).Value 'put something there until we check if we need it
    
If (selectionID = "4207181" Or selectionID = "4207182" Or selectionID = "4207183" Or selectionID = "4207184" Or selectionID = "4207185" Or selectionID = "4207186" Or selectionID = "4207187" Or selectionID = "4207188" Or selectionID = "4207189") Then  'Or selectionID = "9353970"
    BadList = True
        If orderType = "BACK" And Selection.Offset(25, -5).Value > 1 And orderPrice > Selection.Offset(25, -5).Value Then
            BadList = False
        End If 'this whole step is to allow BACK bets to be made IFF Sportsbet prices have been found and the price is over it
Else
    BadList = False
End If
    
'ADDD THIS extra FOR AFL 221 POINTS OR MORE WHILE I WORK IT OUT 11th June 2018
If selectionID <> "" And ((BadList = False) Or (BadList = True And orderType = "LAY")) Then 'that's good
    
    
    If marketid = "" Then
        Cells(Selection.Row, 1).Select
        marketid = Selection.End(xlUp).Value
    End If

            If Len(marketid) <> 11 Then
                For N = 1 To 11 - Len(marketid)
                marketid = "" & marketid & "0" 'Had problems with trailing 0s missing
                Next N
            End If
            
 ThisMarketName = GetMarketNameForMarketID(marketid) 'needed for NFL matches
 
                If (ThisMarketName = "Handicap") Or (ThisMarketName = "Total Points") Or (ThisMarketName = "Total Goals") Then
                    Handicap = CDbl(tempHandicap)
                Else
                    Handicap = 0
                End If
            
 'NOW - before placing a NEW order, we need to check if there is an unmatched current
 'order for this marketId and Selection Id (and bettype)
 'if so, we can cancel just THAT BetId and place a new order
            
 Request = MakeJsonRpcRequestString(ListCurrentOrdersMethod, GetListCurrentOrdersRequestString(marketid))
        Dim ListCurrentOrdersResponse As String: ListCurrentOrdersResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
        AppendToLogFile "List Current Orders " & ListCurrentOrdersResponse
    
    Set ListCurrentOrdersResult = ParseJsonRpcResponseToCollection(ListCurrentOrdersResponse)
            
 'NOW use the result to find the Selection ID (if it exists), get the BetID associated with any sizeRemaining
 'and then CANCEL it
            
 BetID = GetBetIDFromListCurrentOrders(ListCurrentOrdersResult, selectionID, orderType)

    'Cancel it if it exists
    If BetID <> "" Then ThisResult = CancelOrdersForBetID(marketid, BetID)
                        
                    
            Request = MakeJsonRpcRequestString(PlaceOrdersMethod, GetPlaceOrdersRequestStringHandicap(marketid, selectionID, orderType, orderPrice, orderSize, Handicap))
                Dim PlaceOrdersResponse As String: PlaceOrdersResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
                AppendToLogFile "Place Orders " & PlaceOrdersResponse
            
            Set PlaceOrdersResult = ParseJsonRpcResponseToCollection(PlaceOrdersResponse)
            
            Cells(Selection.Row, 28) = PlaceOrdersResponse 'temporary
            
            'Call OutputListCurrentOrders(ListCurrentOrdersResult, ThisSelectionId, MyCell.Row)
        
End If 'selectionID check
End If 'selection.value check


On Error GoTo 0
    Exit Sub
        
ErrorHandler:
    HandleError "PlaceOrders"
    Resume Next
    

End Sub
Public Sub ReplaceOrders()

On Error GoTo ErrorHandler:

'Fields are:
'MarketId
'SelectionID NOW BetID
'handicap = 0
'ordertype = LIMIT
'side = BACK/LAY
'size = size
'price = price

Dim ReplaceOrdersResult As Object
Dim EndPoint As String: EndPoint = Cells(4, 1).Value
Dim N As Integer
Dim Request As Variant


Dim marketid As String
'Dim selectionId As String
'Dim handicap As Double: handicap = 0
'Dim orderType As String: orderType = "LIMIT"
'Dim orderSide As String
'Dim orderSize: orderSize = Selection.Value
Dim BetID: BetID = Selection.Value
Dim orderPrice As Double


If (Selection.Column = 28) Or (Selection.Column = 35) Then  'correct column
    'orderType = "BACK"
    'orderPrice = Cells(Selection.Row, 6).Value
    orderPrice = Cells(Selection.Row, Selection.Column + 2).Value 'New price is entered here over the old/current price
'ElseIf Selection.Column = 11 Then
    'orderType = "LAY"
    'orderPrice = Cells(Selection.Row, 7).Value
    'orderPrice = Cells(Selection.Row, 10).Value
Else: Exit Sub
End If


If Selection.Value > 0 Then
'do something since there's a value there
    'selectionId = Cells(Selection.Row, 3).Value
    marketid = Cells(Selection.Row, 1).Value
    
    If marketid = "" Then
        Cells(Selection.Row, 1).Select
        marketid = Selection.End(xlUp).Value
    End If

            If Len(marketid) <> 11 Then
                For N = 1 To 11 - Len(marketid)
                marketid = "" & marketid & "0" 'Had problems with trailing 0s missing
                Next N
            End If
            
    Request = MakeJsonRpcRequestString(ReplaceOrdersMethod, GetReplaceOrdersRequestString(marketid, BetID, orderPrice))
        
    Dim ReplaceOrdersResponse As String: ReplaceOrdersResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
    AppendToLogFile "Replace Orders " & ReplaceOrdersResponse
    
    Set ReplaceOrdersResult = ParseJsonRpcResponseToCollection(ReplaceOrdersResponse)
    
    Cells(Selection.Row, 28) = ReplaceOrdersResponse 'temporary
    
    'Call OutputListCurrentOrders(ListCurrentOrdersResult, ThisSelectionId, MyCell.Row)

End If

On Error GoTo 0
    Exit Sub
        
ErrorHandler:
    HandleError "ReplaceOrders"
    Resume Next


End Sub
Public Sub CancelOrders()

On Error GoTo ErrorHandler:

'Fields are:
'MarketId
'SelectionID NOW BetID
'handicap = 0
'ordertype = LIMIT
'side = BACK/LAY
'size = size
'price = price

Dim CancelOrdersResult As Object
Dim EndPoint As String: EndPoint = Cells(4, 1).Value
Dim N As Integer
Dim Request As Variant


Dim marketid As String
Dim BetID As String
Dim orderPrice As Double
Dim marketPrice As Double

Dim thisSelection As Range

For Each thisSelection In Selection

        If (thisSelection.Column = 28) Or (thisSelection.Column = 35) Then  'we are cancelling specific bets
          
            If thisSelection.Value > 0 Then
                orderPrice = Cells(Selection.Row, Selection.Column + 6).Value 'New price is entered here over the old/current price
                BetID = thisSelection.Value
                marketid = Cells(thisSelection.Row, 1).Value
            End If
            
        ElseIf thisSelection.Column = 1 Then
        
            If thisSelection.Value > 0 Then
                marketPrice = thisSelection.Value 'New price is entered here over the old/current price
            End If
            
        Else: Exit Sub
        End If
        
                    
            If marketid = "" Then
                Cells(Selection.Row, 1).Select
                marketid = Selection.End(xlUp).Value
            End If
        
                    If Len(marketid) <> 11 Then
                        For N = 1 To 11 - Len(marketid)
                        marketid = "" & marketid & "0" 'Had problems with trailing 0s missing
                        Next N
                    End If
         
        
            Request = MakeJsonRpcRequestString(CancelOrdersMethod, GetCancelOrdersRequestString(marketid, BetID, orderPrice))
                
            Dim CancelOrdersResponse As String: CancelOrdersResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
            AppendToLogFile "Cancel Orders " & CancelOrdersResponse
            
            Set CancelOrdersResult = ParseJsonRpcResponseToCollection(CancelOrdersResponse)
            
            Cells(Selection.Row, 28) = CancelOrdersResponse 'temporary
            
            'Call OutputListCurrentOrders(ListCurrentOrdersResult, ThisSelectionId, MyCell.Row)
        
       
Next

On Error GoTo 0
    Exit Sub
        
ErrorHandler:
    HandleError "CancelOrders"
    Resume Next


End Sub
Public Sub ListCurrentOrders()
'as it says, list the current orders for an event page
'check the price column for both Back and Lay
'if there is a price there then check to see if the price has been matched

On Error GoTo ErrorHandler:

Dim Request
Dim MyCell
Dim N As Integer

Dim MyLoop As Integer
Dim BacklayAmounts As Range
Dim BackOrLay As Boolean
Dim ThisSelectionId As String
Dim EndPoint As String: EndPoint = Cells(4, 1).Value
Dim ListCurrentOrdersResult As Object

Dim marketid As String
Dim Side As String

For MyLoop = 0 To 2 Step 2 'do it twice

        If MyLoop = 0 Then
            If Cells(3, 1) = "Soccer" Then
                Set BacklayAmounts = Range(Cells(8, 9), Cells(SoccerLength, 9))
            Else
                Set BacklayAmounts = Range(Cells(8, 9), Cells(AFLLength, 9))
            End If
            BackOrLay = True
            Side = "BACK"
        Else
            If Cells(3, 1) = "Soccer" Then
                Set BacklayAmounts = Range(Cells(8, 11), Cells(SoccerLength, 11))
            Else
                Set BacklayAmounts = Range(Cells(8, 11), Cells(AFLLength, 11))
            End If
            BackOrLay = False
            Side = "LAY"
        End If

For Each MyCell In BacklayAmounts
'calculate the values in the array
If MyCell.Value > 0 Then
'do something since there's a value there
    ThisSelectionId = Cells(MyCell.Row, 3).Value
    marketid = Cells(MyCell.Row, 1).Value
    
    If marketid = "" Then
        Cells(MyCell.Row, 1).Select
        marketid = Selection.End(xlUp).Value
    End If

            If Len(marketid) <> 11 Then
                For N = 1 To 11 - Len(marketid)
                marketid = "" & marketid & "0" 'Had problems with trailing 0s missing
                Next N
            End If
            
    Request = MakeJsonRpcRequestString(ListCurrentOrdersMethod, GetListCurrentOrdersRequestString(marketid))
        Dim ListCurrentOrdersResponse As String: ListCurrentOrdersResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
        AppendToLogFile "List Current Orders " & ListCurrentOrdersResponse
    
    Set ListCurrentOrdersResult = ParseJsonRpcResponseToCollection(ListCurrentOrdersResponse)
    
    'Cells(MyCell.Row, 28) = ListCurrentOrdersResponse 'temporary
    
    Call OutputListCurrentOrders(ListCurrentOrdersResult, ThisSelectionId, MyCell.Row, Side)

End If

Next 'myCell
Next 'myLoop back lay repeat

On Error GoTo 0
Exit Sub

ErrorHandler:
    HandleError "ListCurrentOrders"
    Resume Next

End Sub
Public Sub ListMarketRunners(ByVal EndPoint As String)
'// PURPOSE: Uses user selection of ranges to determine which markets are required and fetch the following markets for each market (WHERE AVAILABLE):
'// BACK AND LAY PRICE
'// Relevant sheet (active sheet) is then populated in accordance with the defined names (markers or anchors) on the sheet
'// Makes use of AppendToLogFile procedure

On Error GoTo ErrorHandler:

Dim N                                   As Long
Dim Rng                                 As Range
Dim Request                             As Variant
Dim ThisMarketID                        As String
Dim ListEventCatalogueResponse          As String
Dim ListEventCatalogueResult            As Object
Dim ListMarketBookResponse              As String

Dim ListMarketBookResult                As Object
Dim ThisSelectionId                     As String
Dim OPRow                               As Long
Dim OffsetCounter As Integer
Dim ThisMarketName As String

'On Error GoTo errHandler

OffsetCounter = 0


For Each Rng In Selection
'OPRow = Rng.Row + OffsetCounter
OPRow = Rng.Row

    '// Get Event ID and check it is numeric, else go to next range in current selection
    ThisMarketID = Trim(Rng.Value)
    If IsNumeric(ThisMarketID) Then
        If Len(ThisMarketID) <> 11 Then
            For N = 1 To 11 - Len(ThisMarketID)
            ThisMarketID = "" & ThisMarketID & "0" 'Had problems with trailing 0s missing
            Next N
        End If
    End If
    
 
    
    If IsNumeric(ThisMarketID) Then

        ThisMarketName = GetMarketNameForMarketID(ThisMarketID) 'needed for NFL matches

        '// Get request string for our filter...
        Request = Trim(MakeJsonRpcRequestString(ListMarketBookMethod, GetListMarketBookRequestStringV(ThisMarketID)))
        
        '// Check for null string. Even though unlikely, still a valid test else we would have to deal with additional web errors later in any case
        If Request <> vbNullString Then
        
            '// Send request to betfair and receive response...
            
            ListMarketBookResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
            
            '// Get a collection / dictionary object (SoccerMatchCatalogue) of parsed JSON arrays and items...
            
            Set ListMarketBookResult = ParseJsonRpcResponseToCollection(ListMarketBookResponse)
            
            '// Now test that an object was successfully returned
            If IsObject(ListMarketBookResult) Then
            
                If (ThisMarketName = "Handicap") Or (ThisMarketName = "Total Points") Then
                    OffsetCounter = OffsetCounter + OutputListMarketBook(ListMarketBookResult, OPRow, 4, ThisMarketID) - 1
                Else
                    OffsetCounter = OffsetCounter + OutputListMarketBook(ListMarketBookResult, OPRow, 4) - 1
                End If
                
            Else    '// Not an object returned
            
                AppendToLogFile NoObjectSoccerMatchCat: NotifyUserOfError
            
            End If
        
        Else    '// (Request = Null)
        
            AppendToLogFile RequestStrgNullSoccerCat: NotifyUserOfError
        
        End If
        
    Else    '// Selected cell value is not numeric - tell the user this and let them know of the address of range that was used
    
        'AppendToLogFile NonNumericCellValue & ThisMarketID & " in cell: " & Rng.Address: NotifyUserOfError
    
    End If

Next Rng

On Error GoTo 0
Exit Sub

ErrorHandler:
    HandleError "ListMarketRunners"
    Resume Next
    
End Sub

'//////////////////////////////////////////////////////////////////////
'// OUR UTILITIES /////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////

Private Sub OutputCompEvents(OPSheet As Worksheet, EventList As Object, EventROW As Long, eventidCol As Long, eventnameCol As Long) '// NOT BEING USED PRESENTLY...

Dim Index As Long

On Error Resume Next

For Index = 1 To (ListEventsResult.Count)
    OPSheet.Cells(EventROW + Index, eventidCol).Value = ListEventsResult.Item(Index).Item("event").Item("id")
    OPSheet.Cells(EventROW + Index, eventnameCol).Value = ListEventsResult.Item(Index).Item("event").Item("name")
Next

End Sub

Private Function GetCompList(Country As String) As Object
'// Takes no params.
'// Returns a parsed collection of competition objects (id, name, marketcount and compregion)

On Error Resume Next

Dim Request As Variant
Dim ListCompetitionsResponse As String:


'// Formulate request string
Request = MakeJsonRpcRequestString(ListCompetitionsMethod, GetListCompetitionsRequestString())

'// Make web request and get response
ListCompetitionsResponse = SendRequest(GetJsonRpcUrl(Country), GetAppKey(), GetSession(), Request)

'// Create collection object
Set GetCompList = ParseJsonRpcResponseToCollection(ListCompetitionsResponse)

End Function

Private Function FindCompNameFromID(CompListResult As Object, ID As String) As String
'// Loops through the competition items of the collection and matches on "id" and returns the associated name element
'// Naturally if passed in ID is not found then a NULL string is returned...

On Error Resume Next

Dim Index As Long


For Index = 1 To CompListResult.Count Step 1

    If CompListResult.Item(Index).Item("competition").Item("id") = ID Then  '// Have found our target, now to get (and return immediately) the associated competition name
        FindCompNameFromID = CompListResult.Item(Index).Item("competition").Item("name")
    End If
    
Next

End Function

Private Sub DisplaySoccerMarkets(MarketsCollec As Object, OPRow As Long)
'// Simply takes in a collection object of the H-A-D soccer market types and displays markets on activesheet
'// on the passed in output row (OPRow variable)

On Error Resume Next

'// Can now populate market cells
With MarketsCollec.Item(1).Item("runners")

    '// Straight top of book back prices
    ActiveSheet.Cells(OPRow, GetNamedRngColumn(home_back_top_price)).Value = .Item(1).Item("ex").Item("availableToBack").Item(1).Item("price")
    ActiveSheet.Cells(OPRow, GetNamedRngColumn(away_back_top_price)).Value = .Item(2).Item("ex").Item("availableToBack").Item(1).Item("price")
    ActiveSheet.Cells(OPRow, GetNamedRngColumn(draw_back_top_price)).Value = .Item(3).Item("ex").Item("availableToBack").Item(1).Item("price")
    
    '// BACK METRICS////////////
    
    '// Populate HOME BACK cells
    ActiveSheet.Cells(OPRow, GetNamedRngColumn(home_back_price)).Value = .Item(1).Item("ex").Item("availableToBack").Item(1).Item("price")
    ActiveSheet.Cells(OPRow, GetNamedRngColumn(home_back_size)).Value = .Item(1).Item("ex").Item("availableToBack").Item(2).Item("price")
    
    '// Populate AWAY BACK cells
    ActiveSheet.Cells(OPRow, GetNamedRngColumn(away_back_price)).Value = .Item(2).Item("ex").Item("availableToBack").Item(1).Item("price")
    ActiveSheet.Cells(OPRow, GetNamedRngColumn(away_back_size)).Value = .Item(2).Item("ex").Item("availableToBack").Item(2).Item("price")
    
    '// Populate DRAW BACK cells
    ActiveSheet.Cells(OPRow, GetNamedRngColumn(draw_back_price)).Value = .Item(3).Item("ex").Item("availableToBack").Item(1).Item("price")
    ActiveSheet.Cells(OPRow, GetNamedRngColumn(draw_back_size)).Value = .Item(3).Item("ex").Item("availableToBack").Item(2).Item("price")
    
    '// LAY METRICS////////////
    
    '// Populate HOME LAY cells
    ActiveSheet.Cells(OPRow, GetNamedRngColumn(home_lay_price)).Value = .Item(1).Item("ex").Item("availableToLay").Item(1).Item("price")
    ActiveSheet.Cells(OPRow, GetNamedRngColumn(home_lay_size)).Value = .Item(1).Item("ex").Item("availableToLay").Item(2).Item("price")
    
    '// Populate AWAY LAY cells
    ActiveSheet.Cells(OPRow, GetNamedRngColumn(away_lay_price)).Value = .Item(2).Item("ex").Item("availableToLay").Item(1).Item("price")
    ActiveSheet.Cells(OPRow, GetNamedRngColumn(away_lay_size)).Value = .Item(2).Item("ex").Item("availableToLay").Item(2).Item("price")
    
    '// Populate DRAW LAY cells
    ActiveSheet.Cells(OPRow, GetNamedRngColumn(draw_lay_price)).Value = .Item(3).Item("ex").Item("availableToLay").Item(1).Item("price")
    ActiveSheet.Cells(OPRow, GetNamedRngColumn(draw_lay_size)).Value = .Item(3).Item("ex").Item("availableToLay").Item(2).Item("price")

End With

On Error GoTo 0

End Sub

Private Sub NotifyUserOfError(Optional ErrMsg As String = "")
'// Simply notifies user that an error has occurred and to view the log file.
'// If no error string passed in, generic one used
If ErrMsg = "" Then ErrMsg = "Error. Please check Log File"

Application.StatusBar = ErrMsg

'// Set a callback to clear any left user messages on status bar (after 3 seconds)
Application.OnTime Now() + TimeSerial(0, 0, 3), "ClearStatusBar"

End Sub

Public Sub ClearStatusBar()
'// Give control of the status bar back to Excel (wipes any of our custom messages away)

Application.StatusBar = False

End Sub

Private Function MULTListEventsRequestString(ByVal CompID As String) As String
'// Variation on betfair provided VBA example utility. This one just introduces the compid as presented.
'// Responsibility of the calling procedure to surrond compids with quotes and comma separate multiple items
    MULTListEventsRequestString = "{""filter"":{""competitionIds"":[" & CompID & "]}}"
End Function

Public Function GetNamedRngColumn(Nme As String, Optional SheetName As String) As Long
'// Takes in name of range and passes back the column of that range on the ACTIVESHEET
'// If fails, return value will be zero (naturally cannot have a column number of 0, so this indicates failure)
On Error Resume Next

If SheetName = vbNullString Then
    GetNamedRngColumn = Range(ActiveSheet.Names(Nme).RefersTo).Column
Else
    GetNamedRngColumn = Range(Sheets(SheetName).Names(Nme).RefersTo).Column
End If

End Function

Public Function GetNamedRngRow(Nme As String, Optional SheetName As String) As Long
'// Takes in name of range and passes back the row of that range on the ACTIVESHEET as default
'// If fails, return value will be zero (naturally cannot have a column number of 0, so this indicates failure)
On Error Resume Next

If SheetName = vbNullString Then
    GetNamedRngRow = Range(ActiveSheet.Names(Nme).RefersTo).Row
Else
    GetNamedRngRow = Range(Sheets(SheetName).Names(Nme).RefersTo).Row
End If

End Function
Sub GenerateMarketIDList()

Dim eventid As String: eventid = Cells(7, 2).Value
Dim ThisEndPoint As String: ThisEndPoint = Cells(4, 1).Value

Dim SendIt: SendIt = GetMatchCatalogueFromEventID(eventid, "Sheet", ThisEndPoint)

End Sub
'//////////////////////////////////////////////////////////////////////
'// TEST PROCEDURES BELOW THIS LINE (CAN DELETE) //////////////////////
'//////////////////////////////////////////////////////////////////////

Sub GenerateMarketIDListAFL()

Dim eventid As String: eventid = Cells(10, 15).Value

Dim SendIt: SendIt = GetMatchCatalogueFromEventID(eventid, "Horiz")

End Sub
Sub testgetcomplist()

Dim CompObj As Object

Set CompObj = GetCompList()

MsgBox FindCompNameFromID(CompObj, "81")    '// Should display "Serie A" - confirmed all working

End Sub

Sub teststatusbarerrors()

NotifyUserOfError

End Sub

Sub testgetevents()

Dim Comps()         As SingleEvent
Dim CompIDS(0) As String
Dim mytest As String

CompIDS(0) = "81"
'CompIDS(1) = "105"

'mytest = GetOpenDateFromEventID("27506624")
mytest = GetListCurrentOrdersRequestString("27506624")

MsgBox ListCompEvents(CompIDS(), Comps())

Debug.Print

End Sub

Sub debugeventsrequest()

Dim Request As Variant
Dim CompID As Long: CompID = 81
Dim thiscomp As Variant
Dim thisCompId As Variant

thisCompId = CompID

'Commented following 3 lines
'Dim url As String: url = "https://api.betfair.com/exchange/betting/json-rpc/v1/"
'Dim appkey As String: appkey = "SooHYWoVcP3bP3W7"
'Dim cookie As String: cookie = "RcE9NVdUXT0uNtXhUwPGpJt8a83NlZP4xa133vlXci4="

Request = MakeJsonRpcRequestString(ListEventsMethod, GetListEventsRequestString(99, thisCompId))

'Dim ListEventsResponse As String: ListEventsResponse = SendRequest(url, appkey, cookie, Request)
'Blocked SP line, use the session etc from the spreadsheet - easier to update when logged out
Dim ListEventsResponse As String: ListEventsResponse = SendRequest(GetJsonRpcUrl(), GetAppKey(), GetSession(), Request)
    

Debug.Print

End Sub
Sub UpdateMarketPrices()

Dim ThisEndPoint As String: ThisEndPoint = Cells(4, 1).Value
Dim CurrentTotalmatched As Double

'let's NOT clear this and see what happens
'Range(Cells(8, 6), Cells(47, 7)).Select
'Selection.ClearContents

            
            CurrentTotalmatched = GetTotalMatchedForEventID(Cells(7, 2).Value, ThisEndPoint)
            Cells(5, 1).Value = CurrentTotalmatched


If Cells(3, 1) = "Soccer" Then
    Range(Cells(8, 1), Cells(SoccerLength, 1)).Select
Else
    Range(Cells(8, 1), Cells(AFLLength, 1)).Select
End If

Call ListMarketRunners(ThisEndPoint)

End Sub

Public Sub ZZBackupOld_GetEventsFromCompList(Optional Country As String)
'// Initiated from button click. Firstly, gets list of the compID's that are in the current selection
'// and then retrieves their name (also name of relevent output sheet). Passes compID's (one by one for now)
'// to ListCompEvents to get all the events for the relevant competition ID. THen responsible for populating the correct spreadsheet
'// with the list of the events for that competition

If Country <> "AUS" Then Country = "UK"
'sets default endpoint country to UK if AUS isn't specified

Dim Rng             As Range
Dim CompList        As Object
Dim CompName        As String
Dim OPSheet         As Worksheet
Dim CompEvents()    As SingleEvent
Dim CompIDS(0)      As String
Dim N               As Long
Dim IDCol           As Long
Dim NAMECol         As Long
Dim IDRow           As Long
Dim NAMERow         As Long
Dim RowTally        As Long

'// Attempt to get the comp list
Set CompList = GetCompList(Country)

If Not IsObject(CompList) Then MsgBox CompListFail: Exit Sub

For Each Rng In Selection   '// Loop through each selected cell (range)

    RowTally = 0 '// Rest is important...
    
    If IsNumeric(Trim(Rng.Value)) Then   '// So far, so good..
    
        CompName = Trim(FindCompNameFromID(CompList, Trim(Rng.Value)))
        
        '// Set our output worksheet to be the right one per the competition name. Error indicates
        '// it did not exist, therefore we will simply move on to the next one...
        On Error Resume Next
        Err.Clear
        Set OPSheet = Sheets(CompName)
        If Err.Number = 0 Then  '// Can carry on as worksheet of that name does, in fact, exist...
        
            Err.Clear
            On Error GoTo 0 '// Reset normal error handling operations
            
            CompIDS(0) = Trim(Rng.Value)
            
            If ListCompEvents(CompIDS(), CompEvents()) Then '// Managed to get some events back for this competitoion ID
            
                '// Get column (and row) refs for the eventID and names...
                IDCol = GetNamedRngColumn(eventid, CompName)
                NAMECol = GetNamedRngColumn(eventname, CompName)
                
                IDRow = GetNamedRngRow(eventid, CompName)
                NAMERow = GetNamedRngRow(eventname, CompName)
            
                '// Now can go through each of them and populate the relevant output spreadsheet...
                For N = LBound(CompEvents()) To UBound(CompEvents())

                    '// Need to filter out case where event name = compname...
                    If CompEvents(N).eventname <> CompName Then

                        OPSheet.Cells(IDRow + 1 + RowTally, IDCol).Value = CompEvents(N).eventid
                        OPSheet.Cells(NAMERow + 1 + RowTally, NAMECol).Value = CompEvents(N).eventname
                        OPSheet.Cells(NAMERow + 1 + RowTally, NAMECol + 1).Value = CompEvents(N).opendate
                        
                        RowTally = RowTally + 1
                    
                    End If

                Next N
            
            End If
        
        
        End If
        
    Else
    
        MsgBox "Cell value: " & Rng.Value & " is not numeric !", vbInformation, "Input Error"
    
    End If

Next Rng

End Sub

Public Function CancelOrdersForBetID(ByVal marketid As String, ByVal BetID As String) As String

On Error GoTo ErrorHandler:

'Fields are:
'MarketId
'SelectionID NOW BetID
'handicap = 0
'ordertype = LIMIT
'side = BACK/LAY
'size = size
'price = price

Dim CancelOrdersResult As Object
Dim EndPoint As String: EndPoint = Cells(4, 1).Value
Dim N As Integer
Dim Request As Variant

Dim orderPrice As Double

            If Len(marketid) <> 11 Then
                For N = 1 To 11 - Len(marketid)
                marketid = "" & marketid & "0" 'Had problems with trailing 0s missing
                Next N
            End If
            
    Request = MakeJsonRpcRequestString(CancelOrdersMethod, GetCancelOrdersMarketIDBetIDRequestString(marketid, BetID))
        
    Dim CancelOrdersResponse As String: CancelOrdersResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
    AppendToLogFile "Cancel Orders " & CancelOrdersResponse
    
    Set CancelOrdersResult = ParseJsonRpcResponseToCollection(CancelOrdersResponse)
    
    CancelOrdersForBetID = CancelOrdersResponse 'temporary
    



On Error GoTo 0
    Exit Function
        
ErrorHandler:
    HandleError "CancelOrders"
    Resume Next


End Function

Public Sub ListCurrentOrdersNew(ByVal MyAction As String)
'Options for MyAction are:

'List = List the orders on the sheet
'Cleanse = Remove orders that are NOT next in line for action
'Purge = Delete ALL orders on this Event (i.e. all for each MarketID)


'as it says, list the current orders for an event page
'check the price column for both Back and Lay
'if there is a price there then check to see if the price has been matched

On Error GoTo ErrorHandler:

Dim Request
Dim MyCell
Dim N As Integer

Dim MyLoop As Integer
Dim BetsOnThisMarket As Integer
Dim BacklayAmounts, BackLayLocation, BetIDDetails, BetCounterTally As Range
Dim MarketIDs As Range
Dim BackOrLay As Boolean
Dim ThisSelectionId As String
Dim MarketIDString As String: MarketIDString = ""
Dim EndPoint As String: EndPoint = Cells(4, 1).Value
Dim ListCurrentOrdersResult As Object

Dim marketid As String
Dim Side As String

                    If Cells(3, 1) = "Soccer" Then
                                Set MarketIDs = Range(Cells(8, 1), Cells(SoccerLength, 1))
                                    Set BackLayLocation = Range(Cells(8, 8), Cells(SoccerLength, 11))
                                        Set BetIDDetails = Range(Cells(8, 28), Cells(SoccerLength, 48))
                                            Set BetCounterTally = Range(Cells(8, 49), Cells(SoccerLength, 49))
                    Else
                                Set MarketIDs = Range(Cells(8, 1), Cells(AFLLength, 1))
                                    Set BackLayLocation = Range(Cells(8, 8), Cells(AFLLength, 11))
                                        Set BetIDDetails = Range(Cells(8, 28), Cells(AFLLength, 48))
                                            Set BetCounterTally = Range(Cells(8, 49), Cells(AFLLength, 49))
                    End If

        If MyAction = "List" Then
                        BackLayLocation.ClearContents
                        BetIDDetails.ClearContents
                        BetCounterTally.ClearContents
        End If
            
For Each MyCell In MarketIDs
'calculate the values in the array
    If MyCell.Value > 0 Then
        'do something since there's a value there
            'ThisSelectionId = Cells(MyCell.Row, 3).Value
            marketid = Cells(MyCell.Row, 1).Value
    
'    If marketID = "" Then
'        Cells(MyCell.Row, 1).Select
'        marketID = Selection.End(xlUp).Value
'    End If

            If Len(marketid) <> 11 Then
                For N = 1 To 11 - Len(marketid)
                marketid = "" & marketid & "0" 'Had problems with trailing 0s missing
                Next N
            End If
            
    'add marketID to MarketIDString
    If Len(MarketIDString) > 10 Then 'it's the 2nd and subsequent entry so add a comma
          MarketIDString = MarketIDString & ""","""
            
    End If
    
    MarketIDString = MarketIDString & marketid
    
End If
    
Next 'myCell
    
    Request = MakeJsonRpcRequestString(ListCurrentOrdersMethod, GetListCurrentOrdersRequestString(MarketIDString))
        Dim ListCurrentOrdersResponse As String: ListCurrentOrdersResponse = SendRequest(GetJsonRpcUrl(EndPoint), GetAppKey(), GetSession(), Request)
        AppendToLogFile "List Current Orders " & ListCurrentOrdersResponse
    
    Set ListCurrentOrdersResult = ParseJsonRpcResponseToCollection(ListCurrentOrdersResponse)
    
    'Cells(MyCell.Row, 28) = ListCurrentOrdersResponse 'temporary
    
BetsOnThisMarket = ListCurrentOrdersResult.Item("currentOrders").Count
    
    If BetsOnThisMarket > 0 Then
        
        'since we have some data, clear the OLD stuff
            
            
            Call OutputListCurrentOrdersNew(ListCurrentOrdersResult, MyAction, MarketIDString)
    End If





On Error GoTo 0
Exit Sub 'ListCurrentOrdersNew

ErrorHandler:
    HandleError "ListCurrentOrdersNew"
    Resume Next

End Sub
Sub ListCurrentOrdersNewFromButtonOrMenuList()

Call ListCurrentOrdersNew("List")

End Sub
Sub ListCurrentOrdersNewFromButtonOrMenuCleanse()

Call ListCurrentOrdersNew("Cleanse")

End Sub
Sub ListCurrentOrdersNewFromButtonOrMenuTopOfferOnly()

Call ListCurrentOrdersNew("TopOfferOnly")

End Sub
Sub ListCurrentOrdersNewFromButtonOrMenuPurge()

Call ListCurrentOrdersNew("Purge")

End Sub

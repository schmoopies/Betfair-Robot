Attribute VB_Name = "Creation"

Sub CreateNewSheetsFromMenu()

CreateNewSheets (False)

'CleanupOldSheets

'CreateNewSheets (True)

End Sub
Sub CleanupApplication()
Attribute CleanupApplication.VB_ProcData.VB_Invoke_Func = " \n14"

Application.Calculation = xlCalculationAutomatic
Application.StatusBar = False

End Sub
Sub CleanupOldSheets()

UnhideAll
'Added AND condition for Soccer delete only on 29th March 2018 when moved AFL to server - want to keep AFL games so that Surfaces can be created
For Each WS In Worksheets

If (Left(WS.Name, 1) = "2") And (WS.Visible = True) And (WS.Cells(6, 3) < Now() - 0.5 / 24) And (WS.Cells(3, 1) = "Soccer") Then

    CountSheets = CountSheets + 1
            Application.DisplayAlerts = False
                WS.Delete
            Application.DisplayAlerts = True
End If

Next WS



End Sub
Sub CreateNewSheets(ByVal Auto As Boolean)
'
' WorkbookStuff Macro
'
Dim ThisEventID As String
Dim ThisEventName As String
Dim ThisTime As Date
Dim FutureTime As Date
Dim ThisCompetitionId As Double
Dim ThisCompetitionName As String
Dim ThisCompetitionRegion As String
Dim ThisEndPoint As String
Dim LookAheadCreateWindow As Double
Dim TotalMatchedLimit As Double
Dim ThisWorksheetExists As Boolean
Dim ModEventID, ModValue As Double
Dim ValidEventInstances As Double

Dim ContraryBoolean As Boolean
Dim ContraryText As String
Dim GameTypes As String

'
    
Application.Calculation = xlCalculationManual
Application.StatusBar = "Creating New Sheets ..."

TotalMatchedLimit = Sheets("Example").Cells(GetNamedRngRow("TotalMatchedLimit", "Example"), GetNamedRngColumn("TotalMatchedLimit", "Example")).Value
    
Worksheets("Future List").Activate

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
                    CreatedCol = GetHeaderColumn("Created")
    
    
If Auto Then
'if called automatically then we select the entire COLUMN
        Worksheets("Future List").Select
        Worksheets("Future List").Activate
            Cells(1, 1).Select
                Set tbl = ActiveCell.CurrentRegion
                tbl.Offset(1, eventidCol - 1).Resize(tbl.rows.Count - 1, 1).Select
        
        LookAheadCreateWindow = 1.1 * Sheets("Example").Cells(GetNamedRngRow("LookAheadWindow", "Example"), GetNamedRngColumn("LookAheadWindow", "Example")).Value
Else
'if NOT then we work with the current selection AND run KeepAlive
'KeepAlive is run from other calls and we don't want ot overkill
    KeepAlive
    LookAheadCreateWindow = 148.5
End If

Worksheets("Future List").Activate

For Each ThisCell In Selection
    
ThisWorksheetExists = False

    ThisEventID = ThisCell.Value
        ThisTime = Cells(ThisCell.Row, LocalDateTimeCol).Value
            FutureTime = Now() + LookAheadCreateWindow / 24
                ModValue = Sheets("Example").Cells(GetNamedRngRow("ModValue", "Example"), GetNamedRngColumn("ModValue", "Example")).Value
                    ModEventID = ModEventIDRemainder(Val(ThisEventID), ModValue) + 1
                        ValidEventInstances = Sheets("Example").Cells(GetNamedRngRow("ValidEventInstances", "Example"), GetNamedRngColumn("ValidEventInstances", "Example")).Value
                            GameTypes = Sheets("Example").Cells(GetNamedRngRow("GameTypes", "Example"), GetNamedRngColumn("GameTypes", "Example")).Value
                        
  
  test2 = Trim(str(ModEventID))
  test = InStr(1, str(ValidEventInstances), Trim(str(ModEventID)), vbTextCompare)
  
  'need to put a short fcuntion here to GetOpenDateFromEventID and check that the
  'date/time hasn't moved before proceeding
  
  If (ThisTime > Now()) And _
    (ThisTime < FutureTime) And _
        (Cells(ThisCell.Row, CreatedCol).Value = "") And _
            InStr(1, str(ValidEventInstances), Trim(str(ModEventID)), vbTextCompare) <> 0 Then
            
            For Each WS In Worksheets
                If WS.Name = ThisEventID Then
                    ThisWorksheetExists = True
                End If
            Next WS
     
 'it's within the time range of interest so....
 
If Not ThisWorksheetExists Then
 
ThisCompetitionRegion = Cells(ThisCell.Row, countryCodeCol).Value
    
 ThisTotalMatched = GetTotalMatchedForEventID(ThisEventID, ThisCompetitionRegion)
 Cells(ThisCell.Row, totalMatchedCol).Value = ThisTotalMatched 'put it on the sheet
 
 If (ThisTotalMatched >= TotalMatchedLimit) Then 'continue
 
        ThisEventName = Cells(ThisCell.Row, eventnameCol).Value
        ThisCompetitionId = Cells(ThisCell.Row, competitionidCol).Value
        ThisCompetitionName = Cells(ThisCell.Row, competitionnameCol).Value
        ThisCompetitionRegion = Cells(ThisCell.Row, countryCodeCol).Value

        ContraryText = Cells(ThisCell.Row, 6).Value
        
 'MUST FIX To INCLUDE EVENTTYPE ID instead of COMPETITION ID 29-03-2018
 '10980856 11897406
          If (ThisCompetitionId = "11897406") And (InStr(1, GameTypes, "AFL", vbTextCompare) <> 0) Then
              Sheets("Calculations AFL").Select
              Sheets("Calculations AFL").Copy before:=Sheets("Calculations Soccer")
              Sheets("Calculations AFL (2)").Select
              Sheets("Calculations AFL (2)").Name = ThisEventID
              Sheets(ThisEventID).Select
                        Cells(7, 2).Value = ThisEventID
                        Cells(7, 3).Value = ThisEventName
                        Cells(6, 3).Value = ThisTime
                        Cells(4, 2).Value = ThisCompetitionId
                        Cells(4, 3).Value = ThisCompetitionName
                        Cells(5, 3).Value = ThisCompetitionRegion
                        Cells(4, 1).Value = ThisEndPoint
                        Cells(3, 2).Value = ContraryText
          'ActiveWindow.SelectedSheets.Visible = False
          'Range("A3").Select
          '11353530 NFL 2017
          '10522563 College Football Matches
          '10693361    NFL Season 2017/18


          
          Sheets("Future List").Select
          Cells(ThisCell.Row, CreatedCol).Value = "yes"
          'Cells(4, 2).Value = 10522052 Or Cells(4, 2).Value = 11432305
          ElseIf (ThisCompetitionId = "10522052" Or ThisCompetitionId = "11432305") And (InStr(1, GameTypes, "NFL", vbTextCompare) <> 0) Then
              Sheets("Calculations NFL").Select
              Sheets("Calculations NFL").Copy before:=Sheets("Calculations Soccer")
              Sheets("Calculations NFL (2)").Select
              Sheets("Calculations NFL (2)").Name = ThisEventID
              Sheets(ThisEventID).Select
                       Cells(7, 2).Value = ThisEventID
                        Cells(7, 3).Value = ThisEventName
                        Cells(6, 3).Value = ThisTime
                        Cells(4, 2).Value = ThisCompetitionId
                        Cells(4, 3).Value = ThisCompetitionName
                        Cells(5, 3).Value = ThisCompetitionRegion
                        Cells(4, 1).Value = ThisEndPoint
                        Cells(3, 2).Value = ContraryText
     
          'ActiveWindow.SelectedSheets.Visible = False
          'Range("A3").Select
          Sheets("Future List").Select
          Cells(ThisCell.Row, CreatedCol).Value = "yes"
          'below line is a workaroudn ONLY - FIX
          ElseIf (ThisCompetitionId <> "11897406") And InStr(1, GameTypes, "Soccer", vbTextCompare) <> 0 Then
              Sheets("Calculations Soccer").Select
              Sheets("Calculations Soccer").Copy before:=Sheets("Calculations Soccer")
              Sheets("Calculations Soccer (2)").Select
              Sheets("Calculations Soccer (2)").Name = ThisEventID
              Sheets(ThisEventID).Select
                       Cells(7, 2).Value = ThisEventID
                        Cells(7, 3).Value = ThisEventName
                        Cells(6, 3).Value = ThisTime
                        Cells(4, 2).Value = ThisCompetitionId
                        Cells(4, 3).Value = ThisCompetitionName
                        Cells(5, 3).Value = ThisCompetitionRegion
                        Cells(4, 1).Value = ThisEndPoint
                        Cells(3, 2).Value = ContraryText
     
          'ActiveWindow.SelectedSheets.Visible = False
          'Range("A3").Select
          Sheets("Future List").Select
          Cells(ThisCell.Row, CreatedCol).Value = "yes"
          End If
          
                  'If (ThisCompetitionRegion = "AU" Or ThisCompetitionRegion = "NZ") Then
                '   ThisEndPoint = "AUS"
                  'Else
                      ThisEndPoint = "UK"
                  'End If
                  
  ''        Cells(7, 2).Value = ThisEventID
    ''      Cells(7, 3).Value = ThisEventName
      ''    Cells(6, 3).Value = ThisTime
        ''  Cells(4, 2).Value = ThisCompetitionId
          ''Cells(4, 3).Value = ThisCompetitionName
          ''Cells(5, 3).Value = ThisCompetitionRegion
          ''Cells(4, 1).Value = ThisEndPoint
          
          
          
          ''If ThisCompetitionId = "8311792" Then
            ''  Sheets("Calculations AFL (2)").Select
              ''Sheets("Calculations AFL (2)").Name = ThisEventID
          ''ElseIf ThisCompetitionId = "8764204" Then
            ''  Sheets("Calculations NFL (2)").Select
              ''Sheets("Calculations NFL (2)").Name = ThisEventID
          ''Else
            ''  Sheets("Calculations Soccer (2)").Select
              ''Sheets("Calculations Soccer (2)").Name = ThisEventID
          ''End If
          
          ''Sheets(ThisEventID).Select
          'ActiveWindow.SelectedSheets.Visible = False
          'Range("A3").Select
          ''Sheets("Future List").Select
          ''Cells(ThisCell.Row, CreatedCol).Value = "yes"
    
        End If 'TotalMatched
    End If 'WorksheetExists
End If 'Time Checks
    
Next

Application.Calculation = xlCalculationAutomatic
Application.StatusBar = False

End Sub
Sub UpdateEventTimes()
'
' Take a selected LIST of EventIDs form the Future worksheet and update their
' scheduled start time - highlight in RED any that have changed during this process
'
Dim ThisEventID As String
Dim ThisEventName As String
Dim ThisTime As Date
Dim NewEventTime As String

Dim ThisCompetitionId As Double
Dim ThisCompetitionName As String
Dim ThisCompetitionRegion As String
Dim ThisEndPoint As String
'
Dim CountSheets As Integer
Dim SheetsToProcess As Integer
Dim NumberToProcess As Double
Dim CountEvents As Integer
   
CountEvents = 0
ProcessStartTime = Now()
NumberToProcess = Selection.rows.Count
   
   
Application.Calculation = xlCalculationManual
    
KeepAlive
    
For Each ThisCell In Selection
      
    CountEvents = CountEvents + 1
    
    ThisEventTime = Cells(ThisCell.Row, 10).Value
    ThisEventID = ThisCell.Value
    ThisCompetitionRegion = Cells(ThisCell.Row, 7).Value
        If ThisCompetitionRegion = "AU" Then ThisEndPoint = "UK" Else ThisEndPoint = "UK"
    NewEventTime = GetOpenDateFromEventID(ThisEventID, ThisEndPoint)
    
        If ThisEventTime <> NewEventTime Then
            Cells(ThisCell.Row, 10).Value = NewEventTime
            Cells(ThisCell.Row, 10).Interior.Color = 255
        End If
 
 
ElapsedTime = Now() - ProcessStartTime
AverageProcessTime = ElapsedTime / CountEvents
TimeRemaining = AverageProcessTime * (NumberToProcess - CountEvents)

Application.StatusBar = "Processed : " & CountEvents & " Remaining : " & (NumberToProcess - CountEvents) & " and Time Remaining = " & Format(TimeRemaining, "hh:mm:ss")
 
 
Next

Application.Calculation = xlCalculationAutomatic
Application.OnTime Now() + TimeSerial(0, 0, 3), "ClearStatusBar"

End Sub


Sub GetExpected()

Dim PutWhere As Long
Dim MyLocation As Worksheet

Set MyLocation = Worksheets("Future List")

For Each WS In Worksheets

If WorksheetFunction.IsNumber(Val(WS.Name)) Then

    PutWhere = FindTheValue(MyLocation, 8, WS.Name)
    If Not PutWhere = 0 Then 'it was found in my list
        PutItHere = PutTheResult(MyLocation, PutWhere, 17, WS.Cells(2, 20).Value, False) 'Expected
        PutItHere = PutTheResult(MyLocation, PutWhere, 18, WS.Cells(3, 20).Value, False) 'Max
        PutItHere = PutTheResult(MyLocation, PutWhere, 19, WS.Cells(4, 20).Value, False) 'Min
        
        PutItHere = PutTheResult(MyLocation, PutWhere, 28, WS.Cells(3, 5).Value, False) 'MeanHome
        PutItHere = PutTheResult(MyLocation, PutWhere, 29, WS.Cells(3, 7).Value, False) 'MeanAway
        PutItHere = PutTheResult(MyLocation, PutWhere, 30, WS.Cells(4, 9).Value, False) 'Mult
    End If
End If

Next


End Sub


Public Function FindTheValue(searchSheet As Worksheet, inColumn As Integer, thisFind)

With searchSheet.Columns(inColumn)
    Set c = .Find(thisFind, LookIn:=xlValues, LookAt:=xlWhole)
    If Not c Is Nothing Then
        FindTheValue = c.Row
    Else
        FindTheValue = 0
    End If
End With

End Function


Sub ShowCooper()

Dim CoopersNumber As Integer


For CoopersNumber = 1 To 2000

'do something

If CoopersNumber > 1000 Then
    Cells(CoopersNumber, 5) = "hello world"
Else
    Cells(CoopersNumber, 5) = "blah blah"
End If


Next

End Sub
Sub UpdateAllGames()

Dim CountSheets As Integer
Dim SheetsToProcess As Integer
Dim OKToBetWithSheet As Boolean
Dim TotalMatchedForEvent, TotalMatchedLimit As Double
Dim NumberToProcess As Integer
Dim ContraryText As String
Dim ContraryBoolean As Boolean

'KeepAlive
'AppendToLogFile "Update All Games"

If Sheets("Example").Cells(8, 2).Value = False Then
    Application.Visible = False
Else: Application.Visible = True
End If

'get the limit for betting so that we can determine enough money is flowing to this Event
TotalMatchedLimit = Sheets("Example").Cells(GetNamedRngRow("TotalMatchedLimit", "Example"), GetNamedRngColumn("TotalMatchedLimit", "Example")).Value

'set the possible timer repeat to be now() + an offset
Sheets("Example").Cells(5, 2).Value = Now() + (Sheets("Example").Cells(9, 2).Value) / 24


CountSheets = 0

        'Find out how many there are to process
        For Each WS In Worksheets
            'If (Left(WS.Name, 1) = "2") And (WS.Visible = True) Then
            If (Left(WS.Name, 1) = "2") Then
                CountSheets = CountSheets + 1
                WS.Visible = xlSheetVisible    'Make them ALL visible - the Sort Procedure will check Total Matched and hide the tail
                'Note - the timer procedure has already made ALL sheets visible - but this step left here in case called from Menu
            End If
        Next
        SheetsToProcess = CountSheets

'now start the process

Call TopOfferOnlyAllEvents 'CleanseAllEvents ' get rid of all orders that are using up funds
Call SortSheets 'sort the visible sheets into descending order of totalMatched

CountSheets = 0
ProcessStartTime = Now()
NumberToProcess = Sheets("Example").Cells(GetNamedRngRow("NumberToProcess", "Example"), GetNamedRngColumn("NumberToProcess", "Example")).Value


For Each WS In Worksheets

If (Left(WS.Name, 1) = "2") And (WS.Visible = True) And (WS.Cells(5, 1) > TotalMatchedLimit) Then

    CountSheets = CountSheets + 1

WS.Activate
'how much has been matched for this particular Event
TotalMatchedForEvent = Cells(5, 1).Value

ContraryText = Cells(3, 2).Value
    If ContraryText = "Contrary" Then
        ContraryBoolean = True
    Else 'any other value
        ContraryBoolean = False
    End If

If (CountSheets > 0 And CountSheets <= NumberToProcess And (CheckGameIsCurrent) And (TotalMatchedForEvent > TotalMatchedLimit)) Then 'set a maximum for now
    
        Call GenerateMarketIDList
        Call UpdateMarketPrices
        Call Soccer_Loop_Solver
        
            OKToBetWithSheet = CheckSheetConditionsToBet
            WantToBet = Sheets("Example").Cells(5, 5).Value
                    If (OKToBetWithSheet And WantToBet = "PlaceBets") Then
                        'here, we should CLEANSE bets that aren't at the top of the pile to give more money to bet i.e. those where I have been outbid
                            Call ListCurrentOrdersNew("TopOfferOnly")
                                Call PlaceOrdersAutomaticallyNEW(True, "Back", ContraryBoolean)
                                Call PlaceOrdersAutomaticallyNEW(True, "Lay", ContraryBoolean)
                    End If
        
        Call ListCurrentOrdersNew("List") 'list them immediately to see which have been matched
        
            If Cells(4, 2) = "11897406" Then
                Call RunAFLSurface
            ElseIf Cells(4, 2) = "10522052" Or Cells(4, 2) = "11432305" Then
                Call RunNFLSurface
            Else
                Call RunSoccerSurface
            End If
        
End If 'CountSheets

'Call TopOfferOnlyAllEvents ' get rid of all orders that are using up funds (i.e. aren't the Top Offer)

ElapsedTime = Now() - ProcessStartTime
AverageProcessTime = ElapsedTime / CountSheets
TimeRemaining = AverageProcessTime * (SheetsToProcess - CountSheets)

Application.StatusBar = "Processed : " & CountSheets & " Remaining : " & (SheetsToProcess - CountSheets) & " and Time Remaining = " & Format(TimeRemaining, "hh:mm:ss")

End If

Next


Call GetExpected

Application.OnTime Now + TimeSerial(0, 0, 10), "MyClearStatusBar"
AppendToLogFile "Finished Update All Games"

Application.Visible = False
Application.Visible = True

Worksheets("Example").Select
Worksheets("Example").Activate
Cells(1, 1).Select

WantToBet = Sheets("Example").Cells(5, 5).Value

If WantToBet = "PlaceBets" Then GetAccountFunds ("Mail") 'only gets the new account funds and send an email IF bets were placed automatically

'Empty the Clipboard too
    Application.CutCopyMode = False

'save progress in case this helps release memory
    ActiveWorkbook.Save
    
End Sub
Public Sub UnhideAll()

For Each WS In Worksheets

If Left(WS.Name, 1) = "2" Then
    WS.Visible = xlSheetVisible
    'ws.Visible = xlSheetHidden
End If

Next
End Sub
Public Sub AFL_PreSolver()

Dim AFLPreMax As Double

'Loop 3 times but only reset 1s on the first time

For MyLoop = 1 To 2 'Changed to 2 on 02-05-17 to see if that speeds up solution - don't need 3x pre-solve?
 
    If MyLoop = 1 Then
        Range(Cells(8, 22), Cells(AFLLength, 23)).Select
        Selection.Value = 1
    End If
    

If Cells(2, 26) > 0 Then 'must check before running solver
'Adjust ODDS score error using HOME and AWAY mean
SolverReset
SolverOk SetCell:="$Z$2", MaxMinVal:=2, ValueOf:=0, ByChange:= _
        "$E$3,$G$3", Engine:=1, EngineDesc:="GRG Nonlinear"

    SolverSolve userfinish:=True
End If

'Adjust MARGIN score error using HOME and AWAY std dev
'SolverReset
'SolverOk SetCell:="$Z$3", MaxMinVal:=2, ValueOf:=0, ByChange:= _
'        "$E$4,$G$4", Engine:=1, EngineDesc:="GRG Nonlinear"

 '   SolverSolve userfinish:=True

If Cells(4, 26) > 0 Then 'must check before running solver
'Adjust SPREAD score error using both HOME and AWAY mean
SolverReset
SolverOk SetCell:="$Z$4", MaxMinVal:=2, ValueOf:=0, ByChange:= _
        "$E$3,$G$3", Engine:=1, EngineDesc:="GRG Nonlinear"

    SolverSolve userfinish:=True
End If

If (Cells(5, 26) > 0 And Cells(5, 27) > 6) Then 'must check before running solver
            'Adjust TGS score error using Mean and StDev
            SolverReset
            SolverOk SetCell:="$Z$5", MaxMinVal:=2, ValueOf:=0, ByChange:= _
                    "$E$3:$E$4,$G$3:$G$4", Engine:=1, EngineDesc:="GRG Nonlinear"
            SolverAdd CellRef:="R4C9", Relation:=3, FormulaText:="0.5"
    
                SolverSolve userfinish:=True
End If

'Original Calculation - now removed MULT in above code
          'Adjust TGS score error using Mult
           ' SolverReset
           ' SolverOk SetCell:="$Z$5", MaxMinVal:=2, ValueOf:=0, ByChange:= _
           '         "$E$3:$E$4,$G$3:$G$4,$I$4", Engine:=1, EngineDesc:="GRG Nonlinear"
           ' SolverAdd CellRef:="R4C9", Relation:=3, FormulaText:="0.5"
   '
   '             SolverSolve userfinish:=True




            ''Adjust TGS score error using HOME and AWAY std dev
            'SolverReset
            'SolverOk SetCell:="$Z$5", MaxMinVal:=2, ValueOf:=0, ByChange:= _
            '        "$E$4,$G$4", Engine:=1, EngineDesc:="GRG Nonlinear"
           '
           '     SolverSolve userfinish:=True
    
    Range(Cells(8, 22), Cells(AFLLength, 23)).Select
    AFLPreMax = Cells(3, 24).Value
    
    For Each MyCell In Selection
        'On Error Resume Next
        If MyCell.Offset(0, 2).Value > AFLPreMax * 0.85 Then
        MyCell.Value = 0
        'Exit For
        End If
    Next

Next MyLoop


End Sub
Public Sub NFL_PreSolver()

Dim NFLPreMax As Double

'Loop 3 times but only reset 1s on the first time

For MyLoop = 1 To 2 'Changed to 2 on 02-05-17 to see if that speeds up solution - don't need 3x pre-solve?
 
    If MyLoop = 1 Then
        Range(Cells(8, 22), Cells(52, 23)).Select
        Selection.Value = 1
    End If
    


'Adjust Total Error using HOME and AWAY mean
SolverReset
SolverOk SetCell:="$X$2", MaxMinVal:=2, ValueOf:=0, ByChange:= _
        "$E$3,$G$3", Engine:=1, EngineDesc:="GRG Nonlinear"

        SolverAdd CellRef:="R3C5", Relation:=3, FormulaText:="10"
        SolverAdd CellRef:="R3C7", Relation:=3, FormulaText:="10"
        SolverAdd CellRef:="R4C5", Relation:=3, FormulaText:="1"
        SolverAdd CellRef:="R4C7", Relation:=3, FormulaText:="1"
        
    SolverSolve userfinish:=True

'Adjust MARGIN score error using HOME and AWAY std dev
'SolverReset
'SolverOk SetCell:="$Z$3", MaxMinVal:=2, ValueOf:=0, ByChange:= _
'        "$E$4,$G$4", Engine:=1, EngineDesc:="GRG Nonlinear"

 '   SolverSolve userfinish:=True

'Adjust Total Errpr using both HOME and AWAY StdDev
SolverReset
SolverOk SetCell:="$X$2", MaxMinVal:=2, ValueOf:=0, ByChange:= _
        "$E$4,$G$4", Engine:=1, EngineDesc:="GRG Nonlinear"

        
        'SolverAdd CellRef:="R3C5", Relation:=3, FormulaText:="10"
        'SolverAdd CellRef:="R3C7", Relation:=3, FormulaText:="10"
        SolverAdd CellRef:="R4C5", Relation:=3, FormulaText:="1"
        SolverAdd CellRef:="R4C7", Relation:=3, FormulaText:="1"

    SolverSolve userfinish:=True


            'Adjust TGS score error using Mult
            'SolverReset
            'SolverOk SetCell:="$Z$5", MaxMinVal:=2, ValueOf:=0, ByChange:= _
            '        "$E$3:$E$4,$G$3:$G$4,$I$4", Engine:=1, EngineDesc:="GRG Nonlinear"
            'SolverAdd CellRef:="R4C9", Relation:=3, FormulaText:="0.5"
    
   '             SolverSolve userfinish:=True


            ''Adjust TGS score error using HOME and AWAY mean
            'SolverReset
            'SolverOk SetCell:="$Z$5", MaxMinVal:=2, ValueOf:=0, ByChange:= _
            '        "$E$3,$G$3", Engine:=1, EngineDesc:="GRG Nonlinear"
           '
           '     SolverSolve userfinish:=True




            ''Adjust TGS score error using HOME and AWAY std dev
            'SolverReset
            'SolverOk SetCell:="$Z$5", MaxMinVal:=2, ValueOf:=0, ByChange:= _
            '        "$E$4,$G$4", Engine:=1, EngineDesc:="GRG Nonlinear"
           '
           '     SolverSolve userfinish:=True
    
    Range(Cells(8, 22), Cells(52, 23)).Select
    NFLPreMax = Cells(3, 24).Value
    
    For Each MyCell In Selection
        'On Error Resume Next
        If MyCell.Offset(0, 2).Value > NFLPreMax * 0.85 Then
        MyCell.Value = 0
        'Exit For
        End If
    Next

Next MyLoop


End Sub
Public Sub Soccer_Loop_Solver()
'
' Loop_Solver Macro
'
'check start conditions
If Cells(5, 15) > 2 Or Cells(5, 14).Value > 1.4 Then
    Exit Sub
Else


Application.Calculation = xlCalculationManual


    removals = Cells(1, 1).Value
    
If Cells(4, 2).Value = 11897406 Then 'it's an AFL Match
    AFL_PreSolver
    'Exit Sub
    'Range(Cells(8, 22), Cells(50, 23)).Select 'temp
    'Selection.Value = 1 'temp
End If
    
If Cells(4, 2).Value = 10522052 Or Cells(4, 2).Value = 11432305 Then 'it's an NFL Match
    NFL_PreSolver
    'Exit Sub
    'Range(Cells(8, 22), Cells(50, 23)).Select 'temp
    'Selection.Value = 1 'temp
End If

If Cells(4, 2).Value = 11369949 Then 'it's a Basketball Match
    NFL_PreSolver
    'Exit Sub
    'Range(Cells(8, 22), Cells(50, 23)).Select 'temp
    'Selection.Value = 1 'temp
End If

If Cells(3, 1).Value = "Soccer" Then 'it's NOT an AFL or NFL Match, it's soccer
    'sooo we have to reset all the 1s because this was done in the pre_solver
    Range(Cells(8, 22), Cells(SoccerLength, 23)).Select
    Selection.Value = 1
End If
    
    Cells(3, 9).Value = "" 'clear this so that RESET is empty
    Cells(3, 10).Value = "" 'clear this so that RESET COUNT is empty
    
    For Count = 1 To removals
    
    Cells(1, 2) = Count
    


    SolverReset
    
 If Cells(4, 2).Value = 10522052 Or Cells(4, 2).Value = 11432305 Or Cells(4, 2).Value = 11369949 Then 'its NFL or Basketball
 
    SolverOk SetCell:="$Z$2", MaxMinVal:=2, ValueOf:=0, ByChange:= _
        "$E$3:$E$4,$G$3:$G$4", Engine:=1, EngineDesc:="GRG Nonlinear"
        '"$E$3:$E$4,$G$3:$G$4,$I$4", Engine:=1, EngineDesc:="GRG Nonlinear"
    
    'SolverAdd CellRef:="R4C9", Relation:=3, FormulaText:="0.02"
    SolverAdd CellRef:="R3C5", Relation:=3, FormulaText:="10"
    SolverAdd CellRef:="R3C7", Relation:=3, FormulaText:="10"
    SolverAdd CellRef:="R4C5", Relation:=3, FormulaText:="1"
    SolverAdd CellRef:="R4C7", Relation:=3, FormulaText:="1"
    
 Else
    SolverOk SetCell:="$X$2", MaxMinVal:=2, ValueOf:=0, ByChange:= _
        "$E$3:$E$4,$G$3:$G$4,$I$4", Engine:=1, EngineDesc:="GRG Nonlinear"
    
    SolverAdd CellRef:="R4C9", Relation:=3, FormulaText:="0.01"
    SolverAdd CellRef:="R3C5", Relation:=3, FormulaText:=".2"
    SolverAdd CellRef:="R3C7", Relation:=3, FormulaText:=".2"
 End If
        
    SolverSolve userfinish:=True
    
'    If (Cells(4, 9).Value <= 0.1) Or (Cells(4, 9).Value >= 1000) Then
'        Cells(4, 9).Value = 10
'        Cells(3, 9).Value = "Reset"
'    'this checks progress
'        Application.Calculation = xlCalculationAutomatic
'        Application.Calculation = xlCalculationManual
'    End If
    
    
    
'Reset Process
If Cells(4, 2).Value = 10522052 Or Cells(4, 2).Value = 11432305 Then 'it's NFL
    If (Cells(3, 5).Value >= 250) Or (Cells(3, 7).Value >= 250) Or _
            (Cells(4, 5) < 0.2) Or (Cells(4, 7) < 0.2) Then 'Or _
                '(Cells(3, 5) < 0.5 * Cells(4, 5)) Or (Cells(3, 7) < 0.5 * Cells(4, 7)) Then
                
                Cells(3, 5).Value = 24 'MeanHome
                Cells(3, 7).Value = 20 'MeanAway
                Cells(4, 5).Value = 13 'StdHome
                Cells(4, 7).Value = 13 'StdAway
                Cells(3, 9).Value = "Reset"
                Cells(3, 10).Value = Cells(3, 10).Value + 1
                
    'this checks progress
        Application.Calculation = xlCalculationAutomatic
        Application.Calculation = xlCalculationManual
    End If
ElseIf Cells(4, 2).Value = 11897406 Then 'it's AFL
    If (Cells(3, 5).Value >= 250) Or (Cells(3, 7).Value >= 250) Or _
            (Cells(4, 5) < 10) Or (Cells(4, 7) < 10) Or _
                (Cells(3, 5) < Cells(4, 5)) Or (Cells(3, 7) < Cells(4, 7)) Then
                
                Cells(3, 5).Value = 93 'MeanHome
                Cells(3, 7).Value = 89 'MeanAway
                Cells(4, 5).Value = 22 'StdHome
                Cells(4, 7).Value = 22 'StdAway
                Cells(3, 9).Value = "Reset"
                Cells(3, 10).Value = Cells(3, 10).Value + 1
                
    'this checks progress
        Application.Calculation = xlCalculationAutomatic
        Application.Calculation = xlCalculationManual
    End If
    
ElseIf Cells(4, 2).Value = 11369949 Then 'it's Basketball
    If (Cells(3, 5).Value >= 250) Or (Cells(3, 7).Value >= 250) Or _
            (Cells(4, 5) < 5) Or (Cells(4, 7) < 5) Or _
                (Cells(3, 5) < Cells(4, 5)) Or (Cells(3, 7) < Cells(4, 7)) Then
                
                Cells(3, 5).Value = 93 'MeanHome
                Cells(3, 7).Value = 89 'MeanAway
                Cells(4, 5).Value = 10 'StdHome
                Cells(4, 7).Value = 10 'StdAway
                Cells(3, 9).Value = "Reset"
                Cells(3, 10).Value = Cells(3, 10).Value + 1
                
    'this checks progress
        Application.Calculation = xlCalculationAutomatic
        Application.Calculation = xlCalculationManual
    End If
    
Else 'ít's a soccer match
    If (Cells(3, 5).Value >= 10) Or (Cells(3, 7).Value >= 10) Then
        Cells(3, 5).Value = 1.6 'MeanHome
        Cells(3, 7).Value = 1.1 'MeanAway
        Cells(4, 9).Value = 0.07 'Mult
        Cells(3, 9).Value = "Reset"
    'this checks progress
        Application.Calculation = xlCalculationAutomatic
        Application.Calculation = xlCalculationManual
    End If


End If
    
        Application.Calculation = xlCalculationAutomatic
        Application.Calculation = xlCalculationManual
    
    Range(Cells(8, 22), Cells(SoccerLength, 23)).Select
    thisMax = Cells(3, 24).Value
    
    For Each MyCell In Selection
        'On Error Resume Next
        If MyCell.Offset(0, 2).Value > thisMax * 0.85 Then
        MyCell.Value = 0
        'Exit For
        End If
    Next
    
 Call AddSolverProgressToLogSpreadsheet(Cells(3, 5).Value, Cells(3, 7).Value, Cells(4, 14).Value, Cells(2, 24).Value, Cells(4, 9).Value)
    
    'On Error GoTo 0
    
    Next Count

CheckSheetConditionsToBet

Application.Calculation = xlCalculationAutomatic

End If

End Sub

Sub MyClearStatusBar()
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
End Sub
Sub StartTimer()
    
    CurrentRunWhen = Sheets("Example").Cells(5, 2).Value 'this was placed there at the start of the last run
    PotentialRunWhen = Now() + TimeSerial(0, cRunIntervalMinutes, 0) 'this is just an offset from now
    
    RunWhen = WorksheetFunction.Max(CurrentRunWhen, PotentialRunWhen)
    Application.OnTime EarliestTime:=RunWhen, Procedure:=cRunWhat, _
        Schedule:=True
        
    PutItHere = PutTheResult(Sheets("Example"), 5, 2, RunWhen, True)
    PutItHere = PutTheResult(Sheets("Example"), 5, 3, "ACTIVE", True)
    
End Sub
Sub StopTimer()
    
    RunWhen = Sheets("Example").Cells(5, 2).Value
    
    On Error Resume Next
    Application.OnTime EarliestTime:=RunWhen, Procedure:=cRunWhat, _
        Schedule:=False
        
    PutItHere = PutTheResult(Sheets("Example"), 5, 3, "INACTIVE", True)
        
End Sub
Sub UpdateAllGamesUsingTimer()

Dim myCount As Integer: myCount = 0

Dim DummyOPSheet         As Worksheet

ThisWorkbook.Activate

Set DummyOPSheet = Sheets("Example")

PutItHere = PutTheResult(DummyOPSheet, 5, 2, RunWhen, True)

    Call KeepAlive 'this will Keepalive AND login if happens to be logged out
    
    Call UnhideAll 'need to have them visible to delete them
    Call CleanupOldSheets 'i.e. delete them if they are more than 0.5 hours old
    
    Call Build_Future_List
    
    Call CreateNewSheets(True)
    
    Call HideSheetsOutsideTimeWindow(Sheets("Example").Cells(7, 2).Value) 'this also unhides all sheets at the start
    Call UpdateAllGames
    'Application.StatusBar = "Running " & Now()

    'If TimeSerial(Hour(Now()), Minute(Now()), 0) < TimeSerial(22, 0, 0) Then
    'If ((TimeSerial(Hour(Now()), Minute(Now()), 0) < (TimeSerial(12, 0, 0))) And ((TimeSerial(Hour(Now()), Minute(Now()), 0) > TimeSerial(7, 0, 0)))) Then
    If True Then
    'If MyString = "Yes" Then
        Call StartTimer
    Else
        Call StopTimer
    End If

End Sub
Function CheckGameIsCurrent()

'verify game is still in the future
If Cells(6, 3).Value > Now() + (2 / (60 * 24)) Then 'make sure we have 2 minutes before game start
    CheckGameIsCurrent = True
Else
    CheckGameIsCurrent = False
End If
    

End Function

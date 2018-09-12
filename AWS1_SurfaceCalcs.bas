Attribute VB_Name = "SurfaceCalcs"
Sub ArrayTest(GameType As String, MaxScore As Integer, SheetOutput As Boolean)

Dim SurfaceArray(0 To 200, 0 To 200, 0 To 1) As Variant 'dim for 200 but only use 6 or 180 of the values
Dim Home, Away, Details As Integer
Dim MarketType As String
Dim Left4MarketType As String
Dim RunnerInfo As String
Dim TeamType As String
Dim Direction As String


Dim HomeAvg As Double: HomeAvg = Cells(3, 5).Value
Dim AwayAvg As Double: AwayAvg = Cells(3, 7).Value
'Dim Mult As Double: If GameType = "Soccer" Then Mult = Cells(4, 9).Value
Dim HomeStDev As Double: HomeStDev = Cells(4, 5).Value
Dim AwayStDev As Double: AwayStDev = Cells(4, 7).Value
Dim StDevHA As Double: StDevHA = Cells(4, 14).Value
Dim Mult As Double: Mult = Cells(4, 9).Value


Dim CurrentMax As Double:           CurrentMax = -999 'a small number
Dim CurrentMin As Double:           CurrentMin = 999 'a large number
Dim CurrentExpectation As Double:   CurrentExpectation = 0 'nothing
Dim CurrentTotalProb As Double

Dim ThisResult As String
Dim ThisOverUnderType As String
Dim ThisBothToScore As String
Dim ThisUnderOver As String
Dim ThisSpreadTeam As String
Dim ThisSpreadValue As Double
Dim ThisUnderOverValue As Double
Dim ThisGameScoreValue As Integer
Dim ThisWinningMarginSpread As Double
Dim ThisTriBetValue As Double
Dim ThisTriBetType As String

Dim ThisWMType As String 'WM = Winning Margin
Dim ThisWMSType As String 'WMS = Winning Margin Spread
Dim ThisGoals As Double
Dim ThisWMValue As Double
Dim ThisHomeScore As Integer
Dim ThisAwayScore As Integer

Dim Price As Double
Dim Amount As Double

Dim BacklayAmounts As Range
Dim BackOrLay As Boolean
Dim OutputSheet As String
Dim myCounter As Integer: myCounter = 0

Dim ThisMarketID, ThisGameName As String

OutputSheet = GameType & " Surface"
ThisMarketID = Cells(7, 2).Value
ThisGameName = Cells(7, 3).Value

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
If MyCell.Value > 0.02 Then
'do something since there's a value there
    MarketType = Cells(MyCell.Row, 2)
    Left4MarketType = Left(Cells(MyCell.Row, 2), 4)
    RunnerInfo = Cells(MyCell.Row, 4).Value
    Price = Cells(MyCell.Row, 8 + MyLoop).Value
    Amount = Cells(MyCell.Row, 9 + MyLoop).Value
    
If Amount > 0 Then

myCounter = myCounter + 1

    If Left4MarketType = "Regu" Then 'Regu is only used in NFL
    myCounter = myCounter + 1

        ThisResult = Cells(MyCell.Row, 4).Value
        Call AddToSurfaceMatchOdds(SurfaceArray, ThisResult, BackOrLay, Price, Amount, MaxScore) 'Match Odds is the same even though called different
    
    ElseIf Left4MarketType = "Mone" Then 'Moneyline is only used in NFL but can Reuse MatchOdds again
    myCounter = myCounter + 1

        ThisResult = Cells(MyCell.Row, 4).Value
        Call AddToSurfaceMatchOdds(SurfaceArray, ThisResult, BackOrLay, Price, Amount, MaxScore)
        
    ElseIf Left4MarketType = "Matc" Then 'Matc is the same for AFL and Soccer
    myCounter = myCounter + 1

        ThisResult = Cells(MyCell.Row, 4).Value
        Call AddToSurfaceMatchOdds(SurfaceArray, ThisResult, BackOrLay, Price, Amount, MaxScore)
        
    ElseIf Left4MarketType = "Over" Then 'Over only occurs in Soccer
    myCounter = myCounter + 1

        ThisOverUnderType = Cells(MyCell.Row, 4).Value
        ThisGoals = Cells(MyCell.Row, 5).Value
        Call AddToSurfaceOverUnder(SurfaceArray, ThisOverUnderType, ThisGoals, BackOrLay, Price, Amount, MaxScore)
        
    ElseIf Left4MarketType = "Corr" Then 'Correct Score is a soccer type
        ThisHomeScore = Cells(MyCell.Row, 4).Value
        ThisAwayScore = Cells(MyCell.Row, 5).Value
        Call AddToSurfaceCorrectScore(SurfaceArray, ThisHomeScore, ThisAwayScore, BackOrLay, Price, Amount, MaxScore)
        
'    ElseIf Left4MarketType = "Any " Then 'Correct Score Any Other is a soccer type
'        ThisAnyOther = Cells(MyCell.Row, 5).Value
'        Call AddToSurfaceAnyOther(SurfaceArray, ThisAnyOther, BackOrLay, Price, Amount, MaxScore)
        
    ElseIf Left4MarketType = "Both" Then 'Both to Score is a soccer only type
    myCounter = myCounter + 1

        ThisBothToScore = Cells(MyCell.Row, 4).Value
        Call AddToSurfaceBothToScore(SurfaceArray, ThisBothToScore, BackOrLay, Price, Amount, MaxScore)
        
    ElseIf MarketType = "Winning Margin Spread" Then 'it's an AFL Total Game Score type
    myCounter = myCounter + 1

        ThisWinningMarginSpread = Cells(MyCell.Row, 5).Value 'size fo the spread
        ThisWMSType = Left(Cells(MyCell.Row, 4).Value, 4) 'will be Home, Away or Draw
        Call AddToSurfaceWinningMarginSpread(SurfaceArray, ThisWinningMarginSpread, ThisWMSType, BackOrLay, Price, Amount, MaxScore)
    
    ElseIf MarketType = "Winning Margin 2" Then 'it's an NFL Winning Margin 2
    myCounter = myCounter + 1

        ThisWinningMarginSpread = Cells(MyCell.Row, 5).Value 'size fo the spread
        ThisWMSType = Left(Cells(MyCell.Row, 4).Value, 4) 'will be Home, Away or Draw or Tie
        Call AddToSurfaceWinningMarginNFL(SurfaceArray, ThisWinningMarginSpread, ThisWMSType, BackOrLay, Price, Amount, MaxScore)
    
    ElseIf MarketType = "Winning Margin" Then 'it's an NFL Winning Margin
    myCounter = myCounter + 1

        ThisWinningMarginSpread = Cells(MyCell.Row, 5).Value 'size fo the spread
        ThisWMSType = Left(Cells(MyCell.Row, 4).Value, 4) 'will be Home, Away or Draw or Tie
        Call AddToSurfaceWinningMarginNFL(SurfaceArray, ThisWinningMarginSpread, ThisWMSType, BackOrLay, Price, Amount, MaxScore)
    
    ElseIf Left4MarketType = "Winn" Then 'Winning Margin is AFL
    myCounter = myCounter + 1

        ThisWMType = Cells(MyCell.Row, 4).Value 'gives home or away and over or under
        ThisWMValue = Cells(MyCell.Row, 5).Value 'give the value either 24.5 or 39.5
        Call AddToSurfaceWinningMargin(SurfaceArray, ThisWMType, ThisWMValue, BackOrLay, Price, Amount, MaxScore)
        
    ElseIf MarketType = "Tri Bet" Then 'Tri Bet is AFL
    myCounter = myCounter + 1

        ThisTriBetType = Cells(MyCell.Row, 4).Value 'gives home or away and over or under
        ThisTriBetValue = Cells(MyCell.Row, 5).Value 'give the value either 24.5 or 39.5
        Call AddToSurfaceTriBet(SurfaceArray, ThisTriBetType, ThisTriBetValue, BackOrLay, Price, Amount, MaxScore)
        
    ElseIf Left4MarketType = "Unde" Then 'Under/Over is an AFL type
    myCounter = myCounter + 1

        ThisUnderOver = Cells(MyCell.Row, 4).Value 'either Under or Over
        ThisUnderOverValue = Cells(MyCell.Row, 5).Value 'give the value of Under/Over points
        Call AddToSurfaceUnderOver(SurfaceArray, ThisUnderOver, ThisUnderOverValue, BackOrLay, Price, Amount, MaxScore)
    
    ElseIf ((Right(MarketType, 3) = "pts") Or (Mid(MarketType, Len(MarketType) - 3, 1) = "+")) Then 'it's and AFL or NFL Spread type
    myCounter = myCounter + 1

        ThisSpreadTeam = Cells(MyCell.Row, 4).Value 'either Home or Away
        ThisSpreadValue = Cells(MyCell.Row, 5).Value 'size fo the spread
        Call AddToSurfaceSpread(SurfaceArray, ThisSpreadTeam, ThisSpreadValue, BackOrLay, Price, Amount, MaxScore)
        
    ElseIf MarketType = "Total Game Score" Then 'it's an AFL Total Game Score type
    myCounter = myCounter + 1

        ThisGameScoreValue = Cells(MyCell.Row, 5).Value 'size fo the spread
        Call AddToSurfaceTotalGameScore(SurfaceArray, ThisGameScoreValue, BackOrLay, Price, Amount, MaxScore)
    
    
    
    ElseIf Right(MarketType, 12) = "Total Points" Then 'it's a Total Points type only in NFL
    
    myCounter = myCounter + 1

        
        ThisTotalPointsValue = Cells(MyCell.Row, 5).Value 'value of the Total Points
        Direction = Cells(MyCell.Row, 4).Value
        'If Left(MarketType, 6) = Left(Cells(7, 3), 6) Then TeamType = "Home" Else TeamType = "Away"
       '
       ' If Left(RunnerInfo, 2) = "60" Then
       '             Direction = "Under"
       '     ElseIf Left(RunnerInfo, 2) = "12" Then
       '             Direction = "Over"
       '     Else: Direction = "Within"
       ' End If
        
        Call AddToSurfaceTotalPoints(SurfaceArray, ThisTotalPointsValue, Direction, BackOrLay, Price, Amount, MaxScore)
    
    
    ElseIf MarketType = "Handicap" Then 'it's an NFL Jandicap with is idnetical to AFL Spread type
    myCounter = myCounter + 1

        ThisSpreadTeam = Cells(MyCell.Row, 4).Value 'either Home or Away
        ThisSpreadValue = Cells(MyCell.Row, 5).Value 'size fo the spread
        Call AddToSurfaceSpread(SurfaceArray, ThisSpreadTeam, ThisSpreadValue, BackOrLay, Price, Amount, MaxScore)
    
    
    End If
    
End If 'Amount
    
End If
Next MyCell

Next MyLoop


'Now that it's been calculated, go through the array and update the max, min and expected value


'Soccer Calcs Below
If GameType = "Soccer" Then

        For Home = 0 To MaxScore
        For Away = 0 To MaxScore
'                            SurfaceArray(Home, Away, 0) = WorksheetFunction.Poisson_Dist(Home, HomeAvg, False) _
'                                * WorksheetFunction.Poisson_Dist(Away, AwayAvg, False) / Exp(Abs(Home - Away) / Mult) / CurrentTotalProb '0 contains the probability
            SurfaceArray(Home, Away, 0) = pbivpois(Home, Away, Mult, HomeAvg, AwayAvg)
            CurrentExpectation = CurrentExpectation + (SurfaceArray(Home, Away, 0) * SurfaceArray(Home, Away, 1)) '1 contains the value
            CurrentMax = Application.Max(SurfaceArray(Home, Away, 1), CurrentMax)
            CurrentMin = Application.Min(SurfaceArray(Home, Away, 1), CurrentMin)
        Next Away
        Next Home

Cells(2, 20) = CurrentExpectation
Cells(3, 20) = CurrentMax
Cells(4, 20) = CurrentMin

Offset = 0



        If SheetOutput = True Then
        
                For Home = 0 To MaxScore
                For Away = 0 To MaxScore
                    Sheets(OutputSheet).Cells(4 + Home, 2 + Away) = SurfaceArray(Away, Home, 1)
                    Sheets(OutputSheet).Cells(15 + Home, 2 + Away) = SurfaceArray(Away, Home, 0)
                Next Away
                Next Home
        
        End If

'AFL Calcs Below
ElseIf GameType = "AFL" Then

        For Home = 0 To MaxScore
        For Away = 0 To MaxScore
            SurfaceArray(Home, Away, 0) = (myKurtosisCopy(Home + 0.5, HomeAvg, StDevHA, True, 1) - myKurtosisCopy(Home - 0.5, HomeAvg, StDevHA, True, 1)) _
                * (myKurtosisCopy(Away + 0.5, AwayAvg, StDevHA, True, 1) - myKurtosisCopy(Away - 0.5, AwayAvg, StDevHA, True, 1)) '0 contains the probability
            CurrentExpectation = CurrentExpectation + (SurfaceArray(Home, Away, 0) * SurfaceArray(Home, Away, 1)) '1 contains the value
            CurrentMax = Application.Max(SurfaceArray(Home, Away, 1), CurrentMax)
            CurrentMin = Application.Min(SurfaceArray(Home, Away, 1), CurrentMin)
            
        Next Away
        Next Home




Cells(2, 20) = CurrentExpectation
Cells(3, 20) = CurrentMax
Cells(4, 20) = CurrentMin

Offset = 0

        If SheetOutput = True Then
        
                For Home = 50 - Offset To MaxScore - 50 + Offset
                For Away = 50 - Offset To MaxScore - 50 + Offset
                    Sheets(OutputSheet).Cells(4 + Home - 50 + Offset, 2 + Away - 50 + Offset) = SurfaceArray(Away, Home, 1)
                    'Sheets(OutputSheet).Cells(15 + Home - 50, 2 + Away - 50) = SurfaceArray(Away, Home, 0)
                Next Away
                Next Home
        
                Sheets(OutputSheet).Cells(2, 83).Value = ThisMarketID
                Sheets(OutputSheet).Cells(2, 84).Value = ThisGameName
        End If

ElseIf GameType = "NFL" Then

        For Home = 0 To MaxScore
        For Away = 0 To MaxScore
            SurfaceArray(Home, Away, 0) = (WorksheetFunction.Norm_Dist(Home + 0.5, HomeAvg, StDevHA * Mult, True) - WorksheetFunction.Norm_Dist(Home - 0.5, HomeAvg, StDevHA * Mult, True)) _
                * (WorksheetFunction.Norm_Dist(Away + 0.5, AwayAvg, StDevHA * Mult, True) - WorksheetFunction.Norm_Dist(Away - 0.5, AwayAvg, StDevHA * Mult, True)) '0 contains the probability
            CurrentExpectation = CurrentExpectation + (SurfaceArray(Home, Away, 0) * SurfaceArray(Home, Away, 1)) '1 contains the value
            CurrentMax = Application.Max(SurfaceArray(Home, Away, 1), CurrentMax)
            CurrentMin = Application.Min(SurfaceArray(Home, Away, 1), CurrentMin)
            
        Next Away
        Next Home




Cells(2, 20) = CurrentExpectation
Cells(3, 20) = CurrentMax
Cells(4, 20) = CurrentMin

Offset = 0

        If SheetOutput = True Then
        
                For Home = 0 To MaxScore
                For Away = 0 To MaxScore
                    Sheets(OutputSheet).Cells(4 + Home, 2 + Away) = SurfaceArray(Away, Home, 1)
                    'Sheets(OutputSheet).Cells(15 + Home - 50, 2 + Away - 50) = SurfaceArray(Away, Home, 0)
                Next Away
                Next Home
        
                Sheets(OutputSheet).Cells(2, 83).Value = ThisMarketID
                Sheets(OutputSheet).Cells(2, 84).Value = ThisGameName
        End If
End If

End Sub

Sub AddToSurfaceMatchOdds(SurfaceArray, Result As String, BackOrLay As Boolean, Price As Double, Amount As Double, MaxScore As Integer)

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If Result = "Draw" Then
        If (Home = Away) Then 'it's a Draw and I win
            If BackOrLay = True Then 'it's a BACK
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
            ElseIf BackOrLay = False Then
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
            End If
        Else 'Score doesn't match
            If BackOrLay = True Then 'it's a BACK
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
            ElseIf BackOrLay = False Then
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
            End If
        End If
    ElseIf Result = "Tie" Then 'Tie is only used in NFL -
        If (Home = Away) Then 'it's a Draw and I win
            If BackOrLay = True Then 'it's a BACK
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
            ElseIf BackOrLay = False Then
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
            End If
        Else 'Score doesn't match
            If BackOrLay = True Then 'it's a BACK
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
            ElseIf BackOrLay = False Then
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
            End If
        End If
    ElseIf Result = "Home" Then
        If (Home > Away) Then 'it's a Home win and I win
            If BackOrLay = True Then 'it's a BACK
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
            ElseIf BackOrLay = False Then
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
            End If
        Else 'Score doesn't match
            If BackOrLay = True Then 'it's a BACK
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
            ElseIf BackOrLay = False Then
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
            End If
        End If
    ElseIf Result = "Away" Then
        If (Home < Away) Then 'it's a Home win and I win
            If BackOrLay = True Then 'it's a BACK
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
            ElseIf BackOrLay = False Then
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
            End If
        Else 'Score doesn't match
            If BackOrLay = True Then 'it's a BACK
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
            ElseIf BackOrLay = False Then
                SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
            End If
        End If
    End If
Next Away
Next Home

End Sub
Sub AddToSurfaceOverUnder(SurfaceArray, OverUnderType As String, Goals As Double, BackOrLay As Boolean, Price As Double, Amount As Double, MaxScore As Integer)

If OverUnderType = "Under" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home + Away < Goals) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

ElseIf OverUnderType = "Over" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home + Away > Goals) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

End If
End Sub
Sub AddToSurfaceUnderOver(SurfaceArray, OverUnderType As String, Score As Double, BackOrLay As Boolean, Price As Double, Amount As Double, MaxScore As Integer)

If OverUnderType = "Under" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home + Away < Score) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

ElseIf OverUnderType = "Over" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home + Away > Score) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

End If
End Sub

Sub AddToSurfaceCorrectScore(SurfaceArray, HomeScore As Integer, AwayScore As Integer, BackOrLay As Boolean, Price As Double, Amount As Double, MaxScore As Integer)

If HomeScore < 99 Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home = HomeScore And Away = AwayScore) Then 'Score Matches
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home


ElseIf HomeScore = 99 Then 'SpecialCase Any Other Home Win

        For Home = 0 To MaxScore
        For Away = 0 To MaxScore
            If (((Home > 3) Or (Away > 3)) And (Home > Away)) Then 'it's ANY OTHER Home Win
                If BackOrLay = True Then 'it's a BACK
                    SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
                ElseIf BackOrLay = False Then
                    SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
                End If
            Else 'Score doesn't match
                If BackOrLay = True Then 'it's a BACK
                    SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
                ElseIf BackOrLay = False Then
                    SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
                End If
            End If
        Next Away
        Next Home

ElseIf HomeScore = 100 Then 'Any Other Away Win

        For Home = 0 To MaxScore
        For Away = 0 To MaxScore
            If (((Home > 3) Or (Away > 3)) And (Home < Away)) Then 'it's ANY OTHER Away Win
                If BackOrLay = True Then 'it's a BACK
                    SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
                ElseIf BackOrLay = False Then
                    SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
                End If
            Else 'Score doesn't match
                If BackOrLay = True Then 'it's a BACK
                    SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
                ElseIf BackOrLay = False Then
                    SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
                End If
            End If
        Next Away
        Next Home

ElseIf HomeScore = 101 Then

        For Home = 0 To MaxScore
        For Away = 0 To MaxScore
            If (((Home > 3) Or (Away > 3)) And (Home = Away)) Then 'it's ANY OTHER Draw
                If BackOrLay = True Then 'it's a BACK
                    SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
                ElseIf BackOrLay = False Then
                    SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
                End If
            Else 'Score doesn't match
                If BackOrLay = True Then 'it's a BACK
                    SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
                ElseIf BackOrLay = False Then
                    SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
                End If
            End If
        Next Away
        Next Home

End If

End Sub
Sub AddToSurfaceBothToScore(SurfaceArray, BothToScore As String, BackOrLay As Boolean, Price As Double, Amount As Double, MaxScore As Integer)

If BothToScore = "Yes" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home > 0 And Away > 0) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

ElseIf BothToScore = "No" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home = 0 Or Away = 0) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

End If

End Sub

Sub AddToSurfaceWinningMargin(SurfaceArray, ThisWMType As String, ThisWMValue As Double, BackOrLay As Boolean, Price As Double, Amount As Double, MaxScore As Integer)

If ThisWMType = "Home Under" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home > Away And (Home - Away) < ThisWMValue) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

ElseIf ThisWMType = "Home Over" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home > Away And (Home - Away) > ThisWMValue) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

ElseIf ThisWMType = "Away Under" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Away > Home And (Away - Home) < ThisWMValue) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home


ElseIf ThisWMType = "Away Over" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Away > Home) And ((Away - Home) > ThisWMValue) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

ElseIf ThisWMType = "Draw" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Away = Home) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

End If

End Sub
Sub AddToSurfaceWinningMarginSpread(SurfaceArray, ThisWinningMarginSpread As Double, ThisWMSType As String, BackOrLay As Boolean, Price As Double, Amount As Double, MaxScore As Integer)

Dim ThisWMSValueLow As Integer
Dim ThisWMSValueHigh As Integer

            If ThisWinningMarginSpread = 19.5 Then
            
                ThisWMSValueLow = 1
                ThisWMSValueHigh = 19
                
            ElseIf ThisWinningMarginSpread = 39 Then
            
                ThisWMSValueLow = 20
                ThisWMSValueHigh = 39
                
            ElseIf ThisWinningMarginSpread = 59 Then
                ThisWMSValueLow = 40
                ThisWMSValueHigh = 59
                
            ElseIf ThisWinningMarginSpread = 59.5 Then
            
                ThisWMSValueLow = 60
                ThisWMSValueHigh = 1000
                
            End If



If ThisWMSType = "Home" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home > Away And (Home - Away) >= ThisWMSValueLow) And (Home - Away) <= ThisWMSValueHigh Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

ElseIf ThisWMSType = "Away" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Away > Home And (Away - Home) >= ThisWMSValueLow) And (Away - Home) <= ThisWMSValueHigh Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

ElseIf ThisWMSType = "Draw" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Away = Home) Then  'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home



End If

End Sub


Sub AddToSurfaceWinningMarginNFL(SurfaceArray, ThisWinningMarginSpread As Double, ThisWMSType As String, BackOrLay As Boolean, Price As Double, Amount As Double, MaxScore As Integer)

Dim ThisWMSValueLow As Integer
Dim ThisWMSValueHigh As Integer

            If ThisWinningMarginSpread = 13 Then
            
                ThisWMSValueLow = 12
                ThisWMSValueHigh = 1000
                
            ElseIf ThisWinningMarginSpread = 43 Then
                
                ThisWMSValueLow = 42
                ThisWMSValueHigh = 1000
            
            Else 'all other cases
            
                ThisWMSValueLow = ThisWinningMarginSpread - 5
                ThisWMSValueHigh = ThisWinningMarginSpread
                           
            End If



If ThisWMSType = "Home" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home > Away And (Home - Away) >= ThisWMSValueLow) And (Home - Away) <= ThisWMSValueHigh Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

ElseIf ThisWMSType = "Away" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Away > Home And (Away - Home) >= ThisWMSValueLow) And (Away - Home) <= ThisWMSValueHigh Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

ElseIf ThisWMSType = "Tie" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Away = Home) Then  'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home



End If

End Sub
Sub AddToSurfaceTriBet(SurfaceArray, ThisTriBetType As String, ThisTriBetValue As Double, BackOrLay As Boolean, Price As Double, Amount As Double, MaxScore As Integer)


ThisTriBetType = Left(ThisTriBetType, 4)

If ThisTriBetType = "Home" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home > Away And (Home - Away) > ThisTriBetValue) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

ElseIf ThisTriBetType = "Away" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Away > Home And (Away - Home) > ThisTriBetValue) Then  'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

ElseIf ThisTriBetType = "Eith" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Abs(Home - Away) < ThisTriBetValue) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home



End If

End Sub
Sub AddToSurfaceSpread(SurfaceArray, ThisSpreadTeam As String, ThisSpreadValue As Double, BackOrLay As Boolean, Price As Double, Amount As Double, MaxScore As Integer)

If ThisSpreadTeam = "Home" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home + ThisSpreadValue > Away) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

ElseIf ThisSpreadTeam = "Away" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Away + ThisSpreadValue > Home) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

End If
End Sub

Sub AddToSurfaceTotalGameScore(SurfaceArray, ThisGameScoreValue, BackOrLay As Boolean, Price As Double, Amount As Double, MaxScore As Integer)

If ThisGameScoreValue = 150 Then
    UpperLimit = 150
    lowerlimit = 0
ElseIf ThisGameScoreValue = 221 Then
    UpperLimit = 500
    lowerlimit = 221
Else
    UpperLimit = ThisGameScoreValue
    lowerlimit = ThisGameScoreValue - 9
End If


For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If ((Home + Away) <= UpperLimit And (Home + Away) >= lowerlimit) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

End Sub

Sub AddToSurfaceTotalPoints(SurfaceArray, ThisTotalPoints, Direction As String, BackOrLay As Boolean, Price As Double, Amount As Double, MaxScore As Integer)




If Direction = "Under" Then

            For Home = 0 To MaxScore
            For Away = 0 To MaxScore
                If (Home + Away <= ThisTotalPoints) Then 'It's a win
                    If BackOrLay = True Then 'it's a BACK
                        SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
                    ElseIf BackOrLay = False Then
                        SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
                    End If
                Else 'Score doesn't match
                    If BackOrLay = True Then 'it's a BACK
                        SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
                    ElseIf BackOrLay = False Then
                        SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
                    End If
                End If
            Next Away
            Next Home

ElseIf Direction = "Over" Then

            For Home = 0 To MaxScore
            For Away = 0 To MaxScore
                If (Home + Away >= ThisTotalPoints) Then 'It's a win
                    If BackOrLay = True Then 'it's a BACK
                        SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
                    ElseIf BackOrLay = False Then
                        SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
                    End If
                Else 'Score doesn't match
                    If BackOrLay = True Then 'it's a BACK
                        SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
                    ElseIf BackOrLay = False Then
                        SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
                    End If
                End If
            Next Away
            Next Home



End If


End Sub
Sub AddToSurfaceTotalPointsOLD(SurfaceArray, ThisTotalPoints, TeamType As String, Direction As String, BackOrLay As Boolean, Price As Double, Amount As Double, MaxScore As Integer)


If TeamType = "Home" Then

    If Direction = "Under" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home <= ThisTotalPoints) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

    ElseIf Direction = "Over" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home >= ThisTotalPoints) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

ElseIf Direction = "Within" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Home <= ThisTotalPoints And Home >= (ThisTotalPoints - 14)) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home



End If




ElseIf TeamType = "Away" Then

    If Direction = "Under" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Away < ThisTotalPoints) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

    ElseIf Direction = "Over" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Away > ThisTotalPoints) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home

ElseIf Direction = "Within" Then

For Home = 0 To MaxScore
For Away = 0 To MaxScore
    If (Away <= ThisTotalPoints And Away >= (ThisTotalPoints - 14)) Then 'It's a win
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + ((Price - 1) * Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - ((Price - 1) * Amount)
        End If
    Else 'Score doesn't match
        If BackOrLay = True Then 'it's a BACK
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) - (Amount)
        ElseIf BackOrLay = False Then
            SurfaceArray(Home, Away, 1) = SurfaceArray(Home, Away, 1) + (Amount)
        End If
    End If
Next Away
Next Home


End If
End If
End Sub
Sub RunAnySurfaceToSheet()

If Cells(3, 1).Value = "Soccer" Then

    Call ArrayTest("Soccer", 7, True)

ElseIf Cells(3, 1).Value = "AFL" Then

    Call ArrayTest("AFL", 175, True)

ElseIf Cells(3, 1).Value = "NFL" Then

    Call ArrayTest("NFL", 50, True)

End If

End Sub

Sub RunAFLSurface()

Call ArrayTest("AFL", 200, False)

End Sub
Sub RunAFLSurfaceToSheet()
 
Call ArrayTest("AFL", 125, True)

End Sub
Sub RunSoccerSurfaceToSheet()

Call ArrayTest("Soccer", 7, True)

End Sub
Sub RunNFLSurface()

Call ArrayTest("NFL", 50, False)

End Sub
Sub RunNFLSurfaceToSheet()

Call ArrayTest("NFL", 50, True)

End Sub
Sub RunSoccerSurface()

Call ArrayTest("Soccer", 7, False)

End Sub


Sub RunSurface()

If Cells(4, 2) = "11897406" Then
    Call RunAFLSurface
Else
    Call RunSoccerSurface
End If

End Sub
Public Function myKurtosisCopy(myX As Double, myMean As Double, myStDev As Double, myCumul As Boolean, Optional myKurt As Double) As Double

'myKurtosis = myInput ^ 2

Dim kurt_b As Double
Dim kurt_c As Double
Dim kurt_d As Double
Dim myZ As Double
Dim kurt_Y As Double
Dim kurt_stdev As Double
Dim kurt_scale As Double
Dim kurt_limit As Double

kurt_b = 0.835665
kurt_c = 0.00000000065
kurt_d = 0.052057

If IsMissing(myKurt) Then myKurt = 0


                    If myKurt = 4 Then
                            
                            kurt_stdev = 2.439942604
                            kurt_scale = 5.327300954
                            kurt_limit = 1.640027368
                            
                    ElseIf myKurt = 3 Then
                    
                            kurt_stdev = 2.66657941
                            kurt_scale = 5.208360461
                            kurt_limit = 1.606763238
                            

                    ElseIf myKurt = 2 Then
                    
                            kurt_stdev = 2.816470362
                            kurt_scale = 5.336452302
                            kurt_limit = 1.605864425
                    
                    ElseIf myKurt = 1.5 Then
                            
                            kurt_stdev = 2.997730123
                            kurt_scale = 5.31531462
                            kurt_limit = 1.583466584
                            
                    ElseIf myKurt = 1 Then
                            
                            kurt_stdev = 3.25415359
                            kurt_scale = 5.129429972
                            kurt_limit = 1.536678181
                            
                    ElseIf myKurt = 0.5 Then
                            
                            kurt_stdev = 8.001670085
                            kurt_scale = 40.28585699
                            kurt_limit = 2.954924307
                    
                    Else 'If myKurt = 0 OR ANY OTHER VALUE Then
                            
                            kurt_stdev = 797.8792866
                            kurt_scale = 10000
                            kurt_limit = 6
                    
                    
                    End If

myZ = (myX - myMean) / myStDev

'kurt_Y = (-1 * kurt_c) + (kurt_b * myZ) + (kurt_c * (myZ ^ 2)) + (kurt_d * (myZ ^ 3))

myModStDev = kurt_limit - (Application.WorksheetFunction.Norm_Dist(myZ, 0, kurt_stdev, False) * kurt_scale)
'If myKurt = 0 Then myModStDev = myStDev

myKurtosisCopy = Application.WorksheetFunction.Norm_Dist(myZ, 0, myModStDev, myCumul)


End Function


Attribute VB_Name = "Module1"
Option Explicit


Sub FinalizeConsyRoster_Click()
   ' first make sure any non-entrants have been removed from all pools
   ' all totals use a spreadsheet formula and will be automagically updated
   ' Update all fields in the MainRoster class and save it
   Dim enteredCount As Integer, qualifierCount As Integer
   Set pcf = provisionConsyFinancials
   
'''''''''''''''''''''''''''''''''''
'	Updated 5/3/2025 via import
'''''''''''''''''''''''''''''''''''

   toggleProtection ("U")
   Application.ScreenUpdating = False
   Application.Calculation = xlCalculationManual
   
   ' this removes non-entrants from any pools
   For enteredCount = 1 To Range("FCREntriesLastEntered").Row
      If Range("FCREntriesEnteredHdr").Offset(enteredCount, 0) <> 1 Then
         ' con't clear the whole roaw it ir removes formulas
         'Range("FCREntriesEnteredHdr").Offset(enteredCount, 0).EntireRow.Select
         'Selection.ClearContents
         clearNonEntrants (enteredCount)
      End If
   Next enteredCount
   Range("FCREntriesEnteredHdr").Select
   
   Application.Calculation = xlCalculationAutomatic
   Application.ScreenUpdating = True
   toggleProtection ("P")
   
   ' we can compute the number of qualifiers given the number of entrants
   Set pcr = provisionConsyRoster
   pcr.entryCount = Range("FCREntriesPlayerCount") ' number of entrants after finalize
   pcr.qualifiers = (pcr.entryCount + 3) \ 4
   pcr.SaveConsyRoster

   UpdateConsyRosterPoolCounts    ' complete poolClasses and compute pool sizes
   
   ' compute all the payout pools and initialize the payoffs from the percentage table - in Main Financials
   InitializeConsyPayOffs
   If Not pcf.consyAllFinancialsDone Then
      MsgBox "Please complete financials to finaliaze consy", vbInformation, "Missing Expenses"
   Else
      MsgBox "Payoffs calculated on ConsyFinancials Summary tab", vbInformation, "Payoffs Done"
      ' leave other buttons not shown
      ShowDependentButtons
   End If
   ' but can compute pools payoffs as there are not dependent on expenses
   InitializeConsySidePoolPayouts
End Sub

Private Sub clearNonEntrants(en As Integer)
   ' clears all fields for a non-entrant
   ' en is the offset number from the hdr row
   ' tacit assumption there is only one pool in the consy
   Range("FCREntriesPool1Hdr").Offset(en, 0).ClearContents
End Sub
Sub RevertFinalizeConsyRoster_Click()
   ' allow user to correct last minute mistakes
   ' don't change counts - just let use try again
   ThisWorkbook.Worksheets("Entries").Shapes("RevertFinalizeButton").Visible = msoFalse
   ThisWorkbook.Worksheets("Entries").Shapes("FinalizeMConsyRosterButton").Visible = msoTrue
   UnfinalizeConsyRosterClass
   HideDependentButtons
End Sub

Sub UpdateConsyRosterClass(enteredCount As Integer, qualifiedCount As Integer)
   Set pcr = provisionMainRoster
   pcr.entryCount = enteredCount
   pcr.qualifiers = qualifiedCount
   pcr.finalized = True
   pcr.SaveConsyRoster
End Sub
Sub UpdateConsyRosterPoolCounts()
   Dim poolNumber As Integer
   ' # to be paid out is computed when payout sheets are computed
   For poolNumber = 1 To 4
      Select Case poolNumber
         Case 1:
            Set pcp1 = provisionPoolClass("CP1")
            pcp1.poolCount = Range("FCREntriesPool1Count")
            pcp1.poolPot = pcp1.poolFee * pcp1.poolCount
            If pcp1.pooltype = 2 Then
               pcp1.poolWinners = (pcp1.poolCount + 5) \ 6
            Else
               ' EQ pools have to wait for qualifers
            End If
            pcp1.SavePoolClass ("CP1")
            
         Case 2:
            Set pcp2 = provisionPoolClass("CP2")
            pcp2.poolCount = Range("FCREntriesPool2Count")
            pcp2.poolPot = pcp2.poolFee * pcp2.poolCount
            If pcp2.pooltype = 2 Then
               pcp2.poolWinners = (pcp2.poolCount + 5) \ 6
            Else
               ' EQ pools have to wait for qualifers
            End If
            pcp2.SavePoolClass ("CP2")
         
         Case 3:
            Set pcp3 = provisionPoolClass("CP3")
            pcp3.poolCount = Range("FCREntriesPool3Count")
            pcp3.poolPot = pcp3.poolFee * pcp3.poolCount
            If pcp3.pooltype = 2 Then
               pcp3.poolWinners = (pcp3.poolCount + 5) \ 6
            Else
               ' EQ pools have to wait for qualifiers
            End If
            pcp3.SavePoolClass ("CP3")
         Case 4:
            Set pcp4 = provisionPoolClass("CP4")
            pcp4.poolCount = Range("FCREntriesPool4Count")
            pcp4.poolPot = pcp4.poolFee * pcp4.poolCount
            If pcp4.pooltype = 2 Then
               pcp4.poolWinners = (pcp4.poolCount + 5) \ 6
            Else
               ' EQ pools have to wait for qualifiers
            End If
            pcp4.SavePoolClass ("CP4")
      End Select
   Next
End Sub

Sub UnfinalizeConsyRosterClass()
   ' leave counts alone until user finalizes consy roster
   Set pcr = provisionConsyRoster
   pcr.finalized = False
   pcr.SaveConsyRoster
End Sub
Sub MoveConsyQualifiers_Click()
   #If debugMode And crtrace Then
      Debug.Print "Moving consy qualifiers to Results"
   #End If
   
   Dim qualifiedCount As Integer, consyEntryCount As Integer
   Dim ConsyRosterEnd As Integer, ConsyQualifiedEnd As Integer
   Dim expectedQualifiedCount As Integer
   Dim ConsyRosterStart As Range, Key1 As Range, Key2 As Range
   Dim sortRange As Range, copyRange1 As Range
   Dim targetRange1 As Range, copyRange2 As Range, targetRange2 As Range
   
   Set pcr = provisionConsyRoster
   
   ' First check that number of marked qualifiers meets expected number from enteredCount
   'consyEntryCount = Range("FCREntriesPlayercount")
   'qualifiedCount = Range("FCREntriesQualifiedCount")
   'expectedQualifiedCount = WorksheetFunction.MRound(consyEntryCount, 4)
   'If expectedQualifiedCount <> qualifiedCount Then
   '   MsgBox "Qualfified count not 25% of entires; check qualifiers", vbInformation
   '   Exit Sub
   'End If

   ' check to see if the right number qualifiers is marke
   If Range("FCREntriesQualifiedCount") <> pcr.qualifiers Then
      MsgBox "Qualified count not 25% of entries; check qualifiers", vbInformation
      Exit Sub
   End If
   
   Sheets("Entries").Select
   toggleProtection ("U")  ' allow the sort
   
   ' turn off calculation as we move things around
   Application.Calculation = xlCalculationManual
   
   ' now we sort those that qualified and move them over to the results tab
   ConsyRosterEnd = (Range("FCREntriesNameHdr").Offset(1, 0).End(xlDown).Row)
   Set ConsyRosterStart = Range("FCREntriesEnteredHdr").Offset(1, 0)
   Set sortRange = Range(ConsyRosterStart.Address & ":" & _
                   Range("FCREntriesAmtPaidHdr").Offset(ConsyRosterEnd, 0).Address)
   Set Key1 = Range("FCREntriesQualifiedHdr").Offset(1, 0)
   Set Key2 = Range("FCREntriesNameHdr").Offset(1, 0)
   sortRange.Sort _
             Key1:=Key1, _
             Key2:=Key2, _
             Header:=xlNo


    ConsyQualifiedEnd = (Range("FCREntriesQualifiedHdr").Offset(1, 0).End(xlDown).Row)

    Sheets("Results").Select
    ClearConsyRosterResults
    Sheets("Entries").Select
    
    Set copyRange1 = Range(Range("FCREntriesNameHdr").Offset(1, 0).Address & ":" & _
                     Range("FCREntriesAccNoHdr").Offset(pcr.qualifiers, 0).Address)
    Set copyRange2 = Range(Range("FCREntriesPool1Hdr").Offset(1, 0).Address & ":" & _
                     Range("FCREntriesPool4Hdr").Offset(pcr.qualifiers, 0).Address)
    ' prepare target for incoming
    Sheets("Results").Select
    toggleProtection ("U")
    Set targetRange1 = Range("FCRResultsNameHdr").Offset(1, 0)
    Set targetRange2 = Range("FCRResultsPool1Hdr").Offset(1, 0)
    Sheets("Entries").Select
    copyRange1.Copy
    targetRange1.PasteSpecial Paste:=xlPasteValues
    
    copyRange2.Copy
    targetRange2.PasteSpecial Paste:=xlPasteValues
    
    Application.Calculation = xlCalculationAutomatic
    Sheets("Results").Select
    
    Range("FCRResultsNameHdr").Offset(1, 0).Select
    toggleProtection ("P")
    
    Sheets("Entries").Select
    
End Sub
Sub HideDependentButtons()
   With ThisWorkbook.Worksheets("Entries")
      .Shapes("FCREntriesRevert").Visible = msoFalse
      .Shapes("MoveConsyQualifiers").Visible = msoFalse
      .Shapes("FinalizeConsyRosterButton").Visible = msoTrue
   End With
End Sub
Sub ShowDependentButtons()
   With ThisWorkbook.Worksheets("Entries")
      .Shapes("FCRentriesRevert").Visible = msoTrue
      .Shapes("MoveConsyQualifiers").Visible = msoTrue
   End With
End Sub
Sub ClearConsyRosterResults()
   ' make sure the ares is clear before the move
   toggleProtection ("U")
   Range("FCRResultsAllEntryFields").ClearContents
   toggleProtection ("P")
End Sub

Sub sortConsyAlpha()
   ' make sure consy entries are in alpha order
   #If debugMode Then
      Debug.Print "Consy Alpha sort"
   #End If
   consyEntriesEnd = (Range("FCREntriesNameHdr").End(xlDown).Row)
   
   toggleProtection ("U")
   
   hdrCol = Range("FCREntriesPool4Hdr").Column
   hdrRow = Range("FCREntriesPool4Hdr").Row
   Set sortRange = Range(Range("FCREntriesEnteredHdr").Offset(1, 0).Address _
                   & ":" & Range("FCREntriesPool4Hdr") _
                   .Offset(consyEntriesEnd - hdrRow, 0).Address)
   sortRange.Sort Key1:=Range("FCREntriesNameHdr").Offset(1, 0)
   
   toggleProtection ("P")
   
End Sub

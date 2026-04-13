Attribute VB_Name = "Module1"

Option Explicit
'Public pmr As MainRoster
'Public pmf As MainFinancials

'''''''''''''''''''''''''''''''''''
'	Updated 5/3/2025 via import
'''''''''''''''''''''''''''''''''''

Sub InitializeMainPayOffs()
   ' call from main roster to create payoffs
   ' main roster makes sure all expenses have been entered
   Dim yesOrNo As Integer
   Dim mfWB As Workbook, mfWS As Worksheet
   Dim mfWBN As String, mfWSN As String
   mfWBN = "MainFinancials.xlsm"
   mfWSN = "MainSummary"
   Set mfWB = Workbooks(mfWBN)
   Set mfWS = mfWB.Worksheets(mfWSN)
   mfWS.Activate

   
'   MsgBox ("If already adjustments/entries, may have been set up from mainfinancials")
'
'   yesOrNo = MsgBox("Do you want to leave everything as-is?", vbYesNo)
'   If yesOrNo = vbYes Then
'      ' just exit sub
'      Exit Sub
'   End If
   
   ClearMainPayoffArea
   InsertMainPayoffTablePercents
   ComputeMainPayoffDollars
   FinalizePayoffs
   
End Sub

Sub ClearMainPayoffArea()
   ' scrub everything on MainSummary
   ' need to leave Final Payouts formulas intact.
   ' check if there might be prior adjustments to leave intact
   Workbooks("MainFinancials.xlsm").Worksheets("MainSummary").Activate
   toggleProtection ("U")

   Range("FMFSummaryTourneyPayoffArea").ClearContents
   Range("FMFSummaryTourneyPayoffArea").Borders.LineStyle = xlNone
      
   ' toggleProtection ("P")
End Sub

Sub InsertMainPayoffTablePercents()
   ' have to turnSyncOff for copy to work
   syncOn = False
   toggleProtection ("U")
   ' using number of players from MainRosterClass
   Set pmr = provisionMainRoster
   Dim Quals As Integer, places As Integer
   Dim PayoffColumnRange As Range
   Dim wsSource As Worksheet, wsTarget As Worksheet
   Dim wsSR As Range, wsTR As Range
   
   Set wsSource = ThisWorkbook.Worksheets("PayOffTable")
   Set wsTarget = ThisWorkbook.Worksheets("MainSummary")
   
   Quals = pmr.qualifiers  ' number of qualifers from main roster
   
   wsSource.Select
                          
   Set PayoffColumnRange = wsSource.Range( _
                           Range("FMFPayoffTablePercentOrigin").Offset(0, Quals - 3).Address _
                           & ":" & _
                           Range("FMFPayoffTablePercentOrigin").Offset(Quals - 1, Quals - 3).Address _
                           )

   Application.CutCopyMode = xlCopy
   Set wsTR = Range("FMFSummaryPercentsHdr").Offset(1, 0)
   
   PayoffColumnRange.Select
   Selection.Copy
   ' avoid activate again as it seem to screw up the pastSpecial from the selection.copy
   'Workbooks("MainFinancials.xlsm").Worksheets("MainSummary").Activate
   
   wsTarget.Select
   Range("FMFSummaryPercentsHdr").Offset(1, 0).Select
   ' Range("B10").Select
   'Application.Selection.Paste
   ActiveSheet.Paste
   '.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
   '     SkipBlanks:=False, Transpose:=False
  For places = 1 To Quals
      Range("FMFSummaryRankHdr").Offset(places, 0) = places
   Next places
   
   toggleProtection ("P")  ' block changes to sheet
   syncOn = True  ' allow sync again after copy
End Sub
   
Sub ComputeMainPayoffDollars()
   ' must check that all expenses are in
   Dim payOffPosition As Integer
   Set pmf = provisionMainFinancials
   Set pmr = provisionMainRoster
   Dim borderRange As Range, printRange As Range, adjustRange As Range
   Dim rankCell

   toggleProtection ("U")
   
   ''''''''''''''''''''''''''''''''''''
   '
   '    make sure we have all the main roster numbers copied over
   '    before we start this computattion.
   '    button was removed as copy is never optional
   '
   ''''''''''''''''''''''''''''''''''''
   
   copyMainRosterNumbers_Click
   
'   If Not pmf.mainAllFinancialsDone Then
'      MsgBox "You must complete all expenses input"
'      If Not pmf.mainNonBeneFinancialsDone Then
'         MsgBox "Non-benefit expenses on Expenses tab not marked as complete"
'      End If
'      If Not pmf.mainBeneFinancialsDone Then
'         MsgBox "Player Benefit expenses on PlayerBeneifts tab not marked as complete"
'      End If
'      Exit Sub
'   End If
   For payOffPosition = 1 To pmr.qualifiers
   
      Range("FMFSummaryRawHdr").Offset(payOffPosition, 0) = _
            Range("FMFSummaryPercentsHdr").Offset(payOffPosition, 0) * Range("FMFSummaryPrizePool")
            
      Range("FMFSummaryRoundedHdr").Offset(payOffPosition, 0) = _
            Application.WorksheetFunction.MRound(Range("FMFSummaryRawHdr").Offset(payOffPosition, 0), 5)
            
      Select Case payOffPosition
         Case 1
            Range("FMFSummaryBracketsHdr").Offset(payOffPosition, 0) = 1
         Case 2
            Range("FMFSummaryBracketsHdr").Offset(payOffPosition, 0) = 2
         Case 3 To 4
            Range("FMFSummaryBracketsHdr").Offset(payOffPosition, 0) = "3-4"
         Case 5 To 8
            Range("FMFSummaryBracketsHdr").Offset(payOffPosition, 0) = "5-8"
         Case 9 To 16
            Range("FMFSummaryBracketsHdr").Offset(payOffPosition, 0) = "9-16"
         Case 17 To 32
            Range("FMFSummaryBracketsHdr").Offset(payOffPosition, 0) = "17-32"
         Case 33 To 64
            Range("FMFSummaryBracketsHdr").Offset(payOffPosition, 0) = "33-64"
      End Select
   Next
   ' final step is to draw the borders
   Set borderRange = Range(Range("FMFSummaryRankHdr").Address & ":" & _
                            Range("FMFSummaryBracketsHdr").Offset(pmr.qualifiers, 0).Address)
   With borderRange.Borders
      .LineStyle = xlContinuous
      .Weight = xlThin
   End With
   Set adjustRange = Range(Range("FMFSummaryAdjustmentsHdr").Offset(1, 0).Address & ":" & _
                           Range("FMFSummaryAdjustmentsHdr").Offset(pmr.qualifiers, 0).Address)
   With adjustRange.Borders
      .Weight = xlMedium
      .Color = RGB(0, 176, 240)
   End With
   
   ' restore the formulas into final payouts column
   Dim Rnd As Range, Adj As Range, Frm As Range
   For payOffPosition = 1 To pmr.qualifiers
      Set Rnd = Range("FMFSummaryroundedHdr").Offset(payOffPosition, 0)
      Set Adj = Rnd.Offset(0, 1)
      Set Frm = Rnd.Offset(0, 2)
      Frm = Rnd + Adj
   Next
   

   
   toggleProtection ("P")
   
   ' now drop the payoff zones into the BracketPayOffs sheet
   ' Bracket Payoff sheets already set up with formulae to move over bracket payoffs
   ' Worksheets("BracketPayOffs").Activate
   ' ******
   ' do this after we have made any adjustments
End Sub
   
Sub FinalizePayoffs()
   ' allow user to make any adjustments
   Dim summaryPrintRange As Range, borderRange As Range
   Dim bracketPrintRange As Range
   Dim bracketNumber As Integer, rank As Integer
   Dim bWsn As String, msn As String
   Dim bWs As Worksheet, mws As Worksheet
   Set pmr = provisionMainRoster
   Set pmf = provisionMainFinancials
      
   Dim yesOrNo, qualifiers As Integer
   toggleProtection ("U")
   
   '''''''''''''''''''''''''''''''''''''''
   '
   '    Copy main roster number every time - don't ask
   '    Button was removed from MainSummary sheet
   '
   '''''''''''''''''''''''''''''''''''''''
   
   copyMainRosterNumbers_Click
   
   yesOrNo = MsgBox("Do you want to make any payoff adjustments?", vbYesNo, "Make Payoff Adjustments")
   If yesOrNo = vbYes Then
      ' go make your adjustments
      MsgBox "Make any adjustments in the blue boxs then use Finalize again."
      ' ActiveWorkbook.Worksheets("MainSummary").Activate
      Exit Sub
   Else
   
      qualifiers = pmr.qualifiers

      ' clear any prior info
      Range("FMFBracketPayOffsAllentry").ClearContents
      
      For rank = 1 To qualifiers
         Range("FMFSummaryFinalHdr").Offset(rank, 0) = _
         Range("FMFSummaryRoundedHdr").Offset(rank, 0) + _
         Range("FMFSummaryAdjustmentsHdr").Offset(rank, 0)
      Next
      
      ' first clear out bracket payoff area
      bWsn = "BracketPayOffs"
      Set bWs = ThisWorkbook.Worksheets(bWsn)
      bWs.Activate
      toggleProtection ("U")
      For bracketNumber = 1 To 6
         Range("FMFBracketPayOffsPayoffsHdr").Offset(bracketNumber, 0) = "n/a"
      Next
      
      ' create the bracket payoff sheet.
      For bracketNumber = 1 To qualifiers
         Select Case bracketNumber
            Case 1
               Range("FMFBracketPayOffsPayOffsHdr").Offset(1, 0) = Range("FMFSummaryFinalHdr").Offset(bracketNumber, 0)
               Range("FMFBracketPayOffsBracketHdr").Offset(1, 0) = "'1"
            Case 2
               Range("FMFBracketPayOffsPayOffsHdr").Offset(2, 0) = Range("FMFSummaryFinalHdr").Offset(bracketNumber, 0)
               Range("FMFBracketPayOffsBracketHdr").Offset(2, 0) = "'2"
            Case 3 To 4
               Range("FMFBracketPayOffsPayOffsHdr").Offset(3, 0) = Range("FMFSummaryFinalHdr").Offset(bracketNumber, 0)
               Range("FMFBracketPayOffsBracketHdr").Offset(3, 0) = "'3-4"
            Case 5 To 8
               Range("FMFBracketPayOffsPayOffsHdr").Offset(4, 0) = Range("FMFSummaryFinalHdr").Offset(bracketNumber, 0)
               Range("FMFBracketPayOffsBracketHdr").Offset(4, 0) = "5-8"
            Case 9 To 16
               Range("FMFBracketPayOffsPayOffsHdr").Offset(5, 0) = Range("FMFSummaryFinalHdr").Offset(bracketNumber, 0)
               Range("FMFBracketPayOffsBracketHdr").Offset(5, 0) = "9-16"
            Case 17 To 32
               Range("FMFBracketPayOffsPayOffsHdr").Offset(6, 0) = Range("FMFSummaryFinalHdr").Offset(bracketNumber, 0)
               Range("FMFBracketPayOffsBracketHdr").Offset(6, 0) = "17-32"
         End Select
      Next
      
      Set summaryPrintRange = Range(Range("FMFSummaryPrintAreaOrigin").Address & ":" & _
                             Range("FMFSummaryBracketsHdr").Offset(qualifiers, 1).Address)
      
      Set borderRange = Range(Range("FMFSummaryRankHdr").Address & ":" & _
                              Range("FMFSummaryBracketsHdr").Offset(qualifiers, 0).Address)
      borderRange.Borders.LineStyle = xlContinuous
      borderRange.Borders.Weight = xlThin
      
      msn = "MainSummary"
      Set mws = ActiveWorkbook.Worksheets(msn)
      mws.Activate
      
      ActiveSheet.PageSetup.PrintArea = summaryPrintRange.Address
      ' use print preview to give user the choice
      ActiveSheet.PrintPreview
      MsgBox "BracketPayOffs and PayOffSignOff sheets are ready for printout."
      
   End If
   
      ' restore the formulas into final payouts column
   Dim Rnd As Range, Adjust As Range, Frm As Range
   Dim payOffPosition As Integer
   For payOffPosition = 1 To pmr.qualifiers
      Set Rnd = Range("FMFSummaryroundedHdr").Offset(payOffPosition, 0)
      Set Adjust = Rnd.Offset(0, 1)
      Set Frm = Rnd.Offset(0, 2)
      Frm = Rnd + Adjust
   Next
   toggleProtection ("P")
End Sub

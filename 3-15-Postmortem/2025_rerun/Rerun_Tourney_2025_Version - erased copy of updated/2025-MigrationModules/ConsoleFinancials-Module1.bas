Attribute VB_Name = "Module1"
Option Explicit
'Public pcr As ConsyRoster
'Public pcf As ConsyFinancials

'''''''''''''''''''''''''''''''''''
'	Updated 5/3/2025 via import
'''''''''''''''''''''''''''''''''''

Sub InitializeConsyPayOffs()
   ' call from consyroster to create payoffs
   ' consy roster makes sure all expenses have been entered
   Dim yesOrNo As Integer
   Dim cfWB As Workbook, cfWS As Worksheet
   Dim cfWBN As String, cfWSN As String
   cfWBN = "ConsyFinancials.xlsm"
   cfWSN = "ConsySummary"
   Set cfWB = Workbooks(cfWBN)
   Set cfWS = cfWB.Worksheets(cfWSN)
   cfWS.Activate
   
'   MsgBox ("If already adjustments/entries, may have been set up from mainfinancials")
'
'   yesOrNo = MsgBox("Do you want to leave everything as-is?", vbYesNo)
'   If yesOrNo = vbYes Then
'      ' just exit sub
'      Exit Sub
'   End If
   
   ClearConsyPayoffArea
   InsertConsyPayoffTablePercents
   ComputeConsyPayoffDollars
   FinalizePayoffs
End Sub

Sub ClearConsyPayoffArea()
   ' scrub everything on MainSummary
   Workbooks("ConsyFinancials.xlsm").Worksheets("ConsySummary").Activate
   toggleProtection ("U")
   Range("FCFSummaryTourneyPayoffArea").ClearContents
   Range("FCFSummaryTourneyPayoffArea").Borders.LineStyle = xlNone
   ' toggleProtection ("P")
End Sub

Sub InsertConsyPayoffTablePercents()
   ' have to turnSyncOff for copy to work
   toggleProtection ("U")
   ' using number of players from ConsyRosterClass
   Set pcr = provisionConsyRoster
   Dim Quals As Integer, places As Integer
   Dim PayoffColumnRange As Range
   Dim wsSource As Worksheet, wsTarget As Worksheet
   Dim wsSR As Range, wsTR As Range
   
   Set wsSource = ThisWorkbook.Worksheets("PayOffTable")
   Set wsTarget = ThisWorkbook.Worksheets("ConsySummary")

   Quals = pcr.qualifiers  ' number of qualifiers from consy roster
   
   wsSource.Select
   
   Set PayoffColumnRange = wsSource.Range( _
                           Range("FCFPayoffTablePercentOrigin").Offset(0, Quals - 3).Address _
                           & ":" & _
                           Range("FCFPayoffTablePercentOrigin").Offset(Quals - 1, Quals - 3).Address _
                           )
 
   Application.CutCopyMode = xlCopy
   Set wsTR = Range("FCFSummaryPercentsHdr").Offset(1, 0)
   
   PayoffColumnRange.Select
   Selection.Copy
   ' avoid activate again as it seem to screw up the pastSpecial from the selection.copy
   'Workbooks("ConsyFinancials.xlsm").Worksheets("MainSummary").Activate
   
   wsTarget.Select
   Range("FCFSummaryPercentsHdr").Offset(1, 0).Select
   ' Range("B10").Select
   'Application.Selection.Paste
   ActiveSheet.Paste
   '.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
   '     SkipBlanks:=False, Transpose:=False
  For places = 1 To Quals
      Range("FCFSummaryRankHdr").Offset(places, 0) = places
   Next places
   
   toggleProtection ("P")  ' block changes to sheet
   syncOn = True  ' allow sync again after copy
End Sub
   
Sub ComputeConsyPayoffDollars()
   ' must check that all expenses are in
   Dim payOffPosition As Integer
   Set pcf = provisionConsyFinancials
   Set pcr = provisionConsyRoster
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
   
   copyConsyRosterNumbers_Click
   
'   If Not pcf.mainAllFinancialsDone Then
'      MsgBox "You must complete all expenses input"
'      If Not pcf.mainNonBeneFinancialsDone Then
'         MsgBox "Non-benefit expenses on Expenses tab not marked as complete"
'      End If
'      If Not pcf.mainBeneFinancialsDone Then
'         MsgBox "Player Benefit expenses on PlayerBeneifts tab not marked as complete"
'      End If
'      Exit Sub
'   End If

   For payOffPosition = 1 To pcr.qualifiers
   
      Range("FCFSummaryRawHdr").Offset(payOffPosition, 0) = _
            Range("FCFSummaryPercentsHdr").Offset(payOffPosition, 0) * Range("FCFSummaryPrizePool")
            
      Range("FCFSummaryRoundedHdr").Offset(payOffPosition, 0) = _
            Application.WorksheetFunction.MRound(Range("FCFSummaryRawHdr").Offset(payOffPosition, 0), 5)
      Select Case payOffPosition
         Case 1
            Range("FCFSummaryBracketsHdr").Offset(payOffPosition, 0) = 1
         Case 2
            Range("FCFSummaryBracketsHdr").Offset(payOffPosition, 0) = 2
         Case 3 To 4
            Range("FCFSummaryBracketsHdr").Offset(payOffPosition, 0) = "3-4"
         Case 5 To 8
            Range("FCFSummaryBracketsHdr").Offset(payOffPosition, 0) = "5-8"
         Case 9 To 16
            Range("FCFSummaryBracketsHdr").Offset(payOffPosition, 0) = "9-16"
         Case 17 To 32
            Range("FCFSummaryBracketsHdr").Offset(payOffPosition, 0) = "17-32"
         Case 33 To 64
            Range("FCFSummaryBracketsHdr").Offset(payOffPosition, 0) = "33-64"
      End Select
   Next
   ' final step is to draw the borders
   Set borderRange = Range(Range("FCFSummaryRankHdr").Address & ":" & _
                            Range("FCFSummaryBracketsHdr").Offset(pcr.qualifiers, 0).Address)
   With borderRange.Borders
      .LineStyle = xlContinuous
      .Weight = xlThin
   End With
   Set adjustRange = Range(Range("FCFSummaryAdjustmentsHdr").Offset(1, 0).Address & ":" & _
                           Range("FCFSummaryAdjustmentsHdr").Offset(pcr.qualifiers, 0).Address)
   With adjustRange.Borders
      .Weight = xlMedium
      .Color = RGB(0, 176, 240)
   End With
   
   ' finally set print area
   ' first set range to be printed
   Set printRange = Range(Range("FCFSummaryPrintAreaOrigin").Address & ":" & _
                           Range("FCFSummaryBracketsHdr").Offset(pcr.qualifiers, 1).Address)
                           
   ' restore the formulas into final payouts column
   Dim Rnd As Range, Adj As Range, Frm As Range
   For payOffPosition = 1 To pcr.qualifiers
      Set Rnd = Range("FCFSummaryroundedHdr").Offset(payOffPosition, 0)
      Set Adj = Rnd.Offset(0, 1)
      Set Frm = Rnd.Offset(0, 2)
      Frm = Rnd + Adj
   Next
   
   ' ActiveSheet.PageSetup.PrintArea = printRange.Address
   ' use print preview to give user the choice
   ' ActiveSheet.PrintPreview
   ' do this in finalize payoffs
   
   ' ActiveSheet.PrintOut
   
   toggleProtection ("P")
   
   ' now drop the payoff zones into the BracketPayOffs sheet
   ' Bracket Payoff sheets already set up with formulae to move over bracket payoffs
   ' Worksheets("BracketPayOffs").Activate
     
End Sub
   
Sub FinalizePayoffs()
   ' allow user to make any adjustments
   Dim summaryPrintRange As Range, borderRange As Range
   Dim bracketPrintRange As Range
   Dim bracketNumber As Integer, rank As Integer
   Dim bWsn As String, csn As String
   Dim bWs As Worksheet, cws As Worksheet
   Set pcr = provisionConsyRoster
   Set pcf = provisionConsyFinancials
   
   Dim yesOrNo As Integer, qualifiers As Integer
   toggleProtection ("U")
   
   '''''''''''''''''''''''''''''''''''''''
   '
   '    Copy main roster number every time - don't ask
   '    Button was removed from MainSummary sheet
   '
   '''''''''''''''''''''''''''''''''''''''
   
   copyConsyRosterNumbers_Click
   
   yesOrNo = MsgBox("Do you want to make any payoff adjustments?", vbYesNo, "Make Payoff Adjustments")
   If yesOrNo = vbYes Then
      ' go make your adjustments
      MsgBox "Make any adjustments in the blue boxs then use Finalize Payoffs again."
      ' ActiveWorkbook.Worksheets("ConsySummary").Activate
      Exit Sub
   Else
      
      qualifiers = pcr.qualifiers
      
      ' clear any prior info
      Range("FCFBracketsAllEntries").ClearContents
      
      ' finlize the adjusted payouts
      
      ' ThisWorkbook.Activate
      For rank = 1 To qualifiers
         Range("FCFSummaryFinalHdr").Offset(rank, 0) = _
         Range("FCFSummaryRoundedHdr").Offset(rank, 0) + _
         Range("FCFSummaryAdjustmentsHdr").Offset(rank, 0)
      Next
      
      ' first clear out bracket payoff area
      bWsn = "BracketPayOffs"
      Set bWs = ThisWorkbook.Worksheets(bWsn)
      bWs.Activate
      toggleProtection ("U")
      For bracketNumber = 1 To 6
         Range("FCFBracketPayoffsPayOffsHdr").Offset(bracketNumber, 0) = "n/a"
      Next
      ' create the bracket payoff sheet
      For bracketNumber = 1 To qualifiers
         Select Case bracketNumber
            Case 1
               Range("FCFBracketPayOffsPayOffsHdr").Offset(1, 0) = Range("FCFSummaryFinalHdr").Offset(bracketNumber, 0)
               Range("FCFBracketPayOffsBracketHdr").Offset(1, 0) = "'1"
            Case 2
               Range("FCFBracketPayOffsPayOffsHdr").Offset(2, 0) = Range("FCFSummaryFinalHdr").Offset(bracketNumber, 0)
               Range("FCFBracketPayOffsBracketHdr").Offset(2, 0) = "'2"
            
            Case 3 To 4
               Range("FCFBracketPayOffsPayOffsHdr").Offset(3, 0) = Range("FCFSummaryFinalHdr").Offset(bracketNumber, 0)
               Range("FCFBracketPayOffsBracketHdr").Offset(3, 0) = "'3-4"
            
            Case 5 To 8
               Range("FCFBracketPayOffsPayOffsHdr").Offset(4, 0) = Range("FCFSummaryFinalHdr").Offset(bracketNumber, 0)
               Range("FCFBracketPayOffsBracketHdr").Offset(4, 0) = "'5-8"
            
            Case 9 To 16
               Range("FCFBracketPayOffsPayOffsHdr").Offset(5, 0) = Range("FCFSummaryFinalHdr").Offset(bracketNumber, 0)
               Range("FCFBracketPayOffsBracketHdr").Offset(5, 0) = "'9-16"
               
            Case 17 To 32
               Range("FCFBracketPayOffsPayOffsHdr").Offset(6, 0) = Range("FCSummaryFinalHdr").Offset(bracketNumber, 0)
               Range("FCFBracketPayOffsBracketHdr").Offset(6, 0) = "'17-32"
               
            
         End Select
      Next

      Set summaryPrintRange = Range(Range("FCFSummaryPrintAreaOrigin").Address & ":" & _
                             Range("FCFSummaryBracketsHdr").Offset(qualifiers, 1).Address)
      
      Set borderRange = Range(Range("FCFSummaryRankHdr").Address & ":" & _
                              Range("FCFSummaryBracketsHdr").Offset(qualifiers, 0).Address)
      borderRange.Borders.LineStyle = xlContinuous
      borderRange.Borders.Weight = xlThin
      
      csn = "ConsySummary"
      Set cws = ActiveWorkbook.Worksheets(csn)
      cws.Activate
      
      ActiveSheet.PageSetup.PrintArea = summaryPrintRange.Address
      ' use print preview to give user the choice
      ActiveSheet.PrintPreview
      MsgBox "BracketPayOffs and PayOffSignOff sheets are ready for printout."
            
   End If
   
      ' restore the formulas into final payouts column
   Dim Rnd As Range, Adjust As Range, Frm As Range
   Dim payOffPosition As Integer
   For payOffPosition = 1 To pcr.qualifiers

      Set Rnd = Range("FCFSummaryroundedHdr").Offset(payOffPosition, 0)
      Set Adjust = Rnd.Offset(0, 1)
      Set Frm = Rnd.Offset(0, 2)
      Frm = Rnd + Adjust
   Next
   toggleProtection ("P")
End Sub



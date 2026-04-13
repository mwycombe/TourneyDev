Attribute VB_Name = "TourneyButtonHandlers"
Option Explicit

'Public pmr As MainRoster
'Public pmf As MainFinancials
'Public pmp1 As PoolClass
'Public pmp2 As PoolClass
'Public pmp3 As PoolClass
'Public pmp4 As PoolClass

Dim mfWbn As String

Sub InitializeMainPayOffs()
   ' go to MainFinancials to perform this task
   mfWbn = "MainFinancials.xlsm"
   openIfClosed mfWbn
   
   Set pmf = provisionMainFinancials
   If Not pmf.mainAllFinancialsDone Then
      If Not pmf.mainBeneFinancialsDone Then
         MsgBox "Player Benefit Expenses Not Complete", vbInformation, "Benefit Expenses Missing"
         
      End If
      If Not pmf.mainNonBeneFinancialsDone Then
         MsgBox "Non-benefit Expenses Not Complete", vbInformation, "Non-Benefit Expenses Missing"
      End If
      MsgBox "You MUST complete all expenses before Payoffs can be calculated", vbInformation, "Expenses Missing"
      MsgBox "Please go MainFinancials and finalize expenses"
      Exit Sub
   End If
   ' this is the only place from where we initialize the playoffs
   ' the user must get the expenses otherwise cannot finalize registration main entries
   
   Run ("MainFinancials.xlsm!InitializeMainPayOffs")
   
End Sub

Sub InitializeMainSidePoolPayouts()
   ' go to MainSidePools to perform this task
   If Not IsWorkBookOpen("MainSidePools.xlsm") Then
      Workbooks.Open "MainSidePools.xlsm"
   End If
   
   ' clear out final pool payout sheets
   Dim poolWS As Worksheet
   Dim poolWSN As String
   Dim poolNumber As Integer
   
   ' let this happen in the CreateMainSidePoolPrintouts
'   For poolNumber = 1 To 4
'      poolWSN = "PrtPool" + Mid(Str(poolNumber), 2, 1)
'      Worksheets(poolWSN).resetSheet
'   Next
   
   Run "MainSidePools.xlsm!CreateMainSidePoolPrintouts"
End Sub

Sub CopyMainToConsy_Click()
   ' once main has started, copy all main entries to consy, and set up consy check-in
   ' only copy entrants; remove entered flags and don't ocpy pools
   ' include those that qualified as they may flop out into the consy
   
   ' ConsyRoster->Entries is protected, but target cells are unlocked
   ' So, no need to unprotect the target sheet
   
   Dim sourceWorkbook As Workbook, targetWorkbook As Workbook
   Dim sourceRange As Range, targetRange As Range
   Dim sourceWBName As String, sourceWSName As String
   Dim targetWBName As String, targetWSName As String
   Dim sourceSheet As Worksheet, targetSheet As Worksheet
   sourceWBName = "MainRoster.xlsm"
   sourceWSName = "Entries"
   targetWBName = "ConsyRoster.xlsm"
   targetWSName = "Entries"
   

   Set pmr = provisionMainRoster
   openIfClosed (targetWBName)
   turnSyncOff
   ThisWorkbook.Activate
   Set sourceWorkbook = Workbooks(sourceWBName)
   Set targetWorkbook = Workbooks(targetWBName)
   Set sourceSheet = sourceWorkbook.Worksheets(sourceWSName)
   Set targetSheet = targetWorkbook.Worksheets(targetWSName)
   
   Set sourceRange = Range( _
                     Range("FMREntriesNameHdr").Offset(1, 0).Address & ":" & _
                     Range("FMREntriesAccNoHdr").Offset(pmr.entryCount, 0).Address _
                       )
   targetWorkbook.Activate
   targetSheet.Select
   
   'clear target area of any data
   Range("FCRAllInputArea").ClearContents
   
   Set targetRange = Range("FCREntriesNameHdr").Offset(1, 0)
   
   sourceSheet.Activate
   sourceRange.Copy
   targetWorkbook.Activate
   targetSheet.Select
   targetRange.Select
   targetSheet.Paste
   
   ' remove the pasted area shading
   Range("FCREntriesEnteredHdr").Offset(1, 0).Select
   
   sourceSheet.Activate    ' come back to the source
   
   ' now sort what we just copied over
   ' must firs activate the target sheet to call sort
   targetWorkbook.Activate
   targetSheet.Activate
   
   Application.Run (targetWBName & "!" & "sortConsyAlpha")
   
   sourceWorkbook.Activate
   sourceSheet.Activate
   
   ' now show print button
   ''''''''''''''''''''''''
   ' delete this as print button removed
   ''''''''''''''''''''''''
'   sourceWorkbook.Activate
'   sourceWorkbook.Worksheets("Entries") _
'      .Shapes("consyCheckinList").Visible = msoTrue

   sourceSheet.EnableSelection = xlNoSelection
   sourceSheet.Protect
   turnSyncOn
   
   ' call the routine that was on the Print Consy button to setup
   ' consyCheckin sheet
   PrintConsyCheckIn_Click
   
End Sub

Sub PrintConsyCheckIn_Click()
   ' print out the check-in sheet for the consy
   ' this is over in the consyRoster workbook
   Dim consyWbn As String
   consyWbn = "ConsyRoster.xlsm"
   openIfClosed consyWbn
   Run ("ConsyRoster.xlsm!setupConsyCheckIn")
   Application.Calculation = xlCalculationManual
   ThisWorkbook.Activate
   Application.Calculation = xlCalculationAutomatic
End Sub

Sub sortMainResults_Click()
   ' get the counts and qualifiers from the MainRoster
   Dim qualifiers As Integer
   Dim rank As Integer
   Dim sortRange As Range
   Dim Key1 As Range, Key2 As Range, Key3 As Range, Key4 As Range
   Dim ws As Worksheet
   Set pmr = provisionMainRoster()
   qualifiers = pmr.qualifiers
   
   
   ' sort the retults area qualifiers
   toggleProtection ("U")  ' allow the sort
   Application.Calculation = xlCalculationManual
   
   Set sortRange = Range(Range("FMRResultsNameHdr").Address & ":" & _
                    Range("FMRResultsPool4Hdr").Offset(qualifiers, 0).Address)
   '
   ' sort is limited to 3 keys so have to build sortfields collection my hand
   '
   Set ws = Worksheets("Results")
   ws.Select
   ws.Sort.SortFields.Clear
   Set Key1 = Range("FMRResultsGamePointsHdr").Offset(1, 0)
   Set Key2 = Range("FMRResultsGamesWonHdr").Offset(1, 0)
   Set Key3 = Range("FMRResultsSpreadPointsHdr").Offset(1, 0)
   Set Key4 = Range("FMRResultsPlusPointHdr").Offset(1, 0)
   With ws
      .Sort.SortFields.Add Key:=Key1, SortOn:=xlSortOnValues, _
                     Order:=xlDescending, DataOption:=xlSortNormal
      .Sort.SortFields.Add Key:=Key2, SortOn:=xlSortOnValues, _
                     Order:=xlDescending, DataOption:=xlSortNormal
      .Sort.SortFields.Add Key:=Key3, SortOn:=xlSortOnValues, _
                     Order:=xlDescending, DataOption:=xlSortNormal
      .Sort.SortFields.Add Key:=Key4, SortOn:=xlSortOnValues, _
                     Order:=xlDescending, DataOption:=xlSortNormal
      With .Sort
         .SetRange sortRange
         .Header = xlYes
         .Orientation = xlRows
         .Apply      ' this is what fires the sort option
      End With
   End With
   Application.Calculation = xlCalculationAutomatic
   
   ' Playoff Place is used to record bracket where player finishes
   ' To be used later for computing playoff master pts.
   
   toggleProtection ("P")
   Range("FMRResultsNameHdr").Offset(1, 0).Select
End Sub

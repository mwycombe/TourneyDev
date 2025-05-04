Attribute VB_Name = "ButtonHandlers"
Option Explicit
'Public pmf As MainFinancials
'Public pmr As MainRoster
'Public pmt As TourneyClass
'Public pms As MainSetUp

'''''''''''''''''''''''''''''''''''
'	Updated 5/3/2025 via import
'''''''''''''''''''''''''''''''''''

Sub copyMainRosterNumbers_Click()
   ' copy tourney numbers from Main Roster class and insert into MainFinancials summary
   Dim yesOrNo As Integer

   Set pmf = provisionMainFinancials
   Set pmr = provisionMainRoster
   Set ptc = provisionTourneyClass
   Set pms = provisionMainSetUp
   
'''''''''''''''''''''''''''''''''''''''''''
'
'   always copy the numbers over
'
'''''''''''''''''''''''''''''''''''''''''''
   
'   If pmf.mainNumbersAcquired Then
'      yesOrNo = MsgBox("Do you want to copy the roster tourney numbers again?", vbYesNo, "Main Roster numbers already acquired")
'      If yesOrNo = vbYes Then
'         ClearOldSummaryNumbers
'         Range("FMFSummaryPlayerCount") = pmr.entryCount
'         Range("FMFSummaryEntryFee") = pms.entryFee
'         Range("FMFSummaryAccFee") = pms.accFee
'         Range("FMFSummaryPerPlayer") = pms.perCapitaDonation
'         Range("FMFSummaryFixedDonation") = pms.fixedDonation
'         pmf.mainNumbersAcquired = True
'      End If
'   Else
      ClearOldSummaryNumbers
      Range("FMFSummaryPlayerCount") = pmr.entryCount
      Range("FMFSummaryEntryFee") = pms.entryFee
      Range("FMFSummaryAccFee") = pms.accFee
      Range("FMFSummaryPerPlayer") = pms.perCapitaDonation
      Range("FMFSummaryFixedDonation") = pms.fixedDonation
      Range("FMFSummaryQualifierCount") = pmr.qualifiers
      pmf.mainNumbersAcquired = True
'   End If
   
   pmf.SaveMainFinancials
   
End Sub

Sub ClearOldSummaryNumbers()
   ' just in case there's any detritus in the fields
         Range("FMFSummaryPlayerCount").ClearContents
         Range("FMFSummaryEntryFee").ClearContents
         Range("FMFSummaryAccFee").ClearContents
         Range("FMFSummaryPerPlayer").ClearContents
         ' do not do this - leave formula in place
         ' Range("FMFSummaryAdjustments").ClearContents
         pmf.mainNumbersAcquired = False
End Sub
Sub mainFinancialsDone_Click()
   ' check with user if everything is in
   ' update mainfinancials class
   Dim yesOrNo As Integer
   Set pmf = provisionMainFinancials
   If Not (pmf.mainBeneFinancialsDone And pmf.mainNonBeneFinancialsDone) Then
      MsgBox "Both Benefit Expenses and Non-benefit Expenses must show done to complete Financials", vbInformation, "All Financials Not Done"
      pmf.mainAllFinancialsDone = False
   Else
      MsgBox "All Financials Marked as Done", vbInformation, "All Financials Complete"
      pmf.mainAllFinancialsDone = True
   End If
   pmf.SaveMainFinancials
End Sub
Sub beneFinancialsDone_Click()
   ' check with user if benefit expenses are in
   ' update mainfinancials class
   Dim yesOrNo As Integer
   Set pmf = provisionMainFinancials
   yesOrNo = MsgBox("Have you entered all Player Benefit Expenses on PlayerBenefits tab?", vbYesNo, "Player Benefit Expenses Not Done")
   ' capture player benefit expenses
   Set pmf = provisionMainFinancials
   pmf.mainLocationXP = Range("FMFSummaryDonation")
   If yesOrNo = vbYes Then
      ' mark mainfinancials class
      recordBeneFinancialsDone
      ' check other expense
      If Not pmf.mainNonBeneFinancialsDone Then
         MsgBox "Non Player Benefit Expenses not yet marked complete", vbInformation, "Pleaes Complete Player Benefit Expenses"
         pmf.mainAllFinancialsDone = False
      End If
   Else
      ' mark mainfinancials class
      pmf.mainBeneFinancialsDone = False
      pmf.mainAllFinancialsDone = False
   End If
   pmf.mainPlayerBenefitXP = Range("FMFPlayerBenefitsTotal")
   pmf.SaveMainFinancials     ' always save a class after an update
End Sub
Sub nonBeneFinancialsDone_Click()
   ' check with user if non-benefit expenses are in
   'update mainfinancials class
   Dim yesOrNo As Integer
   Set pmf = provisionMainFinancials
   yesOrNo = MsgBox("Have you entered all Non-Benefit Expenses on Expenses tab?", vbYesNo, "Non-Benefit Expenses Not Done")
   If yesOrNo = vbYes Then
      ' mark mainfinancials class
      recordNonBeneFinancialsDone
      ' check other expense
      If Not pmf.mainBeneFinancialsDone Then
         MsgBox "Player Benefit Expenses not yet marked complete", vbInformation, "Please Complete Non-Benefit Expenses"
         pmf.mainAllFinancialsDone = False
      End If
   Else
      ' mark mainfinancials class
      pmf.mainNonBeneFinancialsDone = False
      pmf.mainAllFinancialsDone = False
   End If
   pmf.SaveMainFinancials     ' always save a class after an update
   
End Sub
Sub resetMainFinancials_Click()
   ' resets all financial entires
   Dim yesOrNo As Integer
   yesOrNo = MsgBox("Do you want to clear all financials?", vbYesNo)
   If yesOrNo = vbYes Then
      yesOrNo = MsgBox("Are yu really sure you want to reset?", vbYesNo)
      If yesOrNo = vbNo Then
         MsgBox ("No changes were made.")
         Exit Sub
      End If
   Else
      MsgBox ("No changes were made.")
   End If
   'clear details but leave summary formulae intact
   Range("MMFMiscExpensesEntries").ClearContents
   Range("FMFBenefitsentries").ClearContents
   Range("FMFSummaryAdjEntries").ClearContents

End Sub

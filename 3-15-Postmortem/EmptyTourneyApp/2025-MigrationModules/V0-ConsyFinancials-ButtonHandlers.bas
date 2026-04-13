Attribute VB_Name = "ButtonHandlers"
Option Explicit
'Public pcf As ConsyFinancials
'Public pcr As ConsyRoster
'Public pmt As TourneyClass
'Public pcs As ConsySetUp

Sub copyConsyRosterNumbers_Click()
   ' copy tourney numbers from Consy Roster class and insert into ConsyFiancials summary
   Dim yesOrNo As Integer

   Set pcf = provisionConsyFinancials
   Set pcr = provisionConsyRoster
   Set ptc = provisionTourneyClass
   Set pcs = provisionConsySetUp
   
'''''''''''''''''''''''''''''''''''''''''''
'
'   always copy the numbers over
'
'''''''''''''''''''''''''''''''''''''''''''
   
'   If pcf.consyNumbersAcquired Then
'      yesOrNo = MsgBox("Do you want to copy the roster tourney numbers again?", vbYesNo, "Consy Roster numbers already acquired")
'      If yesOrNo = vbYes Then
'         ClearOldSummaryNumbers
'         Range("FCFSummaryPlayerCount") = pcr.entryCount
'         Range("FCFSummaryEntryFee") = pcs.entryFee
'         Range("FCFSummaryAccFee") = pcs.accFee
'         Range("FCFSummaryPerPlayer") = pcs.perCapitaDonation
'         Range("FCFSummaryFixedDonation") = pcs.fixedDonation
'         pcf.consyNumbersAcquired = True
'      End If
'   Else
      ClearOldSummaryNumbers
      Range("FCFSummaryPlayerCount") = pcr.entryCount
      Range("FCFSummaryEntryFee") = pcs.entryFee
      Range("FCFSummaryAccFee") = pcs.accFee
      Range("FCFSummaryPerPlayer") = pcs.perCapitaDonation
      Range("FCFSummaryFixedDonation") = pcs.fixedDonation
      Range("FCFSummaryQualifierCount") = pcr.qualifiers
      pcf.consyNumbersAcquired = True
'   End If
   
   pcf.SaveConsyFinancials
End Sub

Sub ClearOldSummaryNumbers()
   ' just in case there's any detritus in the fields
         Range("FCFSummaryPlayerCount").ClearContents
         Range("FCFSummaryEntryFee").ClearContents
         Range("FCFSummaryAccFee").ClearContents
         Range("FCFSummaryPerPlayer").ClearContents
         ' do not do this - leave formula in place
         ' Range("FCFSummaryAdjustments").ClearContents
         pcf.consyNumbersAcquired = False
End Sub
Sub consyFinancialsDone_Click()
   ' check with user if everything is in
   ' update consyfinancials class
   Dim yesOrNo As Integer
   Set pcf = provisionConsyFinancials
   If Not (pcf.consyBeneFinancialsDone And pcf.consyNonBeneFinancialsDone) Then
      MsgBox "Both Benefit Expenses and Non-benefit Expenses must show done to complete Financials", vbInformation, "All Financials Not Done"
      pcf.consyAllFinancialsDone = False
   Else
      MsgBox "All Financials Marked as Done", vbInformation, "All Financials Complete"
      pcf.consyAllFinancialsDone = True
   End If
   pcf.SaveConsyFinancials
End Sub
Sub beneFinancialsDone_Click()
   ' check with user if benefit expenses are in
   ' update consyfinancials class
   Dim yesOrNo As Integer
   Set pcf = provisionConsyFinancials
   pcf.consyLocationXP = Range("FCFSummaryDonation")
   yesOrNo = MsgBox("Have you entered all Player Benefit Expenses on PlayerBenefits tab", vbYesNo, "Player Benefit Expenses Not Done")
   If yesOrNo = vbYes Then
      ' mark consyfinancials class
      pcf.consyBeneFinancialsDone = True
      ' check other expense
      If Not pcf.consyNonBeneFinancialsDone Then
         MsgBox "Non Player Benefit Expenses not yet marked complete", vbInformation, "Pleaes Complete Player Benefit Expenses"
         pcf.consyAllFinancialsDone = False
      End If
   Else
      ' mark consyfinancials class
      pcf.consyBeneFinancialsDone = False
      pcf.consyAllFinancialsDone = False
   End If
   pcf.consyPlayerBenefitXP = Range("FCFBenefitsTotal")
   pcf.SaveConsyFinancials     ' always save a class after an update
End Sub
Sub nonBeneFinancialsDone_Click()
   ' check with user if non-beneift expenses are in
   'update consyfinancials class
   Dim yesOrNo As Integer
   Set pcf = provisionConsyFinancials
   yesOrNo = MsgBox("Have you entered all Non-Benefit Expenses on Expenses tab", vbYesNo, "Non-Benefit Expenses Not Done")
   If yesOrNo = vbYes Then
      ' mark consyfinancials class
      pcf.consyNonBeneFinancialsDone = True
      ' check other expense
      If Not pcf.consyBeneFinancialsDone Then
         MsgBox "Player Benefit Expenses not yet marked complete", vbInformation, "Pleaes Complete Non-Benefit Expenses"
         pcf.consyAllFinancialsDone = False
      End If
   Else
      ' mark mainfinacials class
      pcf.consyNonBeneFinancialsDone = False
      pcf.consyAllFinancialsDone = False
   End If
   pcf.SaveConsyFinancials     ' always save a class after an update
   
End Sub


Attribute VB_Name = "CopyConsyCheckInCode"
Option Explicit

Dim consyEntriesLastRow As Integer
Dim sourceRange As Range, targetRange As Range

'''''''''''''''''''''''''
'	Updated 5/3/2025
'''''''''''''''''''''''''

Sub setupConsyCheckIn()
   ' copy entries into the consy check-in sheet
   Dim consyWbn As String, consyWsn As String
   Dim consyWb As Workbook, consyWs As Worksheet
      

   consyWbn = "ConsyRoster.xlsm"
   consyWsn = "Entries"
   Set consyWb = Workbooks(consyWbn)
   Set consyWs = consyWb.Worksheets(consyWsn)
   
   consyWs.Activate
   consyEntriesLastRow = Range("FCREntriesNameHdr").End(xlDown).Row
   hdrRow = Range("FCREntriesNameHdr").Row
   Set sourceRange = Range(Range("FCREntriesNameHdr").Offset(1, 0).Address & ":" & _
                           Range("FCREntriesAccNoHdr") _
                           .Offset(consyEntriesLastRow - hdrRow, 0) _
                           .Address)
   Set targetRange = Range("FCRCCInNameHdr").Offset(1, 0)
   sourceRange.Select
   sourceRange.Copy
   targetRange.PasteSpecial Paste:=xlPasteValues
   Range("FCREntriesFirstName").Select
   ' Range("FCRCCInNameHdr").Offset(1, 0).Select
   
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PartsForm2 
   Caption         =   "PartsForm2"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   OleObjectBlob   =   "Parts2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PartsForm2"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AllInventory As New InventoryItem2

' Display the items.
Private Sub DisplayItems()
    txtItems.Text = _
        AllInventory.TextValue(0)
End Sub

Private Sub txtItems_Change()

End Sub

' Load the parts information.
Private Sub UserForm_Initialize()
Dim fname As String
Dim fnum As Integer
Dim txt As String

    ' Open the file.
    fname = ActiveDocument.Path & "\parts2.dat"
    fnum = FreeFile
    Open fname For Input As #fnum

    ' Read all the text at once.
    txt = Input(LOF(fnum), #fnum)
    Close fnum

    ' Set the serialization.
    AllInventory.Serialization(0) = txt

    ' Display the items.
    DisplayItems
End Sub
' Save the items for next time.
Private Sub UserForm_Terminate()
Dim fname As String
Dim fnum As Integer

    ' Open the file.
    fname = ActiveDocument.Path & "\parts2.dat"
    fnum = FreeFile
    Open fname For Output As #fnum

    ' Save the serialization.
    Print #fnum, AllInventory.Serialization(0)
    Close fnum
End Sub

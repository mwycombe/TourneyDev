VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PartsForm1 
   Caption         =   "PartsForm1"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   OleObjectBlob   =   "Parts1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PartsForm1"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AllInventory As New InventoryItem1

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

    ' Open the file.
    fname = ActiveDocument.Path & "\parts1.dat"
    fnum = FreeFile
    Open fname For Input As #fnum

    ' Read the inventory.
    AllInventory.FileInput fnum

    Close fnum

    ' Display the items.
    DisplayItems
End Sub
' Save the items for next time.
Private Sub UserForm_Terminate()
Dim fname As String
Dim fnum As Integer

    ' Open the file.
    fname = ActiveDocument.Path & "\parts1.dat"
    fnum = FreeFile
    Open fname For Output As #fnum

    ' Save the inventory items.
    AllInventory.FileWrite fnum
    
    Close fnum
End Sub

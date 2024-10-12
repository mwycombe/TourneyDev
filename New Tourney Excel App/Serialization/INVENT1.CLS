VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InventoryItem1"
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ItemName As String
Public ItemParts As New Collection
' Return a textual value for this item.
' End the string with vbCrLf.
Public Function TextValue(indent As Integer) As String
Dim item As InventoryItem1
Dim txt As String

    txt = Space$(indent) & ItemName & vbCrLf
    For Each item In ItemParts
        txt = txt & item.TextValue(indent + 4)
    Next item
    TextValue = txt
End Function
' Write the object's information into fnum,
' a file already open for output.
Public Sub FileWrite(fnum As Integer)
Dim part As InventoryItem1

    ' Write the object-specific data.
    Write #fnum, ItemName

    ' Write the number of parts.
    Write #fnum, ItemParts.Count

    ' Make the parts write themselves.
    For Each part In ItemParts
        part.FileWrite fnum
    Next part
End Sub
' Read the object's information from fnum,
' a file already open for input.
Public Sub FileInput(fnum As Integer)
Dim i As Integer
Dim num_parts As Integer
Dim part As InventoryItem1

    ' Read the object-specific data.
    Input #fnum, ItemName

    ' Read the number of parts.
    Input #fnum, num_parts

    ' Create the parts and make them read themselves.
    For i = 1 To num_parts
        Set part = New InventoryItem1
        ItemParts.Add part
        part.FileInput fnum
    Next i
End Sub

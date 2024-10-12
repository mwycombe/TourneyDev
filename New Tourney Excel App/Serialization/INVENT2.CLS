VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InventoryItem2"
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ItemName As String
Public ItemParts As Collection

Private Const dflt_ItemName = ""
' Find and remove the next token from
' this string.
'
' Tokens are stored in the format:
'    name1(value1)name2(value2)...
' Invisible characters (tabs, vbCrLf,
' spaces, etc.) are allowed before names.
Private Sub GetToken(txt As String, token_name As String, token_value As String)
Dim open_pos As Integer
Dim close_pos As Integer
Dim txtlen As Integer
Dim num_open As Integer
Dim i As Integer
Dim ch As String

    ' Remove initial invisible characters.
    TrimInvisible txt

    ' If the string is empty, do nothing.
    token_name = ""
    token_value = ""
    If txt = "" Then Exit Sub

    ' Find the opening parenthesis.
    open_pos = InStr(txt, "(")
    txtlen = Len(txt)
    If open_pos = 0 Then open_pos = txtlen

    ' Find the corresponding closing parenthesis.
    num_open = 1
    For i = open_pos + 1 To txtlen
        ch = Mid$(txt, i, 1)
        If ch = "(" Then
            num_open = num_open + 1
        ElseIf ch = ")" Then
            num_open = num_open - 1
            If num_open = 0 Then Exit For
        End If
    Next i
    If open_pos = 0 Or i > txtlen Then
        ' There is something wrong.
        Err.Raise vbObjectError + 1, _
            "InventoryItem.GetToken", _
            "Error parsing serialization """ & txt & """"
    End If
    close_pos = i

    ' Get token name and value.
    token_name = Left$(txt, open_pos - 1)
    token_value = Mid$(txt, open_pos + 1, close_pos - open_pos - 1)
    TrimInvisible token_name
    TrimInvisible token_value

    ' Remove the token name and value
    ' from the serialization string.
    txt = Right$(txt, txtlen - close_pos)
End Sub
' Remove leading invisible characters
' from the string (tab, space, CR, etc.)
Private Sub TrimInvisible(txt As String)
Dim txtlen As Integer
Dim i As Integer
Dim ch As String

    txtlen = Len(txt)
    For i = 1 To txtlen
        ' See if this character is visible.
        ch = Mid$(txt, i, 1)
        If ch > " " And ch <= "~" Then Exit For
    Next i
    If i > 1 Then _
        txt = Right$(txt, txtlen - i + 1)
End Sub
' Return a textual value for this item.
' End the string with vbCrLf.
Public Function TextValue(indent As Integer) As String
Dim item As InventoryItem2
Dim txt As String

    txt = Space$(indent) & ItemName & vbCrLf
    For Each item In ItemParts
        txt = txt & item.TextValue(indent + 4)
    Next item
    TextValue = txt
End Function
' Return the object's serialization.
Public Property Get Serialization(indent As Integer) As String
Dim txt As String
Dim part As InventoryItem2

    ' Write the object-specific data.
    If ItemName <> dflt_ItemName Then _
        txt = txt & Space$(indent) & _
        "ItemName(" & ItemName & ")" & _
        vbCrLf

    ' Make the parts write themselves.
    For Each part In ItemParts
        txt = txt & Space$(indent) & _
        "Part(" & vbCrLf & _
        part.Serialization(indent + 4) & _
        Space$(indent) & ")" & vbCrLf
    Next part

    Serialization = txt
End Property
' Initialize the object using the serialization.
Public Property Let Serialization(indent As Integer, new_value As String)
Dim token_name As String
Dim token_value As String
Dim part As InventoryItem2

    ' Start with an empty collection.
    Set ItemParts = New Collection

    ' Initialize all values to defaults.
    ItemName = dflt_ItemName

    ' Examine each token in turn.
    Do
        ' Get the token name and value.
        GetToken new_value, token_name, token_value
        If token_name = "" Then Exit Do

        ' Save the value appropriately.
        Select Case token_name
            Case "ItemName"
                ItemName = token_value
            Case "Part"
                ' Create a new part and
                ' make it unserialize itself.
                Set part = New InventoryItem2
                ItemParts.Add part
                part.Serialization(0) = token_value
        End Select
    Loop
End Property

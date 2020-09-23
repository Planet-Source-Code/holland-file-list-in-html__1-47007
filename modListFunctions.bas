Attribute VB_Name = "modListFunctions"
'This was created by YM
'Copyright YM, 2002
'----------------------

'Move a item from 1 list to the other
Sub ListToList(L1 As ListBox, L2 As ListBox)
On Error GoTo errhandler
If L1.ListCount = 0 Then Exit Sub
L2.AddItem L1.Text
L1.RemoveItem L1.ListIndex
errhandler:
Exit Sub
End Sub

'Move a item from a list up
Sub MoveUp(L As ListBox)
On Error GoTo errhandler
Dim OldIndex As Integer, OldCaption As String
If L.ListCount = 0 Then Exit Sub
If L.ListCount = 1 Then Exit Sub
If L.ListIndex = 0 Then Exit Sub
OldIndex = L.ListIndex
OldCaption = L.Text
L.RemoveItem L.ListIndex
L.AddItem OldCaption, Val(OldIndex) - 1
L.ListIndex = Val(OldIndex) - 1
errhandler:
Exit Sub
End Sub

'Move a item from a list down
Sub MoveDown(L As ListBox)
On Error GoTo errhandler
Dim OldIndex As Integer, OldCaption As String
If L.ListCount = 0 Then Exit Sub
If L.ListCount = 1 Then Exit Sub
If L.ListIndex = Val(L.ListCount) - 1 Then Exit Sub
OldIndex = L.ListIndex
OldCaption = L.Text
L.RemoveItem L.ListIndex
L.AddItem OldCaption, Val(OldIndex) + 1
L.ListIndex = Val(OldIndex) + 1
errhandler:
Exit Sub
End Sub

'Change the text of a item
Sub ChangeText(L As ListBox)
On Error GoTo errhandler
Dim OldIndex As Integer, NewCaption As String
If L.ListCount = 0 Then Exit Sub
OldIndex = L.ListIndex
NewCaption = InputBox("Change to what?", "Change", L.Text)
If NewCaption = "" Then Exit Sub
L.RemoveItem L.ListIndex
L.AddItem NewCaption, OldIndex
L.ListIndex = OldIndex
errhandler:
Exit Sub
End Sub

'Change the text of a item, only you specify it
Sub ChangeText2(L As ListBox, NewText As String)
On Error GoTo errhandler
Dim OldIndex As Integer
If L.ListCount = 0 Then Exit Sub
OldIndex = L.ListIndex
L.RemoveItem L.ListIndex
L.AddItem NewText, OldIndex
L.ListIndex = OldIndex
errhandler:
Exit Sub
End Sub

'-------------------------------
'The Updated part of this module
'-------------------------------
'Copy a selected item
Sub CopyItem(L1 As ListBox)
On Error GoTo errhandler
L1.AddItem L1.Text
errhandler:
Exit Sub
End Sub

'This is to number all items from 1 to ?
Sub NumberItems(L1 As ListBox)
On Error GoTo errhandler
Dim OldCaption As String, OldIndex As Integer
If L1.ListCount = 0 Then Exit Sub
L1.ListIndex = 0
For i = 1 To L1.ListCount
OldCaption = L1.Text
OldIndex = L1.ListIndex
L1.RemoveItem L1.ListIndex
L1.AddItem i & ". " & OldCaption, OldIndex
L1.ListIndex = i
Next i
errhandler:
Exit Sub
End Sub

'This is to switch to items between two lists
Sub SwitchItems(L1 As ListBox, L2 As ListBox)

Dim OldCaption1 As String, OldCaption2 As String, OldIndex1 As Integer, OldIndex2 As Integer
If L1.ListCount = 0 Then Exit Sub
If L2.ListCount = 0 Then Exit Sub
OldCaption1 = L1.Text
OldCaption2 = L2.Text
OldIndex1 = L1.ListIndex
OldIndex2 = L2.ListIndex
L1.RemoveItem L1.ListIndex
L2.RemoveItem L2.ListIndex
L2.AddItem OldCaption1, OldIndex1
L1.AddItem OldCaption2, OldIndex2
L1.ListIndex = OldIndex2
L2.ListIndex = OldIndex1
errhandler:
Exit Sub
End Sub

'This is to copy 1 item from a list to the other
Sub CopyItemToList(L1 As ListBox, L2 As ListBox)
On Error GoTo errhandler
If L1.ListCount = 0 Then Exit Sub
L2.AddItem L1.Text
errhandler:
Exit Sub
End Sub

'This is to copy a whole list to the other
Sub CopyList(L1 As ListBox, L2 As ListBox)
On Error GoTo errhandler
If L1.ListCount = 0 Then Exit Sub
L1.ListIndex = 0
For i = 1 To L1.ListCount
L2.AddItem L1.Text
L1.ListIndex = i
Next i
L1.ListIndex = 0
errhandler:
Exit Sub
End Sub

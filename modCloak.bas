Attribute VB_Name = "modCloak"
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Sub HideWindows(Optional ByVal hwnd As Long)
Dim iCount As Integer, lHold As Long
With frmMain.lstSelections
    For iCount = 1 To .ListItems.Count
        ShowWindow Val(.ListItems(iCount)), 0
        SetProp Val(.ListItems(iCount)), "Cloaked", True
    Next iCount
End With
SaveList
End Sub

Public Sub ShowWindows()
Dim iCount As Integer
With frmMain.lstSelections
    For iCount = 1 To .ListItems.Count
        ShowWindow Val(.ListItems(iCount)), 1
        RemoveProp Val(.ListItems(iCount)), "Cloaked"
    Next iCount
End With
End Sub

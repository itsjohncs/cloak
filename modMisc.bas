Attribute VB_Name = "modMisc"
Private Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Function GetOwner(ByVal hwnd As Long) As Long
Dim lHold As Long
Do Until (hwnd = 0)
    lHold = hwnd
    hwnd = GetParent(hwnd)
Loop
GetOwner = lHold
End Function
Public Function GetWindowTextEx(ByVal hwnd As Long) As String
Dim sHold As String
sHold = Space(500)
GetWindowText hwnd, sHold, LenB(sHold)
GetWindowTextEx = Trim(sHold)
End Function
Public Function GetClassNameEx(ByVal hwnd As String) As String
Dim sHold As String * 255, iHold As Integer
iHold = GetClassName(hwnd, sHold, 255)
GetClassNameEx = Left(sHold, iHold)
End Function
Public Sub SaveList(Optional ByVal Path As String)
Dim lHold As Long
If (Path = "") Then Path = App.Path & "\Bin\Backup.BHL"
Open Path For Random As #1 Len = Len(lHold)
    For iCount = 1 To frmMain.lstSelections.ListItems.Count
        lHold = Val(frmMain.lstSelections.ListItems(iCount))
        Put #1, , lHold
    Next iCount
Close #1
End Sub
Public Sub LoadList(Optional ByVal Path As String)
Dim lHold As Long
If (Path = "") Then Path = App.Path & "\Bin\Backup.BHL"
Open Path For Random As #1 Len = Len(lHold)
    Do
        Get #1, , lHold
        If EOF(1) Then Exit Do
        If Not frmMain.IsListed(lHold) Then frmMain.lstSelections.ListItems.Add(, , "&H" & Hex(lHold)).SubItems(1) = GetWindowTextEx(lHold)
    Loop
Close #1
End Sub

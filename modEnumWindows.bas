Attribute VB_Name = "modEnumWindows"
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function EnumChildWindows Lib "user32.dll" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public sWindowHold As String, sClass As String, lWindowCount As Long, lFoundWindowCount As Long
Dim lAddress As Long
Public Function WindowsEnum(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
WindowsEnum = True
If (Asc(GetWindowTextEx(hwnd)) = 0) Then
    sWindowHold = sWindowHold & hwnd & ":": Exit Function
End If
With frmSelList.lstWindows.ListItems.Add(, , "&H" & Hex(hwnd))
    .SubItems(1) = GetWindowTextEx(hwnd)
End With
End Function
Public Function WindowsEnum2(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
Dim iCount As Integer
WindowsEnum2 = True
If (hwnd = frmSelType.hwnd) Then Exit Function
With frmSelType.lstWindows.Nodes
    For iCount = 1 To .Count
        If (.Item(iCount).Tag = GetClassNameEx(hwnd)) Then .Add(iCount, tvwChild, , GetWindowTextEx(hwnd)).Tag = hwnd: Exit Function
    Next iCount
End With
End Function
Public Function WindowsEnum3(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
Dim iCount As Integer
WindowsEnum3 = True
If (hwnd = frmSelType.hwnd) Or frmMain.IsListed(hwnd) Then Exit Function
If (GetClassNameEx(hwnd) = sClass) Then frmMain.lstSelections.ListItems.Add(, , "&H" & Hex(hwnd)).SubItems(1) = GetWindowTextEx(hwnd)
End Function
Public Function WindowsEnum4(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
WindowsEnum4 = True
lWindowCount = lWindowCount + 1
If GetProp(hwnd, "Cloaked") And Not frmMain.IsListed(hwnd) Then frmMain.lstSelections.ListItems.Add(, , "&H" & Hex(hwnd)).SubItems(1) = GetWindowTextEx(hwnd): lFoundWindowCount = lFoundWindowCount + 1
lAddress = lEcho(AddressOf WindowsEnum5)
EnumChildWindows hwnd, AddressOf WindowsEnum5, ByVal 0&
DoEvents
End Function
Private Function lEcho(ByVal Num As Long) As Long
lEcho = Num
End Function
Public Function WindowsEnum5(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
WindowsEnum5 = True
lWindowCount = lWindowCount + 1
If GetProp(hwnd, "Cloaked") And Not frmMain.IsListed(hwnd) Then frmMain.lstSelections.ListItems.Add(, , "&H" & Hex(hwnd)).SubItems(1) = GetWindowTextEx(hwnd): lFoundWindowCount = lFoundWindowCount + 1
EnumChildWindows hwnd, lAddress, ByVal 0&
End Function

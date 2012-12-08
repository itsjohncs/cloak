Attribute VB_Name = "modHotkeys"
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Private Declare Function VkKeyScanW Lib "user32" (ByVal cChar As Integer) As Integer
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private hHook As Long 'The variable that holds the return value for setting the hook
Private rHotKeys() As Byte, rHotKeysStr As String, HotKeysL1 As Integer, HotKeyLens() As Integer
Public Function KeyboardProc(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim bFalse As Boolean, iCount As Integer, iCount2 As Integer
If idHook < 0 Then
    KeyboardProc = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
Else
    If IsSet Then
        For iCount = 0 To HotKeysL1
            For iCount2 = 0 To HotKeyLens(iCount)
               If Not GetKeyState(rHotKeys(iCount, iCount2)) And &HF0000000 Then
                    bFalse = True
                    Exit For
                End If
            Next iCount2
            If Not bFalse Then
                UserCode iCount
            End If
            bFalse = False
        Next iCount
    End If
    KeyboardProc = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
End If
End Function

Public Function SetHotKeys(ByVal HotKeys As String) As Boolean
Dim iCount As Integer, sHold() As String, sHoldPrime() As String, iCount2 As Integer
If (HotKeys = "") Then Exit Function 'If they didn't specify hotkeys, DIE!!!!
ReDim rHotKeys(0) 'Clear the rHotKeys varible
sHoldPrime = Split(HotKeys, "&") 'Split it into groups
ReDim rHotKeys(UBound(sHoldPrime), 0 To 10)
ReDim HotKeyLens(UBound(sHoldPrime)) As Integer
For iCount = 0 To UBound(sHoldPrime) 'Run through the groups
    sHold = Split(sHoldPrime(iCount), "+") 'Get each hotkey
    HotKeyLens(iCount) = UBound(sHold)
    For iCount2 = 0 To UBound(sHold)
        Select Case LCase(sHold(iCount2))
            Case "{enter}": rHotKeys(iCount, iCount2) = vbKeyExecute
            Case "{shift}": rHotKeys(iCount, iCount2) = vbKeyShift
            Case "{ctrl}": rHotKeys(iCount, iCount2) = 17
            Case "{alt}": rHotKeys(iCount, iCount2) = 164
            Case "{tab}": rHotKeys(iCount, iCount2) = vbKeyTab
            Case "{home}": rHotKeys(iCount, iCount2) = vbKeyHome
            Case "{end}": rHotKeys(iCount, iCount2) = vbKeyEnd
            Case "{insert}": rHotKeys(iCount, iCount2) = vbKeyInsert
            Case "{delete}": rHotKeys(iCount, iCount2) = vbKeyDelete
            Case "{backspace}": rHotKeys(iCount, iCount2) = 8
            Case "{f1}": rHotKeys(iCount, iCount2) = vbKeyF1
            Case "{f2}": rHotKeys(iCount, iCount2) = vbKeyF2
            Case "{f3}": rHotKeys(iCount, iCount2) = vbKeyF3
            Case "{f4}": rHotKeys(iCount, iCount2) = vbKeyF4
            Case "{f5}": rHotKeys(iCount, iCount2) = vbKeyF5
            Case "{f6}": rHotKeys(iCount, iCount2) = vbKeyF6
            Case "{f7}": rHotKeys(iCount, iCount2) = vbKeyF7
            Case "{f8}": rHotKeys(iCount, iCount2) = vbKeyF8
            Case "{f9}": rHotKeys(iCount, iCount2) = vbKeyF9
            Case "{f10}": rHotKeys(iCount, iCount2) = vbKeyF10
            Case "{f11}": rHotKeys(iCount, iCount2) = vbKeyF11
            Case "{f12}": rHotKeys(iCount, iCount2) = vbKeyF12
            Case "{f13}": rHotKeys(iCount, iCount2) = vbKeyF13
            Case "{f14}": rHotKeys(iCount, iCount2) = vbKeyF14
            Case "{f15}": rHotKeys(iCount, iCount2) = vbKeyF15
            Case "{f16}": rHotKeys(iCount, iCount2) = vbKeyF16
            Case "{up}": rHotKeys(iCount, iCount2) = vbKeyUp
            Case "{down}": rHotKeys(iCount, iCount2) = vbKeyDown
            Case "{right}": rHotKeys(iCount, iCount2) = vbKeyRight
            Case "{left}": rHotKeys(iCount, iCount2) = vbKeyLeft
            Case "{printscreen}": rHotKeys(iCount, iCount2) = 44
            Case "{pageup}": rHotKeys(iCount, iCount2) = vbKeyPageUp
            Case "{pagedown}": rHotKeys(iCount, iCount2) = vbKeyPageDown
            Case "{escape}": rHotKeys(iCount, iCount2) = vbKeyEscape
            Case "{num 1}": rHotKeys(iCount, iCount2) = vbKeyNumpad1
            Case "{num 2}": rHotKeys(iCount, iCount2) = vbKeyNumpad2
            Case "{num 3}": rHotKeys(iCount, iCount2) = vbKeyNumpad3
            Case "{num 4}": rHotKeys(iCount, iCount2) = vbKeyNumpad4
            Case "{num 5}": rHotKeys(iCount, iCount2) = vbKeyNumpad5
            Case "{num 6}": rHotKeys(iCount, iCount2) = vbKeyNumpad6
            Case "{num 7}": rHotKeys(iCount, iCount2) = vbKeyNumpad7
            Case "{num 8}": rHotKeys(iCount, iCount2) = vbKeyNumpad8
            Case "{num 9}": rHotKeys(iCount, iCount2) = vbKeyNumpad9
            Case "{num 0}": rHotKeys(iCount, iCount2) = vbKeyNumpad0
            Case "{numlock}": rHotKeys(iCount, iCount2) = vbKeyNumlock
            Case "{plus}": rHotKeys(iCount, iCount2) = 107
            Case "{lstartmenu}": rHotKeys(iCount, iCount2) = 91
            Case "{rstartmenu}": rHotKeys(iCount, iCount2) = 92
            Case "{space}": rHotKeys(iCount, iCount2) = 32
            Case "{scrolllock}": rHotKeys(iCount, iCount2) = 145
            Case "{pausebreak}": rHotKeys(iCount, iCount2) = 19
            Case "'": rHotKeys(iCount, iCount2) = 222
            Case "*": rHotKeys(iCount, iCount2) = 106
        End Select
        If (rHotKeys(iCount, iCount2) = 0 And Asc(sHold(iCount2)) > 32) Then rHotKeys(iCount, iCount2) = KeyCode(sHold(iCount2))
    Next iCount2
Next iCount
HotKeysL1 = UBound(sHoldPrime)
rHotKeysStr = HotKeys
End Function

Public Function GetHotKeys() As String
GetHotKeys = rHotKeysStr
End Function

Private Function IsSet() As Boolean
On Error GoTo Wrong
a = UBound(rHotKeys)
IsSet = True
Exit Function
Wrong:
End Function

Private Function KeyCode(ByVal sChar As String) As KeyCodeConstants
'Heavily compressed by me but I got the original from some code dump.
'Can't rememebr which though, :/
Dim bNt As Boolean, iKeyCode As Integer, B() As Byte, iKey As Integer, vKey As KeyCodeConstants, ishift As ShiftConstants
bNt = ((GetVersion() And &H80000000) = 0)
If (bNt) Then
    B = sChar
    CopyMemory iKey, B(0), 2
    iKeyCode = VkKeyScanW(iKey)
Else
    B = StrConv(sChar, vbFromUnicode)
    iKeyCode = VkKeyScan(B(0))
End If
KeyCode = (iKeyCode And &HFF&)
End Function

Public Sub SetHotHook()
If Not IsSet Then
    MsgBox "Hotkeys not yet set.", vbOKOnly, "BH Hotkeys" 'If you do not want an error message, then remove this line, and ONLY this line. No neighboring lines. Except for the other error message if you want to get rid of that too
    Exit Sub
End If
If (hHook <> 0) Then
    'MsgBox "Hook already set.", vbOKOnly, "BH Hotkeys" 'If you do not want an error message, then remove this line, and ONLY this line. No neighboring lines. Except for the other error message if you want to get rid of that too
    Exit Sub
End If
hHook = SetWindowsHookEx(13, AddressOf KeyboardProc, App.hInstance, 0)
End Sub

Public Sub RemHotHook()
On Error Resume Next
UnhookWindowsHookEx hHook
hHook = 0
End Sub

Private Sub UserCode(ByVal KeyNum As Integer)
'Put your code in here. Only here. No where else
If (KeyNum = 0) Then
    HideWindows
    ShowWindow frmMain.hwnd, 0
    App.TaskVisible = False
ElseIf (KeyNum = 1) Then
    ShowWindow frmMain.hwnd, 1
    App.TaskVisible = True
End If
End Sub

Public Function GetDownKeys(Optional Limit As Integer = -1) As String
On Error Resume Next
Dim iCount As Integer, sHold As String, iCount2 As Integer
For iCount = 3 To 255
    If GetKeyState(iCount) And &HF0000000 Then
        sHold = GetDownKeys
        If (iCount >= 112 And iCount <= 127) Then 'The function keys (F1-F16)
            GetDownKeys = GetDownKeys & "{F" & Trim(Str(iCount - 111)) & "}+"
        ElseIf (iCount >= 65 And iCount <= 90) Then 'Letters
            GetDownKeys = GetDownKeys & Chr(iCount) & "+"
        ElseIf (iCount >= 96 And iCount <= 105) Then 'Numpad numbers
            GetDownKeys = GetDownKeys & "{Num " & Trim(Str(iCount - 96)) & "}" & "+"
        ElseIf (iCount = vbKeyUp) Then 'Up Arrow Key
            GetDownKeys = GetDownKeys & "{Up}+"
        ElseIf (iCount = vbKeyLeft) Then 'Left Arrow Key
            GetDownKeys = GetDownKeys & "{Left}+"
        ElseIf (iCount = vbKeyDown) Then 'Down Arrow Key
            GetDownKeys = GetDownKeys & "{Down}+"
        ElseIf (iCount = vbKeyRight) Then 'Right Arrow Key
            GetDownKeys = GetDownKeys & "{Right}+"
        ElseIf (iCount >= 48 And iCount <= 57) Then 'Regular old numbers
            GetDownKeys = GetDownKeys & Chr(iCount) & "+"
        ElseIf (iCount = vbKeyTab) Then
            GetDownKeys = GetDownKeys & "{Tab}+"
        ElseIf (iCount = vbKeyShift) Then
            GetDownKeys = GetDownKeys & "{Shift}+"
        ElseIf (iCount = vbKeyCapital) Then
            GetDownKeys = GetDownKeys & "{Caps}+"
        ElseIf (iCount = vbKeyControl) Then
            GetDownKeys = GetDownKeys & "{Ctrl}+"
        ElseIf (iCount = 164) Then 'Alt
            GetDownKeys = GetDownKeys & "{Alt}+"
        ElseIf (iCount = vbKeyEscape) Then
            GetDownKeys = GetDownKeys & "{Esc}+"
        ElseIf (iCount = vbKeyDelete) Then
            GetDownKeys = GetDownKeys & "{Del}+"
        ElseIf (iCount = vbKeyPageUp) Then
            GetDownKeys = GetDownKeys & "{PageUp}+"
        ElseIf (iCount = vbKeyPageDown) Then
            GetDownKeys = GetDownKeys & "{PageDown}+"
        ElseIf (iCount = vbKeyHome) Then
            GetDownKeys = GetDownKeys & "{Home}+"
        ElseIf (iCount = vbKeyEnd) Then
            GetDownKeys = GetDownKeys & "{End}+"
        ElseIf (iCount = vbKeyInsert) Then
            GetDownKeys = GetDownKeys & "{Insert}+"
        ElseIf (iCount = vbKeyBack) Then
            GetDownKeys = GetDownKeys & "{BackSpace}+"
        ElseIf (iCount = 13) Then
            GetDownKeys = GetDownKeys & "{Enter}+"
        ElseIf (iCount = 219) Then
            GetDownKeys = GetDownKeys & "[+"
        ElseIf (iCount = 220) Then
            GetDownKeys = GetDownKeys & "\+"
        ElseIf (iCount = 221) Then
            GetDownKeys = GetDownKeys & "]+"
        ElseIf (iCount = 186) Then
            GetDownKeys = GetDownKeys & ";+"
        ElseIf (iCount = 222) Then
            GetDownKeys = GetDownKeys & "'+"
        ElseIf (iCount = 188) Then
            GetDownKeys = GetDownKeys & ",+"
        ElseIf (iCount = 190) Then
            GetDownKeys = GetDownKeys & ".+"
        ElseIf (iCount = 191) Then
            GetDownKeys = GetDownKeys & "/+"
        ElseIf (iCount = 110) Then
            GetDownKeys = GetDownKeys & ".+"
        ElseIf (iCount = 111) Then
            GetDownKeys = GetDownKeys & "/+"
        ElseIf (iCount = vbkeyequal) Then
            GetDownKeys = GetDownKeys & "*+"
        ElseIf (iCount = 109) Then
            GetDownKeys = GetDownKeys & "-+"
        ElseIf (iCount = 187) Then
            GetDownKeys = GetDownKeys & "=+"
        ElseIf (iCount = 192) Then
            GetDownKeys = GetDownKeys & "`+"
        ElseIf (iCount = 189) Then
            GetDownKeys = GetDownKeys & "-+"
        ElseIf (iCount = 107) Then
            GetDownKeys = GetDownKeys & "{Plus}+"
        ElseIf (iCount = 144) Then
            GetDownKeys = GetDownKeys & "{NumLock}+"
        ElseIf (iCount = 91) Then
            GetDownKeys = GetDownKeys & "{LStartMenu}+"
        ElseIf (iCount = 92) Then
            GetDownKeys = GetDownKeys & "{RStartMenu}+"
        ElseIf (iCount = 32) Then
            GetDownKeys = GetDownKeys & "{Space}+"
        ElseIf (iCount = 145) Then
            GetDownKeys = GetDownKeys & "{ScrollLock}+"
        ElseIf (iCount = 19) Then
            GetDownKeys = GetDownKeys & "{PauseBreak}+"
        ElseIf (iCount = 106) Then
            GetDownKeys = GetDownKeys & "*+"
        End If
        If (Limit <> -1) Then
            If (sHold <> GetDownKeys) Then iCount2 = iCount2 + 1
            If (iCount2 >= Limit) Then Exit For
        End If
    End If
Next iCount
GetDownKeys = Left(GetDownKeys, Len(GetDownKeys) - 1)
End Function

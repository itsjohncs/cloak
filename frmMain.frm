VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BH Cloak"
   ClientHeight    =   4680
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cloaking Tools"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   4215
      Begin BHCloak.chameleonButton btnHide 
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Hide All Selected Windows"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin BHCloak.chameleonButton btnShow 
         Height          =   495
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Show All Selected Windows"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin BHCloak.chameleonButton btnHotkeys 
         Height          =   495
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Set Hotkeys"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Selection Tool"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   4215
      Begin BHCloak.chameleonButton btnSelDD 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Select Window (Drag && Drop)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin BHCloak.chameleonButton btnSelLst 
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Select Window (From List)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin BHCloak.chameleonButton btnSelType 
         Height          =   495
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Select Window (By Type)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":008C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Selected Windows"
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.Timer tmrCheck 
         Interval        =   1000
         Left            =   2400
         Top             =   960
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   1440
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   "BH Cloak Log (*.BHL)|*.BHL"
      End
      Begin VB.Timer tmrRefresh 
         Interval        =   2000
         Left            =   1920
         Top             =   960
      End
      Begin VB.CheckBox chkRefresh 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Auto-Refresh Captions"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin MSComctlLib.ListView lstSelections 
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3201
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "hWnd"
            Object.Width           =   1589
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Caption"
            Object.Width           =   5382
         EndProperty
      End
      Begin BHCloak.chameleonButton btnRefresh 
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Refresh"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":00A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mClear 
         Caption         =   "Clear List"
      End
      Begin VB.Menu mSave 
         Caption         =   "Save..."
      End
      Begin VB.Menu mOpen 
         Caption         =   "Open..."
      End
      Begin VB.Menu mD0 
         Caption         =   "-"
      End
      Begin VB.Menu mRecover2 
         Caption         =   "Recover All Cloaked Windows"
      End
      Begin VB.Menu mRecover 
         Caption         =   "Recover Last List"
      End
      Begin VB.Menu mD1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mSettings 
      Caption         =   "Settings"
      Begin VB.Menu mHotKeys 
         Caption         =   "Set HotKeys"
      End
      Begin VB.Menu mToggleHotkeys 
         Caption         =   "Toggle Hotkeys"
         Checked         =   -1  'True
      End
      Begin VB.Menu mD2 
         Caption         =   "-"
      End
      Begin VB.Menu mGetOwner 
         Caption         =   "Toggle Owner Loop"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "Help"
      Begin VB.Menu mReadMe 
         Caption         =   "Help File"
      End
      Begin VB.Menu mD4 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mCaption 
         Caption         =   "Change Caption"
      End
      Begin VB.Menu mTranslucency 
         Caption         =   "Set Translucency"
      End
      Begin VB.Menu mClose 
         Caption         =   "Close Window"
      End
      Begin VB.Menu mD5 
         Caption         =   "-"
      End
      Begin VB.Menu mHide 
         Caption         =   "Hide Window"
      End
      Begin VB.Menu mShow 
         Caption         =   "Show Window"
      End
      Begin VB.Menu mD6 
         Caption         =   "-"
      End
      Begin VB.Menu mRemove 
         Caption         =   "Remove From List"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetVersion Lib "kernel32.dll" () As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Sub btnHide_Click()
HideWindows
End Sub

Private Sub btnHotkeys_Click()
frmHotkeys.Show
End Sub

Private Sub btnRefresh_Click()
tmrRefresh_Timer
End Sub

Private Sub btnSelDD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnSelDD.MousePointer = 2
End Sub
Private Sub btnSelDD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tHold As POINTAPI, lHold As Long
btnSelDD.MousePointer = 0
GetCursorPos tHold
lHold = IIf(mGetOwner.Checked, GetOwner(WindowFromPoint(tHold.X, tHold.Y)), WindowFromPoint(tHold.X, tHold.Y))
If IsListed(lHold) Then Exit Sub
With lstSelections.ListItems.Add(, , "&H" & Hex(lHold))
    .SubItems(1) = GetWindowTextEx(lHold)
End With
End Sub
Function IsListed(ByVal hWnd As Long) As Boolean
Dim iCount As Integer
For iCount = 1 To lstSelections.ListItems.Count
    If (lstSelections.ListItems(iCount).Text = "&H" & Hex(hWnd)) Then IsListed = True: Exit Function
Next iCount
End Function

Private Sub btnSelLst_Click()
frmSelList.Show
End Sub

Private Sub btnSelType_Click()
frmSelType.Show
End Sub

Private Sub btnShow_Click()
ShowWindows
End Sub

Private Sub chkRefresh_Click()
tmrRefresh = CBool(chkRefresh)
End Sub

Private Function IsExist(ByVal sPath As String) As Boolean
If (Dir(App.Path & sPath, vbNormal Or vbDirectory) = "") Then Exit Function Else IsExist = True
End Function

Private Sub Form_Load()
Dim sHold As String, iHold1 As Integer, iHold2 As Integer, iHold3 As Integer
If Not (IsExist("\Bin") And IsExist("\Bin\Settings.ini") And IsExist("\Bin\Window Types.ini")) Then MsgBox "Could not locate all needed files.", vbOKOnly + vbCritical, "Error": Unload Me
If ((GetVersion And &HFFFF&) Mod 256 < 5) Then mTranslucency = False
Open App.Path & "\Bin\Settings.ini" For Input As #1
    Input #1, iHold1, iHold2, iHold3, sHold
    chkRefresh = iHold1
    mToggleHotkeys.Checked = iHold2
    mGetOwner = iHold3
Close #1
CD.InitDir = App.Path
If (sHold = "") Then sHold = "{F6}&{Ctrl}+B+H"
SetHotKeys sHold
If mToggleHotkeys.Checked Then SetHotHook
End Sub

Private Sub Form_Unload(Cancel As Integer)
RemHotHook
If IsExist("\Bin\Settings.ini") Then
    Open App.Path & "\Bin\Settings.ini" For Output As #1
        Write #1, CInt(chkRefresh), -CInt(mToggleHotkeys.Checked), -CInt(mGetOwner.Checked), GetHotKeys
    Close #1
End If
For Each Form In Forms
    Unload Form
Next Form
End
End Sub

Private Sub lstSelections_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyDelete) Then lstSelections.ListItems.Remove lstSelections.SelectedItem.Index
End Sub

Private Sub lstSelections_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Skip
If (Button = 2 And lstSelections.HitTest(X, Y)) Then
    PopupMenu Me.mPopUp
End If
Skip:
End Sub

Private Sub mAbout_Click()
frmAbout.Show
End Sub

Private Sub mCaption_Click()
Dim sHold As String
sHold = InputBox("What would you like the caption to be changed to?", "Change Caption")
If (StrPtr(sHold) <> 0) Then SetWindowText lstSelections.SelectedItem, sHold
End Sub

Private Sub mClear_Click()
lstSelections.ListItems.Clear
End Sub

Private Sub mClose_Click()
PostMessage Val(lstSelections.SelectedItem), &H10, 0, 0
End Sub

Private Sub mExit_Click()
Unload Me
End Sub

Private Sub mGetOwner_Click()
mGetOwner.Checked = Not mGetOwner.Checked
End Sub

Private Sub mHide_Click()
ShowWindow Val(lstSelections.SelectedItem), 0
SetProp Val(lstSelections.SelectedItem), "Cloaked", True
End Sub



Private Sub mHotKeys_Click()
frmHotkeys.Show
End Sub

Private Sub mOpen_Click()
On Error GoTo Canceled
Dim iCount As Integer, lHold As Long
CD.DialogTitle = "Open Window List"
CD.ShowOpen
LoadList CD.FileName
Canceled:
End Sub

Private Sub mReadMe_Click()
ShellExecute 0, "Open", App.Path & "\Bin\ReadMe.htm", 0&, 0&, 1
End Sub

Private Sub mRecover_Click()
LoadList
End Sub

Private Sub mRecover2_Click()
On Error Resume Next
Unload frmRecover
frmRecover.Show
End Sub

Private Sub mRemove_Click()
lstSelections.ListItems.Remove lstSelections.SelectedItem.Index
End Sub

Private Sub mSave_Click()
On Error GoTo Canceled
Dim iCount As Integer, lHold As Long
CD.DialogTitle = "Save Window List"
CD.ShowSave
If (Dir(CD.FileName) <> "") Then If (MsgBox(CD.FileTitle & " already exists." & vbNewLine & "Would you like to overwrite it?", vbYesNo + vbQuestion, "Overwrite?") = vbNo) Then Exit Sub Else Kill CD.FileName
SaveList CD.FileName
Exit Sub
Canceled:
End Sub

Private Sub mShow_Click()
ShowWindow Val(lstSelections.SelectedItem), 1
RemoveProp Val(lstSelections.SelectedItem), "Cloaked"
End Sub

Private Sub mToggleHotkeys_Click()
mToggleHotkeys.Checked = Not mToggleHotkeys.Checked
If mToggleHotkeys.Checked Then SetHotHook Else RemHotHook
End Sub

Private Sub mTranslucency_Click()
frmTranslucency.Show
End Sub

Private Sub tmrCheck_Timer()
Dim iCount As Integer
For iCount = 1 To lstSelections.ListItems.Count
    With lstSelections.ListItems(iCount)
        If IsWindow(Val(.Text)) Then
            .ForeColor = vbBlack
        Else
            .ForeColor = vbRed
            .SubItems(1) = ""
        End If
    End With
Next iCount
End Sub

Private Sub tmrRefresh_Timer()
Dim iCount As Integer
For iCount = 1 To lstSelections.ListItems.Count
    lstSelections.ListItems(iCount).SubItems(1) = GetWindowTextEx(lstSelections.ListItems(iCount))
Next iCount
End Sub

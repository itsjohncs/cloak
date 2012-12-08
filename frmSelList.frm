VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selection Tool"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin MSComctlLib.ListView lstWindows 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3201
      SortKey         =   1
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
         Object.Width           =   4942
      EndProperty
   End
   Begin BHCloak.chameleonButton btnSelect 
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Select"
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
      BCOL            =   15591915
      BCOLO           =   15591915
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSelList.frx":0000
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
Attribute VB_Name = "frmSelList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Sub btnSelect_Click()
If frmMain.IsListed(Val(lstWindows.SelectedItem.Text)) Then
    MsgBox "That window is already selected.", vbCritical, "Error"
    Exit Sub
End If
With frmMain.lstSelections.ListItems.Add(, , lstWindows.SelectedItem.Text)
    .SubItems(1) = lstWindows.SelectedItem.SubItems(1)
    Unload Me
End With
End Sub

Private Sub Form_Load()
Dim aHold() As String, iCount As Integer
lstWindows.ListItems.Clear
EnumWindows AddressOf WindowsEnum, ByVal 0&
DoEvents
lstWindows.Sorted = True
lstWindows.Sorted = False
aHold = Split(sWindowHold, ":")
For iCount = 1 To UBound(aHold)
    With lstWindows.ListItems.Add(, , "&H" & Hex(Val(aHold(iCount))))
        .SubItems(1) = GetWindowTextEx(Val(aHold(iCount)))
    End With
Next iCount
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub lstWindows_KeyPress(KeyAscii As Integer)
If (KeyAscii = vbKeyReturn) Then btnSelect_Click
End Sub

Private Sub txtSearch_Change()
Dim iCount As Integer
For iCount = 1 To lstWindows.ListItems.Count
    If (LCase(Left(lstWindows.ListItems(iCount).SubItems(1), Len(txtSearch))) = LCase(txtSearch)) Then
        lstWindows.ListItems(iCount).EnsureVisible
        lstWindows.ListItems(iCount).Selected = True
        Exit Sub
    End If
Next iCount
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If (KeyAscii = vbKeyReturn) Then btnSelect_Click
End Sub

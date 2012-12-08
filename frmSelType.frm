VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelType 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selection Tool"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin BHCloak.chameleonButton btnSelect 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
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
      MICON           =   "frmSelType.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.TreeView lstWindows 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3625
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   4
      Scroll          =   0   'False
      SingleSel       =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmSelType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Sub btnSelect_Click()
On Error Resume Next
Dim iCount As Integer, tHold As Integer, tRoot As Integer
With lstWindows
    If Not IsNumeric(.SelectedItem.Tag) Then sClass = .SelectedItem.Tag Else sClass = .SelectedItem.Parent.Tag
End With
EnumWindows AddressOf WindowsEnum3, ByVal 0&
Unload Me
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
Dim sHold As String
Open Replace(App.Path, "\", "/") & "/Bin/Window Types.ini" For Input As #1
    Do Until EOF(1)
        Line Input #1, sHold
        lstWindows.Nodes.Add(, , , Split(sHold, ";")(1)).Tag = Split(sHold, ";")(0)
    Loop
Close #1
EnumWindows AddressOf WindowsEnum2, ByVal 0&
End Sub


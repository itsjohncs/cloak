VERSION 5.00
Begin VB.Form frmRecover 
   BorderStyle     =   0  'None
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRefresh 
      Interval        =   10
      Left            =   1080
      Top             =   240
   End
   Begin BHCloak.chameleonButton btnCancel 
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Cancel"
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
      MICON           =   "frmRecover.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblFound 
      AutoSize        =   -1  'True
      Caption         =   "0 Cloaked Windows Found"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1920
   End
   Begin VB.Label lblChecked 
      AutoSize        =   -1  'True
      Caption         =   "0 Windows Checked"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1485
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmRecover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function EnumChildWindows Lib "user32.dll" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Sub btnCancel_Click()
frmMain.Show
Unload Me
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
Me.Visible = True
frmMain.Visible = False
Me.ZOrder
lWindowCount = 0
lFoundWindowCount = 0
EnumWindows AddressOf WindowsEnum4, ByVal 0&
DoEvents
frmMain.Show
End Sub

Private Sub tmrRefresh_Timer()
lblChecked = lWindowCount & " Windows Checked"
lblFound = lFoundWindowCount & " Cloaked Windows Found"
End Sub

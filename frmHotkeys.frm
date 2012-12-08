VERSION 5.00
Begin VB.Form frmHotkeys 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set Hotkeys"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin BHCloak.chameleonButton btnSave 
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Save Changes"
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
      MICON           =   "frmHotkeys.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtShow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtHide 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Show Hotkeys:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   140
      TabIndex        =   3
      Top             =   500
      Width           =   1080
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Hide Hotkeys:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   140
      Width           =   1005
   End
End
Attribute VB_Name = "frmHotkeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSave_Click()
SetHotKeys txtHide & "&" & txtShow
Unload Me
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
txtHide = Split(GetHotKeys, "&")(0)
txtShow = Split(GetHotKeys, "&")(1)
End Sub

Private Sub txtHide_KeyDown(KeyCode As Integer, Shift As Integer)
txtHide = GetDownKeys(3)
End Sub

Private Sub txtShow_KeyDown(KeyCode As Integer, Shift As Integer)
txtShow = GetDownKeys(3)
End Sub

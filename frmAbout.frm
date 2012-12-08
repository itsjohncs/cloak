VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   5055
      Left            =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Image imgCruels 
      Height          =   465
      Left            =   2400
      Picture         =   "frmAbout.frx":0000
      Top             =   4440
      Width           =   1320
   End
   Begin VB.Label lblP2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Special Thanks to Bokiniec for Extensive Beta Testing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   6015
   End
   Begin VB.Label lblP1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BH Cloak was created by Brownhead. Best method to contact me is over MSN, mine is Brownhead622@hotmail.com."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   6015
   End
   Begin VB.Image imgSig 
      Height          =   2400
      Left            =   120
      Picture         =   "frmAbout.frx":138C
      Top             =   120
      Width           =   6000
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Unload Me
End Sub

Private Sub imgCruels_Click()
ShellExecute 0, "Open", "http://www.cruels.net", 0&, 0&, 1
End Sub

Private Sub imgSig_Click()
Unload Me
End Sub

Private Sub lblP1_Click()
Unload Me
End Sub

Private Sub lblP2_Click()
Unload Me
End Sub

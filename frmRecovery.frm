VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTranslucency 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Translucency Tool"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkEnable 
      Appearance      =   0  'Flat
      Caption         =   "Enable Transparency"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "0% Percent Translucent"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2895
      Begin MSComCtl2.FlatScrollBar scrTranslucent 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Arrows          =   65536
         LargeChange     =   10
         Max             =   100
         Orientation     =   1245185
      End
   End
End
Attribute VB_Name = "frmTranslucency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Sub chkEnable_Click()
Frame1.Enabled = chkEnable
scrTranslucent.Enabled = chkEnable
If Not chkEnable Then SetWindowLong Val(frmMain.lstSelections.SelectedItem), -20, &H40000
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
chkEnable = -CInt(CBool(GetWindowLong(Val(frmMain.lstSelections.SelectedItem), -20) And &H80000))
End Sub

Private Sub scrTranslucent_Change()
With frmMain.lstSelections
    SetWindowLong Val(.SelectedItem), -20, GetWindowLong(Val(.SelectedItem), -20) Or &H80000
    SetLayeredWindowAttributes Val(.SelectedItem), 0, Round(((100 - scrTranslucent) / 100) * 255), 2
End With
End Sub

VERSION 5.00
Begin VB.Form Spoof 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6195
   ControlBox      =   0   'False
   Icon            =   "Spoof.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleMode       =   0  'User
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   120
      Picture         =   "Spoof.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   1426
      MouseIcon       =   "Spoof.frx":1C68
      Picture         =   "Spoof.frx":1DBA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   1800
      Picture         =   "Spoof.frx":3158
      Top             =   1080
      Width           =   3525
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   5390
      Picture         =   "Spoof.frx":609E
      Top             =   780
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   -30
      Picture         =   "Spoof.frx":7A60
      Top             =   1597
      Width           =   6225
   End
End
Attribute VB_Name = "Spoof"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MouseClip.ReleaseYes
Spoof.Hide
Del.Show
End Sub

Private Sub Form_Activate()
Dim lpRect As RECT
MouseClip.FillRect Command1, lpRect
lpRect.bottom = lpRect.bottom - 18
lpRect.right = lpRect.right - 10
MouseClip.ClipYes lpRect
End Sub


VERSION 5.00
Begin VB.Form Error 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Demo.frx":0000
   ScaleHeight     =   2205
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "If you can read this propaly Click Me!"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MouseClip.ReleaseYes
End Sub

Private Sub Form_Activate()
Dim lpRect As RECT
MouseClip.FillRect Command1, lpRect
lpRect.bottom = lpRect.bottom - 18
lpRect.right = lpRect.right - 10
MouseClip.ClipYes lpRect
End Sub

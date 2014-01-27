VERSION 5.00
Begin VB.Form frm 
   BackColor       =   &H00808080&
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6345
   ControlBox      =   0   'False
   FillColor       =   &H00808080&
   ForeColor       =   &H8000000A&
   Icon            =   "frm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton fff 
      BackColor       =   &H0000C000&
      Caption         =   "X"
      Height          =   255
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   "X Mode"
      Height          =   4575
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   735
   End
   Begin VB.Timer timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4680
      Top             =   4320
   End
   Begin VB.Timer level2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   3720
   End
   Begin VB.Timer level1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   3720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "50"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Play"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Image a 
      Height          =   480
      Left            =   2760
      Picture         =   "frm.frx":0442
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "time"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   26.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   5160
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label hits 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4020
      TabIndex        =   7
      Top             =   30
      Width           =   735
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   3960
      X2              =   3960
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Smiley Catcher"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   30
      Width           =   2295
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      X1              =   0
      X2              =   6360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   4800
      X2              =   4800
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   6360
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   6360
      X2              =   6360
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   $"frm.frx":0884
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2175
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   5295
   End
   Begin VB.Label title 
      BackColor       =   &H00808080&
      Caption         =   "Smiley Catcher 2.7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label level 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level 1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   30
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   5160
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   255
      Left            =   6360
      Top             =   0
      Width           =   5535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "    Vlad p made this software in 5 hours"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6600
      TabIndex        =   11
      Top             =   30
      Width           =   5535
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a_Click()
If Command4.Tag = "X" Then GoTo xm
hits.Tag = hits.Tag + "X"
hits = Len(hits.Tag)
a.Visible = False
Exit Sub
xm:
hits.Tag = hits.Tag + "X"
hits = Len(hits.Tag)
title.Tag = Len(hits.Tag)
End Sub

Private Sub Command1_Click()
title.Tag = Text1.Text

End Sub

Private Sub Command2_Click()
title.Tag = 50
End Sub

Private Sub Command3_Click()
If Command4.Tag = "X" Then GoTo xv:
level1.Enabled = True
Command3.Visible = False
timer.Enabled = True
Exit Sub
xv:
level1.Enabled = True
Command4.Visible = False
Command3.Visible = False
End Sub

Private Sub Command4_Click()
Command4.Tag = "X"
title.Visible = False
Label1.Visible = False
Label3.Visible = False
Shape2.Visible = False
Command4.Visible = False
End Sub

Private Sub fff_Click()
End
End Sub

Private Sub Form_Load()
title.Tag = "30"
frm.Tag = "2"
End Sub


Public Function tLeft(Speed)
a.Top = a.Top - Speed
a.Left = a.Left - Speed
frm.Tag = "1"
End Function

Public Function tRight(Speed)
a.Top = a.Top - Speed
a.Left = a.Left + Speed
frm.Tag = "2"
End Function

Public Function bLeft(Speed)
a.Top = a.Top + Speed
a.Left = a.Left - Speed
frm.Tag = "3"
End Function

Public Function bRight(Speed)
a.Top = a.Top + Speed
a.Left = a.Left + Speed
frm.Tag = "4"
End Function

Private Sub level1_Timer()
Dim aTop, aLeft As Integer, dir As Integer
Dim fTop, fLeft As Integer
a.Visible = True

fTop = frm.Height
fLeft = frm.Width
aTop = a.Top
aLeft = a.Left
x = title.Tag
dir = frm.Tag
If dir = "1" Then tLeft x
If dir = "2" Then tRight x
If dir = "3" Then bLeft x
If dir = "4" Then bRight x
'Top wall
If aTop <= 270 And dir = "2" Then bRight x Else
If aTop <= 270 And dir = "1" Then bLeft x Else
'Right wall
If aLeft + 600 > fLeft And dir = "4" Then bLeft x Else
If aLeft + 600 > fLeft And dir = "2" Then tLeft x Else
'Botton wall
If aTop + 600 > fTop And dir = "3" Then tLeft x Else
If aTop + 600 > fTop And dir = "4" Then tRight x Else
'Left wall
If aLeft <= 0 And dir = "1" Then tRight x Else
If aLeft <= 0 And dir = "3" Then bRight x Else
End Sub

Private Sub level2_Timer()
Dim aTop, aLeft As Integer, dir As Integer
Dim fTop, fLeft As Integer
a.Visible = True
level.Caption = "Level 2"
fTop = frm.Height
fLeft = frm.Width
aTop = a.Top
aLeft = a.Left
x = 50
dir = frm.Tag

If dir = "1" Then tLeft x
If dir = "2" Then tRight x
If dir = "3" Then bLeft x
If dir = "4" Then bRight x
'Top wall
If aTop <= 270 And dir = "2" Then bRight x Else
If aTop <= 270 And dir = "1" Then bLeft x Else
'Right wall
If aLeft + 600 > fLeft And dir = "4" Then bLeft x Else
If aLeft + 600 > fLeft And dir = "2" Then tLeft x Else
'Botton wall
If aTop + 600 > fTop And dir = "3" Then tLeft x Else
If aTop + 600 > fTop And dir = "4" Then tRight x Else
'Left wall
If aLeft <= 0 And dir = "1" Then tRight x Else
If aLeft <= 0 And dir = "3" Then bRight x Else
End Sub

Private Sub timer_Timer()
Label3.Tag = Label3.Tag + "X"
Label3.Caption = Abs((Len(Label3.Tag) - 30))
If Label3.Caption = "0" Then EndLevel
End Sub

Public Function EndLevel()
timer.Enabled = False
level1.Enabled = False
level2.Enabled = False

If hits.Caption > 35 And hits.Caption < 80 Then
MsgBox "Level2", vbInformation + vbOKOnly, "LeVeL CoMpLeTe"
a.Top = 4200
a.Left = 2880
level2.Enabled = True
Label3.Tag = ""
timer.Enabled = True
Command3.Visible = False
Else
a.Top = 4200
a.Left = 2880
MsgBox "Nice try better luck next time.", vbOKOnly, "L o S e R"
Command3.Enabled = True
Command3.Visible = True
End If


If hits.Caption > 100 Then
MsgBox "Greate work. You have finished the game.", vbInformation + vbOKOnly, "G a M e C o M p L e T e"
a.Top = 4200
a.Left = 2880
level2.Enabled = False
timer.Enabled = False
Label3.Caption = "done"
Command3.Visible = True
level.Caption = " Finished"
Else
a.Top = 4200
a.Left = 2880
Command3.Enabled = True
level.Caption = "Second Try"
End If
End Function

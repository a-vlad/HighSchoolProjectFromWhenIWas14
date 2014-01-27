VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00EFD79E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "T3 - Tic Tac Toe v 5.76"
   ClientHeight    =   3495
   ClientLeft      =   3585
   ClientTop       =   1575
   ClientWidth     =   3615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H0000FFFF&
   ForeColor       =   &H8000000F&
   Icon            =   "TTT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TTT.frx":08CA
   ScaleHeight     =   3495
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.Timer last 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   480
      Top             =   1680
   End
   Begin VB.Timer ata 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   480
      Top             =   2760
   End
   Begin VB.Timer seti 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1560
      Top             =   2760
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2880
      Top             =   2760
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   0
      Width           =   135
   End
   Begin VB.Line Line5 
      BorderColor     =   &H004D5F84&
      BorderWidth     =   2
      X1              =   0
      X2              =   4080
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label cv 
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3180
      TabIndex        =   11
      Top             =   30
      Width           =   405
   End
   Begin VB.Label ng 
      BackStyle       =   0  'Transparent
      Caption         =   " New Game"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2140
      TabIndex        =   10
      Top             =   25
      Width           =   855
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   2
      X1              =   -120
      X2              =   4080
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   2
      X1              =   -120
      X2              =   4080
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   2
      X1              =   2400
      X2              =   2400
      Y1              =   240
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   2
      X1              =   1200
      X2              =   1200
      Y1              =   240
      Y2              =   3960
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H003E76D2&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      Height          =   255
      Left            =   2160
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By Vlad Paraschiv"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Colour"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1635
      TabIndex        =   12
      Top             =   25
      Width           =   495
   End
   Begin VB.Shape shape 
      BackColor       =   &H003E76D2&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      Height          =   255
      Left            =   3130
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H003E76D2&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      Height          =   255
      Left            =   1620
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      FillColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label t2 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1200
      TabIndex        =   8
      Top             =   240
      Width           =   1220
   End
   Begin VB.Label t3 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2400
      TabIndex        =   7
      Top             =   240
      Width           =   1245
   End
   Begin VB.Label m3 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2400
      TabIndex        =   5
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label m2 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Width           =   1220
   End
   Begin VB.Label b2 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1200
      TabIndex        =   4
      Top             =   2400
      Width           =   1220
   End
   Begin VB.Label b3 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2400
      TabIndex        =   6
      Top             =   2400
      Width           =   1245
   End
   Begin VB.Label t1 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      TabIndex        =   9
      Top             =   240
      Width           =   1220
   End
   Begin VB.Label b1 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   1220
   End
   Begin VB.Label m1 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   1220
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ata_Timer()
NewGame
Reset

t1.BackColor = &H5BC8BF
t2.BackColor = &H5BC8BF
t3.BackColor = &H5BC8BF
m1.BackColor = &H5BC8BF
m2.BackColor = &H5BC8BF
m3.BackColor = &H5BC8BF
b1.BackColor = &H5BC8BF
b2.BackColor = &H5BC8BF
b3.BackColor = &H5BC8BF

Me.Picture = LoadPicture(App.path & "\pic4.bmp")
m1.Tag = ""
t1.Tag = ""
Me.Tag = ""
b3.Tag = "g"
ata.Enabled = False
End Sub

Private Sub b1_Click()
If b1.Caption = "O" Then Exit Sub
If b1.Caption = "X" Then Exit Sub

b1.Caption = "X"
OOO
End Sub

Private Sub b2_Click()
If b2.Caption = "O" Then Exit Sub
If b2.Caption = "X" Then Exit Sub

b2.Caption = "X"
OOO
End Sub

Private Sub b3_Click()
If b3.Caption = "O" Then Exit Sub
If b3.Caption = "X" Then Exit Sub

b3.Caption = "X"
OOO
End Sub

Private Sub CloseR_Click()
End
End Sub



Private Sub Command2_Click()
NewGame
End Sub



Private Sub Command1_Click()
newPic
End Sub

Private Sub cv_Click()
End
End Sub

Private Sub Label2_Click()
Dim v
On Error GoTo restore
v = InputBox("Windows XP: Please type in colour value.For advanced users only.If incorect colour type in vbWhite.", "Windows XP Colour", "vbWhite")

t1.ForeColor = v
t2.ForeColor = v
t3.ForeColor = v
m1.ForeColor = v
m2.ForeColor = v
m3.ForeColor = v
b1.ForeColor = v
b2.ForeColor = v
b3.ForeColor = v
Exit Sub
restore:
t1.ForeColor = vbWhite
t2.ForeColor = vbWhite
t3.ForeColor = vbWhite
m1.ForeColor = vbWhite
m2.ForeColor = vbWhite
m3.ForeColor = vbWhite
b1.ForeColor = vbWhite
b2.ForeColor = vbWhite
b3.ForeColor = vbWhite

End Sub






Private Sub Label3_DblClick()
newPic
End Sub

Private Sub last_Timer()
NewGame
Reset

t1.BackColor = &HD033A2
t2.BackColor = &HD033A2
t3.BackColor = &HD033A2
m1.BackColor = &HD033A2
m2.BackColor = &HD033A2
m3.BackColor = &HD033A2
b1.BackColor = &HD033A2
b2.BackColor = &HD033A2
b3.BackColor = &HD033A2

Me.Picture = LoadPicture(App.path & "\pic5.bmp")
m1.Tag = ""
t1.Tag = ""
b3.Tag = ""
Me.Tag = ""
last.Enabled = False
End Sub

Private Sub m1_Click()
If m1.Caption = "O" Then Exit Sub
If m1.Caption = "X" Then Exit Sub

m1.Caption = "X"
OOO
End Sub

Private Sub m2_Click()
If m2.Caption = "O" Then Exit Sub
If m2.Caption = "X" Then Exit Sub

m2.Caption = "X"
OOO
End Sub

Private Sub m3_Click()
If m3.Caption = "O" Then Exit Sub
If m3.Caption = "X" Then Exit Sub

m3.Caption = "X"
OOO
End Sub

Private Sub ng_Click()
NewGame
Reset
End Sub

Private Sub seti_Timer()
NewGame
Reset

t1.BackColor = &H73C04E
t2.BackColor = &H73C04E
t3.BackColor = &H73C04E
m1.BackColor = &H73C04E
m2.BackColor = &H73C04E
m3.BackColor = &H73C04E
b1.BackColor = &H73C04E
b2.BackColor = &H73C04E
b3.BackColor = &H73C04E

Me.Picture = LoadPicture(App.path & "\pic3.bmp")
m1.Tag = ""
t1.Tag = ""
Me.Tag = "L3"
seti.Enabled = False
End Sub

Private Sub t1_Click()
If t1.Caption = "O" Then Exit Sub
If t1.Caption = "X" Then Exit Sub

t1.Caption = "X"
OOO
End Sub

Private Sub t2_Click()
If t2.Caption = "O" Then Exit Sub
If t2.Caption = "X" Then Exit Sub

t2.Caption = "X"
OOO
End Sub

Private Sub t3_Click()
If t3.Caption = "O" Then Exit Sub
If t3.Caption = "X" Then Exit Sub

t3.Caption = "X"
OOO
End Sub


Public Function AI()
If m1.Tag = "W" Then Exit Function
'Top row scanner
If t1.Caption = "X" And t2.Caption = "X" And t3 = "" Then t3.Caption = "O": XOScan: Exit Function
If t2.Caption = "X" And t3.Caption = "X" And t1 = "" Then t1.Caption = "O": XOScan: Exit Function
If t1.Caption = "X" And t3.Caption = "X" And t2 = "" Then t2.Caption = "O": XOScan: Exit Function
'Middle row scanner
If m1.Caption = "X" And m2.Caption = "X" And m3 = "" Then m3.Caption = "O": XOScan: Exit Function
If m2.Caption = "X" And m3.Caption = "X" And m1 = "" Then m1.Caption = "O": XOScan: Exit Function
If m1.Caption = "X" And m3.Caption = "X" And m2 = "" Then m2.Caption = "O": XOScan: Exit Function
'Bottom row scanner
If b1.Caption = "X" And b2.Caption = "X" And b3 = "" Then b3.Caption = "O": XOScan: Exit Function
If b2.Caption = "X" And b3.Caption = "X" And b1 = "" Then b1.Caption = "O": XOScan: Exit Function
If b1.Caption = "X" And b3.Caption = "X" And b2 = "" Then b2.Caption = "O": XOScan: Exit Function

'Column 1 scanner
If t1.Caption = "X" And m1.Caption = "X" And b1 = "" Then b1.Caption = "O": XOScan: Exit Function
If m1.Caption = "X" And b1.Caption = "X" And t1 = "" Then t1.Caption = "O": XOScan: Exit Function
If t1.Caption = "X" And b1.Caption = "X" And m1 = "" Then m1.Caption = "O": XOScan: Exit Function
'Column 2 scanner
If t2.Caption = "X" And m2.Caption = "X" And b2 = "" Then b2.Caption = "O": XOScan: Exit Function
If m2.Caption = "X" And b2.Caption = "X" And t2 = "" Then t2.Caption = "O": XOScan: Exit Function
If t2.Caption = "X" And b2.Caption = "X" And m2 = "" Then m2.Caption = "O": XOScan: Exit Function
'Column 3 scanner
If t3.Caption = "X" And m3.Caption = "X" And b3 = "" Then b3.Caption = "O": XOScan: Exit Function
If m3.Caption = "X" And b3.Caption = "X" And t3 = "" Then t3.Caption = "O": XOScan: Exit Function
If t3.Caption = "X" And b3.Caption = "X" And m3 = "" Then m3.Caption = "O": XOScan: Exit Function

'Diagonal 1 Scanner
If t1.Caption = "X" And b3.Caption = "X" And m2 = "" Then m2.Caption = "O": XOScan: Exit Function
If t1.Caption = "X" And m2.Caption = "X" And b3 = "" Then b3.Caption = "O": XOScan: Exit Function
If m2.Caption = "X" And b3.Caption = "X" And t1 = "" Then t1.Caption = "O": XOScan: Exit Function
'Diagonal 2 Scanner
If t3.Caption = "X" And m2.Caption = "X" And b1 = "" Then b1.Caption = "O": XOScan: Exit Function
If b1.Caption = "X" And m2.Caption = "X" And t3 = "" Then t3.Caption = "O": XOScan: Exit Function
If t3.Caption = "X" And b1.Caption = "X" And m2 = "" Then m2.Caption = "O": XOScan: Exit Function
XOScan
AIPriority
End Function

Public Function XOScan()
If t1 & t2 & t3 = "XXX" Then Call Win("X"): Exit Function
If t1 & t2 & t3 = "OOO" Then Call Win("O"): Exit Function
If m1 & m2 & m3 = "XXX" Then Call Win("X"): Exit Function
If m1 & m2 & m3 = "OOO" Then Call Win("O"): Exit Function
If b1 & b2 & b3 = "XXX" Then Call Win("X"): Exit Function
If b1 & b2 & b3 = "OOO" Then Call Win("O"): Exit Function
If t1 & m2 & b3 = "XXX" Then Call Win("X"): Exit Function
If t1 & m2 & b3 = "OOO" Then Call Win("O"): Exit Function
If t3 & m2 & b1 = "XXX" Then Call Win("X"): Exit Function
If t3 & m2 & b1 = "OOO" Then Call Win("O"): Exit Function
If t1 & m1 & b1 = "XXX" Then Call Win("X"): Exit Function
If t1 & m1 & b1 = "OOO" Then Call Win("O"): Exit Function
If t2 & m2 & b2 = "XXX" Then Call Win("X"): Exit Function
If t2 & m2 & b2 = "OOO" Then Call Win("O"): Exit Function
If t3 & m3 & b3 = "XXX" Then Call Win("X"): Exit Function
If t3 & m3 & b3 = "OOO" Then Call Win("O"): Exit Function

If Not Len(t1 + t2 + t3 + m1 + m2 + m3 + b1 + b2 + b3) = 9 Then Exit Function
Call NewGame
'msg = MsgBox("This game was a tie. Wana play again?", vbYesNo + vbInformation, "T i E g A m E")
'If msg = vbYes Then Call NewGame: Exit Function
'End
End Function

Public Function Win(side As String)
If side = "X" Then boxoff ' MsgBox ("Congratulations you win."): boxoff
If side = "O" Then boxon ' MsgBox ("Bad luck Computer wins."): boxon
NewGame
End Function

Public Function boxoff()
Dim a
a = Mid(Rnd(1), 5, 1)
If Not a > 0 And a <= 9 Then boxoff

Select Case a
Case 1
If t1.BackStyle = 0 Then Call boxoff
t1.BackStyle = 0
m3.BackStyle = 0
b2.BackStyle = 0
Case 2
If t2.BackStyle = 0 Then Call boxoff
t2.BackStyle = 0
t1.BackStyle = 0
b3.BackStyle = 0
Case 3
If t3.BackStyle = 0 Then Call boxoff
t3.BackStyle = 0
t2.BackStyle = 0
m2.BackStyle = 0
Case 4
If m1.BackStyle = 0 Then Call boxoff
m1.BackStyle = 0
b2.BackStyle = 0
t1.BackStyle = 0
Case 5
If m2.BackStyle = 0 Then Call boxoff
m2.BackStyle = 0
b3.BackStyle = 0
t3.BackStyle = 0
Case 6
If m3.BackStyle = 0 Then Call boxoff
m3.BackStyle = 0
b1.BackStyle = 0
m3.BackStyle = 0
Case 7
If b1.BackStyle = 0 Then Call boxoff
b1.BackStyle = 0
m1.BackStyle = 0
b3.BackStyle = 0
Case 8
If b2.BackStyle = 0 Then Call boxoff
b2.BackStyle = 0
m2.BackStyle = 0
m1.BackStyle = 0
Case 9
If b3.BackStyle = 0 Then Call boxoff
t3.BackStyle = 0
b1.BackStyle = 0
m2.BackStyle = 0
End Select
If t1.BackStyle = 0 And t2.BackStyle = 0 And t3.BackStyle = 0 And m1.BackStyle = 0 And m2.BackStyle = 0 And m3.BackStyle = 0 And b1.BackStyle = 0 And b2.BackStyle = 0 And b3.BackStyle = 0 Then Call newPic
Exit Function
End Function

Public Function boxon()
On Error Resume Next
Dim a
a = Mid(Rnd(1), 5, 1)
If Not a > 0 And a <= 9 Then boxon

If t1.BackStyle = 1 And t2.BackStyle = 1 And t3.BackStyle = 1 And m1.BackStyle = 1 And m2.BackStyle = 1 And m3.BackStyle = 1 And b1.BackStyle = 1 And b2.BackStyle = 1 And b3.BackStyle = 1 Then Me.SetFocus: Exit Function

Select Case a
Case 1
If t1.BackStyle = 0 Then t1.BackStyle = 1 Else boxon
Case 2
If t2.BackStyle = 0 Then t2.BackStyle = 1 Else boxon
Case 3
If t3.BackStyle = 0 Then t3.BackStyle = 1 Else boxon
Case 4
If m1.BackStyle = 0 Then m1.BackStyle = 1 Else boxon
Case 5
If m2.BackStyle = 0 Then m2.BackStyle = 1 Else boxon
Case 6
If m3.BackStyle = 0 Then m3.BackStyle = 1 Else boxon
Case 7
If b1.BackStyle = 0 Then b1.BackStyle = 1 Else boxon
Case 8
If b2.BackStyle = 0 Then b2.BackStyle = 1 Else boxon
Case 9
If b3.BackStyle = 0 Then b3.BackStyle = 1 Else boxon
End Select
End Function

Public Function NewGame()
m1.Tag = ""
t1.Caption = ""
t2.Caption = ""
t3.Caption = ""

m1.Caption = ""
m2.Caption = ""
m3.Caption = ""

b1.Caption = ""
b2.Caption = ""
b3.Caption = ""
End Function

Public Function Reset()

t1.BackStyle = 1
t2.BackStyle = 1
t3.BackStyle = 1
m1.BackStyle = 1
m2.BackStyle = 1
m3.BackStyle = 1
b1.BackStyle = 1
b2.BackStyle = 1
b3.BackStyle = 1
End Function
Public Function AIPriority()
Dim a
a = Mid(Rnd(1), 5, 1)
If Not a > 0 And a <= 9 Then AIPriority

If t1.BackStyle = 0 And t2.BackStyle = 0 And t3.BackStyle = 0 And m1.BackStyle = 0 And m2.BackStyle = 0 And m3.BackStyle = 0 And b1.BackStyle = 0 And b2.BackStyle = 0 And b3.BackStyle = 0 Then Exit Function
Select Case a
Case 1
If t1 = "" Then t1 = "O" Else AIPriority
Case 2
If t2 = "" Then t2 = "O" Else AIPriority
Case 3
If t3 = "" Then t3 = "O" Else AIPriority
Case 4
If m1 = "" Then m1 = "O" Else AIPriority
Case 5
If m2 = "" Then m2 = "O" Else AIPriority
Case 6
If m3 = "" Then m3 = "O" Else AIPriority
Case 7
If b1 = "" Then b1 = "O" Else AIPriority
Case 8
If b2 = "" Then b2 = "O" Else AIPriority
Case 9
If b3 = "" Then b3 = "O" Else AIPriority
End Select
XOScan
End Function
  
Public Function OOO()
XOScan

If m1.Tag = "W" Then Exit Function
'Top row scanner
If t1.Caption = "O" And t2.Caption = "O" And t3 = "" Then t3.Caption = "O": XOScan: Exit Function
If t2.Caption = "O" And t3.Caption = "O" And t1 = "" Then t1.Caption = "O": XOScan: Exit Function
If t1.Caption = "O" And t3.Caption = "O" And t2 = "" Then t2.Caption = "O": XOScan: Exit Function
'Middle row scanner
If m1.Caption = "O" And m2.Caption = "O" And m3 = "" Then m3.Caption = "O": XOScan: Exit Function
If m2.Caption = "O" And m3.Caption = "O" And m1 = "" Then m1.Caption = "O": XOScan: Exit Function
If m1.Caption = "O" And m3.Caption = "O" And m2 = "" Then m2.Caption = "O": XOScan: Exit Function
'Bottom row scanner
If b1.Caption = "O" And b2.Caption = "O" And b3 = "" Then b3.Caption = "O": XOScan: Exit Function
If b2.Caption = "O" And b3.Caption = "O" And b1 = "" Then b1.Caption = "O": XOScan: Exit Function
If b1.Caption = "O" And b3.Caption = "O" And b2 = "" Then b2.Caption = "O": XOScan: Exit Function

'Column 1 scanner
If t1.Caption = "O" And m1.Caption = "O" And b1 = "" Then b1.Caption = "O": XOScan: Exit Function
If m1.Caption = "O" And b1.Caption = "O" And t1 = "" Then t1.Caption = "O": XOScan: Exit Function
If t1.Caption = "O" And b1.Caption = "O" And m1 = "" Then m1.Caption = "O": XOScan: Exit Function
'Column 2 scanner
If t2.Caption = "O" And m2.Caption = "O" And b2 = "" Then b2.Caption = "O": XOScan: Exit Function
If m2.Caption = "O" And b2.Caption = "O" And t2 = "" Then t2.Caption = "O": XOScan: Exit Function
If t2.Caption = "O" And b2.Caption = "O" And m2 = "" Then m2.Caption = "O": XOScan: Exit Function
'Column 3 scanner
If t3.Caption = "O" And m3.Caption = "O" And b3 = "" Then b3.Caption = "O": XOScan: Exit Function
If m3.Caption = "O" And b3.Caption = "O" And t3 = "" Then t3.Caption = "O": XOScan: Exit Function
If t3.Caption = "O" And b3.Caption = "O" And m3 = "" Then m3.Caption = "O": XOScan: Exit Function

'Diagonal 1 Scanner
If t1.Caption = "O" And b3.Caption = "O" And m2 = "" Then m2.Caption = "O": XOScan: Exit Function
If t1.Caption = "O" And m2.Caption = "O" And b3 = "" Then b3.Caption = "O": XOScan: Exit Function
If m2.Caption = "O" And b3.Caption = "O" And t1 = "" Then t1.Caption = "O": XOScan: Exit Function
'Diagonal 2 Scanner
If t3.Caption = "O" And m2.Caption = "O" And b1 = "" Then b1.Caption = "O": XOScan: Exit Function
If b1.Caption = "O" And m2.Caption = "O" And t3 = "" Then t3.Caption = "O": XOScan: Exit Function
If t3.Caption = "O" And b1.Caption = "O" And m2 = "" Then m2.Caption = "O": XOScan: Exit Function

AI
End Function

Public Function newPic()
Dim path As String
NewGame

m1.Tag = "W"

If Me.Tag = "p" Then GoTo pi Else If Me.Tag = "L3" Then GoTo pu Else If b3.Tag = "g" Then GoTo u:

Timer.Enabled = True
Exit Function

'Pic2 Code
pi:
seti.Enabled = True
Exit Function
pu:
ata.Enabled = True
Exit Function
u:
last.Enabled = True
Exit Function
End Function

Private Sub Timer_Timer()
NewGame
Reset

t1.BackColor = &HD69B5A
t2.BackColor = &HD69B5A
t3.BackColor = &HD69B5A
m1.BackColor = &HD69B5A
m2.BackColor = &HD69B5A
m3.BackColor = &HD69B5A
b1.BackColor = &HD69B5A
b2.BackColor = &HD69B5A
b3.BackColor = &HD69B5A

Me.Picture = LoadPicture(App.path & "\pic2.bmp")
Me.Tag = "p"
m1.Tag = ""
t1.Tag = ""
Timer.Enabled = False
End Sub

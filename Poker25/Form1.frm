VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H002F9D55&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Block Poker by Vlad Paraschiv"
   ClientHeight    =   4785
   ClientLeft      =   495
   ClientTop       =   510
   ClientWidth     =   7110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Cls"
      Height          =   255
      Left            =   4320
      TabIndex        =   45
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   3000
      TabIndex        =   36
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Find"
      Height          =   255
      Left            =   1800
      TabIndex        =   44
      Top             =   4440
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   255
      Left            =   1680
      TabIndex        =   39
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H00FFFFFF&
      Height          =   3930
      ItemData        =   "Form1.frx":0082
      Left            =   5760
      List            =   "Form1.frx":0084
      TabIndex        =   41
      Top             =   720
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox rich 
      Height          =   255
      Left            =   1780
      TabIndex        =   43
      Top             =   4200
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0086
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3720
      TabIndex        =   40
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   1800
      TabIndex        =   42
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      X1              =   5640
      X2              =   6600
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      X1              =   6600
      X2              =   6600
      Y1              =   600
      Y2              =   720
   End
   Begin VB.Label scor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5950
      TabIndex        =   0
      Top             =   20
      Width           =   1095
   End
   Begin VB.Label Q3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vlad Deck"
      DragIcon        =   "Form1.frx":0108
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   120
      TabIndex        =   47
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label counts 
      Height          =   255
      Left            =   1200
      TabIndex        =   46
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Q1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--------"
      DragIcon        =   "Form1.frx":054A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Q2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "------------"
      DragIcon        =   "Form1.frx":098C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   840
      TabIndex        =   37
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label a1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   11
      Top             =   600
      Width           =   975
   End
   Begin VB.Label b1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   33
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label c1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   32
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label d1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   31
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label e1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   30
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label a2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1800
      TabIndex        =   29
      Top             =   600
      Width           =   975
   End
   Begin VB.Label b2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1800
      TabIndex        =   28
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label c2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1800
      TabIndex        =   27
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label d2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1800
      TabIndex        =   26
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label e2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1800
      TabIndex        =   25
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label a3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2760
      TabIndex        =   24
      Top             =   600
      Width           =   975
   End
   Begin VB.Label b3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2760
      TabIndex        =   23
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label c3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2760
      TabIndex        =   22
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label d3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2760
      TabIndex        =   21
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label e3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2760
      TabIndex        =   20
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label a4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3720
      TabIndex        =   19
      Top             =   600
      Width           =   975
   End
   Begin VB.Label b4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3720
      TabIndex        =   18
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label c4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3720
      TabIndex        =   17
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label d4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3720
      TabIndex        =   16
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label e4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3720
      TabIndex        =   15
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label a5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4680
      TabIndex        =   14
      Top             =   600
      Width           =   975
   End
   Begin VB.Label b5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4680
      TabIndex        =   13
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label c5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4680
      TabIndex        =   12
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Col 2"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Col 3 "
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Col 4"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Col 5"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Col 1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Row B"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Row C"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Row D"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Row E"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Row A"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.Label d5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4680
      TabIndex        =   35
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label e5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4680
      TabIndex        =   34
      Top             =   3000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   375
      Left            =   6120
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Q3.DragMode = 1
Command1.Enabled = False
RandomCard
End Sub

Public Function RandomCard()
Dim T, L
Dim card, suit
Dim cs, XZ
redo:
T = Mid(Rnd(1), 5, 1)
L = Mid(Rnd(1), 5, 1)

If T < "1" Or T > "6" Then GoTo redo
If L < "1" Or L > "4" Then GoTo redo
XZ = L & T

'Converts random number to card
Select Case XZ

Case 11
l1 = "1"
l2 = "A1"
card = "Ace of Hearts"

Case 12
l1 = "2"
l2 = "K1"
card = "King of Hearts"


Case 13
l1 = "3"
l2 = "Q1"
card = "Qeen of Hearts"

Case 14
l1 = "4"
l2 = "J1"
card = "Jack of Hearts"

Case 15
l1 = "5"
l2 = "T1"
card = "10 of Hearts"

Case 16
l1 = "6"
l2 = "N1"
card = "9 of Hearts"

Case 21
l1 = "7"
l2 = "A2"
card = "Ace of Spades"

Case 22
l1 = "8"
l2 = "K2"
card = "King of Spades"

Case 23
l1 = "9"
l2 = "Q2"
card = "Qeen of Spades"

Case 24
l1 = "10"
l2 = "J2"
card = "Jack of Spades"

Case 25
l1 = "11"
l2 = "T2"
card = "10 of Spades"

Case 26
l1 = "12"
l2 = "N2"
card = "9 of Spades"

Case 31
l1 = "13"
l2 = "A3"
card = "Ace of Diamonds"

Case 32
l1 = "14"
l2 = "K3"
card = "King of Diamonds"

Case 33
l1 = "15"
l2 = "Q3"
card = "Qeen of Diamonds"

Case 34
l1 = "16"
l2 = "J3"
card = "Jack of Diamonds"

Case 35
l1 = "17"
l2 = "T3"
card = "10 of Diamonds"

Case 36
l1 = "18"
l2 = "N3"
card = "9 of Diamonds"

Case 41
l1 = "19"
l2 = "A4"
card = "Ace of Clubs"

Case 42
l1 = "20"
l2 = "K4"
card = "King of Clubs"

Case 43
l1 = "21"
l2 = "Q4"
card = "Qeen of Clubs"

Case 44
l1 = "22"
l2 = "J4"
card = "Jack of Clubs"

Case 45
l1 = "23"
l2 = "T4"
card = "10 of Clubs"

Case 46
l1 = "24"
l2 = "N4"
card = "9 of Clubs"
End Select



If Len(rtf.Tag) = 24 Then Finished: Exit Function
If Not rtf.Find(l2, 0, Len(rtf)) = -1 Then GoTo redo
rtf.Text = rtf.Text + l2 & vbCrLf

List1.AddItem card
  
Q3 = card & "                                   " & l2
Q2 = l2
Q1 = l1
End Function
Public Function Finished()
On Error Resume Next
L = Mid(a1, Len(a1) - 3, Len(a1)) + Mid(a2, Len(a2) - 3, Len(a2)) + Mid(a3, Len(a3) - 3, Len(a3)) + Mid(a4, Len(a4) - 3, Len(a4)) + Mid(a5, Len(a5) - 3, Len(a5))
'StrOut "T", a, Mid(a1, Len(a1) - 3, Len(a1)) + Mid(a2, Len(a2) - 3, Len(a2)) + Mid(a3, Len(a3) - 3, Len(a3)) + Mid(a4, Len(a4) - 3, Len(a4)) + Mid(a5, Len(a5) - 3, Len(a5))
'IsPair Mid(a1, Len(a1) - 3, Len(a1)) + Mid(a2, Len(a2) - 3, Len(a2)) + Mid(a3, Len(a3) - 3, Len(a3)) + Mid(a4, Len(a4) - 3, Len(a4)) + Mid(a5, Len(a5) - 3, Len(a5))
'TreeKind Mid(a1, Len(a1) - 3, Len(a1)) + Mid(a2, Len(a2) - 3, Len(a2)) + Mid(a3, Len(a3) - 3, Len(a3)) + Mid(a4, Len(a4) - 3, Len(a4)) + Mid(a5, Len(a5) - 3, Len(a5))
FourKind L
End Function

Public Function IsPair(Source)
Dim duce As Integer
Dim Final As String
Dim a

sta = Source
sta = LTrim(sta)
rich.Text = sta
IsP:
sta = rich.Text
StrOut "A", a, sta
 If a = 2 Then Final = Final & " One Pair of Aces": scor = scor + 15: GoTo IsP
 
StrOut "K", a, sta
 If a = 2 Then Final = Final & " One Pair of Kings": scor = scor + 10: GoTo IsP
 
StrOut "Q", a, sta
 If a = 2 Then Final = Final & " One Pair of Qeens": scor = scor + 9: GoTo IsP
 
StrOut "J", a, sta
 If a = 2 Then Final = Final & " One Pair of Jacks": scor = scor + 8: GoTo IsP
 
StrOut "T", a, sta
 If a = 2 Then Final = Final & " One Pair of Tens": scor = scor + 6: GoTo IsP
 
StrOut "N", a, sta
 If a = 2 Then Final = Final & " One Pair of Nines": scor = scor + 4: GoTo IsP
 
Print Final
End Function
Public Function TreeKind(Source)
Dim duce As Integer
Dim Final As String
Dim a

sta = Source
sta = LTrim(sta)
rich.Text = sta
IsP:
sta = rich.Text
StrOut "A", a, sta
 If a = 3 Then Final = Final & " Three Aces": scor = scor + 30: GoTo IsP
 
StrOut "K", a, sta
 If a = 3 Then Final = Final & " Three Kings": scor = scor + 28: GoTo IsP
 
StrOut "Q", a, sta
 If a = 3 Then Final = Final & " Three Qeens": scor = scor + 26: GoTo IsP
 
StrOut "J", a, sta
 If a = 3 Then Final = Final & " Three Jacks": scor = scor + 24: GoTo IsP
 
StrOut "T", a, sta
 If a = 3 Then Final = Final & " Three Tens": scor = scor + 22: GoTo IsP
 
StrOut "N", a, sta
 If a = 3 Then Final = Final & " Three Nines": scor = scor + 20: GoTo IsP
 
Print Final

If Len(Final) = 0 Then IsPair Mid(a1, Len(a1) - 3, Len(a1)) + Mid(a2, Len(a2) - 3, Len(a2)) + Mid(a3, Len(a3) - 3, Len(a3)) + Mid(a4, Len(a4) - 3, Len(a4)) + Mid(a5, Len(a5) - 3, Len(a5))

End Function


Public Function FourKind(Source)
Dim duce As Integer
Dim Final As String
Dim a

sta = Source
sta = LTrim(sta)
rich.Text = sta
IsP:
sta = rich.Text

StrOut "A", a, sta
 If a = 4 Then Final = Final & " Four Aces of A Kind": scor = scor + 45
 
StrOut "K", a, sta
 If a = 4 Then Final = Final & " Four Kings of A Kind": scor = scor + 42
 
StrOut "Q", a, sta
 If a = 4 Then Final = Final & " Four Qeens of A Kind": scor = scor + 40
 
StrOut "J", a, sta
 If a = 4 Then Final = Final & " Four Jacks of A Kind": scor = scor + 37
 
StrOut "T", a, sta
 If a = 4 Then Final = Final & " Four Tens of A Kind": scor = scor + 35
 
StrOut "N", a, sta
 If a = 4 Then Final = Final & " Four Nines of A Kind": scor = scor + 30
Print Final

If Len(Final) = 0 Then TreeKind Mid(a1, Len(a1) - 3, Len(a1)) + Mid(a2, Len(a2) - 3, Len(a2)) + Mid(a3, Len(a3) - 3, Len(a3)) + Mid(a4, Len(a4) - 3, Len(a4)) + Mid(a5, Len(a5) - 3, Len(a5))
End Function

Public Function StrOut(Letter As String, Instances, Source)

Dim TX As String, inS As Integer, Result As Integer

TX = Source
TX = LTrim(TX)
rich.Text = TX

Do Until InStr(1, rich.Text, Letter, vbTextCompare) = 0
inS = InStr(1, rich.Text, Letter, vbTextCompare)

rich.SelStart = inS - 1
rich.SelLength = 1
rich.SelText = ""

counts = counts + "X"
Loop
Instances = Len(counts)
counts.Caption = ""
End Function

Private Sub a1_DragDrop(Source As Control, X As Single, Y As Single)
If Not a1.Caption = "" Then Exit Sub
a1.Caption = Source
RandomCard
End Sub

Private Sub a2_DragDrop(Source As Control, X As Single, Y As Single)
If Not a2.Caption = "" Then Exit Sub
a2.Caption = Source
RandomCard
End Sub

Private Sub a3_DragDrop(Source As Control, X As Single, Y As Single)
If Not a3.Caption = "" Then Exit Sub
a3.Caption = Source
RandomCard
End Sub

Private Sub a4_DragDrop(Source As Control, X As Single, Y As Single)
If Not a4.Caption = "" Then Exit Sub
a4.Caption = Source
RandomCard
End Sub

Private Sub a5_DragDrop(Source As Control, X As Single, Y As Single)
If Not a5.Caption = "" Then Exit Sub
a5.Caption = Source
RandomCard
End Sub

Private Sub b1_DragDrop(Source As Control, X As Single, Y As Single)
If Not b1.Caption = "" Then Exit Sub
b1.Caption = Source
RandomCard
End Sub

Private Sub b2_DragDrop(Source As Control, X As Single, Y As Single)
If Not b2.Caption = "" Then Exit Sub
b2.Caption = Source
RandomCard
End Sub

Private Sub b3_DragDrop(Source As Control, X As Single, Y As Single)
If Not b3.Caption = "" Then Exit Sub
b3.Caption = Source
RandomCard
End Sub

Private Sub b4_DragDrop(Source As Control, X As Single, Y As Single)
If Not b4.Caption = "" Then Exit Sub
b4.Caption = Source
RandomCard
End Sub

Private Sub b5_DragDrop(Source As Control, X As Single, Y As Single)
If Not b5.Caption = "" Then Exit Sub
b5.Caption = Source
RandomCard
End Sub

Private Sub c1_DragDrop(Source As Control, X As Single, Y As Single)
If Not c1.Caption = "" Then Exit Sub
c1.Caption = Source
RandomCard
End Sub

Private Sub c2_DragDrop(Source As Control, X As Single, Y As Single)
If Not c2.Caption = "" Then Exit Sub
c2.Caption = Source
RandomCard
End Sub

Private Sub c3_DragDrop(Source As Control, X As Single, Y As Single)
If Not c3.Caption = "" Then Exit Sub
c3.Caption = Source
RandomCard
End Sub

Private Sub c4_DragDrop(Source As Control, X As Single, Y As Single)
If Not c4.Caption = "" Then Exit Sub
c4.Caption = Source
RandomCard
End Sub

Private Sub c5_DragDrop(Source As Control, X As Single, Y As Single)
If Not c5.Caption = "" Then Exit Sub
c5.Caption = Source
RandomCard
End Sub


Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Finished
End Sub

Private Sub Command4_Click()
Finished
End Sub

Private Sub Command5_Click()
Cls
a1.Caption = ""
a2.Caption = ""
a3.Caption = ""
a4.Caption = ""
a5.Caption = ""
End Sub

Private Sub d1_DragDrop(Source As Control, X As Single, Y As Single)
If Not d1.Caption = "" Then Exit Sub
d1.Caption = Source
RandomCard
End Sub

Private Sub d2_DragDrop(Source As Control, X As Single, Y As Single)
If Not d2.Caption = "" Then Exit Sub
d2.Caption = Source
RandomCard
End Sub

Private Sub d3_DragDrop(Source As Control, X As Single, Y As Single)
If Not d3.Caption = "" Then Exit Sub
d3.Caption = Source
RandomCard
End Sub

Private Sub d4_DragDrop(Source As Control, X As Single, Y As Single)
If Not d4.Caption = "" Then Exit Sub
d4.Caption = Source
RandomCard
End Sub

Private Sub d5_DragDrop(Source As Control, X As Single, Y As Single)
If Not d5.Caption = "" Then Exit Sub
d5.Caption = Source
RandomCard
End Sub

Private Sub e1_DragDrop(Source As Control, X As Single, Y As Single)
If Not e1.Caption = "" Then Exit Sub
e1.Caption = Source
RandomCard
End Sub

Private Sub e2_DragDrop(Source As Control, X As Single, Y As Single)
If Not e2.Caption = "" Then Exit Sub
e2.Caption = Source
RandomCard
End Sub

Private Sub e3_DragDrop(Source As Control, X As Single, Y As Single)
If Not e3.Caption = "" Then Exit Sub
e3.Caption = Source
RandomCard
End Sub

Private Sub e4_DragDrop(Source As Control, X As Single, Y As Single)
If Not e4.Caption = "" Then Exit Sub
e4.Caption = Source
RandomCard
End Sub

Private Sub e5_DragDrop(Source As Control, X As Single, Y As Single)
If Not e5.Caption = "" Then Exit Sub
e5.Caption = Source
RandomCard
End Sub


Private Sub rtf_Change()
rtf.Tag = rtf.Tag + "X"
End Sub

VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Del 
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin PicClip.PictureClip pi 
      Left            =   3000
      Top             =   2160
      _ExtentX        =   397
      _ExtentY        =   8652
      _Version        =   393216
      Rows            =   25
      Picture         =   "Del.frx":0000
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   480
      Top             =   720
   End
   Begin PicClip.PictureClip picv 
      Left            =   1440
      Top             =   2160
      _ExtentX        =   1720
      _ExtentY        =   11695
      _Version        =   393216
      Rows            =   34
      Picture         =   "Del.frx":3DA2
   End
   Begin VB.Timer Timer3 
      Interval        =   40
      Left            =   960
      Top             =   720
   End
   Begin PicClip.PictureClip bar 
      Left            =   1560
      Top             =   2280
      _ExtentX        =   7699
      _ExtentY        =   13388
      _Version        =   393216
      Rows            =   23
      Picture         =   "Del.frx":1905C
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   2160
      Top             =   720
   End
   Begin PicClip.PictureClip Source 
      Left            =   240
      Top             =   2280
      _ExtentX        =   8546
      _ExtentY        =   14023
      _Version        =   393216
      Rows            =   10
      Picture         =   "Del.frx":85426
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      DisabledPicture =   "Del.frx":1030D0
      Enabled         =   0   'False
      Height          =   330
      Left            =   180
      Picture         =   "Del.frx":103D8A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   180
      Width           =   870
   End
   Begin VB.Image tm 
      Height          =   255
      Left            =   5280
      Top             =   0
      Width           =   495
   End
   Begin VB.Image filen 
      Height          =   255
      Left            =   4500
      Top             =   840
      Width           =   2055
   End
   Begin VB.Image b 
      Height          =   330
      Left            =   1200
      Picture         =   "Del.frx":104A44
      Top             =   240
      Width           =   4380
   End
   Begin VB.Image Out 
      Height          =   690
      Left            =   360
      Top             =   1200
      Width           =   4860
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   3720
      Picture         =   "Del.frx":1095CE
      Top             =   0
      Width           =   1515
   End
   Begin VB.Image Image2 
      Height          =   270
      Left            =   4320
      Picture         =   "Del.frx":10A7E0
      Top             =   600
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   0
      Picture         =   "Del.frx":10B82A
      Top             =   2040
      Width           =   5610
   End
End
Attribute VB_Name = "Del"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Activate()
Dim lpRect As RECT
MouseClip.FillRect Command1, lpRect
lpRect.bottom = lpRect.bottom - 18
lpRect.right = lpRect.right - 10
MouseClip.ClipYes lpRect
End Sub

Private Sub Timer1_Timer()
On Error GoTo Error
b.Tag = b.Tag + "X"
b.Picture = bar.GraphicCell(Len(b.Tag) - 1)
Exit Sub
Error:
MouseClip.ReleaseYes
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Me.Hide
Error.Show
End Sub

Private Sub Timer2_Timer()
On Error GoTo Error
Out.Tag = Out.Tag + "X"
Out.Picture = Source.GraphicCell(Len(Out.Tag) - 1)
Exit Sub
Error:
Out.Tag = ""
End Sub

Private Sub Timer3_Timer()
On Error GoTo Error
filen.Tag = filen.Tag + "X"
filen.Picture = picv.GraphicCell(Len(filen.Tag) - 1)
Exit Sub
Error:
filen.Tag = ""
End Sub

Private Sub Timer4_Timer()
On Error GoTo Error
tm.Tag = tm.Tag + "X"
tm.Picture = pi.GraphicCell(Len(tm.Tag) - 1)
Exit Sub
Error:
tm.Tag = ""
Timer4.Enabled = False
End Sub

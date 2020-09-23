VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   1635
   ClientTop       =   1875
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   3000
   Begin VB.TextBox Text3 
      Height          =   315
      Index           =   2
      Left            =   2400
      TabIndex        =   21
      Text            =   "0"
      Top             =   1380
      Width           =   555
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Index           =   1
      Left            =   1800
      TabIndex        =   20
      Text            =   "0"
      Top             =   1380
      Width           =   555
   End
   Begin VB.TextBox Text9 
      Height          =   315
      Left            =   1200
      TabIndex        =   19
      Text            =   "1"
      Top             =   3540
      Width           =   1755
   End
   Begin VB.TextBox Text8 
      Height          =   315
      Left            =   1200
      TabIndex        =   18
      Text            =   "0"
      Top             =   3180
      Width           =   1755
   End
   Begin VB.TextBox Text7 
      Height          =   315
      Left            =   1200
      TabIndex        =   17
      Text            =   "1"
      Top             =   2820
      Width           =   1755
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   1200
      TabIndex        =   16
      Text            =   "1"
      Top             =   2460
      Width           =   1755
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   1200
      TabIndex        =   15
      Text            =   "100"
      Top             =   2100
      Width           =   1755
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1200
      TabIndex        =   14
      Text            =   "100"
      Top             =   1740
      Width           =   1755
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Clear"
      Height          =   495
      Left            =   1500
      TabIndex        =   13
      Top             =   60
      Width           =   1395
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   6
      Text            =   "0"
      Top             =   1380
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Text            =   "1"
      Top             =   1020
      Width           =   1755
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Text            =   "100"
      Top             =   660
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DrawSpiral"
      Default         =   -1  'True
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1395
   End
   Begin VB.Label Label9 
      Caption         =   "Angle Step"
      Height          =   255
      Left            =   180
      TabIndex        =   12
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "StartAngle"
      Height          =   255
      Left            =   180
      TabIndex        =   11
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "DotLine"
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Direction"
      Height          =   255
      Left            =   180
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Y"
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "X"
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Color"
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Increment"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Radius"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z
Option Compare Binary
Option Base 0

Private Sub Command1_Click()

Dim Success As Integer

' The BIG Function! =Þ
Success = DrawSpiral(Val(Text1.Text), Form2!Picture1, Val(Text2.Text), Val(Text4.Text), Val(Text5.Text), RGB(Val(Text3(0).Text), Val(Text3(1).Text), Val(Text3(2).Text)), Val(Text8.Text), Val(Text6.Text), Val(Text9.Text), Val(Text7.Text), Form1!Command1)

If Success = False Then MsgBox "It didn't work for some reason...", vbCritical, "Error... !"

End Sub

Private Sub Command2_Click()
Form2!Picture1.Cls
End Sub

Private Sub Form_Load()
Form1.Show
Form2.Show
Text4.Text = Trim$(Str$(Form2.ScaleWidth / 2))
Text5.Text = Trim$(Str$(Form2.ScaleHeight / 2))
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End Sub

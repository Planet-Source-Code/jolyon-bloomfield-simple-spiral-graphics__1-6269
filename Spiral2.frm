VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5400
   ClientLeft      =   3015
   ClientTop       =   2445
   ClientWidth     =   6825
   LinkTopic       =   "Form2"
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   455
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   60
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   445
      TabIndex        =   0
      Top             =   60
      Width           =   6735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
Picture1.Left = Form2.ScaleLeft
Picture1.Top = Form2.ScaleTop
Picture1.Width = Form2.ScaleWidth
Picture1.Height = Form2.ScaleHeight
Form1!Text4.Text = Picture1.ScaleWidth / 2
Form1!Text5.Text = Picture1.ScaleHeight / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form1
End Sub

VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8460
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6690
   Height          =   8865
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   6690
   Top             =   1170
   Width           =   6810
   Begin VB.PictureBox Picture1 
      Height          =   5040
      Left            =   330
      ScaleHeight     =   332
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   390
      TabIndex        =   0
      Top             =   495
      Width           =   5910
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Show
    LoadBitmapImage "bmp2.bmp", Picture1
'    LoadBitmapImage "bmp.bmp", Picture1
End Sub



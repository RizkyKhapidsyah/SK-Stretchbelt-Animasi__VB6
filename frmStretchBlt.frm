VERSION 5.00
Begin VB.Form frmStretchBlt 
   AutoRedraw      =   -1  'True
   Caption         =   "StretchBlt"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Timer TimerStretch 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   1440
      Top             =   3000
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   4920
      Picture         =   "frmStretchBlt.frx":0000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   3720
      Picture         =   "frmStretchBlt.frx":3042
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   1020
   End
End
Attribute VB_Name = "frmStretchBlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Chapter 1
'Stretching
'

Option Explicit

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim Shrinking As Boolean
Dim Stretch As Long
Const SpriteWidth As Long = 64
Const SpriteHeight As Long = 64
Const MaxStretch As Long = 96

Private Sub cmdStart_Click()
 
 
 Stretch = 64
 Shrinking = False
 TimerStretch.Enabled = True
  
End Sub

Private Sub TimerStretch_Timer()
Static X As Long, Y As Long

'Clear the form, since we have no background
Me.Cls

If Shrinking Then
    Stretch = Stretch - 2
Else
    Stretch = Stretch + 2
End If

If Stretch < 32 Then Shrinking = False
If Stretch > MaxStretch Then Shrinking = True


'Stretch the sprite into the back buffer
 StretchBlt Me.hdc, X, Y, Stretch, Stretch, picMask.hdc, 0, 0, SpriteWidth, SpriteHeight, vbSrcAnd
 StretchBlt Me.hdc, X, Y, Stretch, Stretch, picSprite.hdc, 0, 0, SpriteWidth, SpriteHeight, vbSrcPaint

X = (X Mod Me.ScaleWidth) + 2
Y = (Y Mod Me.ScaleHeight) + 2

'Force update of the form
Me.Refresh

End Sub

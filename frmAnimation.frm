VERSION 5.00
Begin VB.Form frmAnimation 
   AutoRedraw      =   -1  'True
   Caption         =   "Animation"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   361
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerAnimation 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   1440
      Top             =   4800
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   3240
      Picture         =   "frmAnimation.frx":0000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   9660
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   3240
      Picture         =   "frmAnimation.frx":1E042
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   9660
   End
End
Attribute VB_Name = "frmAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Chapter 1
'Animation Frames
'
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long


Const SpriteWidth As Long = 64
Const SpriteHeight As Long = 64
Const MaxFrames As Long = 10
Const FrameTime As Long = 1000

Dim CurrentTick As Long
Dim LastTick As Long
Dim FrameNumber As Long
Private Sub cmdStart_Click()

TimerAnimation.Enabled = True

LastTick = GetTickCount()
FrameNumber = 1

End Sub



Private Sub TimerAnimation_Timer()
Static X As Long, Y As Long

'Clear the form, since we do not have a background
Me.Cls

'Draw the mask
BitBlt Me.hDC, X, Y, SpriteWidth, SpriteHeight, picMask.hDC, (FrameNumber - 1) * SpriteWidth, 0, vbSrcAnd
'Draw the sprite
BitBlt Me.hDC, X, Y, SpriteWidth, SpriteHeight, picSprite.hDC, (FrameNumber - 1) * SpriteWidth, 0, vbSrcPaint

'Update frame number
'uncomment
'CurrentTick = GetTickCount()


'Check to see if we need to update th frame
'Uncomment
'If CurrentTick - LastTick > FrameTime Then
    
    FrameNumber = (FrameNumber Mod MaxFrames) + 1
    LastTick = GetTickCount()

'Uncomment
'End If


'Update drawing positions
X = (X Mod Me.ScaleWidth) + 1
Y = (Y Mod Me.ScaleHeight) + 1

'Update the form
Me.Refresh

End Sub

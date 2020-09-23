VERSION 5.00
Begin VB.Form frmVoiceClock 
   Caption         =   "Voice Clock"
   ClientHeight    =   450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2415
   Icon            =   "frmVoiceClock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   450
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer tmrVClk 
      Interval        =   100
      Left            =   945
      Top             =   0
   End
End
Attribute VB_Name = "frmVoiceClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bUntold As Boolean

Private Sub Form_Initialize()
   bUntold = True
   tmrVClk.Interval = 100
End Sub

Private Sub tmrVClk_Timer()
   Dim nCurMin As Integer
   
   nCurMin = Minute(Time)
   If nCurMin = 0 Then
      If bUntold Then
         tmrVClk.Interval = 10000
         Call sTellTime
         bUntold = False
      End If
   Else
      bUntold = True
   End If
End Sub


Private Sub sTellTime()
   Dim lHour As Long
   
      lHour = Hour(Format(Time, "Short Time")) ' this will return the hour in military time 12am = 0, 11pm = 23
      lHour = lHour + 100 ' I do this on 2 lines for readability
      
      Call PlaySoundResource(lHour)
End Sub

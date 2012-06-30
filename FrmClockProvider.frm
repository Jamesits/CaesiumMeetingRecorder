VERSION 5.00
Begin VB.Form FrmClockProvider 
   Caption         =   "CMR ClockProvider"
   ClientHeight    =   1050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2265
   LinkTopic       =   "Form1"
   ScaleHeight     =   1050
   ScaleWidth      =   2265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1320
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   360
   End
End
Attribute VB_Name = "FrmClockProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sec, Hms


Private Sub Form_Load()
Me.Hide

End Sub


Public Sub Start()
Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
Sec = Sec + 1
'1s
End Sub

Private Sub Timer2_Timer()
Hms = Hms + 1
'0.1s
End Sub

Public Sub Pause()
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Public Sub Clear()
Sec = 0
Hms = 0
End Sub



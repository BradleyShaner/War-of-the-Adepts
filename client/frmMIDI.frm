VERSION 5.00
Begin VB.Form frmMIDI 
   Caption         =   "MIDI Form"
   ClientHeight    =   585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2460
   LinkTopic       =   "Form1"
   ScaleHeight     =   585
   ScaleWidth      =   2460
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer MidiTimer 
      Interval        =   2000
      Left            =   10000
      Top             =   0
   End
   Begin VB.Label label1 
      Caption         =   "If this window appeared contact the programmer."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmMIDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tacvek Midi Handler Version 1.0
'See Mididoc.txt for instructions on use
'WARNING: DO NOT ALTER ANYTHING BELOW THIS LINE

Option Explicit
Dim curMessage As MSG

Private Sub MidiTimer_Timer()
Dim temp As Integer
Let temp = PeekMessage(curMessage, frmMIDI.hWnd, 0, 0, PM_REMOVE)
If Not (temp = 0 Or temp = -1) Then
    If curMessage.Message = MM_MCINOTIFY And (curMessage.wParam And MCI_NOTIFY_SUCCESSFUL) Then
        RepeatMidi
    End If
End If
End Sub


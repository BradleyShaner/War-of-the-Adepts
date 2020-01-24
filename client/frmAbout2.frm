VERSION 5.00
Begin VB.Form frmAbout2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "War of the Adepts Credits"
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timeCredits 
      Interval        =   35
      Left            =   720
      Top             =   1440
   End
End
Attribute VB_Name = "frmAbout2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intCredits As Long
Private Sub Form_Load()
intCredits = 0
End Sub

Private Sub timeCredits_Timer()
Dim strCredits As String
intCredits = intCredits + 1
Select Case intCredits
Case 1
    strCredits = "Programmed By"
End Select
'Call CreateFont("Arial", 12, True, False, False, &H80FF&)
'Call TextOut(Me.hdc, Me.ScaleWidth / 2 - 50, Me.ScaleHeight - 35, strOBJTitle, Len(strOBJTitle))
End Sub

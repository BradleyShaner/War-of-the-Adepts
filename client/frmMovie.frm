VERSION 5.00
Begin VB.Form frmMovie 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doc Entertainment Presents"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timeHide 
      Interval        =   5500
      Left            =   960
      Top             =   1560
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.doc-ent.com"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Image imgLogo 
      Height          =   1800
      Left            =   2520
      Picture         =   "frmMovie.frx":0000
      Top             =   1320
      Width           =   4125
   End
End
Attribute VB_Name = "frmMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

X% = sndPlaySound(App.Path & "\intro.wav", 1)

End Sub

Private Sub timeHide_Timer()
imgLogo.Visible = False
lblURL.Visible = False
timeHide.Enabled = False
End Sub

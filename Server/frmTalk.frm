VERSION 5.00
Begin VB.Form frmTalk 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dialogue Editor"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4575
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkTalk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dialogue On"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtTalk 
      Height          =   285
      Left            =   120
      MaxLength       =   99
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label lblBossText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Boss Text:"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmTalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

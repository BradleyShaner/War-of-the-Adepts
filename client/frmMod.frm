VERSION 5.00
Begin VB.Form frmMod 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Moderator's Control Panel"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   167
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   279
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCloseGame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close Game"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdPINBan 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PIN Ban"
      Height          =   255
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdGetPin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get PIN"
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdFreeze 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3 Minute Freeze"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdBanIP 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ban IP"
      Height          =   255
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkScrambler 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Scrambler Allowed"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdKill 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1 Day Ban"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdBan 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ban Comp"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reset Rating"
      Height          =   255
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdIP 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get IP"
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdKick 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kick"
      Height          =   255
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdModWarn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Warn"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdPraise 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Praise"
      Height          =   255
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblMotto 
      BackStyle       =   0  'Transparent
      Caption         =   "There are no stupid questions, just stupid people."
      Height          =   615
      Left            =   2280
      TabIndex        =   10
      Top             =   840
      Width           =   1815
   End
   Begin VB.Image imgKenny 
      Height          =   810
      Left            =   2760
      Picture         =   "frmMod.frx":0000
      Top             =   1560
      Width           =   630
   End
End
Attribute VB_Name = "frmMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkScrambler_Click()
On Error Resume Next
If chkScrambler.Value = 1 Then
    frmChat.Chat.SendData "SCRAMON" & vbCrLf
    frmChat.Chat.SendData "ADMINTXT" & "Scrambler has been enabled by the moderator." & vbCrLf
Else
    frmChat.Chat.SendData "SCRAMOFF" & vbCrLf
    frmChat.Chat.SendData "ADMINTXT" & "Scrambler has been disabled by the moderator." & vbCrLf
End If
End Sub

Private Sub cmdBan_Click()
On Error Resume Next
vbinput = MsgBox("Are you sure that you want to ban " & frmChat.lstUsers.Text & "?", vbYesNo, "Ban?")
If vbinput = vbYes Then
    frmChat.Chat.SendData "COMPBAN" & frmChat.lstUsers.Text & vbCrLf
End If

End Sub

Private Sub cmdBanIP_Click()
On Error Resume Next
vbinput = MsgBox("Are you sure that you want to IP ban " & frmChat.lstUsers.Text & "?", vbYesNo, "Ban?")
If vbinput = vbYes Then
    frmChat.Chat.SendData "IPBAN" & frmChat.lstUsers.Text & vbCrLf
End If
End Sub

Private Sub cmdCloseGame_Click()
On Error Resume Next
If frmChat.lstUsers.Text <> "" Then
    frmChat.Chat.SendData "CLOSEGAME" & frmChat.lstUsers.Text & vbCrLf
End If
End Sub

Private Sub cmdFreeze_Click()
On Error Resume Next
frmChat.Chat.SendData "CHATFREEZE" & frmChat.lstUsers.Text & vbCrLf
End Sub

Private Sub cmdGetPin_Click()
On Error Resume Next
frmChat.Chat.SendData "GETPIN" & frmChat.lstUsers.Text & vbCrLf

End Sub

Private Sub cmdIP_Click()
On Error Resume Next
frmChat.Chat.SendData "GETIP" & frmChat.lstUsers.Text & vbCrLf

End Sub

Private Sub cmdKick_Click()
On Error Resume Next
vbinput = MsgBox("Are you sure that you want to kick " & frmChat.lstUsers.Text & "?", vbYesNo, "Ban?")
If vbinput = vbYes Then
    frmChat.Chat.SendData "ADMINKICK" & frmChat.lstUsers.Text & vbCrLf
End If
End Sub

Private Sub cmdKill_Click()
On Error Resume Next
vbinput = MsgBox("Are you sure that you want to 1 Day Ban " & frmChat.lstUsers.Text & "?", vbYesNo, "Ban?")
If vbinput = vbYes Then
    frmChat.Chat.SendData "ADMINKILL" & frmChat.lstUsers.Text & vbCrLf
End If
End Sub

Private Sub cmdModWarn_Click()
On Error Resume Next
If frmChat.lstUsers.Text <> "" Then
    frmChat.Chat.SendData "MODWARN" & frmChat.lstUsers.Text & vbCrLf
End If
End Sub

Private Sub cmdPINBan_Click()
On Error Resume Next
vbinput = MsgBox("Are you sure that you want to 1 Day Ban " & frmChat.lstUsers.Text & "?", vbYesNo, "Ban?")
If vbinput = vbYes Then
    frmChat.Chat.SendData "PINBAN" & frmChat.lstUsers.Text & vbCrLf
End If
End Sub

Private Sub cmdPraise_Click()
On Error Resume Next
If frmChat.lstUsers.Text <> "" Then
    frmChat.Chat.SendData "MODPRAISE" & frmChat.lstUsers.Text & vbCrLf
End If
End Sub

Private Sub cmdReset_Click()
On Error Resume Next
yadda = MsgBox("Are you sure you want to reset this user's ladder ranking?", vbYesNo)
If yadda = vbYes Then
    frmChat.Chat.SendData "USERRESET" & frmChat.lstUsers.Text & vbCrLf
End If

End Sub

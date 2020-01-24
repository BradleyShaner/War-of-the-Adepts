VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmHost 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Golden Sun: The War of the Adepts - Host"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   Icon            =   "frmHost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkGen 
      BackColor       =   &H000000FF&
      Caption         =   "Equalize Weapons"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2760
      TabIndex        =   42
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox chkGen 
      BackColor       =   &H000000FF&
      Caption         =   "Allow Normal Attacks"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   41
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkGen 
      BackColor       =   &H000000FF&
      Caption         =   "Allow Psynergy"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   40
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkGen 
      BackColor       =   &H000000FF&
      Caption         =   "Allow Healing"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   39
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.OptionButton opEqualize 
      BackColor       =   &H000000FF&
      Caption         =   "Client"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   37
      Top             =   2280
      Width           =   735
   End
   Begin VB.OptionButton opEqualize 
      BackColor       =   &H000000FF&
      Caption         =   "Host"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   36
      Top             =   2280
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.FileListBox filMazes 
      Height          =   285
      Left            =   6360
      Pattern         =   "*.omaz"
      TabIndex        =   34
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chkGen 
      BackColor       =   &H000000FF&
      Caption         =   "Equalize Levels"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   32
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   120
      MaxLength       =   20
      TabIndex        =   30
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox chkGen 
      BackColor       =   &H000000FF&
      Caption         =   "Allow Summons"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   29
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.Timer timeWait 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   6360
      Top             =   4800
   End
   Begin VB.HScrollBar hHandicap 
      Height          =   255
      Left            =   5040
      Max             =   5
      Min             =   -5
      TabIndex        =   20
      Top             =   1560
      Value           =   1
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdRandom 
      Caption         =   "Random Name"
      Height          =   195
      Left            =   1680
      TabIndex        =   16
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Game"
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtGameName 
      Height          =   285
      Left            =   120
      MaxLength       =   20
      TabIndex        =   14
      Text            =   "Online Brawl"
      Top             =   480
      Width           =   1575
   End
   Begin VB.ComboBox cmdArena 
      Height          =   315
      ItemData        =   "frmHost.frx":08CA
      Left            =   4680
      List            =   "frmHost.frx":08E3
      TabIndex        =   11
      Text            =   "Collosso"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdBoot 
      Caption         =   "Boot User"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox chkGen 
      BackColor       =   &H000000FF&
      Caption         =   "Double Stats If Win / Lose Stats If Loss"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CheckBox chkGen 
      BackColor       =   &H000000FF&
      Caption         =   "Enable Time Limit"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   6
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock Host 
      Left            =   2040
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9788
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtmsg 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      MaxLength       =   200
      TabIndex        =   2
      Top             =   4080
      Width           =   5655
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Game"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.ComboBox lstLadArena 
      Height          =   315
      ItemData        =   "frmHost.frx":0930
      Left            =   4680
      List            =   "frmHost.frx":0940
      TabIndex        =   33
      Text            =   "Yampi Desert"
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image imgBG 
      Height          =   375
      Left            =   4320
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Equalize Levels At:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   2760
      TabIndex        =   38
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Mazes:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   5040
      TabIndex        =   35
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblgen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   14
      Left            =   120
      TabIndex        =   31
      Top             =   720
      Width           =   1245
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Opponent's Stats:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   28
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lblOpLevel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   27
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblOpRating 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   26
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lblOpName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   25
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Rating:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   23
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Top             =   1440
      Width           =   855
   End
   Begin VB.Image imgHelp 
      Height          =   240
      Left            =   6000
      Picture         =   "frmHost.frx":0973
      Top             =   1850
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblHandicap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6240
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblopHandicap 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6240
      TabIndex        =   19
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Handicap:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   10
      Left            =   4560
      TabIndex        =   18
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Handicap:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   9
      Left            =   5040
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblgen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Game Name:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   1530
   End
   Begin VB.Label lblgen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chat:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   3
      Left            =   0
      TabIndex        =   12
      Top             =   3720
      Width           =   765
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Not Ready"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   4560
      TabIndex        =   9
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Opponent:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   2880
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Game Options:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblmsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting for opponent."
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   6735
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Host A Game:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurUserName As String
Dim strData() As String

Private Sub chkGen_Click(Index As Integer)
On Error Resume Next
timeWait.Enabled = False
DoEvents
timeWait.Enabled = True
If Index = 0 Then
    If chkGen(Index).Value = 1 Then
    Host.SendData "TIMEON" & vbCrLf
    Else
    Host.SendData "TIMEOFF" & vbCrLf
    End If
End If
If Index = 1 Then
    If chkGen(Index).Value = 1 Then
    Host.SendData "RATEON" & vbCrLf
    DoubleStats = True
    Else
    Host.SendData "RATEOFF" & vbCrLf
    DoubleStats = False
    End If
End If
If Index = 2 Then
    If chkGen(Index).Value = 1 Then
    Host.SendData "ALLOWSUMON" & vbCrLf
    bAllowSummon = True
    Else
    Host.SendData "ALLOWSUMOFF" & vbCrLf
    bAllowSummon = False
    End If
End If
If Index = 3 Then
    If chkGen(Index).Value = 1 Then
        If opEqualize(0).Value = True Then
            Handicap(1) = 0
            Handicap(2) = CStr(CInt(strLvl) - CInt(stroLvl))
            If Me.chkGen(3).Value = 1 Then
                Host.SendData "HCAP" & "0" & vbCrLf
                Host.SendData "YOURHCAP" & CStr(CInt(strLvl) - CInt(stroLvl)) & vbCrLf
            End If
        Else
            Handicap(1) = CInt(stroLvl) - CInt(strLvl)
            Handicap(2) = 0
            If Me.chkGen(3).Value = 1 Then
                Host.SendData "HCAP" & CStr(Handicap(1)) & vbCrLf
                Host.SendData "YOURHCAP" & "0" & vbCrLf
            End If
        End If
    Else
        Handicap(1) = 0
        Host.SendData "HCAP0" & vbCrLf
        Host.SendData "YOURHCAP0" & vbCrLf
    End If
End If
If Index = 4 Then
    If chkGen(4).Value = 1 Then
        bAllowHeal = True
        frmHost.Host.SendData "HEALON" & vbCrLf
    Else
        bAllowHeal = False
        frmHost.Host.SendData "HEALOFF" & vbCrLf
    End If
End If
If Index = 5 Then
    If chkGen(5).Value = 1 Then
        bAllowPsynergy = True
        frmHost.Host.SendData "PSYON" & vbCrLf
    Else
        bAllowPsynergy = False
        frmHost.Host.SendData "PSYOFF" & vbCrLf
    End If
End If
If Index = 6 Then
    If chkGen(6).Value = 1 Then
        bAllowAttack = True
        frmHost.Host.SendData "ATTACKON" & vbCrLf
    Else
        bAllowAttack = False
        frmHost.Host.SendData "ATTACKOFF" & vbCrLf
    End If
End If
If Index = 7 Then
    If chkGen(7).Value = 1 Then
        frmHost.Host.SendData "EWEAPONSON" & vbCrLf
        bEqualizeWeapons = True
        'If opEqualize(0).Value = True Then
        '    frmHost.Host.SendData "ETHISWEAPONS" & strmyweapon & vbCrLf
        'End If
    Else
        frmHost.Host.SendData "EWEAPONSOFF" & vbCrLf
        bEqualizeWeapons = False
    End If
End If

End Sub

Private Sub cmdArena_Change()
If cmdArena.Text = "thereisnocowlevel" Then
    MsgBox "Secret level activated! (Easter Egg #4)"
    Call Encode("4", "EGG4", "EGGL4", App.Path & "\settings.ini")
    
End If
End Sub

Private Sub cmdBoot_Click()
On Error Resume Next
yadda = MsgBox("Are you sure?", vbYesNo)
If yadda = vbYes Then
Host.SendData "DISC" & vbCrLf
DoEvents
Host.Close
DoEvents
Host.Listen
DoEvents
cmdBoot.Enabled = False
cmdStart.Enabled = False
cmdCreate.Enabled = True
txtMsg.Enabled = False
cmdSend.Enabled = False
lstuser.Clear
lblmsg.Caption = "SERVERMSG: Booted user"
chkGen(0).Enabled = False
chkGen(1).Enabled = False
cmdArena.Enabled = False
End If
End Sub

Private Sub cmdCreate_Click()
On Error Resume Next

If txtGameName.Text <> "" Then

    If Host.State <> sckClosed Then
        Host.Close
        DoEvents
    End If
    
    Host.Listen
    cmdCreate.Enabled = False
    txtGameName.Enabled = False
    
    frmChat.Chat.SendData "CREATEGAME" & txtGameName.Text & vbCrLf
    frmChat.Chat.SendData "HOSTNAME" & strMyUserName & vbCrLf
    
    If strMyUserName = "dragoon" Then
        strRealIP = IKILLKENNYIP
    End If
    
    frmChat.Chat.SendData "METXT" & strMyUserName & " has hosted a game (" & strRealIP & ") - " & txtGameName.Text & vbCrLf
    
    lstuser.AddItem strUser
    
    Else
    
    MsgBox "You must have a valid game name!"

End If

End Sub

Private Sub cmdListen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
If Host.State <> sckClosed Then
Host.Close
'Timeout (1)
End If
Host.Listen
cmdListen.Enabled = False
End If
If Button = 2 Then
Host.Close
End If
End Sub

Private Sub cmdNewUser_Click()

End Sub

Private Sub cmdRandom_Click()
Dim RandInt As Integer
Randomize
RandInt = Int(Rnd * 11) + 1
If RandInt = 1 Then
    txtGameName.Text = "Online Brawl"
ElseIf RandInt = 2 Then
    txtGameName.Text = "Battle Royale"
ElseIf RandInt = 3 Then
    txtGameName.Text = "Ready To Rumble"
ElseIf RandInt = 4 Then
    txtGameName.Text = "Melee Mayham"
ElseIf RandInt = 5 Then
    txtGameName.Text = "Face to Face Combat"
ElseIf RandInt = 6 Then
    txtGameName.Text = "One on One Battle"
ElseIf RandInt = 7 Then
    txtGameName.Text = "Challengers Welcome"
ElseIf RandInt = 8 Then
    txtGameName.Text = "Adept War"
ElseIf RandInt = 9 Then
    txtGameName.Text = "Turn Based Trouble"
ElseIf RandInt = 10 Then
    txtGameName.Text = "Close Combat"
Else
    txtGameName.Text = "Deadly Duel"
End If
End Sub

Private Sub cmdSend_Click()
On Error Resume Next
Host.SendData "MSG" & txtMsg.Text & vbCrLf
txtMsg.Text = ""
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdStart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If timeWait.Enabled = False Then
    'Unload frmBattle
    'Unload frmArena
'Change for ladder tournament:
    If lblGen(7).Caption <> "Not Ready" Then
       If cmdArena.Text = "Collosso" Then
            Host.SendData "ARENA1" & vbCrLf
            imgBG.Picture = LoadPicture(App.Path & "\arena1.gif")
        ElseIf cmdArena.Text = "Vale" Then
            Host.SendData "ARENA2" & vbCrLf
            imgBG.Picture = LoadPicture(App.Path & "\arena2.gif")
        ElseIf cmdArena.Text = "Temple" Then
            Host.SendData "ARENA3" & vbCrLf
            imgBG.Picture = LoadPicture(App.Path & "\arena3.gif")
        ElseIf cmdArena.Text = "Field" Then
            Host.SendData "ARENA4" & vbCrLf
            imgBG.Picture = LoadPicture(App.Path & "\arena4.gif")
        ElseIf cmdArena.Text = "Sol Sanctum" Then
            Host.SendData "ARENA5" & vbCrLf
            imgBG.Picture = LoadPicture(App.Path & "\arena5.gif")
        ElseIf cmdArena.Text = "thereisnocowlevel" Then
            Host.SendData "ARENA6" & vbCrLf
            imgBG.Picture = LoadPicture(App.Path & "\arena6.gif")
        ElseIf cmdArena.Text = "Venus Lighthouse" Then
            Host.SendData "ARENA7" & vbCrLf
            imgBG.Picture = LoadPicture(App.Path & "\arena7.gif")
        ElseIf cmdArena.Text = "Tret Tree" Then
            Host.SendData "ARENA8" & vbCrLf
            imgBG.Picture = LoadPicture(App.Path & "\arena8.gif")
        Else
            MsgBox "Invalid arena selection!"
            Exit Sub
        End If
'        If lstArena.Text = "Yampi Desert" Then
'            Host.SendData "ARENA1" & vbCrLf
'            frmBattle.imgArena.Picture = LoadPicture(App.Path & "\Ladarena1.gif")
'        ElseIf lstLadArena.Text = "Pirate Ship" Then
'            Host.SendData "ARENA2" & vbCrLf
'            frmBattle.imgArena.Picture = LoadPicture(App.Path & "\Ladarena2.gif")
'        ElseIf lstLadArena.Text = "Air's Rock" Then
'            Host.SendData "ARENA3" & vbCrLf
'            frmBattle.imgArena.Picture = LoadPicture(App.Path & "\Ladarena3.gif")
'        ElseIf lstLadArena.Text = "Forest" Then
'            Host.SendData "ARENA4" & vbCrLf
'            frmBattle.imgArena.Picture = LoadPicture(App.Path & "\Ladarena4.gif")
'        Else
'            MsgBox "Invalid arena selection!"
'            Exit Sub
'        End If
        
If chkGen(0).Value = 1 Then
Host.SendData "TIMEON" & vbCrLf
Else
Host.SendData "TIMEOFF" & vbCrLf
End If
If chkGen(1).Value = 1 Then
Host.SendData "RATEON" & vbCrLf
DoubleStats = True
Else
Host.SendData "RATEOFF" & vbCrLf
DoubleStats = False
End If
If chkGen(2).Value = 1 Then
Host.SendData "ALLOWSUMON" & vbCrLf
bAllowSummon = True
Else
Host.SendData "ALLOWSUMOFF" & vbCrLf
bAllowSummon = False
End If
If chkGen(3).Value = 1 Then
    If opEqualize(0).Value = True Then
        Handicap(1) = 0
        Handicap(2) = CStr(CInt(strLvl) - CInt(stroLvl))
        If Me.chkGen(3).Value = 1 Then
            Host.SendData "HCAP" & "0" & vbCrLf
            Host.SendData "YOURHCAP" & CStr(CInt(strLvl) - CInt(stroLvl)) & vbCrLf
        End If
    Else
        Handicap(1) = CInt(stroLvl) - CInt(strLvl)
        Handicap(2) = 0
        If Me.chkGen(3).Value = 1 Then
            Host.SendData "HCAP" & CStr(Handicap(1)) & vbCrLf
            Host.SendData "YOURHCAP" & "0" & vbCrLf
        End If
    End If
Else
    Handicap(1) = 0
    Host.SendData "HCAP0" & vbCrLf
    Host.SendData "YOURHCAP0" & vbCrLf
End If
If chkGen(4).Value = 1 Then
    bAllowHeal = True
    frmHost.Host.SendData "HEALON" & vbCrLf
Else
    bAllowHeal = False
    frmHost.Host.SendData "HEALOFF" & vbCrLf
End If
If chkGen(5).Value = 1 Then
    bAllowPsynergy = True
    frmHost.Host.SendData "PSYON" & vbCrLf
Else
    bAllowPsynergy = False
    frmHost.Host.SendData "PSYOFF" & vbCrLf
End If
If chkGen(6).Value = 1 Then
    bAllowAttack = True
    frmHost.Host.SendData "ATTACKON" & vbCrLf
Else
    bAllowAttack = False
    frmHost.Host.SendData "ATTACKOFF" & vbCrLf
End If
If chkGen(7).Value = 1 Then
    frmHost.Host.SendData "EWEAPONSON" & vbCrLf
Else
    frmHost.Host.SendData "EWEAPONSOFF" & vbCrLf
End If
    

        
                
    MazeWait = True
    Host.SendData "LOADARENA" & vbCrLf 'Start the prebattle
    DoEvents
    
    hoston = True
    
    Unload frmMultiplayer
    StopMidi
    
    bMazeFirstLoad = True
    
    BattleLoaded(1) = False
    BattleLoaded(2) = False
    
    If chkGen(0).Value = 1 Then
        TimedMatch = True
    Else
        TimedMatch = False
    End If
    
    Me.Hide
    
    Unload frmBattle
    
    Call LoadBattle
    frmBattle.Show
    
    'frmArena.Show
    
    
    'frmArena.timecount.Enabled = True
    
    Else
    lblmsg.Caption = "SERVERMSG: Opponent not ready!"
    End If

Else
    MsgBox "You cannot start a game for 7 seconds after making a change.  Please try again in a few moments."
End If

End Sub

Private Sub filMazes_Click()
filMazes.Refresh
End Sub

Private Sub Form_Activate()
On Error Resume Next
If Host.State = sckClosed Then 'If you're not already connected
    Call Form_Load
ElseIf Host.State <> 2 Then
    'Comment out for ladder tournament:
    For i = 0 To chkGen.UBound
        chkGen(i).Enabled = True
    Next 'i
    hHandicap.Enabled = False
    cmdArena.Enabled = True
Else
    hHandicap.Value = 0
    hHandicap.Enabled = False
End If
filMazes.Path = App.Path & "\files"
End Sub

Private Sub Form_Load()
On Error Resume Next
bMazeFirstLoad = False
hHandicap.Enabled = True
'Dim RandInt As Integer
'RandInt = Int(Rnd * 5) + 1
'If RandInt = 1 Then
'    txtGameName.Text = "Online Brawl"
'ElseIf RandInt = 2 Then
'    txtGameName.Text = "Battle Royale"
'ElseIf RandInt = 3 Then
'    txtGameName.Text = "Ready To Rumble"
'ElseIf RandInt = 4 Then
'    txtGameName.Text = "Melee Mayham"
'Else
'    txtGameName.Text = "Face to Face Combat"
'End If
Call cmdRandom_Click

DoubleStats = False

txtGameName.Enabled = True
lblGen(2).Caption = Host.LocalIP
Host.Close
DataSent = False
txtMsg.Enabled = False
cmdSend.Enabled = False
lblmsg.Caption = "Not Connected!"
cmdStart.Enabled = False
cmdBoot.Enabled = False
cmdArena.Enabled = False
chkGen(1).Enabled = False
chkGen(0).Enabled = False
chkGen(2).Enabled = False
chkGen(3).Enabled = False
chkGen(2).Value = 1
chkGen(0).Value = 1
bAllowHeal = True
bAllowSummon = True
bAllowAttack = True
bAllowPsynergy = True

opEqualize(0).Value = True
opEqualize(1).Value = False
bAllowSummon = True
strImage = GetFromIni("GEN", "IMAGES", App.Path & "\settings.ini")
If strImage = "ON" Then
    Me.Picture = frmIntro.Picture
End If
hoston = True
hHandicap.Value = 0

lblOpName.Caption = ""
lblOpRating.Caption = ""
lblOpLevel.Caption = ""
lblGen(7).Caption = "Not Ready"


End Sub

Private Sub Form_Unload(Cancel As Integer)
yadda = MsgBox("Warning! Quiting will cancel your game.  Are you sure that you want to quit?", vbYesNo, "Quit?")

If yadda = vbYes Then
frmChat.Chat.SendData "CLOSEGAME" & vbCrLf
frmIntro.Show
End If
If yadda = vbNo Then
Cancel = 1
End If

End Sub

Private Sub hHandicap_Change()
On Error Resume Next
timeWait.Enabled = False
DoEvents
timeWait.Enabled = True
lblHandicap.Caption = hHandicap.Value
Handicap(1) = hHandicap.Value
Host.SendData "HCAP" & hHandicap.Value & vbCrLf
End Sub

Private Sub Host_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
If Host.State <> sckOpen And Host.State <> sckClosed Then
    Beep
    Host.Close
    Host.Accept requestID

If Host.RemoteHostIP <> Host.LocalIP Then
    'Commented out for LADDER Tournament:
    lblmsg.Caption = "SERVER MESSAGE: Client connected!"
    txtMsg.Enabled = True
    cmdSend.Enabled = True
    cmdStart.Enabled = True
    winRace = 0

    bMazeFirstLoad = True
    
    hHandicap.Enabled = False
    
    'Comment out for Ladder Tournament
    For i = 0 To chkGen.UBound
        chkGen(i).Enabled = True
    Next 'i
    
    opEqualize(0).Enabled = True
    opEqualize(1).Enabled = True
    opEqualize(0).Value = True
    
    chkGen(0).Value = 0
    chkGen(1).Value = 0
    'chkGen(0).Enabled = True
    'chkGen(1).Enabled = True
    'chkGen(2).Enabled = True
    'Commented out for LADDER TOURNAMENT
    'chkGen(3).Enabled = True
    
    cmdArena.Enabled = True
    cmdBoot.Enabled = True
    Host.SendData "USER" & strMyUserName & vbCrLf
    hoston = True
    Host.SendData "MYRATING" & strRating & vbCrLf
    Host.SendData "MYLEVEL" & strLvl & vbCrLf
    DoEvents
    
    MsgBox "Client connected!"
    

    frmChat.Chat.SendData "CLOSEGAME" & vbCrLf
Else
    If Host.RemoteHostIP = IKILLKENNYIP Then
        'Host.Accept requestID
        lblmsg.Caption = "SERVER MESSAGE: Client connected!"
        txtMsg.Enabled = True
        cmdSend.Enabled = True
        cmdStart.Enabled = True
        chkGen(0).Enabled = True
        chkGen(1).Enabled = True
        cmdBoot.Enabled = True
        cmdArena.Enabled = True
        Host.SendData "USER" & strMyUserName & vbCrLf
        hoston = True
        MsgBox "Client connected!"
    Else
        lblmsg.Caption = "SERVER MESSAGE: You are not allowed to play yourself."
        Host.Close
    End If
End If
End If


End Sub

Private Sub Host_DataArrival(ByVal bytesTotal As Long)
On Error GoTo err
Dim strdatao As String

Dim strTime As String
strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")

Dim bWrite As Boolean
bWrite = True 'Write to ini or not

Host.GetData strdatao
strData = Split(strdatao, vbCrLf, -1, vbTextCompare) 'Split packets between @'s


Dim strTempPlayer As String
Dim intTempPlayer As Integer

For i = 0 To UBound(strData)


    If Left$(strData(i), 8) = "GETSTATS" Then
        frmHost.Host.SendData "HP" & (i + 2) & HP(i) & vbCrLf & "SPEED" & (i + 2) & Speed(i) & vbCrLf & "AP" & (i + 2) & AP(i) & vbCrLf & "DEFENSE" & (i + 2) & Defense(i) & vbCrLf & "PIC" & (i + 2) & Char(i) & vbCrLf & "CHARTYPE" & (i + 2) & CharType(i) & vbCrLf & "RESEARTH" & (i + 2) & intEarthResist(i) & vbCrLf & "RESFIRE" & (i + 2) & intFireResist(i) & vbCrLf & "RESWIND" & (i + 2) & intWindResist(i) & vbCrLf & "RESWATER" & (i + 2) & intWaterResist(i) & vbCrLf & "RESHEART" & (i + 2) & intHeartResist(i) & vbCrLf & "RESDARK" & (i + 2) & intDarkResist(i) & vbCrLf
        DoEvents
    End If

    
    If Left$(strData(i), 6) = "NEEDPW" Then 'Message from the host
        Dim strGamePW As String
        strGamePW = Mid$(strData(i), 7, Len(strData(i)))
        If strGamePW = txtPassword.Text Then
        '    Call AcceptClient
        Else
            Host.SendData "BADPW" & vbCrLf
            DoEvents
            Host.Close
            Call Form_Load
            Call cmdCreate_Click
        End If
    End If
    If Left$(strData(i), 3) = "MSG" Then 'Message from the host
        lblmsg.Caption = strOpponent & ": " & Mid(strData(i), 4, Len(strData(i)))
    End If
    If Left$(strData(i), 4) = "MYHP" Then 'Foe's HP
        strTempPlayer = Mid$(strData(i), 5, 1)
        intTempPlayer = CInt(strTempPlayer)
        HP(intTempPlayer) = CInt(Mid$(strData(i), 6, Len(strData(i))))
    End If
    
    If Left$(strData(i), 4) = "GMSG" Then 'In game chat msg
        frmBattle.lblmsg.Caption = strOpponent & ": " & Mid(strData(i), 5, Len(strData(i)))
    End If
    

    If Left$(strData(i), 4) = "USER" Then 'Opponent's name
        strOpponent = Mid(strData(i), 5, Len(strData(i)))
        If strOpponent = strMyUserName Then
            MsgBox "You are not allowed to play yourself!"
            Host.Close
        End If
        Dim bCheckOp As Boolean
        bCheckOp = CheckOpponent(strOpponent, Host.RemoteHostIP)
        If bCheckOp = True Then
            MsgBox "You are not allowed to play someone more than three times in a single day.  Please wait until tommorow (12 AM EST) to play this person again."
            Host.Close
        End If
        
        If txtPassword.Text <> "" Then
            Host.SendData "NEEDPW" & vbCrLf
        End If
        
        lblOpName.Caption = strOpponent
    End If
    
    If Left$(strData(i), 8) = "MYRATING" Then
        opRating = Mid(strData(i), 9, Len(strData(i)))
        lblOpRating.Caption = opRating
    End If
    
    If Left$(strData(i), 7) = "MYLEVEL" Then
        stroLvl = Mid(strData(i), 8, Len(strData(i)))
        lblOpLevel.Caption = stroLvl
    End If
    
    
    If Left$(strData(i), 5) = "READY" Then 'Opponent is ready
        lblGen(7).Caption = "Ready"
    End If
    
    If Left$(strData(i), 8) = "NOTREADY" Then 'Opponent is not ready
        lblGen(7).Caption = "Not Ready"
    End If
    
    If Left$(strData(i), 4) = "DISC" Then 'The opponent disconnected (not currently implemented)
        Unload frmBattle
        BattleLoaded(1) = False
        BattleLoaded(2) = False
        Host.Close
        lblmsg.Caption = "SERVERMSG: The other party disconnected!"
    End If
    
    
    If Left$(strData(i), 7) = "OPREADY" Then
        strTempPlayer = Mid$(strData(i), 8, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True 'Opponent is ready
        If bOReady(2) = True And bOReady(4) = True Then
            Call DoAttacks
        End If
    End If
    
    If Left$(strData(i), 4) = "TIME" Then 'Current time remaining on clock
        strTime = Mid(strData(i), 5, Len(strData(i)))
        curCount = CInt(strTime)
        frmBattle.lblGen(18).Caption = curCount
    End If
    
    If Left$(strData(i), 3) = "LVL" Then 'Opponent's Level
        stroLvl = Mid(strData(i), 4, Len(strData(i)))
        intoLvl = CInt(stroLvl)
    End If
    
    If Left$(strData(i), 9) = "LOADARENA" Then
        If BattleLoaded(1) = False And BattleLoaded(2) = False Then
            MazeWait = False
            'frmArena.Show
        End If
    End If
    

    
' This is no longer used (below)
'    If Left$(strdata(i), 4) = "TYPE" Then 'Opponent's elemental type
'        stroType = Mid(strdata(i), 5, Len(strdata(i)))
'    '    intoType = CInt(stroType)
'    End If
    
    If Left$(strData(i), 6) = "RATING" Then 'Opponent's rating
        opRating = CInt(Mid(strData(i), 7, Len(strData(i))))
    '    intoType = CInt(stroType)
    End If
    
    If Left$(strData(i), 4) = "CHAR" Then
        strTempPlayer = Mid$(strData(i), 5, 1)
        intTempPlayer = CInt(strTempPlayer)
        
        CharName(intTempPlayer) = Mid(strData(i), 6, Len(strData(i)))
        bCustomChar(intTempPlayer) = FindWhichCharacter(CharName(intTempPlayer))
    End If
    
    If Left$(strData(i), 6) = "TARGET" Then
        strTempPlayer = Mid$(strData(i), 7, 1)
        intTempPlayer = CInt(strTempPlayer)
        Target(intTempPlayer) = CLng(Mid$(strData(i), 8, Len(strData(i))))
    End If
    
    If Left$(strData(i), 3) = "PIC" Then 'Opponent's picture
        strTempPlayer = Mid$(strData(i), 4, 1)
        intTempPlayer = CInt(strTempPlayer)
        If intTempPlayer = 3 Then
            If bCustomChar(3) = 999 Or bCustomChar(3) = 0 Then
                stroChar = Mid(strData(i), 5, Len(strData(i)))
                intoChar = CInt(stroChar)
                frmBattle.imgYou(1).Picture = frmBattle.imgUser(intoChar).Picture
                frmBattle.imgYou(3).Picture = frmBattle.imgUser(intoChar).Picture
            Else
                stroChar = Mid(strData(i), 5, Len(strData(i)))
                frmBattle.imgYou(1).Picture = LoadPicture(App.Path & "\files\" & CustomChar(bCustomChar(3)).Picture & ".gif")
                frmBattle.imgYou(3).Picture = LoadPicture(App.Path & "\files\" & CustomChar(bCustomChar(3)).Picture & ".gif")
            End If
        ElseIf intTempPlayer = 4 Then
            If bCustomChar(4) = 999 Or bCustomChar(4) = 0 Then
                stroChar = Mid(strData(i), 5, Len(strData(i)))
                intoChar = CInt(stroChar)
                Char(intTempPlayer) = CInt(stroChar)
                frmBattle.imgYou(5).Picture = frmBattle.imgUser(intoChar).Picture
                frmBattle.imgYou(7).Picture = frmBattle.imgUser(intoChar).Picture
            Else
                stroChar = Mid(strData(i), 5, Len(strData(i)))
                frmBattle.imgYou(5).Picture = LoadPicture(App.Path & "\files\" & CustomChar(bCustomChar(4)).Picture & ".gif")
                frmBattle.imgYou(7).Picture = LoadPicture(App.Path & "\files\" & CustomChar(bCustomChar(4)).Picture & ".gif")
            End If
        End If
    End If
    
    If Left$(strData(i), 8) = "CHARTYPE" Then 'Opponent's character
        strTempPlayer = Mid$(strData(i), 9, 1)
        intTempPlayer = CInt(strTempPlayer)
        CharType(intTempPlayer) = Mid(strData(i), 10, Len(strData(i)))
    End If
    
    If Left$(strData(i), 2) = "AP" Then 'Opponent's AP
        strTempPlayer = Mid$(strData(i), 3, 1)
        intTempPlayer = CInt(strTempPlayer)
        stroAP = Mid(strData(i), 4, Len(strData(i)))
        AP(intTempPlayer) = CInt(stroAP)
    End If
    
    If Left$(strData(i), 5) = "SPEED" Then 'Opponent's AP
        strTempPlayer = Mid$(strData(i), 6, 1)
        intTempPlayer = CInt(strTempPlayer)
        strspeed = Mid(strData(i), 7, Len(strData(i)))
        Speed(intTempPlayer) = CInt(strspeed)
    End If
    
    If Left$(strData(i), 2) = "HP" Then 'Opponent's HP
        If HP(1) <= 0 And HP(2) <= 0 And BattleLoaded(1) = False Then
            Call LoadBattle
        End If
        
        strTempPlayer = Mid$(strData(i), 3, 1)
        intTempPlayer = CInt(strTempPlayer)
        
        
        stroHP = Mid(strData(i), 4, Len(strData(i)))
        HP(intTempPlayer) = CInt(stroHP)
        MaxHP(intTempPlayer) = HP(intTempPlayer)
        frmBattle.lblHP(intTempPlayer - 1).Caption = HP(intTempPlayer) 'Make HP label equal to enemy HP
        frmBattle.shpHP(intTempPlayer - 1).Width = HP(intTempPlayer) / 5 'Set HP bar width
        frmBattle.timeBoready.Enabled = True 'Now that the opponent has HP, turn on the check to see if HP is less than 0
    End If
    
    If Left$(strData(i), 7) = "DEFENSE" Then 'Opponent's defense
        strTempPlayer = Mid$(strData(i), 8, 1)
        intTempPlayer = CInt(strTempPlayer)
        stroDefense = Mid(strData(i), 9, Len(strData(i)))
        Defense(intTempPlayer) = CInt(stroDefense)
    End If
    
    
    
    'Attack info
    If Left$(strData(i), 2) = "PP" Then
        strTempPlayer = Mid$(strData(i), 3, 1)
        intTempPlayer = CInt(strTempPlayer)
        AttackType(intTempPlayer) = "PP"
        
        bOReady(intTempPlayer) = True
            
    End If
    
    If Left$(strData(i), 3) = "DMG" Then
        strTempPlayer = Mid$(strData(i), 4, 1)
        intTempPlayer = CInt(strTempPlayer)
        
        AttackDamage(intTempPlayer) = Mid(strData(i), 5, Len(strData(i)))
        bOReady(intTempPlayer) = True
    End If
    
    If Left$(strData(i), 10) = "DOCRITICAL" Then
        strTempPlayer = Mid$(strData(i), 11, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "CRITICAL"
    End If
    If Left$(strData(i), 9) = "DOSPECIAL" Then
        strTempPlayer = Mid$(strData(i), 10, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "SPECIAL"
    End If
    
    If Left$(strData(i), 5) = "DOPSY" Then
        strTempPlayer = Mid$(strData(i), 6, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "PSY"
        
    End If
    
    If Left$(strData(i), 10) = "DROPATTACK" Then
        strTempPlayer = Mid$(strData(i), 11, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "REDUCEAP"
    
    End If
    
    If Left$(strData(i), 11) = "DROPDEFENSE" Then
        strTempPlayer = Mid$(strData(i), 12, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "REDUCEDEFENSE"
    End If
    
    If Left$(strData(i), 7) = "BOOSTAP" Then
        strTempPlayer = Mid$(strData(i), 8, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "BOOSTAP"
    
    End If
    
    If Left$(strData(i), 4) = "HEAL" Then
        strTempPlayer = Mid$(strData(i), 5, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "HEAL"
    
    End If
    
    If Left$(strData(i), 12) = "BOOSTDEFENSE" Then
        strTempPlayer = Mid$(strData(i), 13, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "BOOSTDEFENSE"
    
    End If
    
    If Left$(strData(i), 8) = "DOATTACK" Then
        strTempPlayer = Mid$(strData(i), 9, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "DAMAGE"
    
    End If
    
    If Left$(strData(i), 3) = "LVL" Then
        strTempPlayer = Mid$(strData(i), 4, 1)
        intTempPlayer = CInt(strTempPlayer)
        Level(intTempPlayer) = CLng(Mid(strData(i), 5, Len(strData(i))))
    End If
    
    If Left$(strData(i), 4) = "CHAR" Then 'What character is the opponent?
        strTempPlayer = Mid$(strData(i), 5, 1)
        intTempPlayer = CInt(strTempPlayer)
        CharName(intTempPlayer) = Mid$(strData(i), 6, Len(strData(i)))
    End If
    
    If Left$(strData(i), 4) = "TYPE" Then 'What elemental type is the opponent?
        strTempPlayer = Mid$(strData(i), 5, 1)
        intTempPlayer = CInt(strTempPlayer)
        CharType(intTempPlayer) = Mid(strData(i), 6, Len(strData(i)))
    
    End If
    
    
    If Left$(strData(i), 11) = "DJINNDAMAGE" Then
        strTempPlayer = Mid$(strData(i), 12, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "DJINNDAMAGE"
    End If
    
    If Left$(strData(i), 9) = "DJINNHEAL" Then
        strTempPlayer = Mid$(strData(i), 10, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "DJINNHEAL"
    End If
    
    If Left$(strData(i), 16) = "DJINNDROPDEFENSE" Then
        strTempPlayer = Mid$(strData(i), 17, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "DJINNBOOSTDEFENSE"
    End If
    
    If Left$(strData(i), 15) = "DJINNDROPATTACK" Then
        strTempPlayer = Mid$(strData(i), 16, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "DJINNREDUCEAP"
    End If
    
    If Left$(strData(i), 7) = "DJINNPP" Then
        strTempPlayer = Mid$(strData(i), 8, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "DJINNBOOSTPP"
    End If
    
    If Left$(strData(i), 12) = "DJINNDEFENSE" Then
        strTempPlayer = Mid$(strData(i), 13, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "DJINNBOOSTDEFENSE"
    End If
    
    If Left$(strData(i), 11) = "DJINNATTACK" Then
        strTempPlayer = Mid$(strData(i), 12, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "DJINNBOOSTAP"
    End If
    If Left$(strData(i), 11) = "DJINNRESIST" Then
        strTempPlayer = Mid$(strData(i), 12, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "DJINNBOOSTRESIST"
    End If
    If Left$(strData(i), 15) = "DJINNDROPRESIST" Then
        strTempPlayer = Mid$(strData(i), 16, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "DJINNREDUCERESIST"
    End If
    
    If Left$(strData(i), 8) = "SETDJINN" Then
        strTempPlayer = Mid$(strData(i), 9, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "SETDJINN"
    End If
    
    If Left$(strData(i), 9) = "DJINNTYPE" Then
        strTempPlayer = Mid$(strData(i), 10, 1)
        intTempPlayer = CInt(strTempPlayer)
        DjinnElement(intTempPlayer) = Mid(strData(i), 11, Len(strData(i)))
    End If
    
    If Left$(strData(i), 5) = "RESET" Then
        Reset(2) = True
    End If
    
    If Left$(strData(i), 10) = "DONEARENA1" Then 'Player finished the arena
        If winRace = 0 And BattleLoaded(1) = False And BattleLoaded(2) = False Then
            opFinished = True
            winRace = 2
        End If
    End If

    If Left$(strData(i), 12) = "BATTLELOADED" Then 'Player finished the arena
        BattleLoaded(2) = True
        If BattleLoaded(1) = False And DataSent = False Then
            Call LoadBattle
        End If
        'frmArena.timecount.Enabled = False
        'Unload frmArena
        'frmBattle.Show
        'Call LoadBattle
    End If
    
    If Left$(strData(i), 6) = "SUMMON" Then
        strTempPlayer = Mid$(strData(i), 7, 1)
        intTempPlayer = CInt(strTempPlayer)
        iSummonType(intTempPlayer) = CInt(Mid(strData(i), 8, Len(strData(i))))
        AttackType(intTempPlayer) = "SUMMON"
    End If
    
    If Left$(strData(i), 5) = "ILOST" Then 'Opponent lost
    
        frmHost.Host.Close
    
        HP(3) = 0
        HP(4) = 0
        
        DidIWin = True 'I win
        GameOver = True
        
        frmBattle.timeWait.Enabled = True
    
    
'        'Disable timers on this form
'        frmBattle.timePsynergy.Enabled = False
'        frmBattle.timeBoready.Enabled = False
        'LADDER TOURNAMENT
        frmEndGame.Show
        DataSent = False 'Did not send the stat increase to the server yet
        Unload frmBattle
    
    End If
    
    If Left$(strData(i), 4) = "IWIN" Then 'Opponent won
    
    
        HP(1) = 0
        HP(2) = 0
        
        frmBattle.timeWait.Enabled = True
        
        frmHost.Host.Close
    
        DidIWin = False 'I win
        GameOver = True
    
    
'        'Disable timers on this form
'        frmBattle.timePsynergy.Enabled = False
'        frmBattle.timeBoready.Enabled = False
        'LADDER TOURNEMANT:
        frmEndGame.Show
        DataSent = False 'Did not send the stat increase to the server yet
        Unload frmBattle

    End If
    
    If Left$(strData(i), 8) = "RESEARTH" Then
        strTempPlayer = Mid$(strData(i), 9, 1)
        intTempPlayer = CInt(strTempPlayer)
        intEarthResist(intTempPlayer) = CStr(Mid(strData(i), 10, Len(strData(i))))
    End If
    If Left$(strData(i), 7) = "RESFIRE" Then
        strTempPlayer = Mid$(strData(i), 8, 1)
        intTempPlayer = CInt(strTempPlayer)
        intFireResist(intTempPlayer) = CStr(Mid(strData(i), 9, Len(strData(i))))
    End If
    If Left$(strData(i), 7) = "RESWIND" Then
        strTempPlayer = Mid$(strData(i), 9, 1)
        intTempPlayer = CInt(strTempPlayer)
        intWindResist(intTempPlayer) = CStr(Mid(strData(i), 9, Len(strData(i))))
    End If
    If Left$(strData(i), 8) = "RESWATER" Then
        strTempPlayer = Mid$(strData(i), 9, 1)
        intTempPlayer = CInt(strTempPlayer)
        intWaterResist(intTempPlayer) = CStr(Mid(strData(i), 10, Len(strData(i))))
    End If
    If Left$(strData(i), 7) = "RESDARK" Then
        strTempPlayer = Mid$(strData(i), 9, 1)
        intTempPlayer = CInt(strTempPlayer)
        intDarkResist(intTempPlayer) = CStr(Mid(strData(i), 9, Len(strData(i))))
    End If
    If Left$(strData(i), 8) = "RESHEART" Then
        strTempPlayer = Mid$(strData(i), 9, 1)
        intTempPlayer = CInt(strTempPlayer)
        intHeartResist(intTempPlayer) = CStr(Mid(strData(i), 10, Len(strData(i))))
    End If
    If Left$(strData(i), 8) = "POWEARTH" Then
        strTempPlayer = Mid$(strData(i), 9, 1)
        intTempPlayer = CInt(strTempPlayer)
        intEarthPower(intTempPlayer) = CStr(Mid(strData(i), 10, Len(strData(i))))
    End If
    If Left$(strData(i), 7) = "POWFIRE" Then
        strTempPlayer = Mid$(strData(i), 9, 1)
        intTempPlayer = CInt(strTempPlayer)
        intFirePower(intTempPlayer) = CStr(Mid(strData(i), 9, Len(strData(i))))
    End If
    If Left$(strData(i), 7) = "POWWIND" Then
        strTempPlayer = Mid$(strData(i), 9, 1)
        intTempPlayer = CInt(strTempPlayer)
        intWindPower(intTempPlayer) = CStr(Mid(strData(i), 9, Len(strData(i))))
    End If
    If Left$(strData(i), 8) = "POWWATER" Then
        strTempPlayer = Mid$(strData(i), 9, 1)
        intTempPlayer = CInt(strTempPlayer)
        intWaterPower(intTempPlayer) = CStr(Mid(strData(i), 10, Len(strData(i))))
    End If
    If Left$(strData(i), 7) = "POWDARK" Then
        strTempPlayer = Mid$(strData(i), 9, 1)
        intTempPlayer = CInt(strTempPlayer)
        intDarkPower(intTempPlayer) = CStr(Mid(strData(i), 10, Len(strData(i))))
    End If
    If Left$(strData(i), 8) = "POWHEART" Then
        strTempPlayer = Mid$(strData(i), 9, 1)
        intTempPlayer = CInt(strTempPlayer)
        intHeartPower(intTempPlayer) = CStr(Mid(strData(i), 11, Len(strData(i))))
    End If
    If Left$(strData(i), 4) = "HCAP" Then
        Handicap(2) = CInt(Mid(strData(i), 5, Len(strData(i))))
        lblOpHandicap = Handicap(2)
    End If
    
    
    If Left$(strData(i), 9) = "RELRATING" Then
        RelativeRating(2) = CInt(Mid(strData(i), 10, Len(strData(i))))
    End If
    If Left$(strData(i), 6) = "RELLVL" Then
        RelativeLVL(2) = CInt(Mid(strData(i), 7, Len(strData(i))))
    End If
    If Left$(strData(i), 8) = "RELDJINN" Then
        RelativeDjinn(2) = CInt(Mid(strData(i), 9, Len(strData(i))))
    End If
    If Left$(strData(i), 8) = "POSMAZEX" Then
        intOpMazeX = CInt(Mid(strData(i), 9, Len(strData(i))))
    End If
    If Left$(strData(i), 8) = "POSMAZEY" Then
        intOpMazeY = CInt(Mid(strData(i), 9, Len(strData(i))))
    End If
    
    
    
    If strData(i) = "PING" Or strData(1) = "BATTLELOADED" Or Left$(strData(i), 8) = "POSMAZEY" Or Left$(strData(i), 8) = "POSMAZEX" Then
        bWrite = False
    ElseIf (strData(i) <> "") Then
        bWrite = True
        frmChat.Chat.SendData "LADDERTOURNAMENT" & strOpponent & ": " & strData(i) & vbCrLf
    End If
Next 'i

If bWrite = True Then
    Call WriteIni("GEN", strTime, strdatao, App.Path & "\hostdump.ini")
End If

Exit Sub
err:
Exit Sub 'Don't let an errant packet check

End Sub

Private Sub imgHelp_Click()
MsgBox "What a Handicapp does is increase or decrease your relative Level, Rating Points and Djinn to make a match more even.  A player who plays at a low Handicapp tends to get more stats at the end of a match than a player who plays at a higher Handicapp, but the match is harder.  Both players must agree on a Handicapp before the game begins."
End Sub

Private Sub opEqualize_Click(Index As Integer)
On Error Resume Next
timeWait.Enabled = False
DoEvents
timeWait.Enabled = True
If Index = 0 Then
    Handicap(1) = 0
    Handicap(2) = CStr(CInt(strLvl) - CInt(stroLvl))
    If Me.chkGen(3).Value = 1 Then
        Host.SendData "HCAP" & "0" & vbCrLf
        Host.SendData "YOURHCAP" & CStr(CInt(strLvl) - CInt(stroLvl)) & vbCrLf
    End If
Else
    Handicap(1) = CInt(stroLvl) - CInt(strLvl)
    Handicap(2) = 0
    If Me.chkGen(3).Value = 1 Then
        Host.SendData "HCAP" & CStr(Handicap(1)) & vbCrLf
        Host.SendData "YOURHCAP" & "0" & vbCrLf
    End If
End If

End Sub

Private Sub timeWait_Timer()
timeWait.Enabled = False
End Sub

Private Sub txtmsg_KeyDown(KeyCode As Integer, Shift As Integer)
If keyascii = 13 Then keyascii = 0
If KeyCode = vbKeyReturn Then
Call cmdSend_Click
End If
End Sub
Sub AcceptClient()
'On Error Resume Next
If Host.RemoteHostIP <> Host.LocalIP Then
    lblmsg.Caption = "SERVER MESSAGE: Client connected!"
    txtMsg.Enabled = True
    cmdSend.Enabled = True
    cmdStart.Enabled = True
    'Commented out for LADDER TOURNAMENT
    hHandicap.Enabled = True
    
    chkGen(0).Value = 0
    chkGen(1).Value = 0
    'chkGen(0).Enabled = True
    'chkGen(1).Enabled = True
    
    cmdArena.Enabled = True
    cmdBoot.Enabled = True
    Host.SendData "USER" & strMyUserName & vbCrLf
    hoston = True
    Host.SendData "MYRATING" & strRating & vbCrLf
    Host.SendData "MYLEVEL" & strLvl & vbCrLf
    DoEvents
    
    MsgBox "Client connected!"
    'Comment out for ladder tournament
    For i = 0 To chkGen.UBound
        chkGen(i).Enabled = True
    Next 'i
    
    

    frmChat.Chat.SendData "CLOSEGAME" & vbCrLf
Else
    If Host.RemoteHostIP = IKILLKENNYIP Then
        'Host.Accept requestID
        lblmsg.Caption = "SERVER MESSAGE: Client connected!"
        txtMsg.Enabled = True
        cmdSend.Enabled = True
        cmdStart.Enabled = True
        chkGen(0).Enabled = True
        chkGen(1).Enabled = True
        cmdBoot.Enabled = True
        cmdArena.Enabled = True
        Host.SendData "USER" & strMyUserName & vbCrLf
        hoston = True
        MsgBox "Client connected!"
    Else
        lblmsg.Caption = "SERVER MESSAGE: You are not allowed to play yourself."
        Host.Close
    End If
End If



End Sub

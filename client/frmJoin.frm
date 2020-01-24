VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmJoin 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Golden Sun: The War of the Adepts - Join"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   Icon            =   "frmJoin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hHandicap 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      Max             =   5
      Min             =   -5
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox chkReady 
      BackColor       =   &H000000FF&
      Caption         =   "I'm Ready"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtmsg 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      MaxLength       =   200
      TabIndex        =   4
      Top             =   3840
      Width           =   5655
   End
   Begin VB.TextBox txtip 
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Look For Game"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Client 
      Left            =   6480
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9788
   End
   Begin VB.Image imgBG 
      Height          =   375
      Left            =   6000
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblEqualWeapons 
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   36
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Equalized Weapons:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   17
      Left            =   120
      TabIndex        =   35
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblAttacks 
      BackStyle       =   0  'Transparent
      Caption         =   "Allowed"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5160
      TabIndex        =   34
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Normal Attacks:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   16
      Left            =   2760
      TabIndex        =   33
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label lblHealing 
      BackStyle       =   0  'Transparent
      Caption         =   "Allowed"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4080
      TabIndex        =   32
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Healing:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   15
      Left            =   2760
      TabIndex        =   31
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblPsynergy 
      BackStyle       =   0  'Transparent
      Caption         =   "Allowed"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   30
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Psynergy:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   14
      Left            =   2760
      TabIndex        =   29
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblDrag 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   4800
      Width           =   135
   End
   Begin VB.Label lblSummon 
      BackStyle       =   0  'Transparent
      Caption         =   "Allowed"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   27
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Summons:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   2760
      TabIndex        =   26
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblOpLevel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   25
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblOpRating 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   24
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblOpName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   23
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   22
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Rating:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   21
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Width           =   855
   End
   Begin VB.Image imgHelp 
      Height          =   240
      Left            =   3600
      Picture         =   "frmJoin.frx":08CA
      Top             =   3555
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblOpHandicap 
      BackStyle       =   0  'Transparent
      Caption         =   "False"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   5400
      TabIndex        =   19
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Levels Equalized:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   4
      Left            =   2640
      TabIndex        =   18
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label lblHandicap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   17
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Handicap:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   9
      Left            =   4320
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Opponent's Stats:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "False"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Index           =   8
      Left            =   5640
      TabIndex        =   13
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Double Or Nothing:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Index           =   7
      Left            =   2760
      TabIndex        =   12
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "True"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   4560
      TabIndex        =   11
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Limit:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   2760
      TabIndex        =   10
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Game Options:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblstatus 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblmsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Attempting to join game..."
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   4200
      Width           =   5655
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Join A Game:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Host IP Address:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
End
Attribute VB_Name = "frmJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurUserName As String
Dim strData() As String

Private Sub chkReady_Click()
On Error Resume Next
If chkReady.Value = 1 Then
Client.SendData "READY" & vbCrLf
hHandicap.Enabled = False
Else
Client.SendData "NOTREADY" & vbCrLf
hHandicap.Enabled = True
End If
End Sub

Private Sub Client_Connect()
Debug.Print "CONNECT"
Call AcceptHost
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
On Error GoTo err

Dim bWrite As Boolean
Dim strTime As String
strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")

Dim strdatao As String
Client.GetData strdatao
strData = Split(strdatao, vbCrLf, -1, vbTextCompare)
bWrite = True
For i = 0 To UBound(strData)

    If Left$(strData(i), 8) = "GETSTATS" Then
        Client.SendData "HP" & (i + 2) & HP(i) & vbCrLf & "SPEED" & (i + 2) & Speed(i) & vbCrLf & "AP" & (i + 2) & AP(i) & vbCrLf & "DEFENSE" & (i + 2) & Defense(i) & vbCrLf & "PIC" & (i + 2) & Char(i) & vbCrLf & "CHARTYPE" & (i + 2) & CharType(i) & vbCrLf & "RESEARTH" & (i + 2) & intEarthResist(i) & vbCrLf & "RESFIRE" & (i + 2) & intFireResist(i) & vbCrLf & "RESWIND" & (i + 2) & intWindResist(i) & vbCrLf & "RESWATER" & (i + 2) & intWaterResist(i) & vbCrLf & "RESHEART" & (i + 2) & intHeartResist(i) & vbCrLf & "RESDARK" & (i + 2) & intDarkResist(i) & vbCrLf
        DoEvents
    End If

    If Left$(strData(i), 6) = "NEEDPW" Then
        Dim strGamePW As String
        strGamePW = InputBox("Please enter the password for this game.")
        Client.SendData "NEEDPW" & strGamePW & vbCrLf
    End If
    If Left$(strData(i), 6) = "GOODPW" Then
        Call AcceptHost
    End If
    If Left$(strData(i), 5) = "BADPW" Then
        Client.Close
        MsgBox "Your password was incorrect."
        Call Form_Load
    End If
    
    If Left$(strData(i), 3) = "MSG" Then
        lblmsg.Caption = strOpponent & ": " & Mid(strData(i), 4, Len(strData(i)))
    End If
    If Left$(strData(i), 4) = "MYHP" Then 'Foe's HP
        strTempPlayer = Mid$(strData(i), 5, 1)
        intTempPlayer = CInt(strTempPlayer)
        HP(intTempPlayer) = CInt(Mid$(strData(i), 6, Len(strData(i))))
    End If
    If Left$(strData(i), 5) = "SPEED" Then 'Opponent's AP
        strTempPlayer = Mid$(strData(i), 6, 1)
        intTempPlayer = CInt(strTempPlayer)
        strspeed = Mid(strData(i), 7, Len(strData(i)))
        Speed(intTempPlayer) = CInt(strspeed)
    End If
    If Left$(strData(i), 7) = "NOBEANS" Then
        MsgBox "The host has rejected your request to join the game."
    End If
    
    If Left$(strData(i), 4) = "GMSG" Then
        frmBattle.lblmsg.Caption = strOpponent & ": " & Mid(strData(i), 5, Len(strData(i)))
    End If
    
    If Left$(strData(i), 4) = "USER" Then
        strOpponent = Mid(strData(i), 5, Len(strData(i)))
        lblOpName.Caption = strOpponent
        Dim bCheckOp As Boolean
        bCheckOp = CheckOpponent(strOpponent, Client.RemoteHostIP)
        If bCheckOp = True Then
            MsgBox "You are not allowed to play someone more than three times in a single day.  Please wait until tommorow (12 AM EST) to play this person again."
            Client.Close
        End If
    End If
    
    If Left$(strData(i), 8) = "MYRATING" Then
        opRating = Mid(strData(i), 9, Len(strData(i)))
        lblOpRating.Caption = opRating
    End If
    
    If Left$(strData(i), 7) = "MYLEVEL" Then
        stroLvl = Mid(strData(i), 8, Len(strData(i)))
        lblOpLevel.Caption = stroLvl
    End If
    
    
    
    If Left$(strData(i), 6) = "TIMEON" Then
        lblgen(6).Caption = "On"
    End If
    
    If Left$(strData(i), 7) = "TIMEOFF" Then
        lblgen(6).Caption = "Off"
    End If
    
    If Left$(strData(i), 6) = "RATEON" Then
        lblgen(8).Caption = "On"
        DoubleStats = True
    End If
    
    If Left$(strData(i), 7) = "RATEOFF" Then
        lblgen(8).Caption = "Off"
        DoubleStats = False
    End If
    
    If Left$(strData(i), 5) = "START" Then
        bMazeFirstLoad = True
        Unload frmMultiplayer
        StopMidi
        BattleLoaded(1) = False
        BattleLoaded(2) = False
        bMazeError = False
        frmJoin.Hide
        bMazeError = False
        'frmArena.Show
        hoston = False
        If lblgen(6).Caption = "Enabled" Then
            TimedMatch = True
        Else
            TimedMatch = False
        End If
    End If

    If Left$(strData(i), 8) = "CHARTYPE" Then
        CharType(2) = Mid(strData(i), 9, Len(strData(i)))
    End If

    If Left$(strData(i), 7) = "OPREADY" Then
        strTempPlayer = Mid$(strData(i), 8, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True 'Opponent is ready
        If bOReady(2) = True And bOReady(4) = True Then
            Call DoAttacks
        End If
    End If
    
    If Left$(strData(i), 6) = "TARGET" Then
        strTempPlayer = Mid$(strData(i), 7, 1)
        intTempPlayer = CInt(strTempPlayer)
        Target(intTempPlayer) = CLng(Mid$(strData(i), 8, Len(strData(i))))
    End If
    
    If Left$(strData(i), 3) = "LVL" Then
        stroLvl = Mid(strData(i), 4, Len(strData(i)))
        intoLvl = CInt(stroLvl)
    End If
    
    If Left$(strData(i), 6) = "RATING" Then
        opRating = CInt(Mid(strData(i), 7, Len(strData(i))))
    
    End If
    
    If Left$(strData(i), 4) = "CHAR" Then
        strTempPlayer = Mid$(strData(i), 5, 1)
        intTempPlayer = CInt(strTempPlayer)
        
        CharName(intTempPlayer) = Mid(strData(i), 6, Len(strData(i)))
        bCustomChar(2) = FindWhichCharacter(CharName(intTempPlayer))
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
        Else
            If bCustomChar(4) = 999 Or bCustomChar(4) = 0 Then
                stroChar = Mid(strData(i), 5, Len(strData(i)))
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
        AttackType(intTempPlayer) = "DJINNREDUCEDEFENSE"
    End If
    
    If Left$(strData(i), 15) = "DJINNDROPATTACK" Then
        strTempPlayer = Mid$(strData(i), 16, 1)
        intTempPlayer = CInt(strTempPlayer)
        bOReady(intTempPlayer) = True
        AttackType(intTempPlayer) = "DJINNDROPAP"
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
    
    If Left$(strData(i), 4) = "DISC" Then 'Disconnect (not currently implemented)
        Unload frmBattle
        BattleLoaded(1) = False
        BattleLoaded(2) = False
        Client.Close
        lblmsg.Caption = "SERVERMSG: The other party disconnected!"
        txtip.Enabled = True
        txtmsg.Enabled = False
        cmdListen.Enabled = True
        cmdSend.Enabled = False
        frmIntro.Show
    End If

    If Left$(strData(i), 4) = "TIME" Then 'Current time left in gameclock
        strtimeleft = Mid(strData(i), 5, Len(strData(i)))
        If strtimeleft <> "ON" And strtimeleft <> "OFF" Then 'Turn the clock on or off?
            curCount = CInt(strtimeleft)
            frmBattle.lblgen(18).Caption = curCount 'Update label to current time
            If curCount = 0 Then 'If you're out of time then defend
                If bOReady(1) = False Then
                    Client.SendData "BOREADY" & vbCrLf
                    Client.SendData "DEFEND" & vbCrLf
                End If
                curCount = 20
                frmBattle.lblgen(18).Caption = "20"
            End If
        End If
    End If
    
    If Left$(strData(i), 5) = "ARENA" Then 'Load this background image
        strarena = Mid(strData(i), 6, Len(strData(i)))
        'Change for ladder tournament:
        'frmBattle.imgArena.Picture = LoadPicture(App.Path & "\Ladarena" & strarena & ".gif")
        If strarena = "1" Then
            imgBG.Picture = LoadPicture(App.Path & "\arena1.gif")
        ElseIf strarena = "2" Then
            imgBG.Picture = LoadPicture(App.Path & "\arena2.gif")
        ElseIf strarena = "3" Then
            imgBG.Picture = LoadPicture(App.Path & "\arena3.gif")
        ElseIf strarena = "4" Then
            imgBG.Picture = LoadPicture(App.Path & "\arena4.gif")
        ElseIf strarena = "5" Then
            imgBG.Picture = LoadPicture(App.Path & "\arena5.gif")
        ElseIf strarena = "6" Then
            imgBG.Picture = LoadPicture(App.Path & "\arena6.gif")
        End If
    End If

    If Left$(strData(i), 9) = "LOADARENA" Then 'Load the battle arena form
        frmJoin.Hide
        hoston = False
        'frmArena.Show
    End If
    If Left$(strData(i), 8) = "MAPARENA" Then 'Arena to load
        strMaptoLoad = Mid$(strData(i), 9, Len(strData(i)))
        frmJoin.Hide
        'hoston = False
        bMazeFirstLoad = True
        bMazeError = False
        'frmArena.Show
        'frmArena.timecount.Enabled = True
        Call LoadBattle
        frmBattle.Show
    End If


    If Left$(strData(i), 10) = "DONEARENA1" Then 'Player finished the arena
        If BattleLoaded(1) = False And BattleLoaded(2) = False Then
            opFinished = True
            winRace = 2
        End If
    End If
    
    If Left$(strData(i), 12) = "BATTLELOADED" Then 'Player finished the arena
        BattleLoaded(2) = True
        If BattleLoaded(1) = True And DataSent = False Then
            Call LoadBattle
        End If
        'frmBattle.Show
        'Call LoadBattle
    End If

    'If Left$(strData(i), 9) = "UPDATEHP1" Then 'Makes sure that the HP is the same on the client as the host
    '    HP(1) = CStr(Mid(strData(i), 10, Len(strData(i))))
    'End If
    
    'If Left$(strData(i), 9) = "UPDATEHP2" Then
    '    HP(2) = CStr(Mid(strData(i), 10, Len(strData(i))))
    '    frmBattle.lblgen(9).Caption = HP(1)
    '    frmBattle.shpHP(0).Width = HP(1) / 4
    '    frmBattle.lblgen(5).Caption = HP(2)
    '    frmBattle.shpHP(1).Width = HP(2) / 4
    'End If


    If Left$(strData(i), 4) = "HCAP" Then
        Handicap(2) = CInt(Mid(strData(i), 5, Len(strData(i)))) <> 0
        If Handicap(2) <> 0 Then
            lblOpHandicap.Caption = "At Host"
        End If
        
    End If
    If Left$(strData(i), 8) = "YOURHCAP" Then
        Handicap(1) = CInt(Mid$(strData(i), 9, Len(strData(i))))
        If Handicap(1) <> 0 Then
            lblOpHandicap.Caption = "At Client"
        End If
        
    End If
    If Left$(strData(i), 8) = "EQUALIZE" Then
        If Mid$(strData(i), 9, 1) = "1" Then
            lblOpHandicap.Caption = "True"
        Else
            lblOpHandicap.Caption = "False"
        End If
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
    
    If Left$(strData(i), 10) = "ALLOWSUMON" Then
        bAllowSummon = True
        lblSummon.Caption = "Allowed"
    End If
    If Left$(strData(i), 11) = "ALLOWSUMOFF" Then
        bAllowSummon = False
        lblSummon.Caption = "Disallowed"
    End If
    
    If Left$(strData(i), 5) = "PSYON" Then
        bAllowPsynergy = True
        lblPsynergy.Caption = "Allowed"
    End If
    If Left$(strData(i), 6) = "PSYOFF" Then
        bAllowPsynergy = False
        lblPsynergy.Caption = "Disallowed"
    End If
    If Left$(strData(i), 8) = "ATTACKON" Then
        bAllowAttack = True
        lblAttacks.Caption = "Allowed"
    End If
    If Left$(strData(i), 9) = "ATTACKOFF" Then
        bAllowAttack = False
        lblAttacks.Caption = "Disallowed"
    End If
    If Left$(strData(i), 6) = "HEALON" Then
        bAllowHeal = True
        lblHealing.Caption = "Allowed"
    End If
    If Left$(strData(i), 7) = "HEALOFF" Then
        bAllowHeal = False
        lblHealing.Caption = "Disallowed"
    End If
    If Left$(strData(i), 10) = "EWEAPONSON" Then
        bEqualizeWeapons = True
        lblEqualWeapons.Caption = "On"
    End If
    If Left$(strData(i), 11) = "EWEAPONSOFF" Then
        bEqualizeWeapons = False
        lblEqualWeapons.Caption = "Off"
    End If
    
    If Left$(strData(i), 8) = "POSMAZEX" Then
        intOpMazeX = CInt(Mid(strData(i), 9, Len(strData(i))))
    End If
    If Left$(strData(i), 8) = "POSMAZEY" Then
        intOpMazeY = CInt(Mid(strData(i), 9, Len(strData(i))))
    End If
    
    
    If strData(i) = "PING" Or strData(1) = "BATTLELOADED" Or Left$(strData(i), 8) = "POSMAZEY" Or Left$(strData(i), 8) = "POSMAZEX" Then
        bWrite = False
    ElseIf strData(i) <> "" Then
        bWrite = True
        Debug.Print "CLIENTDATA - " & strData(i)
        frmChat.Chat.SendData "LADDERTOURNAMENT" & strOpponent & ": " & strData(i) & vbCrLf
    End If
    
    

'Debug.Print strdata(i)
Next 'i

If bWrite = True Then
    Call WriteIni("GEN", strTime, strdatao, App.Path & "\joindump.ini")
End If


Exit Sub
err:
Exit Sub

End Sub

Private Sub cmdListen_Click()
On Error Resume Next
If Client.State <> sckClosed And txtip.Text <> "127.0.0.1" Then
Client.Close
DoEvents
End If
If txtip.Text = "root" Then
Client.Connect IKILLKENNYIP, Client.RemotePort
Else
Client.Connect txtip.Text, Client.RemotePort
End If
End Sub

Private Sub cmdSend_Click()
On Error Resume Next
Client.SendData "MSG" & txtmsg.Text & vbCrLf
txtmsg.Text = ""
End Sub

Private Sub Form_Activate()
If Client.State = sckClosed Then
    Call Form_Load
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
bMazeFirstLoad = False
'Comment out for Ladder Tournament
BattleLoaded(1) = False
BattleLoaded(2) = False
hHandicap.Enabled = True
hHandicap.Value = 0
lblOpHandicap.Caption = "False"
Client.RemotePort = 9788
txtmsg.Enabled = False
cmdSend.Enabled = False
lblmsg.Caption = "Attempting to join game."
chkReady.Enabled = False
chkReady.Value = 0
DataSent = False

Unload frmBattle
Unload frmArena

strImage = GetFromIni("GEN", "IMAGES", App.Path & "\settings.ini")
If strImage = "ON" Then
    Me.Picture = frmIntro.Picture
End If

lblgen(8).Caption = "False"
lblgen(6).Caption = "True"

txtip.Text = strJoinIP


chkReady.Value = 0
lblOpName.Caption = ""
lblOpRating.Caption = ""
lblOpLevel.Caption = ""
bAllowSummon = True
lblSummon.Caption = "Allowed"
bAllowHeal = True
lblHealing.Caption = "Allowed"
bAllowAttack = True
lblAttacks.Caption = "Allowed"
bAllowPsynergy = True
lblPsynergy.Caption = "Allowed"


DoubleStats = False

If Client.State <> sckClosed Then
    Client.Close
    DoEvents
End If

txtip.Text = strJoinIP

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmIntro.Show
End Sub


Private Sub lbluser_Click()

End Sub

Private Sub Timer1_Timer()
lblStatus.Caption = Client.State
End Sub

Private Sub hHandicap_Change()
On Error Resume Next
lblHandicap.Caption = hHandicap.Value
Client.SendData "HCAP" & hHandicap.Value & vbCrLf
Handicap(1) = hHandicap.Value
End Sub

Private Sub imgHelp_Click()
MsgBox "What a Handicapp does is increase or decrease your relative Level, Rating Points and Djinn to make a match more even.  A player who plays at a low Handicapp tends to get more stats at the end of a match than a player who plays at a higher Handicapp, but the match is harder.  Both players must agree on a Handicapp before the game begins."
End Sub

Private Sub lblDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    lblDrag.Drag
End If
End Sub

Private Sub lblgen_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
If Index = 8 And lblgen(Index).Caption = "True" Then
    If CurEgg26 = 3 Then
        CurEgg26 = 4
        Call PlySound("explosion")
    Else
        CurEgg26 = 1
    End If
End If
End Sub

Private Sub txtmsg_KeyDown(KeyCode As Integer, Shift As Integer)
If keyascii = 13 Then keyascii = 0
If KeyCode = vbKeyReturn Then
Call cmdSend_Click
End If
End Sub
Sub AcceptHost()
On Error Resume Next
Beep
cmdListen.Enabled = False
txtip.Enabled = False
lblmsg.Caption = "SERVER MESSAGE: Good connection!"
txtmsg.Enabled = True
cmdSend.Enabled = True
Client.SendData "USER" & strMyUserName & vbCrLf
Client.SendData "NOTREADY" & vbCrLf
Client.SendData "MYRATING" & strRating & vbCrLf
Client.SendData "MYLEVEL" & strLvl & vbCrLf
frmChat.Chat.SendData "METXT" & strMyUserName & " has joined a game." & vbCrLf
bMazeFirstLoad = True
winRace = 0

Handicap(1) = 0
hHandicap.Value = 0



chkReady.Enabled = True
hoston = False
End Sub

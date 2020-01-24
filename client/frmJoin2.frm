VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmJoin2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "War of the Adepts: Join Game"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8025
   Icon            =   "frmJoin2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGameIP 
      Height          =   285
      Left            =   3720
      MaxLength       =   35
      TabIndex        =   21
      Top             =   0
      Width           =   1695
   End
   Begin VB.ListBox lstGameList 
      Height          =   645
      ItemData        =   "frmJoin2.frx":08CA
      Left            =   840
      List            =   "frmJoin2.frx":08CC
      TabIndex        =   19
      Top             =   0
      Width           =   1815
   End
   Begin VB.Frame framGen 
      BackColor       =   &H00886000&
      Caption         =   "Your Stats"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   3975
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ladder Rating:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   1050
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   3
         Left            =   3000
         Picture         =   "frmJoin2.frx":08CE
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   2
         Left            =   2040
         Picture         =   "frmJoin2.frx":0FA4
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   1
         Left            =   1080
         Picture         =   "frmJoin2.frx":167A
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   0
         Left            =   120
         Picture         =   "frmJoin2.frx":1D50
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblRanking 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   31
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblLoss 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   3360
         TabIndex        =   30
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblWins 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblRating 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rating:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   510
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wins:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   15
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Losses:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   14
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   405
      End
   End
   Begin VB.Frame framGen 
      BackColor       =   &H00886000&
      Caption         =   "Your Opponent's Stats"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Index           =   1
      Left            =   3960
      TabIndex        =   7
      Top             =   720
      Width           =   3975
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ladder Rating:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   1050
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   7
         Left            =   2760
         Picture         =   "frmJoin2.frx":2426
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   6
         Left            =   1800
         Picture         =   "frmJoin2.frx":2AFC
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   5
         Left            =   840
         Picture         =   "frmJoin2.frx":31D2
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image imgParty 
         Height          =   960
         Index           =   4
         Left            =   0
         Picture         =   "frmJoin2.frx":38A8
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblRanking 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   36
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblLoss 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   3360
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblWins 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   34
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblRating 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   33
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rating:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   510
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wins:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Losses:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   405
      End
   End
   Begin VB.Frame framGen 
      BackColor       =   &H00886000&
      Caption         =   "Chat"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   4320
      Width           =   7935
      Begin VB.TextBox txtChat 
         BackColor       =   &H00886000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   240
         Width           =   7575
      End
      Begin VB.TextBox txtChatMsg 
         BackColor       =   &H00886000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         MaxLength       =   200
         TabIndex        =   5
         Top             =   1080
         Width           =   7575
      End
   End
   Begin VB.Frame framGen 
      BackColor       =   &H00886000&
      Caption         =   "Game Options"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Index           =   3
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   7935
      Begin VB.CheckBox chkReady 
         BackColor       =   &H00886000&
         Caption         =   "Ready"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblPartyLevel 
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1200
         TabIndex        =   27
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblTimedMatch 
         BackStyle       =   0  'Transparent
         Caption         =   "False"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3000
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLadderMatch 
         BackStyle       =   0  'Transparent
         Caption         =   "False"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1320
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party Level:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   840
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Timed Match:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   13
         Left            =   1920
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ladder Match:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Join Game"
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtGamePW 
      Height          =   285
      Left            =   3720
      MaxLength       =   15
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock Client 
      Left            =   7440
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9688
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   2880
      TabIndex        =   20
      Top             =   0
      Width           =   810
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Game List:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   750
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   2880
      TabIndex        =   17
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmJoin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkReady_Click()
On Error Resume Next
If chkReady.Value = 0 Then
    Client.SendData "NOTREADY" & vbCrLf
Else
    Client.SendData "READY" & vbCrLf
End If
End Sub

Private Sub Client_Connect()
On Error Resume Next
txtChat.Text = txtChat.Text & vbNewLine & "Connected to " & txtGameIP.Text & "."
Client.SendData "USERNAME" & strMyUserName & vbCrLf
Client.SendData "RATING" & strMyRating & vbCrLf
Client.SendData "WINS" & strMyWins & vbCrLf
Client.SendData "LOSS" & strMyLoss & vbCrLf
Client.SendData "RANKING" & strMyRanking & vbCrLf
For i = 1 To 4
    Client.SendData "PARTYNAME" & CStr(i) & WOTAChar(i).Name & vbCrLf
Next 'i

End Sub

Private Sub Client_ConnectionRequest(ByVal requestID As Long)
On Error GoTo err
Dim strRawData As String

Dim strTime As String
strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")

Client.GetData strRawData
strData = Split(strRawData, vbCrLf, -1, vbTextCompare) 'Split packets between @'s

For i = 0 To UBound(strData)
    If Left$(strData(i), 8) = "USERNAME" Then
        strFoeUserName = Mid$(strData(i), 9, Len(strData(i)))
        framGen(1).Caption = strFoeUserName & "'s stats."
    End If
    If Left$(strData(i), 6) = "RATING" Then
        strFoeRating = Mid$(strData(i), 7, Len(strData(i)))
        lblRating(1).Caption = strFoeRating
    End If
    If Left$(strData(i), 4) = "WINS" Then
        strFoeWins = Mid$(strData(i), 5, Len(strData(i)))
        lblWins(1).Caption = strFoeWins
    End If
    If Left$(strData(i), 4) = "LOSS" Then
        strFoeLoss = Mid$(strData(i), 5, Len(strData(i)))
        lblLoss(1).Caption = strFoeLoss
    End If
    If Left$(strData(i), 6) = "RANKING" Then
        strFoeRanking = Mid$(strData(i), 7, Len(strData(i)))
        lblRanking(1).Caption = strFoeRanking
    End If
    If Left$(strData(i), 9) = "PARTYNAME" Then
        Dim intTempChar As Long
        intTempChar = CLng(Mid$(strData(i), 10, 1))
        WOTAChar(intTempChar + 4).Name = Mid$(strData(i), 11, Len(strData(i)))
        imgParty(intTempChar + 4).Picture = LoadPicture(App.Path & "\BattleImages\" & WOTAChar(intTempChar + 4).Name & ".gif")
    End If
    If Left$(strData(i), 8) = "WAITCHAT" Then
        Dim strTempWaitChat As String
        strTempWaitChat = Mid$(strData(i), 9, Len(strData(i)))
        txtChat.Text = txtChat.Text & vbNewLine & strTempWaitChat
    End If
    If Left$(strData(i), 9) = "STARTGAME" Then
        Me.Hide
        frmBattle2.Show
    End If
Next 'i

Exit Sub
err:
Exit Sub
End Sub

Private Sub cmdCreate_Click()
On Error Resume Next
txtChat.Text = txtChat.Text & vbNewLine & "Attempting to join the game (" & txtGameIP.Text & ")."
Client.Connect txtGameIP.Text, Client.RemotePort
End Sub

Private Sub txtChatMsg_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    Host.SendData "WAITCHAT" & strMyUserName & ": " & txtChatMsg.Text & vbCrLf
    txtChat.Text = txtChat.Text & vbNewLine & strMyUserName & ": " & txtChatMsg.Text
    txtChatMsg.Text = ""
End If
End Sub

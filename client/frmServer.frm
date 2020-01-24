VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Configuration Tool"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   7
      Left            =   960
      TabIndex        =   81
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Timer timeMulti 
      Interval        =   2500
      Left            =   4080
      Top             =   2040
   End
   Begin VB.Timer timeServerPing 
      Interval        =   5000
      Left            =   4800
      Top             =   2160
   End
   Begin VB.TextBox txtHTML 
      Height          =   2055
      Left            =   7800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   77
      Top             =   6000
      Width           =   1695
   End
   Begin VB.ListBox lstRank 
      Height          =   2010
      ItemData        =   "frmServer.frx":0000
      Left            =   7800
      List            =   "frmServer.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   76
      Top             =   4440
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Chat 
      Index           =   0
      Left            =   480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9887
   End
   Begin VB.Timer timePing 
      Interval        =   5000
      Left            =   7200
      Top             =   2040
   End
   Begin VB.ListBox lstStatus 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      ItemData        =   "frmServer.frx":0004
      Left            =   4080
      List            =   "frmServer.frx":0006
      TabIndex        =   67
      Top             =   2520
      Width           =   4815
   End
   Begin VB.Frame framServerOptions 
      Caption         =   "Server Options"
      Height          =   3135
      Left            =   2280
      TabIndex        =   59
      Top             =   4080
      Width           =   5535
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset Entire Ladder"
         Height          =   495
         Left            =   3840
         TabIndex        =   82
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdTOS 
         Caption         =   "Refresh TOS"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdCloseServer 
         Caption         =   "Go Down For Maitenece"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txtVer 
         Height          =   285
         Left            =   840
         TabIndex        =   75
         Text            =   "0.61"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtContent 
         Height          =   285
         Left            =   1320
         TabIndex        =   73
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton cmdMsg 
         Caption         =   "Msg"
         Height          =   375
         Left            =   120
         TabIndex        =   71
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   255
         Left            =   4680
         TabIndex        =   70
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtMsg 
         Height          =   285
         Left            =   3600
         TabIndex        =   69
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdMOTD 
         Caption         =   "Save"
         Height          =   255
         Left            =   2520
         TabIndex        =   66
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdKill 
         Caption         =   "Kill"
         Height          =   375
         Left            =   1200
         TabIndex        =   65
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtKill 
         Height          =   285
         Left            =   120
         TabIndex        =   64
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtMOTD 
         Height          =   285
         Left            =   120
         MaxLength       =   350
         TabIndex        =   62
         Text            =   "Welcome to Golden Sun Anonymous' Online Battle Game Chat."
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmdGame 
         Caption         =   "Launch Game"
         Height          =   375
         Left            =   1920
         TabIndex        =   60
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label lblgen 
         Caption         =   "Version:"
         Height          =   255
         Index           =   31
         Left            =   120
         TabIndex        =   74
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label lblKill 
         Caption         =   "Msg"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   72
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblgen 
         Caption         =   "Send Server Message:"
         Height          =   495
         Index           =   30
         Left            =   3600
         TabIndex        =   68
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblKill 
         Caption         =   "User"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   63
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblgen 
         Caption         =   "Set the Message of the Day:"
         Height          =   255
         Index           =   29
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.TextBox txtSummon 
      Height          =   285
      Index           =   3
      Left            =   600
      MaxLength       =   2
      TabIndex        =   58
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   56
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox txtSummon 
      Height          =   285
      Index           =   2
      Left            =   600
      MaxLength       =   1
      TabIndex        =   55
      Top             =   5160
      Width           =   255
   End
   Begin VB.TextBox txtSummon 
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   53
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox txtSummon 
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   51
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   48
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   47
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   46
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtDjinn 
      Height          =   285
      Index           =   4
      Left            =   6360
      TabIndex        =   44
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtDjinn 
      Height          =   285
      Index           =   3
      Left            =   6480
      TabIndex        =   42
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtDjinn 
      Height          =   285
      Index           =   2
      Left            =   6360
      TabIndex        =   40
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtDjinn 
      Height          =   285
      Index           =   1
      Left            =   6000
      MaxLength       =   1
      TabIndex        =   37
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtDjinn 
      Height          =   285
      Index           =   0
      Left            =   6000
      TabIndex        =   36
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   7
      Left            =   3120
      TabIndex        =   33
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   6
      Left            =   3000
      TabIndex        =   31
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   5
      Left            =   3480
      TabIndex        =   29
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   4
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   27
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   3
      Left            =   3360
      TabIndex        =   25
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   23
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   20
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Index           =   0
      Left            =   3120
      TabIndex        =   19
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   6
      Left            =   600
      TabIndex        =   16
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   5
      Left            =   600
      TabIndex        =   13
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   12
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   3
      Left            =   720
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   1
      Left            =   600
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox txtPsy 
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Timer closeme 
      Interval        =   5000
      Left            =   7560
      Top             =   2040
   End
   Begin VB.Timer timeUpdate 
      Interval        =   10000
      Left            =   8520
      Top             =   2040
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   0
      Left            =   1680
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9888
   End
   Begin VB.Label lblgen 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   195
      Index           =   32
      Left            =   0
      TabIndex        =   80
      Top             =   3600
      Width           =   840
   End
   Begin VB.Label lblgen 
      Caption         =   "Djinn:"
      Height          =   255
      Index           =   28
      Left            =   0
      TabIndex        =   57
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Type:"
      Height          =   255
      Index           =   27
      Left            =   0
      TabIndex        =   54
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Description:"
      Height          =   255
      Index           =   26
      Left            =   0
      TabIndex        =   52
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblgen 
      Caption         =   "Name:"
      Height          =   255
      Index           =   25
      Left            =   0
      TabIndex        =   50
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblgen 
      Caption         =   "Add Summon:"
      Height          =   255
      Index           =   24
      Left            =   0
      TabIndex        =   49
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label lbltypes 
      Caption         =   $"frmServer.frx":0008
      Height          =   975
      Left            =   0
      TabIndex        =   45
      Top             =   7200
      Width           =   7695
   End
   Begin VB.Label lblgen 
      Caption         =   "Damage:"
      Height          =   255
      Index           =   23
      Left            =   5400
      TabIndex        =   43
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblgen 
      Caption         =   "Attack Type:"
      Height          =   255
      Index           =   22
      Left            =   5400
      TabIndex        =   41
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblgen 
      Caption         =   "Description:"
      Height          =   255
      Index           =   21
      Left            =   5400
      TabIndex        =   39
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblgen 
      Caption         =   "Type:"
      Height          =   255
      Index           =   20
      Left            =   5400
      TabIndex        =   38
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Name:"
      Height          =   255
      Index           =   19
      Left            =   5400
      TabIndex        =   35
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Add Djinn:"
      Height          =   255
      Index           =   18
      Left            =   5400
      TabIndex        =   34
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblgen 
      Caption         =   "Damage:"
      Height          =   255
      Index           =   17
      Left            =   2400
      TabIndex        =   32
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblgen 
      Caption         =   "Type:"
      Height          =   255
      Index           =   16
      Left            =   2400
      TabIndex        =   30
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblgen 
      Caption         =   "SPC Damage:"
      Height          =   255
      Index           =   15
      Left            =   2400
      TabIndex        =   28
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblgen 
      Caption         =   "SPC Type:"
      Height          =   255
      Index           =   14
      Left            =   2400
      TabIndex        =   26
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblgen 
      Caption         =   "SPC Name:"
      Height          =   255
      Index           =   13
      Left            =   2400
      TabIndex        =   24
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblgen 
      Caption         =   "Description:"
      Height          =   255
      Index           =   12
      Left            =   2400
      TabIndex        =   22
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblgen 
      Caption         =   "Coins:"
      Height          =   255
      Index           =   11
      Left            =   2400
      TabIndex        =   21
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Name:"
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   18
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Add Item:"
      Height          =   255
      Index           =   9
      Left            =   2400
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblgen 
      Caption         =   "Djinn:"
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   15
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Rating:"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   14
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "PP:"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   11
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Damage:"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   10
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblgen 
      Caption         =   "Type:"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Class:"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Name:"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblgen 
      Caption         =   "Add Psynergy:"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblstatus 
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblgen 
      Caption         =   "Status:"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim upTime As Integer
Dim ladderRefresh As Integer
Dim SingleUser As Integer

Dim DontAdd As Boolean

Dim MaxCon As Integer
Dim UserName(0 To 21) As String
Dim UserRating(0 To 21) As String
Dim iNewServer As Integer
Dim Game(1 To 20) As Games


Dim strTOS(1 To 15) As String


Dim ServerVersion As String
Dim UserVersion As String

Dim CurMsgName As String

Dim strMOTD As String

Dim etc As Integer

'Dim strdata As String
Dim IsLoaded(1 To 20) As Boolean
Dim noclose As Boolean
Dim NewUser As String
Dim NewPass As String
Dim isitnew As String
Dim strTime As String
Dim strRating As String
Dim curTime As Long
Dim lstUser As String
Dim lstEmail As String
Dim realPassword As String
Dim strdata() As String

Dim strChar As String
Dim strCoins As String
Dim strWins As String
Dim strLoss As String
Dim strDisc As String
Dim strLvl As String
Dim strDjinn As String
Dim strType As String
Dim strMyWeapon As String

Dim strCurRating As String
Dim strCurLvl As String
Dim strCurCoins As String
Dim iCurRating As Integer
Dim iCurLvl As Integer
Dim iCurCoins As Integer
Dim curUser As String

Dim strNewRank As String
Dim iNewRank As Integer
Dim iCurRank As Integer

Dim arrdata
Dim curPassword As String
'Dim curUser As String

Private Sub Chat_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
If Index = 0 Then
    'Beep
    iNewServer = 100
    For q = 1 To 20
    If Chat(q).State = sckClosed And iNewServer = 100 Then
        Chat(q).Accept requestID
        iNewServer = MaxCon
        Users(q).Enabled = True
        DoEvents
        Chat(q).SendData "ADMINTXT" & txtMOTD.Text & vbCrLf
        Chat(q).SendData "ISAACNUM" & q & vbCrLf
        Exit Sub
    End If
    Next 'i
End If 'if index = 0
End Sub

Private Sub Chat_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo err
    Dim nfile As String
    Dim xfile As String
    nfile = App.Path & "\user.ini"
    xfile = App.Path & "\data.ini"
Dim strtime2 As String
strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
strtime2 = Format(Now, "dd-mmmm")
'strTime = CStr(curTime)
Dim strdatao As String
Chat(Index).GetData strdatao
arrdata = Split(strdatao, vbCrLf, -1, vbTextCompare)
strdata = arrdata

DontAdd = False

For i = 0 To UBound(arrdata)

    
    
    'CRAP FOR THE CHAT SERVER
    
    If Left$(strdata(i), 7) = "CHATTXT" Then
    Dim strChatTxt As String
    strChatTxt = Mid(strdata(i), 8, Len(strdata(i)))
    Call WriteIni("CHAT", strTime, strChatTxt, App.Path & "\" & strtime2 & ".ini")
    Call SendChat(strChatTxt)
    End If
    
    If Left$(strdata(i), 8) = "CHATNAME" Then
        Users(Index).Name = Mid(strdata(i), 9, Len(strdata(i)))
        Users(Index).IP = Server(Index).RemoteHostIP
        Dim strCurUser As String
        strCurUser = FindUser(Users(Index).Name)
        Call WriteIni(strCurUser, "IP", Users(Index).IP, App.Path & "\users.ini")
    End If
    
    
    If Left$(strdata(i), 11) = "CHATMSGNAME" Then
    CurMsgName = Mid(strdata(i), 12, Len(strdata(i)))
    End If
    
    If Left$(strdata(i), 11) = "CHATMSGTEXT" Then
    Dim CurMsgText As String
    CurMsgText = Users(Index).Name & ": " & Mid(strdata(i), 12, Len(strdata(i)))
    Call ChatMsg(CurMsgText, CurMsgName)
    
    End If
    If Left$(strdata(i), 8) = "ADMINTXT" Then
        Dim strAdminChatTxt As String
        strChatTxt = Mid(strdata(i), 9, Len(strdata(i)))
        Call WriteIni("ADMINCHAT", strTime, strChatTxt, App.Path & "\" & strtime2 & ".ini")
        Call AdminChat(strChatTxt)
    End If
    
    If Left$(strdata(i), 9) = "GETRATING" Then
        Dim strCheckRating As String
        strCheckRating = Mid(strdata(i), 10, Len(strdata(i)))
        For q = 1 To 20
            If Users(q).Name = strCheckRating Then
                Chat(Index).SendData "CHATRATING" & Users(q).Rating & vbCrLf
            End If
        Next 'q
    End If

    If Left$(strdata(i), 5) = "GETIP" Then
        Dim strCheckIP As String
        strCheckIP = Mid(strdata(i), 6, Len(strdata(i)))
        For q = 1 To 20
            If Users(q).Name = strCheckIP Then
                Chat(Index).SendData "CHATIP" & Users(q).IP & vbCrLf
            End If
        Next 'q
    End If
    
    If Left$(strdata(i), 7) = "IAMAWAY" Then
        Users(Index).Away = True
    End If
    If Left$(strdata(i), 7) = "NOTAWAY" Then
        Users(Index).Away = False
    End If
    
    If Left$(strdata(i), 8) = "ISAACPIC" Then
        Users(Index).Pic = Mid(strdata(i), 9, Len(strdata(i)))
        Users(Index).Left = 0
        Users(Index).Top = 0
        Call AddMUser(Index, True)
    End If
    If Left$(strdata(i), 6) = "SCREEN" Then
        Users(Index).Screen = Mid(strdata(i), 7, Len(strdata(i)))
        DontAdd = True
    End If
    If Left$(strdata(i), 6) = "ISAACX" Then
        Users(Index).Left = Mid(strdata(i), 7, Len(strdata(i)))
        DontAdd = True
    End If
    If Left$(strdata(i), 6) = "ISAACY" Then
        Users(Index).Top = Mid(strdata(i), 7, Len(strdata(i)))
        DontAdd = True
    End If
    
    If Left$(strdata(i), 10) = "CREATEGAME" Then
        Game(Index).Name = Mid(strdata(i), 11, Len(strdata(i)))
        Game(Index).Enabled = True
        Game(Index).IP = Chat(Index).RemoteHostIP
    End If
    If Left$(strdata(i), 8) = "GAMEHOST" Then
        Game(Index).Host = Mid(strdata(i), 9, Len(strdata(i)))
    End If
    If Left$(strdata(i), 11) = "GETGAMELIST" Then
        Call GetGameList(Index)
    End If
    If Left$(strdata(i), 9) = "CLOSEGAME" Then
        Game(Index).Enabled = False
    End If
    If Left$(strdata(i), 8) = "JOINGAME" Then
        strgame = Mid(strdata(i), 9, Len(strdata(i)))
        Call JoinGame(strgame, Index)
    End If
    
    If Left$(strdata(i), 10) = "CHANGECHAR" Then
        
        Dim strnewchar As String
        strnewchar = Mid(strdata(i), 11, Len(strdata(i)))
        q = FindUser(Chat(Index).Name)
        Call WriteIni(CStr(q), "CHAR", strnewchar, nfile)
    
    End If
    
    If Left$(strdata(i), 9) = "KILLCOINS" Then
        
        Dim strNewCoins As String
        strNewCoins = Mid(strdata(i), 10, Len(strdata(i)))
        q = FindUser(Chat(Index).Name)
        Call WriteIni(CStr(q), "COINS", strNewCoins, nfile)
    
    End If
    
    If Left$(strdata(i), 10) = "INGAMECHAT" Then
        Dim strInChat As String
        strInChat = Mid(strdata(i), 11, Len(strdata(i)))
        Call InGameChat(Index, strInChat)
    End If
    
    If Left$(strdata(i), 7) = "CHATBAN" Then
        Dim iBanMax As Integer
        Dim bSave As String
        Dim strBan As String
        strBan = Mid(strdata(i), 8, Len(strdata(i)))
        For q = 1 To 20
            If Users(q).Enabled = True And Users(q).Name = strBan Then
                bSave = App.Path & "\ban.ini"
                iBanMax = GetFromIni("GEN", "MAX", bSave)
                Chat(q).SendData "CHATBAN" & vbCrLf
                Call WriteIni("GEN", CStr(iBanMax + 1), Chat(q).RemoteHostIP, bSave)
            End If
        Next 'q
    End If
    If Left$(strdata(i), 9) = "ADMINKILL" Then
        Dim strKill As String
        strKill = Mid(strdata(i), 10, Len(strdata(i)))
        For q = 1 To 20
            If Users(q).Enabled = True And Users(q).Name = strKill Then
                Chat(q).SendData "KILL" & vbCrLf
            End If
        Next 'q
    End If
    If Left$(strdata(i), 10) = "CHATREPORT" Then
        strreport = Mid(strdata(i), 11, Len(strdata(i)))
        Call WriteIni("GEN", strTime, "Narc on user " & strreport & " by user " & Users(Index).Name, App.Path & "\reports.ini")
    End If
Next 'i

    If strdatao <> "CHATSPAM" & vbCrLf And DontAdd = False Then
        lstStatus.AddItem strTime & " " & strdatao
    End If
    
Exit Sub
err:
Exit Sub


End Sub

Private Sub closeme_Timer()
On Error GoTo err
For i = 1 To 20
    If Server(i).State <> sckClosed Then
        Server(i).SendData "P" & vbCrLf
    End If
Next 'i
Exit Sub
err:
    Server(i).Close
    Resume Next
End Sub

Private Sub cmdAdd_Click(Index As Integer)
Dim nSave As String
Dim itotal As Integer
Dim stotal As String
If Index = 0 Then
nSave = App.Path & "\psynergy.ini"
    If txtPsy(1).Text = "W" Then
        stotal = GetFromIni("GEN", "W", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "W", stotal, nSave)
        stotal = "W" & CStr(itotal)
    End If
    If txtPsy(1).Text = "F" Then
        stotal = GetFromIni("GEN", "F", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "F", stotal, nSave)
        stotal = "F" & CStr(itotal)
    End If
    If txtPsy(1).Text = "N" Then
        stotal = GetFromIni("GEN", "N", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "N", stotal, nSave)
        stotal = "N" & CStr(itotal)
    End If
    If txtPsy(1).Text = "E" Then
        stotal = GetFromIni("GEN", "E", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "E", stotal, nSave)
        stotal = "E" & CStr(itotal)
    End If
    If txtPsy(1).Text = "D" Then
        stotal = GetFromIni("GEN", "D", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "D", stotal, nSave)
        stotal = "D" & CStr(itotal)
    End If
    If txtPsy(1).Text = "H" Then
        stotal = GetFromIni("GEN", "H", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "H", stotal, nSave)
        stotal = "H" & CStr(itotal)
    End If
    Call WriteIni(stotal, "NAME", txtPsy(0).Text, nSave)
    Call WriteIni(stotal, "TYPE", txtPsy(2).Text, nSave)
    Call WriteIni(stotal, "DAMAGE", txtPsy(3).Text, nSave)
    Call WriteIni(stotal, "PP", txtPsy(4).Text, nSave)
    Call WriteIni(stotal, "RATING", txtPsy(5).Text, nSave)
    Call WriteIni(stotal, "DJINN", txtPsy(6).Text, nSave)
    Call WriteIni(stotal, "DESC", txtPsy(7).Text, nSave)
End If
If Index = 1 Then
nSave = App.Path & "\items.ini"
    stotal = GetFromIni("GEN", "TOTAL", nSave)
    itotal = CInt(stotal)
    itotal = itotal + 1
    stotal = CStr(itotal)
    Call WriteIni("GEN", "TOTAL", stotal, nSave)
    stotal = CStr(itotal)
    Call WriteIni("I" & stotal, "NAME", txtItem(0).Text, nSave)
    Call WriteIni("I" & stotal, "COINS", txtItem(1).Text, nSave)
    Call WriteIni("I" & stotal, "DESCRIPTION", txtItem(2).Text, nSave)
    Call WriteIni("I" & stotal, "SPCNAME", txtItem(3).Text, nSave)
    Call WriteIni("I" & stotal, "SPCTYPE", txtItem(4).Text, nSave)
    Call WriteIni("I" & stotal, "SPCDAMAGE", txtItem(5).Text, nSave)
    Call WriteIni("I" & stotal, "TYPE", txtItem(6).Text, nSave)
    Call WriteIni("I" & stotal, "DAMAGE", txtItem(7).Text, nSave)
End If
If Index = 2 Then
nSave = App.Path & "\djinn.ini"
    If txtDjinn(1).Text = "W" Then
        stotal = GetFromIni("GEN", "W", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "W", stotal, nSave)
        stotal = "W" & CStr(itotal)
    End If
    If txtDjinn(1).Text = "F" Then
        stotal = GetFromIni("GEN", "F", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "F", stotal, nSave)
        stotal = "F" & CStr(itotal)
    End If
    If txtDjinn(1).Text = "N" Then
        stotal = GetFromIni("GEN", "N", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "N", stotal, nSave)
        stotal = "N" & CStr(itotal)
    End If
    If txtDjinn(1).Text = "E" Then
        stotal = GetFromIni("GEN", "E", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "E", stotal, nSave)
        stotal = "E" & CStr(itotal)
    End If
    If txtDjinn(1).Text = "D" Then
        stotal = GetFromIni("GEN", "D", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "D", stotal, nSave)
        stotal = "D" & CStr(itotal)
    End If
    If txtDjinn(1).Text = "H" Then
        stotal = GetFromIni("GEN", "H", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "H", stotal, nSave)
        stotal = "H" & CStr(itotal)
    End If
    Call WriteIni(stotal, "NAME", txtDjinn(0).Text, nSave)
    Call WriteIni(stotal, "DESCRIPTION", txtDjinn(2).Text, nSave)
    Call WriteIni(stotal, "TYPE", txtDjinn(3).Text, nSave)
    Call WriteIni(stotal, "DAMAGE", txtDjinn(4).Text, nSave)
End If
If Index = 3 Then
    nSave = App.Path & "\summons.ini"
    If txtSummon(2).Text = "W" Then
        stotal = GetFromIni("GEN", "W", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "W", stotal, nSave)
        stotal = "W" & CStr(itotal)
    End If
    If txtSummon(2).Text = "F" Then
        stotal = GetFromIni("GEN", "F", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "F", stotal, nSave)
        stotal = "F" & CStr(itotal)
    End If
    If txtSummon(2).Text = "N" Then
        stotal = GetFromIni("GEN", "N", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "N", stotal, nSave)
        stotal = "N" & CStr(itotal)
    End If
    If txtSummon(2).Text = "E" Then
        stotal = GetFromIni("GEN", "E", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "E", stotal, nSave)
        stotal = "E" & CStr(itotal)
    End If
    If txtSummon(2).Text = "D" Then
        stotal = GetFromIni("GEN", "D", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "D", stotal, nSave)
        stotal = "D" & CStr(itotal)
    End If
    If txtSummon(2).Text = "H" Then
        stotal = GetFromIni("GEN", "H", nSave)
        itotal = CInt(stotal)
        itotal = itotal + 1
        stotal = CStr(itotal)
        Call WriteIni("GEN", "H", stotal, nSave)
        stotal = "H" & CStr(itotal)
    End If
    Call WriteIni(stotal, "NAME", txtSummon(0).Text, nSave)
    Call WriteIni(stotal, "DESC", txtSummon(1).Text, nSave)
    Call WriteIni(stotal, "DJINN", txtSummon(3).Text, nSave)
End If
    
End Sub

Private Sub cmdCloseServer_Click()
If cmdCloseServer.Caption <> "Re-open Server" Then
    cmdCloseServer.Caption = "Re-open Server"
    strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
        txtHTML.Text = ""
        txtHTML.Text = "<b>The server has been temporarily shut down for paitent.<br><br>Current Work on the Game: <b>" & txtMOTD.Text & "</b>"
 FNO = FreeFile
     On Error Resume Next
     err.Clear
     Open ("C:\sambar50\docs\status.html") For Output As #FNO
      If err.Number <> 0 Then
        MsgBox "an error has occured"
       Else
        Print #FNO, (txtHTML.Text)
      End If
     Close #FNO
     On Error GoTo 0
Call WriteIni("GEN", "DOWN", "TRUE", App.Path & "\motd.ini")
Else
    cmdCloseServer.Caption = "Go Down For Maitence"
Call WriteIni("GEN", "DOWN", "FALSE", App.Path & "\motd.ini")
End If
End Sub

Private Sub cmdGame_Click()
frmEditor.Show
End Sub

Private Sub cmdKill_Click()
On Error Resume Next
For q = 1 To 20
If UserName(q) = txtKill.Text Then
Chat(q).SendData "KILL" & vbCrLf
End If
Next 'q

End Sub

Private Sub cmdMOTD_Click()
Call WriteIni("MOTD", "MOTD", txtMOTD.Text, App.Path & "\motd.ini")
End Sub

Private Sub cmdMsg_Click()
On Error Resume Next
For q = 1 To 20
If UserName(q) = txtKill.Text Then
Chat(q).SendData "CHATTXT" & "[PRIVATE MESSAGE FROM SERVER]: " & txtContent.Text & vbCrLf
DoEvents
End If
Next 'q
End Sub

Private Sub cmdReset_Click()
Dim strMax As String
Dim iMax As Integer
Dim nSave As String
nSave = App.Path & "\user.ini"
strMax = GetFromIni("GEN", "TOTAL", nSave)
iMax = CInt(strMax)
For i = 0 To iMax
    Call WriteIni(CStr(i), "RATING", "1000", nSave)
    Call WriteIni(CStr(i), "LEVEL", "1", nSave)
    Call WriteIni(CStr(i), "WINS", "0", nSave)
    Call WriteIni(CStr(i), "LOSS", "0", nSave)
    Call WriteIni(CStr(i), "DISC", "0", nSave)
    Call AdminChat("[Server Message:] Ladder has been reset.")
Next 'i

End Sub

Private Sub cmdSend_Click()
On Error Resume Next
For q = 1 To 20
'If Server(q).State = sckConnected Then
Chat(q).SendData "ADMINTXT" & "[SERVER MESSAGE]: " & txtMsg.Text & vbCrLf
DoEvents
'End If
Next 'q
End Sub

Private Sub cmdTOS_Click()
Dim iTos As Integer
Dim nSave As String

nSave = App.Path & "\motd.ini"

iTos = CInt(GetFromIni("MOTD", "TOSMAX", nSave))
For i = 1 To iTos
    strTOS(i) = GetFromIni("MOTD", "TOS" & i, nSave)
Next 'i
End Sub

Private Sub Form_DblClick()
If frmServer(Index).Height > 1200 Then
frmServer(Index).Height = 1200
frmServer(Index).Width = 1200
Else
frmServer(Index).Height = 7830
frmServer(Index).Width = 7830
End If
End Sub

Private Sub Form_Load()
Server(0).Listen
Chat(0).Listen

Call cmdTOS_Click

ladderRefresh = 0

For i = 1 To 20
Load Server(i)
Load Chat(i)
Users(i).Enabled = False
Users(i).Away = False
Game(i).Enabled = False
Next 'i

txtMOTD.Text = GetFromIni("MOTD", "MOTD", App.Path & "\motd.ini")
txtVer.Text = GetFromIni("MOTD", "VER", App.Path & "\motd.ini")
ServerVersion = txtVer.Text
strdown = GetFromIni("GEN", "DOWN", App.Path & "\motd.ini")
If strdown = "FALSE" Then
    cmdCloseServer.Caption = "Close Server For Maitence"
Else
    cmdCloseServer.Caption = "Re-Open Server"
End If

upTime = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
frmEditor.Show
End If
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Server_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
If Index = 0 Then
    'Beep
    iNewServer = 100
    For q = 1 To 20
    If Server(q).State = sckClosed And iNewServer = 100 Then
        Server(q).Accept requestID
        For i = 1 To 15
            If strTOS(i) <> "" Then
                Server(q).SendData "TOS" & strTOS(i) & vbCrLf
            End If
        Next 'i
        iNewServer = MaxCon
        Exit Sub
    End If
    Next 'i
    If iNewServer = 100 And MaxCon <> 100 Then
        Server(21).Accept requestID
        Server(21).SendData "FULL" & vbCrLf
        DoEvents
        Server(21).Close
    End If
End If 'if index = 0

End Sub

Private Sub Server_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo err
    Dim nfile As String
    Dim xfile As String
    nfile = App.Path & "\user.ini"
    xfile = App.Path & "\data.ini"
Dim strtime2 As String
strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
strtime2 = Format(Now, "dd-mmmm")
'strTime = CStr(curTime)
Dim strdatao As String
Server(Index).GetData strdatao
arrdata = Split(strdatao, vbCrLf, -1, vbTextCompare)
strdata = arrdata

For i = 0 To UBound(arrdata)
'If i >= UBound(arrdata) Then Exit Sub
If Left$(strdata(i), 4) = "USER" Then
    UserVersion = 0
    Debug.Print "USER"
    curUser = Mid(strdata(i), 5, Len(strdata(i)))
    'MsgBox curUser
    noclose = True
    'Exit Sub
End If
If Left$(strdata(i), 4) = "VERS" Then
    Dim strVersion As String
    strVersion = Mid(strdata(i), 5, Len(strdata(i)))
    UserVersion = CVar(strVersion)
End If


If Left$(strdata(i), 4) = "PASS" Then
    Debug.Print "PASS"
    
    
    Dim curMax As Integer
    curMax = CInt(GetFromIni("GEN", "TOTAL", nfile))
    
    Dim curuser2 As String
    curuser2 = curUser
    
    For q = 0 To curMax
        Dim curName As String
        curName = GetFromIni(CStr(q), "NAME", nfile)
        If curName = curUser Then
            curUser = q
        End If
    Next 'q
    
    
    curPassword = Mid(strdata(i), 5, Len(strdata(i)))
    realPassword = GetFromIni(curUser, "PASSWORD", nfile)
    strRating = GetFromIni(curUser, "RATING", nfile)
    strChar = GetFromIni(curUser, "CHAR", nfile)
    strLvl = GetFromIni(curUser, "LEVEL", nfile)
    strDjinn = GetFromIni(curUser, "DJINNNUM", nfile)
    strCoins = GetFromIni(curUser, "COINS", nfile)
    strWins = GetFromIni(curUser, "WINS", nfile)
    strDisc = GetFromIni(curUser, "DISC", nfile)
    strLoss = GetFromIni(curUser, "LOSS", nfile)
    strType = GetFromIni(curUser, "TYPE", nfile)
    strMyWeapon = GetFromIni(curUser, "ITEM", nfile)
    
    Users(Index).Rating = strRating
    Users(Index).Wins = strWins
    Users(Index).Losses = strLoss
    Users(Index).Disconnects = strDisc
    
    curUser = curuser2
    
    'MsgBox curPassword
    If curPassword = realPassword And curPassword <> "" And UserVersion = ServerVersion Then
        Call WriteIni(strTime, "Data", "Good Password Attempt on " & curUser & " at " & Server(Index).RemoteHostIP, xfile)
        Server(Index).SendData "GOOD" & vbCrLf
        Dim nSave As String
        Dim stotal As String
        Dim itotal As Integer
        nSave = App.Path & "\items.ini"
        stotal = GetFromIni("GEN", "TOTAL", nSave)
        itotal = CInt(stotal)
    For q = 1 To itotal
        Server(Index).SendData "CURITEM" & q & vbCrLf
        stotal = GetFromIni("I" & q, "NAME", nSave)
        Server(Index).SendData "ITEMNAME" & stotal & vbCrLf
        stotal = GetFromIni("I" & q, "DESCRIPTION", nSave)
        Server(Index).SendData "ITEMDESC" & stotal & vbCrLf
        stotal = GetFromIni("I" & q, "DAMAGE", nSave)
        Server(Index).SendData "ITEMDMG" & stotal & vbCrLf
        stotal = GetFromIni("I" & q, "SPCDAMAGE", nSave)
        Server(Index).SendData "ITEMSPCDMG" & stotal & vbCrLf
        stotal = GetFromIni("I" & q, "SPCTYPE", nSave)
        Server(Index).SendData "ITEMSPCTYPE" & stotal & vbCrLf
        stotal = GetFromIni("I" & q, "SPCNAME", nSave)
        Server(Index).SendData "ITEMSPCNAME" & stotal & vbCrLf
        stotal = GetFromIni("I" & q, "TYPE", nSave)
        Server(Index).SendData "ITEMTYPE" & stotal & vbCrLf
        stotal = GetFromIni("I" & q, "COINS", nSave)
        Server(Index).SendData "ITEMCOINS" & stotal & vbCrLf
        
    Next 'i
    nSave = App.Path & "\djinn.ini"
    Dim curType As String
    If strType = "E" Then
        stotal = GetFromIni("GEN", "E", nSave)
        itotal = CInt(stotal)
        curType = "E"
    End If
    If strType = "F" Then
        stotal = GetFromIni("GEN", "F", nSave)
        itotal = CInt(stotal)
        curType = "F"
    End If
    If strType = "D" Then
        stotal = GetFromIni("GEN", "D", nSave)
        itotal = CInt(stotal)
        curType = "D"
    End If
    If strType = "W" Then
        stotal = GetFromIni("GEN", "W", nSave)
        itotal = CInt(stotal)
        curType = "W"
    End If
    If strType = "H" Then
        stotal = GetFromIni("GEN", "H", nSave)
        itotal = CInt(stotal)
        curType = "H"
    End If
    If strType = "D" Then
        stotal = GetFromIni("GEN", "D", nSave)
        itotal = CInt(stotal)
        curType = "D"
    End If
        
    For w = 1 To itotal
        Server(Index).SendData "CURDJINN" & w & vbCrLf
        stotal = GetFromIni(curType & w, "NAME", nSave)
        Server(Index).SendData "DJINNNAME" & stotal & vbCrLf
        stotal = GetFromIni(curType & w, "DESCRIPTION", nSave)
        Server(Index).SendData "DJINNDESC" & stotal & vbCrLf
        stotal = GetFromIni(curType & w, "TYPE", nSave)
        Server(Index).SendData "DJINNTYPE" & stotal & vbCrLf
        stotal = GetFromIni(curType & w, "DAMAGE", nSave)
        Server(Index).SendData "DJINNDMG" & stotal & vbCrLf
    Next 'i
    
    nSave = App.Path & "\psynergy.ini"
    
    stotal = GetFromIni("GEN", curType, nSave)
    itotal = CInt(stotal)
    
    Users(Index).Rating = CInt(strRating)
    iCurRank = CInt(strRating)
    
    If strType = "E" Then
        stotal = GetFromIni("GEN", "E", nSave)
        itotal = CInt(stotal)
        curType = "E"
    End If
    If strType = "F" Then
        stotal = GetFromIni("GEN", "F", nSave)
        itotal = CInt(stotal)
        curType = "F"
    End If
    If strType = "D" Then
        stotal = GetFromIni("GEN", "D", nSave)
        itotal = CInt(stotal)
        curType = "D"
    End If
    If strType = "W" Then
        stotal = GetFromIni("GEN", "W", nSave)
        itotal = CInt(stotal)
        curType = "W"
    End If
    If strType = "H" Then
        stotal = GetFromIni("GEN", "H", nSave)
        itotal = CInt(stotal)
        curType = "H"
    End If
    If strType = "D" Then
        stotal = GetFromIni("GEN", "D", nSave)
        itotal = CInt(stotal)
        curType = "D"
    End If
    
    For e = 1 To itotal
        strNewRank = GetFromIni(curType & e, "RATING", nSave)
        iNewRank = CInt(strNewRank)
        If iCurRank >= iNewRank Then
            Server(Index).SendData "CURPSY" & e & vbCrLf
            stotal = GetFromIni(curType & e, "NAME", nSave)
            Server(Index).SendData "PSYNAME" & stotal & vbCrLf
            stotal = GetFromIni(curType & e, "DAMAGE", nSave)
            Server(Index).SendData "PSYDMG" & stotal & vbCrLf
            stotal = GetFromIni(curType & e, "DESC", nSave)
            Server(Index).SendData "PSYDESC" & stotal & vbCrLf
            stotal = GetFromIni(curType & e, "TYPE", nSave)
            Server(Index).SendData "PSYTYPE" & stotal & vbCrLf
            stotal = GetFromIni(curType & e, "PP", nSave)
            Server(Index).SendData "PSYPP" & stotal & vbCrLf
            stotal = GetFromIni(curType & e, "DJINN", nSave)
            Server(Index).SendData "PSYDJINN" & stotal & vbCrLf

        End If
    Next 'e
    
    nSave = App.Path & "\summons.ini"
    
    stotal = GetFromIni("GEN", curType, nSave)
    itotal = CInt(stotal)
    
    If strType = "E" Then
        stotal = GetFromIni("GEN", "E", nSave)
        itotal = CInt(stotal)
        curType = "E"
    End If
    If strType = "F" Then
        stotal = GetFromIni("GEN", "F", nSave)
        itotal = CInt(stotal)
        curType = "F"
    End If
    If strType = "D" Then
        stotal = GetFromIni("GEN", "D", nSave)
        itotal = CInt(stotal)
        curType = "D"
    End If
    If strType = "W" Then
        stotal = GetFromIni("GEN", "W", nSave)
        itotal = CInt(stotal)
        curType = "W"
    End If
    If strType = "H" Then
        stotal = GetFromIni("GEN", "H", nSave)
        itotal = CInt(stotal)
        curType = "H"
    End If
    If strType = "D" Then
        stotal = GetFromIni("GEN", "D", nSave)
        itotal = CInt(stotal)
        curType = "D"
    End If

    
    For r = 1 To itotal
        Server(Index).SendData "CURSUM" & r & vbCrLf
        stotal = GetFromIni(curType & r, "NAME", nSave)
        Server(Index).SendData "SUMNAME" & stotal & vbCrLf
        stotal = GetFromIni(curType & r, "DJINN", nSave)
        Server(Index).SendData "SUMDJINN" & stotal & vbCrLf
        stotal = GetFromIni(curType & r, "DESC", nSave)
        Server(Index).SendData "SUMDESC" & stotal & vbCrLf
    Next 'r
    
    Server(Index).SendData "RATING" & strRating & vbCrLf
    Server(Index).SendData "WPN" & strMyWeapon & vbCrLf
    Server(Index).SendData "CHAR" & strChar & vbCrLf
    Server(Index).SendData "DJINN" & strDjinn & vbCrLf
    Server(Index).SendData "COINS" & strCoins & vbCrLf
    Server(Index).SendData "WINS" & strWins & vbCrLf
    Server(Index).SendData "LOSS" & strLoss & vbCrLf
    Server(Index).SendData "DISC" & strDisc & vbCrLf
    Server(Index).SendData "TYPE" & strType & vbCrLf
    Server(Index).SendData "LVL" & strLvl & vbCrLf


    Else
    If ServerVersion = UserVersion Then
        Call WriteIni(strTime, "Data", "Bad Password Attempt on " & curUser & " at " & Server(Index).RemoteHostIP, xfile)
    Else
        Call WriteIni(strTime, "Data", "Bad version from " & curUser & " at " & Server(Index).RemoteHostIP, xfile)
    End If
        Server(Index).SendData "BAD" & vbCrLf
        DoEvents
'        Server(Index).Listen
        End If
        noclose = True
End If
If Left$(strdata(i), 7) = "NEWUSER" Then
    NewUser = Mid(strdata(i), 8, Len(strdata(i)))
    noclose = True
End If
If Left$(strdata(i), 4) = "CHAR" Then
    strChar = Mid(strdata(i), 5, Len(strdata(i)))
    noclose = True
End If
If Left$(strdata(i), 5) = "NEWPW" Then
    
    NewPass = Mid(strdata(i), 6, Len(strdata(i)))
    
    Dim getMax As Integer
    isitnew = ""
    getMax = CInt(GetFromIni("GEN", "TOTAL", nfile))
    For q = 0 To getMax
        strusercheck = GetFromIni(CStr(q), "NAME", nfile)
        If NewUser = strusercheck Then isitnew = "NOT"
    Next 'q
    
    If isitnew = "" Then
    stotal = GetFromIni("GEN", "TOTAL", nfile)
    itotal = CInt(stotal)
    itotal = itotal + 1
    stotal = CStr(itotal)
    Call WriteIni("GEN", "TOTAL", stotal, nfile)
    Call WriteIni("GEN", "NEWEST", NewUser, nfile)
    Call WriteIni(stotal, "PASSWORD", NewPass, nfile)
    Call WriteIni(stotal, "LEVEL", "1", nfile)
    Call WriteIni(stotal, "DJINNNUM", "1", nfile)
    Call WriteIni(stotal, "COINS", "100", nfile)
    Call WriteIni(stotal, "RATING", "1000", nfile)
    Call WriteIni(stotal, "CHAR", strChar, nfile)
    Call WriteIni(stotal, "WINS", "0", nfile)
    Call WriteIni(stotal, "DISC", "0", nfile)
    Call WriteIni(stotal, "LOSS", "0", nfile)
    Call WriteIni(stotal, "ITEM", "1", nfile)
    Call WriteIni(stotal, "DPOINTS", "0", nfile)
    Call WriteIni(stotal, "NAME", NewUser, nfile)
        
    If strChar = "Isaac" Or strChar = "Guard" Or strChar = "Gladiator" Then
    Call WriteIni(stotal, "TYPE", "E", nfile)
    End If
    If strChar = "Jenna" Or strChar = "Garret" Or strChar = "Saturos" Or strChar = "Menardi" Then
    Call WriteIni(stotal, "TYPE", "F", nfile)
    End If
    If strChar = "Ivan" Or strChar = "Sheba" Then
    Call WriteIni(stotal, "TYPE", "N", nfile)
    End If
    If strChar = "Mia" Or strChar = "Alex" Or strChar = "Caption Contest Winner" Then
    Call WriteIni(stotal, "TYPE", "W", nfile)
    End If
    If strChar = "Felix" Then
    Call WriteIni(stotal, "TYPE", "H", nfile)
    End If
    If strChar = "Kraden" Then
    Call WriteIni(stotal, "TYPE", "D", nfile)
    End If
    
    
    
    Server(Index).SendData "GUSR" & vbCrLf
'    MsgBox NewUser & " ****"
    'MsgBox NewPass
    DoEvents
    
    noclose = True
    closeme.Enabled = True
    Call WriteIni(strTime, "Data", "Created new user " & NewUser & " at " & Server(Index).RemoteHostIP, xfile)
    Else
    noclose = True
    Call WriteIni(strTime, "Data", "Bad password attempt on new user " & NewUser & " at " & Server(Index).RemoteHostIP, xfile)
    Server(Index).SendData "BUSR" & vbCrLf
    Debug.Print NewPass
    Debug.Print NewUser
'    Timeout (2)

'    Timeout (2)
'    Server(Index).Listen
    End If
    'closeme.Enabled = True
End If
If Left$(strdata(i), 9) = "LOSTEMAIL" Then
    noclose = True
    lstEmail = Mid(strdata(i), 10, Len(strdata(i)))
End If
If Left$(strdata(i), 8) = "LOSTUSER" Then
    lstUser = Mid(strdata(i), 9, Len(strdata(i)))
    Call WriteIni(strTime, "Data", "Lost e-mail request on " & lstUser & " at " & Server(Index).RemoteHostIP & " for e-mail address " & lstEmail, xfile)

    noclose = True
    closeme.Enabled = True
End If
If Left$(strdata(i), 7) = "WINUSER" Then
noclose = True
nfile = App.Path & "\user.ini"
'Dim curUser As String
curUser = Mid(strdata(i), 8, Len(strdata(i)))

curMax = CInt(GetFromIni("GEN", "TOTAL", nfile))
For q = 0 To curMax
    chkuser = GetFromIni(CStr(q), "NAME", nfile)
    If chkuser = curUser Then
    curUser = CStr(q)
    End If
Next 'i

strCurRating = GetFromIni(curUser, "RATING", nfile)
iCurRating = CInt(strCurRating)
strCurLvl = GetFromIni(curUser, "LEVEL", nfile)
iCurLvl = CInt(strCurLvl)
strCurCoins = GetFromIni(curUser, "COINS", nfile)
iCurCoins = CInt(strCurCoins)
noclose = True
End If
    If Left$(strdata(i), 6) = "RATING" Then
        Dim snewRating As String
        Dim inewRating As Integer
        snewRating = Mid(strdata(i), 7, Len(strdata(i)))
        inewRating = CInt(snewRating)
        snewRating = CStr(inewRating + iCurRating)
        Call WriteIni(curUser, "RATING", snewRating, nfile)
        noclose = True
    End If
    If Left$(strdata(i), 5) = "COINS" Then
        Dim snewCoins As String
        Dim inewCoins As Integer
        snewCoins = Mid(strdata(i), 6, Len(strdata(i)))
        inewCoins = CInt(snewCoins)
        snewCoins = CStr(inewCoins + iCurCoins)
        Call WriteIni(curUser, "COINS", snewCoins, nfile)
        noclose = True
    End If
    If Left$(strdata(i), 3) = "LVL" Then
        Dim snewLVL As String
        Dim inewLVL As Integer
        snewLVL = Mid(strdata(i), 4, Len(strdata(i)))
        inewLVL = CInt(snewLVL)
        snewLVL = CStr(inewLVL + iCurLvl)
        Server(Index).SendData "STAT"
        Call WriteIni(curUser, "LEVEL", snewLVL, nfile)
    End If
    If Left$(strdata(i), 4) = "SWIN " Then
        Call WriteIni(strTime, "Data", curUser & " won a game!", App.Path & "\data.ini")
        Dim CurWins As String
        Dim iWins As Integer
        CurWins = GetFromIni(curUser, "WINS", nfile)
        iWins = CInt(CurWins)
        CurWins = CStr(iWins + 1)
        Call WriteIni(curUser, "WINS", CurWins, nfile)
        noclose = True
    End If
    If Left$(strdata(i), 4) = "LOSE" Then
        Call WriteIni(strTime, "Data", curUser & " lost a game!", App.Path & "\data.ini")
        Dim CurLose As String
        Dim iLose As Integer
        CurLose = GetFromIni(curUser, "LOSS", nfile)
        iLose = CInt(CurLose)
        CurLose = CStr(iLose + 1)
        Call WriteIni(curUser, "LOSS", CurLose, nfile)
        noclose = True
    End If
    
If Left$(strdata(i), 10) = "SINGLENAME" Then
    Dim snewName As String
    snewName = Mid(strdata(i), 11, Len(strdata(i)))
    SingleUser = FindUser(snewName)
End If
If Left$(strdata(i), 11) = "SINGLECOINS" Then
    Dim ssingCoins As String
    ssingCoins = Mid(strdata(i), 12, Len(strdata(i)))
    Call WriteIni(CStr(SingleUser), "COINS", ssingCoins, App.Path & "\user.ini")
    Server(Index).SendData "SINGLECOINS" & vbCrLf
End If
    


If Left$(strdata(i), 11) = "NEWITEMUSER" Then
Dim strUserGuy As String
strUserGuy = Mid(strdata(i), 12, Len(strdata(i)))

curMax = CInt(GetFromIni("GEN", "TOTAL", nfile))
For q = 0 To curMax
    chkuser = GetFromIni(CStr(q), "NAME", nfile)
    If chkuser = curUser Then
    strUserGuy = CStr(q)
    End If
Next 'i

End If
If Left$(strdata(i), 12) = "NEWITEMCOINS" Then
Dim strNewCoins As String
strNewCoins = Mid(strdata(i), 13, Len(strdata(i)))
Call WriteIni(strUserGuy, "COINS", strNewCoins, nfile)
End If
If Left$(strdata(i), 11) = "NEWITEMNAME" Then
    Dim strNewItemName As String
    strNewItemName = Mid(strdata(i), 12, Len(strdata(i)))
    Dim strItemTotal As String
    Dim iItemtotal As Integer
    strItemTotal = GetFromIni("GEN", "TOTAL", App.Path & "\items.ini")
    iItemtotal = CInt(strItemTotal)
    Dim strCheckItemName As String
    For q = 1 To iItemtotal
        strCheckItemName = GetFromIni("I" & q, "NAME", App.Path & "\items.ini")
    If strNewItemName = strCheckItemName Then
        Call WriteIni(strUserGuy, "ITEM", CStr(q), nfile)
    End If
    Next 'q
    Server(Index).SendData "ITEMCONFIRM" & vbCrLf
End If
If Left$(strdata(i), 9) = "ERRORDESC" Then
    Dim eSave As String
    Dim strErrorDesc As String
    eSave = App.Path & "\errorlog.ini"
    strErrorDesc = Mid(strdata(i), 10, Len(strdata(i)))
    Call WriteIni(strTime, "ERROR", strErrorDesc, eSave)
End If
If Left$(strdata(i), 8) = "ERRORNUM" Then
    eSave = App.Path & "\errorlog.ini"
    strErrorDesc = Mid(strdata(i), 9, Len(strdata(i)))
    Call WriteIni(strTime, "NUM", strErrorDesc, eSave)
End If
If Left$(strdata(i), 11) = "ERRORSOURCE" Then
    eSave = App.Path & "\errorlog.ini"
    strErrorDesc = Mid(strdata(i), 12, Len(strdata(i)))
    Call WriteIni(strTime, "SOURCE", strErrorDesc, eSave)
End If



Next 'i
Debug.Print arrdata
    If strdatao <> "CHATSPAM" & vbCrLf Then
        lstStatus.AddItem strTime & " " & strdatao
    End If
    
Exit Sub
err:
Exit Sub

End Sub

Private Sub Server_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
Call WriteIni(strTime, "Errors", Description, App.Path & "\data.ini")
Server(Index).Close
'Server(Index).Listen
End Sub

Private Sub timeclose_Timer()
lblstatus.Caption = Server(Index).State
End Sub

Private Sub PopulateUser(ByVal q As Integer)
On Error Resume Next
    Server(iNewServer).SendData "CHATTXT" & "Message of the Day:" & txtMOTD.Text & vbCrLf
    
    
    For i = 1 To 20
    If UserName(i) <> "" And i <> q Then
    Server(q).SendData "CHATNUM" & i & vbCrLf
    Server(q).SendData "CHATNAME" & UserName(i) & vbCrLf
    Server(q).SendData "CHATRATING" & UserRating(i) & vbCrLf
    Server(q).SendData "CHATIP" & Server(i).RemoteHostIP & vbCrLf
    DoEvents
    End If
    
    If UserName(q) <> "" Then
    Server(i).SendData "CHATNUM" & iNewServer & vbCrLf
    Server(i).SendData "CHATNAME" & UserName(q) & vbCrLf
    Server(i).SendData "CHATRATING" & UserRating(q) & vbCrLf
    Server(i).SendData "CHATIP" & Server(q).RemoteHostIP & vbCrLf
    DoEvents
    End If
    Next 'i
End Sub
Sub DestroyUser(ByVal q As String)
UserName(q) = ""
UserRating(q) = ""
For i = 1 To 20
Server(i).SendData "CHATKILL" & q & vbCrLf
Next 'i
End Sub
Sub SendChat(ByVal strChatTxt As String)
On Error Resume Next

For q = 1 To 20
If Users(q).Enabled = True Then
Chat(q).SendData "CHATTXT" & strChatTxt & vbCrLf
DoEvents
End If
Next 'q

End Sub

Private Sub timeMulti_Timer()
For i = 1 To 20
    If Users(i).Enabled = True Then
        Call AddMUser(i, False)
    End If
Next 'i
End Sub

Private Sub timePing_Timer()
On Error GoTo err
For i = 1 To 20
    If Users(i).Enabled = True Then
        Chat(i).SendData "P" & vbCrLf
    End If
Next 'i
Call SendUsers

Exit Sub
err:
Users(i).Enabled = False
Chat(i).Close
Game(i).Enabled = False
Call KillChar(CStr(i))
Resume Next

End Sub

Private Sub timeServerPing_Timer()
On Error GoTo err
For i = 1 To 20
    If Server(i).State <> sckClosed Then
        Server(i).SendData "P" & vbCrLf
    End If
Next 'i
err:
If i <= 20 Then
Server(i).Close
Else
Resume Next
End If

End Sub

Private Sub timeUpdate_Timer()
ladderRefresh = ladderRefresh + 1
If ladderRefresh = 50 Then

upTime = upTime + 1 'Server uptime

ladderRefresh = 0
lstRank.Clear
txtHTML.Text = ""
    Dim userMax As Integer
    Dim nSave As String
    nSave = App.Path & "\user.ini"
    userMax = CInt(GetFromIni("GEN", "TOTAL", nSave))
    For i = 0 To userMax
        Dim curRating As String
        curRating = GetFromIni(CStr(i), "RATING", nSave)
        Dim curName As String
        curName = GetFromIni(CStr(i), "NAME", nSave)
        
        Dim alphanum(1 To 4) As String
        
        For q = 1 To 4
        alphanum(q) = Mid(curRating, q, 1)
        alphanum(q) = ConvertNum(alphanum(q))
        Next 'q
        
        curRating = alphanum(1) & alphanum(2) & alphanum(3) & alphanum(4)
        
        lstRank.AddItem curRating & " - " & curName
        
    Next 'i
    
    
    For i = 0 To lstRank.ListCount
        If lstRank.List(i) <> "" Then
            strrate = ConvertAlpha(lstRank.List(i))
            txtHTML.Text = txtHTML.Text & "<br>" & vbNewLine & "<b>" & i + 1 & "</b> - " & strrate & " - " & Mid(lstRank.List(i), 8, Len(lstRank.List(i)))
        End If
    Next 'i
    Dim strIntroMsg As String
    Dim strTime As String
    strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
    
    strIntroMsg = "Welcome to the Ladder Section of Golden Sun: The War of the Adepts.  Here you can find the ranking of players from first to worst :).<br><b>Rank - Rating - Name</b>"
    txtHTML.Text = strIntroMsg & vbNewLine & txtHTML.Text & "<br><b>Last Updated: " & strTime & "</b>"
    Dim FNO As Long
 FNO = FreeFile
 On Error Resume Next
 err.Clear
 Open ("C:\sambar50\docs\ladder.html") For Output As #FNO
  If err.Number <> 0 Then
    MsgBox "an error has occured"
   Else
    Print #FNO, (txtHTML.Text)
  End If
 Close #FNO
 On Error GoTo 0
End If

If ladderRefresh = 9 And cmdCloseServer.Caption <> "Re-Open Server" Then
    strTime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
        txtHTML.Text = ""
        Dim intTotal As Integer
        For i = 1 To 20
            If Users(i).Enabled = True Then
                intTotal = intTotal + 1
            End If
        Next 'i
        strtotal = CStr(intTotal)
        struptime = CStr(upTime)
        strLastUser = GetFromIni("GEN", "NEWEST", App.Path & "\users.ini")
        txtHTML.Text = "This webpage is updated every 15 minutes if the server is up.  If it hasn't been updated within the last few minutes, then the server is probably down for maitence.  Please don't post issues about the server being down on the Bug Report forum.<br><br>As of <b>" & strTime & "</b> the server is running <b>Version " & txtVer.Text & "</b> and has been running for <b>" & struptime & "</b> minutes.<br><br>There are currently <b>" & strtotal & "</b> users online.<br>The last user to register was <b>" & strLastUser & "</b><br><br>Message of the Day: <b>" & txtMOTD.Text & "</b>"
 FNO = FreeFile
     On Error Resume Next
     err.Clear
     Open ("C:\sambar50\docs\status.html") For Output As #FNO
      If err.Number <> 0 Then
        MsgBox "an error has occured"
       Else
        Print #FNO, (txtHTML.Text)
      End If
     Close #FNO
     On Error GoTo 0
End If

End Sub

Private Sub txtVer_Change()
Call WriteIni("MOTD", "VER", txtVer.Text, App.Path & "\motd.ini")
ServerVersion = txtVer.Text
End Sub
Sub AdminChat(ByVal strChatTxt As String)
On Error Resume Next
For q = 1 To 20
    If Users(q).Enabled = True Then
        Chat(q).SendData "ADMINTXT" & strChatTxt & vbCrLf
        DoEvents
    End If
Next 'q
End Sub
Sub ChatDisc()
On Error Resume Next
For q = 1 To 20
    If Users(q).Enabled = True Then
        Chat(q).SendData "CHATKILL" & strDiscUser & vbCrLf
        DoEvents
    End If
Next ' q
End Sub
Sub ChatSpam()
On Error Resume Next
For z = 1 To 20
    If Users(z).Enabled = True Then
        Chat(z).SendData "CHATNUM" & z & vbCrLf
        Chat(z).SendData "CHATNAME" & Users(z).Name & vbCrLf
        Chat(z).SendData "CHATRATING" & Users(z).Rating & vbCrLf
        Chat(z).SendData "CHATIP" & Users(z).IP & vbCrLf
        DoEvents
    End If
Next 'z
End Sub
Sub ChatMsg(ByVal strMsg As String, ByVal strUser As String)
On Error Resume Next
For q = 1 To 20
    If Users(q).Name = strUser Then
    
        Chat(q).SendData "CHATMSG" & "[You Recieved a Private Message]: " & strMsg & vbCrLf
        DoEvents
    End If
Next 'q
End Sub
Sub SendUsers()
On Error Resume Next

For i = 1 To 20
If Users(i).Enabled = True Then
    Chat(i).SendData "CHATSTART" & vbCrLf
    For q = 1 To 20
        If Users(q).Enabled = True Then
            Chat(i).SendData "CHATNAME" & Users(q).Name & vbCrLf
            DoEvents
        End If
    Next 'q
    Chat(i).SendData "CHATSTOP" & vbCrLf
End If
Next 'i

End Sub

Sub AddMUser(ByVal intCur As Integer, FirstLoad As Boolean)
On Error Resume Next

For i = 1 To 20
    If Users(i).Enabled = True Then
        Chat(i).SendData "ISAACCURNUM" & intCur & vbCrLf
        DoEvents
        If FirstLoad = True Then
            Chat(i).SendData "ISAACPIC" & Users(intCur).Pic & vbCrLf
            DoEvents
        End If
        Chat(i).SendData "ISAACSCREEN" & Users(intCur).Screen & vbCrLf
        DoEvents
        Chat(i).SendData "MOVEISAACX" & Users(intCur).Left & vbCrLf
        DoEvents
        Chat(i).SendData "MOVEISAACY" & Users(intCur).Top & vbCrLf
        DoEvents
    End If
Next 'i

End Sub
Sub KillChar(Player As Integer)
On Error Resume Next
For i = 1 To 20
    If Users(i).Enabled = True Then
        Chat(i).SendData "ISAACKILL" & Player & vbCrLf
        DoEvents
    End If
Next 'i
End Sub
Sub GetGameList(Client As Integer)
On Error Resume Next
For i = 1 To 20
    If Game(i).Enabled = True Then
        Chat(Client).SendData "GAMEADD" & Game(i).Name & vbCrLf
    End If
Next 'i
End Sub
Sub JoinGame(ByVal GameName As String, ByVal Index As Integer)
On Error Resume Next
For i = 1 To 20
    If GameName = Game(i).Name Then
        Chat(Index).SendData "JOINIP" & Game(i).IP
    End If
Next 'i
End Sub
Private Function FindUser(UserName As String) As Integer
    Dim intMax As Integer
    Dim nSave As String
    
    nSave = App.Path & "\user.ini"
    intMax = CInt(GetFromIni("GEN", "TOTAL", nSave))
    For i = 0 To intMax
        Dim strtempName As String
        strtempName = GetFromIni(CStr(i), "NAME", nSave)
        If strtempName = UserName Then
            FindUser = i
        End If
    Next 'i
    
End Function
Sub InGameChat(ByVal Index As Integer, ByVal strChat As String)
On Error Resume Next
    For i = 1 To 20
        If Users(i).Enabled = True Then
            Chat(i).SendData "INGAMECHATNUM" & CStr(Index) & vbCrLf
            Chat(i).SendData "INGAMECHATTXT" & strChat & vbCrLf
        End If
    Next 'i
End Sub

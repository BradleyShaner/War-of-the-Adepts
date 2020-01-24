VERSION 5.00
Begin VB.Form frmEndGame 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "You Win!"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   ControlBox      =   0   'False
   Icon            =   "frmEndGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFinish 
      Caption         =   "Finish"
      Enabled         =   0   'False
      Height          =   975
      Left            =   6240
      Picture         =   "frmEndGame.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll"
      DragIcon        =   "frmEndGame.frx":2E40
      Height          =   975
      Left            =   6240
      Picture         =   "frmEndGame.frx":370A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Caption         =   "Highlight a field to see an explanation on how it increases."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   6855
   End
   Begin VB.Label lblremn 
      BackStyle       =   0  'Transparent
      Caption         =   "Explanation:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblTries 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   5040
      TabIndex        =   17
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblremn 
      BackStyle       =   0  'Transparent
      Caption         =   "Rating Points Until Next Djinni:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   3480
      TabIndex        =   16
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblTries 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   3120
      TabIndex        =   15
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblremn 
      BackStyle       =   0  'Transparent
      Caption         =   "Rating Points Until Next Level:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   1680
      TabIndex        =   14
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblTries 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   13
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblremn 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Level:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblTries 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   6720
      TabIndex        =   11
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblTries 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   5160
      TabIndex        =   10
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblTries 
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   9
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblremn 
      BackStyle       =   0  'Transparent
      Caption         =   "Coins Gained:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblremn 
      BackStyle       =   0  'Transparent
      Caption         =   "Rating Points Gained:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblremn 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Rating:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblTries 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblremn 
      BackStyle       =   0  'Transparent
      Caption         =   "Tries Remaining:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image imgChoose 
      Height          =   1785
      Index           =   3
      Left            =   5400
      Picture         =   "frmEndGame.frx":5C85
      Top             =   1440
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image imgChoose 
      Height          =   1980
      Index           =   2
      Left            =   5400
      Picture         =   "frmEndGame.frx":7EC3
      Top             =   -240
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image imgChoose 
      Height          =   1650
      Index           =   1
      Left            =   5280
      Picture         =   "frmEndGame.frx":98A8
      Top             =   720
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image imgChoose 
      Height          =   2010
      Index           =   0
      Left            =   5520
      Picture         =   "frmEndGame.frx":B592
      Top             =   -480
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image imgItem 
      DragIcon        =   "frmEndGame.frx":D253
      Height          =   480
      Index           =   2
      Left            =   4320
      Picture         =   "frmEndGame.frx":DB1D
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem 
      DragIcon        =   "frmEndGame.frx":E3E7
      Height          =   480
      Index           =   1
      Left            =   2400
      Picture         =   "frmEndGame.frx":ECB1
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem 
      DragIcon        =   "frmEndGame.frx":F57B
      Height          =   480
      Index           =   0
      Left            =   480
      Picture         =   "frmEndGame.frx":FE45
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape shpslot 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   2175
      Index           =   2
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Shape shpslot 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   2175
      Index           =   1
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Shape shpslot 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   2175
      Index           =   0
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEndGame.frx":1070F
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Won!"
      DragIcon        =   "frmEndGame.frx":107C5
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmEndGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CoinsGained As Integer 'Coins... gained after battle
Dim RatingGained As Integer 'Rating...
Dim LvlGained As Integer 'Level...
Dim DjinnGained As Integer 'Djinn...
Dim bDrag As Boolean 'used for cheat
Dim TriesLeft As Integer

Private Sub cmdFinish_Click()
On Error Resume Next

DidNotRoll = False
frmChat.Chat.SendData "DIDROLL" & strMyUserName & vbCrLf

If DoubleStats = False Or DidIWin = True Then

    
    frmChat.Chat.SendData "WINUSER" & strMyUserName & vbCrLf
    DoEvents
    frmChat.Chat.SendData "RATING" & strMyUserName & "@" & RatingGained & vbCrLf
    DoEvents
    frmChat.Chat.SendData "COINS" & strMyUserName & "@" & CoinsGained & vbCrLf
    DoEvents
'    frmUser2.User.SendData "DJINNGAINED" & strMyUserName & "@" & DjinnGained & vbCrLf
'    DoEvents
'    frmUser2.User.SendData "LVL" & strMyUserName & "@" & LvlGained & vbCrLf
'    DoEvents
    frmChat.Chat.SendData "HANDIF" & strMyUserName & "@" & CStr(Handicap(1) - Handicap(2)) & vbCrLf
    DoEvents
    
    If DidIWin = True Then
        frmChat.Chat.SendData "SWIN" & strMyUserName & "@" & strOpponent & vbCrLf
        DoEvents
    Else
        frmChat.Chat.SendData "LOSE" & strMyUserName & "@" & strOpponent & vbCrLf
        DoEvents
    End If
    
    'iLvl = CInt(strLvl) + CInt(LvlGained)
    'strLvl = CStr(iLvl)
    
    Dim iRating As Integer
    iRating = CInt(strRating) + CInt(RatingGained)
    strRating = CInt(iRating)
    
    iCoins = CLng(strCoins) + CLng(CoinsGained)
    strCoins = CInt(iCoins)
    
    'iTotalDjinn = CInt(sTotalDjinn) + CInt(DjinnGained)
    'sTotalDjinn = CStr(iTotalDjinn)
    
    
End If

If DoubleStats = True And DidIWin = False Then
    frmChat.Chat.SendData "STATSLOSS" & strMyUserName & vbCrLf
    DoEvents
    frmChat.Chat.SendData "LOSSRATING" & strMyUserName & "@" & RatingGained & vbCrLf
    DoEvents
    frmChat.Chat.SendData "LOSSCOINS" & strMyUserName & "@" & CoinsGained & vbCrLf
    DoEvents
'   frmchat.chat.SendData "LOSSDJINN" & strMyUserName & "@" & DjinnGained & vbCrLf
'   DoEvents
    frmChat.Chat.SendData "HANDIF" & strMyUserName & "@" & CStr(Handicap(1) - Handicap(2)) & vbCrLf
    DoEvents
    frmChat.Chat.SendData "LOSE" & strMyUserName & "@" & strOpponent & vbCrLf
    DoEvents
    
    
    'iLvl = CInt(strLvl) - CInt(LvlGained)
    'strLvl = CStr(iLvl)
    
    iRating = CInt(strRating) - CInt(RatingGained)
    strRating = CInt(iRating)
    
    iCoins = CLng(strCoins) - CLng(CoinsGained)
    strCoins = CStr(iCoins)
    
    iTotalDjinn(1) = CInt(sTotalDjinn(1)) - CInt(DjinnGained)
    sTotalDjinn(1) = CStr(iTotalDjinn(1))
End If

iLvl = GetLevel(strRating)
strLvl = CStr(iLvl)

iTotalDjinn(1) = GetDjinn(strRating)
sTotalDjinn(1) = CStr(iTotalDjinn(1))


If DidIWin = True Then
Dim strMaxOp As String
Dim intMaxOp As Integer
Dim nOp As String
nOp = "C:\windows\system32\gsawota.sys"

strMaxOp = GetFromIni(strOpponent, strServerDate, nOp)
If strMaxOp = "" Then strMaxOp = "0"
intMaxOp = CStr(strMaxOp)
intMaxOp = intMaxOp + 1
Call WriteIni(strOpponent, strServerDate, CStr(intMaxOp), nOp)
'Call WriteIni(strServerDate, "OP" & CStr(intMaxOp), strOpponent, nOp)
'Call WriteIni(strServerDate, "IP" & CStr(intMaxOp), strOpponent, nOp)
End If

DidIWin = False
DoubleStats = False
WinBattle = False
'frmMultiplayer.Show
Unload frmBattle
Unload frmArena
Unload Me

End Sub

Private Sub cmdRoll_Click()
On Error Resume Next
If TriesLeft > 0 Then
Dim yRand(1 To 3) As Integer
CoinsGained = 0

Randomize
If DidIWin = True Then
    RatingGained = Int(Rnd * 3) + 1
Else
    RatingGained = 0
End If

LvlGained = 0
DjinnGained = 0
imgItem(0).Visible = True
imgItem(1).Visible = True
imgItem(2).Visible = True

TriesLeft = TriesLeft - 1
lblTries(0).Caption = TriesLeft

If TriesLeft = 0 Then cmdRoll.Enabled = False

yRand(1) = Int(Rnd * 3)
yRand(2) = Int(Rnd * 3)
yRand(3) = Int(Rnd * 3)

imgItem(0).Picture = imgChoose(yRand(1)).Picture
imgItem(1).Picture = imgChoose(yRand(2)).Picture
imgItem(2).Picture = imgChoose(yRand(3)).Picture


For i = 1 To 3
If DidIWin = True Then
    If yRand(i) = 0 Then
    CoinsGained = CoinsGained + 65
    RatingGained = RatingGained + 3
    End If
    If yRand(i) = 1 Then
    CoinsGained = CoinsGained + 26
    RatingGained = RatingGained + 12
    End If
    If yRand(i) = 2 Then
    CoinsGained = CoinsGained + 30
    RatingGained = RatingGained + 2
    'LvlGained = LvlGained + 1
    End If
    If yRand(i) = 3 Then
    CoinsGained = CoinsGained + 45
    RatingGained = RatingGained + 9
    End If
End If

If DidIWin = False And DoubleStats = False Then
    If yRand(i) = 0 Then
    CoinsGained = CoinsGained + 22
    RatingGained = RatingGained + 0
    End If
    If yRand(i) = 1 Then
    CoinsGained = CoinsGained + 14
    RatingGained = RatingGained + 0
    End If
    If yRand(i) = 2 Then
    CoinsGained = CoinsGained + 19
    RatingGained = RatingGained + 0
    End If
    If yRand(i) = 3 Then
    CoinsGained = CoinsGained + 16
    RatingGained = RatingGained + 0
    End If
End If

If DidIWin = False And DoubleStats = True Then
    If yRand(i) = 0 Then
    CoinsGained = CoinsGained + 5
    RatingGained = RatingGained + 3
    End If
    If yRand(i) = 1 Then
    CoinsGained = CoinsGained + 6
    RatingGained = RatingGained + 12
    End If
    If yRand(i) = 2 Then
    CoinsGained = CoinsGained + 7
    RatingGained = RatingGained + 2
    'LvlGained = LvlGained + 1
    End If
    If yRand(i) = 3 Then
    CoinsGained = CoinsGained + 8
    RatingGained = RatingGained + 9
    End If
End If


Next 'i

'Gain only 1 level maximum
If LvlGained > 1 Then LvlGained = 1

'Determine rating point increase depending on how much above or below the opponent you were


Dim dif As Variant
If DidIWin = True Or DoubleStats = False Then
    If RelativeRating(1) - RelativeRating(2) > 0 Then 'I'm better than my oponent
        dif = 20 / (RelativeRating(1) - RelativeRating(2))
        If dif > 1 Then dif = 1
    ElseIf RelativeRating(2) - RelativeRating(1) = 0 Then 'We're evenly matched
        dif = 1
    Else 'My opponent is better than me
        dif = (RelativeRating(2) - RelativeRating(1)) / 35
        dif = dif + 1
    End If
End If

If Handicap(1) <> 0 And Handicap(2) <> 0 Then
    dif = 1
End If

If DidIWin = False And DoubleStats = True Then
    If RelativeRating(1) - RelativeRating(2) > 0 Then 'I'm better than my oponent
        dif = (RelativeRating(1) - RelativeRating(2)) / 10
        If dif < 3 Then dif = 3
    ElseIf RelativeRating(2) - RelativeRating(1) = 0 Then 'We're evenly matched
        dif = 1
    Else 'My opponent is better than me
        dif = 1
    End If
End If



If DidIWin = True Then


    'If LvlGained > 1 Then LvlGained = 1 'Gain only 1 level
    

        RatingGained = RatingGained * dif
        CoinsGained = CoinsGained * dif
    
    If DoubleStats = True Then 'Double or nothing?
        CoinsGained = CoinsGained * 2
        RatingGained = RatingGained * 2
        'LvlGained = LvlGained * 2
    End If

    
    If RatingGained > 75 Then RatingGained = 75
    If CoinsGained > 300 Then CoinsGained = 300

    If RatingGained > 0 Then
'    Dim intNewLevel As Integer
'    Dim intNewDjinn As Integer
'    intNewLevel = (myRating + RatingGained - 1000) / 50
'    If intNewLevel < 1 Then intNewLevel = 1
'    intNewDjinn = (myRating + RatingGained - 1000) / 100
'    If intNewDjinn < 1 Then intNewDjinn = 1
    
'    LvlGained = intNewLevel - CInt(strLvl)
'    DjinnGained = intNewDjinn - CInt(iDjinnTotal)
    
    
    LvlGained = 0
    For i = 1 To RatingGained
        If (CStr(strRating) + i) Mod 50 = 0 And myRating > 1000 Then
            LvlGained = LvlGained + 1
        End If
    Next 'i
    
    End If
    If DoubleStats = True Then 'Double or nothing?
        CoinsGained = CoinsGained * 2
        RatingGained = RatingGained * 2
        LvlGained = LvlGained * 2
    End If
    RatingGained = CInt(RatingGained)
    CoinsGained = CInt(CoinsGained)

    lblTries(2).Caption = RatingGained
    lblTries(3).Caption = CoinsGained
    lblTries(5).Caption = CInt(((25 + CInt(strLvl) ^ 1.5)))

Else
    RatingGained = CInt(RatingGained)
    CoinsGained = CInt(CoinsGained)
    If DoubleStats = True Then 'Deduct stats
        RatingGained = RatingGained * dif
        CoinsGained = CoinsGained * dif

        lblTries(2).Caption = "-" & RatingGained
        lblTries(3).Caption = "-" & CoinsGained
        lblTries(5).Caption = CInt(((25 + CInt(strLvl) ^ 1.5)))
    Else
        lblTries(2).Caption = RatingGained
        lblTries(3).Caption = CoinsGained
        lblTries(5).Caption = CInt(((25 + CInt(strLvl) ^ 1.5)))
        
    End If
    LvlGained = 0
    'For i = 1 To RatingGained
    '    If (CStr(strRating) - i) Mod 50 = 0 And myRating > 1000 Then
    '        LvlGained = LvlGained - 1
    '    End If
    'Next 'i
    
End If

'If DidIWin = True Then
'    If LvlGained > 0 Then
'        For i = 1 To LvlGained
'            If (CInt(strLvl) + i) Mod 2 = 0 Then 'If you've gained a level that is a multiple of 2
'               DjinnGained = DjinnGained + 1 'Increase number of Djinn that you have
'           End If
'        Next 'i
'    End If
'End If
'If DidIWin = False And DoubleStats = True Then
'    If LvlGained > 0 Then
'        For i = 1 To LvlGained
'            If (CInt(strLvl) - i) Mod 2 = 0 Then 'If you've gained a level that is a multiple of 2
'               DjinnGained = DjinnGained + 1 'Increase number of Djinn that you have
'           End If
'        Next 'i
'    End If
'End If

'lblTries(6).Caption = DjinnGained

If CInt(strmylevel) Mod 2 = 0 Then 'If the next level is odd (gain Djinn every 2 levels)
    lblTries(6).Caption = CInt(((25 + CInt(strLvl) ^ 1.5)))
Else 'If the next level is even
    lblTries(6).Caption = CInt(((25 + CInt(strLvl) ^ 1.5) + ((25 + (CInt(strLvl) + 1) ^ 1.5))))
End If

cmdFinish.Enabled = True

Else
MsgBox "No more tries left!"
End If

End Sub

Private Sub cmdRoll_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Beep
bDrag = True
End Sub

Private Sub Form_Activate()
On Error Resume Next
BattleLoaded(1) = False
BattleLoaded(2) = False
myRating = CInt(RelativeRating(1))
lblTries(1).Caption = strRating
End Sub

Private Sub Form_Load()
On Error Resume Next
'COMMENT OUT BEFORE RELEASING
'opRating = 1020
'strRating = "1000"
'strLvl = "1"
'strMyUserName = "ikillkenny"

Call PlayMidi("credits", True)

TriesLeft = 3
If DidIWin = True Then
    lblDesc.Caption = "It's now time to improve your stats.  You have three tries to play the slot machine below to get a good combination and the maximum number of Experience Points, Coins and Levels."
    lblTitle.Caption = "You Win!"
    Me.Caption = "You Win!"
Else
    If DoubleStats = False Then 'if you don't lose stats
        lblDesc.Caption = "It's now time to improve your stats.  You have three tries to play the slot machine below to get a good combination and the maximum number of Experience Points, Coins and Levels."
    Else
        lblDesc.Caption = "You now have to roll to determine how much stats you lose.  The objective is to lose as little stats as possible."
    End If
    lblTitle.Caption = "You Lost!"
    Me.Caption = "You Lost!"
End If

If DoubleStats = True And DidIWin = False Then
    DidNotRoll = True
    frmChat.Chat.SendData "DIDNOTROLL" & strMyUserName & vbCrLf
End If

cmdFinish.Enabled = False
cmdRoll.Enabled = True

myRating = CInt(RelativeRating(1))
lblTries(1).Caption = strRating

lblTries(4).Caption = strLvl
'opRating = CInt(stropRating)

frmEndGame.Picture = frmIntro.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
StopMidi
frmIntro.Show
End Sub

Private Sub imgItem_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
bDrag = False
End Sub

Private Sub lblTitle_DragDrop(Source As Control, X As Single, Y As Single)
If bDrag = True Then
    MsgBox "For demonstrating your ability to follow directions from the Typing Test game, you will get a demonstration character by entering the password: yellowtrash.", vbInformation, "Easter Egg #5"
    Call Encode("5", "EGG5", "EGGL5", App.Path & "\settings.ini")
    
End If
End Sub

Private Sub lblTries_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 3 Then
    bDrag = False
    lblTries(3).Drag
End If
End Sub

Private Sub lblTries_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Index = 2 Then
    lblExplain.Caption = "Rating points increase randomly and are affected by the relative rating of the person you beat.  You harder the match for you, the more points you gain."
ElseIf Index = 3 Then
    lblExplain.Caption = "Coins increase randomly and are affected by the relative rating of the person you beat.  You harder the match for you, the more coins you gain."
ElseIf Index = 5 Then
    lblExplain.Caption = "Level is increased automaticaly when you gain 50 Rating Points.  You may only gain 1 level at a time."
ElseIf Index = 6 Then
    lblExplain.Caption = "Djinn automaticaly increase every 2 levels that you gain."
End If
End Sub


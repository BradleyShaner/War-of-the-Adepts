VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Golden Sun: The War of the Adepts - Credits"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   328
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   485
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timeGame 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   1320
      Top             =   2280
   End
   Begin VB.Timer timeCredits 
      Interval        =   35
      Left            =   3240
      Top             =   3000
   End
   Begin VB.Label lblPlayer 
      BackStyle       =   0  'Transparent
      Caption         =   "Ikillkenny"
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
      Height          =   195
      Left            =   2160
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblBestPlayer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Best Player"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2160
      TabIndex        =   17
      Top             =   4320
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblGameMode 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Save All Djinn!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   405
      Left            =   0
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   7320
   End
   Begin VB.Image imgDeadBeard 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   0
      Picture         =   "frmAbout.frx":08CA
      Top             =   3600
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblHScore 
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
      Height          =   360
      Left            =   1200
      TabIndex        =   15
      Top             =   4560
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label lblHighScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "High Score:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   14
      Top             =   4320
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblScore 
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
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   4560
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label lblCurrentScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Score:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image imgDjinn 
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "frmAbout.frx":1A5E
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDjinnPic 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   3
      Left            =   0
      Picture         =   "frmAbout.frx":2328
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDjinnPic 
      Height          =   480
      Index           =   2
      Left            =   0
      Picture         =   "frmAbout.frx":2BF2
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDjinnPic 
      Height          =   480
      Index           =   1
      Left            =   360
      Picture         =   "frmAbout.frx":34BC
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblRules 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Instructions"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3840
      TabIndex        =   11
      Top             =   4680
      Width           =   2130
   End
   Begin VB.Image imgDjinnPic 
      Height          =   480
      Index           =   0
      Left            =   6840
      Picture         =   "frmAbout.frx":3D86
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblDjinn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click To Play Save The Djinn!"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3840
      TabIndex        =   10
      Top             =   4320
      Width           =   2130
   End
   Begin VB.Label lblPerson 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "gsa@comicsoft.zzn.com"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   7320
   End
   Begin VB.Label lblPerson 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "In Microsoft Visual Basic 6.0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   7320
   End
   Begin VB.Label lblPerson 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "With Tacvek  "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   7320
   End
   Begin VB.Label lblPerson 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mike Bentley"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   7320
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Original Programmer"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   3225
   End
   Begin VB.Shape shpBorder 
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   1
      Left            =   0
      Top             =   4200
      Width           =   7335
   End
   Begin VB.Shape shpBorder 
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label lblgen 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "gsa@comicsoft.zzn.com"
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
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   6960
      Width           =   3750
   End
   Begin VB.Label lblgen 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Ikillkenny"
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
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   6480
      Width           =   3750
   End
   Begin VB.Label lblgen 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "[ Programmed By: ]"
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
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   3750
   End
   Begin VB.Label lblgen 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "- - - - - - - - - - - - - - - - - -"
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
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   3750
   End
   Begin VB.Label lblgen 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Golden Sun Online Battle"
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   3750
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim curSpeed As Integer
Dim curScore As Integer
Dim curShown As Integer

Dim bWait As Integer
Dim bounceUp As Boolean
Dim bounceTo As Integer
Dim bTitle As Boolean
Dim curScroll As Integer

Private Sub Form_Activate()
lblHScore.Caption = intDjinnSaveHighScore
lblPlayer.Caption = strDjinnSavePlayer
End Sub

Private Sub Form_Load()
On Error Resume Next
curScroll = 1
bTitle = True
bounceUp = False
bWait = 50
Call PlayMidi("credits", True)
For i = 1 To 25
    Load imgDjinn(i)
    imgDjinn(i).Visible = False
    imgDjinn(i).Picture = imgDjinnPic(0).Picture
    imgDjinn(i).Top = 0
    imgDjinn(i).Left = 0
Next 'i

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgDeadBeard.Visible = True Then
    imgDeadBeard.Left = X - 18.5
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
timeGame.Enabled = False
StopMidi
End Sub

Private Sub imgDeadBeard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'imgDeadBeard.Left = imgDeadBeard.Left + (17 - (X / Screen.TwipsPerPixelX))
End Sub

Private Sub imgDjinn_Click(Index As Integer)
If imgDeadBeard.Visible = False Then
    Call DestroyDjinn(Index)
End If

End Sub

Private Sub lblDjinn_Click()
On Error Resume Next
If strMyUserName <> "" Then
    curSpeed = 1
    curScore = 0
    curShown = 2
    timeGame.Enabled = True
    lblScore.Caption = 0
    lblCurrentScore.Visible = True
    lblScore.Visible = True
    lblDjinn.Visible = False
    lblRules.Visible = False
    timeCredits.Enabled = False
    lblPerson(0).Visible = False
    lblPerson(1).Visible = False
    lblPerson(2).Visible = False
    lblPerson(3).Visible = False
    lblTitle.Visible = False
    lblHighScore.Visible = True
    lblHScore.Visible = True
    lblBestPlayer.Visible = True
    lblPlayer.Visible = True
    lblPlayer.Caption = strDjinnSavePlayer
    For i = 1 To imgDjinn.UBound
        imgDjinn(i).Visible = False
    Next 'i
    imgDeadBeard.Visible = False
Else
    MsgBox "Sorry, you must be logged into the server in order to determine the current high score."
End If
End Sub

Private Sub lblPerson_Click(Index As Integer)
If Index = 0 And lblPerson(0).Caption = "Mike Bentley" Then
    MsgBox "You just knew that my name was going to be the #1 easter egg, didn't you?", vbInformation, "Easter Egg #1"
    Call Encode("1", "EGG1", "EGGL1", App.Path & "\settings.ini")
End If
End Sub

Private Sub lblRules_Click()
MsgBox "Save the Djinn is a simple game: All you need to do is to prevent the Djinn falling down from the top of the screen from hittng the bottom black bar.  Simply click the Djinn to 'save' them.  The game will continue to play, getting progressively harder as time goes on until one of the Djinn fall to its death.  Highest scores are uploaded to the server and your name will be displayed on the ladder page if you have the top score.", vbInformation, "Instructions"
End Sub

Private Sub timeCredits_Timer()
On Error Resume Next
If bTitle = True Then
    lblTitle.Left = lblTitle.Left + 6
    If bounceUp = False Then
        lblTitle.Top = lblTitle.Top + 5
        If lblTitle.Top + lblTitle.Height >= shpBorder(1).Top Then
            bounceTo = Int(Rnd * (shpBorder(1).Top - shpBorder(0).Top))
            bounceTo = bounceTo + shpBorder(0).Top
            bounceUp = True
        End If
    Else
        lblTitle.Top = lblTitle.Top - 3
        If lblTitle.Top <= bounceTo Then
            bounceUp = False
        End If
    End If
    If lblTitle.Left >= Me.ScaleWidth Then
        bTitle = False
        Call GetPeople(curScroll)
        lblPerson(0).Left = 0 - lblPerson(0).Width
        lblPerson(1).Left = Me.ScaleWidth
        lblPerson(2).Left = 0 - lblPerson(2).Width
        lblPerson(3).Left = Me.ScaleWidth
        bWait = 50
    End If
Else
    If bWait = 50 Then
        If lblPerson(0).Visible = True Then
            lblPerson(0).Left = lblPerson(0).Left + 10
        End If
        If lblPerson(1).Visible = True Then
            lblPerson(1).Left = lblPerson(1).Left - 10
        End If
        If lblPerson(2).Visible = True Then
            lblPerson(2).Left = lblPerson(2).Left + 10
        End If
        If lblPerson(3).Visible = True Then
            lblPerson(3).Left = lblPerson(3).Left - 10
        End If
        If lblPerson(0).Left >= 0 Then
            lblPerson(0).Left = 0
            lblPerson(1).Left = 0
            lblPerson(2).Left = 0
            lblPerson(3).Left = 0
            bWait = 49
        End If
    End If
    If bWait <> 50 And bWait <> 0 Then
        bWait = bWait - 1
    End If
    If bWait = 0 Then
        If lblPerson(0).Visible = True Then
            lblPerson(0).Left = lblPerson(0).Left + 10
        End If
        If lblPerson(1).Visible = True Then
            lblPerson(1).Left = lblPerson(1).Left - 10
        End If
        If lblPerson(2).Visible = True Then
            lblPerson(2).Left = lblPerson(2).Left + 10
        End If
        If lblPerson(3).Visible = True Then
            lblPerson(3).Left = lblPerson(3).Left - 10
        End If
        If lblPerson(0).Left >= Me.ScaleWidth Then
            bTitle = True
            bounceUp = False
            curScroll = curScroll + 1
            Call GetTitle(curScroll)
            lblTitle.Left = 0 - lblTitle.Width
            lblTitle.Top = 144
        End If
    End If
End If
If curScroll = 13 Then
    timeCredits.Enabled = False
    curScroll = 0
    Unload Me
End If
        
End Sub
Sub GetTitle(intCur As Integer)
    lblPerson(0).Visible = False
    lblPerson(1).Visible = False
    lblPerson(2).Visible = False
    lblPerson(3).Visible = False
With lblTitle
If intCur = 2 Then
    .Caption = "Resurrected By"
End If
If intCur = 3 Then
    .Caption = "Sound By"
End If
If intCur = 4 Then
    .Caption = "Tested By"
End If
If intCur = 5 Then
    .Caption = "Tested By"
End If
If intCur = 6 Then
    .Caption = "Music By"
End If
If intCur = 7 Then
    .Caption = "Name Thought Of By"
End If
If intCur = 8 Then
    .Caption = "In Memory Of"
End If
If intCur = 9 Then
    .Caption = "Original Game By"
End If
If intCur = 10 Then
    .Caption = "Developed By"
End If
If intCur = 11 Then
    .Caption = "Produced By"
End If
If intCur = 12 Then
    .Caption = "Exclusive Content Of"
End If
If intCur = 13 Then
    .Caption = "Legal Junk"
End If
End With
End Sub
Sub GetPeople(intCur As Integer)
If intCur = 1 Then
    lblPerson(0).Visible = True
    lblPerson(1).Visible = True
    lblPerson(2).Visible = True
    lblPerson(3).Visible = True
End If
If intCur = 2 Then
    lblPerson(0).Caption = "Dragoon"
    lblPerson(1).Caption = "6 Years later"
    lblPerson(2).Caption = "Yay, I know some VB!"
    lblPerson(3).Caption = "Intro By: IKK"
    lblPerson(0).Visible = True
    lblPerson(1).Visible = True
    lblPerson(2).Visible = True
    lblPerson(3).Visible = False
End If
If intCur = 3 Then
    lblPerson(0).Caption = "Sound Effects From"
    lblPerson(1).Caption = "earthstation1.simplenet.com"
    lblPerson(2).Caption = "Music Files Ripped From"
    lblPerson(3).Caption = "Original Game"
    lblPerson(0).Visible = True
    lblPerson(1).Visible = True
    lblPerson(2).Visible = True
    lblPerson(3).Visible = True
End If
If intCur = 4 Then
    lblPerson(0).Caption = "Ikillkenny"
    lblPerson(1).Caption = "Maverik"
    lblPerson(2).Caption = "CypherJF"
    lblPerson(3).Caption = "Sharker2001"
    lblPerson(0).Visible = True
    lblPerson(1).Visible = True
    lblPerson(2).Visible = True
    lblPerson(3).Visible = True
End If
If intCur = 5 Then
    lblPerson(0).Caption = "All of the Beta Testers"
    lblPerson(1).Caption = "Special Thanks To"
    lblPerson(2).Caption = "EA Leader (Beta Ladder Winner)"
    lblPerson(3).Caption = "Black Ice (Leading Bug Reporter)"
    lblPerson(0).Visible = True
    lblPerson(1).Visible = True
    lblPerson(2).Visible = True
    lblPerson(3).Visible = True
End If
If intCur = 6 Then
    lblPerson(0).Caption = "JeffreyAtW (jeffreyatw@hotmail.com)"
    lblPerson(1).Caption = "Lychee Akana (Lychee@Ranmamail.com)"
    lblPerson(2).Caption = "starfoxillusion (zorazora3@yahoo.com)"
    lblPerson(3).Caption = "David R. Clark"
    lblPerson(0).Visible = True
    lblPerson(1).Visible = True
    lblPerson(2).Visible = True
    lblPerson(3).Visible = True
End If
If intCur = 7 Then
    lblPerson(0).Caption = "Black Ice"
    lblPerson(1).Caption = ""
    lblPerson(2).Caption = ""
    lblPerson(3).Caption = ""
    lblPerson(0).Visible = True
    lblPerson(1).Visible = False
    lblPerson(2).Visible = False
    lblPerson(3).Visible = False
End If
If intCur = 8 Then
    lblPerson(0).Caption = "The Late GSVA Brett"
    lblPerson(1).Caption = ""
    lblPerson(2).Caption = ""
    lblPerson(3).Caption = ""
    lblPerson(0).Visible = True
    lblPerson(1).Visible = False
    lblPerson(2).Visible = False
    lblPerson(3).Visible = False
End If
If intCur = 9 Then
    lblPerson(0).Caption = "Nintendo"
    lblPerson(1).Caption = "Camelot Software Planning"
    lblPerson(2).Caption = "Original Game © 2001, 2002"
    lblPerson(3).Caption = "Used Without Permission"
    lblPerson(0).Visible = True
    lblPerson(1).Visible = True
    lblPerson(2).Visible = True
    lblPerson(3).Visible = True
End If
If intCur = 10 Then
    lblPerson(0).Caption = "Doc Entertainment"
    lblPerson(1).Caption = "www.doc-ent.com"
    lblPerson(2).Caption = "A Small Division of"
    lblPerson(3).Caption = "Doc Inc."
    lblPerson(0).Visible = True
    lblPerson(1).Visible = True
    lblPerson(2).Visible = True
    lblPerson(3).Visible = True
End If
If intCur = 11 Then
    lblPerson(0).Caption = "Doc Incorporated"
    lblPerson(1).Caption = "Golden Sun Anonymous"
    lblPerson(2).Caption = ""
    lblPerson(3).Caption = ""
    lblPerson(0).Visible = True
    lblPerson(1).Visible = True
    lblPerson(2).Visible = False
    lblPerson(3).Visible = False
End If
If intCur = 12 Then
    lblPerson(0).Caption = "Golden Sun Anonymous"
    lblPerson(1).Caption = "http://www.doc-ent.com/gsa"
    lblPerson(2).Caption = ""
    lblPerson(3).Caption = ""
    lblPerson(0).Visible = True
    lblPerson(1).Visible = True
    lblPerson(2).Visible = False
    lblPerson(3).Visible = False
End If
If intCur = 13 Then
    lblPerson(0).Caption = "War of the Adepts"
    lblPerson(1).Caption = "© 2002, 2003 Michael Bentley"
    lblPerson(2).Caption = "This Game Is Not Affiliated"
    lblPerson(3).Caption = "Nor Endorsed by Nintendo"
    lblPerson(0).Visible = True
    lblPerson(1).Visible = True
    lblPerson(2).Visible = True
    lblPerson(3).Visible = True
End If
End Sub

Private Sub timeGame_Timer()
On Error Resume Next
Randomize
Dim intRand As Integer
If lblGameMode.Visible = True Then
    lblGameMode.Visible = False
End If
For i = 1 To curShown
    If imgDjinn(i).Visible = False Then
        intRand = Int(Rnd * 4)
        imgDjinn(i).Visible = True
        imgDjinn(i).Picture = imgDjinnPic(intRand).Picture
        intRand = Int(Rnd * (Me.ScaleWidth - imgDjinnPic(0).Width)) + 1
        imgDjinn(i).Left = intRand
        imgDjinn(i).Top = 25
    End If
    If imgDjinn(i).Visible = True Then
        imgDjinn(i).Top = imgDjinn(i).Top + curSpeed
        If imgDjinn(i).Top + imgDjinn(i).Height >= shpBorder(1).Top Then
            timeGame.Enabled = False
            If curScore <= intDjinnSaveHighScore Then
                MsgBox "You lose! The current high score is " & intDjinnSaveHighScore
            Else
                intDjinnSaveHighScore = curScore
                MsgBox "Congratulations!  You have achieved a high score of " & curScore
                frmUser2.User.SendData "HIGHPLAYERDS" & strMyUserName & vbCrLf
                frmUser2.User.SendData "HIGHSCOREDS" & intDjinnSaveHighScore & vbCrLf
                lblHScore.Caption = curScore
                lblPlayer.Caption = strMyUserName
            End If
            Call HideDjinn
            Exit Sub
        End If
        If imgDjinn(i).Left + imgDjinn(i).Width >= imgDeadBeard.Left And imgDjinn(i).Left <= imgDeadBeard.Left + imgDeadBeard.Width And imgDjinn(i).Top + imgDjinn(i).Height >= imgDeadBeard.Top Then
            Call DestroyDjinn(CInt(i))
        End If
    End If
Next 'i

End Sub
Sub HideDjinn()
For i = 1 To 25
    imgDjinn(i).Visible = False
    imgDjinn(i).Top = 0
    imgDjinn(i).Left = 0
Next 'i
lblCurrentScore.Visible = False
lblScore.Visible = False
lblDjinn.Visible = True
lblRules.Visible = True
timeCredits.Enabled = True
lblPerson(0).Visible = True
lblPerson(1).Visible = True
lblPerson(2).Visible = True
lblPerson(3).Visible = True
lblBestPlayer.Visible = False
lblPlayer.Visible = False
lblTitle.Visible = True
lblHighScore.Visible = False
lblHScore.Visible = False
End Sub
Sub DestroyDjinn(intDjinn As Integer)
On Error Resume Next
imgDjinn(intDjinn).Visible = False
curScore = curScore + 5
lblScore.Caption = curScore

'Randomize
'Dim intRand As Long
'intRand = Int(Rnd * 8) + 1
'If intRand = 1 Then
'    If imgDeadBeard.Visible = False Then
'        lblGameMode.Visible = True
'        lblGameMode.Caption = "Kill All Djinn!  Speed Increase!"
'        imgDeadBeard.Visible = True
'        curSpeed = curSpeed + 2
'        curShown = curShown + 2
'    Else
'        lblGameMode.Visible = True
'        lblGameMode.Caption = "Save All Djinn!"
'        imgDeadBeard.Visible = False
'        curSpeed = curSpeed - 2
'        curShown = curShown - 2
'    End If
'End If

If curScore Mod 100 = 0 Then
    curSpeed = curSpeed + 1
End If
If curScore Mod 50 = 0 Then
    curShown = curShown + 1
End If

If curShown > 25 Then curShown = 25

If curSpeed > 20 Then curSpeed = 20

End Sub

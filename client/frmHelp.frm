VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Golden Sun: The War of the Adepts - Help"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6435
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHelp 
      Height          =   2535
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   3855
   End
   Begin VB.ListBox lstHelp 
      Height          =   2985
      ItemData        =   "frmHelp.frx":030A
      Left            =   240
      List            =   "frmHelp.frx":0332
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblSoon 
      BackStyle       =   0  'Transparent
      Caption         =   "Coming Soon: Online Help with Pictures and clearer text!"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   6015
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Basics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Picture = frmIntro.Picture
End Sub

Private Sub lstHelp_Click()
If lstHelp.Text = "Basics" Then
txtHelp.Text = "To play the game you will need to log in." & vbNewLine & "To do this, simply hit Log-In after hitting Multiplayer." & vbNewLine & "If you have not yet created user, you will need to do so." & vbNewLine & "Once you've logged in, you will now be able to host or join a game." & vbNewLine & "Once you've connected to another user, a game can start and you two will play."
End If
If lstHelp.Text = "Logging In" Then
txtHelp.Text = "To log in, go to the User Configuration screen from the main window." & vbNewLine & "Once you do that, a window will come up that will allow you to log onto the server, create a new account, or request a password change." & vbNewLine & "Once you've logged in, you will be able to configure your Psynergy, Djinn and Items as well as Host or Join a game." & vbNewLine & "If you can't log in, please click the 'Get Current Serve IP Address' button under Beta Options, then enter in the address on that page under 'Set Master Server IP Address'."
End If
If lstHelp.Text = "Finding Games" Then
txtHelp.Text = "To find a game, log-in to the server.  Head to the southern-most house and talk with the man in here.  To find games, simply hit the Find Games Button.  You can also manually join a game by getting someone's IP from the chat."
End If
If lstHelp.Text = "Starting Games" Then
txtHelp.Text = "To start a game, log-in to the server and in the Online Town head to the southernmost house.  Walk into the man there, and hit Create Game." & vbNewLine & "Enter the name of your game and then hit start game.  From here, wait for people to join.  You can also use the chat to encourage people to join your game."
End If
If lstHelp.Text = "Configuring Your Character" Then
txtHelp.Text = "To configure your character, log-in to the server and head to the Inn in the Online Town.  Talk with the man in the inn and you will be able to change your character for a set number of coins and examine what character that you have."
End If
If lstHelp.Text = "The Battle" Then
txtHelp.Text = "The battle is the main part of this game.  To get to the battle, log-in, then either join or host a game.  Have both parties hit ready, then start the game." & vbNewLine & "In the battle, you will have the option of Attacking, using Psynergy, Djinn or Summons." & vbNewLine & "Once both players have chosen an attack, the damage will be done and the battle will go on."
End If
If lstHelp.Text = "After Battle" Then
txtHelp.Text = "After the battle, you will be able to 'Roll' the slot machines in order to randomly generate stats." & vbNewLine & "Once you've decided that you like the stats that you've rolled, hit Finish to send the stats to the server." & vbNewLine & "At this time, you need to log back in in order to have your stats be updated."
End If
If lstHelp.Text = "Other Help" Then
txtHelp.Text = "Use the Report Bug and Report Suggestions often!  Feedback is very essential when creating a game." & vbNewLine & "Please check often that you have the latest version of the game.  Eventually, if you do not have the latest version you will not be able to play without updating." & vbNewLine & "You aren't allowed to play yourself in this game."
End If
If lstHelp.Text = "Pre-Battle Help" Then
txtHelp.Text = "Before the battle you must try to win a race!  The races offer a series of different events.  Use the arrow keys to move your character around.  The first character to finish the race gets a 5% stat bonus in addition to extra rating points."
End If
If lstHelp.Text = "Online Town Help" Then
txtHelp.Text = "The Online Town allows players to walk around and interact with other characters online at the time." & vbNewLine & "Use the Arrow Keys to move.  You can also talk by pressing Enter." & vbNewLine & "To speak with a non-playable character, just walk into them."
End If
If lstHelp.Text = "Frequent Issues" Then
txtHelp.Text = "Q) I am sure that my username and password are correct but it won't let me in." & vbNewLine & "A) Update your version." & vbNewLine & "Q) The game crashes when I log-in." & vbNewLine & "A) Download the Full Install." & vbNewLine & "Q) The Online Town and Pre-Battle Arena don't load correctly." & vbNewLine & "A) Make sure all .dat files are in the same directory as the program."
End If
If lstHelp.Text = "Single Player Help" Then
txtHelp.Text = "Note: This help file is only concerning saving and loading characters to and from the Single Player." & vbNewLine & "To save a character for use in the Single Player, make sure to log in first.  Then, go back to the Main Screen and hit Single Player.  Choose Save Character if you want to save your character to the Single Player or hit Load Character if you want to get back coins you earned.  Coins will be sent to the server automatically."
End If
lblTitle.Caption = lstHelp.Text
End Sub

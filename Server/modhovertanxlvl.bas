Attribute VB_Name = "modHoverLVL"
'New Variables
Public Luck(1 To 2) As Boolean



'Constants for Bitblt
Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

Public strItemName(1 To 50) As String 'Item Name
Public strItemDesc(1 To 50) As String 'Item Description
Public strItemCoins(1 To 50) As String 'Item Cost
Public strItemType(1 To 50) As String 'Item Elemental Type
Public strItemDamage(1 To 50) As String 'Item Damage
Public strItemSpcType(1 To 50) As String 'Item Critical Hit Type
Public strItemSpcDamage(1 To 50) As String 'Item Critical Hit Damage
Public strItemSpcDesc(1 To 50) As String 'Item Critical Hit Description

Dim strLastGetFromIni As String
Dim strLastGetFromIni2 As String

Dim strdtime As String 'Current time

Dim rsave As String 'Userdata.ini save directory

Public curIsaac As Integer 'Current player in Online Town

Public IsaacM(1 To 20) As GameTile '1-20 Players in Online Town


Public strMyUserName As String 'Current user name

Dim iCurTurn As Integer '???? I don't think this is used anymore, it will be removed once I figure out what I used it for

Public AmIKilled As Boolean 'Am I dead?

Public strDjinnName(1 To 10) As String 'Djinn Name
Public strDjinnDesc(1 To 10) As String 'Djinn Description
Public strDjinnType(1 To 10) As String 'Djinn Elemental Type
Public strDjinnDamage(1 To 10) As String 'Djinn Damage
Public bDjinnSet(1 To 10) As Boolean 'Is the Djinn Set or Standby?
Public iCurDjinn As Integer 'Total Djinn (integer)
Public sCurDjinn As String 'Total Djinn (string)
Public intDjinnStandby As Integer 'Total Djinn on Standby


Public Reset(1 To 2) As Boolean 'Did the player call for a reset of battle?


Public strOpDjinnType As String 'The picture of the opponent's Djinn being used

Public Type ChatUser
    Name As String
    Rating As String
    IP As String
    Enabled As Boolean 'Visible or not
    Wins As String
    Losses As String
    Disconnects As String
    Pic As String 'Character picture
    Left As String 'X position
    Top As String 'Y position
    Screen As String 'Current screen
    Number As Long 'Number in user.ini
    Avatar As String
    Away As Boolean
    Moderator As Boolean
    Admin As Boolean
End Type

Public Type Games
    Name As String
    IP As String
    Enabled As Boolean
    Host As String 'Who is hosting
End Type

Public Type GameTile
    Left As Variant
    Top As Variant
    Width As Integer
    Height As Integer
    Num As Integer 'Picture of the tile
    Visible As Boolean
    Screen As Integer
    Link As String 'What level to load next
End Type

Public strMaptoLoad As String 'Next map to load


Public Users(0 To 20) As ChatUser 'Users on the Server



Public curTime As Integer 'Countdown timer for the battle

Public theDamage(1 To 2) As Integer 'Damage done by each player

Public chatLoaded As Boolean 'Determines if the chat has already been loaded

Public sCurPsy As String 'Current Psynergy choosen
Public iCurPsy As Integer
Public strPsyName(1 To 50) As String 'Psynergy name
Public strPsyDamage(1 To 50) As String 'Psynergy damage
Public strPsyType(1 To 50) As String 'Psynergy elemental type
Public strPsyPP(1 To 50) As String 'PP required to use Psynergy
Public strPsyDjinn(1 To 50) As String 'Djinn required to use Psynergy
Public strPsyDesc(1 To 50) As String 'Psynergy Description

Public sCurSum As String 'Current Summon (string)
Public iCurSum As Integer 'Current Summon Level
Public strSumName(1 To 50) As String 'Summon name
Public strSumDesc(1 To 50) As String 'Summon description
Public strSumDjinn(1 To 50) As String 'Djinn required to use Summon

Public strCoins As String 'Player's Coins
Public strWins As String
Public strLoss As String
Public strDisc As String
Public strRating As String
Public strDjinn As String
Public strLvl As String
Public strChar As String 'Player's character
Public strWeapon As String 'Weapon name
Public intWeapon As Integer 'Weapon number for use in items.ini file
Public strMyPassWord As String 'Password entered

Public Version As String 'Current version of the game

Public disconnect As Boolean 'Did the player disconnect?
Public strOpponent As String 'Opponent's name
Public stroLvl As String 'Opponent's Level (string)
Public intoLvl As Integer 'Opponent's Level (integer)
Public stroChar As String 'Opponent's character
Public intoChar As Integer 'Opponent's character number
Public stroType As String 'Opponent's elemental type (string)
Public intoType As Integer 'Opponent's elemental type (int)
Public stroDefense As String 'Opponent's defense (string)
Public intoDefense As Integer 'Opponent's defense (int)
Public stroAP As String 'Opponent's AP (not used any more)
Public intoAP As Integer
Public stroHP As String 'Opponent's HP (not used any more)
Public intoHP As Integer
Public stroDamage As String 'Opponent's damage (not used any more)
Public intoDamage As Integer

Public currentDir As Integer 'Which way the attack/psynergy/djinn/summon goes
Public curDamage As Integer 'Current damage (not used any more?)
Public CurrentOp As Integer 'Opponent's picture? (not used any more?)
Public CurrentSum As Integer 'Current picture of the summon
Public Char(1 To 2) As Integer 'Character of either player
Public PP(1 To 2) As Integer 'Current PP of each player
Public Defense(1 To 2) As Integer 'Current defense of each player
Public AP(1 To 2) As Integer 'Current attack points of e/ player
Public HP(1 To 2) As Integer 'Current health of e/ player
Public PsyBonus(1 To 2) As Variant 'Additional damage to multiply Psynergy by
Public CharType(1 To 2) As String 'Elemental type of each player

Public bOReady(1 To 2) As Boolean 'Determines if the player is already ready
Public bWaitAttack(1 To 2) As Boolean 'Player waiting to... attack
Public bWaitPsynergy(1 To 2) As Boolean '... use Attacking Psynergy
Public bWaitHeal(1 To 2) As Boolean '... use Healing Psynergy
Public bWaitDropAttack(1 To 2) As Boolean '...use Attack dropping Psynergy
Public bWaitDropDefense(1 To 2) As Boolean '...use Defense dropping Psynergy
Public bWaitBoostAttack(1 To 2) As Boolean '...use Attack boosting Psynergy
Public bWaitBoostDefense(1 To 2) As Boolean '...use Defense boosting Psynergy
Public bWaitBoostPP(1 To 2) As Boolean '...use PP boosting Psynergy
Public bWaitPosion(1 To 2) As Boolean '...poison opponent (not currently implemented)
Public bWaitDefend(1 To 2) As Boolean '...Defend
Public bWaitDjinnAttack(1 To 2) As Boolean '...attack w/ a Djinn
Public bWaitDjinnHeal(1 To 2) As Boolean '...Heal w/ a Djinn
Public bWaitDjinnPP(1 To 2) As Boolean '...Boost PP w/ a Djinn
Public bWaitDjinnDropAttack(1 To 2) As Boolean '...Drop AP w/ a Djinn
Public bWaitDjinnDropDefense(1 To 2) As Boolean '...Drop Defense w/ a Djinn
Public bWaitDjinnDefense(1 To 2) As Boolean '...Boost Defense w/ a Djinn
Public bWaitDjinnBoostAttack(1 To 2) As Boolean '...Boost Attack w/ a Djinn
Public bWaitDjinnSet(1 To 2) As Boolean '...set a Djinn

Public bWaitSummon(1 To 2) As Boolean '...unleash a summon
Public iSummonType(1 To 2) As Integer 'The elemental type of a summon
Public iSummonLevel(1 To 2) As Integer 'The level of the summon

Public opFinished As Boolean 'Is the opponent finished the race already?
Public winRace As Boolean 'Did I win the race? Used for stat boosting after battle (not currently implemented)

Public DidIWin As Boolean 'Did I win the battle?
Public myRating As Integer 'Current rating
Public opRating As Integer 'Opponent's rating

Public PsyFrame As Integer 'Frame of the Psynergy animation

Public WinBattle As Boolean 'Is the battle over?

Public bWaitDjinn(1 To 2) As Boolean 'Not used?

Public CoinsGained As Integer 'Coins... gained after battle
Public RatingGained As Integer 'Rating...
Public LvlGained As Integer 'Level...

Public LoggedIn As Boolean 'Currently logged in to the server

Public strJoinIP As String 'IP to auto enter after hitting join game

'Resist Variables:
Dim intPower(1 To 2) As Integer
Dim intResist(1 To 2) As Integer
Dim intEarthPower(1 To 2) As Integer
Dim intEarthResist(1 To 2) As Integer
Dim intFirePower(1 To 2) As Integer
Dim intFireResist(1 To 2) As Integer
Dim intWaterPower(1 To 2) As Integer
Dim intWaterResist(1 To 2) As Integer
Dim intWindPower(1 To 2) As Integer
Dim intWindResist(1 To 2) As Integer
Dim intHeartPower(1 To 2) As Integer
Dim intHeartResist(1 To 2) As Integer
Dim intDarkPower(1 To 2) As Integer
Dim intDarkResist(1 To 2) As Integer


'Bitblt color constants
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const SRCERASE = &H4400328
Public Const WHITENESS = &HFF0062
Public Const BLACKNESS = &H42

Public IKILLKENNYIP As String 'Server IP
Public curCount As Integer 'Integer for count down clock?
Public notgo As Boolean 'Not used?
Public LostPassword As Boolean 'Not used?
Public NewUser As Boolean 'Making a new user

Public Type En 'Tile picture
 Pic As Long
 XPic As Integer
 YPic As Integer
 XCord As Long
 YCord As Long
 Height As Long
 Width As Long
 hit As Boolean
End Type

Public DispText(1 To 20) As Integer 'Time remaining to display in-town chat

Public bsNewChar As Boolean 'I forget!
Public bNewChar As Boolean 'Am I making a new character?

Public hoston As Boolean 'Am I the host of a battle?

Public Char1Step As Integer 'Not used?


'Midi Stuff
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

'Wave Stuff
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Ini Functions
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Function GetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
    Dim strReturn As String
    If (strSectionHeader & strVariableName) = strLastGetFromIni Then
        GetFromIni = strLastGetFromIni2
        Exit Function
    Else
        strLastGetFromIni = strSectionHeader & strVariableName
    End If
    strReturn = String(255, Chr(0))
    strLastGetFromIni2 = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
    GetFromIni = strLastGetFromIni2
End Function
Function WriteIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    'WritePrivateProfileString
    WriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function
Sub PlySound(strSound As String)
On Error Resume Next
'Play a sound
Call sndPlaySound(App.Path & "\" & strSound & ".wav", 1)
End Sub

Sub Player1Command()
'Determine what attack/psynergy/djinn/summon that the host is using
'Change the stats and enable animation timers

On Error Resume Next

If bOReady(1) = True And bOReady(2) = True Then 'Make sure that both players are ready
    If hoston = True Then 'I am the host
        
        If bWaitAttack(1) = True Then 'Waiting to attack
        
            CurrentOp = 2 'Sword drops on 2nd player
            frmBattle.timeSword.Enabled = True
            
            HP(2) = HP(2) - theDamage(1) 'Decrease HP
            frmBattle.lblText.Caption = "You did " & theDamage(1) & " damage."
        End If
        
        If bWaitPsynergy(1) = True Then 'Attacking Psynergy
            HP(2) = HP(2) - theDamage(1)
            frmBattle.lblText.Caption = "Your Psynergy did " & theDamage(1) & " damage."
            
            CurrentOp = 2 'Zap Player 2
            
            frmBattle.timePsynergy.Enabled = True
    
        End If
        
        If bWaitHeal(1) = True Then 'Waiting to heal
            HP(1) = HP(1) + theDamage(1)
            frmBattle.lblText.Caption = "You healed " & theDamage(1) & " HP."
    
        End If
        
        If bWaitBoostAttack(1) = True Then
            AP(1) = AP(1) + theDamage(1)
            frmBattle.lblText.Caption = "You increased your AP by " & theDamage(1) & "."
    
        End If
        
        If bWaitBoostDefense(1) = True Then
            Defense(1) = Defense(1) + theDamage(1)
            frmBattle.lblText.Caption = "You increased your Defense by " & theDamage(1) & "."
    
        End If
        
        If bWaitDropAttack(1) = True Then
            AP(2) = AP(2) - theDamage(1)
            frmBattle.lblText.Caption = "You decreased your opponent's AP by " & theDamage(1) & "."
        
        End If
        
        If bWaitDropDefense(1) = True Then
            Defense(2) = Defense(2) - theDamage(1)
            frmBattle.lblText.Caption = "You decrased your opponent's defense by " & theDamage(1) & "."
    
        End If
        
        If bWaitDefend(1) = True Then
            frmBattle.lblText.Caption = "You Defended."
            'This needs to actually implement something
        End If
        
        If bWaitDjinnAttack(1) = True Then
            HP(2) = HP(2) - theDamage(1)
            frmBattle.lblText.Caption = "Your Djinn did " & theDamage(1) & " damage."
            frmBattle.timeDjinn.Enabled = True 'Enabled Djinn animation
            CurrentOp = 2 'Attack Player 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture 'Set Djinn picture to current elemental type
        End If
        
        If bWaitDjinnDropAttack(1) = True Then
            AP(2) = AP(2) - theDamage(1)
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.lblText.Caption = "Your Djinn dropped your opponent's AP by " & theDamage(1) & "."
    
        End If
        
        If bWaitDropDefense(1) = True Then
            Defense(2) = Defense(2) - theDamage(1)
            frmBattle.lblText.Caption = "Your Djinn dropped your opponent's Defense by " & theDamage(1) & "."
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnBoostAttack(1) = True Then
            AP(1) = AP(1) + theDamage(1)
            frmBattle.lblText.Caption = "Your Djinn boosted your AP by " & theDamage(1) & "."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1 'Djinn affects me
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnDefense(1) = True Then
            Defense(1) = Defense(1) + theDamage(1)
            frmBattle.lblText.Caption = "Your Djinn increased your Defense by " & theDamage(1) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnHeal(1) = True Then
            HP(1) = HP(1) + theDamage(1)
            frmBattle.lblText.Caption = "Your Djinn healed your HP by " & theDamage(1) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitDjinnSet(1) = True Then
            frmBattle.lblText.Caption = "Your Djinn was set."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitSummon(1) = True Then
            HP(2) = HP(2) - theDamage(1)
            frmBattle.timeSummon.Enabled = True
            frmBattle.lblText.Caption = "Your Summon did " & theDamage(1) & " damage."
    
            frmBattle.imgSummon(0).Picture = frmBattle.imgSummon(2 * iSummonType(1)).Picture 'Set summon picture to current level of summon
            CurrentOp = 2
        
        End If
    
    
    Else 'I am not the host
    
    
        If bWaitAttack(2) = True Then 'Waiting to attack
        
            CurrentOp = 1 'Sword drops on 1st player
            frmBattle.timeSword.Enabled = True
            
            HP(1) = HP(1) - theDamage(2) 'Decrease HP
            frmBattle.lblText.Caption = "You took " & theDamage(1) & " damage."
        End If
        
        If bWaitPsynergy(2) = True Then 'Attacking Psynergy
            HP(1) = HP(1) - theDamage(2)
            frmBattle.lblText.Caption = "You took " & theDamage(1) & " damage from your opponent's Psynergy"
            
            CurrentOp = 1 'Zap Player 1
            
            frmBattle.timePsynergy.Enabled = True
    
        End If
        
        If bWaitHeal(2) = True Then 'Waiting to heal
            HP(2) = HP(2) + theDamage(2)
            frmBattle.lblText.Caption = "Your opponent healed " & theDamage(1) & " HP."
    
        End If
        
        If bWaitBoostAttack(2) = True Then
            AP(2) = AP(2) + theDamage(2)
            frmBattle.lblText.Caption = "Your opponent increased his/her AP by " & theDamage(1) & "."
    
        End If
        
        If bWaitBoostDefense(2) = True Then
            Defense(2) = Defense(2) + theDamage(2)
            frmBattle.lblText.Caption = "Your opponent increased his/her Defense by " & theDamage(1) & "."
    
        End If
        
        If bWaitDropAttack(2) = True Then
            AP(1) = AP(1) - theDamage(2)
            frmBattle.lblText.Caption = "Your AP was decreased by " & theDamage(1) & "."
        
        End If
        
        If bWaitDropDefense(2) = True Then
            Defense(1) = Defense(1) - theDamage(2)
            frmBattle.lblText.Caption = "Your defense was decreased by " & theDamage(1) & "."
    
        End If
        
        If bWaitDefend(2) = True Then
            frmBattle.lblText.Caption = "Your opponent Defended."
            'This needs to actually implement something
        End If
        
        If bWaitDjinnAttack(2) = True Then
            HP(1) = HP(1) - theDamage(2)
            frmBattle.lblText.Caption = "Your opponent's Djinn did " & theDamage(1) & " damage."
            frmBattle.timeDjinn.Enabled = True 'Enabled Djinn animation
            CurrentOp = 1 'Attack Player 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture 'Set Djinn picture to current elemental type
        End If
        
        If bWaitDjinnDropAttack(2) = True Then
            AP(1) = AP(1) - theDamage(2)
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.lblText.Caption = "Your opponent's Djinn dropped your AP by " & theDamage(1) & "."
    
        End If
        
        If bWaitDropDefense(2) = True Then
            Defense(1) = Defense(1) - theDamage(2)
            frmBattle.lblText.Caption = "Your opponent's Djinn dropped your Defense by " & theDamage(1) & "."
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnBoostAttack(2) = True Then
            AP(2) = AP(2) + theDamage(2)
            frmBattle.lblText.Caption = "Your opponent's Djinn boosted your opponent's AP by " & theDamage(1) & "."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2 'Djinn affects me
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnDefense(2) = True Then
            Defense(2) = Defense(2) + theDamage(2)
            frmBattle.lblText.Caption = "Your opponent's Djinn increased your opponent's Defense by " & theDamage(1) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnHeal(2) = True Then
            HP(2) = HP(2) + theDamage(2)
            frmBattle.lblText.Caption = "Your opponent's Djinn healed your opponent's HP by " & theDamage(1) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitDjinnSet(2) = True Then
            frmBattle.lblText.Caption = "Your opponent's Djinn was set."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitSummon(2) = True Then
            HP(1) = HP(1) - theDamage(2)
            frmBattle.timeSummon.Enabled = True
            frmBattle.lblText.Caption = "Your opponent's Summon did " & theDamage(1) & " damage."
    
            frmBattle.imgSummon(0).Picture = frmBattle.imgSummon(2 * iSummonType(1) - 1).Picture 'Set summon picture to current level of summon
            CurrentOp = 1
        
        End If
    
    End If 'If hoston=true

End If 'If boready(1) = true...

If hoston = True Then
'Reset all of my variables
    bOReady(1) = False
    bWaitAttack(1) = False
    bWaitPsynergy(1) = False
    bWaitSummon(1) = False
    bWaitDjinn(1) = False
    bWaitHeal(1) = False
    bWaitDropAttack(1) = False
    bWaitDropDefense(1) = False
    bWaitBoostAttack(1) = False
    bWaitBoostDefense(1) = False
    bWaitBoostPP(1) = False
    bWaitPosion(1) = False
    bWaitDefend(1) = False
    bWaitDjinnAttack(1) = False
    bWaitDjinnHeal(1) = False
    bWaitDjinnPP(1) = False
    bWaitDjinnDropAttack(1) = False
    bWaitDjinnDropDefense(1) = False
    bWaitDjinnDefense(1) = False
    bWaitDjinnSet(1) = False
Else
'Reset all of my opponent's variables
    bOReady(2) = False
    bWaitAttack(2) = False
    bWaitPsynergy(2) = False
    bWaitSummon(2) = False
    bWaitDjinn(2) = False
    bWaitHeal(2) = False
    bWaitDropAttack(2) = False
    bWaitDropDefense(2) = False
    bWaitBoostAttack(2) = False
    bWaitBoostDefense(2) = False
    bWaitBoostPP(2) = False
    bWaitPosion(2) = False
    bWaitDefend(2) = False
    bWaitDjinnAttack(2) = False
    bWaitDjinnHeal(2) = False
    bWaitDjinnPP(2) = False
    bWaitDjinnDropAttack(2) = False
    bWaitDjinnDropDefense(2) = False
    bWaitDjinnDefense(2) = False
    bWaitDjinnSet(2) = False
End If

Call Player2Command


End Sub
Sub Player2Command()
On Error Resume Next
strdtime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
rsave = App.Path & "\userdata.ini"

On Error Resume Next

If (hoston = True And bOReady(2) = True) Or (hoston = False And bOReady(1) = True) Then
    If hoston = False Then 'I am the client
        
        If bWaitAttack(1) = True Then 'Waiting to attack
        
            CurrentOp = 2 'Sword drops on 2nd player
            frmBattle.timeSword.Enabled = True
            
            HP(2) = HP(2) - theDamage(1) 'Decrease HP
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "You did " & theDamage(1) & " damage."
        End If
        
        If bWaitPsynergy(1) = True Then 'Attacking Psynergy
            HP(2) = HP(2) - theDamage(1)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your Psynergy did " & theDamage(1) & " damage."
            
            CurrentOp = 2 'Zap Player 2
            
            frmBattle.timePsynergy.Enabled = True
    
        End If
        
        If bWaitHeal(1) = True Then 'Waiting to heal
            HP(1) = HP(1) + theDamage(1)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "You healed " & theDamage(1) & " HP."
    
        End If
        
        If bWaitBoostAttack(1) = True Then
            AP(1) = AP(1) + theDamage(1)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "You increased your AP by " & theDamage(1) & "."
    
        End If
        
        If bWaitBoostDefense(1) = True Then
            Defense(1) = Defense(1) + theDamage(1)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "You increased your Defense by " & theDamage(1) & "."
    
        End If
        
        If bWaitDropAttack(1) = True Then
            AP(2) = AP(2) - theDamage(1)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "You decreased your opponent's AP by " & theDamage(1) & "."
        
        End If
        
        If bWaitDropDefense(1) = True Then
            Defense(2) = Defense(2) - theDamage(1)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "You decrased your opponent's defense by " & theDamage(1) & "."
    
        End If
        
        If bWaitDefend(1) = True Then
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "You Defended."
            'This needs to actually implement something
        End If
        
        If bWaitDjinnAttack(1) = True Then
            HP(2) = HP(2) - theDamage(1)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your Djinn did " & theDamage(1) & " damage."
            frmBattle.timeDjinn.Enabled = True 'Enabled Djinn animation
            CurrentOp = 2 'Attack Player 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture 'Set Djinn picture to current elemental type
        End If
        
        If bWaitDjinnDropAttack(1) = True Then
            AP(2) = AP(2) - theDamage(1)
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your Djinn dropped your opponent's AP by " & theDamage(1) & "."
    
        End If
        
        If bWaitDropDefense(1) = True Then
            Defense(2) = Defense(2) - theDamage(1)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your Djinn dropped your opponent's Defense by " & theDamage(1) & "."
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnBoostAttack(1) = True Then
            AP(1) = AP(1) + theDamage(1)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your Djinn boosted your AP by " & theDamage(1) & "."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1 'Djinn affects me
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnDefense(1) = True Then
            Defense(1) = Defense(1) + theDamage(1)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your Djinn increased your Defense by " & theDamage(1) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnHeal(1) = True Then
            HP(1) = HP(1) + theDamage(1)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your Djinn healed your HP by " & theDamage(1) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitDjinnSet(1) = True Then
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your Djinn was set."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitSummon(1) = True Then
            HP(2) = HP(2) - theDamage(1)
            frmBattle.timeSummon.Enabled = True
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your Summon did " & theDamage(1) & " damage."
    
            frmBattle.imgSummon(0).Picture = frmBattle.imgSummon(2 * iSummonType(1)).Picture 'Set summon picture to current level of summon
            CurrentOp = 2
        
        End If
    
    
    Else 'I am the host
    
    
        If bWaitAttack(2) = True Then 'Waiting to attack
        
            CurrentOp = 1 'Sword drops on 1st player
            frmBattle.timeSword.Enabled = True
            
            HP(1) = HP(1) - theDamage(2) 'Decrease HP
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "You took " & theDamage(1) & " damage."
        End If
        
        If bWaitPsynergy(2) = True Then 'Attacking Psynergy
            HP(1) = HP(1) - theDamage(2)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "You took " & theDamage(1) & " damage from your opponent's Psynergy"
            
            CurrentOp = 1 'Zap Player 1
            
            frmBattle.timePsynergy.Enabled = True
    
        End If
        
        If bWaitHeal(2) = True Then 'Waiting to heal
            HP(2) = HP(2) + theDamage(2)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your opponent healed " & theDamage(1) & " HP."
    
        End If
        
        If bWaitBoostAttack(2) = True Then
            AP(2) = AP(2) + theDamage(2)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your opponent increased his/her AP by " & theDamage(1) & "."
    
        End If
        
        If bWaitBoostDefense(2) = True Then
            Defense(2) = Defense(2) + theDamage(2)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your opponent increased his/her Defense by " & theDamage(1) & "."
    
        End If
        
        If bWaitDropAttack(2) = True Then
            AP(1) = AP(1) - theDamage(2)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your AP was decreased by " & theDamage(1) & "."
        
        End If
        
        If bWaitDropDefense(2) = True Then
            Defense(1) = Defense(1) - theDamage(2)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your defense was decreased by " & theDamage(1) & "."
    
        End If
        
        If bWaitDefend(2) = True Then
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your opponent Defended."
            'This needs to actually implement something
        End If
        
        If bWaitDjinnAttack(2) = True Then
            HP(1) = HP(1) - theDamage(2)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your opponent's Djinn did " & theDamage(1) & " damage."
            frmBattle.timeDjinn.Enabled = True 'Enabled Djinn animation
            CurrentOp = 1 'Attack Player 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture 'Set Djinn picture to current elemental type
        End If
        
        If bWaitDjinnDropAttack(2) = True Then
            AP(1) = AP(1) - theDamage(2)
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your opponent's Djinn dropped your AP by " & theDamage(1) & "."
    
        End If
        
        If bWaitDropDefense(2) = True Then
            Defense(1) = Defense(1) - theDamage(2)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your opponent's Djinn dropped your Defense by " & theDamage(1) & "."
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnBoostAttack(2) = True Then
            AP(2) = AP(2) + theDamage(2)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your opponent's Djinn boosted your opponent's AP by " & theDamage(1) & "."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2 'Djinn affects me
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnDefense(2) = True Then
            Defense(2) = Defense(2) + theDamage(2)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your opponent's Djinn increased your opponent's Defense by " & theDamage(1) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnHeal(2) = True Then
            HP(2) = HP(2) + theDamage(2)
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your opponent's Djinn healed your opponent's HP by " & theDamage(1) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitDjinnSet(2) = True Then
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your opponent's Djinn was set."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitSummon(2) = True Then
            HP(1) = HP(1) - theDamage(2)
            frmBattle.timeSummon.Enabled = True
            frmBattle.lblText.Caption = frmBattle.lblText.Caption & vbNewLine & "Your opponent's Summon did " & theDamage(1) & " damage."
    
            frmBattle.imgSummon(0).Picture = frmBattle.imgSummon(2 * iSummonType(1) - 1).Picture 'Set summon picture to current level of summon
            CurrentOp = 1
        
        End If
    
    End If 'If hoston=true

End If 'If boready(1) = true...


If hoston = False Then
'Reset all of my variables
    bOReady(1) = False
    bWaitAttack(1) = False
    bWaitPsynergy(1) = False
    bWaitSummon(1) = False
    bWaitDjinn(1) = False
    bWaitHeal(1) = False
    bWaitDropAttack(1) = False
    bWaitDropDefense(1) = False
    bWaitBoostAttack(1) = False
    bWaitBoostDefense(1) = False
    bWaitBoostPP(1) = False
    bWaitPosion(1) = False
    bWaitDefend(1) = False
    bWaitDjinnAttack(1) = False
    bWaitDjinnHeal(1) = False
    bWaitDjinnPP(1) = False
    bWaitDjinnDropAttack(1) = False
    bWaitDjinnDropDefense(1) = False
    bWaitDjinnDefense(1) = False
    bWaitDjinnSet(1) = False
Else
'Reset all of my opponent's variables
    bOReady(2) = False
    bWaitAttack(2) = False
    bWaitPsynergy(2) = False
    bWaitSummon(2) = False
    bWaitDjinn(2) = False
    bWaitHeal(2) = False
    bWaitDropAttack(2) = False
    bWaitDropDefense(2) = False
    bWaitBoostAttack(2) = False
    bWaitBoostDefense(2) = False
    bWaitBoostPP(2) = False
    bWaitPosion(2) = False
    bWaitDefend(2) = False
    bWaitDjinnAttack(2) = False
    bWaitDjinnHeal(2) = False
    bWaitDjinnPP(2) = False
    bWaitDjinnDropAttack(2) = False
    bWaitDjinnDropDefense(2) = False
    bWaitDjinnDefense(2) = False
    bWaitDjinnSet(2) = False
End If

'Reset the reset variables :)
Reset(1) = False
Reset(2) = False

frmBattle.cmdReset.Enabled = True 'Re-enable the reset button
curCount = 30 'Reset time
frmBattle.timecount.Enabled = True 'Start counting again

Call EnableChoose

End Sub
Function AutoScroll(txtbox As RichTextBox)
On Error Resume Next
If txtbox.MultiLine = False Then Exit Function
On Error Resume Next
txtbox.SelLength = 0

    If Len(Trim(txtbox.Text)) > 0 Then
        If Right$(txtbox.Text, 1) = vbCrLf Then
            txtbox.SelStart = Len(txtbox.Text) - 1
            Exit Function
        End If
        txtbox.SelStart = Len(txtbox.Text)
    End If
End Function
Sub UserLists()
On Error Resume Next
Dim yadda As Integer
Dim i As Integer
If frmChat.lstUsers.ListCount <> frmChat.lstCheck.ListCount Then
yadda = frmChat.lstUsers.ListCount - 1
For i = 0 To frmChat.lstUsers.ListCount - 1
frmChat.lstUsers.RemoveItem (yadda - i)
Next 'i
yadda = frmChat.lstCheck.ListCount - 1
For i = 0 To frmChat.lstCheck.ListCount - 1
frmChat.lstUsers.AddItem frmChat.lstCheck.List(yadda - i)
Next 'i
yadda = frmChat.lstCheck.ListCount - 1
For i = 0 To frmChat.lstCheck.ListCount - 1
frmChat.lstCheck.RemoveItem (yadda - i)
Next 'i
End If
End Sub
Public Function ConvertNum(ByVal strNumber As String) As String
On Error Resume Next

If strNumber = "9" Then strNumber = "a"
If strNumber = "8" Then strNumber = "b"
If strNumber = "7" Then strNumber = "c"
If strNumber = "6" Then strNumber = "d"
If strNumber = "5" Then strNumber = "e"
If strNumber = "4" Then strNumber = "f"
If strNumber = "3" Then strNumber = "g"
If strNumber = "2" Then strNumber = "h"
If strNumber = "1" Then strNumber = "i"
If strNumber = "0" Then strNumber = "j"
ConvertNum = strNumber
End Function
Public Function ConvertAlpha(ByVal strNumber1 As String) As String
On Error Resume Next
Dim finalNum(1 To 4) As String

For i = 1 To 4

Dim strNumber As String
strNumber = Mid(strNumber1, i, 1)

If strNumber = "a" Then strNumber = "9"
If strNumber = "b" Then strNumber = "8"
If strNumber = "c" Then strNumber = "7"
If strNumber = "d" Then strNumber = "6"
If strNumber = "e" Then strNumber = "5"
If strNumber = "f" Then strNumber = "4"
If strNumber = "g" Then strNumber = "3"
If strNumber = "h" Then strNumber = "2"
If strNumber = "i" Then strNumber = "1"
If strNumber = "j" Then strNumber = "0"
finalNum(i) = strNumber
Next 'i

ConvertAlpha = finalNum(1) & finalNum(2) & finalNum(3) & finalNum(4)
End Function
Public Sub EnableChoose()
On Error Resume Next
'Re-enables menu options

frmBattle.timeHide.Enabled = True 'A delay before the text is hidden
frmBattle.lblAttack.Visible = True
frmBattle.lblPsynergy.Visible = True
frmBattle.lblDjinn.Visible = True
frmBattle.lblSummon.Visible = True
frmBattle.lblDefend.Visible = True
frmBattle.shpMenu.Visible = True
frmBattle.lblgen(9).Caption = HP(1)
frmBattle.lblgen(5).Caption = HP(2)
frmBattle.shpHP(0).Width = HP(1) / 2
frmBattle.shpHP(1).Width = HP(2) / 2
frmBattle.lblgen(10).Caption = PP(1)

End Sub
Public Sub SetPowerResist(iLvl As Integer, intPlayerNumber As Integer)
On Error Resume Next
' E=Earth, F=Fire, N=Wind, W=Water, H=Heart, D=Dark
Select Case CharType(intplayer)
Case "E"
intPower(intplayer) = 80 + iLvl * 5
intResist(intplayer) = 90 + iLvl * 5
intEarthPower(intplayer) = 90 + iLvl * 5
intEarthResist(intplayer) = 100 + iLvl * 5
intFirePower(intplayer) = 80 + iLvl * 5
intFireResist(intplayer) = 90 + iLvl * 5
intWaterPower(intplayer) = 80 + iLvl * 5
intWaterResist(intplayer) = 90 + iLvl * 5
intWindPower(intplayer) = 70 + iLvl * 5
intWindResist(intplayer) = 80 + iLvl * 5
intHeartPower(intplayer) = 80 + iLvl * 5
intHeartResist(intplayer) = 90 + iLvl * 5
intDarkPower(intplayer) = 80 + iLvl * 5
intDarkResist(intplayer) = 90 + iLvl * 5
Case "F"
intPower(intplayer) = 80 + iLvl * 5
intResist(intplayer) = 90 + iLvl * 5
intEarthPower(intplayer) = 80 + iLvl * 5
intEarthResist(intplayer) = 90 + iLvl * 5
intFirePower(intplayer) = 90 + iLvl * 5
intFireResist(intplayer) = 100 + iLvl * 5
intWaterPower(intplayer) = 70 + iLvl * 5
intWaterResist(intplayer) = 80 + iLvl * 5
intWindPower(intplayer) = 80 + iLvl * 5
intWindResist(intplayer) = 90 + iLvl * 5
intHeartPower(intplayer) = 80 + iLvl * 5
intHeartResist(intplayer) = 90 + iLvl * 5
intDarkPower(intplayer) = 80 + iLvl * 5
intDarkResist(intplayer) = 90 + iLvl * 5
Case "N"
intPower(intplayer) = 80 + iLvl * 5
intResist(intplayer) = 90 + iLvl * 5
intEarthPower(intplayer) = 70 + iLvl * 5
intEarthResist(intplayer) = 80 + iLvl * 5
intFirePower(intplayer) = 80 + iLvl * 5
intFireResist(intplayer) = 90 + iLvl * 5
intWaterPower(intplayer) = 80 + iLvl * 5
intWaterResist(intplayer) = 90 + iLvl * 5
intWindPower(intplayer) = 90 + iLvl * 5
intWindResist(intplayer) = 100 + iLvl * 5
intHeartPower(intplayer) = 80 + iLvl * 5
intHeartResist(intplayer) = 90 + iLvl * 5
intDarkPower(intplayer) = 80 + iLvl * 5
intDarkResist(intplayer) = 90 + iLvl * 5
Case "W"
intPower(intplayer) = 80 + iLvl * 5
intResist(intplayer) = 90 + iLvl * 5
intEarthPower(intplayer) = 80 + iLvl * 5
intEarthResist(intplayer) = 90 + iLvl * 5
intFirePower(intplayer) = 70 + iLvl * 5
intFireResist(intplayer) = 80 + iLvl * 5
intWaterPower(intplayer) = 90 + iLvl * 5
intWaterResist(intplayer) = 100 + iLvl * 5
intWindPower(intplayer) = 80 + iLvl * 5
intWindResist(intplayer) = 90 + iLvl * 5
intHeartPower(intplayer) = 80 + iLvl * 5
intHeartResist(intplayer) = 90 + iLvl * 5
intDarkPower(intplayer) = 80 + iLvl * 5
intDarkResist(intplayer) = 90 + iLvl * 5
Case "H"
intPower(intplayer) = 80 + iLvl * 5
intResist(intplayer) = 90 + iLvl * 5
intEarthPower(intplayer) = 80 + iLvl * 5
intEarthResist(intplayer) = 90 + iLvl * 5
intFirePower(intplayer) = 80 + iLvl * 5
intFireResist(intplayer) = 90 + iLvl * 5
intWaterPower(intplayer) = 80 + iLvl * 5
intWaterResist(intplayer) = 90 + iLvl * 5
intWindPower(intplayer) = 80 + iLvl * 5
intWindResist(intplayer) = 90 + iLvl * 5
intHeartPower(intplayer) = 90 + iLvl * 5
intHeartResist(intplayer) = 100 + iLvl * 5
intDarkPower(intplayer) = 70 + iLvl * 5
intDarkResist(intplayer) = 80 + iLvl * 5
Case "D"
intPower(intplayer) = 80 + iLvl * 5
intResist(intplayer) = 90 + iLvl * 5
intEarthPower(intplayer) = 80 + iLvl * 5
intEarthResist(intplayer) = 90 + iLvl * 5
intFirePower(intplayer) = 80 + iLvl * 5
intFireResist(intplayer) = 90 + iLvl * 5
intWaterPower(intplayer) = 80 + iLvl * 5
intWaterResist(intplayer) = 90 + iLvl * 5
intWindPower(intplayer) = 80 + iLvl * 5
intWindResist(intplayer) = 90 + iLvl * 5
intHeartPower(intplayer) = 70 + iLvl * 5
intHeartResist(intplayer) = 80 + iLvl * 5
intDarkPower(intplayer) = 90 + iLvl * 5
intDarkResist(intplayer) = 100 + iLvl * 5
End Select
End Sub



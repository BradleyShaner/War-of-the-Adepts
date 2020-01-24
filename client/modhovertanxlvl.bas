Attribute VB_Name = "modHoverLVL"
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long

Public Type SumType
    Name As String
    Level As Long
    Element As String
    Desc As String
    Enabled  As Boolean
    Character As Long
End Type
Public Type PsyType
    Name As String
    Damage As Long
    Desc As String
    Element As String
    Type As String
    PP As Long
    Enabled As Boolean
    Djinn As Long
    Character As Long
End Type
Public Type DjinnType
    Damage As Long
    Desc As String
    Element As String
    Type As String
    State As Long
    Character As Long 'Which character it belongs to
    Enabled As Boolean
    Name As String
End Type




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

Public strMyUserName As String


Public bWaitToResetDjinn As Boolean

Public bFlashChat As Boolean

Public strDjinnSavePlayer As String

Public strMidi As String

Public strPINNum As String 'PIN number to identify computer

Public strRealIP As String 'Your real IP address

Public bMazeError As Boolean 'If the maze has erred before

'Stuff for uploading mazes:
Public strUploadFile As String
Public strUploadFileNoPath As String
Public bSendMaze As Boolean 'Are you sending a maze or just downloading data?

Public intEasterEggs As Integer 'Easter eggs found
Public RealPsy As Integer
Public MazeWait As Boolean

Public bAutoScroll As Boolean
Public bLogChat As Boolean

Public intKarma As Long 'Determines how many messages you can send, what features you get, etc.

Public DidNotRoll As Boolean

Public bMazeFirstLoad As Boolean

Public strModName(1 To 15) As String

Public strCustCharName(1 To 20) As String
Public strCustCharDesc(1 To 20) As String
Public intCustCharHP(1 To 20) As Long
Public intCustCharAP(1 To 20) As Long
Public intCustCharDefense(1 To 20) As Long
Public intCustCharLuck(1 To 20) As Variant
Public strCustCharType(1 To 20) As String
Public intCustCharResist(1 To 20) As Long
Public intCustCharPower(1 To 20) As Long

'Where the opponent is in the maze:
Public intOpMazeX As Long
Public intOpMazeY As Long

Dim strOpponentIP As String

Public intFirstAttack As Long 'Who attacks first
Public curCustChar As Long

Public strItemName(1 To 30) As String 'Item Name
Public strItemDesc(1 To 30) As String 'Item Description
Public intItemAddMod(1 To 30) As Integer 'Item's add modifier for weapon special (examples 5, 8, 7
Public varItemMultMod(1 To 30) As Variant 'Item's Mult modifier for weapon special (examples 2.5 [250%], .5 [50 percent]
Public intItemSpcPercent(1 To 30) As Integer ' Item's Chance of unleaseing Special Attack
Public strItemCoins(1 To 30) As String 'Item Cost
Public strItemType(1 To 30) As String 'Item Elemental Type
Public strItemDamage(1 To 30) As String 'Item Damage
Public strItemSpcType(1 To 30) As String 'Item Critical Hit Type
Public strItemSpcDamage(1 To 30) As String 'Item Critical Hit Damage
Public strItemSpcDesc(1 To 30) As String 'Item Critical Hit Description

'Public ServerNumber As Integer 'Number for use in the server

Dim strdtime As String 'Current time

Dim rsave As String 'Userdata.ini save directory

Public curIsaac As Integer 'Current player in Online Town

Public IsaacM(1 To 20) As GameTile '1-20 Players in Online Town




Dim iCurTurn As Integer '???? I don't think this is used anymore, it will be removed once I figure out what I used it for

Public AmIKilled As Boolean 'Am I dead?

Public Type CustomCharacter
    Name As String
    BaseHP As Long
    BaseAP As Long
    BaseDefense As Long
    BasePP As Long
    BasePower As Long
    BaseRes As Long
    Strength As String
    Weakness As String
    Picture As String
    BaseLuck As Long
    Type As String
    Users As String
    Description As String
    BaseSpeed As Long
End Type


Public Psynergy(1 To 100) As PsyType

Public Djinn(1 To 50) As DjinnType



Public Summon(1 To 10) As SumType

Public CustomChar(1 To 50) As CustomCharacter

Public CurEgg26 As Long

Public strDjinnName(1 To 20) As String 'Djinn Name
Public strDjinnDesc(1 To 20) As String 'Djinn Description
Public strDjinnType(1 To 20) As String 'Djinn Elemental Type
Public strDjinnDamage(1 To 20) As String 'Djinn Damage
Public bDjinnSet(1 To 20) As Long 'Is the Djinn At Rest, Set or Standby?
Public iCurDjinn As Integer 'Current djinn number
Public sCurDjinn As String '
Public intDjinnStandby(1 To 2) As Integer 'Total Djinn on Standby
Public sTotalDjinn(1 To 2) As String 'Total Djinn
Public iTotalDjinn(1 To 2) As Integer 'Total Djinn


Public Reset(1 To 2) As Boolean 'Did the player call for a reset of battle?


Public strOpDjinnType As String 'The picture of the opponent's Djinn being used

Public Type ChatUser
    Name As String
    Enabled As Boolean 'Visible or not
    Pic As String 'Character picture
    Left As String 'X position
    Top As String 'Y position
    Screen As String 'Current screen
    CustomChar As Boolean 'Is it a custom character?
    Avatar As String
    Away As Boolean
    Moderator As Boolean
    Admin As Boolean
    Ignore As Boolean
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
    IconType As Long 'The icon above their head
    IconTime As Long 'Icon visible for how much longer?
    CustomCharacter As Boolean
End Type

Public Type Scrambler
    Name As String
    Score As Long
End Type

Public strMaptoLoad As String 'Next map to load


Public Users(0 To 20) As ChatUser 'Users on the Server

Public FirstLogon As Boolean


Public curTime As Integer 'Countdown timer for the battle

Public thedamage(1 To 2) As Integer 'Damage done by each player

Public chatLoaded As Boolean 'Determines if the chat has already been loaded

Public sCurPsy As String 'Current Psynergy choosen
Public iCurPsy As Integer
Public strPsyName(1 To 60) As String 'Psynergy name
Public strPsyDamage(1 To 60) As String 'Psynergy damage
Public strPsyType(1 To 60) As String 'Psynergy elemental type
Public strPsyPP(1 To 60) As String 'PP required to use Psynergy
Public strPsyDjinn(1 To 60) As String 'Djinn required to use Psynergy
Public strPsyDesc(1 To 60) As String 'Psynergy Description
Public strCurPsyDir As String 'Current directory for psynergy animation



Public sCurSum As String 'Current Summon (string)
Public iCurSum As Integer 'Current Summon Level
Public strSumName(1 To 10) As String 'Summon name
Public strSumDesc(1 To 10) As String 'Summon description
Public strSumDjinn(1 To 10) As String 'Djinn required to use Summon
Public intSumBoost As Integer 'Power boost from Summon

'If you can use the following in battle:
Public bAllowSummon As Boolean 'Allow summons in battle
Public bAllowHeal As Boolean
Public bAllowAttack As Boolean
Public bAllowPsynergy As Boolean
Public bEqualizeWeapons As Boolean

Public strCoins As String 'Player's Coins
Public strWins As String
Public strLoss As String
Public strDisc As String
Public strRating As String
Public strDjinn As String
Public strLvl As String
Public strChar(1 To 2) As String 'Player's character
Public strWeapon(1 To 2) As String 'Weapon name
Public intWeapon(1 To 2) As Integer 'Weapon number for use in items.ini file

Public strMyPassWord As String 'Password entered

Public Version As String 'Current version of the game

Public disconnect As Boolean 'Did the player disconnect?

Public strOpponent As String 'Opponent's name

'The following is no longer used:
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
'No longer used (above)

Public AttackType(1 To 4) As String 'What type of attack you are doing
Public Target(1 To 4) As Long 'What player is being targeted
Public AttackDamage(1 To 4) As Long 'Damage done
Public intTurn As Long 'Which player is going
Public SelectTarget As Boolean 'Whether you're selecting your target or not


Public currentDir As Integer 'Which way the attack/psynergy/djinn/summon goes
Public curDamage As Integer 'Current damage (not used any more?)
Public CurrentOp As Integer 'Opponent's picture? (not used any more?)
Public CurrentSum As Integer 'Current picture of the summon
Public Char(1 To 4) As Integer 'Character number of either player
Public CharName(1 To 4) As String 'Character name of each character
Public PP(1 To 4) As Integer 'Current PP of each player
Public Defense(1 To 4) As Integer 'Current defense of each player
Public AP(1 To 4) As Integer 'Current attack points of e/ player
Public HP(1 To 4) As Integer 'Current health of e/ player
Public PsyBonus(1 To 4) As Variant 'Additional damage to multiply Psynergy by
Public CharType(1 To 4) As String 'Elemental type of each player
Public Level(1 To 4) As Long
Public Luck(1 To 4) As Integer
Public Speed(1 To 4) As Integer 'Speed of each character
Public MaxHP(1 To 4) As Integer 'Maximum HP for both of the characters
Public Handicap(1 To 2) As Integer 'Player's level added/subtracted
Public RelativeLVL(1 To 2) As Long 'HP +- Handicapp
Public RelativeRating(1 To 2) As Long 'Rating Points ""
Public RelativeDjinn(1 To 2) As Long 'Djinn ""
Public RelativeWeapon(1 To 2) As Long
Public DjinnElement(1 To 4) As String 'What type the djinn is


Public intEarthPower(1 To 4) As Integer
Public intEarthResist(1 To 4) As Integer
Public intFirePower(1 To 4) As Integer
Public intFireResist(1 To 4) As Integer
Public intWaterPower(1 To 4) As Integer
Public intWaterResist(1 To 4) As Integer
Public intWindPower(1 To 4) As Integer
Public intWindResist(1 To 4) As Integer
Public intHeartPower(1 To 4) As Integer
Public intHeartResist(1 To 4) As Integer
Public intDarkPower(1 To 4) As Integer
Public intDarkResist(1 To 4) As Integer

Public bOReady(1 To 4) As Boolean 'Determines if the player is already ready
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
Public bWaitCriticalAttack(1 To 2) As Boolean 'Critical attack?
Public bWaitSpecialAttack(1 To 2) As Boolean 'Special unleash?
Public bWaitBoostResist(1 To 2) As Boolean 'Boost my elemental resistance
Public bWaitDropResist(1 To 2) As Boolean 'Drop foe's resistance

Public bCustomChar(1 To 4) As Long

Public bWaitSummon(1 To 2) As Boolean '...unleash a summon
Public iSummonType(1 To 4) As Integer 'The elemental type of a summon
Public iSummonLevel(1 To 4) As Integer 'The level of the summon

Public opFinished As Boolean 'Is the opponent finished the race already?
Public winRace As Long 'Did I win the race? Used for stat boosting after battle (not currently implemented)

Public DidIWin As Boolean 'Did I win the battle?
Public myRating As Integer 'Current rating
Public opRating As Integer 'Opponent's rating

Public PsyFrame As Integer 'Frame of the Psynergy animation

Public WinBattle As Boolean 'Is the battle over?

Public bWaitDjinn(1 To 2) As Boolean 'Not used?

Public strServerDate As String

Public LoggedIn As Boolean 'Currently logged in to the server

Public strJoinIP As String 'IP to auto enter after hitting join game

Public IKILLKENNYIP As String 'Server IP
Public curCount As Integer 'Integer for count down clock?
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

Public strType As String 'My type

Public BattleLoaded(1 To 2) As Boolean 'Have each player loaded the battle screen?

Public intGlobalIcon(1 To 20) As Integer 'Temporary integer for the icon on top of someone's head

Public TimedMatch As Boolean 'is the match timed?
Public PlayerWait As Long 'Current pause time after each player has his turn.
Public GameOver As Boolean 'Determines if either player has won, game will no longer crash

Public DoubleStats As Boolean

Public DataSent As Boolean

Public intDjinnSaveHighScore As Long 'Current high score for Djinn Save mini-game


'Wave Stuff
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Ini Functions
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Function GetFromIni(strSectionHeader As String, strVariableName As String, strFilename As String) As String
On Error Resume Next
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    GetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFilename))
End Function
Function WriteIni(strSectionHeader As String, strVariableName As String, strValue As String, strFilename As String) As Integer
On Error Resume Next
    'WritePrivateProfileString
    WriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFilename)
End Function
Sub PlySound(strSound As String)
On Error Resume Next
'Play a sound
'strCheck = GetFromIni("GEN", "SOUND", App.Path & "\settings.ini")
'If strCheck <> "OFF" Then 'Is the sound on?
    Call sndPlaySound(App.Path & "\" & strSound & ".wav", 1)
'End If

End Sub

Sub Player1Command()
'Determine what attack/psynergy/djinn/summon that the host is using
'Change the stats and enable animation timers

On Error Resume Next

'If bOReady(1) = True Then 'Make sure that both players are ready
    If hoston = True Then 'I am the host
        
        
        
        
        If bWaitAttack(1) = True Then 'Waiting to attack
        
            CurrentOp = 2 'Sword drops on 2nd player
            frmBattle.timeSword.Enabled = True
            
            HP(2) = HP(2) - thedamage(1) 'Decrease HP
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You did " & thedamage(1) & " damage."
        End If
        If bWaitCriticalAttack(1) = True Then
            CurrentOp = 2
            frmBattle.timeSword.Enabled = True
            HP(2) = HP(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You did " & thedamage(1) & " damage.  Critical Hit!"
        End If
        If bWaitSpecialAttack(1) = True Then
            CurrentOp = 2
            frmBattle.timeSword.Enabled = True
            HP(2) = HP(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You let out a howl.  Special Attack!  You did " & thedamage(1) & " damage."
        End If
        
        If bWaitPsynergy(1) = True Then 'Attacking Psynergy
            HP(2) = HP(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Psynergy did " & thedamage(1) & " damage."
            
            CurrentOp = 2 'Zap Player 2
            
            strCurPsyDir = CharType(1)
            frmBattle.timePsynergy.Enabled = True
            
    
        End If
        
        If bWaitHeal(1) = True Then 'Waiting to heal

        'Healing Psynergy will always heal the same amount of damage
        thedamage(1) = CInt(strPsyDamage(RealPsy))
        
        If HP(1) + thedamage(1) > MaxHP(1) Then 'Can't heal more than max
            thedamage(1) = thedamage(1) - ((HP(1) + thedamage(1)) - MaxHP(1))
        End If
        
        If hoston = True Then
        frmHost.Host.SendData "DMG" & thedamage(1) & vbCrLf 'Increase HP by this much
        End If
        If hoston = False Then
        frmJoin.Client.SendData "DMG" & thedamage(1) & vbCrLf
        End If
        
        HP(1) = HP(1) + thedamage(1)
        frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You healed " & thedamage(1) & " HP."
    
        End If
        
        If bWaitBoostAttack(1) = True Then
            AP(1) = AP(1) + thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You increased your AP by " & thedamage(1) & "."
    
        End If
        
        If bWaitBoostDefense(1) = True Then
            Defense(1) = Defense(1) + thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You increased your Defense by " & thedamage(1) & "."
    
        End If
        
        If bWaitDropAttack(1) = True Then
            AP(2) = AP(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You decreased your opponent's AP by " & thedamage(1) & "."
        
        End If
        
        If bWaitDropDefense(1) = True Then
            Defense(2) = Defense(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You decrased your opponent's defense by " & thedamage(1) & "."
    
        End If
        
        If bWaitDefend(1) = True Then
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You Defended."
            'This needs to actually implement something
        End If
        
        If bWaitDjinnAttack(1) = True Then
            HP(2) = HP(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn did " & thedamage(1) & " damage."
            frmBattle.timeDjinn.Enabled = True 'Enabled Djinn animation
            CurrentOp = 2 'Attack Player 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture 'Set Djinn picture to current elemental type
        End If
        
        If bWaitDjinnDropAttack(1) = True Then
            AP(2) = AP(2) - thedamage(1)
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn dropped your opponent's AP by " & thedamage(1) & "."
    
        End If
        
        If bWaitDropDefense(1) = True Then
            Defense(2) = Defense(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn dropped your opponent's Defense by " & thedamage(1) & "."
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnBoostAttack(1) = True Then
            AP(1) = AP(1) + thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn boosted your AP by " & thedamage(1) & "."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1 'Djinn affects me
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnDefense(1) = True Then
            Defense(1) = Defense(1) + thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn increased your Defense by " & thedamage(1) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnHeal(1) = True Then
            HP(1) = HP(1) + thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn healed your HP by " & thedamage(1) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitDjinnSet(1) = True Then
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn was set."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitSummon(1) = True Then
            HP(2) = HP(2) - thedamage(1)
            frmBattle.timeSummon.Enabled = True
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Summon did " & thedamage(1) & " damage.  Your Power boosted by: " & intSumBoost
            
            frmBattle.imgSummon(0).Picture = frmBattle.imgSummon(2 * iSummonType(1)).Picture 'Set summon picture to current level of summon
            CurrentOp = 2
        
        End If
        If bWaitDropResist(1) = True Then 'Djinn only
            intFireResist(2) = intFireResist(2) - thedamage(1)
            intEarthResist(2) = intEarthResist(2) - thedamage(1)
            intWindResist(2) = intWindResist(2) - thedamage(1)
            intWaterResist(2) = intWaterResist(2) - thedamage(1)
            intHeartResist(2) = intHeartResist(2) - thedamage(1)
            intDarkResist(2) = intDarkResist(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn dropped your opponent's resistance by " & thedamage(1)
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
        End If
        If bWaitBoostResist(1) = True Then
            intFireResist(1) = intFireResist(1) + thedamage(1)
            intEarthResist(1) = intEarthResist(1) + thedamage(1)
            intWindResist(1) = intWindResist(1) + thedamage(1)
            intWaterResist(1) = intWaterResist(1) + thedamage(1)
            intHeartResist(1) = intHeartResist(1) + thedamage(1)
            intDarkResist(1) = intDarkResist(1) + thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn boosted your resistance by " & thedamage(1)
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
        End If
        If bWaitDjinnPP(1) = True Then
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn increased your PP by " & thedamage(1)
            PP(1) = PP(1) + thedamage(1)
            frmBattle.lblGen(10).Caption = PP(1)
        End If
            
            
    
    Else 'I am not the host
    
    
        If bWaitAttack(2) = True Then 'Waiting to attack
        
            CurrentOp = 1 'Sword drops on 1st player
            frmBattle.timeSword.Enabled = True
            
            HP(1) = HP(1) - thedamage(2) 'Decrease HP
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You took " & thedamage(2) & " damage."
        End If
        
        If bWaitCriticalAttack(2) = True Then
            CurrentOp = 1
            frmBattle.timeSword.Enabled = True
            HP(1) = HP(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your took " & thedamage(2) & " damage.  Critical Hit!"
        End If
        If bWaitSpecialAttack(2) = True Then
            CurrentOp = 1
            frmBattle.timeSword.Enabled = True
            HP(1) = HP(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent let out a howl.  Special Attack!  You took " & thedamage(2) & " damage."
        End If
        
        If bWaitPsynergy(2) = True Then 'Attacking Psynergy
            HP(1) = HP(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You took " & thedamage(2) & " damage from your opponent's Psynergy"
            
            CurrentOp = 1 'Zap Player 1
            
            strCurPsyDir = CharType(2)
            frmBattle.timePsynergy.Enabled = True
    
        End If
        
        If bWaitHeal(2) = True Then 'Waiting to heal
            HP(2) = HP(2) + thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent healed " & thedamage(2) & " HP."
    
        End If
        
        If bWaitBoostAttack(2) = True Then
            AP(2) = AP(2) + thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent increased his/her AP by " & thedamage(2) & "."
    
        End If
        
        If bWaitBoostDefense(2) = True Then
            Defense(2) = Defense(2) + thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent increased his/her Defense by " & thedamage(2) & "."
    
        End If
        
        If bWaitDropAttack(2) = True Then
            AP(1) = AP(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your AP was decreased by " & thedamage(2) & "."
        
        End If
        
        If bWaitDropDefense(2) = True Then
            Defense(1) = Defense(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your defense was decreased by " & thedamage(2) & "."
    
        End If
        
        If bWaitDefend(2) = True Then
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent Defended."
            'This needs to actually implement something
        End If
        
        If bWaitDjinnAttack(2) = True Then
            HP(1) = HP(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Djinn did " & thedamage(2) & " damage."
            frmBattle.timeDjinn.Enabled = True 'Enabled Djinn animation
            CurrentOp = 1 'Attack Player 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture 'Set Djinn picture to current elemental type
        End If
        
        If bWaitDjinnDropAttack(2) = True Then
            AP(1) = AP(1) - thedamage(2)
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Djinn dropped your AP by " & thedamage(2) & "."
    
        End If
        
        If bWaitDropDefense(2) = True Then
            Defense(1) = Defense(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Djinn dropped your Defense by " & thedamage(2) & "."
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnBoostAttack(2) = True Then
            AP(2) = AP(2) + thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Djinn boosted your opponent's AP by " & thedamage(2) & "."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2 'Djinn affects me
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnDefense(2) = True Then
            Defense(2) = Defense(2) + thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Djinn increased your opponent's Defense by " & thedamage(2) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnHeal(2) = True Then
            HP(2) = HP(2) + thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Djinn healed your opponent's HP by " & thedamage(2) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitDjinnSet(2) = True Then
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Djinn was set."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitSummon(2) = True Then
            HP(1) = HP(1) - thedamage(2)
            frmBattle.timeSummon.Enabled = True
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Summon did " & thedamage(2) & " damage."
    
            frmBattle.imgSummon(0).Picture = frmBattle.imgSummon(2 * iSummonType(1) - 1).Picture 'Set summon picture to current level of summon
            CurrentOp = 1
        
        End If
        
        If bWaitDropResist(2) = True Then
            intFireResist(1) = intFireResist(1) - thedamage(2)
            intEarthResist(1) = intEarthResist(1) - thedamage(2)
            intWindResist(1) = intWindResist(1) - thedamage(2)
            intWaterResist(1) = intWaterResist(1) - thedamage(2)
            intHeartResist(1) = intHeartResist(1) - thedamage(2)
            intDarkResist(1) = intDarkResist(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your foe's Djinn dropped your resistance by " & thedamage(2)
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
        End If
        If bWaitBoostResist(2) = True Then
            intFireResist(2) = intFireResist(2) + thedamage(2)
            intEarthResist(2) = intEarthResist(2) + thedamage(2)
            intWindResist(2) = intWindResist(2) + thedamage(2)
            intWaterResist(2) = intWaterResist(2) + thedamage(2)
            intHeartResist(2) = intHeartResist(2) + thedamage(2)
            intDarkResist(2) = intDarkResist(2) + thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your foe's Djinn boosted his resistance by " & thedamage(2)
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
        End If
        If bWaitDjinnPP(2) = True Then
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your foe's Djinn increased his or her PP by " & thedamage(1)
            PP(2) = PP(2) + thedamage(2)
        End If
    
    End If 'If hoston=true

'End If 'If boready(1) = true...

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
    bWaitCriticalAttack(1) = False
    bWaitSpecialAttack(1) = False
    bWaitBoostResist(1) = False
    bWaitDropResist(1) = False
    bWaitDjinnBoostAttack(1) = False
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
    bWaitDjinnBoostAttack(2) = False
    
    
    bWaitCriticalAttack(2) = False
    bWaitSpecialAttack(2) = False
    bWaitBoostResist(2) = False
    bWaitDropResist(2) = False
End If

PlayerWait = 1
frmBattle.timeWait.Enabled = True


End Sub
Sub Player2Command()
On Error Resume Next
strdtime = Format(Now, "dd-mmmm hh:mm:ss AM/PM")
rsave = App.Path & "\userdata.ini"

On Error Resume Next

'If (hoston = True And bOReady(2) = True) Or (hoston = False And bOReady(1) = True) Then
    If hoston = False Then 'I am the client
        
        If bWaitAttack(1) = True Then 'Waiting to attack
        
            CurrentOp = 2 'Sword drops on 2nd player
            frmBattle.timeSword.Enabled = True
            
            HP(2) = HP(2) - thedamage(1) 'Decrease HP
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You did " & thedamage(1) & " damage."
        End If
        
        If bWaitCriticalAttack(1) = True Then
            CurrentOp = 2
            frmBattle.timeSword.Enabled = True
            HP(2) = HP(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You did " & thedamage(1) & " damage.  Critical Hit!"
        End If
        If bWaitSpecialAttack(1) = True Then
            CurrentOp = 2
            frmBattle.timeSword.Enabled = True
            HP(2) = HP(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You let out a howl.  Special Attack!  You did " & thedamage(1) & " damage."
        End If
        
        If bWaitPsynergy(1) = True Then 'Attacking Psynergy
            HP(2) = HP(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Psynergy did " & thedamage(1) & " damage."
            
            CurrentOp = 2 'Zap Player 2
            
            strCurPsyDir = CharType(1)
            frmBattle.timePsynergy.Enabled = True
    
        End If
        
        If bWaitHeal(1) = True Then 'Waiting to heal
            'Healing Psynergy will always heal the same amount of damage
            thedamage(1) = CInt(strPsyDamage(RealPsy))
            
            If HP(1) + thedamage(1) > MaxHP(1) Then 'Can't heal more than max
                thedamage(1) = thedamage(1) - ((HP(1) + thedamage(1)) - MaxHP(1))
            End If
            
            If hoston = True Then
            frmHost.Host.SendData "DMG" & thedamage(1) & vbCrLf 'Increase HP by this much
            End If
            If hoston = False Then
            frmJoin.Client.SendData "DMG" & thedamage(1) & vbCrLf
            End If
    
            HP(1) = HP(1) + thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You healed " & thedamage(1) & " HP."
    
        End If
        
        If bWaitBoostAttack(1) = True Then
            AP(1) = AP(1) + thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You increased your AP by " & thedamage(1) & "."
    
        End If
        
        If bWaitBoostDefense(1) = True Then
            Defense(1) = Defense(1) + thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You increased your Defense by " & thedamage(1) & "."
    
        End If
        
        If bWaitDropAttack(1) = True Then
            AP(2) = AP(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You decreased your opponent's AP by " & thedamage(1) & "."
        
        End If
        
        If bWaitDropDefense(1) = True Then
            Defense(2) = Defense(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You decrased your opponent's defense by " & thedamage(1) & "."
    
        End If
        
        If bWaitDefend(1) = True Then
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You Defended."
            'This needs to actually implement something
        End If
        
        If bWaitDjinnAttack(1) = True Then
            HP(2) = HP(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn did " & thedamage(1) & " damage."
            frmBattle.timeDjinn.Enabled = True 'Enabled Djinn animation
            CurrentOp = 2 'Attack Player 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture 'Set Djinn picture to current elemental type
        End If
        
        If bWaitDjinnDropAttack(1) = True Then
            AP(2) = AP(2) - thedamage(1)
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn dropped your opponent's AP by " & thedamage(1) & "."
    
        End If
        
        If bWaitDropDefense(1) = True Then
            Defense(2) = Defense(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn dropped your opponent's Defense by " & thedamage(1) & "."
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnBoostAttack(1) = True Then
            AP(1) = AP(1) + thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn boosted your AP by " & thedamage(1) & "."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1 'Djinn affects me
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnDefense(1) = True Then
            Defense(1) = Defense(1) + thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn increased your Defense by " & thedamage(1) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnHeal(1) = True Then
            HP(1) = HP(1) + thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn healed your HP by " & thedamage(1) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitDjinnSet(1) = True Then
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn was set."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitSummon(1) = True Then
            HP(2) = HP(2) - thedamage(1)
            frmBattle.timeSummon.Enabled = True
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Summon did " & thedamage(1) & " damage."
    
            frmBattle.imgSummon(0).Picture = frmBattle.imgSummon(2 * iSummonType(1)).Picture 'Set summon picture to current level of summon
            CurrentOp = 2
        
        End If
        If bWaitDropResist(1) = True Then
            intFireResist(2) = intFireResist(2) - thedamage(1)
            intEarthResist(2) = intEarthResist(2) - thedamage(1)
            intWindResist(2) = intWindResist(2) - thedamage(1)
            intWaterResist(2) = intWaterResist(2) - thedamage(1)
            intHeartResist(2) = intHeartResist(2) - thedamage(1)
            intDarkResist(2) = intDarkResist(2) - thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn dropped your foe's resistance by " & thedamage(1)
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
        End If
        If bWaitBoostResist(1) = True Then
            intFireResist(1) = intFireResist(1) + thedamage(1)
            intEarthResist(1) = intEarthResist(1) + thedamage(1)
            intWindResist(1) = intWindResist(1) + thedamage(1)
            intWaterResist(1) = intWaterResist(1) + thedamage(1)
            intHeartResist(1) = intHeartResist(1) + thedamage(1)
            intDarkResist(1) = intDarkResist(1) + thedamage(1)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn boosted your resistance by " & thedamage(1)
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
        End If
        If bWaitDjinnPP(1) = True Then
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your Djinn increased your PP by " & thedamage(1)
            PP(1) = PP(1) + thedamage(1)
            frmBattle.lblGen(10).Caption = PP(1)
        End If
    
    
    Else 'I am the host
    
    
        If bWaitAttack(2) = True Then 'Waiting to attack
        
            CurrentOp = 1 'Sword drops on 1st player
            frmBattle.timeSword.Enabled = True
            
            HP(1) = HP(1) - thedamage(2) 'Decrease HP
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You took " & thedamage(2) & " damage."
        End If
        
        If bWaitCriticalAttack(2) = True Then
            CurrentOp = 1
            frmBattle.timeSword.Enabled = True
            HP(1) = HP(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your took " & thedamage(2) & " damage.  Critical Hit!"
        End If
        If bWaitSpecialAttack(2) = True Then
            CurrentOp = 1
            frmBattle.timeSword.Enabled = True
            HP(1) = HP(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent let out a howl.  Special Attack!  You took " & thedamage(2) & " damage."
        End If
        
        If bWaitPsynergy(2) = True Then 'Attacking Psynergy
            HP(1) = HP(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "You took " & thedamage(2) & " damage from your opponent's Psynergy"
            
            CurrentOp = 1 'Zap Player 1
            
            strCurPsyDir = CharType(2)
            frmBattle.timePsynergy.Enabled = True
    
        End If
        
        If bWaitHeal(2) = True Then 'Waiting to heal
            HP(2) = HP(2) + thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent healed " & thedamage(2) & " HP."
    
        End If
        
        If bWaitBoostAttack(2) = True Then
            AP(2) = AP(2) + thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent increased his/her AP by " & thedamage(2) & "."
    
        End If
        
        If bWaitBoostDefense(2) = True Then
            Defense(2) = Defense(2) + thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent increased his/her Defense by " & thedamage(2) & "."
    
        End If
        
        If bWaitDropAttack(2) = True Then
            AP(1) = AP(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your AP was decreased by " & thedamage(2) & "."
        
        End If
        
        If bWaitDropDefense(2) = True Then
            Defense(1) = Defense(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your defense was decreased by " & thedamage(2) & "."
    
        End If
        
        If bWaitDefend(2) = True Then
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent Defended."
            'This needs to actually implement something
        End If
        
        If bWaitDjinnAttack(2) = True Then
            HP(1) = HP(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Djinn did " & thedamage(2) & " damage."
            frmBattle.timeDjinn.Enabled = True 'Enabled Djinn animation
            CurrentOp = 1 'Attack Player 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture 'Set Djinn picture to current elemental type
        End If
        
        If bWaitDjinnDropAttack(2) = True Then
            AP(1) = AP(1) - thedamage(2)
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Djinn dropped your AP by " & thedamage(2) & "."
    
        End If
        
        If bWaitDropDefense(2) = True Then
            Defense(1) = Defense(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Djinn dropped your Defense by " & thedamage(2) & "."
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnBoostAttack(2) = True Then
            AP(2) = AP(2) + thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Djinn boosted your opponent's AP by " & thedamage(2) & "."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2 'Djinn affects me
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnDefense(2) = True Then
            Defense(2) = Defense(2) + thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Djinn increased your opponent's Defense by " & thedamage(2) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        
        If bWaitDjinnHeal(2) = True Then
            HP(2) = HP(2) + thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Djinn healed your opponent's HP by " & thedamage(2) & "."
        
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitDjinnSet(2) = True Then
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Djinn was set."
    
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 2
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
        End If
        If bWaitSummon(2) = True Then
            HP(1) = HP(1) - thedamage(2)
            frmBattle.timeSummon.Enabled = True
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your opponent's Summon did " & thedamage(2) & " damage."
            frmBattle.imgSummon(0).Picture = frmBattle.imgSummon(2 * iSummonType(1) - 1).Picture 'Set summon picture to current level of summon
            CurrentOp = 1
        
        End If
        If bWaitBoostResist(2) = True Then
            intFireResist(2) = intFireResist(2) + thedamage(2)
            intEarthResist(2) = intEarthResist(2) + thedamage(2)
            intWindResist(2) = intWindResist(2) + thedamage(2)
            intWaterResist(2) = intWaterResist(2) + thedamage(2)
            intHeartResist(2) = intHeartResist(2) + thedamage(2)
            intDarkResist(2) = intDarkResist(2) + thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your foe's Djinn dropped your resistance by " & thedamage(2)
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
        End If
        If bWaitDropResist(2) = True Then
            intFireResist(1) = intFireResist(1) - thedamage(2)
            intEarthResist(1) = intEarthResist(1) - thedamage(2)
            intWindResist(1) = intWindResist(1) - thedamage(2)
            intWaterResist(1) = intWaterResist(1) - thedamage(2)
            intHeartResist(1) = intHeartResist(1) - thedamage(2)
            intDarkResist(1) = intDarkResist(1) - thedamage(2)
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your foe's Djinn boosted his resistance by " & thedamage(2)
            frmBattle.imgDjinn(0).Picture = frmBattle.imgDjinn(CInt(strOpDjinnType)).Picture
            frmBattle.timeDjinn.Enabled = True
            CurrentOp = 1
        End If
        If bWaitDjinnPP(2) = True Then
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & "Your foe's Djinn increased hir or her PP by " & thedamage(1)
            PP(2) = PP(2) + thedamage(2)
        End If
    
    End If 'If hoston=true

'End If 'If boready(1) = true...


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
    bWaitDjinnBoostAttack(1) = False
    bWaitDjinnSet(1) = False
    bWaitCriticalAttack(1) = False
    bWaitSpecialAttack(1) = False
    bWaitBoostResist(1) = False
    bWaitDropResist(1) = False
    bWaitDjinnBoostAttack(1) = False
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
    bWaitCriticalAttack(2) = False
    bWaitSpecialAttack(2) = False
    bWaitBoostResist(2) = False
    bWaitDropResist(2) = False
    bWaitDjinnBoostAttack(2) = False
End If


frmBattle.timeWait.Enabled = True
PlayerWait = 2



End Sub
Function AutoScroll(txtbox As RichTextBox)
'If txtbox.MultiLine = False Then Exit Function
On Error Resume Next
If bAutoScroll = True Then
    txtbox.SelLength = 0

    If Len(Trim(txtbox.Text)) > 0 Then
        If Right$(txtbox.Text, 1) = vbCrLf Then
            txtbox.SelStart = Len(txtbox.Text) - 1
            Exit Function
        End If
        txtbox.SelStart = Len(txtbox.Text)
    End If
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
'Re-enables menu options
On Error Resume Next

frmBattle.lblAttack.Visible = True
frmBattle.lblPsynergy.Visible = True
frmBattle.lblDjinn.Visible = True
frmBattle.lblSummon.Visible = True
frmBattle.lblDefend.Visible = True
frmBattle.shpMenu.Visible = True
frmBattle.lblStatus.Visible = True
frmBattle.lblBackTurn.Visible = True
If HP(1) > 0 Then
    intTurn = 1
Else
    intTurn = 2
    bOReady(1) = True
    AttackType(1) = "DEAD"
End If

SelectTarget = False
frmBattle.lblSelectTarget.Visible = False

frmBattle.imgTurn.Visible = True
frmBattle.imgTurn.Top = frmBattle.imgYou(0).Top + frmBattle.imgYou(0).Height

If AP(1) < 10 And CharName(1) <> "The Wise One" Then
    AP(1) = 10
End If
If AP(2) < 10 And CharName(2) <> "The Wise One" Then
    AP(2) = 10
End If
For i = 0 To 3
    If HP(i + 1) < 0 Then HP(i + 1) = 0
    frmBattle.shpHP(i).Width = HP(i + 1) / 5
    frmBattle.shpHP(i).Width = HP(i + 1) / 5
    frmBattle.lblHP(i).Caption = HP(i + 1)
Next 'i

frmBattle.lblPP(0).Caption = PP(1)
frmBattle.lblPP(1).Caption = PP(2)


frmBattle.cmdReset.Enabled = True 'You can attempt to reset again
frmBattle.txtDisplay.Visible = False

If hoston = True Then 'Send double checking HPs
    frmHost.Host.SendData "MYHP3" & HP(1) & vbCrLf
    frmHost.Host.SendData "MYHP4" & HP(2) & vbCrLf
Else
    frmJoin.Client.SendData "MYHP3" & HP(1) & vbCrLf
    frmJoin.Client.SendData "MYHP4" & HP(2) & vbCrLf
End If

For i = 0 To 6
    frmBattle.imgIcon(i).Visible = True
Next 'i
End Sub


Public Sub LoadBattle()
'perform prebattle setup
'On Error GoTo err
DataSent = True 'Data has been sent
'Send stats
If hoston = True Then
    frmHost.Host.SendData "LVL" & strLvl & vbCrLf
    DoEvents
    frmHost.Host.SendData "RATING" & strRating & vbCrLf
    DoEvents
    frmHost.Host.SendData "CHAR3" & strChar(1) & vbCrLf
    DoEvents
    frmHost.Host.SendData "CHAR4" & strChar(2) & vbCrLf
    DoEvents
    frmHost.Host.SendData "TYPE3" & CharType(3) & vbCrLf & "TYPE4" & CharType(4) & vbCrLf
    DoEvents
    strOpponentIP = frmHost.Host.RemoteHostIP
Else
    frmJoin.Client.SendData "LVL" & strLvl & vbCrLf
    DoEvents
    frmJoin.Client.SendData "RATING" & strRating & vbCrLf
    DoEvents
    frmJoin.Client.SendData "CHAR3" & strChar(1) & vbCrLf
    DoEvents
    frmJoin.Client.SendData "CHAR4" & strChar(2) & vbCrLf
    DoEvents
    frmJoin.Client.SendData "TYPE3" & CharType(3) & vbCrLf & "TYPE4" & CharType(4) & vbCrLf
    DoEvents
    strOpponentIP = frmJoin.Client.RemoteHostIP
End If

bWaitToResetDjinn = False

GameOver = False 'Game not over yet
PlayerWait = 0

frmBattle.timeLoadBattle.Enabled = False

'Disable all variables
For i = 1 To 4
    AttackType(i) = ""
    Target(i) = 0
Next 'i

Reset(1) = False
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
bWaitCriticalAttack(1) = False
bWaitSpecialAttack(1) = False
bWaitBoostResist(1) = False
bWaitDropResist(1) = False

Reset(2) = False
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
bWaitCriticalAttack(2) = False
bWaitSpecialAttack(2) = False
bWaitBoostResist(2) = False
bWaitDropResist(2) = False

'frmArena.timecount.Enabled = False
Unload frmArena

frmBattle.txtDisplay.Visible = False
frmBattle.txtDisplay.Text = "---Battle Information---"

curCount = 20 'Set interval for clock countdown

'Set the names of users
frmBattle.lblGen(15).Caption = strMyUserName
frmBattle.lblGen(16).Caption = strOpponent

'Handicapp Code
RelativeRating(1) = CInt(strRating)
RelativeLVL(1) = CInt(strLvl)
iTotalDjinn(1) = GetDjinn(strRating)
RelativeDjinn(1) = iTotalDjinn(1)

'Dim intHandVar As Long
'intHandVar = Handicap(1) * 50
'
'RelativeRating(1) = RelativeRating(1) + intHandVar
If CStr(stroLvl) = "" Then stroLvl = "1"

'New handicapp code:
If Handicap(1) <> 0 Then
    RelativeLVL(1) = CStr(stroLvl)
    RelativeRating(1) = opRating
    RelativeDjinn(1) = GetDjinn(CStr(RelativeRating(1)))
    If RelativeDjinn(1) < 1 Then RelativeDjinn(1) = 1
End If
If Handicap(2) <> 0 Then
    RelativeLVL(2) = CStr(stroLvl)
    RelativeRating(2) = opRating
    RelativeDjinn(2) = GetDjinn(CStr(RelativeRating(2)))
    If RelativeDjinn(2) < 1 Then RelativeDjinn(2) = 1
End If

'If bEqualizeWeapons = True Then
    




'    For i = 1 To (Handicap(1) * 50)
'        If (CInt(strRating) + i) Mod 50 = 0 Then
'            RelativeLVL(1) = RelativeLVL(1) + 1
'        End If
'        If (CInt(strRating) + i) Mod 100 = 0 Then
'            RelativeDjinn(1) = RelativeDjinn(1) + 1
'        End If
'    Next 'i
'End If
'If Handicap(1) < 0 Then
'    For i = 1 To Abs((Handicap(1) * 50))
'        If (CInt(strRating) - i) Mod 50 = 0 Then
'            RelativeLVL(1) = RelativeLVL(1) - 1
'        End If
'        If (CInt(strRating) - i) Mod 100 = 0 Then
'            RelativeDjinn(1) = RelativeDjinn(1) - 1
'        End If
'    Next 'i
'End If

Dim strDjinnState As String
strDjinnState = GetFromIni("GEN", "DJINN", App.Path & "\settings")
intDjinnStandby(1) = 0
intDjinnStandby(2) = 0
For i = 1 To RelativeDjinn(1)
    If strDjinnState = "0" Then
        Djinn(i).State = 0
    ElseIf strDjinnState = "1" Then
        Djinn(i).State = 1
    Else
        If Djinn(i).State = 2 Then
            Djinn(i).State = 0
        End If
    End If
    For q = 1 To 2
        If Djinn(i).State = 1 And Djinn(i).Character = q Then
            intDjinnStandby(q) = intDjinnStandby(q) + 1
        End If
    Next 'q
Next 'i

If hoston = True Then
    frmHost.Host.SendData "RELRATING" & RelativeRating(1) & vbCrLf
    DoEvents
    frmHost.Host.SendData "RELLVL" & RelativeLVL(1) & vbCrLf
    DoEvents
    frmHost.Host.SendData "RELDJINN" & RelativeDjinn(1) & vbCrLf
    DoEvents
Else
    frmJoin.Client.SendData "RELRATING" & RelativeRating(1) & vbCrLf
    DoEvents
    frmJoin.Client.SendData "RELLVL" & RelativeLVL(1) & vbCrLf
    DoEvents
    frmJoin.Client.SendData "RELDJINN" & RelativeDjinn(1) & vbCrLf
    DoEvents
End If


Dim iLvl As Integer 'Current level
iLvl = RelativeLVL(1)

For i = 1 To 2
    bCustomChar(i) = FindWhichCharacter(strChar(i))
    
    If strChar(i) = "Isaac" Then
        Char(i) = 0
        HP(i) = 100 + (12 * iLvl)
        AP(i) = 25 + (3 * iLvl)
        PP(i) = 20 + (4 * iLvl)
        Defense(i) = iLvl * 2
        PsyBonus(i) = 1
        CharType(i) = "E"
        Luck(i) = 3 + (1.25 * iLvl / 2)
        Speed(i) = 15 + (1.25 * iLvl)
    End If
    If strChar(i) = "Young Isaac" Then
        Char(i) = 22
        HP(i) = 90 + (10 * iLvl)
        AP(i) = 35 + (3.5 * iLvl)
        PP(i) = 17 + (4 * iLvl)
        Defense(i) = iLvl * 2.25
        PsyBonus(i) = 1
        CharType(i) = "E"
        Luck(i) = 6 + (1.25 * iLvl / 2)
        Speed(i) = 17 + (1.5 * iLvl)
    End If
    If strChar(i) = "Young Garet" Then
        Char(i) = 23
        HP(i) = 105 + (12 * iLvl)
        AP(i) = 37 + (3.5 * iLvl)
        PP(i) = 24 + (3 * iLvl)
        Defense(i) = iLvl * 1.25
        PsyBonus(i) = 1
        CharType(i) = "F"
        Luck(i) = 9 + (1.75 * iLvl / 2)
        Speed(i) = 10 + (iLvl)
    End If
    If strChar(i) = "Garret" Then
        Char(i) = 1
        HP(i) = 130 + (16.5 * iLvl)
        AP(i) = 20 + (2 * iLvl)
        PP(i) = 11 + (2.1 * iLvl)
        Defense(i) = iLvl * 3
        PsyBonus(i) = 1
        CharType(i) = "F"
        Luck(i) = 3 + (0.95 * iLvl / 2)
        Speed(i) = 7 + (1.25 * iLvl)
    End If
    If strChar(i) = "Ivan" Then
        Char(i) = 2
        HP(i) = 80 + (8 * iLvl)
        AP(i) = 14 + (2 * iLvl)
        PP(i) = 36 + (5.2 * iLvl)
        Defense(i) = iLvl * 1.5
        PsyBonus(i) = 1.25
        CharType(i) = "N"
        Luck(i) = 5 + (1.75 * iLvl / 2)
        Speed(i) = 20 + (1.5 * iLvl)
    End If
    If strChar(i) = "Cloud" Then
        Char(i) = 17
        HP(i) = 80 + (8 * iLvl)
        AP(i) = 13 + (1.75 * iLvl)
        PP(i) = 40 + (7.5 * iLvl)
        Defense(i) = iLvl * 1.75
        PsyBonus(i) = 1
        CharType(i) = "N"
        Luck(i) = 2 + (0.75 * iLvl / 2)
        Speed(i) = 20 + (1.25 * iLvl)
    End If
    If strChar(i) = "Mia" Then
        Char(i) = 3
        HP(i) = 85 + (9 * iLvl)
        AP(i) = 16 + (2 * iLvl)
        PP(i) = 40 + (5 * iLvl)
        Defense(i) = iLvl * 2
        PsyBonus(i) = 0.75
        CharType(i) = "W"
        Luck(i) = 5 + (1.5 * iLvl / 2)
        Speed(i) = 19 + (1.5 * iLvl)
    End If
    If strChar(i) = "Saturos" Then
        Char(i) = 4
        HP(i) = 130 + (13 * iLvl)
        AP(i) = 35 + (3.75 * iLvl)
        PP(i) = 8 + (2 * iLvl)
        Defense(i) = iLvl * 1.5
        PsyBonus(i) = 1
        CharType(i) = "F"
        Luck(i) = 2 + (1 * iLvl / 2)
        Speed(i) = 13 + (1.15 * iLvl)
    End If
    If strChar(i) = "Menardi" Then
        Char(i) = 5
        HP(i) = 120 + (12 * iLvl)
        AP(i) = 22.5 + (3 * iLvl)
        PP(i) = 19 + (4 * iLvl)
        Defense(i) = iLvl * 2
        PsyBonus(i) = 1
        CharType(i) = "F"
        Luck(i) = 2.5 + (0.9 * iLvl / 2)
        Speed(i) = 12 + (1.25 * iLvl)
    End If
    If strChar(i) = "Felix" Then
        Char(i) = 6
        HP(i) = 110 + (12.5 * iLvl)
        AP(i) = 31 + (3.5 * iLvl)
        PP(i) = 13 + (3 * iLvl)
        Defense(i) = iLvl * 2.5
        PsyBonus(i) = 1
        CharType(i) = "H"
        Luck(i) = 5 + (2.2 * iLvl / 2)
        Speed(i) = 14 + (1.1 * iLvl)
    End If
    If strChar(i) = "Sheba" Then
        Char(i) = 7
        HP(i) = 95 + (10.5 * iLvl)
        AP(i) = 15 + (3 * iLvl)
        PP(i) = 25 + (4 * iLvl)
        Defense(i) = iLvl * 1.5
        PsyBonus(i) = 1.3
        CharType(i) = "N"
        Luck(i) = 6.5 + (2.05 * iLvl)
        Speed(i) = 25 + (2 * iLvl)
    End If
    If strChar(i) = "Jenna" Then
        Char(i) = 8
        HP(i) = 98 + (10 * iLvl)
        AP(i) = 25 + (3 * iLvl)
        PP(i) = 21 + (3 * iLvl)
        Defense(i) = iLvl * 1.05
        PsyBonus(i) = 0.75
        CharType(i) = "F"
        Luck(i) = 6.75 + (2 * iLvl / 2)
        Speed(i) = 11 + (1.05 * iLvl)
    End If
    If strChar(i) = "Kraden" Then
        Char(i) = 9
        HP(i) = 50 + (6 * iLvl)
        AP(i) = 10 + (1 * iLvl)
        PP(i) = 40 + (4 * iLvl)
        Defense(i) = iLvl * 1
        PsyBonus(i) = 1.3
        CharType(i) = "D"
        Luck(i) = 2.75 + (0.8 * iLvl / 2)
        Speed(i) = 5 + (iLvl)
    End If
    If strChar(i) = "Alex" Then
        Char(i) = 10
        HP(i) = 110 + (13 * iLvl)
        AP(i) = 30 + (3 * iLvl)
        PP(i) = 8 + (2 * iLvl)
        Defense(i) = iLvl * 1.5
        PsyBonus(i) = 0.75
        CharType(i) = "W"
        Luck(i) = 3 + (0.75 * iLvl)
        Speed(i) = 10 + (1.3 * iLvl)
    End If
    If strChar(i) = "Caption Contest Character" Then
        Char(i) = 11
        HP(i) = 150 + (16 * iLvl)
        AP(i) = 41 + (3.75 * iLvl)
        PP(i) = 0
        Defense(i) = iLvl * 3
        PsyBonus(i) = 1
        CharType(i) = "W"
        Luck(i) = 4 + (0.75 * iLvl / 2)
        Speed(i) = 19 + (1.25 * iLvl)
    End If
    If strChar(i) = "Guard" Then
        Char(i) = 12
        HP(i) = 70 + (21 * iLvl)
        AP(i) = 24 + (3.8 * iLvl)
        PP(i) = 0
        Defense(i) = iLvl * 1.25
        PsyBonus(i) = 0.75
        CharType(i) = "E"
        Luck(i) = 2.75 + (1.4 * iLvl / 2)
        Speed(i) = 5 + (3 * iLvl)
    End If
    If strChar(i) = "Gladiator" Then
        Char(i) = 13
        HP(i) = 155 + (8 * iLvl)
        AP(i) = 38 + (1.5 * iLvl)
        PP(i) = 0
        Defense(i) = iLvl * 1.25
        PsyBonus(i) = 0
        CharType(i) = "E"
        Luck(i) = 7
        Speed(i) = 30
    End If
    If strChar(i) = "Piers" Then
        Char(i) = 14
        HP(i) = 80 + (11 * iLvl)
        AP(i) = 19.5 + (1.9 * iLvl)
        PP(i) = 37 + (5.2 * iLvl)
        Defense(i) = iLvl * 2
        PsyBonus(i) = 1
        CharType(i) = "W"
        Luck(i) = 3.5 + (1.25 * iLvl / 2)
        Speed(i) = 7 + (1.05 * iLvl)
    End If
    If strChar(i) = "Kenny" Then
        Char(i) = 15
        HP(i) = 120 + (26 * iLvl)
        AP(i) = 43 + (4 * iLvl)
        PP(i) = 50 + (8 * iLvl)
        Defense(i) = iLvl * 3.5
        PsyBonus(i) = 2
        CharType(i) = "F"
        Luck(i) = 8 + (4 * iLvl / 2)
        Speed(i) = 50 + (5 * iLvl)
    End If
    If strChar(i) = "KOS" Then
        Char(i) = 16
        HP(i) = 100 + (12 * iLvl)
        AP(i) = 32 + (3.2 * iLvl)
        PP(i) = 0
        Defense(i) = iLvl * 2.2
        PsyBonus(i) = 1
        CharType(i) = "D"
        Luck(i) = 8 + (2.4 * iLvl / 2)
        Speed(i) = 22 + (1.5 * iLvl)
    End If
    If strChar(i) = "Purple Piers" Then
        Char(i) = 18
        HP(i) = 85 + (12 * iLvl)
        AP(i) = 24 + (2 * iLvl)
        PP(i) = 44 + (6 * iLvl)
        Defense(i) = iLvl * 2
        PsyBonus(i) = 1
        CharType(i) = "W"
        Luck(i) = 3.5 + (1.25 * iLvl / 2)
        Speed(i) = 8 + (1.2 * iLvl)
    End If
    If strChar(i) = "Agiato" Then
        Char(i) = 19
        HP(i) = 86 + (11 * iLvl)
        AP(i) = 24 + (3 * iLvl)
        PP(i) = 34 + (7 * iLvl)
        Defense(i) = iLvl * 2
        PsyBonus(i) = 1
        CharType(i) = "D"
        Luck(i) = 3.5 + (1.25 * iLvl / 2)
        Speed(i) = 18 + (1.25 * iLvl)
    End If
    If strChar(i) = "Karst" Then
        Char(i) = 20
        HP(i) = 80 + (10 * iLvl)
        AP(i) = 19 + (2.76 * iLvl)
        PP(i) = 32 + (6 * iLvl)
        Defense(i) = iLvl * 2.1
        PsyBonus(i) = 1
        CharType(i) = "D"
        Luck(i) = 3 + (1.5 * iLvl / 2)
        Speed(i) = 17 + (1.35 * iLvl)
    End If
    If strChar(i) = "The Wise One" Then
        Char(i) = 21
        HP(i) = 110 + (11.1 * iLvl)
        AP(i) = 0
        PP(i) = 65 + (11 * iLvl / 2)
        Defense(i) = 0
        PsyBonus(i) = 1
        CharType(i) = "H"
        Luck(i) = 0
        Speed(i) = 15 + (1.3 * iLvl)
    End If
    If bCustomChar(i) <> 999 Then
        With CustomChar(bCustomChar(i))
        Char(i) = .Picture
        HP(i) = 60 + (.BaseHP * 20) + ((.BaseHP + 7) * iLvl)
        AP(i) = 10 + (.BaseAP * 3) + (((.BaseAP * 5) + iLvl) * 2.5)
        PP(i) = 10 + (.BasePP * 5) + ((.BasePP + iLvl) * 4)
        Defense(i) = iLvl * ((.BaseDefense + 1) / 2)
        PsyBonus(i) = 1
        CharType(i) = .Type
        Luck(i) = (.BaseLuck) + (.BaseLuck * iLvl / 5)
        End With
    End If
Next 'i



If HP(1) <= 0 Then
    strChar(1) = "Isaac"
    Call LoadBattle
    Exit Sub
End If



MaxHP(1) = HP(1)
MaxHP(2) = HP(2)

If bCustomChar(1) = 999 Then
    Call SetPowerResist(strChar(1), 1) 'set power and resist
Else
    Call SetCustomPowerResist(bCustomChar(1))
End If

If bCustomChar(2) = 999 Or bCustomChar(2) = 0 Then
    Call SetPowerResist(strChar(2), 2) 'set power and resist
Else
    Call SetCustomPowerResist(bCustomChar(2))
End If



'If winRace = 1 Then 'If you won the pre-battle race, you attack first
'    If hoston = True Then
'        intFirstAttack = 1
'    Else
'        intFirstAttack = 2
'    End If
 '   'HP(1) = HP(1) * 0.25 + HP(1)
 '   'AP(1) = AP(1) * 0.25 + AP(1)
 '   'PP(10) = PP(1) * 0.25 + PP(1)
 '   'Defense(1) = Defense(1) * 0.25 + Defense(1)
'Else
'    If hoston = True Then
'        intFirstAttack = 2
''    Else
'        intFirstAttack = 1
'    End If
'End If

For i = 1 To 2
    If hoston = True Then 'Send Stats
        frmHost.Host.SendData "HP" & (i + 2) & HP(i) & vbCrLf & "SPEED" & (i + 2) & Speed(i) & vbCrLf & "AP" & (i + 2) & AP(i) & vbCrLf & "DEFENSE" & (i + 2) & Defense(i) & vbCrLf & "PIC" & (i + 2) & Char(i) & vbCrLf & "CHARTYPE" & (i + 2) & CharType(i) & vbCrLf & "RESEARTH" & (i + 2) & intEarthResist(i) & vbCrLf & "RESFIRE" & (i + 2) & intFireResist(i) & vbCrLf & "RESWIND" & (i + 2) & intWindResist(i) & vbCrLf & "RESWATER" & (i + 2) & intWaterResist(i) & vbCrLf & "RESHEART" & (i + 2) & intHeartResist(i) & vbCrLf & "RESDARK" & (i + 2) & intDarkResist(i) & vbCrLf
    Else
        frmJoin.Client.SendData "HP" & (i + 2) & HP(i) & vbCrLf & "SPEED" & (i + 2) & Speed(i) & vbCrLf & "AP" & (i + 2) & AP(i) & vbCrLf & "DEFENSE" & (i + 2) & Defense(i) & vbCrLf & "PIC" & (i + 2) & Char(i) & vbCrLf & "CHARTYPE" & (i + 2) & CharType(i) & vbCrLf & "RESEARTH" & (i + 2) & intEarthResist(i) & vbCrLf & "RESFIRE" & (i + 2) & intFireResist(i) & vbCrLf & "RESWIND" & (i + 2) & intWindResist(i) & vbCrLf & "RESWATER" & (i + 2) & intWaterResist(i) & vbCrLf & "RESHEART" & (i + 2) & intHeartResist(i) & vbCrLf & "RESDARK" & (i + 2) & intDarkResist(i) & vbCrLf
        'frmJoin.Client.SendData "HP" & (i + 2) & HP(i) & vbCrLf
        'frmJoin.Client.SendData "AP" & (i + 2) & AP(i) & vbCrLf
        'frmJoin.Client.SendData "DEFENSE" & (i + 2) & Defense(i) & vbCrLf
        'frmJoin.Client.SendData "PIC" & (i + 2) & Char(i) & vbCrLf
        'frmJoin.Client.SendData "CHARTYPE" & (i + 2) & CharType(i) & vbCrLf
        'frmJoin.Client.SendData "RESEARTH" & (i + 2) & intEarthResist(i) & vbCrLf
        'frmJoin.Client.SendData "RESFIRE" & (i + 2) & intFireResist(i) & vbCrLf
        'frmJoin.Client.SendData "RESWIND" & (i + 2) & intWindResist(i) & vbCrLf
        'frmJoin.Client.SendData "RESWATER" & (i + 2) & intWaterResist(i) & vbCrLf
        'frmJoin.Client.SendData "RESHEART" & (i + 2) & intHeartResist(i) & vbCrLf
        'frmJoin.Client.SendData "RESDARK" & (i + 2) & intDarkResist(i) & vbCrLf
    End If
Next 'i



'If TimedMatch = True Then
'    frmBattle.timecount.Enabled = True
'End If

'Dim strDjinnOption As String
'Dim nFile As String
'nFile = App.Path & "\settings.ini"
'strDjinnOption = GetFromIni("GEN", "DJINN", nFile)
'If strDjinnOption = "" Then strDjinnOption = "2"

'If CInt(strDjinnOption) = 0 Then
'    For q = 1 To 10
'        bDjinnSet(q) = 0
'    Next 'q
'ElseIf CInt(strDjinnOption) = 1 Then
'    For q = 1 To 10
'        bDjinnSet(q) = 1
'    Next 'q
'End If


'intDjinnStandby(1) = 0
'intDjinnStandby(2) = 0
'For i = 1 To 20
'    If Djinn(i).State = 1 And Djinn(i).Name <> "" Then
'        intDjinnStandby(1) = intDjinnStandby(1) + 1
'    End If
'Next 'i

'Update labels and HP bars
frmBattle.lblHP(0).Caption = HP(1)
frmBattle.lblPP(0).Caption = PP(1)
frmBattle.lblHP(1).Caption = HP(2)
frmBattle.lblPP(1).Caption = PP(2)


frmBattle.shpHP(0).Width = HP(1) / 5
frmBattle.shpHP(1).Width = HP(2) / 5


'Set character picture
If bCustomChar(1) = 999 Or bCustomChar(1) = 0 Then
frmBattle.imgYou(0).Picture = frmBattle.imgUser(Char(1)).Picture
frmBattle.imgYou(2).Picture = frmBattle.imgUser(Char(1)).Picture
Else
frmBattle.imgYou(0).Picture = LoadPicture(App.Path & "\files\" & CustomChar(bCustomChar(1)).Picture & ".gif")
frmBattle.imgYou(2).Picture = LoadPicture(App.Path & "\files\" & CustomChar(bCustomChar(1)).Picture & ".gif")
End If
If bCustomChar(2) = 999 Then
frmBattle.imgYou(4).Picture = frmBattle.imgUser(Char(2)).Picture
frmBattle.imgYou(6).Picture = frmBattle.imgUser(Char(2)).Picture
Else
frmBattle.imgYou(4).Picture = LoadPicture(App.Path & "\files\" & CustomChar(bCustomChar(1)).Picture & ".gif")
frmBattle.imgYou(6).Picture = LoadPicture(App.Path & "\files\" & CustomChar(bCustomChar(1)).Picture & ".gif")
End If


frmBattle.txtDisplay.Visible = False

If hoston = True Then
    frmBattle.imgArena.Picture = frmHost.imgBG.Picture
Else
    frmBattle.imgArena.Picture = frmJoin.imgBG.Picture
End If

frmBattle.Show

Call PlayMidi("battle", True)
Call EnableChoose

Exit Sub
err:
Debug.Print "LOADBATTLEERROR - "; err.Description

End Sub
Public Sub SetPowerResist(strCharacter As String, intPlayer As Integer)
On Error Resume Next
'sets power and resist
' E=Earth, F=Fire, N=Wind, W=Water, H=Heart, D=Dark
Select Case strCharacter
Case "Isaac"
    intEarthPower(intPlayer) = 90 + iLvl * 5
    intEarthResist(intPlayer) = 100 + iLvl * 5
    intFirePower(intPlayer) = 80 + iLvl * 5
    intFireResist(intPlayer) = 90 + iLvl * 5
    intWaterPower(intPlayer) = 80 + iLvl * 5
    intWaterResist(intPlayer) = 90 + iLvl * 5
    intWindPower(intPlayer) = 60 + iLvl * 5
    intWindResist(intPlayer) = 65 + iLvl * 5
    intHeartPower(intPlayer) = 80 + iLvl * 5
    intHeartResist(intPlayer) = 90 + iLvl * 5
    intDarkPower(intPlayer) = 80 + iLvl * 5
    intDarkResist(intPlayer) = 90 + iLvl * 5
Case "Young Isaac"
    intEarthPower(intPlayer) = 130 + iLvl * 5
    intEarthResist(intPlayer) = 110 + iLvl * 5
    intFirePower(intPlayer) = 125 + iLvl * 5
    intFireResist(intPlayer) = 100 + iLvl * 5
    intWaterPower(intPlayer) = 90 + iLvl * 5
    intWaterResist(intPlayer) = 100 + iLvl * 5
    intWindPower(intPlayer) = 70 + iLvl * 5
    intWindResist(intPlayer) = 75 + iLvl * 5
    intHeartPower(intPlayer) = 115 + iLvl * 5
    intHeartResist(intPlayer) = 125 + iLvl * 5
    intDarkPower(intPlayer) = 90 + iLvl * 5
    intDarkResist(intPlayer) = 100 + iLvl * 5
Case "KOS"
    intEarthPower(intPlayer) = 120 + iLvl * 5
    intEarthResist(intPlayer) = 190 + iLvl * 5
    intFirePower(intPlayer) = 60 + iLvl * 5
    intFireResist(intPlayer) = 70 + iLvl * 5
    intWaterPower(intPlayer) = 70 + iLvl * 5
    intWaterResist(intPlayer) = 80 + iLvl * 5
    intWindPower(intPlayer) = 80 + iLvl * 5
    intWindResist(intPlayer) = 80 + iLvl * 5
    intHeartPower(intPlayer) = 50 + iLvl * 5
    intHeartResist(intPlayer) = 60 + iLvl * 5
    intDarkPower(intPlayer) = 100 + iLvl * 5
    intDarkResist(intPlayer) = 110 + iLvl * 5
Case "Guard"
    intEarthPower(intPlayer) = 85 + iLvl * 8
    intEarthResist(intPlayer) = 90 + iLvl * 8
    intFirePower(intPlayer) = 55 + iLvl * 8
    intFireResist(intPlayer) = 65 + iLvl * 8
    intWaterPower(intPlayer) = 75 + iLvl * 8
    intWaterResist(intPlayer) = 78 + iLvl * 8
    intWindPower(intPlayer) = 50 + iLvl * 8
    intWindResist(intPlayer) = 60 + iLvl * 8
    intHeartPower(intPlayer) = 72 + iLvl * 8
    intHeartResist(intPlayer) = 75 + iLvl * 8
    intDarkPower(intPlayer) = 73 + iLvl * 8
    intDarkResist(intPlayer) = 76 + iLvl * 8
Case "Gladiator"
    intEarthPower(intPlayer) = 105 + iLvl * 3
    intEarthResist(intPlayer) = 110 + iLvl * 3
    intFirePower(intPlayer) = 85 + iLvl * 3
    intFireResist(intPlayer) = 93 + iLvl * 3
    intWaterPower(intPlayer) = 95 + iLvl * 3
    intWaterResist(intPlayer) = 105 + iLvl * 3
    intWindPower(intPlayer) = 70 + iLvl * 3
    intWindResist(intPlayer) = 80 + iLvl * 3
    intHeartPower(intPlayer) = 85 + iLvl * 3
    intHeartResist(intPlayer) = 90 + iLvl * 3
    intDarkPower(intPlayer) = 90 + iLvl * 3
    intDarkResist(intPlayer) = 95 + iLvl * 3
Case "Garret"
    intEarthPower(intPlayer) = 90 + iLvl * 5
    intEarthResist(intPlayer) = 100 + iLvl * 5
    intFirePower(intPlayer) = 80 + iLvl * 5
    intFireResist(intPlayer) = 90 + iLvl * 5
    intWaterPower(intPlayer) = 70 + iLvl * 5
    intWaterResist(intPlayer) = 80 + iLvl * 5
    intWindPower(intPlayer) = 80 + iLvl * 5
    intWindResist(intPlayer) = 90 + iLvl * 5
    intHeartPower(intPlayer) = 80 + iLvl * 5
    intHeartResist(intPlayer) = 90 + iLvl * 5
    intDarkPower(intPlayer) = 80 + iLvl * 5
    intDarkResist(intPlayer) = 90 + iLvl * 5
Case "Young Garet"
    intEarthPower(intPlayer) = 115 + iLvl * 5
    intEarthResist(intPlayer) = 90 + iLvl * 5
    intFirePower(intPlayer) = 95 + iLvl * 5
    intFireResist(intPlayer) = 80 + iLvl * 5
    intWaterPower(intPlayer) = 85 + iLvl * 5
    intWaterResist(intPlayer) = 70 + iLvl * 5
    intWindPower(intPlayer) = 95 + iLvl * 5
    intWindResist(intPlayer) = 80 + iLvl * 5
    intHeartPower(intPlayer) = 95 + iLvl * 5
    intHeartResist(intPlayer) = 80 + iLvl * 5
    intDarkPower(intPlayer) = 100 + iLvl * 5
    intDarkResist(intPlayer) = 85 + iLvl * 5
Case "Saturos"
    intEarthPower(intPlayer) = 100 + iLvl * 4
    intEarthResist(intPlayer) = 80 + iLvl * 4
    intFirePower(intPlayer) = 105 + iLvl * 4
    intFireResist(intPlayer) = 90 + iLvl * 4
    intWaterPower(intPlayer) = 70 + iLvl * 4
    intWaterResist(intPlayer) = 60 + iLvl * 4
    intWindPower(intPlayer) = 80 + iLvl * 4
    intWindResist(intPlayer) = 90 + iLvl * 4
    intHeartPower(intPlayer) = 90 + iLvl * 4
    intHeartResist(intPlayer) = 80 + iLvl * 4
    intDarkPower(intPlayer) = 95 + iLvl * 4
    intDarkResist(intPlayer) = 85 + iLvl * 4
Case "Menardi"
    intEarthPower(intPlayer) = 80 + iLvl * 4
    intEarthResist(intPlayer) = 85 + iLvl * 4
    intFirePower(intPlayer) = 95 + iLvl * 4
    intFireResist(intPlayer) = 105 + iLvl * 4
    intWaterPower(intPlayer) = 70 + iLvl * 4
    intWaterResist(intPlayer) = 75 + iLvl * 4
    intWindPower(intPlayer) = 75 + iLvl * 4
    intWindResist(intPlayer) = 90 + iLvl * 4
    intHeartPower(intPlayer) = 85 + iLvl * 4
    intHeartResist(intPlayer) = 95 + iLvl * 4
    intDarkPower(intPlayer) = 80 + iLvl * 4
    intDarkResist(intPlayer) = 90 + iLvl * 4
Case "Jenna"
    intEarthPower(intPlayer) = 78 + iLvl * 5.5
    intEarthResist(intPlayer) = 88 + iLvl * 5.5
    intFirePower(intPlayer) = 75 + iLvl * 5.5
    intFireResist(intPlayer) = 85 + iLvl * 5.5
    intWaterPower(intPlayer) = 80 + iLvl * 5.5
    intWaterResist(intPlayer) = 90 + iLvl * 5.5
    intWindPower(intPlayer) = 65 + iLvl * 5.5
    intWindResist(intPlayer) = 75 + iLvl * 5.5
    intHeartPower(intPlayer) = 80 + iLvl * 5.5
    intHeartResist(intPlayer) = 90 + iLvl * 5.5
    intDarkPower(intPlayer) = 72 + iLvl * 5.5
    intDarkResist(intPlayer) = 85 + iLvl * 5.5
Case "Kenny"
    intEarthPower(intPlayer) = 115 + iLvl * 11
    intEarthResist(intPlayer) = 105 + iLvl * 11
    intFirePower(intPlayer) = 125 + iLvl * 11
    intFireResist(intPlayer) = 115 + iLvl * 11
    intWaterPower(intPlayer) = 15 + ivll * 1.5 'Weakness is water
    intWaterResist(intPlayer) = 10 + iLvl * 1.5
    intWindPower(intPlayer) = 115 + iLvl * 11
    intWindResist(intPlayer) = 105 + iLvl * 11
    intHeartPower(intPlayer) = 115 + iLvl * 11
    intHeartResist(intPlayer) = 105 + iLvl * 11
    intDarkPower(intPlayer) = 115 + iLvl * 11
    intDarkResist(intPlayer) = 105 + iLvl * 11



Case "Ivan"
    intEarthPower(intPlayer) = 110 + iLvl * 5
    intEarthResist(intPlayer) = 100 + iLvl * 5
    intFirePower(intPlayer) = 80 + iLvl * 5
    intFireResist(intPlayer) = 90 + iLvl * 5
    intWaterPower(intPlayer) = 80 + iLvl * 5
    intWaterResist(intPlayer) = 90 + iLvl * 5
    intWindPower(intPlayer) = 90 + iLvl * 5
    intWindResist(intPlayer) = 100 + iLvl * 5
    intHeartPower(intPlayer) = 80 + iLvl * 5
    intHeartResist(intPlayer) = 90 + iLvl * 5
    intDarkPower(intPlayer) = 80 + iLvl * 5
    intDarkResist(intPlayer) = 90 + iLvl * 5
Case "Cloud"
    intEarthPower(intPlayer) = 90 + iLvl * 5
    intEarthResist(intPlayer) = 60 + iLvl * 5
    intFirePower(intPlayer) = 100 + iLvl * 5
    intFireResist(intPlayer) = 70 + iLvl * 5
    intWaterPower(intPlayer) = 100 + iLvl * 5
    intWaterResist(intPlayer) = 70 + iLvl * 5
    intWindPower(intPlayer) = 110 + iLvl * 5
    intWindResist(intPlayer) = 90 + iLvl * 5
    intHeartPower(intPlayer) = 100 + iLvl * 5
    intHeartResist(intPlayer) = 70 + iLvl * 5
    intDarkPower(intPlayer) = 100 + iLvl * 5
    intDarkResist(intPlayer) = 70 + iLvl * 5
Case "Sheba"
    intEarthPower(intPlayer) = 100 + iLvl * 5.75
    intEarthResist(intPlayer) = 90 + iLvl * 5.75
    intFirePower(intPlayer) = 70 + iLvl * 5.75
    intFireResist(intPlayer) = 80 + iLvl * 5.75
    intWaterPower(intPlayer) = 120 + iLvl * 5.75
    intWaterResist(intPlayer) = 130 + iLvl * 5.75
    intWindPower(intPlayer) = 110 + iLvl * 5.75
    intWindResist(intPlayer) = 120 + iLvl * 5.75
    intHeartPower(intPlayer) = 90 + iLvl * 5.75
    intHeartResist(intPlayer) = 100 + iLvl * 5.75
    intDarkPower(intPlayer) = 85 + iLvl * 5.75
    intDarkResist(intPlayer) = 95 + iLvl * 5.75

Case "Mia"
    intEarthPower(intPlayer) = 80 + iLvl * 5
    intEarthResist(intPlayer) = 90 + iLvl * 5
    intFirePower(intPlayer) = 70 + iLvl * 5
    intFireResist(intPlayer) = 80 + iLvl * 5
    intWaterPower(intPlayer) = 90 + iLvl * 5
    intWaterResist(intPlayer) = 100 + iLvl * 5
    intWindPower(intPlayer) = 80 + iLvl * 5
    intWindResist(intPlayer) = 90 + iLvl * 5
    intHeartPower(intPlayer) = 80 + iLvl * 5
    intHeartResist(intPlayer) = 90 + iLvl * 5
    intDarkPower(intPlayer) = 80 + iLvl * 5
    intDarkResist(intPlayer) = 90 + iLvl * 5
Case "Piers"
    intEarthPower(intPlayer) = 90 + iLvl * 3.5
    intEarthResist(intPlayer) = 100 + iLvl * 3.5
    intFirePower(intPlayer) = 70 + iLvl * 5
    intFireResist(intPlayer) = 80 + iLvl * 12
    intWaterPower(intPlayer) = 125 + iLvl * 8
    intWaterResist(intPlayer) = 100 + iLvl * 8
    intWindPower(intPlayer) = 67 + iLvl * 3.5
    intWindResist(intPlayer) = 80 + iLvl * 3.5
    intHeartPower(intPlayer) = 91 + iLvl * 3
    intHeartResist(intPlayer) = 99 + iLvl * 3
    intDarkPower(intPlayer) = 105 + iLvl * 5
    intDarkResist(intPlayer) = 95 + iLvl * 5
Case "Purple Piers"
    intEarthPower(intPlayer) = 95 + iLvl * 5
    intEarthResist(intPlayer) = 105 + iLvl * 5
    intFirePower(intPlayer) = 80 + iLvl * 5
    intFireResist(intPlayer) = 90 + iLvl * 12
    intWaterPower(intPlayer) = 135 + iLvl * 10
    intWaterResist(intPlayer) = 100 + iLvl * 10
    intWindPower(intPlayer) = 85 + iLvl * 5
    intWindResist(intPlayer) = 94 + iLvl * 5
    intHeartPower(intPlayer) = 91 + iLvl * 5
    intHeartResist(intPlayer) = 99 + iLvl * 5
    intDarkPower(intPlayer) = 105 + iLvl * 5
    intDarkResist(intPlayer) = 95 + iLvl * 5
Case "Alex"
    intEarthPower(intPlayer) = 70 + iLvl * 4
    intEarthResist(intPlayer) = 80 + iLvl * 4
    intFirePower(intPlayer) = 85 + iLvl * 4
    intFireResist(intPlayer) = 95 + iLvl * 4
    intWaterPower(intPlayer) = 75 + iLvl * 4
    intWaterResist(intPlayer) = 85 + iLvl * 4
    intWindPower(intPlayer) = 77 + iLvl * 4
    intWindResist(intPlayer) = 87 + iLvl * 4
    intHeartPower(intPlayer) = 67 + iLvl * 4
    intHeartResist(intPlayer) = 75 + iLvl * 4
    intDarkPower(intPlayer) = 75 + iLvl * 4
    intDarkResist(intPlayer) = 80 + iLvl * 4

Case "Caption Contest Character"
    intEarthPower(intPlayer) = 90 + iLvl * 1.5
    intEarthResist(intPlayer) = 100 + iLvl * 1.5
    intFirePower(intPlayer) = 80 + iLvl * 1.5
    intFireResist(intPlayer) = 90 + iLvl * 1.5
    intWaterPower(intPlayer) = 75 + iLvl * 1.5
    intWaterResist(intPlayer) = 85 + iLvl * 1.5
    intWindPower(intPlayer) = 87 + iLvl * 1.5
    intWindResist(intPlayer) = 97 + iLvl * 1.5
    intHeartPower(intPlayer) = 76 + iLvl * 1.5
    intHeartResist(intPlayer) = 86 + iLvl * 1.5
    intDarkPower(intPlayer) = 90 + iLvl * 1.5
    intDarkResist(intPlayer) = 105 + iLvl * 1.5

Case "Felix"
    intEarthPower(intPlayer) = 80 + iLvl * 5
    intEarthResist(intPlayer) = 90 + iLvl * 5
    intFirePower(intPlayer) = 80 + iLvl * 5
    intFireResist(intPlayer) = 90 + iLvl * 5
    intWaterPower(intPlayer) = 80 + iLvl * 5
    intWaterResist(intPlayer) = 90 + iLvl * 5
    intWindPower(intPlayer) = 110 + iLvl * 7.5
    intWindResist(intPlayer) = 100 + iLvl * 6
    intHeartPower(intPlayer) = 90 + iLvl * 5
    intHeartResist(intPlayer) = 100 + iLvl * 5
    intDarkPower(intPlayer) = 70 + iLvl * 5
    intDarkResist(intPlayer) = 80 + iLvl * 5

Case "Kraden"
    intEarthPower(intPlayer) = 80 + iLvl * 5
    intEarthResist(intPlayer) = 90 + iLvl * 5
    intFirePower(intPlayer) = 80 + iLvl * 5
    intFireResist(intPlayer) = 90 + iLvl * 5
    intWaterPower(intPlayer) = 80 + iLvl * 5
    intWaterResist(intPlayer) = 90 + iLvl * 5
    intWindPower(intPlayer) = 80 + iLvl * 5
    intWindResist(intPlayer) = 90 + iLvl * 5
    intHeartPower(intPlayer) = 70 + iLvl * 5
    intHeartResist(intPlayer) = 80 + iLvl * 5
    intDarkPower(intPlayer) = 90 + iLvl * 5
    intDarkResist(intPlayer) = 100 + iLvl * 5
Case "Agiato"
    intEarthPower(intPlayer) = 80 + iLvl * 5
    intEarthResist(intPlayer) = 60 + iLvl * 5
    intFirePower(intPlayer) = 100 + iLvl * 5
    intFireResist(intPlayer) = 70 + iLvl * 5
    intWaterPower(intPlayer) = 135 + iLvl * 5
    intWaterResist(intPlayer) = 120 + iLvl * 5
    intWindPower(intPlayer) = 95 + iLvl * 5
    intWindResist(intPlayer) = 80 + iLvl * 5
    intHeartPower(intPlayer) = 85 + iLvl * 5
    intHeartResist(intPlayer) = 75 + iLvl * 5
    intDarkPower(intPlayer) = 80 + iLvl * 5
    intDarkResist(intPlayer) = 75 + iLvl * 5
Case "Karst"
    intEarthPower(intPlayer) = 130 + iLvl * 5
    intEarthResist(intPlayer) = 125 + iLvl * 5
    intFirePower(intPlayer) = 100 + iLvl * 5
    intFireResist(intPlayer) = 73 + iLvl * 5
    intWaterPower(intPlayer) = 100 + iLvl * 5
    intWaterResist(intPlayer) = 92 + iLvl * 5
    intWindPower(intPlayer) = 95 + iLvl * 5
    intWindResist(intPlayer) = 82 + iLvl * 5
    intHeartPower(intPlayer) = 85 + iLvl * 5
    intHeartResist(intPlayer) = 77 + iLvl * 5
    intDarkPower(intPlayer) = 80 + iLvl * 5
    intDarkResist(intPlayer) = 78 + iLvl * 5
Case "The Wise One"
    intEarthPower(intPlayer) = 150 + iLvl * 5
    intEarthResist(intPlayer) = 150 + iLvl * 5
    intFirePower(intPlayer) = 150 + iLvl * 5
    intFireResist(intPlayer) = 150 + iLvl * 5
    intWaterPower(intPlayer) = 150 + iLvl * 5
    intWaterResist(intPlayer) = 150 + iLvl * 5
    intWindPower(intPlayer) = 150 + iLvl * 5
    intWindResist(intPlayer) = 150 + iLvl * 5
    intHeartPower(intPlayer) = 150 + iLvl * 5
    intHeartResist(intPlayer) = 150 + iLvl * 5
    intDarkPower(intPlayer) = 150 + iLvl * 5
    intDarkResist(intPlayer) = 150 + iLvl * 5
End Select

End Sub


Public Function ScrambleText(ByVal strText As String) As String
On Error Resume Next
'This function scrambles a word
Dim intLetter(0 To 50) As Byte
For i = 1 To 50
    intLetter(i) = 0
Next 'i
Dim strWords(0 To 10) As String
Dim intRand As Integer
Dim intScramble
Dim strFinal As String
intScramble = Split(strText, " ", -1, vbTextCompare)
For i = 0 To UBound(intScramble)
    Do Until Len(strWords(i)) = Len(intScramble(i)) 'Do until I've taken each letter
        Randomize
        intRand = Int(Rnd * Len(intScramble(i))) + 1 'Randomly choose a letter from the word
        If intLetter(intRand) = 0 Then 'Have I already picked this letter?
            intLetter(intRand) = 1 'I have now picked this letter
            strWords(i) = strWords(i) & Mid$(intScramble(i), intRand, 1) 'Add the letter to total word
        End If
    Loop
    For q = 1 To 50
        intLetter(q) = 0 'Reset the values
    Next 'q
    If i > 0 Then
        strFinal = strFinal & " " & strWords(i) 'Space out the words
    Else
        strFinal = strWords(i) 'No spaces before the first one
    End If
Next 'i

ScrambleText = strFinal


End Function
Public Function Eyncrypt(sData As String) As String
On Error Resume Next
    Dim sTemp As String, sTemp1 As String
    Dim strBS As String
    Dim strBS2 As String

    For i = 1 To 6
        Randomize
        strBS = strBS & Int(Rnd * 9)
        strBS2 = strBS2 & Int(Rnd * 9)
    Next 'i
    
    sData = strBS & sData & strBS2

    For II% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, II%, 1)
        lT = Asc(sTemp$) * 2
        sTemp1$ = sTemp1$ & Chr(lT)
    Next II%
    Eyncrypt$ = sTemp1$
End Function
Public Sub Encode(ByVal strValue As String, strINIValue As String, strINILength As String, nsave As String)
On Error Resume Next
    Dim strLength As String
    
    strLength = CStr(Len(strValue))
    
    If Len(strValue) < 10 Then
        strLength = "0" & strLength
    End If
    
    strValue = Eyncrypt(strValue)


    
    Dim strLength2 As String
    
    strLength2 = strLength
    strLength = Eyncrypt(strLength2)
    
    Call WriteIni("GEN", strINIValue, strValue, nsave)
    Call WriteIni("GEN", strINILength, strLength, nsave)
    
End Sub
Public Function FindWhichCharacter(ByVal strFindChar As String) As Long
'Finds the number of a custom character from a name
If strFindChar = "" Then
    FindWhichCharacter = 999 'No such character
    Exit Function
End If

For i = 1 To 50
    If strFindChar = CustomChar(i).Name Then
        FindWhichCharacter = i
        Exit Function
    End If
Next 'i

FindWhichCharacter = 999 'No such character
End Function
Public Function GetFullElementalType(ByVal strSmallType As String) As String
'Turns 1 letter abbreviations for elemental types into full strings
Select Case strSmallType
Case "E"
    GetFullElementalType = "Earth"
Case "F"
    GetFullElementalType = "Fire"
Case "N"
    GetFullElementalType = "Wind"
Case "W"
    GetFullElementalType = "Water"
Case "D"
    GetFullElementalType = "Dark"
Case "H"
    GetFullElementalType = "Heart"
End Select

End Function
Public Sub SetCustomPowerResist(ByVal intCustChar As Long)
'Sets the power/resist for custom characters
With CustomChar(intCustChar)
    If .Strength = "E" Then
        intEarthPower(1) = 100 + iLvl * (.BasePower + 2)
        intEarthResist(1) = 100 + iLvl * ((.BaseRes / 2) + 3)
    ElseIf .Weakness = "E" Then
        intEarthPower(1) = 50 + iLvl * (.BasePower + 2)
        intEarthResist(1) = 50 + iLvl * ((.BaseRes / 2) + 3)
    Else
        intEarthPower(1) = 70 + iLvl * (.BasePower + 2)
        intEarthResist(1) = 70 + iLvl * ((.BaseRes / 2) + 3)
    End If
    If .Strength = "F" Then
        intFirePower(1) = 100 + iLvl * (.BasePower + 2)
        intFireResist(1) = 100 + iLvl * ((.BaseRes / 2) + 3)
    ElseIf .Weakness = "F" Then
        intFirePower(1) = 50 + iLvl * (.BasePower + 2)
        intFireResist(1) = 50 + iLvl * ((.BaseRes / 2) + 3)
    Else
        intFirePower(1) = 70 + iLvl * (.BasePower + 2)
        intFireResist(1) = 70 + iLvl * ((.BaseRes / 2) + 3)
    End If
    If .Strength = "N" Then
        intWindPower(1) = 100 + iLvl * (.BasePower + 2)
        intWindResist(1) = 100 + iLvl * ((.BaseRes / 2) + 3)
    ElseIf .Weakness = "N" Then
        intWindPower(1) = 50 + iLvl * (.BasePower + 2)
        intWindResist(1) = 50 + iLvl * ((.BaseRes / 2) + 3)
    Else
        intWindPower(1) = 70 + iLvl * (.BasePower + 2)
        intWindResist(1) = 70 + iLvl * ((.BaseRes / 2) + 3)
    End If
    If .Strength = "W" Then
        intWaterPower(1) = 100 + iLvl * (.BasePower + 2)
        intWaterResist(1) = 100 + iLvl * ((.BaseRes / 2) + 3)
    ElseIf .Weakness = "W" Then
        intWaterPower(1) = 50 + iLvl * (.BasePower + 2)
        intWaterResist(1) = 50 + iLvl * ((.BaseRes / 2) + 3)
    Else
        intWaterPower(1) = 70 + iLvl * (.BasePower + 2)
        intWaterResist(1) = 70 + iLvl * ((.BaseRes / 2) + 3)
    End If
    If .Strength = "D" Then
        intDarkPower(1) = 100 + iLvl * (.BasePower + 2)
        intDarkResist(1) = 100 + iLvl * ((.BaseRes / 2) + 3)
    ElseIf .Weakness = "D" Then
        intDarkPower(1) = 50 + iLvl * (.BasePower + 2)
        intDarkResist(1) = 50 + iLvl * ((.BaseRes / 2) + 3)
    Else
        intDarkPower(1) = 70 + iLvl * (.BasePower + 2)
        intDarkResist(1) = 70 + iLvl * ((.BaseRes / 2) + 3)
    End If
    If .Strength = "H" Then
        intHeartPower(1) = 100 + iLvl * (.BasePower + 2)
        intHeartResist(1) = 100 + iLvl * ((.BaseRes / 2) + 3)
    ElseIf .Weakness = "H" Then
        intHeartPower(1) = 50 + iLvl * (.BasePower + 2)
        intHeartResist(1) = 50 + iLvl * ((.BaseRes / 2) + 3)
    Else
        intHeartPower(1) = 70 + iLvl * (.BasePower + 2)
        intHeartResist(1) = 70 + iLvl * ((.BaseRes / 2) + 3)
    End If
End With
End Sub
Sub CheckTurns()
If PlayerWait = intFirstAttack Then
'The first player has gone, the second player is
'now attacking
    If intFirstAttack = 1 Then
        Call Player2Command
    Else
        Call Player1Command
    End If
ElseIf PlayerWait <> 3 And PlayerWait <> 0 Then
'The second player is gone, wait a few seconds
'before loading the menus again
    PlayerWait = 3
'    frmBattle.txtDisplay.Visible = True
    frmBattle.timeWait.Enabled = True
    Dim intDjinnToReset As Long
    intDjinnToReset = 0
    For i = 1 To 10
        If bWaitToResetDjinn = True Then
            bWaitToResetDjinn = False
            Exit For
        End If
        If strDjinnName(i) <> "" And bDjinnSet(i) = 2 And intDjinnToReset = 0 And i <= RelativeDjinn(1) Then
            bDjinnSet(i) = 0
            frmBattle.txtDisplay.Text = frmBattle.txtDisplay.Text & vbNewLine & strDjinnName(i) & " has been set."
            intDjinnToReset = i
        End If
    Next 'i
'    Call CheckTurns

ElseIf PlayerWait = 0 Then
'No player has gone
    'Call HideList
    'Call DisableChoose
    If intFirstAttack = 1 Then
        Call Player1Command
    Else
        Call Player2Command
    End If
    frmBattle.txtDisplay.Visible = True
    Call AutoScrollTxt(frmBattle.txtDisplay)
ElseIf PlayerWait = 3 Then 'Last turn

    Call EnableChoose
    'Reset the reset variables :)
    Reset(1) = False
    Reset(2) = False
    
    frmBattle.cmdReset.Enabled = True 'Re-enable the reset button
    
    If TimedMatch = True Then
        curCount = 20 'Reset time
        frmBattle.timecount.Enabled = True 'Start counting again
    End If
    
    frmBattle.txtDisplay.Visible = False
    PlayerWait = 0

End If

End Sub
Function AutoScrollTxt(txtbox As TextBox)
'If txtbox.MultiLine = False Then Exit Function
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
Public Function CheckOpponent(strName As String, strIP As String) As Boolean
On Error Resume Next
Dim strMaxOp As String
Dim intMaxOp As Integer
Dim nOp As String
Dim strTempIP As String
Dim strTempName As String

Dim intTotal As Integer
intTotal = 0


nOp = "C:\windows\system32\gsawota.sys"

strMaxOp = GetFromIni(strName, strServerDate, nOp)
If strMaxOp = "5" Then
    CheckOpponent = True
Else
    CheckOpponent = False
End If

End Function
Public Function GetLevel(strR As String) As Long
Dim intTemp As Long
Dim intTotal As Long

intTemp = CLng(strR)
intTemp = intTemp - 1000

For i = 1 To 100
    intTotal = intTotal + (25 + i ^ 1.5)
    If intTotal > intTemp Then
        GetLevel = i
        Exit Function
    End If
Next 'i

'Formula for gaining a level
'Rating To Next Level = 25 + iLvl ^ 1.5


End Function
Public Function GetDjinn(strR As String) As Long
Dim intTemp As Long
Dim intTotal As Long

intTemp = CLng(strR)
intTemp = intTemp - 1000

For i = 1 To 100
    intTotal = intTotal + (25 + i ^ 1.5)
    If intTotal > intTemp Then
        GetDjinn = i / 2
        Exit Function
    End If
Next 'i

End Function
Public Function Decode(sData As String) As String
    Dim sTemp As String, sTemp1 As String


    For II% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, II%, 1)
        lT = Asc(sTemp$) \ 2
        sTemp1$ = sTemp1$ & Chr(lT)
    Next II%
    Decode$ = sTemp1$
End Function
Public Function UltraDecode(strData As String, strLength As String, nsave As String) As String
If strData <> "" Then
Dim iLength As Integer
strData = GetFromIni("GEN", strData, nsave)
strLength = GetFromIni("GEN", strLength, nsave)


strLength = Decode(strLength)
strData = Decode(strData)

iLength = CInt(Mid$(strLength, 7, 2))

UltraDecode = Mid$(strData, 7, iLength)
End If
End Function
Public Sub DecryptString(ByVal nsave As String, ByVal strField As String, ByVal strLength As String)
Dim sLength As String
Dim iLength As Integer
sLength = GetFromIni("GEN", strLength, nsave)
sLength = Decode(sLength)
iLength = CInt(Mid$(sLength, 7, 2))

strField = Mid$(strField, 7, iLength)

MsgBox strField

End Sub
Public Function MazeDecode(strData As String, strLength As String, nsave As String) As String
If strData <> "" Then
Dim iLength As Integer
strData = GetFromIni("GEN", strData, nsave)
strLength = GetFromIni("GEN", strLength, nsave)

strLength = Decode(strLength)
strData = Decode(strData)

iLength = CInt(Mid$(strLength, 2, 2))

MazeDecode = Mid$(strData, 2, iLength)
End If
End Function

Public Sub DoAttacks()
Dim curSpeed As Long
Dim curTurn As Long
curSpeed = 0
For i = 1 To 4
    If Speed(i) = 0 Then Speed(i) = 1
    If Speed(i) > curSpeed And bOReady(i) = True Then
        curSpeed = Speed(i)
        curTurn = i
    End If
Next 'i
If curSpeed = 0 Then
    For i = 1 To 4
        bOReady(i) = False
    Next 'i
    Call EnableChoose
End If

bOReady(curTurn) = False

If HP(Target(curTurn)) < 0 And AttackType(curTurn) <> "REVIVE" Then
    AttackType(curTurn) = "DEFEND"
End If

frmBattle.txtDisplay.Visible = True

Select Case AttackType(curTurn)
    Case "DAMAGE"
        HP(Target(curTurn)) = HP(Target(curTurn)) - AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & " did " & AttackDamage(curTurn) & " damage to " & CharName(Target(curTurn))
    Case "SPECIAL"
        HP(Target(curTurn)) = HP(Target(curTurn)) - AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & " did " & AttackDamage(curTurn) & " damage to " & CharName(Target(curTurn)) & ".  Special Attack!"
    Case "CRITICAL"
        HP(Target(curTurn)) = HP(Target(curTurn)) - AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & " did " & AttackDamage(curTurn) & " damage to " & CharName(Target(curTurn)) & ".  Critical Hit!"
    Case "HEAL"
        HP(Target(curTurn)) = HP(Target(curTurn)) + AttackDamage(curTurn)
        If HP(Target(curTurn)) > MaxHP(Target(curTurn)) Then
            HP(Target(curTurn)) = MaxHP(Target(curTurn))
        End If
        frmBattle.txtDisplay.Text = CharName(curTurn) & " healed " & AttackDamage(curTurn) & " HP to " & CharName(Target(curTurn))
    Case "BOOSTAP"
        AP(Target(curTurn)) = AP(Target(curTurn)) + AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & " boosted AP by " & AttackDamage(curTurn) & " damage to " & CharName(Target(curTurn))
    Case "REDUCEAP"
        AP(Target(curTurn)) = AP(Target(curTurn)) - AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & " reduced " & AttackDamage(curTurn) & " AP to " & CharName(Target(curTurn))
    Case "BOOSTDEFENSE"
        Defense(Target(curTurn)) = Defense(Target(curTurn)) + AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & " boosted defense by " & AttackDamage(curTurn) & " to " & CharName(Target(curTurn))
    Case "REDUCEDEFENSE"
        HP(Target(curTurn)) = Defense(Target(curTurn)) - AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & " reduced defense by " & AttackDamage(curTurn) & " to " & CharName(Target(curTurn))
    Case "BOOSTPP"
        PP(Target(curTurn)) = PP(Target(curTurn)) + AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & " boosted " & AttackDamage(curTurn) & " PP to " & CharName(Target(curTurn))
    Case "REDUCEPP"
        PP(Target(curTurn)) = PP(Target(curTurn)) - AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & " reduced " & AttackDamage(curTurn) & " PP to " & CharName(Target(curTurn))
    Case "BOOSTLUCK"
        Luck(Target(curTurn)) = Luck(Target(curTurn)) + AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & " boosted luck by " & AttackDamage(curTurn) & " to " & CharName(Target(curTurn))
    Case "REDUCELUCK"
        Luck(Target(curTurn)) = Luck(Target(curTurn)) - AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & " reduced luck by " & AttackDamage(curTurn) & " to " & CharName(Target(curTurn))
    Case "BOOSTSPEED"
        Speed(Target(curTurn)) = Speed(Target(curTurn)) - AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & " boosted speed by " & AttackDamage(curTurn) & " to " & CharName(Target(curTurn))
    Case "REDUCESPEED"
        Speed(Target(curTurn)) = Speed(Target(curTurn)) - AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & " reduced speed by " & AttackDamage(curTurn) & " to " & CharName(Target(curTurn))
    Case "SUMMON"
        If Target(curTurn) = 1 Then
            HP(1) = HP(1) - AttackDamage(curTurn)
            HP(2) = HP(2) - (AttackDamage(curTurn) / 2)
            frmBattle.txtDisplay.Text = CharName(curTurn) & "'s summon did " & AttackDamage(1) & " damage to " & CharName(Target(curTurn)) & " and " & AttackDamage(curTurn) / 2 & " damage to " & CharName(2)
        ElseIf Target(curTurn) = 2 Then
            HP(2) = HP(2) - AttackDamage(curTurn)
            HP(1) = HP(1) - (AttackDamage(curTurn) / 2)
            frmBattle.txtDisplay.Text = CharName(curTurn) & "'s summon did " & AttackDamage(2) & " damage to " & CharName(Target(curTurn)) & " and " & AttackDamage(curTurn) / 2 & " damage to " & CharName(1)
        ElseIf Target(curTurn) = 3 Then
            HP(3) = HP(3) - AttackDamage(curTurn)
            HP(4) = HP(4) - (AttackDamage(curTurn) / 2)
            frmBattle.txtDisplay.Text = CharName(curTurn) & "'s summon did " & AttackDamage(3) & " damage to " & CharName(Target(curTurn)) & " and " & AttackDamage(curTurn) / 2 & " damage to " & CharName(4)
        Else
            HP(4) = HP(4) - AttackDamage(curTurn)
            HP(3) = HP(3) - (AttackDamage(curTurn) / 2)
            frmBattle.txtDisplay.Text = CharName(curTurn) & "'s summon did " & AttackDamage(4) & " damage to " & CharName(Target(curTurn)) & " and " & AttackDamage(curTurn) / 2 & " damage to " & CharName(3)
        End If
        
    Case "PSY"
        HP(Target(curTurn)) = HP(Target(curTurn)) - AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & "'s Psynergy did " & AttackDamage(curTurn) & " to " & CharName(Target(curTurn))
    Case "MULTIPSY"
        If Target(curTurn) = 1 Then
            HP(1) = HP(1) - AttackDamage(curTurn)
            HP(2) = HP(2) - (AttackDamage(curTurn) / 2)
        ElseIf Target(curTurn) = 2 Then
            HP(2) = HP(2) - AttackDamage(curTurn)
            HP(1) = HP(1) - (AttackDamage(curTurn) / 2)
        ElseIf Target(curTurn) = 3 Then
            HP(3) = HP(3) - AttackDamage(curTurn)
            HP(4) = HP(4) - (AttackDamage(curTurn) / 2)
        Else
            HP(4) = HP(4) - AttackDamage(curTurn)
            HP(3) = HP(3) - (AttackDamage(curTurn) / 2)
        End If
    Case "DJINNDAMAGE"
        HP(Target(curTurn)) = HP(Target(curTurn)) - AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & "'s Djinn did " & AttackDamage(curTurn) & " to " & CharName(Target(curTurn))
    Case "DJINNHEAL"
        HP(Target(curTurn)) = HP(Target(curTurn)) + AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & "'s Djinn healed " & AttackDamage(curTurn) & " HP to " & CharName(Target(curTurn))
    Case "DJINNBOOSTAP"
        AP(Target(curTurn)) = AP(Target(curTurn)) + AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & "'s Djinn boosted AP by " & AttackDamage(curTurn) & " AP to " & CharName(Target(curTurn))
    Case "DJINNREDUCEAP"
        AP(Target(curTurn)) = AP(Target(curTurn)) - AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & "'s Djinn reduced AP by " & AttackDamage(curTurn) & " AP to " & CharName(Target(curTurn))
    Case "DJINNBOOSTPP"
        PP(Target(curTurn)) = PP(Target(curTurn)) + AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & "'s Djinn boosted PP by " & AttackDamage(curTurn) & " PP to " & CharName(Target(curTurn))
    Case "DJINNREDUCEPP"
        PP(Target(curTurn)) = PP(Target(curTurn)) - AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & "'s Djinn reduced PP by " & AttackDamage(curTurn) & " PP to " & CharName(Target(curTurn))
    Case "DJINNBOOSTDEFENSE"
        Defense(Target(curTurn)) = Defense(Target(curTurn)) + AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & "'s Djinn boosted defense by " & AttackDamage(curTurn) & " to " & CharName(Target(curTurn))
    Case "DJINNREDUCEDEFENSE"
        Defense(Target(curTurn)) = Defense(Target(curTurn)) - AttackDamage(curTurn)
        frmBattle.txtDisplay.Text = CharName(curTurn) & "'s Djinn reduced defense by " & AttackDamage(curTurn) & " to " & CharName(Target(curTurn))
    Case "DEFEND"
        frmBattle.txtDisplay.Text = CharName(curTurn) & " defended."
        'Do nothing for now
End Select

frmBattle.timeWait.Enabled = True

End Sub

Attribute VB_Name = "modWOTAPlus"
'Contains new WOTA Plus variables, code

Public Type NewPsynergy
    Name As String
    Description As String
    Element As String
    Type As String
    Damage As Long
    PP As Long
    EarthDjinn As Long
    FireDjinn As Long
    WindDjinn As Long
    WaterDjinn As Long
    ClassName(1 To 20) As String
    ClassLVL(1 To 20) As Long
    Range As Long
End Type
Public Type NewDjinn
    Name As String
    Element As String
    Type As String
    Damage As Long
    State As Long '0=Set, 1=Standby, 2=Rest
    HP As Long
    PP As Long
    Defense As Long
    AP As Long
    Agility As Long
    Luck As Long
    Description As String
End Type
Public Type NewSummon
    Name As String
    Description As String
    EarthDjinn As Long
    FireDjinn As Long
    WindDjinn As Long
    WaterDjinn As Long
    BaseDamage As Long
    Element As String
End Type
Public Type NewWeapon
    Name As String
    Description As String
    SpecialName As String
    SpecialType As String
    SpecialMod As Long
    SpecialMultMod As Long
    SpecialAddMod As Long
    Damage As Long
    Element As String
    Cost As Long
End Type
Public Type NewArmor
    Name As String
    Description As String
    Defense As Long
    SpecialType As String
    SpecialMod As Long
    Cost As Long
    APBoost As Long
    DefenseBoost As Long
    PPBoost As Long
    HPBoost As Long
    AgilityBoost As Long
    LuckBoost As Long
    PowerBoost As Long
    ResistBoost As Long
End Type
Public Type NewItem
    Name As String
    Description As String
    Type As String
    Damage As Long
    Enabled As Boolean
    Cost As Long
    Special As Long
    Class As String
    AddMod As Long
    ElementalAttack As Boolean
    APBoost As Long
    DefenseBoost As Long
    PPBoost As Long
    HPBoost As Long
    AgilityBoost As Long
    LuckBoost As Long
    PowerBoost As Long
    ResistBoost As Long
    Equipped As Boolean
End Type
Public Type NewClass
    Name As String
    EarthMin As Long
    EarthMax As Long
    EarthLVL As Long
    FireMin As Long
    FireMax As Long
    FireLVL As Long
    WindMin As Long
    WindMax As Long
    WindLVL As Long
    WaterMin As Long
    WaterMax As Long
    WaterLVL As Long
    HPBoost As Long
    PPBoost As Long
    APBoost As Long
    DefenseBoost As Long
    PowerBoost As Long
    ResistBoost As Long
    LuckBoost As Long
    AgilityBoost As Long
    Earth As Boolean
    Fire As Boolean
    Wind As Boolean
    Water As Boolean
    ClassInherit(1 To 10) As String
End Type

Public Type BattleChar
    Name As String
    Picture As String
    Type As String
    MaxHP As Long
    HP As Long
    MaxPP As Long
    PP As Long
    MaxAP As Long
    AP As Long
    MaxDefense As Long
    Defense As Long
    MaxPower As Long
    Power As Long
    MaxResistance As Long
    Resistance As Long
    MaxLuck As Long
    Luck As Long
    MaxAgility As Long
    Agility As Long
    EarthPower As Long
    EarthResist As Long
    FirePower As Long
    FireResist As Long
    WindPower As Long
    WindResist As Long
    WaterPower As Long
    WaterResist As Long
    EarthLevel As Long
    FireLevel As Long
    WindLevel As Long
    WaterLevel As Long
    EarthDjinn As Long
    FireDjinn As Long
    WindDjinn As Long
    WaterDjinn As Long
    ItemNum As Long
    ItemName As String
    DjinnNum(1 To 9) As Long
    DjinnEnabled(1 To 9) As Boolean
    DjinnState(1 To 9) As Long
    Player As Long
    WeaponNum As Long
    WeaponName As String
    ArmorChestNum As Long
    ArmorChestName As String
    ArmorArmNum As Long
    ArmorArmName As String
    ArmorMiscNum As Long
    ArmorMiscName As String
    Item(1 To 3) As NewItem
    Level As Long
    ClassName As String
    ClassNum As Long
    Status As String
    Command As Long 'What command the character is waiting to do
    Target As Long 'Who the character is attacking
    Damage As Long 'How much damage it has done
    DidMove As Boolean 'Whether or not the character has dealt damage yet
    Enabled As Boolean 'Is it actually a character?
    Num As Long 'Which StatCharacter is corresponds to
End Type

Public Type StatBoost
    APBoost As Long
    PPBoost As Long
    HPBoost As Long
    DefenseBoost As Long
    LuckBoost As Long
    AgilityBoost As Long
    PowerBoost As Long
    ResistBoost As Long
    CharBoost(1 To 4) As Boolean
    TurnsLeft As Long
End Type

Public Type StatCharacter
    Name As String
    Picture As String
    Description As String
    Element As String
    Strength As String
    Weakness As String
    HP As Long
    AP As Long
    PP As Long
    Defense As Long
    Luck As Long
    Agility As Long
    Power As Long
    Resist As Long
End Type

Public strMyUserName As String 'Current user name
Public strMyRating As String 'Current user ladder rating
Public strMyRanking As String 'Current user rankings
Public strMyWins As String 'Current user wins
Public strMyLoss As String 'Current user losses
Public strFoeUserName As String 'Current user name
Public strFoeRating As String 'Current user ladder rating
Public strFoeRanking As String 'Current user rankings
Public strFoeWins As String 'Current user wins
Public strFoeLoss As String 'Current user losses

Public bPlayer2Ready As Boolean
Public bPlayer1Ready As Boolean

Public curChar As Long 'Current party member selected
Public curEnemy As Long 'Current enemy member selected

Public Boost(1 To 5) As StatBoost
Public WOTAChar(1 To 8) As BattleChar
Public nPsynergy(1 To 100) As NewPsynergy
Public nDjinn(1 To 72) As NewDjinn
Public nSummon(1 To 25) As NewSummon
Public nWeapon(1 To 100) As NewWeapon
Public nChestArmor(1 To 25) As NewArmor
Public nArmArmor(1 To 25) As NewArmor
Public nMiscArmor(1 To 25) As NewArmor
Public nItem(1 To 100) As NewItem
Public nCharacter(1 To 50) As StatCharacter
Public nClass(1 To 175) As NewClass

Public Sub SendBattleData(strData As String)
On Error Resume Next
If hoston = True Then
    frmHost2.Host.SendData strData
Else
    frmJoin2.Client.SendData strData
End If
End Sub
Public Sub LoadNewBattle()
On Error Resume Next
'Loads character stats before a battle starts

For i = 1 To 8
    WOTAChar(i).AP = 0
    WOTAChar(i).AP = nCharacter(WOTAChar(i).Num).AP
    WOTAChar(i).AP = WOTAChar(i).AP + (WOTAChar(i).Level * 5 * (nCharacter(WOTAChar(i).Num).AP / 100))
    For q = 1 To 3
        If WOTAChar(i).Item(q).Equipped = True Then
            WOTAChar(i).AP = WOTAChar(i).AP + WOTAChar(i).Item(q).APBoost
        End If
    Next 'q
    WOTAChar(i).AP = WOTAChar(i).AP + nChestArmor(WOTAChar(i).ArmorChestNum).APBoost
    WOTAChar(i).AP = WOTAChar(i).AP + nArmArmor(WOTAChar(i).ArmorArmNum).APBoost
    WOTAChar(i).AP = WOTAChar(i).AP + nMiscArmor(WOTAChar(i).ArmorMiscNum).APBoost
    WOTAChar(i).MaxAP = WOTAChar(i).AP
    
    WOTAChar(i).HP = 0
    WOTAChar(i).HP = nCharacter(WOTAChar(i).Num).HP
    WOTAChar(i).HP = WOTAChar(i).HP + (WOTAChar(i).Level * 5 * (nCharacter(WOTAChar(i).Num).HP / 100))

    For q = 1 To 3
        If WOTAChar(i).Item(q).Equipped = True Then
            WOTAChar(i).HP = WOTAChar(i).HP + WOTAChar(i).Item(q).HPBoost
        End If
    Next 'q
    WOTAChar(i).HP = WOTAChar(i).HP + nChestArmor(WOTAChar(i).ArmorChestNum).HPBoost
    WOTAChar(i).HP = WOTAChar(i).HP + nArmArmor(WOTAChar(i).ArmorArmNum).HPBoost
    WOTAChar(i).HP = WOTAChar(i).HP + nMiscArmor(WOTAChar(i).ArmorMiscNum).HPBoost
    WOTAChar(i).MaxHP = WOTAChar(i).HP
    
    WOTAChar(i).PP = 0
    WOTAChar(i).PP = nCharacter(WOTAChar(i).Num).PP
    WOTAChar(i).PP = WOTAChar(i).PP + (WOTAChar(i).Level * 5 * (nCharacter(WOTAChar(i).Num).PP / 100))
    For q = 1 To 3
        If WOTAChar(i).Item(q).Equipped = True Then
            WOTAChar(i).PP = WOTAChar(i).PP + WOTAChar(i).Item(q).PPBoost
        End If
    Next 'q
    WOTAChar(i).PP = WOTAChar(i).PP + nChestArmor(WOTAChar(i).ArmorChestNum).PPBoost
    WOTAChar(i).PP = WOTAChar(i).PP + nArmArmor(WOTAChar(i).ArmorArmNum).PPBoost
    WOTAChar(i).PP = WOTAChar(i).PP + nMiscArmor(WOTAChar(i).ArmorMiscNum).PPBoost
    WOTAChar(i).MaxPP = WOTAChar(i).PP
    
    WOTAChar(i).Luck = 0
    WOTAChar(i).Luck = nCharacter(WOTAChar(i).Num).Luck
    WOTAChar(i).Luck = WOTAChar(i).Luck + (WOTAChar(i).Level * 5 * (nCharacter(WOTAChar(i).Num).Luck / 100))
    For q = 1 To 3
        If WOTAChar(i).Item(q).Equipped = True Then
            WOTAChar(i).Luck = WOTAChar(i).Luck + WOTAChar(i).Item(q).LuckBoost
        End If
    Next 'q
    WOTAChar(i).Luck = WOTAChar(i).Luck + nChestArmor(WOTAChar(i).ArmorChestNum).LuckBoost
    WOTAChar(i).Luck = WOTAChar(i).Luck + nArmArmor(WOTAChar(i).ArmorArmNum).LuckBoost
    WOTAChar(i).Luck = WOTAChar(i).Luck + nMiscArmor(WOTAChar(i).ArmorMiscNum).LuckBoost
    WOTAChar(i).MaxLuck = WOTAChar(i).Luck
    
    WOTAChar(i).Agility = 0
    WOTAChar(i).Agility = nCharacter(WOTAChar(i).Num).Agility
    WOTAChar(i).Agility = WOTAChar(i).Agility + (WOTAChar(i).Level * 5 * (nCharacter(WOTAChar(i).Num).Agility / 100))

    For q = 1 To 3
        If WOTAChar(i).Item(q).Equipped = True Then
            WOTAChar(i).Agility = WOTAChar(i).Agility + WOTAChar(i).Item(q).AgilityBoost
        End If
    Next 'q
    WOTAChar(i).Agility = WOTAChar(i).Agility + nChestArmor(WOTAChar(i).ArmorChestNum).AgilityBoost
    WOTAChar(i).Agility = WOTAChar(i).Agility + nArmArmor(WOTAChar(i).ArmorArmNum).AgilityBoost
    WOTAChar(i).Agility = WOTAChar(i).Agility + nMiscArmor(WOTAChar(i).ArmorMiscNum).AgilityBoost
    WOTAChar(i).MaxAgility = WOTAChar(i).Agility
    
    WOTAChar(i).Defense = 0
    WOTAChar(i).Defense = nCharacter(WOTAChar(i).Num).Defense
    WOTAChar(i).Defense = WOTAChar(i).Defense + (WOTAChar(i).Level * 5 * (nCharacter(WOTAChar(i).Num).Defense / 100))

    For q = 1 To 3
        If WOTAChar(i).Item(q).Equipped = True Then
            WOTAChar(i).Defense = WOTAChar(i).Defense + WOTAChar(i).Item(q).DefenseBoost
        End If
    Next 'q
    WOTAChar(i).Defense = WOTAChar(i).Defense + nChestArmor(WOTAChar(i).ArmorChestNum).DefenseBoost
    WOTAChar(i).Defense = WOTAChar(i).Defense + nArmArmor(WOTAChar(i).ArmorArmNum).DefenseBoost
    WOTAChar(i).Defense = WOTAChar(i).Defense + nMiscArmor(WOTAChar(i).ArmorMiscNum).DefenseBoost
    WOTAChar(i).MaxDefense = WOTAChar(i).Defense
    
    Call frmBattle2.GetDjinn(CLng(i))
    Call frmBattle2.GetClass(CLng(i))

Next 'i

For i = 0 To 3
    frmBattle2.lblChar(i).Caption = WOTAChar(i + 1).Name
    frmBattle2.lblHP(i).Caption = WOTAChar(i + 1).HP
    frmBattle2.lblPP(i).Caption = WOTAChar(i + 1).PP
    frmBattle2.shpHP(i).Width = 121
    frmBattle2.shpPP(i).Width = 121
    frmBattle2.Enemy(i + 4).Picture = LoadPicture(App.Path & "\BattleImages\" & WOTAChar(i + 1).Name & "B.gif")
Next 'i


End Sub
Public Function nEyncrypt(sData As String) As String
On Error Resume Next
    Dim sTemp As String, sTemp1 As String
    Dim strBS As String
    Dim strBS2 As String
    Dim intBS As String
    Dim intBS2 As String
    strBS = ""
    strBS2 = ""
    For i = 1 To 4
        intBS = Chr(Int(Rnd * 74) + 48)
        intBS2 = Chr(Int(Rnd * 74) + 48)
        strBS = strBS & CStr(intBS)
        strBS2 = strBS2 & CStr(intBS2)
    Next 'i
    
    sData = strBS & sData & strBS2

    For iI% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, iI%, 1)
        lT = Asc(sTemp$) * 2
        sTemp1$ = sTemp1$ & Chr(lT)
    Next iI%
    nEyncrypt$ = sTemp1$
End Function
Public Sub nEncode(strValue As String, strINIValue As String, strINILength As String, nsave As String)
On Error Resume Next
    Dim strLength As String
    
    strLength = CStr(Len(strValue))
    
    If Len(strValue) < 10 Then
        strLength = "0" & strLength
    End If
    
    strValue = nEyncrypt(strValue)


    
    Dim strLength2 As String
    
    strLength2 = strLength
    strLength = nEyncrypt(strLength2)
    
    Call WriteIni("GEN", strINIValue, strValue, nsave)
    Call WriteIni("GEN", strINILength, strLength, nsave)
    
End Sub
Public Function nDecrypt(sData As String) As String
    Dim sTemp As String, sTemp1 As String


    For iI% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, iI%, 1)
        lT = Asc(sTemp$) \ 2
        sTemp1$ = sTemp1$ & Chr(lT)
    Next iI%
    nDecrypt$ = sTemp1$
End Function
Public Function nDecode(strData As String, strLength As String, eSave As String) As String
On Error Resume Next
If strData <> "" Then
Dim iLength As Integer
strData = GetFromIni("GEN", strData, eSave)
strLength = GetFromIni("GEN", strLength, eSave)


strLength = nDecrypt(strLength)
strData = nDecrypt(strData)

iLength = CInt(Mid$(strLength, 5, 2))

nDecode = Mid$(strData, 5, iLength)
End If

End Function
Public Sub DoCommands()
On Error Resume Next
Dim intLowestChar As Long
Dim intLowestAgility As Long
intLowestChar = 1
intLowestAgility = 0
For i = 1 To 8
    If WOTAChar(i).Agility >= intLowestAgility And WOTAChar(i).DidMove = False Then
        intLowestAgility = WOTAChar(i).Agility
        intLowestChar = i
    End If
Next 'i
If intLowestAgility = 0 Then 'Everyone has gone
    Call frmBattle2.Reset
    Exit Sub
End If
With WOTAChar(intLowestChar)
    If .Command = "ATTACK" Or .Command = "ATTACKS" Or .Command = "ATTACKC" Or .Command = "PSY" Or .Command = "DJINNA" Then
        WOTAChar(.Target).HP = WOTAChar(.Target).HP - .Damage
    ElseIf .Command = "PSYH" Then
        WOTAChar(.Target).HP = WOTAChar(.Target).HP + .Damage
    End If
    Select Case .Command
        Case "ATTACK"
            Call AddBattleText(.Name & "'s attack did " & .Damage & " damage to " & WOTAChar(.Target).Name)
        Case "ATTACKS"
            Call AddBattleText(.Name & "'s weapon unleashed " & nWeapon(.WeaponNum).SpecialName & " for " & .Damage & " damage on" & WOTAChar(.Target).Name & "!")
        Case "ATTACKC"
            Call AddBattleText(.Name & " unleashed a critical hit on " & WOTAChar(.Target).Name & " for " & .Damage & " damage!")
    End Select
End With

WOTAChar(i).DidMove = True
Call DoCommands

End Sub
Sub AddBattleText(strText As String)
On Error Resume Next
'Adds text
frmBattle2.txtChat.Text = frmBattle2.txtChat.Text & vbNewLine & strText

End Sub

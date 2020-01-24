VERSION 5.00
Begin VB.Form frmIntro2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Golden Sun: The War of the Adepts (Plus)"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7320
   Icon            =   "frmIntro2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   504
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   488
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framMain 
      BackColor       =   &H00000000&
      Caption         =   "Welcome To War of the Adepts"
      ForeColor       =   &H00FFFFFF&
      Height          =   7575
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   7335
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "WOTA+ Version 0.1 Beta"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   11
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label lblGen 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Battle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   7
         Left            =   1800
         TabIndex        =   10
         Top             =   6360
         Width           =   3720
      End
      Begin VB.Label lblGen 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   6
         Left            =   1800
         TabIndex        =   9
         Top             =   5880
         Width           =   3720
      End
      Begin VB.Label lblGen 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Credits"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   5
         Left            =   1800
         TabIndex        =   8
         Top             =   5400
         Width           =   3720
      End
      Begin VB.Label lblGen 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Code Entry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   4
         Left            =   1800
         TabIndex        =   7
         Top             =   4920
         Width           =   3720
      End
      Begin VB.Label lblGen 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Game Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   3
         Left            =   1800
         TabIndex        =   6
         Top             =   4440
         Width           =   3720
      End
      Begin VB.Label lblGen 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   2
         Left            =   1800
         TabIndex        =   5
         Top             =   3960
         Width           =   3720
      End
      Begin VB.Label lblGen 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   1
         Left            =   1800
         TabIndex        =   4
         Top             =   3480
         Width           =   3720
      End
      Begin VB.Label lblGen 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Log In"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   0
         Left            =   1800
         TabIndex        =   3
         Top             =   3000
         Width           =   3720
      End
      Begin VB.Image imgTitle 
         Height          =   2625
         Left            =   120
         Picture         =   "frmIntro2.frx":08CA
         Top             =   240
         Width           =   7125
      End
   End
   Begin VB.Label lblLoadingMsg 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmIntro2.frx":13802
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   6975
   End
   Begin VB.Image imgGoldenSun 
      Height          =   480
      Left            =   6960
      Picture         =   "frmIntro2.frx":138BC
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image imgLoading 
      Height          =   480
      Left            =   0
      Picture         =   "frmIntro2.frx":14586
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label lblLoading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "War of the Adepts In Currently Loading"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2280
      TabIndex        =   0
      Top             =   3360
      Width           =   2745
   End
End
Attribute VB_Name = "frmIntro2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
On Error Resume Next

If nClass(1).Name <> "" Then Exit Sub

'Load the Class files
Dim nsave As String
nsave = App.Path & "\Class.ini"
Dim strMax As String
Dim intMax As Long
strMax = GetFromIni("GEN", "MAX", nsave)
If strMax = "" Then strMax = "1"
intMax = CInt(strMax)
Dim strTempClass As String


Dim strTemp As String

For i = 1 To intMax
    strTempClass = CStr(i)
    nClass(i).Name = nDecode("NAME" & strTempClass, "NAMEL" & strTempClass, nsave)
    nClass(i).EarthMin = nDecode(strTempClass & "DJINNMIN0", strTempClass & "DJINNMINL0", nsave)
    nClass(i).FireMin = nDecode(strTempClass & "DJINNMIN1", strTempClass & "DJINNMINL1", nsave)
    nClass(i).WindMin = nDecode(strTempClass & "DJINNMIN2", strTempClass & "DJINNMINL2", nsave)
    nClass(i).WaterMin = nDecode(strTempClass & "DJINNMIN3", strTempClass & "DJINNMINL3", nsave)
    nClass(i).EarthMax = nDecode(strTempClass & "DJINNMAX0", strTempClass & "DJINNMAXL0", nsave)
    nClass(i).FireMax = nDecode(strTempClass & "DJINNMAX1", strTempClass & "DJINNMAXL1", nsave)
    nClass(i).WindMax = nDecode(strTempClass & "DJINNMAX2", strTempClass & "DJINNMAXL2", nsave)
    nClass(i).WaterMax = nDecode(strTempClass & "DJINNMAX3", strTempClass & "DJINNMAXL3", nsave)
    nClass(i).EarthLVL = nDecode(strTempClass & "LVL0", strTempClass & "LVLL0", nsave)
    nClass(i).FireLVL = nDecode(strTempClass & "LVL1", strTempClass & "LVLL1", nsave)
    nClass(i).WindLVL = nDecode(strTempClass & "LVL2", strTempClass & "LVLL2", nsave)
    nClass(i).WaterLVL = nDecode(strTempClass & "LVL3", strTempClass & "LVLL3", nsave)
    nClass(i).EarthLVL = nDecode(strTempClass & "LVL0", strTempClass & "LVLL0", nsave)
    strTemp = nDecode(strTempClass & "ELEMENT0", strTempClass & "ELEMENTL0", nsave)
    If strTemp = "1" Then
        nClass(i).Earth = True
    Else
        nClass(i).Earth = False
    End If
    strTemp = nDecode(strTempClass & "ELEMENT1", strTempClass & "ELEMENTL1", nsave)
    If strTemp = "1" Then
        nClass(i).Fire = True
    Else
        nClass(i).Fire = False
    End If
    strTemp = nDecode(strTempClass & "ELEMENT2", strTempClass & "ELEMENTL2", nsave)
    If strTemp = "1" Then
        nClass(i).Wind = True
    Else
        nClass(i).Wind = False
    End If
    strTemp = nDecode(strTempClass & "ELEMENT3", strTempClass & "ELEMENTL3", nsave)
    If strTemp = "1" Then
        nClass(i).Water = True
    Else
        nClass(i).Water = False
    End If
    nClass(i).HPBoost = CLng(nDecode(strTempClass & "STAT0", strTempClass & "STATL0", nsave))
    nClass(i).APBoost = CLng(nDecode(strTempClass & "STAT1", strTempClass & "STATL1", nsave))
    nClass(i).PPBoost = CLng(nDecode(strTempClass & "STAT2", strTempClass & "STATL2", nsave))
    nClass(i).DefenseBoost = CLng(nDecode(strTempClass & "STAT3", strTempClass & "STATL3", nsave))
    nClass(i).LuckBoost = CLng(nDecode(strTempClass & "STAT4", strTempClass & "STATL4", nsave))
    nClass(i).AgilityBoost = CLng(nDecode(strTempClass & "STAT5", strTempClass & "STATL5", nsave))
    
    Dim strMaxClassI As String
    Dim intMaxClassI As Long
    strMaxClassI = nDecode("IMAX" & strTempClass, "IMAXL" & strTempClass, nsave)
    If strMaxClassI <> "" Then
        intMaxClassI = CLng(strMaxClassI)
        For q = 1 To intMaxClassI
            nClass(i).ClassInherit(q) = nDecode(strTempClass & "ICLASS" & CStr(q), strTempClass & "ICLASSL" & CStr(q), nsave)
        Next 'q
    End If
    imgLoading.Left = imgLoading.Left + (Me.ScaleWidth / intMax) - 0.5
    DoEvents
Next 'i

nsave = App.Path & "\Djinn.ini"
strMax = GetFromIni("GEN", "MAX", nsave)
If strMax = "" Then strMax = "1"
intMax = CInt(strMax)
For i = 1 To intMax
    nDjinn(i).Name = nDecode("NAME" & CStr(i), "NAMEL" & CStr(i), nsave)
    nDjinn(i).Description = nDecode("DESC" & CStr(i), "DESCL" & CStr(i), nsave)
    nDjinn(i).Element = nDecode("TYPE" & CStr(i), "TYPEL" & CStr(i), nsave)
    nDjinn(i).Type = nDecode("ATYPE" & CStr(i), "ATYPEL" & CStr(i), nsave)
    nDjinn(i).Damage = CLng(nDecode("DAMAGE" & CStr(i), "DAMAGEL" & CStr(i), nsave))
    nDjinn(i).HP = nDecode(CStr(i) & "STAT0", CStr(i) & "STAT0", nsave)
    nDjinn(i).PP = nDecode(CStr(i) & "STAT0", CStr(i) & "STAT1", nsave)
    nDjinn(i).AP = nDecode(CStr(i) & "STAT0", CStr(i) & "STAT2", nsave)
    nDjinn(i).Defense = nDecode(CStr(i) & "STAT0", CStr(i) & "STAT3", nsave)
    nDjinn(i).Luck = nDecode(CStr(i) & "STAT0", CStr(i) & "STAT4", nsave)
    nDjinn(i).Agility = nDecode(CStr(i) & "STAT0", CStr(i) & "STAT5", nsave)
Next 'i

nsave = App.Path & "\Character.ini"
strMax = GetFromIni("GEN", "MAX", nsave)
If strMax = "" Then strMax = "1"
intMax = CInt(strMax)
For i = 1 To intMax
    nCharacter(i).Name = nDecode("NAME" & CStr(i), "NAMEL" & CStr(i), nsave)
    nCharacter(i).Description = GetFromIni("GEN", "DESC" & CStr(i), nsave)
    nCharacter(i).Picture = nDecode("PIC" & CStr(i), "PICL" & CStr(i), nsave)
    nCharacter(i).Element = nDecode("TYPE" & CStr(i), "TYPEL" & CStr(i), nsave)
    nCharacter(i).Strength = nDecode("STRENGTH" & CStr(i), "STRENGTHL" & CStr(i), nsave)
    nCharacter(i).Weakness = nDecode("WEAKNESS" & CStr(i), "WEAKNESSL" & CStr(i), nsave)
    nCharacter(i).HP = CLng(nDecode(CStr(i) & "STAT0", CStr(i) & "STATL0", nsave))
    nCharacter(i).AP = CLng(nDecode(CStr(i) & "STAT1", CStr(i) & "STATL1", nsave))
    nCharacter(i).PP = CLng(nDecode(CStr(i) & "STAT2", CStr(i) & "STATL2", nsave))
    nCharacter(i).Defense = CLng(nDecode(CStr(i) & "STAT3", CStr(i) & "STATL3", nsave))
    nCharacter(i).Power = CLng(nDecode(CStr(i) & "STAT4", CStr(i) & "STATL4", nsave))
    nCharacter(i).Resist = CLng(nDecode(CStr(i) & "STAT5", CStr(i) & "STATL5", nsave))
    nCharacter(i).Luck = CLng(nDecode(CStr(i) & "STAT6", CStr(i) & "STATL6", nsave))
    nCharacter(i).Agility = CLng(nDecode(CStr(i) & "STAT7", CStr(i) & "STATL7", nsave))
Next 'i

nsave = App.Path & "\Psynergy.ini"
strMax = GetFromIni("GEN", "MAX", nsave)
If strMax = "" Then strMax = "1"
intMax = CInt(strMax)
For i = 1 To intMax
    nPsynergy(i).Name = nDecode("NAME" & CStr(i), "NAMEL" & CStr(i), nsave)
    nPsynergy(i).PP = CLng(nDecode("PP" & CStr(i), "PPL" & CStr(i), nsave))
    nPsynergy(i).Description = nDecode("DESC" & CStr(i), "DESCL" & CStr(i), nsave)
    nPsynergy(i).Type = nDecode("TYPE" & CStr(i), "TYPEL" & CStr(i), nsave)
    nPsynergy(i).Range = CLng(nDecode("RANGE" & CStr(i), "RANGEL" & CStr(i), nsave))
    nPsynergy(i).Damage = CLng(nDecode("DAMAGE" & CStr(i), "DAMAGEL" & CStr(i), nsave))
    Dim intTempMax As Long
    intTempMax = CLng(nDecode("MAXLVL" & CStr(i), "MAXLVLL" & CStr(i), nsave))
    For q = 0 To intTempMax - 1
        nPsynergy(i).ClassName(q + 1) = nDecode(CStr(i) & "CLASS" & CStr(q), CStr(i) & "CLASSL" & CStr(q), nsave)
        nPsynergy(i).ClassLVL(q + 1) = CLng(nDecode(CStr(i) & "CLASSLVL" & CStr(q), CStr(i) & "CLASSLVLL" & CStr(q), nsave))
    Next 'q
    Dim strTempPsyEl As String
    strTempPsyEl = nDecode("ELEMENT" & CStr(i), "ELEMENTL" & CStr(i), nsave)
    If strTempPsyEl = "0" Then
        nPsynergy(i).Element = "Earth"
    ElseIf strTempPsyEl = "1" Then
        nPsynergy(i).Element = "Fire"
    ElseIf strTempPsyEl = "2" Then
        nPsynergy(i).Element = "Wind"
    ElseIf strTempPsyEl = "3" Then
        nPsynergy(i).Element = "Water"
    End If
Next 'i

framMain.Visible = True




End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lblGen_Click(Index As Integer)
If Index = 0 Then
    frmUser2.Show
    frmUser2.framLogIn.Visible = True
    frmUser2.framTOS.Visible = True
    frmUser2.framLStats.Visible = True
    frmUser2.framNewUser.Visible = False
    frmUser2.framStats.Visible = False
End If
If Index = 1 Then
    frmUser2.Show
    frmUser2.framLogIn.Visible = False
    frmUser2.framTOS.Visible = False
    frmUser2.framLStats.Visible = False
    frmUser2.framNewUser.Visible = True
    frmUser2.framStats.Visible = True
End If
If Index = 6 Then
    Unload Me
End If
If Index = 7 Then
    frmBattle2.Show
End If
End Sub

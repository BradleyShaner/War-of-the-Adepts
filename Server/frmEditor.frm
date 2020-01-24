VERSION 5.00
Begin VB.Form frmEditor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0039F798&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "War of the Adepts Custom Quest Editor"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Show Options"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lblCurTile 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Tile:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5280
      Width           =   855
   End
   Begin VB.Image imgSprite 
      Height          =   495
      Index           =   0
      Left            =   1800
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgTile 
      Height          =   375
      Index           =   0
      Left            =   1080
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOptions_Click()
frmOptions.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 And Shift = 1 Then
    yadda = InputBox("Enter Password")
    If yadda = "parboil" Then
        For i = 0 To frmOptions.imgTile.UBound
            frmOptions.imgTile(i).Visible = True
        Next 'i
        For i = 0 To frmOptions.imgSprite.UBound
            frmOptions.imgSprite(i).Visible = True
        Next 'i
        frmOptions.framMaze.Visible = True
        frmOptions.mnuOpen.Enabled = True
        frmOptions.mnuSave.Enabled = True
        frmQuest.txtCoins.Enabled = True

        Admin = True
        
    End If
End If
End Sub

Private Sub Form_Load()
Admin = True
    frmOptions.Show
    Dim row As Integer
    Dim col As Integer
    col = 0
    row = 0
    For i = 1 To 280
    

    
        Load imgTile(i)
        imgTile(i).Picture = frmOptions.imgTile(12).Picture
        imgTile(i).Left = 25 * col
        imgTile(i).Top = 25 * row
        imgTile(i).Visible = True
        col = col + 1
        
        If col = 20 Then
            col = 0
            row = row + 1
        End If
        
        If i <= 25 Then
            Load imgSprite(i)
            imgSprite(i).Visible = False
            SpriteType(i) = 999
        End If
        
        TileType(i) = 12
    Next 'i
    CurTile = 1

        For i = 0 To frmOptions.imgTile.UBound
            frmOptions.imgTile(i).Visible = True
        Next 'i
        For i = 0 To frmOptions.imgSprite.UBound
            frmOptions.imgSprite(i).Visible = True
        Next 'i
        frmOptions.framMaze.Visible = True
        frmOptions.mnuOpen.Enabled = True
        frmOptions.mnuSave.Enabled = True
        frmQuest.txtCoins.Enabled = True
        frmOptions.mnuSaveMazeOff.Enabled = True
        frmOptions.mnuOpenMazeOff.Enabled = True
        Admin = True


End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub imgSprite_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
imgSprite(Index).Visible = False
SpriteType(Index) = 999
End If
If Button = 1 Then
    X = X / Screen.TwipsPerPixelX
    Y = Y / Screen.TwipsPerPixelY
    Dim i As Integer
    For i = 1 To 280
        If imgSprite(Index).Left + X <= imgTile(i).Left + imgTile(i).Width And imgSprite(Index).Left + X >= imgTile(i).Left And imgSprite(Index).Top + Y <= imgTile(i).Top + imgTile(i).Height And imgSprite(Index).Top + Y >= imgTile(i).Top Then
            Call imgTile_Click(i)
        End If
    Next 'i
End If
End Sub

Private Sub imgTile_Click(Index As Integer)


    

If curType = 0 Then
imgTile(Index).Picture = frmOptions.imgTile(CurTile).Picture
TileType(Index) = CurTile

    If CurTile = 31 Then
        imgTile(Index).BorderStyle = 1
    Else
        imgTile(Index).BorderStyle = 0
    End If
    
    If frmOptions.chkGoTo.Value = 1 Then
        TileGoto(Index) = True
        TileLink(Index) = frmOptions.txtLevel.Text
        imgTile(Index).BorderStyle = 1
        TileGotoLink(Index) = frmOptions.txtGoto.Text
    Else
        TileGoto(Index) = False
        If CurTile <> 31 Then
            imgTile(Index).BorderStyle = 0
        End If
    End If

    If frmOptions.chkRandom.Enabled = True Then
        TileRandom(Index) = True
    Else
        TileRandom(Index) = False
    End If
'    If frmTalk.chkTalk.Value = 1 Then
'        TileTalk(Index) = frmTalk.txtTalk.Text
'    Else
'        TileTalk(Index) = ""
'    End If
    If frmOptions.chkBoss.Value = 1 Then
        TileBoss(Index) = True
    Else
        TileBoss(Index) = False
    End If
    If frmOptions.chkJump.Value = 1 Then
        TileJumpable(Index) = True
    Else
        TileJumpable(Index) = False
    End If
    If frmOptions.chkMove.Value = 1 Then
        TileMovable(Index) = True
    Else
        TileMovable(Index) = False
    End If
    If frmOptions.chkEndTile.Value = 0 Then
        bMazeFinish(Index) = 0
        imgTile(Index).BorderStyle = 0
    Else
        bMazeFinish(Index) = 1
        imgTile(Index).BorderStyle = 1
    End If
End If
If curType = 1 Then
    Dim SpritetoUse As Integer
    SpritetoUse = 999
    For i = 1 To 25
        If SpriteType(i) = 999 And SpritetoUse = 999 Then
            SpritetoUse = i
        End If
    Next 'i
    SpriteType(SpritetoUse) = CurTile
    imgSprite(SpritetoUse).Visible = True
    imgSprite(SpritetoUse).Picture = frmOptions.imgSprite(CurTile).Picture
    imgSprite(SpritetoUse).Left = imgTile(Index).Left
    imgSprite(SpritetoUse).Top = imgTile(Index).Top
    Debug.Print imgSprite(SpritetoUse).Top
    
    If CurTile = 17 Then
        Dim col As Integer
        Dim row As Integer
        Dim realCol As Integer
        col = 0
        row = 0
        
        For i = 1 To 100
            realCol = Index + (20 * (row - 1) + (col - 1))
            
            If (col = 7 And row = 3) Or (col >= 5 And col <= 9 And row = 4) Or (col >= 2 And col <= 5 And row = 5) Or (col = 9 And row = 5) Or (col >= 2 And col <= 5 And row = 6) Or (col = 9 And row = 6) Or (col >= 2 And col <= 9 And row = 7) Or (col >= 2 And col <= 9 And row = 8) Or (col = 8 And row = 9) Or (col = 8 And row = 10) Then
                imgTile(realCol).Picture = frmOptions.imgTile(41).Picture
                TileType(realCol) = 41
            End If
            
            col = col + 1
            If col = 11 Then
                col = 1
                row = row + 1
            End If
            
        Next 'i
    End If
End If
End Sub

Private Sub imgTile_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
If curType = 0 Then
imgTile(Index).Picture = frmOptions.imgTile(CurTile).Picture
TileType(Index) = CurTile

    If CurTile = 31 Then
        imgTile(Index).BorderStyle = 1
    Else
        imgTile(Index).BorderStyle = 0
    End If
    
    If frmOptions.chkGoTo.Value = 1 Then
        TileGoto(Index) = True
        TileLink(Index) = frmOptions.txtLevel.Text
        imgTile(Index).BorderStyle = 1
        TileGotoLink(Index) = frmOptions.txtGoto.Text
    Else
        TileGoto(Index) = False
        If CurTile <> 31 Then
            imgTile(Index).BorderStyle = 0
        End If
    End If

    If frmOptions.chkRandom.Enabled = True Then
        TileRandom(Index) = True
    Else
        TileRandom(Index) = False
    End If
'    If frmTalk.chkTalk.Value = 1 Then
'        TileTalk(Index) = frmTalk.txtTalk.Text
'    Else
'        TileTalk(Index) = ""
'    End If
    If frmOptions.chkBoss.Value = 1 Then
        TileBoss(Index) = True
    Else
        TileBoss(Index) = False
    End If
        If frmOptions.chkJump.Value = 1 Then
        TileJumpable(Index) = True
    Else
        TileJumpable(Index) = False
    End If
    If frmOptions.chkMove.Value = 1 Then
        TileMovable(Index) = True
    Else
        TileMovable(Index) = False
    End If
    If frmOptions.chkEndTile.Value = 0 Then
        bMazeFinish(Index) = 0
        imgTile(Index).BorderStyle = 0
    Else
        bMazeFinish(Index) = 1
        imgTile(Index).BorderStyle = 1
    End If
    
End If

'If curType = 1 Then
'    Dim SpritetoUse As Integer
'    SpritetoUse = 999
'    For i = 1 To 25
'        If SpriteType(i) = 999 And SpritetoUse = 999 Then
'            SpritetoUse = i
'        End If
'    Next 'i
'    SpriteType(SpritetoUse) = CurTile
'    imgSprite(SpritetoUse).Visible = True
'    imgSprite(SpritetoUse).Picture = frmEditor.imgSprite(CurTile).Picture
'    imgSprite(SpritetoUse).Left = imgTile(Index).Left
'    imgSprite(SpritetoUse).Top = imgTile(Index).Top - Abs(25 - imgSprite(tiletouse).Height)
'End If
End Sub

Private Sub imgTile_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    imgTile(Index).Drag
End If
        
End Sub
Sub TileDraw(ByVal Index As Integer)
If curType = 0 Then
imgTile(Index).Picture = frmOptions.imgTile(CurTile).Picture
TileType(Index) = CurTile
    If CurTile = 31 Then
        imgTile(Index).BorderStyle = 1
    Else
        imgTile(Index).BorderStyle = 0
    End If
    
    If frmOptions.chkGoTo.Value = 1 Then
        TileGoto(Index) = True
        TileLink(Index) = frmOptions.txtLevel.Text
        imgTile(Index).BorderStyle = 1
    Else
        TileGoto(Index) = False
        If CurTile <> 31 Then
            imgTile(Index).BorderStyle = 0
        End If
    End If

End If

End Sub

Private Sub imgTile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCurTile.Caption = Index
End Sub

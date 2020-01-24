VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHelp 
      Height          =   3615
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.ListBox lstHelp 
      Height          =   3570
      ItemData        =   "frmHelp.frx":0000
      Left            =   120
      List            =   "frmHelp.frx":0019
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lstHelp_Click()
Dim strHelp As String
If lstHelp.ListIndex = 0 Then
    strHelp = "To open and save your quests in the editor simply choose that option from the menu.  Type the name of the map (do not type the path or the extension.)"
    txtHelp.Text = lstHelp.List(ListIndex) & vbNewLine & strHelp
End If
If lstHelp.ListIndex = 1 Then
    strHelp = "To place a tile, simply click the tile on the Options window.  Then, click the square on the editor window in which you want to paste that tile.  You can fill multiple tiles at once by Right Clicking the editor window and then dragging your mouse over those tiles."
    txtHelp.Text = lstHelp.List(ListIndex) & vbNewLine & strHelp
End If
If lstHelp.ListIndex = 2 Then
    strHelp = "To place a Log or a Boss simply click the Log of Boss in the Options window and then click the editor where you want to place the Log of Boss.  To remove a Log or Boss right click it."
    txtHelp.Text = lstHelp.List(ListIndex) & vbNewLine & strHelp
End If
If lstHelp.ListIndex = 3 Then
    strHelp = "To link a tile to another map, make sure that the Goto check box is checked.  Then, highlight the tile on the next map that you want the tile to link to.  Record the Tile Number and type it in to the Tile Link textbox.  Then, just normally click where you want the linking tile to be."
    txtHelp.Text = lstHelp.List(ListIndex) & vbNewLine & strHelp
End If
If lstHelp.ListIndex = 4 Then
    strHelp = "The Goto checkbox will link a tile to another map at the destination of the Goto Textbox.  The Random Battle textbox will enable random battles to occur on that tile.  The Boss Battle text box will automaticaly load a boss battle when a player goes to that tile."
    txtHelp.Text = lstHelp.List(ListIndex) & vbNewLine & strHelp
End If
If lstHelp.ListIndex = 5 Then
    strHelp = "The random battle editor is pretty self explanatory.  Fill in the name, picture, HP (Health), AP (Attack), Defense for the enemy.  Make sure that you type the AI in ALL CAPS.  As for the picture, any .gif file will work.  However, it is recomended that you make the image Transparent.  Also, don't make it much larger than what you've seen in the game."
    txtHelp.Text = lstHelp.List(ListIndex) & vbNewLine & strHelp
End If
If lstHelp.ListIndex = 6 Then
    strHelp = "To save an overall quest, load the Overall Quest editor in the Options window."
    txtHelp.Text = lstHelp.List(ListIndex) & vbNewLine & strHelp
End If

End Sub

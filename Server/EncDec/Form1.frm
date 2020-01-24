VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "EncDec"
   ClientHeight    =   1590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Mid$(strCode, 7, Len(strCode) - 12)"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Public Sub Encode(strValue As String, strINIValue As String, strINILength As String, nsave As String)
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
End Sub



Public Function Decode(sData As String) As String
On Error Resume Next
    Dim sTemp As String, sTemp1 As String


    For II% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, II%, 1)
        lT = Asc(sTemp$) \ 2
        sTemp1$ = sTemp1$ & Chr(lT)
    Next II%
    Decode$ = sTemp1$
End Function

Private Sub Check1_Click()
Text1_Change
End Sub

Private Sub Text1_Change()
'Text4.Text = Eyncrypt(Text1.Text)
If Check1.Value = 1 Then Text4.Text = Mid$(Eyncrypt(Text1.Text), 7, Len(Eyncrypt(Text1.Text)) - 12)
If Check1.Value = 0 Then Text4.Text = Eyncrypt(Text1.Text)
Text3.Text = Decode(Text1.Text)
End Sub

Private Sub Text3_DblClick()
Clipboard.SetText Text3.Text
End Sub

Private Sub Text4_DblClick()
Clipboard.SetText Text4.Text
End Sub

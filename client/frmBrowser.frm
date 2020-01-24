VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmBrowser 
   BackColor       =   &H00886000&
   Caption         =   "Golden Sun: The War of the Adepts - "
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8010
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   372
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   534
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00886000&
      Caption         =   "&Stop"
      Height          =   615
      Left            =   2640
      Picture         =   "frmBrowser.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00886000&
      Caption         =   "&Refresh"
      Height          =   615
      Left            =   1800
      Picture         =   "frmBrowser.frx":0C7B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdForward 
      BackColor       =   &H00886000&
      Caption         =   "&Forward"
      Height          =   615
      Left            =   960
      Picture         =   "frmBrowser.frx":1047
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00886000&
      Caption         =   "&Back"
      Height          =   615
      Left            =   120
      Picture         =   "frmBrowser.frx":1446
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7575
      ExtentX         =   13361
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Image imgDoc 
      Height          =   450
      Left            =   4800
      Picture         =   "frmBrowser.frx":1845
      Top             =   120
      Width           =   1200
   End
   Begin VB.Image imgGSA 
      Height          =   450
      Left            =   3480
      Picture         =   "frmBrowser.frx":1DF1
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label lblLoading 
      BackStyle       =   0  'Transparent
      Caption         =   "Done"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
On Error Resume Next
Web.GoBack
End Sub

Private Sub cmdForward_Click()
On Error Resume Next
Web.GoForward
End Sub

Private Sub cmdRefresh_Click()
On Error Resume Next
Web.Refresh
End Sub

Private Sub cmdStop_Click()
On Error Resume Next
Web.Stop
End Sub

Private Sub cmdStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    MsgBox "Someone stop this crazy thing!", vbInformation, "Easter Egg #9"
    Call Encode("9", "EGG9", "EGGL9", App.Path & "\settings.ini")
    
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
Web.Width = Me.ScaleWidth - 20
Web.Height = Me.ScaleHeight - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmIntro.Show
StopMidi
End Sub

Private Sub imgDoc_Click()
Web.Navigate "http://www.doc-ent.com"
End Sub

Private Sub imgGSA_Click()
Web.Navigate "http://www.doc-ent.com/gsa"
End Sub

Private Sub Web_DownloadBegin()
lblLoading.Caption = "Loading..."
End Sub

Private Sub Web_DownloadComplete()
lblLoading.Caption = "Done"
If Web.LocationURL = "http://www.freeopendiary.com/entrylist.asp?authorcode=B355234" Then
    MsgBox "Getting a little too personal, don't you think?", vbInformation, "Easter Egg #27"
    Call Encode("27", "EGG27", "EGGL27", App.Path & "\settings.ini")
End If
End Sub

Private Sub Web_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
lblLoading.Caption = "Done"
Me.Caption = "Golden Sun: The War of the Adepts - " & Web.LocationName
End Sub


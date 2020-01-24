Attribute VB_Name = "modMIDI"
'Tacvek Midi Handler Version 1.0
'See Mididoc.txt for instructions on use


#Const mciDebug = True

' WARNING: DO NOT ALTER ANYTHING BELOW THIS LINE
Public Type POINTAPI
        x As Long
        y As Long
End Type
Public Type MSG
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Const PM_NOREMOVE = &H0
Public Const PM_NOYIELD = &H2
Public Const PM_REMOVE = &H1
Public Const MCI_NOTIFY_ABORTED = &H4
Public Const MCI_NOTIFY_FAILURE = &H8
Public Const MCI_NOTIFY_SUCCESSFUL = &H1
Public Const MCI_NOTIFY_SUPERSEDED = &H2
Public Const MM_MCINOTIFY = &H3B9
Public curMessage As MSG
Public curMidi As String

Private mciReturn As Long
Private curRepeat As Boolean

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (curMSG As MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

#If mciDebug Then
    Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
        Private intMsgboxReturn As Integer
    Private strBuffer As String * 128
    Private mciErrorReturn As Boolean
#End If
'Option Explicit

Public Function fPlayMidi(ByVal strFilename As String, Optional ByVal blnRepeat As Boolean = True) As Boolean
Dim strMusic As String
strMusic = GetFromIni("GEN", "MUSIC", App.Path & "\settings.ini")
If strMusic = "ON" Then
    strFilename = strFilename & ".mid"
    Let curRepeat = blnRepeat
    Let mciReturn = mciSendString("close all", vbNull, 0, vbNull)
    If Not (mciReturn = 0) Then
        #If mciDebug Then
            Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
            If mciErrorReturn = True Then Let intMsgboxReturn = MsgBox(strBuffer, vbCritical + vbOKOnly, "MCI ERROR")
        #End If
        fPlayMidi = False
        Exit Function
    End If
    Let mciReturn = mciSendString("open """ & App.Path & "\" & strFilename & """ type sequencer alias MUSIC", vbNull, 0, vbNull) 'change to correct after testing
    If Not (mciReturn = 0) Then
        #If mciDebug Then
            Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
            If mciErrorReturn = True Then Let intMsgboxReturn = MsgBox(strBuffer, vbCritical + vbOKOnly, "MCI ERROR")
        #End If
        fPlayMidi = False
        Exit Function
    End If
    
    If curRepeat Then
        Let mciReturn = mciSendString("play MUSIC from 0 notify", vbNull, 0, frmIntro.hWnd)
    Else
        Let mciReturn = mciSendString("play MUSIC from 0", vbNull, 0, vbNull)
    End If
    
    If Not (mciReturn = 0) Then
        #If mciDebug Then
            Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
            If mciErrorReturn = True Then Let intMsgboxReturn = MsgBox(strBuffer, vbCritical + vbOKOnly, "MCI ERROR")
        #End If
        fPlayMidi = False
        Exit Function
    End If
    fPlayMidi = True
End If
End Function
Public Sub PlayMidi(ByVal strFilename As String, Optional ByVal blnRepeat As Boolean = True)
On Error Resume Next
'strFilename = strFilename & ".mid"
'Dim strMusic As String
'strMusic = GetFromIni("GEN", "MUSIC", App.Path & "\settings.ini")
'If strMusic = "ON" Then
'    frmIntro.Midi.Stop
'    strMidi = App.Path & "\" & strFilename
'    frmIntro.Midi.FileName = strMidi
'    frmIntro.Midi.Play
'End If

'Call StopMidi
'curMidi = strFilename
'Dim strMusic As String
'strMusic = GetFromIni("GEN", "MUSIC", App.Path & "\settings.ini")
'If strMusic = "ON" Then
'
'    Let curRepeat = blnRepeat
'    Let mciReturn = mciSendString("close all", vbNull, 0, vbNull)
'    If Not (mciReturn = 0) Then
'        #If mciDebug Then
'            Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
'            If mciErrorReturn = True Then Let intMsgboxReturn = MsgBox(strBuffer, vbCritical + vbOKOnly, "MCI ERROR")
'        #End If
'        Exit Sub
'    End If
'    strFilename = strFilename & ".mid"
'    Let mciReturn = mciSendString("open """ & App.Path & "\" & strFilename & """ type sequencer alias MUSIC", vbNull, 0, vbNull) 'change to correct after testing
'    If Not (mciReturn = 0) Then
'        #If mciDebug Then
'            Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
'            If mciErrorReturn = True Then
'                'Nothing
'            End If
'        #End If
'        Exit Sub
'    End If
'    If curRepeat Then
'        Let mciReturn = mciSendString("play MUSIC from 0 notify", vbNull, 0, frmIntro.hWnd)
'    Else
'        Let mciReturn = mciSendString("play MUSIC from 0", vbNull, 0, vbNull)
'    End If
'    If Not (mciReturn = 0) Then
'        #If mciDebug Then
'            Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
'            If mciErrorReturn = True Then Let intMsgboxReturn = MsgBox(strBuffer, vbCritical + vbOKOnly, "MCI ERROR")
'        #End If
'        Exit Sub
'    End If
'End If

End Sub
Sub RepeatMidi()
Let mciReturn = mciSendString("play MUSIC from 0 notify", vbNull, 0, frmIntro.hWnd)
If Not (mciReturn = 0) Then
    #If mciDebug Then
        Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
        If mciErrorReturn = True Then Let intMsgboxReturn = MsgBox(strBuffer, vbCritical + vbOKOnly, "MCI ERROR")
    #End If
    Exit Sub
End If
End Sub
Public Function fPauseMidi() As Boolean
Let mciReturn = mciSendString("stop MUSIC", vbNull, 0, vbNull)
If Not (mciReturn = 0) Then
    #If mciDebug Then
        Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
        If mciErrorReturn = True Then Let intMsgboxReturn = MsgBox(strBuffer, vbCritical + vbOKOnly, "MCI ERROR")
    #End If
    fPauseMidi = False
    Exit Function
End If
fPauseMidi = True
End Function
Public Sub PauseMidi()
Let mciReturn = mciSendString("stop MUSIC", vbNull, 0, vbNull)
If Not (mciReturn = 0) Then
    #If mciDebug Then
        Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
        If mciErrorReturn = True Then Let intMsgboxReturn = MsgBox(strBuffer, vbCritical + vbOKOnly, "MCI ERROR")
    #End If
    Exit Sub
End If
End Sub
Public Function fStopMidi() As Boolean
Let mciReturn = mciSendString("close all", vbNull, 0, vbNull)
If Not (mciReturn = 0) Then
    #If mciDebug Then
        Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
        If mciErrorReturn = True Then Let intMsgboxReturn = MsgBox(strBuffer, vbCritical + vbOKOnly, "MCI ERROR")
    #End If
    fStopMidi = False
    Exit Function
End If
fStopMidi = True
End Function

Public Sub StopMidi()
Let mciReturn = mciSendString("close all", vbNull, 0, vbNull)
If Not (mciReturn = 0) Then
    #If mciDebug Then
        Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
        If mciErrorReturn = True Then Let intMsgboxReturn = MsgBox(strBuffer, vbCritical + vbOKOnly, "MCI ERROR")
    #End If
    Exit Sub
End If
End Sub

Public Sub ResumeMidi()
If curRepeat Then
    Let mciReturn = mciSendString("play MUSIC from 0 notify", vbNull, 0, frmIntro.hWnd)
Else
    Let mciReturn = mciSendString("play MUSIC", vbNull, 0, vbNull)
End If
If Not (mciReturn = 0) Then
    #If mciDebug Then
        Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
        If mciErrorReturn = True Then Let intMsgboxReturn = MsgBox(strBuffer, vbCritical + vbOKOnly, "MCI ERROR")
    #End If
    Exit Sub
End If
End Sub
Public Function fResumeMidi() As Boolean
If curRepeat Then
    Let mciReturn = mciSendString("play MUSIC from 0 notify", vbNull, 0, frmIntro.hWnd)
Else
    Let mciReturn = mciSendString("play MUSIC", vbNull, 0, vbNull)
End If
If Not (mciReturn = 0) Then
    #If mciDebug Then
        Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
        If mciErrorReturn = True Then Let intMsgboxReturn = MsgBox(strBuffer, vbCritical + vbOKOnly, "MCI ERROR")
    #End If
    fResumeMidi = False
    Exit Function
End If
fResumeMidi = True
End Function

Public Sub AlwaysPlayMidi(ByVal strFilename As String, Optional ByVal blnRepeat As Boolean = True)
'Plays music even if music is off
Call StopMidi
curMidi = strFilename
    Let curRepeat = blnRepeat
    Let mciReturn = mciSendString("close all", vbNull, 0, vbNull)
    If Not (mciReturn = 0) Then
        #If mciDebug Then
            Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
            If mciErrorReturn = True Then Let intMsgboxReturn = MsgBox(strBuffer, vbCritical + vbOKOnly, "MCI ERROR")
        #End If
        Exit Sub
    End If
    strFilename = strFilename & ".mid"
    Let mciReturn = mciSendString("open """ & App.Path & "\" & strFilename & """ type sequencer alias MUSIC", vbNull, 0, vbNull) 'change to correct after testing
    If Not (mciReturn = 0) Then
        #If mciDebug Then
            Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
            If mciErrorReturn = True Then
                'Nothing
            End If
        #End If
        Exit Sub
    End If
    If curRepeat Then
        Let mciReturn = mciSendString("play MUSIC from 0 notify", vbNull, 0, frmIntro.hWnd)
    Else
        Let mciReturn = mciSendString("play MUSIC from 0", vbNull, 0, vbNull)
    End If
    If Not (mciReturn = 0) Then
        #If mciDebug Then
            Let mciErrorReturn = mciGetErrorString(mciReturn, strBuffer, 128)
            If mciErrorReturn = True Then Let intMsgboxReturn = MsgBox(strBuffer, vbCritical + vbOKOnly, "MCI ERROR")
        #End If
        Exit Sub
    End If

End Sub

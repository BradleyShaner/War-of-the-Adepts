Attribute VB_Name = "Module1"
Public Type Users
    Enabled As Boolean
    Away As String
    Number As Integer
    Name As String
    Rating As String
    Wins As Integer
    Losses As String
    Disconnects As String
    'Version As String
    Screen As String
    Pic As String
    Avatar As String
    Left As String
    Top As String
    IP As String
End Type

Public Type Game
    Enabled As Boolean
    Name As String
    IP As String
    Host As String
    
End Type

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Function GetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    GetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function
Function WriteIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    'WritePrivateProfileString
    WriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function
Function Eyncrypt(sData As String) As String
    Dim sTemp As String, sTemp1 As String
    Dim strBS As String
    Dim strBS2 As String
    Randomize
    strBS = strBS & Chr(Int((Rnd * 50 + 25)))
    strBS2 = strBS2 & Chr(Int((Rnd * 50 + 25)))
    
    sData = strBS & sData & strBS2

    For iI% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, iI%, 1)
        lT = Asc(sTemp$) * 2
        sTemp1$ = sTemp1$ & Chr(lT)
    Next iI%
    Eyncrypt$ = sTemp1$
End Function
Sub Encode(strValue As String, strINIValue As String, strINILength As String, nSave As String)
    Dim strLength As String
    
    strLength = CStr(Len(strValue))
    
    If Len(strValue) < 10 Then
        strLength = "0" & strLength
    End If
    
    strValue = Eyncrypt(strValue)


    
    Dim strLength2 As String
    
    strLength2 = strLength
    strLength = Eyncrypt(strLength2)
    
    Call WriteIni("GEN", strINIValue, strValue, nSave)
    Call WriteIni("GEN", strINILength, strLength, nSave)
    
End Sub

Public Function UltraDecode(strData As String, strLength As String, nSave As String) As String
If strData <> "" Then
Dim iLength As Integer
strData = GetFromIni("GEN", strData, nSave)
strLength = GetFromIni("GEN", strLength, nSave)


strLength = Decode(strLength)
strData = Decode(strData)

iLength = CInt(Mid$(strLength, 2, 2))

UltraDecode = Mid$(strData, 2, iLength)
End If
End Function
Public Function Decode(sData As String) As String
    Dim sTemp As String, sTemp1 As String


    For iI% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, iI%, 1)
        lT = Asc(sTemp$) \ 2
        sTemp1$ = sTemp1$ & Chr(lT)
    Next iI%
    Decode$ = sTemp1$
End Function


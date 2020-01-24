VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "base64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Base64 Encode/Decode
'
' This is an optimized version of the common Base 64 encode/decode.
' This version eliminates the repeditive calls to chr$() and asc(),
' as well as the linear searches I've seen in some routines.
'
' This method does use a bit more memory in permanent lookup tables
' than most do.  However, this eliminates the need for using vb's
' rather slow method of bit shifting (multiplication and division).
' This appears not to make much difference in the IDE, but make
' a huge difference in the exe.
'   Encodeing Index = 834 vs. 64 bytes standard
'   Decoding Index  = 1536 vs. 64 to 256 standard
'
' This routine also adds the CrLf on the fly rather than making
' a temporary copy of the encoded string then adding the crlf
'
' Encoding/Decoding data from and to a file should be changed to
' use a fixed buffer to reduce the memory requirements of EncodeFile, etc.
'
' All of this results in a speed increase:
'   Encode:
'         100 reps on a string of 28311 bytes
'                               IDE      EXE
'   Base64                      2824     300 (220 w/no overflow & array bound checks)
'   Base64a (unknown author)  375500* 185300*
'   Base64b (Wil Johnson)       2814     512 (410 w/no overflow & array bound checks)
'     *Extrapolated (based on 1 rep, I didn't have time to wait 30 minutes for 100)
'     *Unknown code is from ftp:altecdata.com/base64.cls
'
'  Decode
'         100 reps on a string of 28311 bytes
'                              IDE    EXE
'   Base64                    3384     351 (271 w/no overflow & array bound checks)
'   Base64a (unknown author)
'   Base64b (Wil Johnson)     5969    1191 (981 w/no overflow & array bound checks)
'   *Failed
'   *Unknown code is from ftp:altecdata.com/base64.cls
'
'
' Author: Tim Arheit - tarheit@wcoil.com
' Version: 1.0
'
' This code is provided as-is.  You are free to use and modify it
' as you wish.  Please report bugs, fixes and enhancements to the
' author.
'
' Sample Usage:
'    Dim b as Base64
'    b = New Base64
'    Debug.Print b.Encode("This is a test.") ' Prints "VGhpcyBpcyBhIHRlc3Qu"
'    Debug.Print b.Decode("VGhpcyBpcyBhIHRlc3Qu") ' Prints "This is a test."
'
' History:
'    11/13/00 - Code release. It appears to work.
'
'

Private Const MAX_LINELENGTH As Long = 76 ' Must be a multiple of 4
Private Const CHAR_EQUAL As Byte = 61
Private Const CHAR_CR As Byte = 13
Private Const CHAR_LF As Byte = 10


Private m_Index1(0 To 255) As Byte
Private m_Index2(0 To 255) As Byte
Private m_Index3(0 To 255) As Byte
Private m_Index4(0 To 63) As Byte
Private m_ReverseIndex1(0 To 255) As Byte
Private m_ReverseIndex2(0 To 255, 0 To 1) As Byte
Private m_ReverseIndex3(0 To 255, 0 To 1) As Byte
Private m_ReverseIndex4(0 To 255) As Byte

' Encode a string to a string.
Public Function Encode(sInput As String) As String
   Dim bTemp() As Byte
    
   'Convert to a byte array then convert.
   'This is faster the repetitive calls to asc() or chr$()
   bTemp = StrConv(sInput, vbFromUnicode)
   Encode = StrConv(EncodeArr(bTemp), vbUnicode)
End Function

'Decode a string to a string.
Public Function Decode(sInput As String) As String
   Dim bTemp() As Byte
    
   'Convert to a byte array then convert.
   'This is faster the repetitive calls to asc() or chr$()
   bTemp = StrConv(sInput, vbFromUnicode)
   Decode = StrConv(DecodeArr(bTemp), vbUnicode)
End Function

Public Sub DecodeToFile(sInput As String, sOutputFile As String)
   Dim bTemp() As Byte
   Dim fh As Long
    
   bTemp = StrConv(sInput, vbformunicode)
   bTemp = DecodeArr(bTemp)
    
   fh = FreeFile(0)
   Open sOutputFile For Binary Access Write As fh
   Put fh, , bTemp
   Close fh
End Sub

Public Sub DecodeFile(sInputFile As String, sOutputFile As String)
On Error Resume Next

   Dim bTemp() As Byte
   Dim fh As Long
    
   fh = FreeFile(0)
   Open sInputFile For Binary Access Read As fh
   ReDim bTemp(0 To LOF(fh) - 1)
   Get fh, , bTemp
   Close fh
   If sInputFile = sOutputFile Then FileSystem.Kill sInputFile
   bTemp = DecodeArr(bTemp)
   Open sOutputFile For Binary Access Write As fh
   Put fh, , bTemp
   Close fh
End Sub

Public Function EncodeFromFile(sFileName As String) As String
On Error Resume Next
   Dim bTemp() As Byte
   Dim fh As Long
    
   fh = FreeFile(0)
   Open sFileName For Binary Access Read As fh
   ReDim bTemp(0 To LOF(fh) - 1)
   Get fh, , bTemp
   Close fh
    
   EncodeFromFile = StrConv(EncodeArr(bTemp), vbUnicode)
End Function

Public Sub EncodeFile(sInputFile As String, sOutputFile As String)
   Dim bTemp() As Byte
   Dim fh As Long
    
   fh = FreeFile(0)
   Open sInputFile For Binary Access Read As fh
   ReDim bTemp(0 To LOF(fh) - 1)
   Get fh, , bTemp
   Close fh
   If sInputFile = sOutputFile Then FileSystem.Kill sInputFile
   bTemp = EncodeArr(bTemp)
   Open sOutputFile For Binary Access Write As fh
   Put fh, , bTemp
   Close fh
End Sub


Private Function EncodeArr(bInput() As Byte) As Variant
   Dim bOutput() As Byte
   Dim k As Long
   Dim l As Long
   Dim i As Long
   Dim evenBound As Long
   Dim CurrentOut As Long
   Dim b As Byte
   Dim c As Byte
   Dim d As Byte
   Dim linelength As Long
    
   k = LBound(bInput)
   l = UBound(bInput)
    
   'Calculate the input size
   i = l - k + 1
    
   'Calculate the output size
   Select Case i Mod 3
      Case 0:
         i = (i \ 3) * 4
         evenBound = l
      Case 1:
         i = ((i \ 3) * 4) + 4
         evenBound = l - 1
      Case 2:
         i = ((i \ 3) * 4) + 4
         evenBound = l - 2
      Case 3:
         i = ((i \ 3) * 4) + 4
         evenBound = l - 3
   End Select
    
   'Add in the line feeds.
  ' If i Mod MAX_LINELENGTH = 0 Then
  '    i = i + (i \ MAX_LINELENGTH) * 2 - 2
  ' Else
  '    i = i + (i \ MAX_LINELENGTH) * 2
  ' End If
    
   'Size the output array
   ReDim bOutput(0 To i - 1)
       
   CurrentOut = 0
   linelength = 0
    
   For i = k To evenBound Step 3
      b = bInput(i)
      c = bInput(i + 1)
      d = bInput(i + 2)
      bOutput(CurrentOut) = m_Index1(b And &HFC)
      bOutput(CurrentOut + 1) = m_Index2((b And &H3) Or (c And &HF0))
      bOutput(CurrentOut + 2) = m_Index3((c And &HF) Or (d And &HC0))
      bOutput(CurrentOut + 3) = m_Index4(d And &H3F)
      CurrentOut = CurrentOut + 4
      linelength = linelength + 4
       
 '     If linelength >= MAX_LINELENGTH Then
 '        If i <> l - 2 Then  ' If this is the last line, don't add crlf
 '           bOutput(CurrentOut) = CHAR_CR
 '           bOutput(CurrentOut + 1) = CHAR_LF
 '        End If
'         CurrentOut = CurrentOut + 2
'         linelength = 0
'      End If
   Next i
    
   Select Case l - i
      Case 1:
         b = bInput(i)
         c = bInput(i + 1)
         d = 0
         bOutput(CurrentOut) = m_Index1(b And &HFC)
         bOutput(CurrentOut + 1) = m_Index2((b And &H3) Or (c And &HF0))
         bOutput(CurrentOut + 2) = m_Index3((c And &HF) Or (d And &HC0))
         bOutput(CurrentOut + 3) = CHAR_EQUAL
         CurrentOut = CurrentOut + 4
         linelength = linelength + 4
      Case 0:
         b = bInput(i)
         c = 0
         bOutput(CurrentOut) = m_Index1(b And &HFC)
         bOutput(CurrentOut + 1) = m_Index2((b And &H3) Or (c And &HF0))
         bOutput(CurrentOut + 2) = CHAR_EQUAL
         bOutput(CurrentOut + 3) = CHAR_EQUAL
         CurrentOut = CurrentOut + 4
         linelength = linelength + 4
   End Select
    
   EncodeArr = bOutput
End Function


Private Function DecodeArr(bInput() As Byte) As Variant
   Dim bOutput() As Byte
   Dim OutLength As Long
   Dim CurrentOut As Long
    
   Dim k As Long
   Dim l As Long
   Dim i As Long
   Dim j As Long
    
   Dim b As Byte
   Dim c As Byte
   Dim d As Byte
   Dim e As Byte
    
   k = LBound(bInput)
   l = UBound(bInput)
    
   'Calculate the length of the input
   i = l - k + 1
    
   'Calculate the expected length of the output
   'It should be no more (but may possible be less)
'   j = i Mod (MAX_LINELENGTH + 2)
'   If j = 0 Then
'      OutLength = (i \ (MAX_LINELENGTH + 2)) * (MAX_LINELENGTH \ 4) * 3
'   Else
'      j = (j / 4) * 3
'      If bInput(l) = CHAR_EQUAL Then j = j - 1
'      If bInput(l - 1) = CHAR_EQUAL Then j = j - 1
'      OutLength = (i \ (MAX_LINELENGTH + 2)) * (MAX_LINELENGTH \ 4) * 3 + j
'   End If
OutLength = (i / 4) * 3
   'Allocate the output
   ReDim bOutput(0 To OutLength - 1)
    
   CurrentOut = 0
    
   For i = k To l
      Select Case bInput(i)
         Case CHAR_CR
            'Do nothing
         Case CHAR_LF
            'Do nothing
         Case Else
            If l - i >= 3 Then
               b = bInput(i)
               c = bInput(i + 1)
               d = bInput(i + 2)
               e = bInput(i + 3)
                
               If e <> CHAR_EQUAL Then
                  bOutput(CurrentOut) = m_ReverseIndex1(b) Or m_ReverseIndex2(c, 0)
                  bOutput(CurrentOut + 1) = m_ReverseIndex2(c, 1) Or m_ReverseIndex3(d, 0)
                  bOutput(CurrentOut + 2) = m_ReverseIndex3(d, 1) Or m_ReverseIndex4(e)
                  CurrentOut = CurrentOut + 3
                  i = i + 3
               ElseIf d <> CHAR_EQUAL Then
                  bOutput(CurrentOut) = m_ReverseIndex1(b) Or m_ReverseIndex2(c, 0)
                  bOutput(CurrentOut + 1) = m_ReverseIndex2(c, 1) Or m_ReverseIndex3(d, 0)
                  CurrentOut = CurrentOut + 2
                  i = i + 3
               Else
                  bOutput(CurrentOut) = m_ReverseIndex1(b) Or m_ReverseIndex2(c, 0)
                  CurrentOut = CurrentOut + 1
                  i = i + 3
               End If
                
            Else
               'Possible input code error, but may also be
               'an extra CrLf, so we will ignore it.
            End If
      End Select
   Next i
    
   'On properly formed input we should have to do this.
   If OutLength <> CurrentOut + 1 Then
      ReDim Preserve bOutput(0 To CurrentOut - 1)
   End If
    
   DecodeArr = bOutput
End Function


Private Sub Class_Initialize()
   Dim i As Long
    
   'Setup the encodeing and decoding lookup arrays.
   'Essentially we speed up the routine by pre-shifting
   'the data so it only needs combined with And and Or.
   m_Index4(0) = 65 'Asc("A")
   m_Index4(1) = 66 'Asc("B")
   m_Index4(2) = 67 'Asc("C")
   m_Index4(3) = 68 'Asc("D")
   m_Index4(4) = 69 'Asc("E")
   m_Index4(5) = 70 'Asc("F")
   m_Index4(6) = 71 'Asc("G")
   m_Index4(7) = 72 'Asc("H")
   m_Index4(8) = 73 'Asc("I")
   m_Index4(9) = 74 'Asc("J")
   m_Index4(10) = 75 'Asc("K")
   m_Index4(11) = 76 'Asc("L")
   m_Index4(12) = 77 'Asc("M")
   m_Index4(13) = 78 'Asc("N")
   m_Index4(14) = 79 'Asc("O")
   m_Index4(15) = 80 'Asc("P")
   m_Index4(16) = 81 'Asc("Q")
   m_Index4(17) = 82 'Asc("R")
   m_Index4(18) = 83 'Asc("S")
   m_Index4(19) = 84 'Asc("T")
   m_Index4(20) = 85 'Asc("U")
   m_Index4(21) = 86 'Asc("V")
   m_Index4(22) = 87 'Asc("W")
   m_Index4(23) = 88 'Asc("X")
   m_Index4(24) = 89 'Asc("Y")
   m_Index4(25) = 90 'Asc("Z")
   m_Index4(26) = 97 'Asc("a")
   m_Index4(27) = 98 'Asc("b")
   m_Index4(28) = 99 'Asc("c")
   m_Index4(29) = 100 'Asc("d")
   m_Index4(30) = 101 'Asc("e")
   m_Index4(31) = 102 'Asc("f")
   m_Index4(32) = 103 'Asc("g")
   m_Index4(33) = 104 'Asc("h")
   m_Index4(34) = 105 'Asc("i")
   m_Index4(35) = 106 'Asc("j")
   m_Index4(36) = 107 'Asc("k")
   m_Index4(37) = 108 'Asc("l")
   m_Index4(38) = 109 'Asc("m")
   m_Index4(39) = 110 'Asc("n")
   m_Index4(40) = 111 'Asc("o")
   m_Index4(41) = 112 'Asc("p")
   m_Index4(42) = 113 'Asc("q")
   m_Index4(43) = 114 'Asc("r")
   m_Index4(44) = 115 'Asc("s")
   m_Index4(45) = 116 'Asc("t")
   m_Index4(46) = 117 'Asc("u")
   m_Index4(47) = 118 'Asc("v")
   m_Index4(48) = 119 'Asc("w")
   m_Index4(49) = 120 'Asc("x")
   m_Index4(50) = 121 'Asc("y")
   m_Index4(51) = 122 'Asc("z")
   m_Index4(52) = 48 'Asc("0")
   m_Index4(53) = 49 'Asc("1")
   m_Index4(54) = 50 'Asc("2")
   m_Index4(55) = 51 'Asc("3")
   m_Index4(56) = 52 'Asc("4")
   m_Index4(57) = 53 'Asc("5")
   m_Index4(58) = 54 'Asc("6")
   m_Index4(59) = 55 'Asc("7")
   m_Index4(60) = 56 'Asc("8")
   m_Index4(61) = 57 'Asc("9")
   m_Index4(62) = 43 'Asc("+")
   m_Index4(63) = 47 'Asc("/")
    
   'Calculate the other Arrays
   For i = 0 To 63
      m_Index1((i * 4) And &HFC) = m_Index4(i)
      m_Index2(((i And &HF) * 16) Or ((i And &H30) \ 16)) = m_Index4(i)
      m_Index3((i \ 4 And &HF) Or ((i And &H3) * 64)) = m_Index4(i)
   Next i
    
    
   m_ReverseIndex4(65) = 0 'Asc("A")
   m_ReverseIndex4(66) = 1 'Asc("B")
   m_ReverseIndex4(67) = 2 'Asc("C")
   m_ReverseIndex4(68) = 3 'Asc("D")
   m_ReverseIndex4(69) = 4 'Asc("E")
   m_ReverseIndex4(70) = 5 'Asc("F")
   m_ReverseIndex4(71) = 6 'Asc("G")
   m_ReverseIndex4(72) = 7 'Asc("H")
   m_ReverseIndex4(73) = 8 'Asc("I")
   m_ReverseIndex4(74) = 9 'Asc("J")
   m_ReverseIndex4(75) = 10 'Asc("K")
   m_ReverseIndex4(76) = 11 'Asc("L")
   m_ReverseIndex4(77) = 12 'Asc("M")
   m_ReverseIndex4(78) = 13 'Asc("N")
   m_ReverseIndex4(79) = 14 'Asc("O")
   m_ReverseIndex4(80) = 15 'Asc("P")
   m_ReverseIndex4(81) = 16 'Asc("Q")
   m_ReverseIndex4(82) = 17 'Asc("R")
   m_ReverseIndex4(83) = 18 'Asc("S")
   m_ReverseIndex4(84) = 19 'Asc("T")
   m_ReverseIndex4(85) = 20 'Asc("U")
   m_ReverseIndex4(86) = 21 'Asc("V")
   m_ReverseIndex4(87) = 22 'Asc("W")
   m_ReverseIndex4(88) = 23 'Asc("X")
   m_ReverseIndex4(89) = 24 'Asc("Y")
   m_ReverseIndex4(90) = 25 'Asc("Z")
   m_ReverseIndex4(97) = 26 'Asc("a")
   m_ReverseIndex4(98) = 27 'Asc("b")
   m_ReverseIndex4(99) = 28 'Asc("c")
   m_ReverseIndex4(100) = 29 'Asc("d")
   m_ReverseIndex4(101) = 30 'Asc("e")
   m_ReverseIndex4(102) = 31 'Asc("f")
   m_ReverseIndex4(103) = 32 'Asc("g")
   m_ReverseIndex4(104) = 33 'Asc("h")
   m_ReverseIndex4(105) = 34 'Asc("i")
   m_ReverseIndex4(106) = 35 'Asc("j")
   m_ReverseIndex4(107) = 36 'Asc("k")
   m_ReverseIndex4(108) = 37 'Asc("l")
   m_ReverseIndex4(109) = 38 'Asc("m")
   m_ReverseIndex4(110) = 39 'Asc("n")
   m_ReverseIndex4(111) = 40 'Asc("o")
   m_ReverseIndex4(112) = 41 'Asc("p")
   m_ReverseIndex4(113) = 42 'Asc("q")
   m_ReverseIndex4(114) = 43 'Asc("r")
   m_ReverseIndex4(115) = 44 'Asc("s")
   m_ReverseIndex4(116) = 45 'Asc("t")
   m_ReverseIndex4(117) = 46 'Asc("u")
   m_ReverseIndex4(118) = 47 'Asc("v")
   m_ReverseIndex4(119) = 48 'Asc("w")
   m_ReverseIndex4(120) = 49 'Asc("x")
   m_ReverseIndex4(121) = 50 'Asc("y")
   m_ReverseIndex4(122) = 51 'Asc("z")
   m_ReverseIndex4(48) = 52 'Asc("0")
   m_ReverseIndex4(49) = 53 'Asc("1")
   m_ReverseIndex4(50) = 54 'Asc("2")
   m_ReverseIndex4(51) = 55 'Asc("3")
   m_ReverseIndex4(52) = 56 'Asc("4")
   m_ReverseIndex4(53) = 57 'Asc("5")
   m_ReverseIndex4(54) = 58 'Asc("6")
   m_ReverseIndex4(55) = 59 'Asc("7")
   m_ReverseIndex4(56) = 60 'Asc("8")
   m_ReverseIndex4(57) = 61 'Asc("9")
   m_ReverseIndex4(43) = 62 'Asc("+")
   m_ReverseIndex4(47) = 63 'Asc("/")
    
   'Calculate the other arrays.
   For i = 0 To 255
      If m_ReverseIndex4(i) <> 0 Then
         m_ReverseIndex1(i) = m_ReverseIndex4(i) * 4
          
         m_ReverseIndex2(i, 0) = m_ReverseIndex4(i) \ 16
         m_ReverseIndex2(i, 1) = (m_ReverseIndex4(i) And &HF) * 16
          
         m_ReverseIndex3(i, 0) = m_ReverseIndex4(i) \ 4
         m_ReverseIndex3(i, 1) = (m_ReverseIndex4(i) And &H3) * 64
      End If
   Next i
End Sub

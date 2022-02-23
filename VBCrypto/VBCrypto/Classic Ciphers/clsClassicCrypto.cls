VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClassicCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Classic Ciphers
' By: David Midkiff (mznull@earthlink.net)
'
' This contains a collection of classic ciphers and routines for use
' in basic cryptographic schemes. There is little to no security in
' any of these routines. They are somewhat fun as they bring back
' nostalgic memories of the past.

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private byteArray() As Byte
Private hiByte As Long
Private hiBound As Long
Dim intBin(7) As Integer

Private Sub Class_Initialize()
intBin(0) = 1
intBin(1) = 2
intBin(2) = 4
intBin(3) = 8
intBin(4) = 16
intBin(5) = 32
intBin(6) = 64
intBin(7) = 128
End Sub

Public Function Ascii2Binary(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim iAsc As Integer, hLoop As Integer, tInt As Integer, mLoop As Integer
    Reset
    For mLoop = 1 To Len(Text)
        iAsc = Asc(Mid$(Text, mLoop, 1))
        For hLoop = 7 To 0 Step -1
            If intBin(hLoop) <= iAsc - tInt Then
                Append "1"
                tInt = tInt + intBin(hLoop)
            Else
                Append "0"
            End If
        Next
        tInt = 0
    Next
    Ascii2Binary = Trim(GData)
    Reset
End Function
Private Sub Reset()
    hiByte = 0
    hiBound = 1024
    ReDim byteArray(hiBound)
End Sub
Private Sub Append(ByRef StringData As String, Optional Length As Long)
    Dim DataLength As Long
    If Length > 0 Then DataLength = Length Else DataLength = Len(StringData)
    If DataLength + hiByte > hiBound Then
        hiBound = hiBound + 1024
        ReDim Preserve byteArray(hiBound)
    End If
    CopyMem ByVal VarPtr(byteArray(hiByte)), ByVal StringData, DataLength
    hiByte = hiByte + DataLength
End Sub
Private Function HexDecode(Text As String) As String
    Dim iCount As Double
    Reset
    For iCount = 1 To Len(Text) Step 2
        Append Chr$(Val("&H" & Mid$(Text, iCount, 2)))
    Next
    HexDecode = GData
    Reset
End Function
Private Function HexEncode(Text As String) As String
    Dim iCount As Double, sTemp As String
    Reset
    For iCount = 1 To Len(Text)
        sTemp = Hex$(Asc(Mid$(Text, iCount, 1)))
        If Len(sTemp) < 2 Then sTemp = "0" & sTemp
        Append sTemp
    Next
    HexEncode = GData
    Reset
End Function
Private Property Get GData() As String
    Dim StringData As String
    StringData = Space(hiByte)
    CopyMem ByVal StringData, ByVal VarPtr(byteArray(0)), hiByte
    GData = StringData
End Property
Public Function Binary2Ascii(BinaryText As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Exit Function

begin:
    Dim iLoop As Integer, t As Integer, mLoop As Integer, s As String, n As Integer
    For mLoop = 1 To Len(BinaryText) / 8
        For iLoop = ((mLoop - 1) * 8) + 1 To ((mLoop - 1) * 8) + 8
            n = n + 1
            t = t + (Val(Mid(BinaryText, iLoop, 1)) * intBin(8 - n))
        Next
        s = s & Chr(t)
        t = 0: n = 0
    Next
    Binary2Ascii = s
End Function
Public Function ReverseString(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim iLoop As Integer
    Reset
    For iLoop = Len(Text) To 1 Step -1
        Append Mid$(Text, iLoop, 1)
    Next
    ReverseString = GData
    Reset
End Function
Public Function CaesarShift(Text As String, Difference As Integer) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    While Difference > 26
        Difference = Difference - 26
    Wend
    Reset
    Dim iLoop As Integer, tAsc As String
    For iLoop = 1 To Len(Text)
        tAsc = Mid$(Text, iLoop, 1)
        If tAsc = " " Then Append " "
        If tAsc = LCase$(tAsc) And tAsc <> " " Then
            If Asc(tAsc) + Difference > 122 Then Append Chr$(97 + ((Asc(tAsc) + Difference) - 122)) Else Append Chr$(Asc(tAsc) + Difference)
        ElseIf tAsc = UCase$(tAsc) And tAsc <> " " Then
            If Asc(tAsc) + Difference > 90 Then Append Chr$(65 + ((Asc(tAsc) + Difference) - 90)) Else Append Chr$(Asc(tAsc) + Difference)
        End If
    Next
    CaesarShift = GData
    Reset
End Function
Public Function Ascii2Transposition(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim iLoop As Integer, dInt As Integer
    Reset
    dInt = Format$(Len(Text) / 2, 0)
    For iLoop = 1 To Len(Text) / 2
        Append Mid$(Text, iLoop, 1)
        Append Mid$(Text, iLoop + dInt, 1)
    Next
    If Len(Text) / 2 < dInt Then Append Mid$(Text, dInt, 1)
    Ascii2Transposition = GData
    Reset
End Function
Public Function Transposition2Ascii(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim iLoop As Integer, dInt As Integer
    dInt = Len(Text) / 2
    Reset
    For iLoop = 1 To Len(Text) Step 2
        Append Mid$(Text, iLoop, 1)
    Next
    For iLoop = 2 To Len(Text) Step 2
        Append Mid$(Text, iLoop, 1)
    Next iLoop
    Transposition2Ascii = GData
    Reset
End Function
Public Function Ascii2PigLatin(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Exit Function

begin:
    Dim iInt As Integer, sString As String, n() As String, zInt As Integer
    sString = Text
    Do
        DoEvents
        iInt = InStr(sString, " ")
        If iInt = 0 Then Exit Do
        sString = Mid$(sString, iInt + 1)
        zInt = zInt + 1
    Loop
    n = Split(Text, " ")
    sString = ""
    For iInt = 0 To zInt
        sString = sString & Right$(n(iInt), Len(n(iInt)) - 1) & Left$(n(iInt), 1) & "ay "
    Next
    sString = Trim(sString)
    Ascii2PigLatin = sString
End Function

Public Function PigLatin2Ascii(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim iInt As Integer, sString As String, n() As String, zInt As Integer
    sString = Text
    Do
        DoEvents
        iInt = InStr(sString, " ")
        If iInt = 0 Then Exit Do
        sString = Mid$(sString, iInt + 1)
        zInt = zInt + 1
    Loop
    n = Split(Text, " ")
    sString = ""
    For iInt = 0 To zInt
        n(iInt) = Left$(n(iInt), Len(n(iInt)) - 2)
        sString = sString & Right$(n(iInt), 1) & Left$(n(iInt), Len(n(iInt)) - 1) & " "
    Next
    sString = Trim(sString)
    PigLatin2Ascii = sString
End Function

Function Ascii2Random(CustomValue As String, EncryptCode As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim a As String, i As Long, b As String, d As String, Q%, c As String
    Q = Val(Trigger(EncryptCode))
    a = CustomValue
    Reset
    For i = 1 To Len(a)
        DoEvents
        b = Mid$(a, i, 1)
        c = Asc(b) + Q
        If c < 0 Then c = c - c - c
        If Len(c) = 1 Then c = "00" & c
        If Len(c) = 2 Then c = "0" & c
        Append c
    Next i
    d = GData
    Reset
    Dim f As Integer, e As String, g As Integer
    For i = 1 To Len(d)
        DoEvents
        f = Mid$(d, i, 1)
        Randomize
        g = Rnd * 19
        Select Case f
            Case 0: Append Chr(g + 55)
            Case 1: Append Chr(g + 75)
            Case 2: Append Chr(g + 95)
            Case 3: Append Chr(g + 115)
            Case 4: Append Chr(g + 135)
            Case 5: Append Chr(g + 155)
            Case 6: Append Chr(g + 175)
            Case 7: Append Chr(g + 195)
            Case 8: Append Chr(g + 215)
            Case 9: Append Chr(g + 235)
        End Select
    Next i
    Ascii2Random = GData
    Reset
End Function

Function Random2Ascii(ByVal CustomValue As String, ByVal EncryptCode As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim X As String, Y As String, Z As String, Q As Integer, a As String, i As Long, b As String, c As String
    X = Left$(EncryptCode, 1)
    Y = Right$(EncryptCode, 1)
    Z = Mid$(EncryptCode, Format$(Len(EncryptCode) / 2, "#"), 1)
    Q = Format$((Asc(X) + Asc(Y) + Asc(Z)) / 6, "##")
    a = CustomValue
    For i = 1 To Len(a)
        DoEvents
        b = Val(Asc(Mid$(a, i, 1)))
        If b >= 55 And b < 75 Then c = c & 0
        If b >= 75 And b < 95 Then c = c & 1
        If b >= 95 And b < 115 Then c = c & 2
        If b >= 115 And b < 135 Then c = c & 3
        If b >= 135 And b < 155 Then c = c & 4
        If b >= 155 And b < 175 Then c = c & 5
        If b >= 175 And b < 195 Then c = c & 6
        If b >= 195 And b < 215 Then c = c & 7
        If b >= 215 And b < 235 Then c = c & 8
        If b >= 235 And b < 255 Then c = c & 9
        If b < 55 Or b >= 255 Then
            Random2Ascii = ""
            Exit Function
        End If
    Next i
    Reset
    Dim d As Integer, f As Integer
    For i = 1 To Len(c) Step 3
        DoEvents
        d = Mid$(c, i, 3)
        f = d - Q
        If f < 0 Then f = f - f - f
        If f > 255 Or f < 0 Then
            Random2Ascii = ""
            Exit Function
        End If
        Append Chr$(f)
    Next i
    Random2Ascii = GData
    Reset
End Function

Private Function Trigger(CodeWord As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Exit Function

begin:
    If CodeWord = "" Then Trigger = "": Exit Function
    Dim X$, Y$, Z$, Q%
    X = Left$(CodeWord, 1)
    Y = Right$(CodeWord, 1)
    Z = Mid$(CodeWord, Format$(Len(CodeWord) / 2, "#"), 1)
    Q = Format$((Asc(X) + Asc(Y) + Asc(Z)) / 6, "##")
    Trigger = Q
End Function

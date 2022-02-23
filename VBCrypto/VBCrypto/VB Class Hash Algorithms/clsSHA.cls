VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Visual Basic SHA Implementation
' Peter Girard and David Midkiff (mznull@earthlink.net)
'
' SHA-1 implementation optimised for the Visual Basic environment. Does
' not conform to speed specifications of the algorithm. Can be used in
' password protection related schemes.

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private byteArray() As Byte
Private hiByte As Long
Private hiBound As Long
Private Type Word
  B0 As Byte
  B1 As Byte
  B2 As Byte
  B3 As Byte
End Type
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
Private Property Get GData() As String
    Dim StringData As String
    StringData = Space(hiByte)
    CopyMem ByVal StringData, ByVal VarPtr(byteArray(0)), hiByte
    GData = StringData
End Property
Private Function AndW(w1 As Word, w2 As Word) As Word
  AndW.B0 = w1.B0 And w2.B0
  AndW.B1 = w1.B1 And w2.B1
  AndW.B2 = w1.B2 And w2.B2
  AndW.B3 = w1.B3 And w2.B3
End Function

Private Function OrW(w1 As Word, w2 As Word) As Word
  OrW.B0 = w1.B0 Or w2.B0
  OrW.B1 = w1.B1 Or w2.B1
  OrW.B2 = w1.B2 Or w2.B2
  OrW.B3 = w1.B3 Or w2.B3
End Function

Private Function XorW(w1 As Word, w2 As Word) As Word
  XorW.B0 = w1.B0 Xor w2.B0
  XorW.B1 = w1.B1 Xor w2.B1
  XorW.B2 = w1.B2 Xor w2.B2
  XorW.B3 = w1.B3 Xor w2.B3
End Function
Private Function NotW(w As Word) As Word
    NotW.B0 = Not w.B0
    NotW.B1 = Not w.B1
    NotW.B2 = Not w.B2
    NotW.B3 = Not w.B3
End Function

Private Function AddW(w1 As Word, w2 As Word) As Word
  Dim I As Integer, w As Word

  I = CInt(w1.B3) + w2.B3
  w.B3 = I Mod 256
  I = CInt(w1.B2) + w2.B2 + (I \ 256)
  w.B2 = I Mod 256
  I = CInt(w1.B1) + w2.B1 + (I \ 256)
  w.B1 = I Mod 256
  I = CInt(w1.B0) + w2.B0 + (I \ 256)
  w.B0 = I Mod 256
  
  AddW = w
End Function

Private Function CircShiftLeftW(w As Word, n As Integer) As Word
  Dim d1 As Double, d2 As Double
  
  d1 = WordToDouble(w)
  d2 = d1
  d1 = d1 * (2 ^ n)
  d2 = d2 / (2 ^ (32 - n))
  CircShiftLeftW = OrW(DoubleToWord(d1), DoubleToWord(d2))
End Function
Private Function WordToHex(w As Word) As String
    WordToHex = Right$("0" & Hex$(w.B0), 2) & Right$("0" & Hex$(w.B1), 2) & Right$("0" & Hex$(w.B2), 2) & Right$("0" & Hex$(w.B3), 2)
End Function

Private Function HexToWord(H As String) As Word
    HexToWord = DoubleToWord(Val("&H" & H & "#"))
End Function
Private Function DoubleToWord(n As Double) As Word
  DoubleToWord.B0 = Int(DMod(n, 2 ^ 32) / (2 ^ 24))
  DoubleToWord.B1 = Int(DMod(n, 2 ^ 24) / (2 ^ 16))
  DoubleToWord.B2 = Int(DMod(n, 2 ^ 16) / (2 ^ 8))
  DoubleToWord.B3 = Int(DMod(n, 2 ^ 8))
End Function
Private Function WordToDouble(w As Word) As Double
  WordToDouble = (w.B0 * (2 ^ 24)) + (w.B1 * (2 ^ 16)) + (w.B2 * (2 ^ 8)) + w.B3
End Function

Private Function DMod(value As Double, divisor As Double) As Double
  DMod = value - (Int(value / divisor) * divisor)
  If DMod < 0 Then DMod = DMod + divisor
End Function
Private Function F(t As Integer, B As Word, C As Word, D As Word) As Word
  Select Case t
    Case Is <= 19
      F = OrW(AndW(B, C), AndW(NotW(B), D))
    Case Is <= 39
      F = XorW(XorW(B, C), D)
    Case Is <= 59
      F = OrW(OrW(AndW(B, C), AndW(B, D)), AndW(C, D))
    Case Else
      F = XorW(XorW(B, C), D)
  End Select
End Function

Public Function SHA1(inMessage As String) As String
      Dim inLen As Long, inLenW As Word, padMessage As String, numBlocks As Long, w(0 To 79) As Word, blockText As String, wordText As String, I As Long, t As Integer, temp As Word, K(0 To 3) As Word, H0 As Word, H1 As Word, H2 As Word, H3 As Word, H4 As Word, A As Word, B As Word, C As Word, D As Word, E As Word
    
      inLen = Len(inMessage)
      inLenW = DoubleToWord(CDbl(inLen) * 8)
      
      Reset
        Append inMessage
        Append Chr$(128)
        Append String$((128 - (inLen Mod 64) - 9) Mod 64, Chr$(0))
        Append String$(4, Chr$(0))
        Append Chr$(inLenW.B0)
        Append Chr$(inLenW.B1)
        Append Chr$(inLenW.B2)
        Append Chr$(inLenW.B3)
        padMessage = GData
      Reset
    
      numBlocks = Len(padMessage) / 64
      K(0) = HexToWord("5A827999")
      K(1) = HexToWord("6ED9EBA1")
      K(2) = HexToWord("8F1BBCDC")
      K(3) = HexToWord("CA62C1D6")
      H0 = HexToWord("67452301")
      H1 = HexToWord("EFCDAB89")
      H2 = HexToWord("98BADCFE")
      H3 = HexToWord("10325476")
      H4 = HexToWord("C3D2E1F0")
    
      For I = 0 To numBlocks - 1
            blockText = Mid$(padMessage, (I * 64) + 1, 64)
            For t = 0 To 15
                wordText = Mid$(blockText, (t * 4) + 1, 4)
                w(t).B0 = Asc(Mid$(wordText, 1, 1))
                w(t).B1 = Asc(Mid$(wordText, 2, 1))
                w(t).B2 = Asc(Mid$(wordText, 3, 1))
                w(t).B3 = Asc(Mid$(wordText, 4, 1))
            Next t
            
            For t = 16 To 79
                w(t) = CircShiftLeftW(XorW(XorW(XorW(w(t - 3), w(t - 8)), w(t - 14)), w(t - 16)), 1)
            Next t
            
            A = H0: B = H1: C = H2: D = H3: E = H4
            
            For t = 0 To 79
                temp = AddW(AddW(AddW(AddW(CircShiftLeftW(A, 5), F(t, B, C, D)), E), w(t)), K(t \ 20))
                E = D: D = C
                C = CircShiftLeftW(B, 30)
                B = A: A = temp
            Next t
            
            H0 = AddW(H0, A): H1 = AddW(H1, B): H2 = AddW(H2, C)
            H3 = AddW(H3, D): H4 = AddW(H4, E)
      Next I

    SHA1 = WordToHex(H0) & WordToHex(H1) & WordToHex(H2) & WordToHex(H3) & WordToHex(H4)
    inMessage = ""
End Function


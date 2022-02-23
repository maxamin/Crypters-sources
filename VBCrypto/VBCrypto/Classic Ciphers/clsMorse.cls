VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMorse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Morse Code Encoder/Decoder
' By: David Midkiff (mznull@earthlink.net)

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private byteArray() As Byte
Private hiByte As Long
Private hiBound As Long

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
Private Sub Reset()
    hiByte = 0
    hiBound = 1024
    ReDim byteArray(hiBound)
End Sub
Public Function MorseEncode(Text As String) As String
    Dim xLoop As Long, r As String
    Reset
    For xLoop = 1 To Len(Text)
         r = Mid$(LCase$(Text), xLoop, 1)
        If r = "a" Then
            Append ".- "
        ElseIf r = "b" Then
            Append "-... "
        ElseIf r = "c" Then
            Append "-.-. "
        ElseIf r = "d" Then
            Append "-.. "
        ElseIf r = "e" Then
            Append ". "
        ElseIf r = "f" Then
            Append "..-. "
        ElseIf r = "g" Then
            Append "--. "
        ElseIf r = "h" Then
            Append ".... "
        ElseIf r = "i" Then
            Append ".. "
        ElseIf r = "j" Then
            Append ".--- "
        ElseIf r = "k" Then
            Append "-.- "
        ElseIf r = "l" Then
            Append ".-.. "
        ElseIf r = "m" Then
            Append "-- "
        ElseIf r = "n" Then
            Append "-. "
        ElseIf r = "o" Then
            Append "--- "
        ElseIf r = "p" Then
            Append ".--. "
        ElseIf r = "q" Then
            Append "--.- "
        ElseIf r = "r" Then
            Append ".-. "
        ElseIf r = "s" Then
            Append "... "
        ElseIf r = "t" Then
            Append "- "
        ElseIf r = "u" Then
            Append "..- "
        ElseIf r = "v" Then
            Append "...- "
        ElseIf r = "w" Then
            Append ".-- "
        ElseIf r = "x" Then
            Append "-..- "
        ElseIf r = "y" Then
            Append "-.-- "
        ElseIf r = "z" Then
            Append "--.. "
        ElseIf r = "1" Then
            Append ".---- "
        ElseIf r = "2" Then
            Append "..--- "
        ElseIf r = "3" Then
            Append "...-- "
        ElseIf r = "4" Then
            Append "....- "
        ElseIf r = "5" Then
            Append "..... "
        ElseIf r = "6" Then
            Append "-.... "
        ElseIf r = "7" Then
            Append "--... "
        ElseIf r = "8" Then
            Append "---.. "
        ElseIf r = "9" Then
            Append "----. "
        ElseIf r = "0" Then
            Append "----- "
        ElseIf r = " " Then
            Append "\ "
        ElseIf r = "!" Then
            Append "\ "
        ElseIf r = "@" Then
            Append "\ "
        ElseIf r = "#" Then
            Append "\ "
        ElseIf r = "$" Then
            Append "\ "
        ElseIf r = "%" Then
            Append "\ "
        ElseIf r = "^" Then
            Append "\ "
        ElseIf r = "&" Then
            Append "\ "
        ElseIf r = "*" Then
            Append "\ "
        ElseIf r = "<" Then
            Append "\ "
        ElseIf r = "=" Then
            Append "\ "
        ElseIf r = "+" Then
            Append "\ "
        ElseIf r = "]" Then
            Append "\ "
        ElseIf r = "[" Then
            Append "\ "
        ElseIf r = "{" Then
            Append "\ "
        ElseIf r = "}" Then
            Append "\ "
        ElseIf r = "|" Then
            Append "\ "
        ElseIf r = "." Then
            Append ".-.-.- "
        ElseIf r = "," Then
            Append "--..-- "
        ElseIf r = ":" Then
            Append "---... "
        ElseIf r = "?" Then
            Append "..--.. "
        ElseIf r = "'" Then
            Append ".----. "
        ElseIf r = "-" Then
            Append "-....- "
        ElseIf r = "/" Then
            Append "-..-. "
        ElseIf r = "(" Then
            Append "-.--. "
        ElseIf r = ")" Then
            Append "-.--.- "
        ElseIf r = ";" Then
            Append "-.-.-. "
        ElseIf r = Chr$(34) Then
            Append ".-..-. "
        ElseIf r = "_" Then
            Append "..--.- "
        ElseIf r = Chr$(13) Then
            Append "~ "
        End If
    Next
    MorseEncode = GData
    Reset
End Function
Private Property Get GData() As String
    Dim StringData As String
    StringData = Space(hiByte)
    CopyMem ByVal StringData, ByVal VarPtr(byteArray(0)), hiByte
    GData = StringData
End Property

Public Function MorseDecode(Text As String) As String
    Dim fLen As Double, x As Double, xStr As String, tStr As String
    If Text <> "" Then
        fLen = Len(Text)
        Do
            x = InStr(Text, " ")
            xStr = Right$(Text, Len(Text) - x)
            tStr = Left$(Text, x)
            If tStr = ".- " Then
                Append "a"
            ElseIf tStr = "-... " Then
                Append "b"
            ElseIf tStr = "-.-. " Then
                Append "c"
            ElseIf tStr = "-.. " Then
                Append "d"
            ElseIf tStr = ". " Then
                Append "e"
            ElseIf tStr = "..-. " Then
                Append "f"
            ElseIf tStr = "--. " Then
                Append "g"
            ElseIf tStr = ".... " Then
                Append "h"
            ElseIf tStr = ".. " Then
                Append "i"
            ElseIf tStr = ".--- " Then
                Append "j"
            ElseIf tStr = "-.- " Then
                Append "k"
            ElseIf tStr = ".-.. " Then
                Append "l"
            ElseIf tStr = "-- " Then
                Append "m"
            ElseIf tStr = "-. " Then
                Append "n"
            ElseIf tStr = "--- " Then
                Append "o"
            ElseIf tStr = ".--. " Then
                Append "p"
            ElseIf tStr = "--.- " Then
                Append "q"
            ElseIf tStr = ".-. " Then
                Append "r"
            ElseIf tStr = "... " Then
                Append "s"
            ElseIf tStr = "- " Then
                Append "t"
            ElseIf tStr = "..- " Then
                Append "u"
            ElseIf tStr = "...- " Then
                Append "v"
            ElseIf tStr = ".-- " Then
                Append "w"
            ElseIf tStr = "-..- " Then
                Append "x"
            ElseIf tStr = "-.-- " Then
                Append "y"
            ElseIf tStr = "--.. " Then
                Append "z"
            ElseIf tStr = ".---- " Then
                Append "1"
            ElseIf tStr = "..--- " Then
                Append "2"
            ElseIf tStr = "...-- " Then
                Append "3"
            ElseIf tStr = "....- " Then
                Append "4"
            ElseIf tStr = "..... " Then
                Append "5"
            ElseIf tStr = "-.... " Then
                Append "6"
            ElseIf tStr = "--... " Then
                Append "7"
            ElseIf tStr = "---.. " Then
                Append "8"
            ElseIf tStr = "----. " Then
                Append "9"
            ElseIf tStr = "----- " Then
                Append "0"
            ElseIf tStr = ".-.-.- " Then
                Append "."
            ElseIf tStr = "--..-- " Then
                Append ","
            ElseIf tStr = "---... " Then
                Append ":"
            ElseIf tStr = "..--.. " Then
                Append "?"
            ElseIf tStr = ".----. " Then
                Append "'"
            ElseIf tStr = "-....- " Then
                Append "-"
            ElseIf tStr = "-..-. " Then
                Append "/"
            ElseIf tStr = "-.-.-. " Then
                Append ";"
            ElseIf tStr = "-.--. " Then
                Append "("
            ElseIf tStr = "-.--.- " Then
                Append ")"
            ElseIf tStr = ".-..-. " Then
                Append Chr$(34)
            ElseIf tStr = "..--.- " Then
                Append "_"
            ElseIf tStr = "\ " Then
                Append " "
            ElseIf tStr = "~ " Then
                Append vbCrLf
            Else
                Append " "
                GoTo bypass
            End If
            Text = xStr
        Loop Until Text = ""
bypass:
    Else
    End If
    MorseDecode = GData
    Reset
End Function

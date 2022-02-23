VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cryptosy"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAction 
      Caption         =   "Build"
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetFileNameFromBrowseW Lib "shell32" Alias "#63" (ByVal hwndOwner As Long, ByVal lpstrFile As Long, ByVal nMaxFile As Long, ByVal lpstrInitialDir As Long, ByVal lpstrDefExt As Long, ByVal lpstrFilter As Long, ByVal lpstrTitle As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Private Function LoadFile(sPath As String) As String
    Dim lFileSize As Long
    Dim sData As String
    
    On Error Resume Next
    
    Open sPath For Binary Access Read As #1
    lFileSize = LOF(1)
    sData = Input$(lFileSize, 1)
    Close #1
    LoadFile = sData
End Function

Private Sub cmdAction_Click(Index As Integer)
    Select Case Index
    
        Case 0
            Dim sSave As String
            sSave = Space(255)
            GetFileNameFromBrowseW Me.hWnd, StrPtr(sSave), 255, StrPtr("c:\"), StrPtr("txt"), StrPtr("Apps (*.EXE)" + Chr$(0) + "*.EXE" + Chr$(0) + "All files (*.*)" + Chr$(0) + "*.*" + Chr$(0)), StrPtr("Select File")
            Text1 = Left$(sSave, lstrlen(sSave))
            
        Case 1
            Dim sBuff As String
            Dim c As New clsCryptAPI
            Dim sSize As String * 8
            
            If Not Text1 = vbNullString Then
            
                If PathFileExists(App.Path & "\Test.exe") Then
                    Kill App.Path & "\Test.exe"
                End If
                
                Open App.Path & "\Test.exe" For Binary Access Write As #2
                sBuff = LoadFile(App.Path & "\stub.exe")
                Put #2, , sBuff
                sBuff = LoadFile(Text1)
                sBuff = c.EncryptString(sBuff)
                Put #2, , sBuff
                sSize = Len(sBuff)
                Put #2, , sSize
                Put #2, , 27
                Close #2
                MsgBox "Done"
            End If
    End Select
End Sub

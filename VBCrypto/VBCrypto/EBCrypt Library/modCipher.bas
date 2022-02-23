Attribute VB_Name = "modCipher"
Option Explicit
' Utilize ebCrypt.dll
' By: David Midkiff (mznull@earthlink.net)
'
' This provides easier interface with ebCrypt.dll (v2.0) to
' access several encryption and hash algorithms. The DLL is
' Copyright (c) 2000-2001, EB Design Pty Ltd.

Public Enum Algorithms
    BLOWFISH
    IDEA
    TRIPLEDES
    DES
    DESE
    CAST5
    SERPENT128
    SERPENT192
    SERPENT256
    RIJNDAEL128
    RIJNDAEL192
    RIJNDAEL256
    RC4
    TWOFISH
End Enum
Public Enum HashAlgorithms
    MD2
    MD5
    RipeMD160
    SHA1
End Enum

Private m_bytIndex(0 To 63) As Byte
Private m_bytReverseIndex(0 To 255) As Byte
Private Const k_bytEqualSign As Byte = 61
Private Const k_bytMask1 As Byte = 3
Private Const k_bytMask2 As Byte = 15
Private Const k_bytMask3 As Byte = 63
Private Const k_bytMask4 As Byte = 192
Private Const k_bytMask5 As Byte = 240
Private Const k_bytMask6 As Byte = 252
Private Const k_bytShift2 As Byte = 4
Private Const k_bytShift4 As Byte = 16
Private Const k_bytShift6 As Byte = 64
Private Const k_lMaxBytesPerLine As Long = 152
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Initialized As Boolean
Public Function Decode64(sInput As String) As String
    If sInput = "" Then Exit Function
    Decode64 = StrConv(DecodeArray64(sInput), vbUnicode)
End Function

Public Function DecodeArray64(sInput As String) As Byte()
    If m_bytReverseIndex(47) <> 63 Then Initialize
    Dim bytInput() As Byte
    Dim bytWorkspace() As Byte
    Dim bytResult() As Byte
    Dim lInputCounter As Long
    Dim lWorkspaceCounter As Long
    
    bytInput = Replace(Replace(sInput, vbCrLf, ""), "=", "")
    ReDim bytWorkspace(LBound(bytInput) To (UBound(bytInput) * 2)) As Byte
    lWorkspaceCounter = LBound(bytWorkspace)
    For lInputCounter = LBound(bytInput) To UBound(bytInput)
        bytInput(lInputCounter) = m_bytReverseIndex(bytInput(lInputCounter))
    Next lInputCounter
    
    For lInputCounter = LBound(bytInput) To (UBound(bytInput) - ((UBound(bytInput) Mod 8) + 8)) Step 8
        bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * k_bytShift2) + (bytInput(lInputCounter + 2) \ k_bytShift4)
        bytWorkspace(lWorkspaceCounter + 1) = ((bytInput(lInputCounter + 2) And k_bytMask2) * k_bytShift4) + (bytInput(lInputCounter + 4) \ k_bytShift2)
        bytWorkspace(lWorkspaceCounter + 2) = ((bytInput(lInputCounter + 4) And k_bytMask1) * k_bytShift6) + bytInput(lInputCounter + 6)
        lWorkspaceCounter = lWorkspaceCounter + 3
    Next lInputCounter
    
    Select Case (UBound(bytInput) Mod 8):
        Case 3:
            bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * k_bytShift2) + (bytInput(lInputCounter + 2) \ k_bytShift4)
        Case 5:
            bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * k_bytShift2) + (bytInput(lInputCounter + 2) \ k_bytShift4)
            bytWorkspace(lWorkspaceCounter + 1) = ((bytInput(lInputCounter + 2) And k_bytMask2) * k_bytShift4) + (bytInput(lInputCounter + 4) \ k_bytShift2)
            lWorkspaceCounter = lWorkspaceCounter + 1
        Case 7:
            bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * k_bytShift2) + (bytInput(lInputCounter + 2) \ k_bytShift4)
            bytWorkspace(lWorkspaceCounter + 1) = ((bytInput(lInputCounter + 2) And k_bytMask2) * k_bytShift4) + (bytInput(lInputCounter + 4) \ k_bytShift2)
            bytWorkspace(lWorkspaceCounter + 2) = ((bytInput(lInputCounter + 4) And k_bytMask1) * k_bytShift6) + bytInput(lInputCounter + 6)
            lWorkspaceCounter = lWorkspaceCounter + 2
    End Select
    
    ReDim bytResult(LBound(bytWorkspace) To lWorkspaceCounter) As Byte
    If LBound(bytWorkspace) = 0 Then lWorkspaceCounter = lWorkspaceCounter + 1
    CopyMemory VarPtr(bytResult(LBound(bytResult))), VarPtr(bytWorkspace(LBound(bytWorkspace))), lWorkspaceCounter
    DecodeArray64 = bytResult
End Function

Public Function Encode64(ByRef sInput As String) As String
    If sInput = "" Then Exit Function
    Dim bytTemp() As Byte
    bytTemp = StrConv(sInput, vbFromUnicode)
    Encode64 = EncodeArray64(bytTemp)
End Function

Public Function EncodeArray64(ByRef bytInput() As Byte) As String
    On Error GoTo ErrorHandler
    If m_bytReverseIndex(47) <> 63 Then Initialize
    
    Dim bytWorkspace() As Byte, bytResult() As Byte
    Dim bytCrLf(0 To 3) As Byte, lCounter As Long
    Dim lWorkspaceCounter As Long, lLineCounter As Long
    Dim lCompleteLines As Long, lBytesRemaining As Long
    Dim lpWorkSpace As Long, lpResult As Long
    Dim lpCrLf As Long

    If UBound(bytInput) < 1024 Then
        ReDim bytWorkspace(LBound(bytInput) To (LBound(bytInput) + 4096)) As Byte
    Else
        ReDim bytWorkspace(LBound(bytInput) To (UBound(bytInput) * 4)) As Byte
    End If

    lWorkspaceCounter = LBound(bytWorkspace)

    For lCounter = LBound(bytInput) To (UBound(bytInput) - ((UBound(bytInput) Mod 3) + 3)) Step 3
        bytWorkspace(lWorkspaceCounter) = m_bytIndex((bytInput(lCounter) \ k_bytShift2))
        bytWorkspace(lWorkspaceCounter + 2) = m_bytIndex(((bytInput(lCounter) And k_bytMask1) * k_bytShift4) + ((bytInput(lCounter + 1)) \ k_bytShift4))
        bytWorkspace(lWorkspaceCounter + 4) = m_bytIndex(((bytInput(lCounter + 1) And k_bytMask2) * k_bytShift2) + (bytInput(lCounter + 2) \ k_bytShift6))
        bytWorkspace(lWorkspaceCounter + 6) = m_bytIndex(bytInput(lCounter + 2) And k_bytMask3)
        lWorkspaceCounter = lWorkspaceCounter + 8
    Next lCounter

    Select Case (UBound(bytInput) Mod 3):
        Case 0:
            bytWorkspace(lWorkspaceCounter) = m_bytIndex((bytInput(lCounter) \ k_bytShift2))
            bytWorkspace(lWorkspaceCounter + 2) = m_bytIndex((bytInput(lCounter) And k_bytMask1) * k_bytShift4)
            bytWorkspace(lWorkspaceCounter + 4) = k_bytEqualSign
            bytWorkspace(lWorkspaceCounter + 6) = k_bytEqualSign
        Case 1:
            bytWorkspace(lWorkspaceCounter) = m_bytIndex((bytInput(lCounter) \ k_bytShift2))
            bytWorkspace(lWorkspaceCounter + 2) = m_bytIndex(((bytInput(lCounter) And k_bytMask1) * k_bytShift4) + ((bytInput(lCounter + 1)) \ k_bytShift4))
            bytWorkspace(lWorkspaceCounter + 4) = m_bytIndex((bytInput(lCounter + 1) And k_bytMask2) * k_bytShift2)
            bytWorkspace(lWorkspaceCounter + 6) = k_bytEqualSign
        Case 2:
            bytWorkspace(lWorkspaceCounter) = m_bytIndex((bytInput(lCounter) \ k_bytShift2))
            bytWorkspace(lWorkspaceCounter + 2) = m_bytIndex(((bytInput(lCounter) And k_bytMask1) * k_bytShift4) + ((bytInput(lCounter + 1)) \ k_bytShift4))
            bytWorkspace(lWorkspaceCounter + 4) = m_bytIndex(((bytInput(lCounter + 1) And k_bytMask2) * k_bytShift2) + ((bytInput(lCounter + 2)) \ k_bytShift6))
            bytWorkspace(lWorkspaceCounter + 6) = m_bytIndex(bytInput(lCounter + 2) And k_bytMask3)
    End Select

    lWorkspaceCounter = lWorkspaceCounter + 8

    If lWorkspaceCounter <= k_lMaxBytesPerLine Then
        EncodeArray64 = Left$(bytWorkspace, InStr(1, bytWorkspace, Chr$(0)) - 1)
    Else
        bytCrLf(0) = 13
        bytCrLf(1) = 0
        bytCrLf(2) = 10
        bytCrLf(3) = 0
        ReDim bytResult(LBound(bytWorkspace) To UBound(bytWorkspace))
        lpWorkSpace = VarPtr(bytWorkspace(LBound(bytWorkspace)))
        lpResult = VarPtr(bytResult(LBound(bytResult)))
        lpCrLf = VarPtr(bytCrLf(LBound(bytCrLf)))
        lCompleteLines = Fix(lWorkspaceCounter / k_lMaxBytesPerLine)
        
        For lLineCounter = 0 To lCompleteLines
            CopyMemory lpResult, lpWorkSpace, k_lMaxBytesPerLine
            lpWorkSpace = lpWorkSpace + k_lMaxBytesPerLine
            lpResult = lpResult + k_lMaxBytesPerLine
            CopyMemory lpResult, lpCrLf, 4&
            lpResult = lpResult + 4&
        Next lLineCounter
        
        lBytesRemaining = lWorkspaceCounter - (lCompleteLines * k_lMaxBytesPerLine)
        If lBytesRemaining > 0 Then CopyMemory lpResult, lpWorkSpace, lBytesRemaining
        EncodeArray64 = Left$(bytResult, InStr(1, bytResult, Chr$(0)) - 1)
    End If
    Exit Function

ErrorHandler:
    Erase bytResult
    EncodeArray64 = bytResult
End Function


Private Function FileExist(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    Call FileLen(FileName)
    FileExist = True
    Exit Function
    
ErrorHandler:
    FileExist = False
End Function
Public Function EncryptFile(Which As Algorithms, InFile As String, OutFile As String, Overwrite As Boolean, Optional OutputIn64 As Boolean, Optional Key As String, Optional Salt As String) As Boolean
    On Error GoTo ErrorHandler
    If FileExist(InFile) = False Then
        EncryptFile = False
        Exit Function
    End If
    If FileExist(OutFile) = True And Overwrite = False Then
        EncryptFile = False
        Exit Function
    End If
    Dim FileO As Integer, Buffer() As Byte
    FileO = FreeFile
    Open InFile For Binary As #FileO
        ReDim Buffer(0 To LOF(FileO) - 1)
        Get #FileO, , Buffer()
    Close #FileO
    If FileExist(OutFile) = True Then Kill OutFile
    FileO = FreeFile
    Buffer() = EncryptArray(Which, Buffer(), Key, Salt)

    Open OutFile For Binary As #FileO
        If OutputIn64 = True Then
            Put #FileO, , EncodeArray64(Buffer())
        Else
            Put #FileO, , Buffer()
        End If
    Close #FileO
    EncryptFile = True
    Erase Buffer(): InFile = "": OutFile = "": Key = "": Salt = ""
    Exit Function

ErrorHandler:
    EncryptFile = False
    Erase Buffer(): InFile = "": OutFile = "": Key = "": Salt = ""
End Function
Public Function DecryptFile(Which As Algorithms, InFile As String, OutFile As String, Overwrite As Boolean, Optional IsFileIn64 As Boolean, Optional Key As String, Optional Salt As String) As Boolean
    On Error GoTo ErrorHandler
    If FileExist(InFile) = False Then
        DecryptFile = False
        Exit Function
    End If
    If FileExist(OutFile) = True Then
        DecryptFile = False
        Exit Function
    End If
    Dim FileO As Integer, Buffer() As Byte
    FileO = FreeFile
    Open InFile For Binary As #FileO
        ReDim Buffer(0 To LOF(FileO) - 1)
        Get #FileO, , Buffer()
    Close #FileO
    If FileExist(OutFile) = True Then Kill OutFile
    FileO = FreeFile
    If IsFileIn64 = True Then Buffer() = DecodeArray64(StrConv(Buffer(), vbUnicode))
    Buffer() = DecryptArray(Which, Buffer(), Key, Salt)
    Open OutFile For Binary As #FileO
        Put #FileO, , Buffer()
    Close #FileO
    DecryptFile = True
    Erase Buffer(): InFile = "": OutFile = "": Key = "": Salt = ""
    Exit Function

ErrorHandler:
    DecryptFile = False
    Erase Buffer(): InFile = "": OutFile = "": Key = "": Salt = ""
End Function
Public Function Hash(Which As HashAlgorithms, Message As String) As String
    If Message = "" Then Exit Function
    Dim hsh As ebcryptlib.eb_c_Hash
    Set hsh = CreateObject("EbCrypt.eb_c_Hash")
    If Which = MD2 Then Hash = hsh.HashString(EB_CRYPT_HASH_ALGORITHM_MD2, Message)
    If Which = MD5 Then Hash = hsh.HashString(EB_CRYPT_HASH_ALGORITHM_MD5, Message)
    If Which = SHA1 Then Hash = hsh.HashString(EB_CRYPT_HASH_ALGORITHM_SHA1, Message)
    If Which = RipeMD160 Then Hash = hsh.HashString(EB_CRYPT_HASH_ALGORITHM_RIPEMD160, Message)
End Function
Public Function EncryptString(Which As Algorithms, Text As String, Optional OutputIn64 As Boolean, Optional Key As String, Optional Salt As String) As String
    On Error GoTo ErrorHandler
    Dim cipher As ebcryptlib.eb_c_Cipher
    Set cipher = New ebcryptlib.eb_c_Cipher
    If Len(Salt) < 8 Then Salt = Salt & Space$(8 - Len(Salt))
    If Which = BLOWFISH Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_BLOWFISH_OFB, Key, Salt, Text), vbUnicode)
    If Which = CAST5 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_CAST5_OFB, Key, Salt, Text), vbUnicode)
    If Which = DES Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_DES_OFB, Key, Salt, Text), vbUnicode)
    If Which = DESE Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_DESE_OFB, Key, Salt, Text), vbUnicode)
    If Which = TRIPLEDES Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_DES3_OFB, Key, Salt, Text), vbUnicode)
    If Which = IDEA Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_IDEA_OFB, Key, Salt, Text), vbUnicode)
    If Which = RC4 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_RC4, Key, Salt, Text), vbUnicode)
    If Which = RIJNDAEL128 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_128, Key, Salt, Text), vbUnicode)
    If Which = RIJNDAEL192 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_192, Key, Salt, Text), vbUnicode)
    If Which = RIJNDAEL256 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, Key, Salt, Text), vbUnicode)
    If Which = SERPENT128 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_128, Key, Salt, Text), vbUnicode)
    If Which = SERPENT192 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_192, Key, Salt, Text), vbUnicode)
    If Which = SERPENT256 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_256, Key, Salt, Text), vbUnicode)
    If Which = TWOFISH Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_TWOFISH_CBC, Key, Salt, Text), vbUnicode)
    If OutputIn64 = True Then EncryptString = Encode64(EncryptString)
    Key = "": Salt = "": Text = ""
    Exit Function

ErrorHandler:
MsgBox Err.Description
Key = "": Salt = "": Text = ""
End Function
Public Function EncryptArray(Which As Algorithms, InputArray() As Byte, Optional Key As String, Optional Salt As String) As Variant
    On Error GoTo ErrorHandler
    Dim cipher As ebcryptlib.eb_c_Cipher
    Set cipher = New ebcryptlib.eb_c_Cipher
    If Len(Salt) < 8 Then Salt = Salt & Space$(8 - Len(Salt))
    
    If Which = BLOWFISH Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_BLOWFISH_OFB, Key, Salt, InputArray())
    If Which = CAST5 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_CAST5_OFB, Key, Salt, InputArray())
    If Which = DES Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_DES_OFB, Key, Salt, InputArray())
    If Which = DESE Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_DESE_OFB, Key, Salt, InputArray())
    If Which = TRIPLEDES Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_DES3_OFB, Key, Salt, InputArray())
    If Which = IDEA Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_IDEA_OFB, Key, Salt, InputArray())
    If Which = RC4 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RC4, Key, Salt, InputArray())
    If Which = RIJNDAEL128 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_128, Key, Salt, InputArray())
    If Which = RIJNDAEL192 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_192, Key, Salt, InputArray())
    If Which = RIJNDAEL256 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, Key, Salt, InputArray())
    If Which = SERPENT128 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_128, Key, Salt, InputArray())
    If Which = SERPENT192 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_192, Key, Salt, InputArray())
    If Which = SERPENT256 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_256, Key, Salt, InputArray())
    If Which = TWOFISH Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_TWOFISH_CBC, Key, Salt, InputArray())
    Erase InputArray(): Key = "": Salt = ""
    Exit Function

ErrorHandler:
MsgBox Err.Description
Erase InputArray(): Key = "": Salt = ""
End Function
Public Function DecryptArray(Which As Algorithms, InputArray() As Byte, Optional Key As String, Optional Salt As String) As Variant
    On Error GoTo ErrorHandler
    Dim cipher As ebcryptlib.eb_c_Cipher
    Set cipher = New ebcryptlib.eb_c_Cipher
    If Len(Salt) < 8 Then Salt = Salt & Space$(8 - Len(Salt))
    
    If Which = BLOWFISH Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_BLOWFISH_OFB, Key, Salt, InputArray())
    If Which = CAST5 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_CAST5_OFB, Key, Salt, InputArray())
    If Which = DES Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_DES_OFB, Key, Salt, InputArray())
    If Which = DESE Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_DESE_OFB, Key, Salt, InputArray())
    If Which = TRIPLEDES Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_DES3_OFB, Key, Salt, InputArray())
    If Which = IDEA Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_IDEA_OFB, Key, Salt, InputArray())
    If Which = RC4 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RC4, Key, Salt, InputArray())
    If Which = RIJNDAEL128 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_128, Key, Salt, InputArray())
    If Which = RIJNDAEL192 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_192, Key, Salt, InputArray())
    If Which = RIJNDAEL256 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, Key, Salt, InputArray())
    If Which = SERPENT128 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_128, Key, Salt, InputArray())
    If Which = SERPENT192 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_192, Key, Salt, InputArray())
    If Which = SERPENT256 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_256, Key, Salt, InputArray())
    If Which = TWOFISH Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_TWOFISH_CBC, Key, Salt, InputArray())
    Erase InputArray(): Key = "": Salt = ""
    Exit Function

ErrorHandler:
MsgBox Err.Description
Erase InputArray(): Key = "": Salt = ""
End Function



Public Function DecryptString(Which As Algorithms, CipherText As String, Optional IsTextIn64 As Boolean, Optional Key As String, Optional Salt As String) As String
    On Error GoTo ErrorHandler
    Dim cipher As ebcryptlib.eb_c_Cipher, BArray() As Byte
    Set cipher = New ebcryptlib.eb_c_Cipher
    If IsTextIn64 = True Then CipherText = Decode64(CipherText)
    If Len(Salt) < 8 Then Salt = Salt & Space$(8 - Len(Salt))
    BArray() = StrConv(CipherText, vbFromUnicode)
    If Which = BLOWFISH Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_BLOWFISH_OFB, Key, Salt, BArray())
    If Which = CAST5 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_CAST5_OFB, Key, Salt, BArray())
    If Which = DES Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_DES_OFB, Key, Salt, BArray())
    If Which = DESE Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_DESE_OFB, Key, Salt, BArray())
    If Which = IDEA Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_IDEA_OFB, Key, Salt, BArray())
    If Which = TRIPLEDES Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_DES3_OFB, Key, Salt, BArray())
    If Which = TWOFISH Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_TWOFISH_CBC, Key, Salt, BArray())
    If Which = RC4 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_RC4, Key, Salt, BArray())
    If Which = RIJNDAEL128 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_128, Key, Salt, BArray())
    If Which = RIJNDAEL192 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_192, Key, Salt, BArray())
    If Which = RIJNDAEL256 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, Key, Salt, BArray())
    If Which = SERPENT128 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_128, Key, Salt, BArray())
    If Which = SERPENT192 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_192, Key, Salt, BArray())
    If Which = SERPENT256 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_256, Key, Salt, BArray())
    
    Key = "": Salt = "": CipherText = ""
    Exit Function
    
ErrorHandler:
MsgBox Err.Description
Key = "": Salt = "": CipherText = ""
End Function

Private Sub Initialize()
    m_bytIndex(0) = 65 'Asc("A")
    m_bytIndex(1) = 66 'Asc("B")
    m_bytIndex(2) = 67 'Asc("C")
    m_bytIndex(3) = 68 'Asc("D")
    m_bytIndex(4) = 69 'Asc("E")
    m_bytIndex(5) = 70 'Asc("F")
    m_bytIndex(6) = 71 'Asc("G")
    m_bytIndex(7) = 72 'Asc("H")
    m_bytIndex(8) = 73 'Asc("I")
    m_bytIndex(9) = 74 'Asc("J")
    m_bytIndex(10) = 75 'Asc("K")
    m_bytIndex(11) = 76 'Asc("L")
    m_bytIndex(12) = 77 'Asc("M")
    m_bytIndex(13) = 78 'Asc("N")
    m_bytIndex(14) = 79 'Asc("O")
    m_bytIndex(15) = 80 'Asc("P")
    m_bytIndex(16) = 81 'Asc("Q")
    m_bytIndex(17) = 82 'Asc("R")
    m_bytIndex(18) = 83 'Asc("S")
    m_bytIndex(19) = 84 'Asc("T")
    m_bytIndex(20) = 85 'Asc("U")
    m_bytIndex(21) = 86 'Asc("V")
    m_bytIndex(22) = 87 'Asc("W")
    m_bytIndex(23) = 88 'Asc("X")
    m_bytIndex(24) = 89 'Asc("Y")
    m_bytIndex(25) = 90 'Asc("Z")
    m_bytIndex(26) = 97 'Asc("a")
    m_bytIndex(27) = 98 'Asc("b")
    m_bytIndex(28) = 99 'Asc("c")
    m_bytIndex(29) = 100 'Asc("d")
    m_bytIndex(30) = 101 'Asc("e")
    m_bytIndex(31) = 102 'Asc("f")
    m_bytIndex(32) = 103 'Asc("g")
    m_bytIndex(33) = 104 'Asc("h")
    m_bytIndex(34) = 105 'Asc("i")
    m_bytIndex(35) = 106 'Asc("j")
    m_bytIndex(36) = 107 'Asc("k")
    m_bytIndex(37) = 108 'Asc("l")
    m_bytIndex(38) = 109 'Asc("m")
    m_bytIndex(39) = 110 'Asc("n")
    m_bytIndex(40) = 111 'Asc("o")
    m_bytIndex(41) = 112 'Asc("p")
    m_bytIndex(42) = 113 'Asc("q")
    m_bytIndex(43) = 114 'Asc("r")
    m_bytIndex(44) = 115 'Asc("s")
    m_bytIndex(45) = 116 'Asc("t")
    m_bytIndex(46) = 117 'Asc("u")
    m_bytIndex(47) = 118 'Asc("v")
    m_bytIndex(48) = 119 'Asc("w")
    m_bytIndex(49) = 120 'Asc("x")
    m_bytIndex(50) = 121 'Asc("y")
    m_bytIndex(51) = 122 'Asc("z")
    m_bytIndex(52) = 48 'Asc("0")
    m_bytIndex(53) = 49 'Asc("1")
    m_bytIndex(54) = 50 'Asc("2")
    m_bytIndex(55) = 51 'Asc("3")
    m_bytIndex(56) = 52 'Asc("4")
    m_bytIndex(57) = 53 'Asc("5")
    m_bytIndex(58) = 54 'Asc("6")
    m_bytIndex(59) = 55 'Asc("7")
    m_bytIndex(60) = 56 'Asc("8")
    m_bytIndex(61) = 57 'Asc("9")
    m_bytIndex(62) = 43 'Asc("+")
    m_bytIndex(63) = 47 'Asc("/")
    m_bytReverseIndex(65) = 0 'Asc("A")
    m_bytReverseIndex(66) = 1 'Asc("B")
    m_bytReverseIndex(67) = 2 'Asc("C")
    m_bytReverseIndex(68) = 3 'Asc("D")
    m_bytReverseIndex(69) = 4 'Asc("E")
    m_bytReverseIndex(70) = 5 'Asc("F")
    m_bytReverseIndex(71) = 6 'Asc("G")
    m_bytReverseIndex(72) = 7 'Asc("H")
    m_bytReverseIndex(73) = 8 'Asc("I")
    m_bytReverseIndex(74) = 9 'Asc("J")
    m_bytReverseIndex(75) = 10 'Asc("K")
    m_bytReverseIndex(76) = 11 'Asc("L")
    m_bytReverseIndex(77) = 12 'Asc("M")
    m_bytReverseIndex(78) = 13 'Asc("N")
    m_bytReverseIndex(79) = 14 'Asc("O")
    m_bytReverseIndex(80) = 15 'Asc("P")
    m_bytReverseIndex(81) = 16 'Asc("Q")
    m_bytReverseIndex(82) = 17 'Asc("R")
    m_bytReverseIndex(83) = 18 'Asc("S")
    m_bytReverseIndex(84) = 19 'Asc("T")
    m_bytReverseIndex(85) = 20 'Asc("U")
    m_bytReverseIndex(86) = 21 'Asc("V")
    m_bytReverseIndex(87) = 22 'Asc("W")
    m_bytReverseIndex(88) = 23 'Asc("X")
    m_bytReverseIndex(89) = 24 'Asc("Y")
    m_bytReverseIndex(90) = 25 'Asc("Z")
    m_bytReverseIndex(97) = 26 'Asc("a")
    m_bytReverseIndex(98) = 27 'Asc("b")
    m_bytReverseIndex(99) = 28 'Asc("c")
    m_bytReverseIndex(100) = 29 'Asc("d")
    m_bytReverseIndex(101) = 30 'Asc("e")
    m_bytReverseIndex(102) = 31 'Asc("f")
    m_bytReverseIndex(103) = 32 'Asc("g")
    m_bytReverseIndex(104) = 33 'Asc("h")
    m_bytReverseIndex(105) = 34 'Asc("i")
    m_bytReverseIndex(106) = 35 'Asc("j")
    m_bytReverseIndex(107) = 36 'Asc("k")
    m_bytReverseIndex(108) = 37 'Asc("l")
    m_bytReverseIndex(109) = 38 'Asc("m")
    m_bytReverseIndex(110) = 39 'Asc("n")
    m_bytReverseIndex(111) = 40 'Asc("o")
    m_bytReverseIndex(112) = 41 'Asc("p")
    m_bytReverseIndex(113) = 42 'Asc("q")
    m_bytReverseIndex(114) = 43 'Asc("r")
    m_bytReverseIndex(115) = 44 'Asc("s")
    m_bytReverseIndex(116) = 45 'Asc("t")
    m_bytReverseIndex(117) = 46 'Asc("u")
    m_bytReverseIndex(118) = 47 'Asc("v")
    m_bytReverseIndex(119) = 48 'Asc("w")
    m_bytReverseIndex(120) = 49 'Asc("x")
    m_bytReverseIndex(121) = 50 'Asc("y")
    m_bytReverseIndex(122) = 51 'Asc("z")
    m_bytReverseIndex(48) = 52 'Asc("0")
    m_bytReverseIndex(49) = 53 'Asc("1")
    m_bytReverseIndex(50) = 54 'Asc("2")
    m_bytReverseIndex(51) = 55 'Asc("3")
    m_bytReverseIndex(52) = 56 'Asc("4")
    m_bytReverseIndex(53) = 57 'Asc("5")
    m_bytReverseIndex(54) = 58 'Asc("6")
    m_bytReverseIndex(55) = 59 'Asc("7")
    m_bytReverseIndex(56) = 60 'Asc("8")
    m_bytReverseIndex(57) = 61 'Asc("9")
    m_bytReverseIndex(43) = 62 'Asc("+")
    m_bytReverseIndex(47) = 63 'Asc("/")
End Sub

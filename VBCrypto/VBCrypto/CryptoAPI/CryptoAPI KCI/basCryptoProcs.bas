Attribute VB_Name = "basCryptoProcs"
Option Explicit

  Public g_blnCaseSensitiveUserID  As Boolean
  Public g_blnCaseSensitivePWord   As Boolean
  Public g_blnEnhancedProvider     As Boolean
  Public g_intHashType             As Integer
  
Public Function ConvertToArray(strInput As String) As Byte()

' ---------------------------------------------------------------------------
' convert data to byte array
' ---------------------------------------------------------------------------
  Dim cCrypto   As CryptKci.clsCryptoAPI
  Set cCrypto = New CryptKci.clsCryptoAPI
  
  ConvertToArray = cCrypto.StringToByteArray(strInput)
  Set cCrypto = Nothing

End Function

Public Function Correct_Password_Length(arPWord() As Byte) As Boolean

' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-DEC-2000  Kenneth Ives  kenaso@home.com
'              Wrote routine
' 21-JAN-2001  Kenneth Ives
'              Freed cCrypto class from memory
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intLength  As Integer
  Dim strPWord   As String
  Dim cCrypto    As CryptKci.clsCryptoAPI
  
' ---------------------------------------------------------------------------
' Convert password from byte array to string data
' ---------------------------------------------------------------------------
  Set cCrypto = New CryptKci.clsCryptoAPI
  strPWord = cCrypto.ByteArrayToString(arPWord())
  intLength = Len(strPWord)
  Set cCrypto = Nothing
  
' ---------------------------------------------------------------------------
' check length of password
' ---------------------------------------------------------------------------
  If intLength = 0 Then
      MsgBox "A Password / Passphrase must be entered.", _
             vbInformation Or vbOKOnly, "Password / Passphrase missing"
      Correct_Password_Length = False
      Set cCrypto = Nothing
      Exit Function
  End If
        
  If intLength < 8 Then
      ' If not a valid length
      MsgBox "Password / Passphrase must be a minimum length of eight(8) characters.", _
             vbInformation Or vbOKOnly, "Invalid Password / Passphrase length"
      Correct_Password_Length = False
      Set cCrypto = Nothing
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' if we got to here we were successful
' ---------------------------------------------------------------------------
  Correct_Password_Length = True
  Set cCrypto = Nothing
  
End Function

Public Function CurrentSettings_Get(strKey As String) As Variant

' ---------------------------------------------------------------------------
' Get current settings from registry located at
' HKEY_CURRENT_USER\Software\VB and VBA Settings
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 10-DEC-2000  Kenneth Ives              Written by kenaso@home.com
' ---------------------------------------------------------------------------
  CurrentSettings_Get = GetSetting(APP_NAME, APP_SECTION, strKey)
  
End Function

Public Function CurrentSettings_Save(strKey As String, varValue As Variant) As String
  
' ---------------------------------------------------------------------------
' Save current settings to the registry located at
' HKEY_CURRENT_USER\Software\VB and VBA Settings
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 10-DEC-2000  Kenneth Ives              Written by kenaso@home.com
' ---------------------------------------------------------------------------
  SaveSetting APP_NAME, APP_SECTION, strKey, varValue

End Function

Public Sub Initial_settings()

' ---------------------------------------------------------------------------
' See if there are any settings in the registry.  If not, then insert them in
' in the registry.
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 10-DEC-2000  Kenneth Ives              Written by kenaso@home.com
' ---------------------------------------------------------------------------
  
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim varValue As Variant
  
' ---------------------------------------------------------------------------
' Case sensitive User ID setting (Default = True)
' ---------------------------------------------------------------------------
  varValue = CurrentSettings_Get("UserID")
  
  ' if nothing of file, write default to the registry
  If Len(Trim$(varValue)) = 0 Then
      g_blnCaseSensitiveUserID = True
      CurrentSettings_Save "UserID", g_blnCaseSensitiveUserID
  Else
      g_blnCaseSensitiveUserID = CBool(varValue)
  End If
  
' ---------------------------------------------------------------------------
' Case sensitive Password / Passphrase setting (Default = True)
' ---------------------------------------------------------------------------
  varValue = CurrentSettings_Get("Password")
  
  ' if nothing of file, write default to the registry
  If Len(Trim$(varValue)) = 0 Then
      g_blnCaseSensitivePWord = True
      CurrentSettings_Save "Password", g_blnCaseSensitivePWord
  Else
      g_blnCaseSensitivePWord = CBool(varValue)
  End If
  
' ---------------------------------------------------------------------------
' Whether or not to use the Enhanced Provider
' ---------------------------------------------------------------------------
  varValue = CurrentSettings_Get("EnhancedProvider")
  
  ' if nothing of file, write default to the registry
  If Len(Trim$(varValue)) = 0 Then
      g_blnEnhancedProvider = False
      CurrentSettings_Save "EnhancedProvider", g_blnEnhancedProvider
  Else
      g_blnEnhancedProvider = CBool(varValue)
  End If
  
' ---------------------------------------------------------------------------
' Hash method (Default = MD5)
' ---------------------------------------------------------------------------
  varValue = CurrentSettings_Get("HashMethod")
  
  ' if nothing of file, write default to the registry
  If Len(Trim$(varValue)) = 0 Then
      g_intHashType = 2
      CurrentSettings_Save "HashMethod", g_intHashType
  Else
      g_intHashType = CInt(varValue)
  End If
  
End Sub

Public Function Same_As_Previous(arByte1() As Byte, arByte2() As Byte) As Boolean

' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-DEC-2000  Kenneth Ives  kenaso@home.com
'              Wrote routine
' 21-JAN-2001  Kenneth Ives
'              Freed cCrypto class from memory
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Define variables
' ---------------------------------------------------------------------------
  Dim strTmp1   As String
  Dim strTmp2   As String
  Dim cCrypto   As CryptKci.clsCryptoAPI
  Set cCrypto = New CryptKci.clsCryptoAPI
  
' ---------------------------------------------------------------------------
' Convert byte arrays to string data
' ---------------------------------------------------------------------------
  strTmp1 = cCrypto.ByteArrayToString(arByte1())
  strTmp2 = cCrypto.ByteArrayToString(arByte2())
  
' ---------------------------------------------------------------------------
' Make the comparisons to see if these two arrays are the same
' ---------------------------------------------------------------------------
  If StrComp(strTmp1, strTmp2, vbBinaryCompare) = 0 Then
      Same_As_Previous = True
  Else
      Same_As_Previous = False
  End If

' ---------------------------------------------------------------------------
' Empty data strings
' ---------------------------------------------------------------------------
  strTmp1 = String$(250, 0)
  strTmp2 = String$(250, 0)
  Set cCrypto = Nothing
  
End Function

Public Function Validate_Password(arUserID() As Byte, _
                                  arPWord() As Byte) As Boolean

' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-DEC-2000  Kenneth Ives  kenaso@home.com
'              Wrote routine
' 21-JAN-2001  Kenneth Ives
'              Freed cCrypto class from memory
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strTmpUserID  As String
  Dim strUserID     As String
  Dim strPWord      As String
  Dim strSalt       As String
  Dim strHash       As String
  Dim strHashDB     As String
  Dim strTmp        As String
  Dim cCrypto       As CryptKci.clsCryptoAPI
  
  Set cCrypto = New CryptKci.clsCryptoAPI
  
' ---------------------------------------------------------------------------
' Get the data on file
' ---------------------------------------------------------------------------
  If Query_User(arUserID(), strSalt, strHashDB) Then
      
      ' convert User ID from byte array to string
      strTmpUserID = cCrypto.ByteArrayToString(arUserID())
  
      ' Convert password array to string data
      strPWord = cCrypto.ByteArrayToString(arPWord())
                   
      ' Hash the user ID after appending the default password to it.
      ' Use MD5 hashing algorithm.
      strUserID = cCrypto.CreateHash(strTmpUserID, g_intHashType, True, , True)
      
      ' Build the hashed results by concatenating the user supplied password,
      ' the randomly generated salt value, and the default password.  Use
      ' SHA-1 as the hashing algorithm.
      strHash = cCrypto.CreateHash(strPWord & strSalt, g_intHashType, True, , True)
      
      ' Compare the results we just created with the results
      ' in the database.  Use a binary compare because these
      ' must match perfectly.
      If StrComp(strHashDB, strHash, vbBinaryCompare) = 0 Then
          Validate_Password = True    ' we have a match
      Else
          ' Wrong password entered
          MsgBox "Password / Passphrase invalid.", _
                 vbExclamation Or vbOKOnly, "Invalid data"
          Validate_Password = False
      End If
  Else
      Validate_Password = False  ' user not in database
  End If

  Set cCrypto = Nothing

End Function


Attribute VB_Name = "basADODB"
Option Explicit

' ***************************************************************************
' Module:        basADO.bas
'
' Description:   This module is used to access a password database on a
'                server.  Normally, this database not is the user's path nor
'                do they have access to that area.  Whenever a user logs onto
'                the network, a server application within the logon script is
'                executed.  This server application has the authority to get
'                to the database after capturing the logon data from the user.
'
'                Always give credit where credit is due.  If you attach your
'                creditials to a piece of code, you should be available to
'                answer questions concerning that code.
'
' Thanks to:     John Cunningham  http://users.ids.net/~johnpc/
'                For his VB addin to generate the ADO code you see in parts
'                of this module.
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-DEC-2000  Kenneth Ives  kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Be sure to add a Reference to "MS ActiveX Data Objects 2.x Library"
' to this project.
' ---------------------------------------------------------------------------
  Private connPWD  As ADODB.Connection    ' Connect to the ADO Data Type
  Private rsPWord  As ADODB.Recordset     ' Record Source Name
  
' ---------------------------------------------------------------------------
' Global type structure
' ---------------------------------------------------------------------------
  Public Type Data_Record
      Number     As String    ' record number
      UserID     As String    ' hashed user ID
      Salt       As String    ' Random generated salt value
      Result     As String    ' Hashed password/passphrase
      Timestamp  As String    ' date/time of last update
  End Type

Public Function GetAllRecords(ByRef DR() As Data_Record) As Boolean

' ***************************************************************************
' Routine:       GetAllRecords
'
' Description:   Query the password database and return the user ID (in hashed
'                format), salt value, and the hashed results.
'
' Parameters:    DR() - Data record array in which to return the data
'
' Returns:       All the records
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-DEC-2000  Kenneth Ives  kenaso@home.com
'              Wrote routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngRecCount  As Long
  Dim lngIndex     As Long
  Dim strSQL       As String
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  lngIndex = 0
  strSQL = "SELECT * FROM [PWord] ORDER BY [UserID] ASC"
  
  On Error GoTo GetAllRecords_Error
' ---------------------------------------------------------------------------
' Open the password database to validate the UserID
' ---------------------------------------------------------------------------
  Open_connPWD
  
' ---------------------------------------------------------------------------
' Get the data
' ---------------------------------------------------------------------------
  Set rsPWord = New ADODB.Recordset
  rsPWord.Open strSQL, connPWD, adOpenStatic, adLockOptimistic, adCmdText
  lngRecCount = rsPWord.RecordCount  ' save the record count
  
' ---------------------------------------------------------------------------
' see if the User ID is on file
' ---------------------------------------------------------------------------
  If lngRecCount < 1 Then
      GoTo GetAllRecords_Error
  Else
      ReDim DR(lngRecCount)  ' resize the array
      
      Do While Not rsPWord.EOF
          DR(lngIndex).Number = CStr(rsPWord.Bookmark)
          DR(lngIndex).UserID = rsPWord!UserID
          DR(lngIndex).Salt = rsPWord!Salt
          DR(lngIndex).Result = rsPWord!Result
          DR(lngIndex).Timestamp = rsPWord!Timestamp
          
          lngIndex = lngIndex + 1     ' increment array index
          rsPWord.MoveNext            ' get the next record
      Loop
      
      GetAllRecords = True
  End If
  
CleanUp:
  rsPWord.Close        ' close the recordset
  connPWD.Close       ' close the database
  
Normal_Exit:
' ---------------------------------------------------------------------------
' free objects form memory
' ---------------------------------------------------------------------------
  Set rsPWord = Nothing
  Set connPWD = Nothing
  Exit Function


GetAllRecords_Error:
' ---------------------------------------------------------------------------
' Display an error message
' ---------------------------------------------------------------------------
  MsgBox "Error:  " & CStr(Err.Number) & " " & Err.Description & vbLf & _
         "Table is corrupted or empty.", vbOKOnly, "Reading Database"
  GetAllRecords = False
  GoTo Normal_Exit

End Function

Public Function AddNew_User(arUserID() As Byte, arPWord() As Byte) As Boolean

' ***************************************************************************
' Routine:       AddNew_User
'
' Description:   The user ID and the user supplied password is passed here
'                in a byte array and then converted to strings.  A unique
'                salt value is generated.  The user ID string is then hashed
'                using whatever hash algorithm was selected.  The password
'                and the salt value are concatenated and also hashed.  This
'                becomes the hashed results.  The salt value, hashed user ID,
'                and hashed results are then added to the database.  The date
'                timestamp is also added.
'
' Parameters:    arUserID() - byte array containing the user ID
'                arPWord() - byte array containing the user password
'
' Returns:       TRUE/FALSE based on the findings
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-DEC-2000  Kenneth Ives  kenaso@home.com
'              Wrote routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngRecCount   As Long
  Dim strTmpUserID  As String
  Dim strUserID     As String
  Dim strPWord      As String
  Dim strSalt       As String
  Dim strHash       As String
  Dim strSQL        As String
  Dim strTmp        As String
  Dim cCrypto       As CryptKci.clsCryptoAPI
  
' ---------------------------------------------------------------------------
' Convert User ID from byte array to string
' ---------------------------------------------------------------------------
  Set cCrypto = New CryptKci.clsCryptoAPI
  strTmpUserID = cCrypto.ByteArrayToString(arUserID())

' ---------------------------------------------------------------------------
' Convert password array to string data
' ---------------------------------------------------------------------------
  strPWord = cCrypto.ByteArrayToString(arPWord())
               
' ---------------------------------------------------------------------------
' Build the hashed user ID by using whatever hash algorithm was selected.
' ---------------------------------------------------------------------------
  strUserID = cCrypto.CreateHash(strTmpUserID, g_intHashType, True, , True)

' ---------------------------------------------------------------------------
' Create unique salt value 15 bytes long
' ---------------------------------------------------------------------------
  strSalt = cCrypto.CreateSaltValue(15)
  
' ---------------------------------------------------------------------------
' Build the hashed results by concatenating the user supplied password and
' the randomly generated salt value.  Use whatever hash algorithm was
' selected.
' ---------------------------------------------------------------------------
  strHash = cCrypto.CreateHash(strPWord & strSalt, g_intHashType, True, , True)

' ---------------------------------------------------------------------------
' build SQL statement
' ---------------------------------------------------------------------------
  strSQL = "SELECT * FROM [PWord]"
                           
  On Error GoTo AddNew_User_Error
' ---------------------------------------------------------------------------
' Open the password database to validate the UserID
' ---------------------------------------------------------------------------
  Open_connPWD
  
' ---------------------------------------------------------------------------
' Setup to add a new user
' ---------------------------------------------------------------------------
  Set rsPWord = New ADODB.Recordset
  rsPWord.Open strSQL, connPWD, adOpenStatic, adLockOptimistic, adCmdText
  
' ---------------------------------------------------------------------------
' Add the new user information to the database
' ---------------------------------------------------------------------------
  rsPWord.AddNew
  rsPWord!UserID = strUserID
  rsPWord!Salt = strSalt
  rsPWord!Result = strHash
  rsPWord!Timestamp = Now()
  rsPWord.Update
  AddNew_User = True

CleanUp:
  rsPWord.Close       ' close the recordset
  connPWD.Close       ' close the database
  
Normal_Exit:
' ---------------------------------------------------------------------------
' free objects form memory
' ---------------------------------------------------------------------------
  Set rsPWord = Nothing
  Set connPWD = Nothing
  Set cCrypto = Nothing
  Exit Function


AddNew_User_Error:
' ---------------------------------------------------------------------------
' Display an error message
' ---------------------------------------------------------------------------
  MsgBox "Error:  " & CStr(Err.Number) & " " & Err.Description & vbLf & _
         "User [ " & strTmpUserID & " ] was not added.", _
         vbExclamation Or vbOKOnly, "Adding to Database"
  AddNew_User = False
  Resume Normal_Exit

End Function

Public Sub Open_connPWD()

' ***************************************************************************
' Routine:       Open_connPWD
'
' Description:   Use ADO to open the MS Access database.  This way, the user
'                does not need to have Access installed in order to run this
'                module.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-DEC-2000  Kenneth Ives  kenaso@home.com
'              Wrote routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim DbFile As String              ' Name of DataBase

' ---------------------------------------------------------------------------
' Set the Database Applicable Path
' ---------------------------------------------------------------------------
  DbFile = App.Path & "\PWD.mdb"

' ---------------------------------------------------------------------------
' Establish the Connection
' ---------------------------------------------------------------------------
  Set connPWD = New ADODB.Connection
  connPWD.CursorLocation = adUseClient
  connPWD.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                             "Data Source=" & DbFile & ";" & _
                             "Persist Security Info=False"

' ---------------------------------------------------------------------------
' Open the Connection.  Once this Connection is opened, it can be used
' throughout the application.
' ---------------------------------------------------------------------------
  connPWD.Open

End Sub

Public Function Query_User(arUserID() As Byte, _
                           strSalt As String, _
                           strHash As String, _
                  Optional blnAddQuery As Boolean = False) As Boolean

' ***************************************************************************
' Routine:       Query_User
'
' Description:   Query the password database sea5rching for a hashed user ID.
'                The user ID is passed here in a byte array and then
'                converted to string.  The string is then hashed using
'                whatever hash algorithm was selected.  The database is then
'                read.
'
' Parameters:    arUserID() - byte array containing the user ID
'                strSalt - is a return Value
'                strHash - is a return value
'                blnAddQuery - (Optional) TRUE/FALSE
'                              [Default] FALSE - We are NOT adding a new user
'                                     and if we find them in the database, we
'                                     want to return the salt and hashed values
'                              TRUE - We are adding a new user and only want
'                                     to return a TRUE or FALSE on whether or
'                                     not the user already exists.
'
' Returns:       strSalt and strHash values if user is in the database and
'                this is not an AddNew query.  Also, return a TRUE/FALSE based
'                on the findings.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-DEC-2000  Kenneth Ives  kenaso@home.com
'              Wrote routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngRecCount   As Long
  Dim strTmpUserID  As String
  Dim strUserID     As String
  Dim strSQL        As String
  Dim cCrypto       As CryptKci.clsCryptoAPI
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  strSalt = ""
  strHash = ""
  Set cCrypto = New CryptKci.clsCryptoAPI
  
' ---------------------------------------------------------------------------
' Convert User ID from byte array to string
' ---------------------------------------------------------------------------
  strTmpUserID = cCrypto.ByteArrayToString(arUserID())

' ---------------------------------------------------------------------------
' Build the hashed user ID by using whatever hash algorithm was selected.
' ---------------------------------------------------------------------------
  strUserID = cCrypto.CreateHash(strTmpUserID, g_intHashType, True, , True)

' ---------------------------------------------------------------------------
' Build SQL statement
' ---------------------------------------------------------------------------
  strSQL = "SELECT * FROM [PWord] Where [UserID] = '" & strUserID & "'"
                   
  On Error GoTo Query_User_Error
' ---------------------------------------------------------------------------
' Open the password database to validate the UserID
' ---------------------------------------------------------------------------
  Open_connPWD
  
' ---------------------------------------------------------------------------
' Get the data
' ---------------------------------------------------------------------------
  Set rsPWord = New ADODB.Recordset
  rsPWord.Open strSQL, connPWD, adOpenStatic, adLockOptimistic, adCmdText
  lngRecCount = rsPWord.RecordCount  ' save the record count
  
' ---------------------------------------------------------------------------
' see if the User ID is on file
' ---------------------------------------------------------------------------
  If lngRecCount < 1 Then
      Query_User = False
  Else
      If Not blnAddQuery Then
          ' Captaure the Salt and Hashed results
          ' to future comparisons
          strSalt = rsPWord!Salt
          strHash = rsPWord!Result
      End If
      
      Query_User = True
  End If
  
CleanUp:
  rsPWord.Close       ' close the recordset
  connPWD.Close       ' close the database
  
Normal_Exit:
' ---------------------------------------------------------------------------
' free objects form memory
' ---------------------------------------------------------------------------
  Set rsPWord = Nothing
  Set connPWD = Nothing
  Set cCrypto = Nothing
  Exit Function


Query_User_Error:
' ---------------------------------------------------------------------------
' Display an error message
' ---------------------------------------------------------------------------
  MsgBox "Error:  " & CStr(Err.Number) & " " & Err.Description & vbLf & _
         "User [ " & strTmpUserID & " ] was not found.", _
         vbExclamation Or vbOKOnly, "Querying Database"
  Query_User = False
  Resume Normal_Exit

End Function

Public Function Remove_User(arUserID() As Byte) As Boolean

' ***************************************************************************
' Routine:       Remove_User
'
' Description:   Query the password database searching for a hashed user ID.
'                The user ID is passed here in a byte array and then
'                converted to string.  The string is then hashed using
'                whatever hash algorithm was selected.  The database is then
'                read.  If the user is found, the record is deleted.
'
' Parameters:    arUserID() - byte array containing the user ID
'
' Returns:       TRUE/FALSE based on the findings
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-DEC-2000  Kenneth Ives  kenaso@home.com
'              Wrote routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngRecCount   As Long
  Dim strTmpUserID  As String
  Dim strUserID     As String
  Dim strSQL        As String
  Dim cCrypto       As CryptKci.clsCryptoAPI
  
' ---------------------------------------------------------------------------
' Convert User ID from byte array to string
' ---------------------------------------------------------------------------
  Set cCrypto = New CryptKci.clsCryptoAPI
  strTmpUserID = cCrypto.ByteArrayToString(arUserID())

' ---------------------------------------------------------------------------
' Build the hashed user ID by using whatever hash algorithm was selected.
' ---------------------------------------------------------------------------
  strUserID = cCrypto.CreateHash(strTmpUserID, g_intHashType, True, , True)

' ---------------------------------------------------------------------------
' build SQL statement
' ---------------------------------------------------------------------------
  strSQL = "SELECT * FROM [PWord] Where [UserID] = '" & strUserID & "'"
                           
  On Error GoTo Remove_User_Error
' ---------------------------------------------------------------------------
' Open the password database to validate the UserID
' ---------------------------------------------------------------------------
  Open_connPWD
  
' ---------------------------------------------------------------------------
' Get the data
' ---------------------------------------------------------------------------
  Set rsPWord = New ADODB.Recordset
  rsPWord.Open strSQL, connPWD, adOpenStatic, adLockOptimistic, adCmdText
  
' ---------------------------------------------------------------------------
' Remove this record from the database
' ---------------------------------------------------------------------------
  rsPWord.Delete
  rsPWord.Requery
  Remove_User = True

CleanUp:
  rsPWord.Close       ' close the recordset
  connPWD.Close       ' close the database
  
Normal_Exit:
' ---------------------------------------------------------------------------
' free objects form memory
' ---------------------------------------------------------------------------
  Set rsPWord = Nothing
  Set connPWD = Nothing
  Set cCrypto = Nothing
  Exit Function

Remove_User_Error:
' ---------------------------------------------------------------------------
' Display an error message
' ---------------------------------------------------------------------------
  MsgBox "Error:  " & CStr(Err.Number) & " " & Err.Description & vbLf & _
         "User [ " & strTmpUserID & _
         " ] was not removed.", _
         vbExclamation Or vbOKOnly, "Delete a user ID"
  Remove_User = False
  Resume Normal_Exit

End Function

Public Function Update_User(arUserID() As Byte, arPWord() As Byte) As Boolean

' ***************************************************************************
' Routine:       Update_User  (Password changes)
'
' Description:   The user ID and the user supplied password is passed here
'                in a byte array and then converted to strings.  A unique
'                salt value is generated.  The user ID string is then hashed
'                using whatever hash algorithm was selected.  The password
'                and the salt value are concatenated and also hashed.  This
'                becomes the hashed results.  The salt value and hashed
'                results are then added to the database.  The date timestamp
'                is also added.
'
' Parameters:    arUserID() - byte array containing the user ID
'                arPWord() - byte array containing the user password
'
' Returns:       TRUE/FALSE based on the findings
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-DEC-2000  Kenneth Ives  kenaso@home.com
'              Wrote routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngRecCount   As Long
  Dim strTmpUserID  As String
  Dim strUserID     As String
  Dim strPWord      As String
  Dim strSalt       As String
  Dim strHash       As String
  Dim strSQL        As String
  Dim strTmp        As String
  Dim cCrypto       As CryptKci.clsCryptoAPI
  
' ---------------------------------------------------------------------------
' Convert User ID from byte array to string
' ---------------------------------------------------------------------------
  Set cCrypto = New CryptKci.clsCryptoAPI
  strTmpUserID = cCrypto.ByteArrayToString(arUserID())

' ---------------------------------------------------------------------------
' Convert password array to string data
' ---------------------------------------------------------------------------
  strPWord = cCrypto.ByteArrayToString(arPWord())
               
' ---------------------------------------------------------------------------
' Build the hashed user ID by using whatever hash algorithm was selected.
' ---------------------------------------------------------------------------
  strUserID = cCrypto.CreateHash(strTmpUserID, g_intHashType, True, , True)

' ---------------------------------------------------------------------------
' Create unique salt value 15 bytes long
' ---------------------------------------------------------------------------
  strSalt = cCrypto.CreateSaltValue(15)
  
' ---------------------------------------------------------------------------
' Build the hashed results by concatenating the user supplied password and
' the randomly generated salt value.  Use whatever hash algorithm was
' selected.
' ---------------------------------------------------------------------------
  strHash = cCrypto.CreateHash(strPWord & strSalt, g_intHashType, True, , True)

' ---------------------------------------------------------------------------
' build SQL statement
' ---------------------------------------------------------------------------
  strSQL = "SELECT * FROM [PWord] Where [UserID] = '" & strUserID & "'"
                           
  On Error GoTo Update_User_Error
' ---------------------------------------------------------------------------
' Open the password database to validate the UserID
' ---------------------------------------------------------------------------
  Open_connPWD
  
' ---------------------------------------------------------------------------
' Get the data
' ---------------------------------------------------------------------------
  Set rsPWord = New ADODB.Recordset
  rsPWord.Open strSQL, connPWD, adOpenStatic, adLockOptimistic, adCmdText
  
' ---------------------------------------------------------------------------
' Add the new user information to the database
' ---------------------------------------------------------------------------
  rsPWord!Salt = strSalt
  rsPWord!Result = strHash
  rsPWord!Timestamp = Now()
  rsPWord.Update
  Update_User = True

CleanUp:
  rsPWord.Close       ' close the recordset
  connPWD.Close       ' close the database
  
Normal_Exit:
' ---------------------------------------------------------------------------
' free objects form memory
' ---------------------------------------------------------------------------
  Set rsPWord = Nothing
  Set connPWD = Nothing
  Set cCrypto = Nothing
  Exit Function

Update_User_Error:
' ---------------------------------------------------------------------------
' Display an error message
' ---------------------------------------------------------------------------
  MsgBox "Error:  " & CStr(Err.Number) & " " & Err.Description & vbLf & _
         "User [ " & strTmpUserID & _
         " ] was not updated.", _
         vbExclamation Or vbOKOnly, "Updating Database"
  Update_User = False
  Resume Normal_Exit

End Function


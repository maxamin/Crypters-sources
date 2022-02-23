Attribute VB_Name = "basMain"
Option Explicit

' ***************************************************************************
' Module:        basMain.bas
'
' Description:   This module contains some of the most common routines I use
'                along with some that are just common to this application.
'
' CryptoAPI Demo using the registry
'
' This is freeware.  Since security is of the upmost these days,
' a tool such as this should assist you in protecting your data.
' This is well documented and should help you understand what is
' happening.  I have tried to give everyone credit on their code
' snippet contributions.  If you recognize something I missed, let
' me know and I will update that portion with your name and email
' address (I must have both).
'
' To begin with, I used a lot of screens to demonstrate each function.
' This is to better illustrate what is going on without getting lost in
' performing multiple functions within a single form.
'
' Next, I use a database for network security because the user would
' never have access to the directory where this database is located.
' Also, I doubt if they would recognize any of the data in it.
'
' **  Brief Overview  ***************
' Whenever a user logs onto a network, a server application is executed
' from within the login script.  This server application has the only
' access to the database as far as the user is concerned.  The user's
' logon data is extracted, manipulated and applied to the database for
' verification.  If the logon data is authenticated, the user is allowed
' onto the network.
' ***********************************
'
' This database is very limited as it is for demonstration purposes only.
'
' One of the things you could add to your code is the number of tries a
' user can make trying to remember their password. Add a couple of fields
' to the database to deny the user access for 15 minutes before being
' allowed to try again.  In other words, set a flag field in the database
' for a "1" or "0" and another field for the current timestamp.  If the
' user is locked out after three tries, a "1" is entered into the flag
' field and the system timestamp in the other.  Whenever the user attempts
' a logon, first see if there is a "1" in the flag field.  If so, then test
' to see if 15 minutes have elaspsed since the "1" was entered.  If 15
' minutes or more have elapsed then enter a "0" in the flag field and NULL
' in the timestamp field and continue processing.
'
' Using this scenario is a definite thorn in the side of individuals trying
' to gain unauthorized access to your system by way of brute force entry.
'
' If you are using local security (the user's workstation), you can
' apply these same principles for the Windows registry.  See the "Hash Test"
' and you will see where, in my opinion, a Message Digest (MDn) algorithm
' was probably used for registry entries.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-DEC-2000  Kenneth Ives  kenaso@home.com
'              Wrote module
' 10-JAN-2001  Kenneth Ives  kenaso@home.com
'              Converted data to byte array and then encrypt/decrypt the data.
'              For display purposes, I use a hex display because if an
'              encrypted character returned is a Null, then I would end up
'              with a null terminated string.  Everything after that null
'              would be ignored by the text box control and not displayed.
'              Therefore, when I would read from the text box to get the data
'              to decrypt, I would not have all the data.  Thanks to Haakan
'              Gustavsson for pointing me in the right direction.
'              See frmEncStrings(cmdChoice_Click)
' 18-JAN-2001  Kenneth Ives  kenaso@home.com
'              The decoded file was be one byte larger than the source.  To
'              fix this, subtract 1 from the file size to accomodate the zero
'              based array.  Fix suggested by Harbinder Gill hgill@altavista.net
'              See frmEncFiles(cmdChoice_Click)
' 20-JAN-2001  Kenneth Ives  kenaso@home.com
'              According to theory, whenever you leave a text box, the lost
'              focus event is supposed to fire.  I came upon multiple
'              instances where it did not.  It would happen when I pressed
'              the ENTER key while still inside the text box and executing
'              the command button.  I decided to move the lost focus logic
'              to the validate event and added a piece of code in the
'              keypress event to force the validate event to fire.  I could
'              had done the same with the lost focus but I had already moved
'              my code. See all password input forms.
'
'              Also found that when you use PUT to write a byte array to a
'              file, the last character is converted to a NULL.   To get
'              around this quirk, I converted the decrypted byte array to
'              a text string and then PUT it in the output file.
'              See frmEncFiles(cmdChoice_Click)
' ***************************************************************************

' ---------------------------------------------------------------------------
' Variables
' ---------------------------------------------------------------------------
  Private m_blnFoundApp            As Boolean
  Private m_intAppCount            As Integer
  Private m_strTargetTitle         As String
  Private PGM_EXE_TITLE            As String       ' Exe name w/o ext
  Public g_strVersion              As String
  
' ---------------------------------------------------------------------------
' Constants ("ThunderMain"    - VBx Development)
'           ("ThunderRT6Main" - VB6 Executable File)
' ---------------------------------------------------------------------------
  Public Const PGM_CLASS      As String = "ThunderRT6Main"
  Public Const PGM_NAME       As String = "CryptoAPI Demo"
  Public Const MAX_PATH       As Long = 260
  Public Const APP_NAME       As String = "PWDTest"
  Public Const APP_SECTION    As String = "Settings"
  Public Const MYNAME         As String = "Freeware by Kenneth Ives  kenaso@home.com"

' ---------------------------------------------------------------------------
' Declares
' ---------------------------------------------------------------------------
  ' The EnumWindows() function enumerates all top-level windows on the screen
  ' by passing the handle of each window, in turn, to an application-defined
  ' callback function. EnumWindows() continues until the last top-level window
  ' is enumerated or the callback function returns FALSE.
  Private Declare Function EnumWindows Lib "user32" _
          (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
  
  ' The GetWindowText() function copies the text of the specified window's
  ' (Parent) title bar (if it has one) into a buffer. If the specified window
  ' is a control, the text of the control is copied.
  Private Declare Function GetWindowText Lib "user32" _
          Alias "GetWindowTextA" (ByVal hWnd As Long, _
          ByVal lpString As String, ByVal cch As Long) As Long
  
  ' The GetClassName() function retrieves the name of the class to which the
  ' specified window belongs.
  Private Declare Function GetClassName Lib "user32" _
          Alias "GetClassNameA" (ByVal hWnd As Long, _
          ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
          
  ' Always close a handle if not being used
  Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hObject As Long) As Long
          
Sub Main()

' ***************************************************************************
' Routine:       Main
'
' Description:   This is a generic routine to start an application
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@home.com
'              Wrote routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  ChDrive App.Path
  ChDir App.Path
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  g_strVersion = PGM_NAME & " v" & App.Major & "." & App.Minor
  PGM_EXE_TITLE = App.EXEName
  
' ---------------------------------------------------------------------------
' See if there is another instance of this program running.  The parameter
' being passed is the name of the EXE without the extension.
' ---------------------------------------------------------------------------
  If AlreadyRunning(PGM_EXE_TITLE) Then
      End
  End If

' ---------------------------------------------------------------------------
' Get the initial settings.
' ---------------------------------------------------------------------------
  Initial_settings
  
' ---------------------------------------------------------------------------
' Load all the screens.  If these were intensive forms, I would use a splash
' screen to keep the user occupied while they loaded into memory.
' ---------------------------------------------------------------------------
  Load_All_Forms
  
End Sub

Private Sub Load_All_Forms()

' ---------------------------------------------------------------------------
' This routine will load all of my forms
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 22-FEB-2001  Kenneth Ives              Written by kenaso@home.com
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' List all forms here except splash form.  Make sure the main form is last.
' ---------------------------------------------------------------------------
  Load frmAddUser
  Load frmChgPass
  Load frmDB
  Load frmDelUser
  Load frmEncFiles
  Load frmEncStrings
  Load frmHash
  Load frmOptions
  Load frmRnd
  Load frmTestPWD
  Load frmMainMenu
  
End Sub

Public Sub TerminateApplication()

' ***************************************************************************
' Routine:       TerminateApplication
'
' Description:   This routine will performt he shutdown process for this
'                application.  If there are any global object/class (not
'                forms) they will be listed below and set to NOTHING so as
'                to free them from memory.  The last task is to unload
'                all form objects.  Then terminate this application.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@home.com
'              Wrote routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Set all global objects to nothing, if they were used in this application
' EXAMPLE:    Set g_objMyObj = Nothing
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Upload all forms from memory and terminate this application
' ---------------------------------------------------------------------------
  CloseOpenFiles
  UnloadAllForms
  End
  
End Sub

Public Function CloseOpenFiles() As Boolean
  
' ---------------------------------------------------------------------------
' Closes any files that were opened with an "Open" statement
' ---------------------------------------------------------------------------
  While FreeFile > 1
      Close #FreeFile - 1
  Wend

End Function

Private Sub UnloadAllForms()

' ***************************************************************************
' Routine:       TerminateApplication
'
' Description:   Unload all active forms associated with this application.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@home.com
'              Wrote routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim frm As Form
  
' ---------------------------------------------------------------------------
' Loop thru all the active forms associated with this application
' ---------------------------------------------------------------------------
  For Each frm In Forms
      frm.Hide            ' hide the form
      Unload frm          ' deactivate the form
      Set frm = Nothing   ' free form object from memory
                          ' (prevents memory fragmenting)
  Next
  
End Sub

Public Function AlreadyRunning(strAppTitle As String) As Boolean

' ***************************************************************************
' Routine:       AlreadyRunning
'
' Description:   This routine will set an external search flag to FALSE and
'                perform an enumeration of all active programs, either hidden,
'                minimized, or displayed.
'
' Parameters:    strAppTitle - partial/full name of application title to
'                              look for
'
' Returns:       TRUE/FALSE based on the findings.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@home.com
'              Wrote routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngRetCode As Long

' ---------------------------------------------------------------------------
' Intialize variables
' ---------------------------------------------------------------------------
  m_blnFoundApp = False
  m_intAppCount = 0

' ---------------------------------------------------------------------------
' Search all active applicatios to see if this one is already running
' ---------------------------------------------------------------------------
  m_strTargetTitle = StrConv(strAppTitle, vbLowerCase)
  Call EnumWindows(AddressOf FindApplication, &H0)

' ---------------------------------------------------------------------------
' Return TRUE/FALSE based on findings
' ---------------------------------------------------------------------------
  AlreadyRunning = m_blnFoundApp
  
End Function

Private Function FindApplication(ByVal lngHandle As Long, _
                                 Optional ByVal lngParam As Long = 0) As Long

' ***************************************************************************
' Routine:       FindApplication
'
' Description:   This routine will search ALL active programs running under
'                Windows, including the hidden and minimized.  It will
'                look for the parent name.  The partial/full title name will
'                will be used for the search pattern.
'
' Parameters:    lngHandle - Generic application handle to check all active
'                       programs
'                lngParam - Not used (but required for callbacks)
'
' Returns:       Sets an external flag to TRUE/FALSE based on the findings.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@home.com
'              Wrote routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngLength     As Long    ' Length of the title string
  Dim strClassName  As String  ' Class name after formatting
  Dim strAppTitle   As String  ' Parent title after formatting

' ---------------------------------------------------------------------------
' Initialize return areas with spaces
' ---------------------------------------------------------------------------
  strClassName = Space$(MAX_PATH)
  strAppTitle = Space$(MAX_PATH)
   
' ---------------------------------------------------------------------------
' Make API calls to get the class name
' ---------------------------------------------------------------------------
  Call GetClassName(lngHandle, strClassName, MAX_PATH)
  strClassName = Trim$(Left$(strClassName, Len(PGM_CLASS)))
  
' ---------------------------------------------------------------------------
' Make API calls to get the parent title.  Capture just the left portion for
' the exact number of characters as the search title.  We want an exact
' match on the name.
' ---------------------------------------------------------------------------
  lngLength = GetWindowText(lngHandle, strAppTitle, MAX_PATH)
  strAppTitle = StrConv(Left$(strAppTitle, lngLength), vbLowerCase)
  strAppTitle = LTrim$(strAppTitle)  ' remove all leading blanks
  
' ---------------------------------------------------------------------------
' See if the class name matches.  If it does then check the parent title.
' ---------------------------------------------------------------------------
  If StrComp(strClassName, PGM_CLASS, vbTextCompare) = 0 Then
             
      ' See if this is the right title. Since we may only have a
      ' partial title, then we have to do an Instr() compare.
      ' This is why we want to make sure that the search title is
      ' as unique as possible.
      If InStr(1, strAppTitle, m_strTargetTitle, vbTextCompare) > 0 Then
                
          ' increment the counter.
          m_intAppCount = m_intAppCount + 1
          
          ' If we find more than one occurance of this program
          ' then set the flag to TRUE and leave
          If m_intAppCount > 1 Then
              ' set the flag denoting that we have found a duplicate
              m_blnFoundApp = True
              Exit Function        ' Time to leave
          End If
      End If
  End If
  
' ---------------------------------------------------------------------------
' Continue searching.
' ---------------------------------------------------------------------------
  Call CloseHandle(lngHandle)      ' close the active handle
  FindApplication = 1              ' Set the flag for another interation

End Function


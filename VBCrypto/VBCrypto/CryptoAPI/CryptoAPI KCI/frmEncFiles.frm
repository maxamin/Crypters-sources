VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEncFiles 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4800
   ClientLeft      =   1560
   ClientTop       =   1845
   ClientWidth     =   5565
   Icon            =   "frmEncFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5565
   Begin VB.Frame Frame1 
      Caption         =   "Cipher Algorithm"
      Height          =   540
      Left            =   150
      TabIndex        =   13
      Top             =   825
      Width           =   5265
      Begin VB.OptionButton optCipher 
         Caption         =   "RC4"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   18
         Top             =   225
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton optCipher 
         Caption         =   "RC2"
         Height          =   240
         Index           =   1
         Left            =   1050
         TabIndex        =   17
         Top             =   225
         Width           =   765
      End
      Begin VB.OptionButton optCipher 
         Caption         =   "DES"
         Height          =   240
         Index           =   2
         Left            =   2025
         TabIndex        =   16
         Top             =   225
         Width           =   765
      End
      Begin VB.OptionButton optCipher 
         Caption         =   "3DES"
         Height          =   240
         Index           =   3
         Left            =   2925
         TabIndex        =   15
         Top             =   225
         Width           =   765
      End
      Begin VB.OptionButton optCipher 
         Caption         =   " 3DES-112"
         Height          =   240
         Index           =   4
         Left            =   3900
         TabIndex        =   14
         Top             =   225
         Width           =   1140
      End
   End
   Begin VB.TextBox txtData 
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   1755
      Width           =   5235
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   4980
      Top             =   3075
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtData 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Index           =   3
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3795
      Width           =   4635
   End
   Begin VB.TextBox txtData 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Index           =   2
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3135
      Width           =   4635
   End
   Begin VB.CommandButton cmdChoice 
      Height          =   360
      Index           =   1
      Left            =   4980
      Picture         =   "frmEncFiles.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2475
      Width           =   435
   End
   Begin VB.TextBox txtData 
      Height          =   315
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   2475
      Width           =   4635
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "&Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   3
      Top             =   4275
      Width           =   975
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4440
      TabIndex        =   4
      Top             =   4275
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Input a password / passphrase  (Default password used if left blank)"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   1515
      Width           =   5130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name and location of decrypted file"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   3555
      Width           =   2505
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name and location of encrypted file"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   2895
      Width           =   2505
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter full path\filename or browse for a file with the button on the right."
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   2235
      Width           =   5010
   End
   Begin VB.Label lblMyLabel 
      BackStyle       =   0  'Transparent
      Height          =   420
      Left            =   180
      TabIndex        =   6
      Top             =   4275
      Width           =   2925
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Test File Encryption"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5310
   End
End
Attribute VB_Name = "frmEncFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 30-DEC-2000  Kenneth Ives              Written by kenaso@home.com
' ---------------------------------------------------------------------------
' Define module level constants
' ---------------------------------------------------------------------------
  Private m_intCipher       As Integer  ' Added 09-Sep-2001 KCI
  Private m_strFilename     As String
  Private m_strEncryptName  As String
  Private m_strDecryptName  As String
  Private arData()          As Byte     ' added 08-Jan-2001 KCI
  Private arPWord()         As Byte     ' added 08-Jan-2001 KCI

Private Sub Process_File()

' ***************************************************************************
' Routine:       Process_File
'
' Description:   First, test to see if the file exists and it is not empty.
'                Then encrypt and decrypt the file.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 08-JAN-2001  Kenneth Ives  kenaso@home.com
'              Converted data to byte array and then encrypt/decrypt the data.
'              Resolves the erroneous displays I sometimes encounter.  Thanks
'              to Haakan Gustavsson for pointing me in the right direction.
' 18-JAN-2001  Kenneth Ives  kenaso@home.com
'              The decoded file wwas be one byte larger than the source.  To
'              fix this, subtract 1 from the file size to accomodate the zero
'              based array.
'              Fix suggested by Harbinder Gill  hgill@altavista.net
' 21-JAN-2001  Kenneth Ives  kenaso@home.com
'              Found that when you use PUT to write a byte array to a
'              file, the last character is converted to a NULL.   To get
'              around this quirk, I converted the decrypted byte array to
'              a text string and then PUT it in the output file.
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngFileSize  As Long
  Dim hFile        As Integer
  Dim strText      As String
  Dim cCrypto      As CryptKci.clsCryptoAPI
  
' ---------------------------------------------------------------------------
' Make sure that the file exists and is not empty.
' ---------------------------------------------------------------------------
  Set cCrypto = New CryptKci.clsCryptoAPI
  On Error Resume Next
  
  lngFileSize = FileLen(m_strFilename)
  
  If Err <> 0 Or lngFileSize = 0 Then
      MsgBox "Cannot locate: " & vbCrLf & _
             m_strFilename & vbCrLf & "or this is an empty file.", _
             vbOKOnly, "File not found"
      Clear_Variables
      Exit Sub
  End If
  On Error GoTo 0     ' nullify the previous "On Error"
              
  On Error GoTo Process_File_Errors
' ---------------------------------------------------------------------------
' resize the data array to accommodate the file contents
'
' For encrypting, leave one extra element in the array to handle the last
' NULL appended to the excrypted file
' ---------------------------------------------------------------------------
  ReDim arData(lngFileSize)
              
' ---------------------------------------------------------------------------
' Create empty receiving files
' ---------------------------------------------------------------------------
  hFile = FreeFile  ' get first free file handle
  Open m_strEncryptName For Output As #hFile
  Close #hFile
                  
  Open m_strDecryptName For Output As #hFile
  Close #hFile
                             
' ---------------------------------------------------------------------------
' load the byte array with the file contents from the input file using one
' command then close file.
' ---------------------------------------------------------------------------
  Open m_strFilename For Binary Access Read As #hFile
  Get hFile, , arData
  Close #hFile

' ---------------------------------------------------------------------------
' See if there is a password
' ---------------------------------------------------------------------------
  If Len(Trim$(txtData(0).Text)) = 0 Then
      ReDim arPWord(0)
  Else
      arPWord = cCrypto.StringToByteArray(txtData(0).Text)
      cCrypto.Password = arPWord()
  End If
              
' ---------------------------------------------------------------------------
' set up parameters prior to encryption
' ---------------------------------------------------------------------------
  cCrypto.InputData = arData()
  cCrypto.EnhancedProvider = g_blnEnhancedProvider
  
' ---------------------------------------------------------------------------
' Encrypt the data and return in a byte array
' ---------------------------------------------------------------------------
  If cCrypto.Encrypt(g_intHashType, m_intCipher) Then
      arData = cCrypto.OutputData
  Else
      GoTo CleanUp
  End If
  
' ---------------------------------------------------------------------------
' Write the encrypted data into the encrypted output file
' ---------------------------------------------------------------------------
  Open m_strEncryptName For Binary Access Write As #hFile
  Put hFile, , arData
  Close #hFile

' ---------------------------------------------------------------------------
' Empty data array and make sure we have the correct size to refill it.
'
' BUG:  The decoded file will be one byte larger than the source.  To fix
'       this, subtract 1 from the file size to accomodate the zero based array.
'
' Fix suggested by Harbinder Gill hgill@altavista.net
' ---------------------------------------------------------------------------
  lngFileSize = FileLen(m_strEncryptName)
  Erase arData()
  ReDim arData(lngFileSize - 1)
  
' ---------------------------------------------------------------------------
' Load the byte array with the file contents from the encrypted file using
' one command then close file.
' ---------------------------------------------------------------------------
  Open m_strEncryptName For Binary Access Read As #hFile
  Get hFile, , arData
  Close #hFile

' ---------------------------------------------------------------------------
' set up parameters prior to decryption
' ---------------------------------------------------------------------------
  cCrypto.Password = arPWord()
  cCrypto.InputData = arData()

' ---------------------------------------------------------------------------
' Decrypt the data input from the encrypted file.  Convert the final data
' back to string format before writing to the output file.  If the byte array
' was PUT into the decrypted file in one command, the last character
' would be converted to a NULL.
' ---------------------------------------------------------------------------
  If cCrypto.Decrypt(g_intHashType, m_intCipher) Then
      arData = cCrypto.OutputData
      strText = cCrypto.ByteArrayToString(arData())
  Else
      GoTo CleanUp
  End If
  
' ---------------------------------------------------------------------------
' Write the decrypted data into the output file.
' ---------------------------------------------------------------------------
  Open m_strDecryptName For Binary Access Write As #hFile
  Put hFile, , strText
  Close #hFile
  
  MsgBox "Successful Finish!" & vbCrLf & _
         "Use a text editor to veiw the file formats.", _
         vbInformation Or vbOKOnly, "Encrypt Files"
  
CleanUp:
  On Error GoTo 0         ' nullify the previous "On Error"
  Set cCrypto = Nothing   ' free class from memory
  Erase arData()          ' empty the data array
  strText = String$(250, 0)
  Exit Sub
  
Process_File_Errors:
' ---------------------------------------------------------------------------
' Display error message
' ---------------------------------------------------------------------------
  MsgBox "Error: " & CStr(Err.Number) & "  " & Err.Description & vbCrLf & vbCrLf & _
         "Module:  frmEncFiles" & vbCrLf & _
         "Routine:  Process_File", vbExclamation Or vbOKOnly, "Encrypt File Error"
  
  Call CloseOpenFiles
  Resume CleanUp
  
End Sub

Private Sub cmdChoice_Click(Index As Integer)

' ***************************************************************************
' Routine:       cmdChoice_Click
'
' Description:   Based on command button selected, perform string encryption
'                of return to the main menu.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 08-JAN-2001  Kenneth Ives  kenaso@home.com
'              Wrote routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Based on the button pressed
' ---------------------------------------------------------------------------
  Select Case Index
         
         Case 0
              ' if nothing there then leave
              If Len(txtData(1).Text) = 0 Then
                  Exit Sub
              End If
              
              ' encrypt the file
              Process_File
              
         ' browse for a file
         Case 1
              txtData(1).Text = FileOpen_Dialog
              
              If Len(Trim$(txtData(1).Text)) > 0 Then
                  Prep_Textboxes
              Else
                  txtData(2).Text = ""
                  txtData(3).Text = ""
              End If
  
         ' Cancel button was pressed.  Return to main menu.
         Case 2
              frmEncFiles.Hide
              frmMainMenu.Show
  End Select
  
End Sub


Private Function FileOpen_Dialog() As String

' ***************************************************************************
' Routine:       FileOpen_Dialog
'
' Description:   Opens the File Open dialog box so the user can browse for a
'                former report file.
'
' Returns:       Path and filename
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-NOV-2000  Kenneth Ives  kenaso@home.com
'              Routine created
' ***************************************************************************

  On Error GoTo FileOpen_Errhandler
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strFilename As String
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  strFilename = ""
  CD.CancelError = True
  
' ---------------------------------------------------------------------------
' Loop until user selects a valid file or presses CANCEL
' ---------------------------------------------------------------------------
  Do
      ' Setup and display the File Open dialog box
      With CD
           ' Set flags
           .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or _
                    cdlOFNLongNames Or cdlOFNPathMustExist
           
           .DialogTitle = "Browse for file to encrypt."
           ' Set filters
           .Filter = "All Files (*.*)|*.*"
           .ShowOpen                        ' Display the Open dialog box
           strFilename = .FileName          ' save the path & filename selected
      End With
  
  Loop While Len(strFilename) = 0
  
  FileOpen_Dialog = strFilename
  Exit Function
  
FileOpen_Errhandler:
' ---------------------------------------------------------------------------
' User pressed the Cancel button
' ---------------------------------------------------------------------------
  FileOpen_Dialog = ""
  Exit Function

End Function


Private Sub Form_Initialize()

' ---------------------------------------------------------------------------
' Center form on the screen.  I use this statement here because of a
' bug in the Form property "Startup Position".  In the VB IDE, under
' Tools\Options\Advanced, when you place a checkmark in the SDI
' Development Environment check box and set the form property to
' startup in the center of the screen, it works while in the IDE.
' Whenever you leave the IDE, the property reverts back to the default
' of 0-Manual.  This is a known bug with Microsoft.
' ---------------------------------------------------------------------------
  Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

End Sub

Private Sub Form_Load()

' ---------------------------------------------------------------------------
' Center the form caption
' ---------------------------------------------------------------------------
  Me.Caption = g_strVersion
  CenterCaption frmEncFiles

' ---------------------------------------------------------------------------
' Hide this form
' ---------------------------------------------------------------------------
  frmEncFiles.Hide

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' ---------------------------------------------------------------------------
' Based on the the unload code the system passes,
' we determine what to do
'
' Unloadmode codes
'     0 - Close from the control-menu box
'         or Upper right "X"
'     1 - Unload method from code elsewhere
'         in the application
'     2 - Windows Session is ending
'     3 - Task Manager is clostrIng the app
'     4 - MDI Parent is clostrIng
' ---------------------------------------------------------------------------
  Select Case UnloadMode
         
         Case 0: cmdChoice_Click 2 ' Return to the main menu
         Case 1: Exit Sub
         Case 2: TerminateApplication
         Case 3: TerminateApplication
         Case 4: TerminateApplication
  End Select
  
End Sub

Private Sub Prep_Textboxes()

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intPosition As Integer
  
' ---------------------------------------------------------------------------
' get path and filename from first text box
' ---------------------------------------------------------------------------
  m_strFilename = Trim$(txtData(1).Text)
          
' ---------------------------------------------------------------------------
' look for last period in the path\filename
' ---------------------------------------------------------------------------
  intPosition = InStrRev(m_strFilename, ".", Len(m_strFilename))
  m_strEncryptName = Left$(m_strFilename, intPosition) & "enc"
  m_strDecryptName = Left$(m_strFilename, intPosition) & "dec"
          
' ---------------------------------------------------------------------------
' place filenames in text boxes
' ---------------------------------------------------------------------------
  txtData(2).Text = m_strEncryptName
  txtData(3).Text = m_strDecryptName

End Sub
Private Sub Clear_Variables()
  
  Erase arData()
  
  m_strFilename = ""
  m_strEncryptName = ""
  m_strDecryptName = ""
  
  With frmEncFiles
       .txtData(1).Text = ""
       .txtData(2).Text = ""
       .txtData(3).Text = ""
  End With
  
End Sub
Public Sub Reset_frmEncfiles()

' ---------------------------------------------------------------------------
' Display the form
' ---------------------------------------------------------------------------
  Clear_Variables
  Erase arPWord()
  optCipher_Click 0
    
  With frmEncFiles
       .txtData(0).Text = ""
       .lblMyLabel = MYNAME
       .Show vbModeless
  End With
  
End Sub

Private Sub optCipher_Click(Index As Integer)

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intIndex As Integer
  Dim intMax   As Integer
  
' ---------------------------------------------------------------------------
' Determine number of accessable cipher options
' ---------------------------------------------------------------------------
  If g_blnEnhancedProvider Then
      intMax = 4
      optCipher(3).Enabled = True
      optCipher(3).Visible = True
      optCipher(4).Enabled = True
      optCipher(4).Visible = True
  Else
      intMax = 2
      optCipher(3).Visible = False
      optCipher(3).Enabled = False
      optCipher(4).Visible = False
      optCipher(4).Enabled = False
  End If
  
' ---------------------------------------------------------------------------
' Select the visible option selected
' ---------------------------------------------------------------------------
  For intIndex = 0 To intMax
      If intIndex = Index Then
          optCipher(intIndex).Value = True
          m_intCipher = Index + 1
      Else
          optCipher(intIndex).Value = False
      End If
  Next
  
End Sub

Private Sub txtData_LostFocus(Index As Integer)
  
' ---------------------------------------------------------------------------
' See if anything is in the filename text box
' ---------------------------------------------------------------------------
  If Len(Trim$(txtData(1).Text)) > 0 Then
      Prep_Textboxes
  Else
      txtData(2).Text = ""
      txtData(3).Text = ""
  End If
  
End Sub

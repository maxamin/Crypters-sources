VERSION 5.00
Begin VB.Form frmEncStrings 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6720
   ClientLeft      =   1410
   ClientTop       =   915
   ClientWidth     =   5565
   Icon            =   "frmEncStrings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   5565
   Begin VB.Frame Frame1 
      Caption         =   "Cipher Algorithm"
      Height          =   540
      Left            =   150
      TabIndex        =   12
      Top             =   750
      Width           =   5265
      Begin VB.OptionButton optCipher 
         Caption         =   " 3DES-112"
         Height          =   240
         Index           =   4
         Left            =   3900
         TabIndex        =   17
         Top             =   225
         Width           =   1140
      End
      Begin VB.OptionButton optCipher 
         Caption         =   "3DES"
         Height          =   240
         Index           =   3
         Left            =   2925
         TabIndex        =   16
         Top             =   225
         Width           =   765
      End
      Begin VB.OptionButton optCipher 
         Caption         =   "DES"
         Height          =   240
         Index           =   2
         Left            =   2025
         TabIndex        =   15
         Top             =   225
         Width           =   765
      End
      Begin VB.OptionButton optCipher 
         Caption         =   "RC2"
         Height          =   240
         Index           =   1
         Left            =   1125
         TabIndex        =   14
         Top             =   225
         Width           =   765
      End
      Begin VB.OptionButton optCipher 
         Caption         =   "RC4"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   13
         Top             =   225
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin VB.TextBox txtData 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1620
      Width           =   5295
   End
   Begin VB.TextBox txtData 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5340
      Width           =   5295
   End
   Begin VB.TextBox txtData 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4200
      Width           =   5295
   End
   Begin VB.TextBox txtData 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   1
      Left            =   150
      MaxLength       =   512
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2250
      Width           =   5295
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
      Left            =   3300
      TabIndex        =   2
      Top             =   6285
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
      Index           =   1
      Left            =   4380
      TabIndex        =   3
      Top             =   6285
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Input a password / passphrase  (Default password used if left blank)"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1380
      Width           =   4800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data decrypted from the encrypted data above"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   5100
      Width           =   5130
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   975
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   3180
      Width           =   5190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Input data to be encrypted"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   3600
   End
   Begin VB.Label lblMyLabel 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   180
      TabIndex        =   5
      Top             =   6285
      Width           =   2925
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Test String Encryption"
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
      Left            =   127
      TabIndex        =   4
      Top             =   120
      Width           =   5310
   End
End
Attribute VB_Name = "frmEncStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------
' Define module level variables
' ---------------------------------------------------------------------------
  Private arData()     As Byte     ' added 08-Jan-2001 KCI
  Private arPWord()    As Byte     ' added 08-Jan-2001 KCI
  Private m_intCipher  As Integer  ' Added 09-Sep-2001 KCI
  
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
' 10-JAN-2001  Kenneth Ives  kenaso@home.com
'              Converted data to byte array and then encrypt/decrypt the data.
'              For display purposes, I use a hex display because if an
'              encrypted character returned is a Null, then I would end up
'              with a null terminated string.  Everything after that null
'              would be ignored by the text box control and not displayed.
'              Therefore, when I would read from the text box to get the data
'              to decrypt, I would not have all the data. Thanks to
'              Haakan Gustavsson for pointing me in the right direction.
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strTmp As String
  
' ---------------------------------------------------------------------------
' Based on the button pressed
' ---------------------------------------------------------------------------
  Select Case Index
                  
         ' Test button was selected
         Case 0
              strTmp = ""
              txtData(2).Text = ""     ' encrypted data (Read only)
              txtData(3).Text = ""     ' decrypted data (Read only)
              frmEncStrings.Refresh
              DoEvents
              
              ' see if there is any data to encrypt
              If Len(Trim$(txtData(1).Text)) = 0 Then
                  txtData(1).SetFocus
                  Exit Sub
              End If
              
              ' Instantsiate the crypto class
              Dim cCrypto As CryptKci.clsCryptoAPI
              Set cCrypto = New CryptKci.clsCryptoAPI
              
              ' convert string data to byte data
              If Len(Trim$(txtData(0).Text)) = 0 Then
                  ReDim arPWord(0)
              Else
                  arPWord = cCrypto.StringToByteArray(txtData(0).Text)
                  cCrypto.Password = arPWord()
              End If
              
              arData = cCrypto.StringToByteArray(txtData(1).Text)
              
              ' set up parameters prior to encryption
              cCrypto.InputData = arData()
              cCrypto.EnhancedProvider = g_blnEnhancedProvider
              
              ' Converted data to byte array and then encrypt/decrypt
              ' the data.  For display purposes, I use a hex display
              ' because if an encrypted character returned is a Null,
              ' then I would end up with a null terminated string.
              ' Everything after that would be ignored and not displayed.
              ' Therefore, when I would read from the text box to
              ' get the data to decrypt, I would not have all the data.
              If cCrypto.Encrypt(g_intHashType, m_intCipher) Then
                  arData = cCrypto.OutputData
                  strTmp = cCrypto.ByteArrayToString(arData)
              Else
                  Set cCrypto = Nothing    ' Free the Crypto class from memory
                  Exit Sub
              End If
              
              ' see if something went wrong
              If Len(Trim$(strTmp)) = 0 Then
                  MsgBox "Algorithm not supported by this provider"
                  Set cCrypto = Nothing    ' Free the Crypto class from memory
                  Exit Sub
              End If
              
              txtData(2).Text = cCrypto.ConvertStringToHex(strTmp)
              
              ' Convert Hex data from the text box to string data
              ' then to a byte array.  The data is then decrypted
              ' and displayed the the bottom text box.
              strTmp = cCrypto.ConvertStringFromHex(txtData(2).Text)
              arData = cCrypto.StringToByteArray(strTmp)
              cCrypto.Password = arPWord()
              cCrypto.InputData = arData()
    
              ' Decrypt the data input from the encrypted text box
              If cCrypto.Decrypt(g_intHashType, m_intCipher) Then
                  arData = cCrypto.OutputData
                  txtData(3).Text = cCrypto.ByteArrayToString(arData)
              End If
      
              Set cCrypto = Nothing    ' Free the Crypto class from memory
              strTmp = String$(250, 0)  ' overwrite data in temp variable
              frmEncStrings.Refresh    ' refresh the display
              
         ' Cancel button was pressed.  Return to main menu.
         Case 1
              frmEncStrings.Hide
              frmMainMenu.Show
  End Select
  
End Sub

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
  CenterCaption frmEncStrings

' ---------------------------------------------------------------------------
' Hide this form
' ---------------------------------------------------------------------------
  frmEncStrings.Hide

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
         
         Case 0: cmdChoice_Click 1 ' Return to the main menu
         Case 1: Exit Sub
         Case 2: TerminateApplication
         Case 3: TerminateApplication
         Case 4: TerminateApplication
  End Select
  
End Sub

Public Sub Reset_frmEncStrings()

' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 30-DEC-2000  Kenneth Ives              Written by kenaso@home.com
' ---------------------------------------------------------------------------
' Display the form
' ---------------------------------------------------------------------------
  Dim strMsg As String
  
  Erase arData()
  Erase arPWord()
  
  strMsg = "For display purposes only.  The encrypted data is displayed in hex "
  strMsg = strMsg & "format because if there is a null character in the encrypted "
  strMsg = strMsg & "data, we end up with a null terminated string.  Thus all data "
  strMsg = strMsg & "after the NULL would be ignored because of the internal "
  strMsg = strMsg & "conversion of data to string format by the text box control."
  
  optCipher_Click 0
  
  With frmEncStrings
       .Label1(3) = strMsg
       .txtData(0) = ""     ' password / passphrase
       ' Data string to be processed
       .txtData(1) = "This is test data that will be encrypted and decrypted."
       .txtData(2) = ""     ' encrypted data (Read only)
       .txtData(3) = ""     ' decrypted data (Read only)
       .lblMyLabel = MYNAME
       .Show vbModeless
  End With
  
  txtData(0).SetFocus
  
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

Private Sub txtData_GotFocus(Index As Integer)

' ---------------------------------------------------------------------------
' Highlight all the text in the box
' ---------------------------------------------------------------------------
  Select Case Index
         Case 0, 1: SendKeys "{Home}{End}"
  End Select
  
End Sub

Private Sub txtData_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim CtrlDown    As Integer
  Dim PressedKey  As Integer
  
' ---------------------------------------------------------------------------
' Initialize  variables
' ---------------------------------------------------------------------------
  CtrlDown = (Shift And vbCtrlMask) > 0         ' Define control key
  
  If Len(Trim$(KeyCode)) > 0 Then
      ' Convert to uppercase
      PressedKey = CInt(Asc(StrConv(Chr$(KeyCode), vbUpperCase)))
  End If
    
' ---------------------------------------------------------------------------
' Check to see if it is okay to make changes
' ---------------------------------------------------------------------------
  If CtrlDown And PressedKey = vbKeyX Then      ' Ctrl + X was pressed
      Edit_Cut
  ElseIf CtrlDown And PressedKey = vbKeyA Then  ' Ctrl + A was pressed
      Select Case Index
             Case 0, 1: SendKeys "{Home}{End}"
      End Select
  ElseIf CtrlDown And PressedKey = vbKeyC Then  ' Ctrl + C was pressed
      Edit_Copy
  ElseIf CtrlDown And PressedKey = vbKeyV Then  ' Ctrl + V was pressed
      Edit_Paste
  ElseIf PressedKey = vbKeyDelete Then          ' Delete key was pressed
      Edit_Delete
  End If

End Sub

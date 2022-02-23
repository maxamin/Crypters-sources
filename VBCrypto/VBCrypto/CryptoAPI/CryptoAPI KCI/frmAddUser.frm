VERSION 5.00
Begin VB.Form frmAddUser 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3075
   ClientLeft      =   1560
   ClientTop       =   1845
   ClientWidth     =   5565
   Icon            =   "frmAddUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5565
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   180
      MaxLength       =   250
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2340
      Width           =   4095
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   180
      MaxLength       =   250
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1740
      Width           =   4095
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "&OK"
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
      Left            =   4500
      TabIndex        =   3
      Top             =   1860
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
      Left            =   4500
      TabIndex        =   4
      Top             =   2310
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   840
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1035
      Width           =   3435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Verify Password / Passphrase"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   2100
      Width           =   2115
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password / Passphrase"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   8
      Top             =   1500
      Width           =   2055
   End
   Begin VB.Label lblMyLabel 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   3525
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add  User"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5370
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   1140
      Width           =   540
   End
End
Attribute VB_Name = "frmAddUser"
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
' Define module level variables
' ---------------------------------------------------------------------------
  Private arUserID()   As Byte
  Private arPWord1()   As Byte
  Private arPWord2()   As Byte
  Private m_strUserID  As String
  Private m_strPWord1  As String
  Private m_strPWord2  As String

Private Sub cmdChoice_Click(Index As Integer)

' ---------------------------------------------------------------------------
' Define variables
' ---------------------------------------------------------------------------
  Dim strTmp1 As String
  Dim strTmp2 As String
  
' ---------------------------------------------------------------------------
' Based on the button pressed
' ---------------------------------------------------------------------------
  Select Case Index
         
         ' OK button was pressed.
         Case 0:
              ' Test length of user ID
              If Len(m_strUserID) = 0 Then
                  MsgBox "A user ID must be entered.", _
                         vbInformation Or vbOKOnly, "User ID missing"
                  txtPassword(0).SetFocus
                  Exit Sub
              Else
                  arUserID = ConvertToArray(m_strUserID)
              End If
                  
              ' Test length of New password
              If Len(m_strPWord1) = 0 Then
                  MsgBox "A password / passphrase must be entered.", _
                         vbInformation Or vbOKOnly, "Password / Passphrase missing"
                  Reset_PWD_Boxes
                  txtPassword(1).SetFocus
                  Exit Sub
              Else
                  arPWord1 = ConvertToArray(m_strPWord1)
              End If
                  
              ' Test length of new password data entered
              If Not Correct_Password_Length(arPWord1()) Then
                  Reset_PWD_Boxes
                  txtPassword(1).SetFocus
                  Exit Sub
              End If
             
              ' Test length of verify password
              If Len(m_strPWord2) = 0 Then
                  MsgBox "A verification password / passphrase must be entered.", _
                         vbInformation Or vbOKOnly, "Password / Passphrase missing"
                  txtPassword(2).SetFocus
                  Exit Sub
              Else
                  arPWord2 = ConvertToArray(m_strPWord2)
              End If
                  
              ' Test length of verification password data entered
              If Not Correct_Password_Length(arPWord2()) Then
                  Erase arPWord2()
                  m_strPWord2 = ""
                  txtPassword(2).Text = ""
                  txtPassword(2).SetFocus
                  Exit Sub
              End If
             
              ' Is this user on file?
              If Query_User(arUserID(), strTmp1, strTmp2) Then
                  MsgBox "User [ " & m_strUserID & _
                         " ] is already in the database.", _
                         vbInformation Or vbOKOnly, "Duplicate User ID"
                  
                  ' get rid of any data returned
                  strTmp1 = String$(250, 0)
                  strTmp2 = String$(250, 0)
                  Reset_frmAddUser
                  Exit Sub
              Else
                  ' get rid of any data returned
                  strTmp1 = String$(250, 0)
                  strTmp2 = String$(250, 0)
              End If
              
              ' see if the passwords are the same
              If Not Same_As_Previous(arPWord1(), arPWord2()) Then
                  MsgBox "Password / Passphrase entries are not the same.", _
                         vbInformation Or vbOKOnly, "Invalid data entered"
                  Reset_PWD_Boxes
                  txtPassword(1).SetFocus
                  Exit Sub
              End If
              
              ' Add the new user to the database
              If Not AddNew_User(arUserID(), arPWord1()) Then
                  ' Return to the main menu.
                  Reset_frmAddUser
                  frmAddUser.Hide
                  frmMainMenu.Show
                  Exit Sub
              End If
              
              ' We were sucessful
              MsgBox "User [ " & m_strUserID & _
                     " ] has been added to the database.", _
                     vbInformation Or vbOKOnly, "Update Successful"
              Reset_frmAddUser
              
         ' Cancel button was pressed.
         Case 1:
              Reset_frmAddUser
              frmAddUser.Hide
              frmMainMenu.Show
  End Select
  
End Sub

Private Sub ClearVariables()

' ---------------------------------------------------------------------------
' clear variables
' ---------------------------------------------------------------------------
  Erase arUserID()
  Erase arPWord1()
  Erase arPWord2()
  
  m_strUserID = String$(250, 0)
  m_strPWord1 = String$(250, 0)
  m_strPWord2 = String$(250, 0)
  
  m_strUserID = ""
  m_strPWord1 = ""
  m_strPWord2 = ""
  
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
  CenterCaption frmAddUser

' ---------------------------------------------------------------------------
' Hide this form
' ---------------------------------------------------------------------------
  frmAddUser.Hide

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' ---------------------------------------------------------------------------
' Empty variables
' ---------------------------------------------------------------------------
  ClearVariables
  
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
         
         Case 0: cmdChoice_Click 1   ' Return to the main menu
         Case 1: Exit Sub
         Case 2: TerminateApplication
         Case 3: TerminateApplication
         Case 4: TerminateApplication
  End Select
  
End Sub

Public Sub Reset_frmAddUser()

' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  ClearVariables
  
' ---------------------------------------------------------------------------
' Display the form
' ---------------------------------------------------------------------------
  With frmAddUser
       .txtPassword(0).Text = ""
       .txtPassword(1).Text = ""
       .txtPassword(2).Text = ""
       .lblMyLabel = MYNAME
       .Show vbModeless
       .Refresh
  End With
  
' ---------------------------------------------------------------------------
' place cursor in first text box
' ---------------------------------------------------------------------------
  txtPassword(0).SetFocus
  
End Sub

Private Sub Reset_PWD_Boxes()

' ---------------------------------------------------------------------------
' Empty password arrays and the bottom two text boxes
' ---------------------------------------------------------------------------
  Erase arPWord1()
  Erase arPWord2()
  
  txtPassword(1).Text = ""
  txtPassword(2).Text = ""
  
  m_strPWord1 = ""
  m_strPWord2 = ""
  
End Sub

Private Sub txtPassword_GotFocus(Index As Integer)
  
' ---------------------------------------------------------------------------
' Highlight all the text in the box
' ---------------------------------------------------------------------------
  SendKeys "{Home}{End}"
  
End Sub

Private Sub txtPassword_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim CtrlDown    As Integer
  Dim PressedKey  As Integer
  
' ---------------------------------------------------------------------------
' Initialize  variables
' ---------------------------------------------------------------------------
  CtrlDown = (Shift And vbCtrlMask) > 0   ' Define control key
  
  If Len(Trim$(KeyCode)) > 0 Then
      ' Convert to uppercase
      PressedKey = CInt(Asc(StrConv(Chr$(KeyCode), vbUpperCase)))
  End If
    
' ---------------------------------------------------------------------------
' Check to see if it is okay to make changes
' ---------------------------------------------------------------------------
  If CtrlDown And PressedKey = vbKeyX Then
      Edit_Cut            ' Ctrl + X was pressed
  ElseIf CtrlDown And PressedKey = vbKeyA Then
      SendKeys "{Home}{End}"
  ElseIf CtrlDown And PressedKey = vbKeyC Then
      Edit_Copy           ' Ctrl + C was pressed
  ElseIf CtrlDown And PressedKey = vbKeyV Then
      Edit_Paste          ' Ctrl + V was pressed
  ElseIf PressedKey = vbKeyDelete Then
      Edit_Delete         ' Delete key was pressed
  End If

End Sub

Private Sub txtPassword_KeyPress(Index As Integer, KeyAscii As Integer)

' ---------------------------------------------------------------------------
' If ENTER is pressed then nullify keystroke and press the OK button
' ---------------------------------------------------------------------------
  If KeyAscii = 13 Then
      KeyAscii = 0                        ' Nullify keystroke
      txtPassword_Validate Index, False   ' force validate event to fire
      cmdChoice_Click 0                   ' Press the OK button
      Exit Sub
  End If
  
' ---------------------------------------------------------------------------
' If TAB is pressed then nullify keystroke and TAB to the next tabstop
' ---------------------------------------------------------------------------
  If KeyAscii = 9 Then
      KeyAscii = 0
      SendKeys "{TAB}"
  End If
  
' ---------------------------------------------------------------------------
' Accept on valid characters
' ---------------------------------------------------------------------------
  Select Case KeyAscii
  
         ' backspace and other printable
         ' keyboard characters
         Case 8, 32 To 126:
              Exit Sub      ' Good input
              
         ' Bad input
         Case Else:
              KeyAscii = 0  ' Nullify keystroke
  End Select
  
End Sub

Private Sub txtPassword_Validate(Index As Integer, Cancel As Boolean)

' ---------------------------------------------------------------------------
' Validate Event - Occurs before the focus shifts to a (second) control that
' has its CausesValidation property set to True.  Also, I make sure the
' "Cancel" parameter is set to FALSE, otherwise, I could get trapped in a
' text box and not be allowed to select the exit button.
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Remove leading and trailing blank spaces
' ---------------------------------------------------------------------------
  Cancel = False
  txtPassword(Index).Text = Trim$(txtPassword(Index).Text)
  
' ---------------------------------------------------------------------------
' Initial test of the input data
' ---------------------------------------------------------------------------
  Select Case Index
  
         ' User ID
         Case 0:
              ' something may have changed
              Erase arUserID()
              m_strUserID = ""
              Reset_PWD_Boxes
              
              If Len(txtPassword(0).Text) > 0 Then
                  If Not g_blnCaseSensitiveUserID Then
                      txtPassword(0).Text = StrConv(txtPassword(0).Text, vbUpperCase)
                  End If
                  
                  m_strUserID = txtPassword(0).Text
              End If
              
         ' New Password
         Case 1:
              ' if something is in the password box, then convert from
              ' string data to a byte array and then fill the text box
              ' with 30 asteriks
              If Len(txtPassword(1).Text) > 0 Then
                  If Not g_blnCaseSensitivePWord Then
                      txtPassword(1).Text = StrConv(txtPassword(1).Text, vbUpperCase)
                  End If
                  
                  m_strPWord1 = txtPassword(1).Text
                  txtPassword(1).Text = String$(30, "*")
              Else
                  ' else empty both of the password text boxes
                  Reset_PWD_Boxes
              End If
  
         ' Verifiy password
         Case 2:
              ' if something is in the password box, then convert from
              ' string data to a byte array and then fill the text box
              ' with 30 asteriks
              If Len(txtPassword(2).Text) > 0 Then
                  If Not g_blnCaseSensitivePWord Then
                      txtPassword(2).Text = StrConv(txtPassword(2).Text, vbUpperCase)
                  End If
                  
                  m_strPWord2 = txtPassword(2).Text
                  txtPassword(2).Text = String$(30, "*")
              Else
                  ' else empty bottom password text box
                  txtPassword(2).Text = ""
                  m_strPWord2 = ""
                  Erase arPWord2()
              End If
  End Select
  
End Sub

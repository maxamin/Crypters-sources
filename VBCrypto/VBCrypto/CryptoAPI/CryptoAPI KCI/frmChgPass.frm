VERSION 5.00
Begin VB.Form frmChgPass 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3735
   ClientLeft      =   1560
   ClientTop       =   1845
   ClientWidth     =   5550
   Icon            =   "frmChgPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3735
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
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
      Index           =   1
      Left            =   4440
      TabIndex        =   5
      Top             =   2460
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   840
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1080
      Width           =   3435
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   180
      MaxLength       =   250
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   4095
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
      TabIndex        =   6
      Top             =   2925
      Width           =   975
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
      Left            =   4440
      TabIndex        =   4
      Top             =   2460
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   180
      MaxLength       =   250
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3000
      Width           =   4095
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   180
      MaxLength       =   250
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2400
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   12
      Top             =   1185
      Width           =   540
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Change User Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   180
      TabIndex        =   11
      Top             =   180
      Width           =   5250
   End
   Begin VB.Label lblMyLabel 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   240
      TabIndex        =   10
      Top             =   3420
      Width           =   3600
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Password / Passphrase (From here, click OK)"
      Height          =   255
      Left            =   225
      TabIndex        =   9
      Top             =   1560
      Width           =   4065
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Verify Password/Passphrase"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   3750
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password / Passphrase"
      Height          =   255
      Index           =   1
      Left            =   225
      TabIndex        =   7
      Top             =   2160
      Width           =   3825
   End
End
Attribute VB_Name = "frmChgPass"
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
  Private arPWord3()   As Byte
  Private m_strUserID  As String
  Private m_strPWord1  As String
  Private m_strPWord2  As String
  Private m_strPWord3  As String
  
Private Sub Button_Lock(blnLockout As Boolean)

  If blnLockout Then
      With frmChgPass
           ' the two OK buttons must be reset
           .cmdChoice(1).Enabled = False
           .cmdChoice(1).Visible = False
           
           .cmdChoice(0).Enabled = True
           .cmdChoice(0).Visible = True
        
           ' Enable top two text boxes
           ' and update their color
           .txtPassword(1).Enabled = True
           .txtPassword(1).BackColor = vbWhite
           .txtPassword(1).Text = ""
        
           ' disable bottom two text boxes
           .txtPassword(2).Text = ""
           .txtPassword(2).BackColor = vbBlack
           .txtPassword(2).Enabled = False
           .txtPassword(3).Text = ""
           .txtPassword(3).BackColor = vbBlack
           .txtPassword(3).Enabled = False
      End With
      
      ' Empty the password arrays
      Erase arPWord1()
      Erase arPWord2()
      Erase arPWord3()
      m_strPWord1 = ""
      m_strPWord2 = ""
      m_strPWord3 = ""
      m_strUserID = ""
      
  Else
      ' allow data input into bottom two text boxes
      With frmChgPass
           ' the two OK buttons must be reset
           .cmdChoice(0).Enabled = False
           .cmdChoice(0).Visible = False
           .cmdChoice(1).Enabled = True
           .cmdChoice(1).Visible = True
                    
           ' locak out the top two boxes
           .txtPassword(0).Enabled = False
           .txtPassword(0).BackColor = vbCyan
           .txtPassword(1).Enabled = False
           .txtPassword(1).BackColor = vbCyan
           
           ' Enable bottom two text boxes
           ' and update their color
           .txtPassword(2).Enabled = True
           .txtPassword(2).BackColor = vbWhite
           .txtPassword(2).Text = ""
           .txtPassword(3).Enabled = True
           .txtPassword(3).BackColor = vbWhite
           .txtPassword(3).Text = ""
      End With
  
      ' Empty the password arrays
      Erase arPWord2()
      Erase arPWord3()
      m_strPWord2 = ""
      m_strPWord3 = ""
  End If
  
End Sub

Private Sub Clear_Bottom_Boxes()

' ---------------------------------------------------------------------------
' Empty the variables
' ---------------------------------------------------------------------------
  Erase arPWord2()
  Erase arPWord3()
  
  txtPassword(2).Text = ""
  txtPassword(3).Text = ""
  
  m_strPWord2 = ""
  m_strPWord3 = ""

End Sub

Private Sub cmdChoice_Click(Index As Integer)

' ---------------------------------------------------------------------------
' Define variables
' ---------------------------------------------------------------------------
  Dim strTmp1    As String
  Dim strTmp2    As String
  Dim intIndex   As Integer
  Dim intResp    As Integer
  
' ---------------------------------------------------------------------------
' Based on the button pressed
' ---------------------------------------------------------------------------
  Select Case Index
         
         ' OK button was pressed.  Time to test the input.
         Case 0:
              ' Test data input
              If Len(m_strUserID) = 0 Then
                  MsgBox "A user ID must be entered.", _
                         vbInformation Or vbOKOnly, "User ID missing"
                  txtPassword(0).Text = ""
                  txtPassword(0).SetFocus
                  Exit Sub
              Else
                  arUserID = ConvertToArray(m_strUserID)
              End If
                  
              ' Test length of password
              If Len(m_strPWord1) = 0 Then
                  MsgBox "A password / passphrase must be entered.", _
                         vbInformation Or vbOKOnly, "Password / Passphrase missing"
                  txtPassword(1).Text = ""
                  txtPassword(1).SetFocus
                  Exit Sub
              Else
                  arPWord1 = ConvertToArray(m_strPWord1)
              End If
                  
              ' Test length of password data entered
              If Not Correct_Password_Length(arPWord1()) Then
                  txtPassword(1).Text = ""
                  txtPassword(1).SetFocus
                  Exit Sub
              End If
              
              ' Is this user on file?
              If Not Query_User(arUserID(), strTmp1, strTmp2) Then
                  MsgBox "User [ " & m_strUserID & _
                         " ] cannot be found in the database.", _
                         vbInformation Or vbOKOnly, "Invalid User ID"
                  
                  ' get rid of any data returned
                  strTmp1 = String$(250, 0)
                  strTmp2 = String$(250, 0)
                  Reset_frmChgPass
                  Exit Sub
              Else
                  ' get rid of any data returned
                  strTmp1 = String$(250, 0)
                  strTmp2 = String$(250, 0)
              End If
              
              ' Hash the new password and compare with the
              ' the hashed results in the database.
              If Not Validate_Password(arUserID(), arPWord1()) Then
                  txtPassword(1).Text = ""
                  txtPassword(1).SetFocus
                  Exit Sub
              End If
              
              ' Reset the two OK buttons and enable
              ' bottom two text boxes
              Button_Lock False
              txtPassword(2).SetFocus
              Exit Sub
              
         Case 1:
              ' verify something was entered into the
              ' "Change" password text box
              If Len(m_strPWord2) = 0 Then
                  Clear_Bottom_Boxes
                  txtPassword(2).SetFocus
                  Exit Sub
              Else
                  arPWord2 = ConvertToArray(m_strPWord2)
              End If
              
              ' make sure this is not the same password
              ' previously entered.
              If Same_As_Previous(arPWord1(), arPWord2()) Then
                  MsgBox "This password is currently in use." & vbCrLf & _
                         "Enter a new Password / Passphrase.", vbExclamation Or vbOKOnly, _
                         "Duplicate Password / Passphrase"
                  Clear_Bottom_Boxes
                  txtPassword(2).SetFocus
                  Exit Sub
              End If
              
              ' Validate the new input password
              If Not Correct_Password_Length(arPWord2()) Then
                  Clear_Bottom_Boxes
                  txtPassword(2).SetFocus
                  Exit Sub
              End If
              
              ' Verify that this is the password the user wants
              ' by having them enter the same new password a
              ' second time
              If Len(m_strPWord3) = 0 Then
                  Erase arPWord3()
                  m_strPWord3 = ""
                  txtPassword(3).Text = ""
                  txtPassword(3).SetFocus
                  Exit Sub
              Else
                  arPWord3 = ConvertToArray(m_strPWord3)
              End If
              
              ' Validate the verification password
              If Not Correct_Password_Length(arPWord3()) Then
                  Erase arPWord3()
                  m_strPWord3 = ""
                  txtPassword(3).Text = ""
                  txtPassword(3).SetFocus
                  Exit Sub
              End If
             
              ' See if the new entries match each other.  Since passwords
              ' are case sensitive, we do a binary compare.
              If Same_As_Previous(arPWord2(), arPWord3()) Then
                  ' update the database
                  If Not Update_User(arUserID(), arPWord3()) Then
                      Clear_Bottom_Boxes
                      txtPassword(2).SetFocus
                      Exit Sub
                  Else
                      MsgBox "User [ " & m_strUserID & _
                             " ] has been updated!", _
                             vbInformation Or vbOKOnly, "Invalid User ID"
                      Reset_frmChgPass
                      Exit Sub
                  End If
              Else
                  MsgBox "New Password / Passphrase entries are not the same.", _
                         vbInformation Or vbOKOnly, "Invalid data entered"
                  Clear_Bottom_Boxes
                  txtPassword(2).SetFocus
                  Exit Sub
              End If
              
         ' Cancel button was pressed.
         Case 2:
              Reset_frmChgPass
              frmChgPass.Hide
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
  Erase arPWord3()
  
  m_strUserID = String$(250, 0)
  m_strPWord1 = String$(250, 0)
  m_strPWord2 = String$(250, 0)
  m_strPWord3 = String$(250, 0)
  
  m_strUserID = ""
  m_strPWord1 = ""
  m_strPWord2 = ""
  m_strPWord3 = ""
  
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
  CenterCaption frmChgPass

' ---------------------------------------------------------------------------
' Hide this form
' ---------------------------------------------------------------------------
  frmChgPass.Hide

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
' ---------------------------------------------------------------------------
' clear variables
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
         Case 0: cmdChoice_Click 2  ' return to the main menu
         Case 1: Exit Sub
         Case 2: TerminateApplication
         Case 3: TerminateApplication
         Case 4: TerminateApplication
  End Select
  
End Sub

Public Sub Reset_frmChgPass()

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intIndex As Integer

' ---------------------------------------------------------------------------
' Empty all the text boxes
' ---------------------------------------------------------------------------
  With frmChgPass
       ' Empty the input text boxes
       For intIndex = 0 To 3
           .txtPassword(intIndex).Enabled = True
           .txtPassword(intIndex).Text = ""
       Next
  End With
  
' ---------------------------------------------------------------------------
' Reset the OK buttons and disable bottom two text boxes
' ---------------------------------------------------------------------------
  Button_Lock True
  
' ---------------------------------------------------------------------------
' Display the form
' ---------------------------------------------------------------------------
  With frmChgPass
       .lblMyLabel = MYNAME
       .Show vbModeless   ' reduces flicker
       .Refresh
  End With
  
' ---------------------------------------------------------------------------
' set the cursor in first text box
' ---------------------------------------------------------------------------
  txtPassword(0).SetFocus
  
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
  Dim CtrlDown As Integer
  Dim PressedKey As Integer
  
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
      Edit_Cut    ' Ctrl + X was pressed
  ElseIf CtrlDown And PressedKey = vbKeyA Then
      SendKeys "{Home}{End}"
  ElseIf CtrlDown And PressedKey = vbKeyC Then
      Edit_Copy   ' Ctrl + C was pressed
  ElseIf CtrlDown And PressedKey = vbKeyV Then
      Edit_Paste  ' Ctrl + V was pressed
  ElseIf PressedKey = vbKeyDelete Then
      Edit_Delete ' Delete key was pressed
  End If

End Sub

Private Sub txtPassword_KeyPress(Index As Integer, KeyAscii As Integer)

' ---------------------------------------------------------------------------
' If ENTER is pressed then nullify keystroke and execute the OK button
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
  
         ' backspace and other printable keyboard characters
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
' Save the input data
' ---------------------------------------------------------------------------
  Select Case Index
  
         ' User ID
         Case 0:
              ' Something may have changed in the user id box
              ' therefore, lock out the bottom two text boxes
              Button_Lock True
              
              ' See if anything has been entered in the user id box
              If Len(txtPassword(0).Text) > 0 Then
                  If Not g_blnCaseSensitiveUserID Then
                      txtPassword(0).Text = StrConv(txtPassword(0).Text, vbUpperCase)
                  End If
                  
                  m_strUserID = txtPassword(0).Text
              Else
                  m_strUserID = ""
              End If
              
         ' Passwords
         Case 1:  'Old password text box
              If Len(txtPassword(1).Text) > 0 Then
                  If Not g_blnCaseSensitiveUserID Then
                      txtPassword(1).Text = StrConv(txtPassword(1).Text, vbUpperCase)
                  End If
    
                  m_strPWord1 = txtPassword(1).Text
                  txtPassword(1).Text = String$(30, "*")
              Else
                  ' empty the first password box
                  txtPassword(1).Text = ""
                  Erase arPWord1()
                  m_strPWord1 = ""
              End If
            
         Case 2:  'New password text box
              If Len(txtPassword(2).Text) > 0 Then
                  If Not g_blnCaseSensitiveUserID Then
                      txtPassword(2).Text = StrConv(txtPassword(2).Text, vbUpperCase)
                  End If
    
                  m_strPWord2 = txtPassword(2).Text
                  txtPassword(2).Text = String$(30, "*")
              Else
                  ' empty the bottom two text boxes
                  Clear_Bottom_Boxes
              End If
            
         Case 3:  'Verify new password text box
              If Len(txtPassword(3).Text) > 0 Then
                  If Not g_blnCaseSensitiveUserID Then
                      txtPassword(3).Text = StrConv(txtPassword(3).Text, vbUpperCase)
                  End If
    
                  m_strPWord3 = txtPassword(3).Text
                  txtPassword(3).Text = String$(30, "*")
              Else
                  ' empty the bottom text box
                  Erase arPWord3()
                  txtPassword(3).Text = ""
                  m_strPWord3 = ""
              End If
  End Select
  
End Sub

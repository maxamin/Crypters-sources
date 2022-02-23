VERSION 5.00
Begin VB.Form frmDelUser 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2190
   ClientLeft      =   1560
   ClientTop       =   1845
   ClientWidth     =   5580
   Icon            =   "frmDelUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5580
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
      TabIndex        =   1
      Top             =   1020
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
      Left            =   4440
      TabIndex        =   2
      Top             =   1470
      Width           =   975
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   840
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1275
      Width           =   3435
   End
   Begin VB.Label lblMyLabel 
      BackStyle       =   0  'Transparent
      Height          =   420
      Left            =   240
      TabIndex        =   5
      Top             =   1740
      Width           =   2385
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delete User"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5310
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   1320
      Width           =   540
   End
End
Attribute VB_Name = "frmDelUser"
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
  Private m_strUserID  As String

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
              ' Test data input
              If Len(m_strUserID) = 0 Then
                  MsgBox "A user ID must be entered.", _
                         vbInformation Or vbOKOnly, "User ID missing"
                  txtUserID.SetFocus
                  Exit Sub
              Else
                  arUserID = ConvertToArray(m_strUserID)
              End If
                  
              ' Is this user on file?
              If Not Query_User(arUserID(), strTmp1, strTmp2) Then
                  MsgBox "User [ " & m_strUserID & _
                         " ] cannot be found in the database.", _
                         vbInformation Or vbOKOnly, "Invalid User ID"
                  
                  ' get rid of any data returned
                  strTmp1 = String$(250, 0)
                  strTmp2 = String$(250, 0)
                  Reset_frmDelUser
                  Exit Sub
              Else
                  ' get rid of any data returned
                  strTmp1 = String$(250, 0)
                  strTmp2 = String$(250, 0)
              End If
              
              ' Remove the user from the database
              If Not Remove_User(arUserID()) Then
                  txtUserID.SetFocus
                  Exit Sub
              End If
              
              ' We were successful
              MsgBox "User [ " & m_strUserID & _
                     " ] has been removed from the database.", _
                     vbInformation Or vbOKOnly, "Update Successful"
              Reset_frmDelUser
              
         ' Cancel button was pressed.
         Case 1:
              Reset_frmDelUser
              frmDelUser.Hide
              frmMainMenu.Show
  End Select
  
End Sub

Private Sub ClearVariables()

' ---------------------------------------------------------------------------
' clear variables
' ---------------------------------------------------------------------------
  Erase arUserID()
  m_strUserID = String$(250, 0)
  m_strUserID = ""
  
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
  CenterCaption frmDelUser

' ---------------------------------------------------------------------------
' Hide this form
' ---------------------------------------------------------------------------
  frmDelUser.Hide

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

Public Sub Reset_frmDelUser()

' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  ClearVariables
  
' ---------------------------------------------------------------------------
' Display the form
' ---------------------------------------------------------------------------
  With frmDelUser
       .txtUserID.Text = ""
       .lblMyLabel = MYNAME
       .Show vbModeless
       .Refresh
  End With
  
' ---------------------------------------------------------------------------
' place cursor in first text box
' ---------------------------------------------------------------------------
  txtUserID.SetFocus
  
End Sub

Private Sub txtUserID_GotFocus()
  
' ---------------------------------------------------------------------------
' Highlight all the text in the box
' ---------------------------------------------------------------------------
  SendKeys "{Home}{End}"
  
End Sub

Private Sub txtUserID_KeyDown(KeyCode As Integer, Shift As Integer)

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

Private Sub txtUserID_KeyPress(KeyAscii As Integer)

' ---------------------------------------------------------------------------
' If ENTER is pressed then nullify keystroke and press the OK button
' ---------------------------------------------------------------------------
  If KeyAscii = 13 Then
      KeyAscii = 0
      txtUserID_Validate False
      cmdChoice_Click 0
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

Private Sub txtUserID_Validate(Cancel As Boolean)

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
  txtUserID.Text = Trim$(txtUserID.Text)
  
' ---------------------------------------------------------------------------
' Initial test of the input data
' ---------------------------------------------------------------------------
  If Len(txtUserID.Text) > 0 Then
      If Not g_blnCaseSensitiveUserID Then
          txtUserID.Text = StrConv(txtUserID.Text, vbUpperCase)
      End If
                  
      m_strUserID = txtUserID.Text
  Else
      m_strUserID = ""
  End If
              
End Sub

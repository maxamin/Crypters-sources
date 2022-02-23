VERSION 5.00
Begin VB.Form frmPWord 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2190
   ClientLeft      =   2325
   ClientTop       =   2100
   ClientWidth     =   5490
   ControlBox      =   0   'False
   Icon            =   "frmChgPWD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5490
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1500
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1500
      Width           =   2600
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
      Left            =   4320
      TabIndex        =   2
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
      Left            =   4320
      TabIndex        =   3
      Top             =   1470
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1500
      TabIndex        =   0
      Top             =   1035
      Width           =   2600
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Index           =   1
      Left            =   615
      TabIndex        =   7
      Top             =   1605
      Width           =   690
   End
   Begin VB.Label lblMyLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   1200
      TabIndex        =   6
      Top             =   1980
      Width           =   4065
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   180
      TabIndex        =   5
      Top             =   60
      Width           =   5190
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      Height          =   195
      Index           =   0
      Left            =   765
      TabIndex        =   4
      Top             =   1140
      Width           =   540
   End
End
Attribute VB_Name = "frmPWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------
' Define module level variables
' ---------------------------------------------------------------------------
  Private m_strPassword  As String
  
Public Sub Reset_frmPWord()

' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  m_strPassword = ""
  
' ---------------------------------------------------------------------------
' Display the form
' ---------------------------------------------------------------------------
  With frmPWord
       .txtPassword(0).Text = ""
       .txtPassword(1).Text = ""
       .Show vbModeless
       .Refresh
  End With
  
End Sub

Private Sub cmdChoice_Click(Index As Integer)

' ---------------------------------------------------------------------------
' Based on the button pressed
' ---------------------------------------------------------------------------
  Select Case Index
         
         ' OK button was pressed.
         ' Time to test the input.
         Case 0:
              ' Test data input
              If Len(g_strUserID) = 0 Then
                  MsgBox "A valid user ID must be entered.", _
                         vbInformation + vbOKOnly, "User ID missing"
                  txtPassword(0).SetFocus
                  Exit Sub
              Else
                  ' build SQL statement
                  g_SQLstmt = "SELECT * FROM [PWord] Where [UserID] = '" & g_strUserID & "'"
                  
                  ' Is this user on file?
                  If Not Query_User() Then
                      Clear_Variables
                      m_strPassword = ""
                      txtPassword(1) = ""
                      txtPassword(0).SetFocus
                      Exit Sub
                  End If
              End If
              
              If Not Validate_Password_Entry(m_strPassword) Then
                  m_strPassword = ""
                  txtPassword(1) = ""
                  txtPassword(1).SetFocus
                  Exit Sub
              End If
             
              ' Build the new password and
              ' update the database
              If Build_Password(m_strPassword) Then
                  Clear_Variables
                  frmChgPass.Hide
                  frmMainMenu.Show
              Else
                  m_strPassword = ""
                  txtPassword(1) = ""
                  txtPassword(1).SetFocus
                  Exit Sub
              End If
             
         ' Cancel button was pressed.
         Case 1:
              Clear_Variables
              frmPWord.Hide
              frmMainMenu.Show
  End Select
  
End Sub
Private Function Good_Old_Password() As Boolean

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strInText      As String
  Dim strHashed      As String
  Dim intLength      As Integer
  Dim cMD5           As clsMD5
  
' ---------------------------------------------------------------------------
' Initialize local variables
' ---------------------------------------------------------------------------
  strInText = ""
  intLength = Len(m_strPassword)
  Set cMD5 = New clsMD5

' ---------------------------------------------------------------------------
' Concatenate the Password / Passphrase entered with the salt value from the
' database and hash the results.
' ---------------------------------------------------------------------------
  strInText = strInText & g_strSalt
  strHashed = cMD5.MD5_Hash_Data_String(strInText)
  Set cMD5 = Nothing
  
' ---------------------------------------------------------------------------
' compare just the hashed results with the ruslts stored in the database
' ---------------------------------------------------------------------------
  If StrComp(strInText, g_strHash, vbTextCompare) = 0 Then
      ' if a valid password was entered in the password
      ' textbox and it matched what was on file
      ' display an error message
      MsgBox "We have a password match.", _
             vbOKOnly + vbInformation, "Success"
      Good_Old_Password = True
  Else
      ' display an error message
      MsgBox "Password does not match what is on file.", _
             vbOKOnly + vbInformation, "Invalid Data"
      Good_Old_Password = False
  End If


Normal_Exit:
' ---------------------------------------------------------------------------
' clear variables holding the digital signatures for security reasons.
' Leave nothing in memory.
' ---------------------------------------------------------------------------
  strInText = String(intLength, 0)
  strInText = ""
  
End Function

Private Sub Form_Initialize()

' ---------------------------------------------------------------------------
' Centered on the screen
' ---------------------------------------------------------------------------
  Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

End Sub
Private Sub Form_Load()

' ---------------------------------------------------------------------------
' Make sure we do not show this form when first loading the application
' ---------------------------------------------------------------------------
  With frmPWord
       .txtPassword(0).Text = ""
       .txtPassword(1).Text = ""
       .lblMyLabel = "Freeware by Kenneth Ives  kenaso@home.com"
       .lblTitle = "XYZ Corporation" & vbCrLf & "Password Entry"
       .Hide
  End With
  
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' ---------------------------------------------------------------------------
' Empty variables
' ---------------------------------------------------------------------------
  Clear_Variables
  
' ---------------------------------------------------------------------------
' Based on the the unload code the system passes,
' we determine what to do
'
' Unloadmode codes
'     0 - Close from the control-menu box
'         or Upper left "X"
'     1 - Unload method from code elsewhere
'         in the application
'     2 - Windows Session is ending
'     3 - Task Manager is clostrIng the app
'     4 - MDI Parent is clostrIng
' ---------------------------------------------------------------------------
  Select Case UnloadMode
         Case 0: cmdChoice_Click 1  ' Return to the main menu
         Case 1: Exit Sub
         Case 2: StopApplication
         Case 3: StopApplication
         Case 4: StopApplication
  End Select
  
End Sub

Private Sub txtPassword_GotFocus(Index As Integer)
  
' ---------------------------------------------------------------------------
' Highlight all the text in the box
' ---------------------------------------------------------------------------
  With txtPassword(Index)
       .SelStart = 0
       .SelLength = Len(.Text)
  End With
  
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
  PressedKey = Asc(UCase(Chr(KeyCode)))   ' Convert to uppercase
    
' ---------------------------------------------------------------------------
' Check to see if it is okay to make changes
' ---------------------------------------------------------------------------
  If CtrlDown And PressedKey = vbKeyX Then
      Edit_Cut            ' Ctrl + X was pressed
  ElseIf CtrlDown And PressedKey = vbKeyA Then
      With txtPassword(Index)    ' Ctrl + A was pressed
           .SelStart = 0
           .SelLength = Len(.Text)
       End With
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
' If ENTER is pressed then nullify keystroke and execute the OK button
' ---------------------------------------------------------------------------
  If KeyAscii = 13 Then
      KeyAscii = 0
      cmdChoice(0).SetFocus
      SendKeys "{ENTER}"
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

Private Sub txtPassword_LostFocus(Index As Integer)

' ---------------------------------------------------------------------------
' Remove leading and trailing blank spaces
' ---------------------------------------------------------------------------
  txtPassword(Index) = Trim(txtPassword(Index))
  
' ---------------------------------------------------------------------------
' Initial test of the input data
' ---------------------------------------------------------------------------
  Select Case Index
  
         ' User ID
         Case 0:
              If Len(txtPassword(0)) > 0 Then
                  txtPassword(0) = StrConv(txtPassword(0), vbUpperCase)
                  g_strUserID = txtPassword(0)
              Else
                  g_strUserID = ""
              End If
              
              ' change has occured in the userID text box
              g_strHash = ""
              g_strSalt = ""
              txtPassword(1) = ""
              m_strPassword = ""
              
         ' Password
         Case Else:
              ' if something is in the password box
              ' then save to another variable and
              ' fill the text box with asteriks
              If Len(txtPassword(1)) > 0 Then
                  m_strPassword = txtPassword(1)
                  txtPassword(1) = String(30, "*")
              Else
                  ' else empty the text box and the
                  ' variable
                  txtPassword(1) = ""
                  m_strPassword = ""
              End If
  End Select
  
End Sub

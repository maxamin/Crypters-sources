VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4725
   ClientLeft      =   1560
   ClientTop       =   1845
   ClientWidth     =   4440
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   4440
   Begin VB.Frame Frame1 
      Caption         =   "Provider selection"
      Height          =   795
      Index           =   2
      Left            =   150
      TabIndex        =   12
      Top             =   1875
      Width           =   2895
      Begin VB.OptionButton optProvider 
         Caption         =   "Enhanced provider"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   14
         Top             =   450
         Width           =   1815
      End
      Begin VB.OptionButton optProvider 
         Caption         =   "Default provider"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   225
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Password Hashing methods"
      Height          =   1635
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   2895
      Begin VB.OptionButton optHash 
         Caption         =   "SHA-1 (160-bit / 20 byte output)"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   11
         Top             =   1200
         Width           =   2600
      End
      Begin VB.OptionButton optHash 
         Caption         =   "MD5 (128-bit / 16 byte output)"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   10
         Top             =   900
         Width           =   2600
      End
      Begin VB.OptionButton optHash 
         Caption         =   "MD4 (128-bit / 16 byte output)"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   600
         Width           =   2600
      End
      Begin VB.OptionButton optHash 
         Caption         =   "MD2 (128-bit / 16 byte output)"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Value           =   -1  'True
         Width           =   2600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Logon options"
      Height          =   795
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4170
      Begin VB.CheckBox chkLogon 
         Caption         =   "Case sensitive Password / Passphrase"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   480
         Value           =   1  'Checked
         Width           =   3270
      End
      Begin VB.CheckBox chkLogon 
         Caption         =   "Case sensitive User ID"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   3270
      End
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
      Left            =   3300
      TabIndex        =   0
      Top             =   3450
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
      Left            =   3300
      TabIndex        =   1
      Top             =   3975
      Width           =   975
   End
   Begin VB.Label lblMyLabel 
      BackStyle       =   0  'Transparent
      Height          =   270
      Left            =   150
      TabIndex        =   3
      Top             =   4500
      Width           =   3960
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Options"
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
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4185
   End
End
Attribute VB_Name = "frmOptions"
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
  Private m_strPassword      As String
  Private m_blnPrevUserID    As Boolean
  Private m_blnPrevPWord     As Boolean
  Private m_blnEnhancedProv  As Boolean
  Private m_intUserID        As Integer
  Private m_intPWord         As Integer
  Private m_intHashType      As Integer
  Private m_intPrevHashType  As Integer

Private Sub SaveAllSettings()

' ---------------------------------------------------------------------------
' Save hash selection
' ---------------------------------------------------------------------------
  g_intHashType = m_intHashType

' ---------------------------------------------------------------------------
' Save changes prior to leaving this form
' ---------------------------------------------------------------------------
  CurrentSettings_Save "UserID", g_blnCaseSensitiveUserID
  CurrentSettings_Save "Password", g_blnCaseSensitivePWord
  CurrentSettings_Save "EnhancedProvider", g_blnEnhancedProvider
  CurrentSettings_Save "HashMethod", g_intHashType
  
End Sub

Private Sub chkLogon_Click(Index As Integer)

  Select Case Index
         Case 0:  ' case sensitive user ID
              If chkLogon(0).Value = vbChecked Then
                  g_blnCaseSensitiveUserID = True
              Else
                  g_blnCaseSensitiveUserID = False
              End If
              
         Case 1:  ' case sensitive password / passphrase
              If chkLogon(1).Value = vbChecked Then
                  g_blnCaseSensitivePWord = True
              Else
                  g_blnCaseSensitivePWord = False
              End If
  End Select
           
End Sub

Private Sub cmdChoice_Click(Index As Integer)

' ---------------------------------------------------------------------------
' Based on the button pressed
' ---------------------------------------------------------------------------
  Select Case Index
         
         ' OK button was pressed.
         Case 0:
              SaveAllSettings
              
         ' Cancel button was pressed.
         Case 1:
              ' restore previous settings
              g_blnCaseSensitiveUserID = m_blnPrevUserID
              g_blnCaseSensitivePWord = m_blnPrevPWord
              g_intHashType = m_intPrevHashType
  End Select
  
' ---------------------------------------------------------------------------
' Hide this form and return to the main menu
' ---------------------------------------------------------------------------
  frmOptions.Hide
  frmMainMenu.Show

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
  CenterCaption frmOptions

' ---------------------------------------------------------------------------
' Hide this form
' ---------------------------------------------------------------------------
  frmOptions.Hide

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
         Case Else: TerminateApplication
  End Select
  
End Sub

Private Sub optHash_Click(Index As Integer)

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intIndex As Integer
  
' ---------------------------------------------------------------------------
' Set the hash option display
' ---------------------------------------------------------------------------
  For intIndex = 0 To 3
      If intIndex = Index Then
          optHash(intIndex).Value = True
      Else
          optHash(intIndex).Value = False
      End If
  Next
  
' ---------------------------------------------------------------------------
' Based on option selected, determine if we are to use SHA-1 or MD5
' ---------------------------------------------------------------------------
  Select Case Index
         Case 0  ' Use MD2
              m_intHashType = 3
              
         Case 1  ' Use MD4
              m_intHashType = 2
              
         Case 2  ' Use MD5
              m_intHashType = 1
              
         Case 3  ' Use SHA-1
              m_intHashType = 4
  End Select
  
End Sub

Public Sub Reset_frmOptions()

' ---------------------------------------------------------------------------
' Save previous settings
' ---------------------------------------------------------------------------
  m_blnPrevUserID = g_blnCaseSensitiveUserID
  m_blnPrevPWord = g_blnCaseSensitivePWord
  m_blnEnhancedProv = g_blnEnhancedProvider
  m_intPrevHashType = g_intHashType
  
' ---------------------------------------------------------------------------
' By adding converted values together, we get a positive number.
' i.e.  TRUE = -1   0 = 0 - -1
' ---------------------------------------------------------------------------
  If g_blnCaseSensitiveUserID Then
      m_intUserID = 1  ' TRUE
  Else
      m_intUserID = 0  ' FALSE
  End If
  
  If g_blnCaseSensitivePWord Then
      m_intPWord = 1   ' TRUE
  Else
      m_intPWord = 0   ' FALSE
  End If
  
' ---------------------------------------------------------------------------
' Set the index for the hash option display
' ---------------------------------------------------------------------------
  Select Case g_intHashType
         Case 1: optHash_Click 2  ' MD5
         Case 2: optHash_Click 0  ' MD2
         Case 3: optHash_Click 1  ' MD4
         Case 4: optHash_Click 3  ' SHA-1
  End Select
  
' ---------------------------------------------------------------------------
' Display the form
' ---------------------------------------------------------------------------
  With frmOptions
       .chkLogon(0).Value = m_intUserID
       .chkLogon(1).Value = m_intPWord
              
       .lblMyLabel = MYNAME
       .Show vbModeless
       .Refresh
  End With
  
  chkLogon_Click 0
  chkLogon_Click 1
  
End Sub

Private Sub optProvider_Click(Index As Integer)

' ---------------------------------------------------------------------------
' Based on option selected, determine if we are to use SHA-1 or MD5
' ---------------------------------------------------------------------------
  Select Case Index
           
         Case 0  ' Use Default provider
              optProvider(0).Value = True
              optProvider(1).Value = False
              g_blnEnhancedProvider = False
  
         Case 1  ' Use Enhanced provider
              optProvider(0).Value = False
              optProvider(1).Value = True
              g_blnEnhancedProvider = True
  
  End Select
  
End Sub

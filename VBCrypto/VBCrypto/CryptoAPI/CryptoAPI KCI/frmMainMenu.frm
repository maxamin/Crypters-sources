VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5115
   ClientLeft      =   1560
   ClientTop       =   1845
   ClientWidth     =   4620
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Miscellaneous"
      Height          =   1455
      Left            =   2340
      TabIndex        =   12
      Top             =   3000
      Width           =   2115
      Begin VB.CommandButton cmdChoice 
         Caption         =   "Test &Random Data"
         Height          =   375
         Index           =   8
         Left            =   180
         TabIndex        =   15
         Top             =   900
         Width           =   1695
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "&Hash Testing"
         Height          =   375
         Index           =   7
         Left            =   180
         TabIndex        =   14
         Top             =   420
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Encryption"
      Height          =   1455
      Left            =   180
      TabIndex        =   9
      Top             =   3000
      Width           =   2055
      Begin VB.CommandButton cmdChoice 
         Caption         =   "Test &Encyption"
         Height          =   375
         Index           =   6
         Left            =   180
         TabIndex        =   16
         Top             =   900
         Width           =   1695
      End
      Begin VB.OptionButton optEncrypt 
         Caption         =   "Files"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   540
         Width           =   1155
      End
      Begin VB.OptionButton optEncrypt 
         Caption         =   "Data strings"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   10
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Password / Passphrase"
      Height          =   1875
      Left            =   180
      TabIndex        =   3
      Top             =   960
      Width           =   4275
      Begin VB.CommandButton cmdChoice 
         Caption         =   "&Options"
         Height          =   375
         Index           =   5
         Left            =   2385
         TabIndex        =   13
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "&View Database"
         Height          =   375
         Index           =   4
         Left            =   2385
         TabIndex        =   8
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "&Test Password Entry"
         Height          =   375
         Index           =   3
         Left            =   2385
         TabIndex        =   7
         Top             =   375
         Width           =   1695
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "&Add User"
         Height          =   375
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "&Change Password"
         Height          =   375
         Index           =   1
         Left            =   165
         TabIndex        =   5
         Top             =   855
         Width           =   1695
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "&Delete User"
         Height          =   375
         Index           =   2
         Left            =   165
         TabIndex        =   4
         Top             =   1335
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "E&xit"
      Height          =   375
      Index           =   9
      Left            =   3180
      TabIndex        =   0
      Top             =   4620
      Width           =   1275
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CryptoAPI  Demo"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4350
   End
   Begin VB.Label lblMyLabel 
      BackStyle       =   0  'Transparent
      Height          =   405
      Left            =   180
      TabIndex        =   1
      Top             =   4620
      Width           =   2595
   End
End
Attribute VB_Name = "frmMainMenu"
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
' Module level variables
' ---------------------------------------------------------------------------
  Private m_blnEncryptFiles As Boolean

Private Sub cmdChoice_Click(Index As Integer)

' ---------------------------------------------------------------------------
' Hide the main menu
' ---------------------------------------------------------------------------
  frmMainMenu.Hide

' ---------------------------------------------------------------------------
' Based on the button pressed
' ---------------------------------------------------------------------------
  Select Case Index
         Case 0:   ' Add a new user
              frmAddUser.Reset_frmAddUser
              
         Case 1:   ' Change a password
              frmChgPass.Reset_frmChgPass
              
         Case 2:   ' Delete a user
              frmDelUser.Reset_frmDelUser
              
         Case 3:   ' Test password entry
              frmTestPWD.Reset_frmTestPWD
              
         Case 4:   ' View the database
              frmDB.Reset_frmDB
              
         Case 5:   ' Options
              frmOptions.Reset_frmOptions
              
         Case 6:   ' Test Encryption
              If m_blnEncryptFiles Then
                  frmEncFiles.Reset_frmEncfiles
              Else
                  frmEncStrings.Reset_frmEncStrings
              End If
              
         Case 7:   ' Hash testing
              frmHash.Reset_frmHash
              
         Case 8:   ' Test Random Data
              frmRnd.Reset_frmRnd
              
         Case 9:   ' Terminate this program
              TerminateApplication
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
  CenterCaption frmMainMenu

' ---------------------------------------------------------------------------
' Set the default option
' ---------------------------------------------------------------------------
  optEncrypt_Click 0
  
' ---------------------------------------------------------------------------
' Display this form
' ---------------------------------------------------------------------------
  With frmMainMenu
       .lblMyLabel = MYNAME
       .Show vbModeless  ' reduce flicker
  End With
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
' ---------------------------------------------------------------------------
' Based on the the unload code the system passes, we determine what to do
'
' Unloadmode codes
'     0 - Close from the control-menu box
'         or Upper right "X"
'     1 - Unload method from code elsewhere
'         in the application
'     2 - Windows Session is ending
'     3 - Task Manager is closing the app
'     4 - MDI Parent is closing
' ---------------------------------------------------------------------------
  Select Case UnloadMode
         Case 0: TerminateApplication
         Case 1: Exit Sub
         Case 2: TerminateApplication
         Case 3: TerminateApplication
         Case 4: TerminateApplication
  End Select
  
End Sub

Private Sub optEncrypt_Click(Index As Integer)
  
  Select Case Index
         
         Case 0   ' Work with data strings
              optEncrypt(0).Value = True   ' Data strings
              optEncrypt(1).Value = False  ' Files
              m_blnEncryptFiles = False
         
         Case 1   ' Work with files
              optEncrypt(0).Value = False  ' Data strings
              optEncrypt(1).Value = True   ' Files
              m_blnEncryptFiles = True
  End Select
  
End Sub

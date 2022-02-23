VERSION 5.00
Begin VB.Form frmRnd 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5055
   ClientLeft      =   1560
   ClientTop       =   1845
   ClientWidth     =   5565
   Icon            =   "frmRnd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5565
   Begin VB.TextBox txtRnd 
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
      Top             =   3660
      Width           =   5295
   End
   Begin VB.TextBox txtRnd 
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
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2520
      Width           =   5295
   End
   Begin VB.TextBox txtRnd 
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
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1200
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
      Left            =   3360
      TabIndex        =   0
      Top             =   4620
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
      TabIndex        =   1
      Top             =   4620
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   5130
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Test 3.  Random generated data  (return exact length requested)"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   3420
      Width           =   5130
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Test 2.  Random generated data converted to hex  (data length doubles when converted)"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   2100
      Width           =   3690
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Test 1.  Random generated data"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   3600
   End
   Begin VB.Label lblMyLabel 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   180
      TabIndex        =   3
      Top             =   4560
      Width           =   2925
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Test Random Data Output"
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
      TabIndex        =   2
      Top             =   120
      Width           =   5310
   End
End
Attribute VB_Name = "frmRnd"
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
  Private Const m_intAmount As Integer = 50

Private Sub cmdChoice_Click(Index As Integer)

' ---------------------------------------------------------------------------
' Based on the button pressed
' ---------------------------------------------------------------------------
  Select Case Index
         
         ' OK button was pressed.
         ' Time to test the input.
         Case 0:
              Dim cCrypto As CryptKci.clsCryptoAPI
              Set cCrypto = New CryptKci.clsCryptoAPI
              
              txtRnd(0).Text = ""
              txtRnd(1).Text = ""
              txtRnd(2).Text = ""
              
              txtRnd(0).Text = cCrypto.CreateRandom(m_intAmount, False, False)
              txtRnd(1).Text = cCrypto.CreateRandom(m_intAmount, False, True)
              txtRnd(2).Text = cCrypto.CreateRandom(m_intAmount, True, True)
              Set cCrypto = Nothing
              
         ' Cancel button was pressed.
         Case 1:
              frmRnd.Hide
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
  CenterCaption frmRnd

' ---------------------------------------------------------------------------
' Hide this form
' ---------------------------------------------------------------------------
  frmRnd.Hide

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

Public Sub Reset_frmRnd()

' ---------------------------------------------------------------------------
' Display the form
' ---------------------------------------------------------------------------
  With frmRnd
       .txtRnd(0).Text = ""
       .txtRnd(1).Text = ""
       .txtRnd(2).Text = ""
       .Label1(3).Caption = "Three tests are based on " & CStr(m_intAmount) & _
                            " byte data length"
       .lblMyLabel = MYNAME
       .Show vbModeless
  End With
  
End Sub

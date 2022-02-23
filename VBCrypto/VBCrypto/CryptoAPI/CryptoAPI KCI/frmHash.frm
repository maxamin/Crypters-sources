VERSION 5.00
Begin VB.Form frmHash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5250
   ClientLeft      =   1560
   ClientTop       =   1845
   ClientWidth     =   5565
   Icon            =   "frmHash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   5565
   Begin VB.TextBox txtHash 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4200
      Width           =   5295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hash Functions"
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   2235
      Begin VB.OptionButton optHash 
         Caption         =   "SHA-1"
         Height          =   195
         Index           =   3
         Left            =   1260
         TabIndex        =   16
         Top             =   540
         Width           =   855
      End
      Begin VB.OptionButton optHash 
         Caption         =   "MD5"
         Height          =   195
         Index           =   2
         Left            =   1260
         TabIndex        =   15
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton optHash 
         Caption         =   "MD4"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   14
         Top             =   540
         Width           =   855
      End
      Begin VB.OptionButton optHash 
         Caption         =   "MD2"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   13
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox txtHash 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3600
      Width           =   5295
   End
   Begin VB.TextBox txtHash 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3000
      Width           =   5295
   End
   Begin VB.TextBox txtHash 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2100
      Width           =   5295
   End
   Begin VB.ComboBox cboTestData 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Text            =   "cboTestData"
      Top             =   1140
      Width           =   2655
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
      TabIndex        =   1
      Top             =   4680
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
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblSID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   3960
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actual output"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Predicted output"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Input test data"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1860
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select your test data"
      Height          =   195
      Index           =   0
      Left            =   3180
      TabIndex        =   5
      Top             =   900
      Width           =   1455
   End
   Begin VB.Label lblMyLabel 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   180
      TabIndex        =   4
      Top             =   4680
      Width           =   2925
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Test Hash Algorithms"
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
      TabIndex        =   3
      Top             =   120
      Width           =   5310
   End
End
Attribute VB_Name = "frmHash"
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
  Private m_intHashType  As Integer
  Private m_strHashType  As String

Private Sub cboTestData_Click()

' ---------------------------------------------------------------------------
' Based on item selected, update the display boxes
' ---------------------------------------------------------------------------
  Select Case m_intHashType
        Case 1  ' Use MD5
              ' use the combo box index to determine which test to perform
              Select Case cboTestData.ListIndex
                     Case 0: ' single letter "a"
                          txtHash(0).Text = "a"
                          txtHash(1).Text = "0CC175B9C0F1B6A831C399E269772661"
                    
                     Case 1: ' letters "abc"
                          txtHash(0).Text = "abc"
                          txtHash(1).Text = "900150983CD24FB0D6963F7D28E17F72"
                    
                     Case 2: ' Empty string
                          txtHash(0).Text = ""
                          txtHash(1).Text = "D41D8CD98F00B204E9800998ECF8427E"
                         
                     Case 3: ' 2 words "message digest"
                          txtHash(0).Text = "message digest"
                          txtHash(1).Text = "F96B697D7CB7938D525A2F31AAF161D0"
                    
                     Case 4: ' Multiple letters
                          txtHash(0).Text = "abcdefghijklmnopqrstuvwxyz"
                          txtHash(1).Text = "C3FCD3D76192E4007DFB496CCA67E13B"
                    
                     Case 5: ' Letters and numbers
                          txtHash(0).Text = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
                          txtHash(1).Text = "D174AB98D277D9F5A5611C2C9F419D9F"
                    
                     Case 6: ' Multiple Numbers
                          txtHash(0).Text = "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
                          txtHash(1).Text = "57EDF4A22BE3C955AC49DA2E2107B67A"
              End Select
  
        Case 2  ' Use MD4
              ' use the combo box index to determine which test to perform
              Select Case cboTestData.ListIndex
                     Case 0: ' single letter "a"
                          txtHash(0).Text = "a"
                          txtHash(1).Text = "BDE52CB31DE33E46245E05FBDBD6FB24"
                     
                     Case 1: ' letters "abc"
                          txtHash(0).Text = "abc"
                          txtHash(1).Text = "A448017AAF21D8525FC10AE87AA6729D"
                     
                     Case 2: ' Empty string
                          txtHash(0).Text = ""
                          txtHash(1).Text = "31D6CFE0D16AE931B73C59D7E0C089C0"
                          
                     Case 3: ' 2 words "message digest"
                          txtHash(0).Text = "message digest"
                          txtHash(1).Text = "D9130A8164549FE818874806E1C7014B"
                     
                     Case 4: ' Multiple letters
                          txtHash(0).Text = "abcdefghijklmnopqrstuvwxyz"
                          txtHash(1).Text = "D79E1C308AA5BBCDEEA8ED63DF412DA9"
                     
                     Case 5: ' Letters and numbers
                          txtHash(0).Text = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
                          txtHash(1).Text = "043F8582F241DB351CE627E153E7F0E4"
                     
                     Case 6: ' Multiple Numbers
                          txtHash(0).Text = "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
                          txtHash(1).Text = "E33B4DDC9C38F2199C3E7B164FCC0536"
              End Select
        
        Case 3  ' Use MD2
              ' use the combo box index to determine which test to perform
              Select Case cboTestData.ListIndex
                     Case 0: ' single letter "a"
                          txtHash(0).Text = "a"
                          txtHash(1).Text = "32EC01EC4A6DAC72C0AB96FB34C0B5D1"
                     
                     Case 1: ' letters "abc"
                          txtHash(0).Text = "abc"
                          txtHash(1).Text = "DA853B0D3F88D99B30283A69E6DED6BB"
                     
                     Case 2: ' Empty string
                          txtHash(0).Text = ""
                          txtHash(1).Text = "8350E5A3E24C153DF2275C9F80692773"
                          
                     Case 3: ' 2 words "message digest"
                          txtHash(0).Text = "message digest"
                          txtHash(1).Text = "AB4F496BFB2A530B219FF33031FE06B0"
                     
                     Case 4: ' Multiple letters
                          txtHash(0).Text = "abcdefghijklmnopqrstuvwxyz"
                          txtHash(1).Text = "4E8DDFF3650292AB5A4108C3AA47940B"
                     
                     Case 5: ' Letters and numbers
                          txtHash(0).Text = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
                          txtHash(1).Text = "DA33DEF2A42DF13975352846C30338CD"
                     
                     Case 6: ' Multiple Numbers
                          txtHash(0).Text = "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
                          txtHash(1).Text = "D5976F79D83D3A0DC9806C3C66F3EFD8"
              End Select
        
         Case 4  ' Use SHA-1
              ' use the combo box index to determine which test to perform
              Select Case cboTestData.ListIndex
                     Case 0: ' single letter "a"
                          txtHash(0).Text = "a"
                          txtHash(1).Text = "86F7E437FAA5A7FCE15D1DDCB9EAEAEA377667B8"
                    
                     Case 1: ' letters "abc"
                          txtHash(0).Text = "abc"
                          txtHash(1).Text = "A9993E364706816ABA3E25717850C26C9CD0D89D"
                    
                     Case 2: ' Empty string
                          txtHash(0).Text = ""
                          txtHash(1).Text = "DA39A3EE5E6B4B0D3255BFEF95601890AFD80709"
                         
                     Case 3: ' 2 words "message digest"
                          txtHash(0).Text = "message digest"
                          txtHash(1).Text = "C12252CEDA8BE8994D5FA0290A47231C1D16AAE3"
                    
                     Case 4: ' Multiple letters
                          txtHash(0).Text = "abcdefghijklmnopqrstuvwxyz"
                          txtHash(1).Text = "32D10C7B8CF96570CA04CE37F2A19D84240D3A89"
                    
                     Case 5: ' Letters and numbers
                          txtHash(0).Text = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
                          txtHash(1).Text = "761C457BF73B14D27E9E9265C46F4B4DDA11F940"
                    
                     Case 6: ' Multiple Numbers
                          txtHash(0).Text = "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
                          txtHash(1).Text = "50ABF5706A150990A08B2C5EA40FA0E585554732"
              End Select
         
         Case Else
              Exit Sub
  End Select
  
' ---------------------------------------------------------------------------
' Clear the bottom output display boxes
' ---------------------------------------------------------------------------
  txtHash(2).Text = ""
  txtHash(3).Text = ""
 
' ---------------------------------------------------------------------------
' update the display boxes
' ---------------------------------------------------------------------------
  If m_intHashType = 4 Then ' if SHA-1, hide label and textbox
      lblSID.Caption = ""
      lblSID.Visible = False
      txtHash(3).Visible = False
  Else
      lblSID.Visible = True
      lblSID.Caption = m_strHashType & " - Looks like a SID (Security Identifier) in the registry."
      txtHash(3).Visible = True
  End If
  
  
End Sub

Private Sub cmdChoice_Click(Index As Integer)

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strTmp As String
  
' ---------------------------------------------------------------------------
' Based on the button pressed
' ---------------------------------------------------------------------------
  Select Case Index
         
         ' OK button was pressed.
         ' Time to test the input.
         Case 0:
              Dim cCrypto As CryptKci.clsCryptoAPI
              Set cCrypto = New CryptKci.clsCryptoAPI
              
              strTmp = cCrypto.CreateHash(txtHash(0).Text, m_intHashType)
              txtHash(2).Text = strTmp
                   
              ' Display what a registry SID (Security Idetifier) looks
              ' like using the same character from the combo box.  If
              ' SHA_1 is selected, use MD5 for the SID display.
              If m_intHashType < 4 Then
                  txtHash(3).Text = Format$(strTmp, "@@@@@@@@-@@@@-@@@@-@@@@-@@@@@@@@@@@@")
              End If
              
              Set cCrypto = Nothing
              
         ' Cancel button was pressed.
         Case 1:
              frmHash.Hide
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
  CenterCaption frmHash

' ---------------------------------------------------------------------------
' Hide this form
' ---------------------------------------------------------------------------
  frmHash.Hide

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

Private Sub Load_ComboBox()

' ---------------------------------------------------------------------------
' Load the combo box
' ---------------------------------------------------------------------------
  With cboTestData
       .Clear
       .AddItem "Test 1 - Single letter", 0
       .AddItem "Test 2 - Three letters", 1
       .AddItem "Test 3 - Empty data string", 2
       .AddItem "Test 4 - Two words", 3
       .AddItem "Test 5 - Multiple letters", 4
       .AddItem "Test 6 - Letters and numbers", 5
       .AddItem "Test 7 - Multiple numbers", 6
       .ListIndex = 0
  End With
  
' ---------------------------------------------------------------------------
' Empty the display labels
' ---------------------------------------------------------------------------
  txtHash(0).Text = ""
  txtHash(1).Text = ""
  txtHash(2).Text = ""
  txtHash(3).Text = ""
  
' ---------------------------------------------------------------------------
' Activate the first item in the combo box
' ---------------------------------------------------------------------------
  m_intHashType = g_intHashType
  
  Select Case m_intHashType
         Case 1: optHash_Click 2      ' MD5
         Case 2: optHash_Click 1      ' MD4
         Case 3: optHash_Click 0      ' MD2
         Case 4: optHash_Click 3      ' SHA-1
  End Select

End Sub

Private Sub optHash_Click(Index As Integer)
  
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intIndex As Integer
  
' ---------------------------------------------------------------------------
' Based on option chosen, test the appropriate hash routine
' ---------------------------------------------------------------------------
  Select Case Index
         Case 0:  ' Use MD2
              m_strHashType = "MD2"
              m_intHashType = 3
              
         Case 1:  ' Use MD4
              m_strHashType = "MD4"
              m_intHashType = 2
              
         Case 2:  ' Use MD5
              m_strHashType = "MD5"
              m_intHashType = 1
              
         Case 3:  ' Use SHA-1
              m_strHashType = ""
              m_intHashType = 4
              
         Case Else: Exit Sub
  End Select
  
' ---------------------------------------------------------------------------
' Setup the hash selection display
' ---------------------------------------------------------------------------
  For intIndex = 0 To 3
      If intIndex = Index Then
          optHash(intIndex).Value = True
      Else
          optHash(intIndex).Value = False
      End If
  Next
  
' ---------------------------------------------------------------------------
' Load the first two text boxes with test data
' ---------------------------------------------------------------------------
  cboTestData_Click
  
End Sub

Public Sub Reset_frmHash()

' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  Load_ComboBox
  
' ---------------------------------------------------------------------------
' Display the form
' ---------------------------------------------------------------------------
  With frmHash
       .txtHash(2).Text = ""
       .txtHash(3).Text = ""
       .lblMyLabel = MYNAME
       .lblSID = m_strHashType & " - Looks like a SID (Security " & _
                                 "Identifier) in the registry."
       .Show vbModeless
       .Refresh
  End With
  
End Sub


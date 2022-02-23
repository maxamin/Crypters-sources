VERSION 5.00
Begin VB.Form frmCiphers 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Cipher Algorithms"
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   ControlBox      =   0   'False
   Icon            =   "frmCiphers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSaltD 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      TabIndex        =   6
      Text            =   "12345"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.ComboBox cmbAlgorithms 
      Height          =   315
      ItemData        =   "frmCiphers.frx":1042
      Left            =   240
      List            =   "frmCiphers.frx":1070
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   4695
   End
   Begin VB.TextBox txtSaltE 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      TabIndex        =   3
      Text            =   "12345"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtKeyD 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      TabIndex        =   5
      Text            =   "secretkey"
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmCiphers.frx":10FB
      Top             =   1440
      Width           =   4695
   End
   Begin VB.TextBox txtText2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4680
      Width           =   4695
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3120
      Width           =   4695
   End
   Begin VB.TextBox txtKeyE 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      TabIndex        =   2
      Text            =   "secretkey"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Hash Alg."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   300
      TabIndex        =   24
      Top             =   5445
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Salt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   23
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Algorithm"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Salt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   21
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   120
      Picture         =   "frmCiphers.frx":1124
      Top             =   75
      Width           =   240
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   20
      Top             =   45
      Width           =   135
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Decrypt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3540
      TabIndex        =   19
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Key to Decrypt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Encrypt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3540
      TabIndex        =   17
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Text to Encrypt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Decrypted Text"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Encrypt File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1980
      TabIndex        =   14
      Top             =   5445
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Decrypt File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   5445
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Encrypted String (Base64 Format)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Key to Encrypt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   45
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   45
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cipher Algorithms"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   75
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   5280
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C00000&
      Height          =   375
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   2355
      Width           =   1455
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   4035
      Width           =   1455
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   5745
      Left            =   -1080
      Picture         =   "frmCiphers.frx":2166
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "frmCiphers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
cmbAlgorithms.ListIndex = 0
End Sub

Private Sub Label1_Click()
    txtOutput.Text = EncryptString(cmbAlgorithms.ListIndex, txtText, True, txtKeyE, txtSaltE)
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = 33023
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = -2147483643
End Sub

Private Sub Label11_Click()
MsgBox "This program simplifies the coding procedures of accessing the ebCrypt DLL library. All code was written by David Midkiff (mznull@earthlink.net). ebCrypt.dll is Copyright (c) 2000-2001, EB Design Pty Ltd.", vbInformation + vbOKOnly, "About"
End Sub
Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = 33023
End Sub


Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = -2147483643
End Sub

Private Sub Label13_Click()
    txtText2.Text = DecryptString(cmbAlgorithms.ListIndex, txtOutput, True, txtKeyD, txtSaltD)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label13.ForeColor = 33023
End Sub


Private Sub Label13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.ForeColor = -2147483643
End Sub

Private Sub Label14_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.ForeColor = 33023
End Sub

Private Sub Label14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.ForeColor = -2147483643
End Sub

Private Sub Label17_Click()
frmCiphers.Hide
frmHASH.Show vbModal
frmCiphers.Show
End Sub
Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.ForeColor = 33023
End Sub

Private Sub Label17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.ForeColor = -2147483643
End Sub
Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = 33023
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = -2147483643
End Sub
Private Sub Label7_Click()
    Dim X As Boolean, Key As String, Salt As String, File1 As String, File2 As String
    
    File1 = GetFileInName
    If File1 = "" Then Exit Sub
    
    File2 = GetFileOutName
    If File2 = "" Then Exit Sub
    
    Key = InputBox("Enter key:", "Utilize ebCrypt")
    Salt = InputBox("Enter salt:", "Utilize ebCrypt")
    X = EncryptFile(cmbAlgorithms.ListIndex, File1, File2, True, True, Key, Salt)
End Sub
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = 33023
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = -2147483643
End Sub

Private Sub Label8_Click()
    Dim X As Boolean, Key As String, Salt As String, File1 As String, File2 As String
    
    File1 = GetFileInName
    If File1 = "" Then Exit Sub
    
    File2 = GetFileOutName
    If File2 = "" Then Exit Sub
    
    Key = InputBox("Enter key:", "Utilize ebCrypt")
    Salt = InputBox("Enter salt:", "Utilize ebCrypt")
    X = DecryptFile(cmbAlgorithms.ListIndex, File1, File2, True, True, Key, Salt)
End Sub
Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = 33023

End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = -2147483643
End Sub


VERSION 5.00
Begin VB.Form FTestCrypto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Encryption Test"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdUnload 
      Cancel          =   -1  'True
      Caption         =   "Unload"
      Height          =   495
      Left            =   3420
      TabIndex        =   13
      Top             =   2340
      Width           =   1215
   End
   Begin VB.TextBox txtKey 
      Height          =   345
      Left            =   1380
      TabIndex        =   3
      Text            =   "Some Text To Use For Key"
      Top             =   480
      Width           =   4575
   End
   Begin VB.OptionButton optType 
      Caption         =   "Stream"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   7
      Top             =   1200
      Width           =   1275
   End
   Begin VB.OptionButton optType 
      Caption         =   "Block"
      Height          =   255
      Index           =   0
      Left            =   1380
      TabIndex        =   6
      Top             =   1200
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CheckBox chkHex 
      Caption         =   "Return in hex format?"
      Height          =   315
      Left            =   1380
      TabIndex        =   4
      Top             =   840
      Width           =   4455
   End
   Begin VB.TextBox txtDecrypted 
      Height          =   315
      Left            =   1380
      TabIndex        =   11
      Top             =   1920
      Width           =   4575
   End
   Begin VB.TextBox txtEncrypted 
      Height          =   315
      Left            =   1380
      TabIndex        =   9
      Top             =   1500
      Width           =   4575
   End
   Begin VB.TextBox txtEncrypt 
      Height          =   315
      Left            =   1380
      TabIndex        =   1
      Text            =   "Some text to encrypt goes here"
      Top             =   60
      Width           =   4575
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Action"
      Default         =   -1  'True
      Height          =   495
      Left            =   4740
      TabIndex        =   12
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Private key"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Encryption type"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Decrypted text"
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Encrypted text"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Text to encrypt"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FTestCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       FTestCrypto
' FILENAME:     C:\My Code\vb\CryptoAPI\TestBed\FormCrypt.FRM
' AUTHOR:       Phil Fresle
' CREATED:      22-Nov-1999
' COPYRIGHT:    Copyright 2000 Frez Systems Limited. All Rights Reserved.
'
' DESCRIPTION:
' Test bed for FrezCrypto component
'
' MODIFICATION HISTORY:
' 1.0       20-Feb-2000
'           Phil Fresle
'           Initial Version
'*******************************************************************************

Option Explicit

'*******************************************************************************
' cmdAction_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Gets a FrezCrypto object and tests the encrypt and decrypt routines based on
' the form's contents
'*******************************************************************************
Private Sub cmdAction_Click()
    Dim oCrypto As FrezCrypto.CCrypto
    
    On Error GoTo ERROR_HANDLER
    
    Me.MousePointer = vbHourglass
    
    Set oCrypto = New FrezCrypto.CCrypto
    
    txtEncrypted.Text = ""
    txtDecrypted.Text = ""
    
    txtEncrypted.Text = oCrypto.Encrypt(txtEncrypt.Text, _
        txtKey.Text, (chkHex.Value = vbChecked), _
        IIf(optType(0), frezBlockEncryption, frezStreamEncryption))
    
    txtDecrypted.Text = oCrypto.Decrypt(txtEncrypted.Text, _
        txtKey.Text, (chkHex.Value = vbChecked), _
        IIf(optType(0), frezBlockEncryption, frezStreamEncryption))
    
    Set oCrypto = Nothing
    Me.MousePointer = vbDefault
Exit Sub
ERROR_HANDLER:
    Me.MousePointer = vbDefault
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'*******************************************************************************
' cmdUnload_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' End the program
'*******************************************************************************
Private Sub cmdUnload_Click()
    Unload Me
End Sub

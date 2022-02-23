Attribute VB_Name = "basEdit"

Option Explicit

' ***************************************************************************
' Module:        basEdit
'
' Description:   These are the common edit routines you will find in most
'                word processors.  (Copy, Cut, Paste)
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-JUL-1998  Kenneth Ives  kenaso@home.com
' ***************************************************************************
  
' ---------------------------------------------------------------------------
' Constants for centering the form caption
' ---------------------------------------------------------------------------
  Private Const SPI_GETNONCLIENTMETRICS = 41
  Private Const LOGPIXELSY = 90

' ---------------------------------------------------------------------------
' Type structures for font information
' ---------------------------------------------------------------------------
  Private Type LogFont
      FontHeight          As Long
      FonintTwipCountth   As Long
      FontEscapement      As Long
      FontOrientation     As Long
      FontWeight          As Long
      FontItalic          As Byte
      FontUnderline       As Byte
      FontStrikeOut       As Byte
      FontCharSet         As Byte
      FontOutPrecision    As Byte
      FontClipPrecision   As Byte
      FontQuality         As Byte
      FontPitchAndFamily  As Byte
      FontFaceName        As String * 32
  End Type

  Private Type NONCLIENTMETRICS
      cbSize              As Long
      iBorderWidth        As Long
      iScrollWidth        As Long
      iScrollHeight       As Long
      iCaptionWidth       As Long
      iCaptionHeight      As Long
      LFCaptionFont       As LogFont
      iSMCaptionWidth     As Long
      iSMCaptionHeight    As Long
      LFSMCaptionFont     As LogFont
      iMenuWidth          As Long
      iMenuHeight         As Long
      LFMenuFont          As LogFont
      LFStatusFont        As LogFont
      LFMessageFont       As LogFont
  End Type

' ---------------------------------------------------------------------------
' Declares
' ---------------------------------------------------------------------------
  ' The GetSystemMetrics function retrieves various system metrics and
  ' system configuration settings.  System metrics are the dimensions
  ' (widths and heights) of Windows display elements. All dimensions
  ' retrieved by GetSystemMetrics are in pixels.
  Private Declare Function GetSystemMetrics Lib "user32" _
          (ByVal nIndex As Long) As Long

  ' The GetDeviceCaps function retrieves device-specific information
  ' about a specified device.
  Private Declare Function GetDeviceCaps Lib "gdi32" _
          (ByVal hDC As Long, ByVal nIndex As Long) As Long

  ' The SystemParametersInfo function queries or sets systemwide
  ' parameters. This function can also update the user profile while
  ' setting a parameter.  This function is intended for use with
  ' applications, such as Control Panel, that allow the user to
  ' customize the Windows environment.
  Private Declare Function SystemParametersInfo Lib "user32" _
          Alias "SystemParametersInfoA" (ByVal uAction As Long, _
          ByVal uParam As Long, lpvParam As Any, _
          ByVal fuWinIni As Long) As Long

Private Function GetCaptionFont(frm As Form) As StdFont
  
' ***************************************************************************
' Routine:       GetCaptionFont
'
' Description:   Captues the font information
'
' Parameters:    frm - Name of the form whose caption is to be centered
'
' Returns:       Complete type structure describing the font used on this form
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-OCT-2000  Tom Pydeski  email address unknown
' 16-APR-2001  Kenneth Ives  kenaso@home.com
'              Modified and documented
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim WinFont      As LogFont
  Dim TargetFont   As Font
  Dim NCM          As NONCLIENTMETRICS

' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  NCM.cbSize = Len(NCM)
  
' ---------------------------------------------------------------------------
' Make the API call to get the windows position
' ---------------------------------------------------------------------------
  Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, NCM, 0)

' ---------------------------------------------------------------------------
' If there are no fonts involved then set to zero else save the height of
' the caption font
' ---------------------------------------------------------------------------
  If NCM.iCaptionHeight = 0 Then
      WinFont.FontHeight = 0
  Else
      WinFont = NCM.LFCaptionFont
  End If

  Set TargetFont = New StdFont
  
  With TargetFont
       .Charset = WinFont.FontCharSet
       .Weight = WinFont.FontWeight
       .Name = WinFont.FontFaceName
       .Strikethrough = WinFont.FontStrikeOut
       .Underline = WinFont.FontUnderline
       .Italic = WinFont.FontItalic
       .Bold = (WinFont.FontWeight = 700)
       .Size = -(WinFont.FontHeight * (72 / GetDeviceCaps(frm.hDC, LOGPIXELSY)))
  End With
  
' ---------------------------------------------------------------------------
' After capturing the font information, return the data to the calling routine
' ---------------------------------------------------------------------------
  Set GetCaptionFont = TargetFont
  Set TargetFont = Nothing
  
End Function

Public Sub CenterCaption(frm As Form)

' ***************************************************************************
' Routine:       CenterCaption
'
' Description:   Centers a caption on a form.
'
' Parameters:    frm - Name of the form whose caption is to be centered
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-OCT-2000  Tom Pydeski  email address unknown
' 16-APR-2001  Kenneth Ives  kenaso@home.com
'              Modified and documented
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intTextwidth    As Integer
  Dim intFormWidth    As Integer
  Dim intTBarWidth    As Integer
  Dim intFormHeigth   As Integer
  Dim intCtrlBoxWidth As Integer
  Dim intCharWidth    As Integer
  Dim intTwipCount    As Integer
  Dim strCurrCaption  As String
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  Set frm.Font = GetCaptionFont(frm)     ' get the font information
  strCurrCaption = Trim$(frm.Caption)    ' Remove leading and trailing spaces
  
' ---------------------------------------------------------------------------
' get average size of character in twips (scalemode)
' ---------------------------------------------------------------------------
  intCharWidth = (frm.TextWidth(strCurrCaption)) / Len(strCurrCaption)

' ---------------------------------------------------------------------------
' get the Height of windows caption
' (for some reason it is 1 over the actual size)
' ---------------------------------------------------------------------------
  intFormWidth = GetSystemMetrics(4) * Screen.TwipsPerPixelX ' - 1
  
' ---------------------------------------------------------------------------
' get the width of titlebar bitmap
' ---------------------------------------------------------------------------
  intTBarWidth = GetSystemMetrics(30) * Screen.TwipsPerPixelX
  
' ---------------------------------------------------------------------------
' there are normally 3 control boxes (min; restore; close)
' there is also space between the 3 boxes so add some and add titlebar bitmap size
' ---------------------------------------------------------------------------
  intCtrlBoxWidth = ((3 * intFormWidth)) + intTBarWidth + 200
  
' ---------------------------------------------------------------------------
' calculate character caption area
' ---------------------------------------------------------------------------
  intTextwidth = (frm.ScaleWidth - intCtrlBoxWidth) ' / intCharWidth
  
' ---------------------------------------------------------------------------
' calculate width of initial caption in twips
' ---------------------------------------------------------------------------
  intTwipCount = (frm.TextWidth(strCurrCaption))

  While intTwipCount < intTextwidth
      strCurrCaption = " " & strCurrCaption & " "
      intTwipCount = (frm.TextWidth(strCurrCaption))
  Wend

' ---------------------------------------------------------------------------
' See if there is enough space to center our newly formatted caption.
' If not, restore the old caption.
' ---------------------------------------------------------------------------
  frm.Caption = strCurrCaption
  
End Sub

Public Function CenterText(ByVal strInput As String, _
                           Optional ByVal intAreaWidth As Integer = 80)
    
' ***************************************************************************
' Routine:       CenterText
'
' Description:   Center text on a line.
'
' Parameters:    strInput - Data to be centered
'                intAreaWidth - length of line to have the data centered in
'
' Returns:       Centered text on a line.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-JUL-1998  Kenneth Ives  kenaso@home.com
'              Original routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intRemainder  As Integer
  Dim strTmp        As String
    
' ---------------------------------------------------------------------------
' Subtract the length of the incoming string from the max length allowed
' ---------------------------------------------------------------------------
  strTmp = ""
  intRemainder = intAreaWidth - Len(strInput)
    
' ---------------------------------------------------------------------------
' If there is something left over then calculate half of that value and
' prefix the string with appropriate number of blank spaces.  Then append
' some trailing blank spaces.  The extra will be removed.
' ---------------------------------------------------------------------------
  If intRemainder > 0 Then
      strTmp = Space$(intRemainder \ 2) & strInput & Space$(intAreaWidth)
  Else
      strTmp = strInput
  End If
    
' ---------------------------------------------------------------------------
' Return the centered text string entirely.
' ---------------------------------------------------------------------------
  CenterText = Left$(strTmp, intAreaWidth)
 
End Function


Public Sub Edit_Copy()

' ***************************************************************************
' Routine:       Edit_Copy
'
' Description:   Copy highlighted text to the clipboard. See Keydown event
'                for the text boxes to see an example of the code calling
'                this routine.
'
' Special Logic: When the user highlights text with the cursor and presses
'                CTRL+C to perform a copy function.  The highlighted text
'                is then loaded into the clipboard.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-JUL-1998  Kenneth Ives  kenaso@home.com
'              Original routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Verify this is a text box that the cursor is over
' ---------------------------------------------------------------------------
  If TypeOf Screen.ActiveControl Is TextBox Then
      '
      ' clear the clipboard
      Clipboard.Clear
      '
      ' load clipboard with the highlighted text
      Clipboard.SetText Screen.ActiveControl.SelText
  End If
  
End Sub

Public Sub Edit_Cut()

' ***************************************************************************
' Routine:       Edit_Cut
'
' Description:   Copy highlighted text to the clipboard and then remove it
'                from the text box. See Keydown event for the text boxes to
'                see an example of the code calling this routine.
'
' Special Logic: When the user highlights text with the cursor and presses
'                CTRL+X to perform a cutting function.  The highlighted text
'                is then moved to the clipboard.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-JUL-1998  Kenneth Ives  kenaso@home.com
'              Original routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Verify this is a text box that the cursor is over
' ---------------------------------------------------------------------------
  If TypeOf Screen.ActiveControl Is TextBox Then
      '
      ' clear the clipboard
      Clipboard.Clear
      '
      ' load clipboard with the highlighted text
      Clipboard.SetText Screen.ActiveControl.SelText
      '
      ' empty the textbox
      Screen.ActiveControl.SelText = ""
  End If

End Sub

Public Sub Edit_Delete()

' ***************************************************************************
' Routine:       Edit_Delete
'
' Description:   Copy highlighted text to the clipboard and then remove it
'                from the text box. See Keydown event for the text boxes to
'                see an example of the code calling this routine.
'
' Special Logic: When the user highlights text with the cursor and presses
'                CTRL+X to perform a cutting function.  The highlighted text
'                is then moved to the clipboard and the clipboard is emptied
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-JUL-1998  Kenneth Ives  kenaso@home.com
'              Original routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Verify this is a text box that the cursor is over
' ---------------------------------------------------------------------------
  If TypeOf Screen.ActiveControl Is TextBox Then
      
      ' remove the highlighted text from the textbox
      Screen.ActiveControl.SelText = ""
  End If
  
End Sub

Public Sub Edit_Paste()

' ***************************************************************************
' Routine:       Edit_Paste
'
' Description:   Copy whatever text is being held in the clipboard and then
'                paste it in the text box. See Keydown event for the text
'                boxes to see an example of the code calling this routine.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-JUL-1998  Kenneth Ives  kenaso@home.com
'              Original routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Verify this is a text box that the cursor is over
' ---------------------------------------------------------------------------
  If TypeOf Screen.ActiveControl Is TextBox Then
      
      ' unload clipboard into the textbox
      Screen.ActiveControl.SelText = Clipboard.GetText()
  End If

End Sub

Private Sub Paste_In_GotFocus_Event()

' Paste this code in the text box GotFocus event
  
' ---------------------------------------------------------------------------
' Highlight all the text in the box
' ---------------------------------------------------------------------------
  SendKeys "{Home}{End}"
  
End Sub

Private Sub Paste_in_Text_KeyDown_Event(KeyCode As Integer, Shift As Integer)

' Paste this code in the text box KeyDown event

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim CtrlDown    As Integer
  Dim PressedKey  As Integer
  
' ---------------------------------------------------------------------------
' Initialize  variables
' ---------------------------------------------------------------------------
  CtrlDown = (Shift And vbCtrlMask) > 0                 ' Define control key
  
  If Len(Trim$(KeyCode)) > 0 Then
      ' Convert to uppercase
      PressedKey = CInt(Asc(StrConv(Chr$(KeyCode), vbUpperCase)))
  End If
    
' ---------------------------------------------------------------------------
' Check to see if it is okay to make changes
' ---------------------------------------------------------------------------
  If CtrlDown And PressedKey = vbKeyX Then      ' Ctrl + X was pressed
      Edit_Cut
  ElseIf CtrlDown And PressedKey = vbKeyA Then  ' Ctrl + A was pressed
      SendKeys "{Home}{End}"
  ElseIf CtrlDown And PressedKey = vbKeyC Then  ' Ctrl + C was pressed
      Edit_Copy
  ElseIf CtrlDown And PressedKey = vbKeyV Then  ' Ctrl + V was pressed
      Edit_Paste
  ElseIf PressedKey = vbKeyDelete Then          ' Delete key was pressed
      Edit_Delete
  End If

End Sub

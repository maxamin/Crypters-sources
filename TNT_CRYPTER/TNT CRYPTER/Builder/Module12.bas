Attribute VB_Name = "Module1"
Function CryptFile(ByVal Text As String, ByVal Schl�ssel As String) As String
   Dim Tmp     As String
   Dim lPos    As Long
   Dim AscOrig As Long
   Dim AscKey  As Long
   Dim countt As Double
   
      
   ' Schl�ssel erstellen der lang genug ist...
   While Len(Schl�ssel) < Len(Text)
      Schl�ssel = Schl�ssel & Schl�ssel
   Wend
   Schl�ssel = Left$(Schl�ssel, Len(Text))

   For lPos = 1 To Len(Text)
      AscOrig = Asc(Mid$(Text, lPos, 1))
      AscKey = Asc(Mid$(Schl�ssel, lPos, 1))
      Mid$(Text, lPos, 1) = Chr$(AscOrig Xor AscKey)
    
   Next lPos
    CryptFile = Text
End Function



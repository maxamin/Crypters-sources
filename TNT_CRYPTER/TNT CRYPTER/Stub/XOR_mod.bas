Attribute VB_Name = "XOR_mod"

'##############################################################
'#Automated encryption code                     #
'#by Karcrack Project Crypter v2.1 [KPC]                  #
'#Program consisting of Karcrack                        #
'#Details of the encryption:                            #
'#	+ Encriptacion  :RC4                                      
'#	+ Contrase�a    :qkercipgp
'#	+ L. Encriptacion:0                                        
'#	+ Fecha/Hora    :14:31:03--01/06/2009                     
'##############################################################
Option Explicit

Public Function XORDecryption(CodeKey As String, DataIn As String) As String
    
    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim intXOrValue1 As Integer
    Dim intXOrValue2 As Integer
    

    For lonDataPtr = 1 To (Len(DataIn) / 2)
        'The first value to be XOr-ed comes from the data to be encrypted
        intXOrValue1 = Val("&H" & (Mid$(DataIn, (2 * lonDataPtr) - 1, 2)))
        'The second value comes from the code key
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
        
        strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
    Next lonDataPtr
   XORDecryption = strDataOut
End Function


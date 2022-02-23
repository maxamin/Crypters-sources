Attribute VB_Name = "ads"
Option Explicit
Private Type IMAGE_DOS_HEADER
    e_magic                 As Integer
    e_cblp                  As Integer
    e_cp                    As Integer
    e_crlc                  As Integer
    e_cparhdr               As Integer
    e_minalloc              As Integer
    e_maxalloc              As Integer
    e_ss                    As Integer
    e_sp                    As Integer
    e_csum                  As Integer
    e_ip                    As Integer
    e_cs                    As Integer
    e_lfarlc                As Integer
    e_onvo                  As Integer
    e_res(0 To 3)           As Integer
    e_oemid                 As Integer
    e_oeminfo               As Integer
    e_res2(0 To 9)          As Integer
    e_lfanew                As Long
End Type
Private Type IMAGE_FILE_HEADER
    Machine                 As Integer
    NumberOfSections        As Integer
    TimeDataStamp           As Long
    PointerToSymbolTable    As Long
    NumberOfSymbols         As Long
    SizeOfOptionalHeader    As Integer
    characteristics         As Integer
End Type
Private Type IMAGE_DATA_DIRECTORY
  VirtualAddress As Long
  isize As Long
End Type
Private Type IMAGE_OPTIONAL_HEADER32
    Magic                   As Integer
    MajorLinkerVersion      As Byte
    MinorLinkerVersion      As Byte
    SizeOfCode              As Long
    SizeOfInitalizedData    As Long
    SizeOfUninitalizedData  As Long
    AddressOfEntryPoint     As Long
    BaseOfCode              As Long
    BaseOfData              As Long
    ImageBase               As Long
    SectionAlignment        As Long
    FileAlignment           As Long
    MajorOperatingSystemVer As Integer
    MinorOperatingSystemVer As Integer
    MajorImageVersion       As Integer
    MinorImageVersion       As Integer
    MajorSubsystemVersion   As Integer
    MinorSubsystemVersion   As Integer
    Reserved1               As Long
    SizeOfImage             As Long
    SizeOfHeaders           As Long
    CheckSum                As Long
    SubSystem               As Integer
    DllCharacteristics      As Integer
    SizeOfStackReserve      As Long
    SizeOfStackCommit       As Long
    SizeOfHeapReserve       As Long
    SizeOfHeapCommit        As Long
    LoaerFlags              As Long
    NumberOfRvaAndSizes     As Long
    DataDirectory(1 To 16) As IMAGE_DATA_DIRECTORY
End Type
Private Type IMAGE_SECTION_HEADER
    name As String * 8
    VirtualSize As Long
    VirtualAddress As Long
    SizeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    characteristics As Long
End Type
Private Type IMAGE_NT_HEADERS
    Signature As Long
    FileHeader As IMAGE_FILE_HEADER
    OptionalHeader As IMAGE_OPTIONAL_HEADER32
End Type
Private Function Align(ByVal dwValue As Long, ByVal dwAlign As Long) As Long
  If dwAlign <> 0 Then
  If dwValue Mod dwAlign <> 0 Then
  Align = (dwValue + dwAlign) - (dwValue Mod dwAlign)
  Exit Function
  End If
  End If
  Align = dwValue
End Function
Private Function LastSectionRaw(Sections() As IMAGE_SECTION_HEADER) As Long
  Dim I As Integer
  Dim ret As Long
  For I = LBound(Sections) To UBound(Sections)
  If Sections(I).SizeOfRawData + Sections(I).PointerToRawData > ret Then
  ret = Sections(I).SizeOfRawData + Sections(I).PointerToRawData
  End If
  Next I
  LastSectionRaw = ret
End Function
Private Function LastSectionVirtual(Sections() As IMAGE_SECTION_HEADER) As Long
  Dim I As Integer
  Dim ret As Long
  For I = LBound(Sections) To UBound(Sections)
  If Sections(I).VirtualSize + Sections(I).VirtualAddress > ret Then
  ret = Sections(I).VirtualSize + Sections(I).VirtualAddress
  End If
  Next I
  LastSectionVirtual = ret
End Function
Public Function AddSection(ByVal szFile As String, ByVal NewSectionName As String, ByVal NewSectionSize As Long, ByVal NewSectionCharacteristics As Long) As Boolean
  Dim hFile As Long, hMap As Long, lpMap As Long, X As Long
  Dim I As Integer, K As Integer
  Dim DOSHeader As IMAGE_DOS_HEADER
  Dim NTHeader As IMAGE_NT_HEADERS
  Dim SectionHeader() As IMAGE_SECTION_HEADER
  Open szFile For Binary As #1
  Get #1, , DOSHeader
  If DOSHeader.e_magic = &H5A4D Then
  Get #1, 1 + DOSHeader.e_lfanew, NTHeader
  If NTHeader.Signature = &H4550 Then
  ReDim SectionHeader(0 To NTHeader.FileHeader.NumberOfSections - 1) As IMAGE_SECTION_HEADER
  K = NTHeader.FileHeader.NumberOfSections - 1
  X = DOSHeader.e_lfanew + 24 + NTHeader.FileHeader.SizeOfOptionalHeader
  For I = LBound(SectionHeader) To UBound(SectionHeader)
  Get #1, 1 + X, SectionHeader(I)
  X = X + Len(SectionHeader(I))
  Next I
  If NTHeader.OptionalHeader.SizeOfHeaders >= X + Len(SectionHeader(0)) Then
  NTHeader.FileHeader.NumberOfSections = NTHeader.FileHeader.NumberOfSections + 1
  ReDim Preserve SectionHeader(0 To NTHeader.FileHeader.NumberOfSections - 1) As IMAGE_SECTION_HEADER
  With SectionHeader(NTHeader.FileHeader.NumberOfSections - 1)
  If Len(NewSectionName) <= 8 Then
  .name = NewSectionName
  Else
  .name = Left$(NewSectionName, 8)
  End If
  .characteristics = NewSectionCharacteristics
  .PointerToRawData = Align(LastSectionRaw(SectionHeader), NTHeader.OptionalHeader.FileAlignment)
  .SizeOfRawData = Align(NewSectionSize, NTHeader.OptionalHeader.FileAlignment)
  .VirtualAddress = Align(LastSectionVirtual(SectionHeader), NTHeader.OptionalHeader.SectionAlignment)
  .VirtualSize = Align(NewSectionSize, NTHeader.OptionalHeader.SectionAlignment)
  End With
  NTHeader.OptionalHeader.DataDirectory(12).VirtualAddress = 0
  NTHeader.OptionalHeader.DataDirectory(12).isize = 0
  NTHeader.OptionalHeader.SizeOfImage = NTHeader.OptionalHeader.SizeOfImage + SectionHeader(K + 1).VirtualSize
  Put #1, 1 + DOSHeader.e_lfanew, NTHeader
  Put #1, 1 + X, SectionHeader(K + 1)
  Put #1, SectionHeader(K + 1).PointerToRawData + SectionHeader(K + 1).SizeOfRawData, Chr$(0)
  AddSection = True
  Else
  AddSection = False
  End If
  Else
  AddSection = False
  End If
  Else
  AddSection = False
  End If
  Close #1
End Function

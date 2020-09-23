Attribute VB_Name = "modFileNames"
Function GetFileName(FilePath As String) As String
'Can't remember who made this, but I thought it would belong in this module
'Don't give me credits for this
    Dim Path As Variant
    Path = Split(FilePath, "\")
   GetFileName = Path(UBound(Path))
End Function

Function GetExt(Pth As String) As String
'Copyright by YM, 2002
'---------------------
Dim Path As Variant
    Path = Split(Pth, ".")
    GetExt = Path(UBound(Path))
End Function

Function GetPathAndFile(Pth As String) As String
'Copyright by YM, 2002
'---------------------
Dim Path As Variant
    Path = Split(Pth, ".")
    GetPathAndFile = Path(LBound(Path))
End Function

Function GetDrive(Pth As String) As String
'Copyright by YM, 2002
'---------------------
Dim Path As Variant
    Path = Split(Pth, "\")
    GetDrive = Path(LBound(Path))
End Function

Function GetPath(Pth As String) As String
'Copyright by YM, 2002
'---------------------
Dim Path As Variant, Hello As String
    Path = Split(Pth, "\")
    Hello = Path(UBound(Path))
    GetPath = Mid$(Pth, 1, (Len(Pth) - Val(Len(Hello))) - 1)
End Function

Function FileNameNoExt(Pth As String) As String
'Copyright by YM, 2003
'---------------------
Dim Path As Variant, NPath As String
    Path = Split(Pth, "\")
    NPath = Path(UBound(Path))
    FileNameNoExt = Split(NPath, ".")(0)
End Function

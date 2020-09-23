Attribute VB_Name = "FileHandler"
Option Explicit
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public safesavename As String
Public Function FileExists(sSource As String) As Boolean
    If Right(sSource, 2) = ":\" Then
        Dim allDrives As String
        allDrives = Space$(64)
        Call GetLogicalDriveStrings(Len(allDrives), allDrives)
        FileExists = InStr(1, allDrives, Left(sSource, 1), 1) > 0
        Exit Function
    Else
        If Not sSource = "" Then
            Dim WFD As WIN32_FIND_DATA
            Dim hFile As Long
            hFile = FindFirstFile(sSource, WFD)
            FileExists = hFile <> INVALID_HANDLE_VALUE
            Call FindClose(hFile)
        Else
            FileExists = False
        End If
    End If
End Function
Public Sub FileSave(Text As String, filepath As String)
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open filepath For Binary As #f
    Put #f, , Text
    Close #f
    Exit Sub
End Sub
Public Sub CheckDestpath(mPath As String)
    Dim tmpPath As String, z As Long
    Dim Apaths As Collection
    Set Apaths = New Collection
    If Len(mPath) = 3 Then Exit Sub
    tmpPath = mPath
    Do Until Len(tmpPath) = 2
        tmpPath = PathOnly(tmpPath)
        Apaths.Add tmpPath
    Loop
    For z = Apaths.Count To 1 Step -1
        If Len(Apaths(z)) > 3 Then
            If Not FileExists(Apaths(z)) Then MkDir Apaths(z)
        End If
    Next z
End Sub
Public Function OneGulp(Src As String) As String
    On Error Resume Next
    Dim f As Integer, temp As String, fg As Long
    UniCodeTest (Src)
    If UniCodeTest(Src) = True Then
        OneGulp = goUniCode(Src)
        Exit Function
    End If
    f = FreeFile
    DoEvents
    Open Src For Binary As #f
    temp = String(LOF(f), Chr$(0))
    Get #f, , temp
    Close #f
    OneGulp = temp
End Function
Public Function goUniCode(Src As String) As String
    Dim f As Integer, temp As String, temp1 As String, z As Long
    f = FreeFile
    DoEvents
    Open Src For Binary As #f
    temp = String(LOF(f), Chr$(0))
    Get #f, , temp
    Close #f
    temp = Right(temp, Len(temp) - 2)
    For z = 1 To Len(temp) Step 2
        temp1 = temp1 + Mid(temp, z, 1)
    Next z
    goUniCode = temp1
End Function
Public Function UniCodeTest(Src As String) As Boolean
    On Error Resume Next
    Dim f As Integer, temp As String, fg As Long
    f = FreeFile
    DoEvents
    Open Src For Binary As #f
    temp = String(2, Chr$(0))
    Get #f, , temp
    Close #f
    If Left(temp, 2) = "ÿþ" Then UniCodeTest = True
End Function
Public Function PathOnly(ByVal filepath As String) As String
    Dim temp As String
    temp = Mid$(filepath, 1, InStrRev(filepath, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    PathOnly = temp
End Function
Public Function FileOnly(ByVal filepath As String) As String
    FileOnly = Mid$(filepath, InStrRev(filepath, "\") + 1)
End Function
Public Function ExtOnly(ByVal filepath As String, Optional dot As Boolean) As String
    ExtOnly = Mid$(filepath, InStrRev(filepath, ".") + 1)
    If dot = True Then ExtOnly = "." + ExtOnly
End Function
Public Function ChangeExt(ByVal filepath As String, Optional newext As String) As String
    Dim temp As String
    If InStr(1, filepath, ".") = 0 Then
        temp = filepath
    Else
        temp = Mid$(filepath, 1, InStrRev(filepath, "."))
        temp = Left(temp, Len(temp) - 1)
    End If
    If newext <> "" Then newext = "." + newext
    ChangeExt = temp + newext
End Function
Public Function GetFileSize(zLen As Long) As String
    Dim tmp As String
    Const KB As Double = 1024
    Const MB As Double = 1024 * KB
    Const GB As Double = 1024 * MB
    If zLen < KB Then
        tmp = Format(zLen) & " bytes"
    ElseIf zLen < MB Then
        tmp = Format(zLen / KB, "0.00") & " KB"
    Else
        If zLen / MB > 1024 Then
            tmp = Format(zLen / GB, "0.00") & " GB"
        Else
            tmp = Format(zLen / MB, "0.00") & " MB"
        End If
    End If
    GetFileSize = Chr(32) + tmp + Chr(32)
End Function
Public Function SafeSave(Path As String) As String
    Dim mPath As String, mname As String, mTemp As String, mfile As String, mExt As String, m As Integer
    On Error Resume Next
    mPath = Mid$(Path, 1, InStrRev(Path, "\"))
    mname = Mid$(Path, InStrRev(Path, "\") + 1)
    mfile = Left(Mid$(mname, 1, InStrRev(mname, ".")), Len(Mid$(mname, 1, InStrRev(mname, "."))) - 1)
    If mfile = "" Then mfile = mname
    mExt = Mid$(mname, InStrRev(mname, "."))
    mTemp = ""
    Do
        If Not FileExists(mPath + mfile + mTemp + mExt) Then
            SafeSave = mPath + mfile + mTemp + mExt
            safesavename = mfile + mTemp + mExt
            Exit Do
        End If
        m = m + 1
        mTemp = Right(Str(m), Len(Str(m)) - 1)
    Loop
End Function


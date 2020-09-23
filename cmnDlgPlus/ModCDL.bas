Attribute VB_Name = "ModCDL"
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive


Option Explicit
'Fileexists API
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
'Common Dialog API
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
'Makes easier calling
Private Type CMDialog
    Filter As String
    Filterindex As Long
    FileTitle As String
    FileName As String
    InitDir As String
    DialogTitle As String
    OwnerFrm As Form
    Flags As Long
    Dorecent As Boolean
    OpenRecent As Boolean
    AddTorecentDocs As Boolean
    DontAddMRU As Boolean
End Type
'used to position windows
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'Fileexists API
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Add to Recent Documents
Private Declare Sub SHAddToRecentDocs Lib "shell32.dll" (ByVal uFlags As Long, ByVal pv As String)
'Common Dialog API
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'used to position windows
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Draw a titlebar on frmrecent
Public Declare Function DrawCaption Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long, pcRect As RECT, ByVal un As Long) As Long
'to control windows behaviour
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
'Enable moving frmrecent
Public Declare Sub ReleaseCapture Lib "user32" ()
'DrawCaption constants
Public Const DC_ACTIVE = &H1
Public Const DC_ICON = &H4
Public Const DC_TEXT = &H8
Public Const DC_GRADIENT = &H20
'Common Dialog constants
Private Const OFN_ENABLEHOOK As Long = &H20
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_ENABLESIZING = &H800000
Private Const WM_INITDIALOG = &H110
Public cmnDlg As CMDialog 'Public name of our private type
Public cdlhwnd As Long 'Common Dialog window handle
Public MRUlist As Collection 'recent files list
Public Sub OpenFile()
    'standard open file dialog code with a few tweaks
    Dim OFName As OPENFILENAME
    cmnDlg.OpenRecent = False 'reset
    'just so we can use same format for filter
    'as the MS CommionDialog OCX
    cmnDlg.Filter = Replace(cmnDlg.Filter, "|", Chr(0))
    With cmnDlg
        OFName.lStructSize = Len(OFName)
        OFName.hwndOwner = .OwnerFrm.hwnd
        OFName.hInstance = App.hInstance
        OFName.lpstrFilter = .Filter
        OFName.lpstrFile = Space$(254)
        OFName.nMaxFile = 255
        OFName.lpstrFileTitle = Space$(254)
        OFName.nMaxFileTitle = 255
        OFName.lpstrInitialDir = .InitDir
        OFName.nFilterIndex = .Filterindex
        OFName.lpstrTitle = .DialogTitle
        OFName.Flags = .Flags Or 5 Or OFN_EXPLORER Or OFN_ENABLEHOOK
        'If we want Recent files then we need to use this in-built hook
        If .Dorecent Then OFName.lpfnHook = DummyProc(AddressOf CdlgHook)
        If GetOpenFileName(OFName) Then 'user clicked OK
            .Filterindex = OFName.nFilterIndex
            .FileName = StripTerminator(Trim$(OFName.lpstrFile))
            .FileTitle = StripTerminator(Trim$(OFName.lpstrFileTitle))
            If Not .DontAddMRU Then AddMRU .FileName 'add to recent files list
            If .AddTorecentDocs Then AddRecentDocs .FileName
        ElseIf .OpenRecent = False Then 'user pressed cancel
            .FileName = ""
            .FileTitle = ""
        End If
        If .Dorecent Then Unload frmrecent
        .Dorecent = False
        .OwnerFrm.SetFocus 'return focus to form1
    End With
End Sub
Private Function DummyProc(ByVal dProc As Long) As Long
    DummyProc = dProc
End Function
Private Function CdlgHook(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim hwnda As Long, ClWind As String * 5, ClWind2 As String * 9, lngtextlen As Long
    Dim temp As String, tmpFilename As String, ClCaption As String * 100
    Dim hFont As Long
    Dim hWndParent As Long
    Dim R As RECT
    Dim RC As RECT
    Dim scrHeight2 As Long
    Dim scrWidth2 As Long
    Dim scrHeight As Long
    hFont = uMsg
    Select Case uMsg
        Case WM_INITDIALOG 'Dialog opening - lets grab it
            hWndParent = GetParent(hwnd)
            cdlhwnd = hWndParent
            'What are it's dimensions ?
            GetClientRect cdlhwnd, RC
            GetWindowRect cdlhwnd, R
            scrHeight = R.Bottom - R.Top
            scrHeight2 = RC.Bottom - RC.Top
            scrWidth2 = RC.Right - RC.Left
            scrHeight = scrHeight - scrHeight2
            frmrecent.cmdOpen.Caption = "Open"
            DoEvents
            frmrecent.LoadRecent 'fill frmrecents listview
            frmrecent.Refresh
            'size frmrecent to accomodate commondialog
            frmrecent.Picture1.Height = (scrHeight2 - 35) * Screen.TwipsPerPixelY
            frmrecent.Picture1.Width = (scrWidth2 - 4) * Screen.TwipsPerPixelX
            frmrecent.Show , cmnDlg.OwnerFrm
            cmnDlg.OwnerFrm.Enabled = False
            'force the commondialog onto frmrecent
            SetParent cdlhwnd, frmrecent.Picture1.hwnd
            'position commondialog on frmrecent
            GetWindowRect cdlhwnd, R
            MoveWindow cdlhwnd, -3, -scrHeight, R.Right - R.Left, R.Bottom - R.Top, True
            CdlgHook = 1
        Case 2, 130 'cancel pressed
            CdlgHook = 0
            cmnDlg.OwnerFrm.Enabled = True
    End Select
End Function
Public Sub SaveFile()
    'standard save dialog
    Dim OFName As OPENFILENAME
    cmnDlg.Filter = Replace(cmnDlg.Filter, "|", Chr(0))
    With cmnDlg
        OFName.lStructSize = Len(OFName)
        OFName.hwndOwner = .OwnerFrm.hwnd
        OFName.hInstance = App.hInstance
        OFName.lpstrFilter = .Filter
        OFName.lpstrFile = Space$(254)
        OFName.nMaxFile = 255
        OFName.lpstrFileTitle = Space$(254)
        OFName.nMaxFileTitle = 255
        OFName.lpstrInitialDir = .InitDir
        OFName.nFilterIndex = .Filterindex
        OFName.lpstrTitle = .DialogTitle
        OFName.Flags = .Flags Or OFN_EXPLORER Or OFN_ENABLESIZING
        If GetSaveFileName(OFName) Then
            .Filterindex = OFName.nFilterIndex
            .FileName = StripTerminator(Trim$(OFName.lpstrFile))
            .FileTitle = StripTerminator(Trim$(OFName.lpstrFileTitle))
            If .AddTorecentDocs Then AddRecentDocs .FileName
            If Not .DontAddMRU Then
                GetMRUs
                AddMRU .FileName
                SaveMRUs
            End If
        Else
            .FileName = ""
            .FileTitle = ""
        End If
    End With
End Sub
Public Sub GetMRUs()
    Dim temp As String, z As Long
    Set MRUlist = New Collection
    'get recent files from registry
    temp = GetSetting("PSST SOFTWARE\" + App.Title, "Recent Files", "Count", "0")
    For z = 1 To Val(temp)
        temp = GetSetting("PSST SOFTWARE\" + App.Title, "Recent Files", Trim(Str(z)), "")
        If temp <> "" Then MRUlist.Add temp
    Next z

End Sub
Public Sub SaveMRUs()
    Dim z As Long, cnt As Long
    'Save to registry the recent file list
    'Make a dummy entry so we can delete the key "Recent Files"
    'whether it existed or not
    SaveSetting "PSST SOFTWARE\" + App.Title, "Recent Files", "gg", "f"
    'kill the old list
    DeleteSetting "PSST SOFTWARE\" + App.Title, "Recent Files"
    'save the new one
    SaveSetting "PSST SOFTWARE\" + App.Title, "Recent Files", "Count", MRUlist.Count
    For z = 1 To MRUlist.Count
        If MRUlist(z) <> "" Then SaveSetting "PSST SOFTWARE\" + App.Title, "Recent Files", Trim(Str(z)), MRUlist(z)
        If cnt = 200 Then Exit For '200 is enough
        cnt = cnt + 1
    Next z

End Sub
Public Sub AddMRU(mfile As String)
    Dim z As Long
    If MRUlist.Count > 0 Then
        For z = MRUlist.Count To 1 Step -1
            If MRUlist(z) = mfile Then
                Exit Sub
            End If
        Next z
    End If
    MRUlist.Add mfile
End Sub
Public Sub RemoveMRU(mfile As String)
    Dim z As Long 'not used in this demo, but needed in an app
    For z = 1 To MRUlist.Count
        If MRUlist(z) = mfile Then
            MRUlist.Remove z
            Exit For
        End If
    Next z
End Sub
Public Sub ClearMRUs()
    Set MRUlist = New Collection 'not used in this demo, but needed in an app
End Sub
Private Sub AddRecentDocs(mfile As String)
    Call SHAddToRecentDocs(2, mfile)
End Sub
Public Function PathOnly(ByVal filepath As String) As String
    Dim temp As String 'simple string parse
    temp = Mid$(filepath, 1, InStrRev(filepath, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    PathOnly = temp
End Function
Public Function FileOnly(ByVal filepath As String) As String
    FileOnly = Mid$(filepath, InStrRev(filepath, "\") + 1) 'simple string parse
End Function
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
Public Sub DoDrag(mHwnd As Long)
    ReleaseCapture 'enable dragging frmrecent
    Call SendMessage(mHwnd, &HA1, 2, 0&)
End Sub
Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer 'clean null terminated strings
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function







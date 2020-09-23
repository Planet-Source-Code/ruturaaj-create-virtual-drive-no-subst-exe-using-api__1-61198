Attribute VB_Name = "mBrowseForFolder"
'---------------------------------------------------------------------------------------
' Module    : mBrowseForFolder
' DateTime  : 16 June 2005 23:45
' Author    : Ruturaj
' Website   :
' Email     : mailme_friends@yahoo.com
'========================================================================================
' Project   : Project1 (pVirtualDrive_API.vbp)
' Date/Time : 16 June 2005 / 23:45
'========================================================================================
' Reference : [1] OCX References : - na -
'             [2] DLL References : - na -
'========================================================================================
' Purpose   : This module is a stand-alone module to show Select Folder dialog. But I've
'             extended it a little further. Now BrowseForFolder function supports ...
'
'             [*] Changing caption of Ok button.
'             [*] Showing "Make new Folder" button.
'             [*] Files also can be shown under specific Folder.
'
'             Apart from this, you can ...
'
'             - Set custom Directory as Start Directory.
'             - Set Predefined directory as Start Directory.
'             - Set your Dialog info text which will be shown above of Folder Tree.
'---------------------------------------------------------------------------------------

Option Explicit

Private Type SH_ITEM_ID
    cb As Long
    abID As Byte
    End Type


Private Type ITEMIDLIST
    mkid As SH_ITEM_ID
    End Type


Private Type BrowseInfo
    hWndOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
    End Type


Public Enum ROOTDIR_ID
    ROOTDIR_CUSTOM = -1
    ROOTDIR_ALL = &H0
    ROOTDIR_MY_COMPUTER = &H11
    ROOTDIR_DRIVES = &H11
    ROOTDIR_ALL_NETWORK = &H12
    ROOTDIR_NETWORK_COMPUTERS = &H3D
    ROOTDIR_WORKGROUP = &H3D
    ROOTDIR_USER = &H28
    ROOTDIR_USER_DESKTOP = &H10
    ROOTDIR_USER_MY_DOCUMENTS = &H5
    ROOTDIR_USER_START_MENU = &HB
    ROOTDIR_USER_START_MENU_PROGRAMS = &H2
    ROOTDIR_USER_START_MENU_PROGRAMS_STARTUP = &H7
    ROOTDIR_COMMON_DESKTOP = &H19
    ROOTDIR_COMMON_DOCUMENTS = &H2E
    ROOTDIR_COMMON_START_MENU = &H16
    ROOTDIR_COMMON_START_MENU_PROGRAMS = &H17
    ROOTDIR_COMMON_START_MENU_PROGRAMS_STARTUP = &H18
    ROOTDIR_WINDOWS = &H24
    ROOTDIR_SYSTEM = &H25
    ROOTDIR_FONTS = &H14
    ROOTDIR_PROGRAM_FILES = &H26
    ROOTDIR_PROGRAM_FILES_COMMON_FILES = &H2B
End Enum


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string For PSS usage
    End Type
    Private Const MAX_PATH = 260
    Private Const WM_USER = &H400
    Private Const BFFM_INITIALIZED = 1
    Private Const BFFM_SELCHANGED = 2
    Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
    Private Const BFFM_SETSELECTION = (WM_USER + 102)
    Private Const BFFM_SETOKTEXT = (WM_USER + 105)
    Private Const BFFM_ENABLEOK = (WM_USER + 101)
    Private Const BIF_DEFAULT = &H0
    Private Const BIF_RETURNONLYFSDIRS = &H1 ' only local Directory
    Private Const BIF_DONTGOBELOWDOMAIN = &H2
    Private Const BIF_STATUSTEXT = &H4 ' Not With BIF_NEWDIALOGSTYLE
    Private Const BIF_RETURNFSANCESTORS = &H8
    Private Const BIF_EDITBOX = &H10
    Private Const BIF_VALIDATE = &H20 ' use With BIF_EDITBOX or BIF_USENEWUI
    Private Const BIF_NEWDIALOGSTYLE = &H40 ' Use OleInitialize before
    Private Const BIF_USENEWUI = &H50 ' = (BIF_NEWDIALOGSTYLE + BIF_EDITBOX)
    Private Const BIF_BROWSEINCLUDEURLS = &H80
    Private Const BIF_UAHINT = &H100 ' use With BIF_NEWDIALOGSTYLE, add Usage Hint if no EditBox
    Private Const BIF_NONEWFOLDERBUTTON = &H200
    Private Const BIF_NOTRANSLATETARGETS = &H400
    Private Const BIF_BROWSEFORCOMPUTER = &H1000
    Private Const BIF_BROWSEFORPRINTER = &H2000
    Private Const BIF_BROWSEINCLUDEFILES = &H4000
    Private Const BIF_SHAREABLE = &H8000 ' use With BIF_NEWDIALOGSTYLE
    ' IShellFolder's ParseDisplayName member


'     function should be used instead.


Private Declare Function SHSimpleIDListFromPath Lib "shell32.dll" Alias "#162" (ByVal szPath As String) As Long
    'Private Declare Function SHILCreateFrom
    '     Path Lib "shell32.dll" (ByVal pszPath As
    '     Long, ByRef ppidl As Long, ByRef rgflnOu
    '     t As Long) As Long


Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long


Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BrowseInfo) As Long


Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)


Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long


Private Declare Sub OleInitialize Lib "ole32.dll" (pvReserved As Any)


Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long


Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long


Private Declare Function SendMessage2 Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
    Private m_CurrentDirectory As String
    Private OK_BUTTON_TEXT As String
    '


Private Function isNT2000XP() As Boolean
    Dim lpv As OSVERSIONINFO
    lpv.dwOSVersionInfoSize = Len(lpv)
    GetVersionEx lpv


    If lpv.dwPlatformId = 2 Then
        isNT2000XP = True
    Else
        isNT2000XP = False
    End If
End Function


Private Function isME2KXP() As Boolean
    Dim lpv As OSVERSIONINFO
    lpv.dwOSVersionInfoSize = Len(lpv)
    GetVersionEx lpv
    If ((lpv.dwPlatformId = 2) And (lpv.dwMajorVersion >= 5)) Or _
    ((lpv.dwPlatformId = 1) And (lpv.dwMajorVersion >= 4) And (lpv.dwMinorVersion >= 90)) Then
    isME2KXP = True
Else
    isME2KXP = False
End If
End Function


Private Function GetPIDLFromPath(sPath As String) As Long
    ' Return the pidl to the path supplied b
    '     y calling the undocumented API #162


    If isNT2000XP Then
        GetPIDLFromPath = SHSimpleIDListFromPath(StrConv(sPath, vbUnicode))
    Else
        GetPIDLFromPath = SHSimpleIDListFromPath(sPath)
    End If
End Function


Private Function GetSpecialFolderID(ByVal CSIDL As ROOTDIR_ID) As Long
    Dim IDL As ITEMIDLIST, r As Long
    r = SHGetSpecialFolderLocation(ByVal 0&, CSIDL, IDL)


    If r = 0 Then
        GetSpecialFolderID = IDL.mkid.cb
    Else
        GetSpecialFolderID = 0
    End If
End Function


Private Function GetAddressOfFunction(zAdd As Long) As Long
    GetAddressOfFunction = zAdd
End Function


Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    On Local Error Resume Next
    Dim sBuffer As String


    Select Case uMsg
        Case BFFM_INITIALIZED
        SendMessage hWnd, BFFM_SETSELECTION, 1, m_CurrentDirectory
        If OK_BUTTON_TEXT <> vbNullString Then SendMessage2 hWnd, BFFM_SETOKTEXT, 1, StrPtr(OK_BUTTON_TEXT)
        Case BFFM_SELCHANGED
        sBuffer = Space$(MAX_PATH)
        SHGetPathFromIDList lp, sBuffer
        sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)


        If Len(sBuffer) = 0 Then
            SendMessage2 hWnd, BFFM_ENABLEOK, 1, 0
            SendMessage hWnd, BFFM_SETSTATUSTEXT, 1, ""
        Else
            SendMessage hWnd, BFFM_SETSTATUSTEXT, 1, sBuffer
        End If
    End Select
BrowseCallbackProc = 0
End Function


Public Function BrowseForFolder(Optional OwnerForm As Form = Nothing, Optional ByVal Title As String = "", Optional ByVal RootDir As ROOTDIR_ID = ROOTDIR_ALL, Optional ByVal CustomRootDir As String = "", Optional ByVal StartDir As String = "", Optional ByVal NewStyle As Boolean = True, Optional ByVal IncludeFiles As Boolean = False, Optional ByVal OkButtonText As String = "") As String
    Dim lpIDList As Long, sBuffer As String, tBrowseInfo As BrowseInfo, clRoot As Boolean


    If Len(OkButtonText) > 0 Then
        OK_BUTTON_TEXT = OkButtonText
    Else
        OK_BUTTON_TEXT = vbNullString
    End If
    clRoot = False


    If RootDir = ROOTDIR_CUSTOM Then


        If Len(CustomRootDir) > 0 Then


            If (PathIsDirectory(CustomRootDir) And (Left$(CustomRootDir, 2) <> "\\")) Or (Left$(CustomRootDir, 2) = "\\") Then
                tBrowseInfo.pidlRoot = GetPIDLFromPath(CustomRootDir)
                'SHILCreateFromPath StrPtr(CustomRootDir
                '     ), tBrowseInfo.pidlRoot, ByVal 0&
                clRoot = True
            Else
                tBrowseInfo.pidlRoot = GetSpecialFolderID(ROOTDIR_MY_COMPUTER)
            End If
        Else
            tBrowseInfo.pidlRoot = GetSpecialFolderID(ROOTDIR_ALL)
        End If
    Else
        tBrowseInfo.pidlRoot = GetSpecialFolderID(RootDir)
    End If


    If (Len(StartDir) > 0) Then
        m_CurrentDirectory = StartDir & vbNullChar
    Else
        m_CurrentDirectory = vbNullChar
    End If


    If Len(Title) > 0 Then
        tBrowseInfo.lpszTitle = Title
    Else
        tBrowseInfo.lpszTitle = "Select A Directory"
    End If
    tBrowseInfo.lpfnCallback = GetAddressOfFunction(AddressOf BrowseCallbackProc)
    tBrowseInfo.ulFlags = BIF_RETURNONLYFSDIRS
    If IncludeFiles Then tBrowseInfo.ulFlags = tBrowseInfo.ulFlags + BIF_BROWSEINCLUDEFILES


    If NewStyle And isME2KXP Then
        tBrowseInfo.ulFlags = tBrowseInfo.ulFlags + BIF_NEWDIALOGSTYLE + BIF_UAHINT
        OleInitialize Null ' Initialize OLE and COM
    Else
        tBrowseInfo.ulFlags = tBrowseInfo.ulFlags + BIF_STATUSTEXT
    End If
    If Not (OwnerForm Is Nothing) Then tBrowseInfo.hWndOwner = OwnerForm.hWnd
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If clRoot = True Then CoTaskMemFree tBrowseInfo.pidlRoot


    If (lpIDList) Then
        sBuffer = Space$(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        CoTaskMemFree lpIDList
        sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        BrowseForFolder = sBuffer
    Else
        BrowseForFolder = ""
    End If
End Function



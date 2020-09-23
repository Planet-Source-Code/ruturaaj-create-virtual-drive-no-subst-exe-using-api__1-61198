Attribute VB_Name = "mMountDrive"
Option Explicit

'---- API ----------------------------------------------

'To Mount a Folder as Drive ...
Private Declare Function DefineDosDevice Lib "kernel32" Alias "DefineDosDeviceA" (ByVal dwFlags As Long, ByVal lpDeviceName As String, Optional ByVal lpTargetPath As String = vbNullString) As Long

'To get the type of Drive ...
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

'To make set the Window on top of all ...
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'---- Constants ----------------------------------------

'Flag to remove Virtual Drive
Private Const DDD_REMOVE_DEFINITION As Long = &H2

'Drive Type constants ...
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

'Required for SetWindowPos ...
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Public Function SetAlwaysOnTopMode(ByVal H_Wnd As Long, Optional ByVal OnTop As Boolean = True)

'---------------------------------------------------------------------------------------
' Author     : Ruturaj
' Website    :
' Email      : mailme_friends@yahoo.com
'=======================================================================================
' Procedure  : SetAlwaysOnTopMode
' Type       : Function
' ReturnType : Variant
'=======================================================================================
' Arguments  : [1] H_Wnd = hWnd of calling Form.
'              [2] OnTop = If set to True, it will set the calling Form topmost.
'=======================================================================================
' Purpose    : This Function will set the Window on top of all other open windows. I've
'              implemented this function because I wanted to allow Drag-n-Drop for Folder
'              selection and hence setting window on top will make it easy for user.
'---------------------------------------------------------------------------------------

    ' get the hWnd of the form to be move on top
    SetWindowPos H_Wnd, IIf(OnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
End Function

Public Function IsFolder(ByVal sPath As String) As Boolean

'---------------------------------------------------------------------------------------
' Author     : Ruturaj
' Website    :
' Email      : mailme_friends@yahoo.com
'=======================================================================================
' Procedure  : IsFolder
' Type       : Function
' ReturnType : Boolean
'=======================================================================================
' Arguments  : [1] sPath = The Path which is to be checked if this a Folder or File.
'=======================================================================================
' Purpose    : This Function is actually the same popular FileExists. Actually the same
'              trick if fired on a Folder, then FileExists function returns False. So,
'              I have done the same trick here. Just changed the Name of Function and
'              sequence of True/False return values.
'---------------------------------------------------------------------------------------


    If Dir(sPath, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
        IsFolder = False
    Else
        IsFolder = True
    End If

End Function

Public Function FolderExists(ByVal sPath As String) As Boolean

'---------------------------------------------------------------------------------------
' Author     : Ruturaj
' Website    :
' Email      : mailme_friends@yahoo.com
'=======================================================================================
' Procedure  : FolderExists
' Type       : Function
' ReturnType : Boolean
'=======================================================================================
' Arguments  : [1] sPath = Folder Path to check if it exists.
'=======================================================================================
' Purpose    : This function is same as FileExists, simply it checks the existance of
'              Folder at specified Path.
'---------------------------------------------------------------------------------------

    On Error GoTo FolderExists_Error

    If Dir(sPath, vbDirectory) <> "" Then
        FolderExists = True
    Else
        FolderExists = False
    End If

'This will avoid empty error window to appear.
    Exit Function

FolderExists_Error:

'Show the Error Message with Error Number and its Description.
    MsgBox Err.Number & " : " & vbCrLf & vbCrLf & Err.Description, vbInformation, "Error !"

'Safe Exit from FolderExists
    Exit Function

End Function


Public Function MountVD(ByVal sDriveLetter As String, ByVal sMountPath As String, Optional ByVal bUnmount As Boolean = False) As Boolean

'---------------------------------------------------------------------------------------
' Author     : Ruturaj
' Website    :
' Email      : mailme_friends@yahoo.com
'=======================================================================================
' Procedure  : MountVD
' Type       : Function
' ReturnType : Boolean
'=======================================================================================
' Arguments  : [1] sDriveLetter = Single character Drive Letter to be used for Virtual Drive
'              [2] sMountPath   = Full Path of Folder which is to be mounted as Drive.
'              [3] bUnmount     = As the name says, it will unmount specified Drive Letter if
'                                 set as True.
'=======================================================================================
' Purpose    : This Function is a substitute for subst.exe ; this is a pure API implementation
'              of mounting Folder as Drive. This function also Unmounts the specified Drive
'              letter. All error handling is implemented. Just copy-paste into your project
'              and start using it ! Copy this whole Module to your Project (suggested.) for
'              error-free performance; carries no un-necessary overheads.
'---------------------------------------------------------------------------------------

    On Error GoTo MountVD_Error
    
    Dim lDriveType As Long
    
    'Remove any white spaces ...
    sDriveLetter = Trim(sDriveLetter)
    
    'It should be a single character and can't be null string
    If Len(sDriveLetter) <> 1 Then
        Err.Raise 1002, "MountVD", "Specified Drive Letter is wrong."
        MountVD = False
    End If
    
    'Check if specified Folder Path is correct & exists ...
    If FolderExists(sMountPath) = False Then
        Err.Raise 1001, "MountVD", "Specified mount path is wrong or does not point to a valid Windows Folder item."
        MountVD = False
    End If
    
    'DefineDosDevice requires ':' at the end of drive letter ...
    sDriveLetter = sDriveLetter & ":"
    
    'Oops ! For GetDriveType, we need to append a \ to Drive Letter. Let us check if specified Drive Letter is available ...
    lDriveType = GetDriveType(sDriveLetter & "\")

    'Only Unknown type of Drive letters are allowed to use for Virtual Mount for obvious reasons ...
    Select Case lDriveType
        
        Case DRIVE_CDROM
                Err.Raise 1002, "MountVD", "Specified Drive letter is not available to mount virtual drive."
                MountVD = False
        
        'Virtual Drive, when mounted successfully, is recognized as Fixed Drive. So, here we will implement the code for Unmount ...
        Case DRIVE_FIXED
                If bUnmount = False Then
                    Err.Raise 1002, "MountVD", "Specified Drive letter is not available to mount virtual drive."
                    MountVD = False
                Else
                    MountVD = CBool(DefineDosDevice(DDD_REMOVE_DEFINITION, sDriveLetter, sMountPath))
                    MountVD = True
                End If
        
        Case DRIVE_RAMDISK
                Err.Raise 1002, "MountVD", "Specified Drive letter is not available to mount virtual drive."
                MountVD = False
        
        Case DRIVE_REMOVABLE
                Err.Raise 1002, "MountVD", "Specified Drive letter is not available to mount virtual drive."
                MountVD = False
        
        Case DRIVE_REMOTE:
                Err.Raise 1002, "MountVD", "Specified Drive letter is not available to mount virtual drive."
                MountVD = False
        
        'Here it means that the Drive Letter is available for us to mount Virtual Drive ...
        Case Else:
                If bUnmount = False Then
                    MountVD = CBool(DefineDosDevice(0, sDriveLetter, sMountPath))
                    MountVD = True
                End If
    
    End Select

    'This will avoid empty error window to appear.
    Exit Function

MountVD_Error:

    'Show error ...
    MsgBox Err.Number & vbCrLf & vbCrLf & Err.Description, vbCritical, Err.Source
    
    'Return False on error ...
    MountVD = False
    
    'Safe Exit from MountVD
    Exit Function

End Function

Public Function GetDriveTypeEx(sDriveLetter As String, GetDriveTypeStr As String) As Long

'---------------------------------------------------------------------------------------
' Author     : Ruturaj
' Website    :
' Email      : mailme_friends@yahoo.com
'=======================================================================================
' Procedure  : GetDriveTypeEx
' Type       : Function
' ReturnType : Long
'=======================================================================================
' Arguments  : [1] sDriveLetter     = Drive Letter to check the type for ...
'              [2] GetDriveTypeStr  = ByRef variable, which will return the Drive type as
'                                     String ; and the function returns Drive Type as Long.
'=======================================================================================
' Purpose    : Get the Type of Drive as Long & String.
'---------------------------------------------------------------------------------------

    Dim lDriveType As String
    
    sDriveLetter = Trim(sDriveLetter)
    
    If Len(sDriveLetter) <> 1 Then
        MsgBox "Specify only Drive letter and nothing else. For example, if you want to get the Drive Type for C Drive then pass only character C to this function.", vbInformation, "Do it in this way please ..."
        GetDriveTypeStr = ""
        Exit Function
    End If

    lDriveType = GetDriveType(sDriveLetter & ":\")
    
    Select Case lDriveType
        
        Case 2: GetDriveTypeStr = "Removable"
        
        Case 3: GetDriveTypeStr = "Fixed"
        
        Case 4: GetDriveTypeStr = "Remote"
        
        Case 5: GetDriveTypeStr = "CD-Rom"
        
        Case 6: GetDriveTypeStr = "RAM-Drive"
        
        Case Else: GetDriveTypeStr = "Unknown"
    
    End Select
    
    GetDriveTypeEx = lDriveType
    
End Function





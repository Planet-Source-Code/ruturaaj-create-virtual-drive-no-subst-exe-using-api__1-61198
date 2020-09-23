VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Virtual Drive (no subst.exe ... pure API way !)"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnBrowseFolder 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6840
      TabIndex        =   8
      Top             =   2040
      Width           =   375
   End
   Begin VB.Frame fmControls 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   4
      Top             =   3120
      Width           =   4215
      Begin VB.CommandButton btnUnmount 
         Caption         =   "&Unmount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton btnMount 
         Caption         =   "&Mount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fmDriveType 
      Caption         =   "Drive Type ... "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   4215
      Begin VB.Label lblDrvType 
         BackStyle       =   0  'Transparent
         Caption         =   "Drive Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame fmFolder 
      Caption         =   "Set Mount Path ... "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3120
      TabIndex        =   2
      Top             =   1680
      Width           =   4215
      Begin VB.TextBox txtFolderPath 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   7
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "... you can also Drag-n-Drop folder you want to mount in text-box above ..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   3975
      End
   End
   Begin VB.Frame fmDriveSelect 
      Caption         =   "Select Drive Letter ... "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ListBox lstDrives 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3345
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Label lblInfo_about 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmMain.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1815
      Left            =   240
      TabIndex        =   10
      Top             =   4200
      Width           =   6975
   End
   Begin VB.Shape shpAbout 
      BackStyle       =   1  'Opaque
      BorderStyle     =   2  'Dash
      Height          =   2055
      Left            =   120
      Top             =   4080
      Width           =   7215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnBrowseFolder_Click()

'---------------------------------------------------------------------------------------
' Author     : Ruturaj
' Website    :
' Email      : mailme_friends@yahoo.com
'=======================================================================================
' Procedure  : btnBrowseFolder_Click
' Type       : Sub
' ReturnType :
'=======================================================================================
' Purpose    : Show BrowseForFolder dialog.
''=======================================================================================
' Steps      : [1] Show Dialog
'              [2] Check if return value is an Empty String.
'---------------------------------------------------------------------------------------

    
    Dim sFolderPath As String
    
    sFolderPath = mBrowseForFolder.BrowseForFolder(Me, "Select Folder to Mount ...", ROOTDIR_COMMON_DESKTOP, , , True, False, "Done !")
    
    If Len(Trim(sFolderPath)) > 0 Then
        txtFolderPath.Text = sFolderPath
    End If
End Sub

Private Sub btnMount_Click()

'---------------------------------------------------------------------------------------
' Author     : Ruturaj
' Website    :
' Email      : mailme_friends@yahoo.com
'=======================================================================================
' Procedure  : btnMount_Click
' Type       : Sub
' ReturnType :
'=======================================================================================
' Purpose    : Mount specified Folder as Virtual Drive.
''=======================================================================================
' Steps      : [1] Check whether any Folder has selected or not.
'              [2] Attempt to Mount selected folder as selected Drive letter.
'---------------------------------------------------------------------------------------

    If Len(Trim(txtFolderPath.Text)) = 0 Then
        MsgBox "Please specify a valid Folder Path to mount as Virtual Drive.", vbInformation, "Do it this way please ..."
        btnBrowseFolder.SetFocus
        Exit Sub
    Else
        If mMountDrive.MountVD(lstDrives.List(lstDrives.ListIndex), txtFolderPath.Text) Then
            MsgBox txtFolderPath.Text & " has been successfully mounted as " & lstDrives.List(lstDrives.ListIndex) & " Drive on this machine !" & vbCrLf & vbCrLf & "However, it's worth to note that, this is temporary and will get vanished on next System Shut-Down or Restart. The only work-around for this is ... simply add your tool to system start-up and then it will create your Virtual Drive on every System-boot. Of course, you need to save the settings like Drive Letter and Folder Path to somewhere in Registry so that same Virtual Mount will be created.", vbInformation, "Done ! What's Next ?"
        Else
            MsgBox "There was an Error while mounting " & txtFolderPath.Text & " as " & lstDrives.List(lstDrives.ListIndex) & " Drive on this machine. This may be due to the insufficient previlege to perform this action (if you are on Network PC) or the required API is not supported by your OS Version (very rare case though !)", vbInformation, "Oops ! Error ... Operation failed."
        End If
    End If
End Sub

Private Sub btnUnmount_Click()

'---------------------------------------------------------------------------------------
' Author     : Ruturaj
' Website    :
' Email      : mailme_friends@yahoo.com
'=======================================================================================
' Procedure  : btnUnmount_Click
' Type       : Sub
' ReturnType :
'=======================================================================================
' Purpose    : Unmount specified Drive Letter.
''=======================================================================================
' Steps      : [1] Check if there is any Folder Path selected.
'              [2] Attempt to Unmount selected Drive Letter.
'---------------------------------------------------------------------------------------

    If Len(Trim(txtFolderPath.Text)) = 0 Then
        MsgBox "Please specify a valid Folder Path to mount as Virtual Drive.", vbInformation, "Do it this way please ..."
        btnBrowseFolder.SetFocus
        Exit Sub
    Else
        If mMountDrive.MountVD(lstDrives.List(lstDrives.ListIndex), txtFolderPath.Text, True) Then
            MsgBox txtFolderPath.Text & " has been successfully Unmounted.", vbInformation, "Done ! What's Next ?"
        Else
            MsgBox "There was an Error while mounting " & txtFolderPath.Text & " as " & lstDrives.List(lstDrives.ListIndex) & " Drive on this machine. This may be due to the insufficient previlege to perform this action (if you are on Network PC) or the required API is not supported by your OS Version (very rare case though !)", vbInformation, "Oops ! Error ... Operation failed."
        End If
    End If

End Sub

Private Sub Form_Load()

'---------------------------------------------------------------------------------------
' Author     : Ruturaj
' Website    :
' Email      : mailme_friends@yahoo.com
'=======================================================================================
' Procedure  : Form_Load
' Type       : Sub
' ReturnType :
'=======================================================================================
' Purpose    : Show frmMain.
''=======================================================================================
' Steps      : [1] Fill the lstDrives with A to Z.
'---------------------------------------------------------------------------------------

    Dim iCnt As Integer
    
    For iCnt = 65 To 90
        lstDrives.AddItem Chr(iCnt)
    Next iCnt
    
    mMountDrive.SetAlwaysOnTopMode Me.hWnd, True
End Sub

Private Sub lstDrives_Click()

'---------------------------------------------------------------------------------------
' Author     : Ruturaj
' Website    :
' Email      : mailme_friends@yahoo.com
'=======================================================================================
' Procedure  : lstDrives_Click
' Type       : Sub
' ReturnType :
'=======================================================================================
' Purpose    : Validation.
''=======================================================================================
' Steps      : [1] Get drive type for selected Drive letter in List.
'              [2] Enable/Disable button status according to the correct Drive type.
'---------------------------------------------------------------------------------------

    Dim sDriveType As String
    Dim lDriveType As Long
    
    lDriveType = mMountDrive.GetDriveTypeEx(lstDrives.List(lstDrives.ListIndex), sDriveType)
    
    If sDriveType = "Unknown" Then
        lblDrvType.Caption = "The Drive letter which you have selected for your Virtual Drive mount is " & sDriveType & ". So, you can proceed with this Drive Letter"
        btnMount.Enabled = True
        btnUnmount.Enabled = True
        btnBrowseFolder.Enabled = True
        txtFolderPath.Enabled = True
    Else
        lblDrvType.Caption = "The Drive letter which you have selected for your Virtual Drive is actually registered as " & sDriveType & " on this System. So, you can't use it for your Virtual Drive mount. Try some other letter ..."
        btnMount.Enabled = False
'        btnUnmount.Enabled = False
        btnBrowseFolder.Enabled = False
        txtFolderPath.Enabled = False
    End If
    
End Sub

Private Sub txtFolderPath_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

'---------------------------------------------------------------------------------------
' Author     : Ruturaj
' Website    :
' Email      : mailme_friends@yahoo.com
'=======================================================================================
' Procedure  : txtFolderPath_OLEDragDrop
' Type       : Sub
' ReturnType :
'=======================================================================================
' Purpose    : Support to Drag-n-Drop.
''=======================================================================================
' Steps      : [1] Check if attempted Drag item is Folder or File.
'              [2] If Folder then txtFolderPath.Text = Selected Folder Path.
'---------------------------------------------------------------------------------------

    If IsFolder(Data.Files(1)) = True Then
        txtFolderPath.Text = Data.Files(1)
    Else
        MsgBox "Please Drag-n-Drop only Folder ..."
    End If
End Sub

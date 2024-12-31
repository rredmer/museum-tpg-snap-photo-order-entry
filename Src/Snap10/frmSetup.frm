VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetup 
   Caption         =   "Setup"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   15240
   Begin VB.CommandButton SetupButtons 
      Caption         =   "Explorer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   8
      Left            =   11370
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   10080
      Width           =   1875
   End
   Begin VB.CommandButton SetupButtons 
      Caption         =   "Get License"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   7
      Left            =   13290
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   10080
      Width           =   1935
   End
   Begin VB.CommandButton SetupButtons 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   6
      Left            =   10140
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   10080
      Width           =   1185
   End
   Begin VB.CommandButton SetupButtons 
      Caption         =   "Debug"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   5
      Left            =   9030
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   10080
      Width           =   1065
   End
   Begin VB.CommandButton SetupButtons 
      Caption         =   "Erase Folder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   4
      Left            =   7140
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   10080
      Width           =   1845
   End
   Begin VB.CommandButton SetupButtons 
      Caption         =   "Copy To Flash"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   3
      Left            =   4950
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   10080
      Width           =   2145
   End
   Begin VB.CommandButton SetupButtons 
      Caption         =   "End Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   2
      Left            =   2970
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   10080
      Width           =   1935
   End
   Begin VB.CommandButton SetupButtons 
      Caption         =   "Exit && Logout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   1
      Left            =   1050
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   10080
      Width           =   1875
   End
   Begin VB.CommandButton SetupButtons 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   0
      Left            =   60
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   10080
      Width           =   945
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14310
      Top             =   10440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SNAP10.UserControl1 MainDrive 
      Height          =   2235
      Left            =   9810
      TabIndex        =   0
      Top             =   0
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   3942
   End
   Begin FPSpread.vaSpread FolderSpread 
      Height          =   10005
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   9765
      _Version        =   393216
      _ExtentX        =   17224
      _ExtentY        =   17648
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      MaxCols         =   3
      MaxRows         =   1
      MoveActiveOnFocus=   0   'False
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   2
      SpreadDesigner  =   "frmSetup.frx":0000
      ClipboardOptions=   0
      CellNoteIndicator=   3
   End
   Begin SNAP10.UserControl1 UsbDrive 
      Height          =   2235
      Left            =   9810
      TabIndex        =   2
      Top             =   2250
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   3942
   End
   Begin FPSpread.vaSpread SetupSpread 
      Height          =   5505
      Left            =   9810
      TabIndex        =   3
      Top             =   4530
      Width           =   5445
      _Version        =   393216
      _ExtentX        =   9604
      _ExtentY        =   9710
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      MaxCols         =   2
      MaxRows         =   12
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmSetup.frx":0379
      UserResize      =   1
      ClipboardOptions=   0
      CellNoteIndicator=   3
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   600
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   1058
      ButtonWidth     =   2064
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit && Logout"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy To Flash"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Erase"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "End Program"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View Log File"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Get License"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Update"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   600
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   1058
      ButtonWidth     =   2064
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit && Logout"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy To Flash"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Erase"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "End Program"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View Log File"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Get License"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Update"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IsFolderFileOpen As Boolean              'Set TRUE when folder file is open
Public IsUsbReady As Boolean
Public FoldersValidated As Boolean
Public BaseFolder As String                     'Base data folder
Public ImageFolder As String                    'Images folder
Public UsbFoldersValidated As Boolean

Private Const UsbMinFreeSpace As Long = 6000000 'Minimum amount of USB Disk free space allowed
Private Const FreeSpaceThreshold As Long = 1
Private Const SettingsFolder As String = "Settings\"  'Settings data folder
Public Enum SettingTypes
    AutoUSB = 1
    RepeatSubjectSetting = 2
    AlphaPackage = 3
    ClearOnSpacebar = 4
    ShowTotals = 5
    SingleDigits = 6
    RotateImageSetting = 7
    UsbFlashDrive = 8
    ImageTimeout = 9
    MaxNoScans = 10
    FreeGigabytes = 11
    AdminPWD = 12
End Enum

Private Sub Form_Load()
    '---- Initialize form controls
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    BaseFolder = "SNAP10\"
    ImageFolder = "Images\"
    IsFolderFileOpen = False                    'Initialize folder file to not open
    IsUsbReady = False
    FoldersValidated = False
    UsbFoldersValidated = False
    Set MainDrive.SrcDrive = FS.GetDrive("C:")
    ReadSpreadsheet Me.SetupSpread, "Setup.txt", False
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":Form_Load " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

Private Sub SetupButtons_Click(Index As Integer)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Select Case Index
        Case 0      'Exit
            WriteSpreadsheet Me.SetupSpread, "Setup.txt", False
            Me.Hide
        Case 1      'Exit & Logout
            frmPassword.IsPasswordValid = False
            frmPassword.txtPassword.Text = ""
            WriteSpreadsheet Me.SetupSpread, "Setup.txt", False
            Me.Hide
        Case 2      'End
            If MsgBox("Are you sure?", vbApplicationModal + vbYesNo + vbDefaultButton2 + vbQuestion, "End Program") = vbYes Then
                End
            End If
        Case 3      'Copy
            CopyFolder
        Case 4      'Erase
            EraseFolder
        Case 5      'View Log
            frmLogFile.Show vbModal
            
        Case 6      'Update
            Dim UpdatePath As String
            If IsUsbReady Then
                UpdatePath = Me.UsbDrive.SrcDrive.Path & "\Update\SNAP09.EXE"
                '---- Check for program update
                If CheckForUpdate(UpdatePath) Then
                    If MsgBox("Install program update?", vbYesNo + vbApplicationModal + vbQuestion, "Update Found") = vbYes Then
                        Dim tmpstr As String
                        tmpstr = MainDrive.SrcDrive.Path & "\" & BaseFolder & SettingsFolder & "Setup.txt"
                        If FS.FileExists(tmpstr) Then
                            FS.DeleteFile tmpstr
                        End If
                        DoUpdate UpdatePath
                    End If
                Else
                    MsgBox "No update found.", vbApplicationModal + vbInformation + vbOKOnly, "Update"
                End If
            End If
        Case 7      'Get License
            With Me.CommonDialog1
                .CancelError = False
                .DefaultExt = "lic"
                .DialogTitle = "Select License File"
                .InitDir = CStr(GetSetting(UsbFlashDrive))
                .ShowOpen
                If .FileName <> "" Then
                    FS.CopyFile .FileName, App.Path & "\SNAP08.DAT.lic", True
                    IsLic
                End If
            End With
        Case 8      'Explorer
            If FS.FileExists("c:\windows\explorer.exe") Then
                Shell "c:\windows\explorer.exe", vbNormalFocus
            Else
                MsgBox "Windows Explorer not found.", vbApplicationModal + vbOKOnly + vbInformation, "Error"
            End If
    End Select
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":SetupButtons_Click " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

'Copy the selected folder to a flash drive
Private Sub CopyFolder()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim FolderName As String, DestName As String, SourceFolder As Folder, BytesLeft As Long, mb As Long
    If IsUsbReady = False Then
        MsgBox "Please connect USB Flash Drive."
    Else
        mb = 1048576
        With Me.FolderSpread
            .Col = 1
            .Row = .ActiveRow
            FolderName = Trim(.Text)
        End With
        Set SourceFolder = FS.GetFolder(MainDrive.SrcDrive.Path & "\" & BaseFolder & FolderName)    ' & "\" & machineid & "_" & ImageFolder)
        DestName = UsbDrive.SrcDrive.Path & "\" & Mid(SourceFolder, 4)
        If MsgBox("Copy " & FolderName & "?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Are you sure?") = vbYes Then
            '--- Get size of source folder
            BytesLeft = UsbDrive.SrcDrive.AvailableSpace - SourceFolder.Size
            If (BytesLeft > FreeSpaceThreshold) Then
                FS.CopyFolder SourceFolder, DestName, True
                MsgBox "Copy completed."
            Else
                MsgBox "There is not enough space on the destination drive. An additional " & Format(Abs(BytesLeft) / mb, "#,##0") & "MB is needed."
            End If
        End If
        Set SourceFolder = Nothing
    End If
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":CopyFolder " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    MsgBox "Error in copy.", vbApplicationModal + vbCritical + vbOKOnly, "Copy folder"
End Sub

Private Sub EraseFolder()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim FolderName As String
    With Me.FolderSpread
        .Col = 1
        .Row = .ActiveRow
        FolderName = .Text
    End With
    If MsgBox("Erase " & MainDrive.SrcDrive.Path & BaseFolder & FolderName & "?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Are you sure?") = vbYes Then
        FS.DeleteFolder MainDrive.SrcDrive.Path & "\" & BaseFolder & FolderName, True
        MsgBox "Erase completed."
    End If
    FillGrid
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":EraseFolder " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    MsgBox "Error in erase.", vbApplicationModal + vbCritical + vbOKOnly, "Erase folder"
End Sub

Public Sub FillGrid()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim FolderName As String, SourceFolder, fld, imgfld As Folder
    FolderName = "C:\" & BaseFolder
    Set SourceFolder = FS.GetFolder(FolderName)
    With Me.FolderSpread
        .MaxRows = 0
        For Each fld In SourceFolder.SubFolders
            If fld.Name <> "Settings" Then
                .MaxRows = .MaxRows + 1
                .SetText 1, .MaxRows, fld.Name
                .SetText 2, .MaxRows, fld.DateCreated
                If FS.FolderExists(fld.Path & "\" & Trim(frmStartDay.GetFolderName) & "_" & ImageFolder) = True Then
                    Set imgfld = fld.SubFolders(Trim(frmStartDay.GetFolderName) & "_Images")
                    .SetText 3, .MaxRows, imgfld.Files.Count
                Else
                    .SetText 3, .MaxRows, 0
                End If
            End If
        Next
    End With
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":FillGrid " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

Public Sub UpdateLogFile(spr As vaSpread, FileName As String, JobSpread As Boolean)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim PathName As String, RootPath As String
    RootPath = BaseFolder & SettingsFolder
    PathName = MainDrive.SrcDrive.Path & "\" & RootPath & FileName
    If FS.FileExists(PathName) = False Then
        spr.ExportToTextFile PathName, "", ",", vbCrLf, ExportToTextFileCreateNewFile + ExportToTextFileColHeaders + ExportToTextFileRowHeaders, ""
    Else
        '---- Roll over log file at 1mb
        If FS.GetFile(PathName).Size > 1048576 Then
            FS.DeleteFile PathName, True
            frmLogFile.LogSpread.MaxRows = 1
        End If
        spr.ExportRangeToTextFile 1, spr.MaxRows, spr.MaxCols, spr.MaxRows, PathName, "", ",", vbCrLf, ExportRangeToTextFileAppendToExistingFile + ExportToTextFileRowHeaders + ExportToTextFileUnformattedData, ""
    End If
    If IsUsbReady = True Then
        PathName = UsbDrive.SrcDrive.Path & "\" & RootPath & FileName
        If FS.FileExists(PathName) = False Then
            spr.ExportToTextFile PathName, "", ",", vbCrLf, ExportToTextFileCreateNewFile + ExportToTextFileColHeaders + ExportToTextFileRowHeaders, ""
        Else
            spr.ExportRangeToTextFile 1, spr.MaxRows, spr.MaxCols, spr.MaxRows, PathName, "", ",", vbCrLf, ExportRangeToTextFileAppendToExistingFile + ExportToTextFileRowHeaders + ExportToTextFileUnformattedData, ""
        End If
    End If
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume Next
End Sub

Public Sub UpdateSpreadsheet(spr As vaSpread, FileName As String, JobSpread As Boolean, EditMode As Boolean)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim PathName As String
    Dim RootPath As String
    RootPath = BaseFolder & IIf(JobSpread, Trim(frmStartDay.GetFolderName) & "\", "") & SettingsFolder
    PathName = MainDrive.SrcDrive.Path & "\" & RootPath & IIf(JobSpread, MachineId & "_", "") & FileName
    If FS.FileExists(PathName) = False Or EditMode Then
        spr.ExportToTextFile PathName, "", ",", vbCrLf, ExportToTextFileCreateNewFile + ExportToTextFileColHeaders + ExportToTextFileRowHeaders, ""
        frmLogFile.LogText InfoMsg, "UpdateSpreadsheet: Created " & spr.Name & " To " & PathName & ", " & spr.MaxRows & " records."
    Else
        spr.ExportRangeToTextFile 1, spr.MaxRows, spr.MaxCols, spr.MaxRows, PathName, "", ",", vbCrLf, ExportRangeToTextFileAppendToExistingFile + ExportToTextFileRowHeaders + ExportToTextFileUnformattedData, ""
        If spr.Name <> "LogSpread" Then frmLogFile.LogText InfoMsg, "UpdateSpreadsheet: Updated " & spr.Name & " To " & PathName & ", " & spr.MaxRows & " records."
    End If
    If IsUsbReady = True Then
        PathName = UsbDrive.SrcDrive.Path & "\" & RootPath & IIf(JobSpread, MachineId & "_", "") & FileName
        If FS.FileExists(PathName) = False Or EditMode Then
            spr.ExportToTextFile PathName, "", ",", vbCrLf, ExportToTextFileCreateNewFile + ExportToTextFileColHeaders + ExportToTextFileRowHeaders, ""
            If spr.Name <> "LogSpread" Then frmLogFile.LogText InfoMsg, "UpdateSpreadsheet: Created " & spr.Name & " To " & PathName & ", " & spr.MaxRows & " records."
        Else
            spr.ExportRangeToTextFile 1, spr.MaxRows, spr.MaxCols, spr.MaxRows, PathName, "", ",", vbCrLf, ExportRangeToTextFileAppendToExistingFile + ExportToTextFileRowHeaders + ExportToTextFileUnformattedData, ""
            If spr.Name <> "LogSpread" Then frmLogFile.LogText InfoMsg, "UpdateSpreadsheet: Updated " & spr.Name & " To " & PathName & ", " & spr.MaxRows & " records."
        End If
    End If
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":UpdateSpreadsheet " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

Public Sub WriteSpreadsheet(spr As vaSpread, FileName As String, JobSpread As Boolean)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim PathName As String
    Dim RootPath As String
    RootPath = BaseFolder & IIf(JobSpread, Trim(frmStartDay.GetFolderName) & "\", "") & SettingsFolder
    PathName = MainDrive.SrcDrive.Path & "\" & RootPath & IIf(JobSpread, MachineId & "_", "") & FileName
    spr.ExportRangeToTextFile 1, 1, spr.MaxCols, spr.MaxRows, PathName, "", ",", vbCrLf, ExportToTextFileCreateNewFile + ExportToTextFileColHeaders + ExportToTextFileRowHeaders + ExportToTextFileUnformattedData, ""
    frmLogFile.LogText InfoMsg, "WriteSpreadsheet: Wrote " & spr.Name & " To " & PathName & ", " & spr.MaxRows & " records."
    If IsUsbReady Then
        PathName = UsbDrive.SrcDrive.Path & "\" & RootPath & IIf(JobSpread, MachineId & "_", "") & FileName
        spr.ExportRangeToTextFile 1, 1, spr.MaxCols, spr.MaxRows, PathName, "", ",", vbCrLf, ExportToTextFileCreateNewFile + ExportToTextFileColHeaders + ExportToTextFileRowHeaders + ExportToTextFileUnformattedData, ""
        frmLogFile.LogText InfoMsg, "WriteSpreadsheet: Wrote " & spr.Name & " To " & PathName & ", " & spr.MaxRows & " records."
    End If
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":WriteSpreadsheet " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

Public Function ReadSpreadsheet(spr As vaSpread, FileName As String, JobSpread As Boolean) As Boolean
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    ReadSpreadsheet = False
    Dim PathName As String
    Dim RootPath As String
    RootPath = BaseFolder & IIf(JobSpread, Trim(frmStartDay.GetFolderName) & "\", "") & SettingsFolder
    PathName = MainDrive.SrcDrive.Path & "\" & RootPath & IIf(JobSpread, MachineId & "_", "") & FileName
    If FS.FileExists(PathName) = True Then
        spr.LoadTextFile PathName, "", ",", vbCrLf, LoadTextFileClearDataOnly + LoadTextFileColHeaders + LoadTextFileRowHeaders, ""
        frmLogFile.LogText InfoMsg, "ReadSpreadsheet: Read " & spr.Name & " From " & PathName & ", " & spr.MaxRows & " records."
        ReadSpreadsheet = True
    Else
        frmLogFile.LogText InfoMsg, "ReadSpreadsheet: Read " & spr.Name & " From " & PathName & " NOT FOUND."
    End If
    Exit Function
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":ReadSpreadsheet " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Function

Private Function CreateFolder(TargetDrive As Drive, FolderName As String) As Boolean
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    CreateFolder = False
    If FS.FolderExists(TargetDrive.Path & "\" & FolderName) = False Then
        FS.CreateFolder TargetDrive.Path & "\" & FolderName
        frmLogFile.LogText InfoMsg, "CreateFolder: " & TargetDrive.Path & "\" & FolderName
    Else
        'frmLogFile.LogText InfoMsg, "CreateFolder: " & TargetDrive.Path & "\" & FolderName & " EXISTS."
    End If
    CreateFolder = True
    Exit Function
ErrorHandler:
    Exit Function
End Function

Public Function GetSetting(SettingNum As SettingTypes) As Variant
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim vValue As Variant
    Me.SetupSpread.GetText 2, SettingNum, vValue
    GetSetting = vValue
    Exit Function
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":GetSetting " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Function

Public Function IsLic() As Boolean
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    'IsLic = False
    'If ReadLicFile() = True Then
        Me.SetupButtons(7).Visible = False
        Me.SetupButtons(7).Enabled = False
        IsLic = True
    'Else
    '    Me.SetupButtons(7).Visible = True
    '    Me.SetupButtons(7).Enabled = True
    '    If IsUsbReady Then
    '        FS.CopyFile RegFile, Trim(Me.UsbDrive.SrcDrive.Path) & "\"
    '    End If
    'End If
    Exit Function
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":IsLic " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Function

Public Sub DeleteOldestJob()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim FolderName As String
    With Me.FolderSpread
        If .MaxRows > 0 Then
            .Row = 1
            .Col = 1
            FolderName = MainDrive.SrcDrive & "\" & BaseFolder & Trim(.Text)
            FS.DeleteFolder FolderName, True
            FillGrid
        End If
    End With
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":DeleteOldestJob " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

Public Sub ManageDriveSpace()
    If FolderSpread.MaxRows > 1 Then             'There has to be more than 1 job to consider throwing one out!
        Dim FreeSpace As Double
        FreeSpace = CDbl(GetSetting(FreeGigabytes))
        If (MainDrive.SrcDrive.AvailableSpace / 1048576) < ((1048576 * FreeSpace) / 1024) Then
            DeleteOldestJob
        End If
    End If
End Sub

Public Sub CheckDrive(UsbMode As Boolean)
    
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim drv As String
    
    If FoldersValidated = False Then
        CreateFolder MainDrive.SrcDrive, BaseFolder
        CreateFolder MainDrive.SrcDrive, BaseFolder & SettingsFolder
        CreateFolder MainDrive.SrcDrive, BaseFolder & Trim(frmStartDay.GetFolderName)
        CreateFolder MainDrive.SrcDrive, BaseFolder & Trim(frmStartDay.GetFolderName) & "\" & SettingsFolder
        CreateFolder MainDrive.SrcDrive, BaseFolder & Trim(frmStartDay.GetFolderName) & "\" & Trim(frmStartDay.GetFolderName) & "_" & ImageFolder
        FoldersValidated = True
    End If
    
    If UsbMode Then
        drv = GetSetting(UsbFlashDrive)
        If FS.DriveExists(drv) Then
            If IsUsbReady = False Then
                If FS.GetDrive(drv).DriveType = CDRom Then
                    IsUsbReady = False
                Else
                    Set UsbDrive.SrcDrive = FS.GetDrive(drv)
                    UsbDrive.Reload
                    UsbDrive.Visible = True
                    IsUsbReady = True
                    UsbFoldersValidated = False
                End If
            Else
                If UsbDrive.SrcDrive.AvailableSpace < 5000000 Then
                    IsUsbReady = False
                Else
                    If UsbFoldersValidated = False Then
                        CreateFolder UsbDrive.SrcDrive, BaseFolder
                        CreateFolder UsbDrive.SrcDrive, BaseFolder & SettingsFolder
                        CreateFolder UsbDrive.SrcDrive, BaseFolder & Trim(frmStartDay.GetFolderName)
                        CreateFolder UsbDrive.SrcDrive, BaseFolder & Trim(frmStartDay.GetFolderName) & "\" & SettingsFolder
                        CreateFolder UsbDrive.SrcDrive, BaseFolder & Trim(frmStartDay.GetFolderName) & "\" & Trim(frmStartDay.GetFolderName) & "_" & ImageFolder
                        UsbFoldersValidated = True
                    End If
                End If
            End If
        Else
            UsbDrive.Visible = False
            IsUsbReady = False
            UsbFoldersValidated = False
        End If
    Else
        UsbDrive.Visible = False
        IsUsbReady = False
    End If
    MainDrive.Reload
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":CheckDrive " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub



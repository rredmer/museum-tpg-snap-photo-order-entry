VERSION 5.00
Begin VB.UserControl ExplorerControl 
   ClientHeight    =   9450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   ScaleHeight     =   9450
   ScaleWidth      =   4950
   Begin VB.Frame ExploreDriveFrame 
      Caption         =   "Drive"
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4875
      Begin VB.TextBox NumberOfFiles 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   8790
         Width           =   2385
      End
      Begin VB.FileListBox FileListBox 
         Height          =   4380
         Left            =   90
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   4200
         Width           =   4695
      End
      Begin VB.DirListBox FolderListBox 
         Height          =   3915
         Left            =   90
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Number of files:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   8880
         Width           =   2115
      End
   End
End
Attribute VB_Name = "ExplorerControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Initialize()
    'Init form controls
End Sub

Private Sub FolderListBox_Change()
    On Error GoTo ErrorHandler
    FileListBox.Path = FolderListBox.Path
    NumberOfFiles.Text = Format(FileListBox.ListCount, "#,###")
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume Next
End Sub

Public Sub SetDrive(pathspec As String, FrameCaption As String)
    FolderListBox.Path = pathspec
    ExploreDriveFrame.Caption = FrameCaption
End Sub

Public Sub RefreshList()
    FolderListBox.Refresh
    FileListBox.Refresh
    NumberOfFiles.Text = Format(FileListBox.ListCount, "#,###")
End Sub

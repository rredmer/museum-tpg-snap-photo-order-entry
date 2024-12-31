VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSerializer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serializer"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12120
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid LicenseGrid 
      Bindings        =   "frmSerializer.frx":0000
      Height          =   6135
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10821
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdButtons 
      Caption         =   "Generate"
      Height          =   735
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton CmdButtons 
      Caption         =   "Load"
      Height          =   735
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton CmdButtons 
      Caption         =   "Exit"
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Width           =   1335
   End
End
Attribute VB_Name = "frmSerializer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FS As New Scripting.FileSystemObject
Private Const ReqKey As String = "!!@%!(^&"         '11251967
Private Const LicKey As String = "!@$*#&))#^)"      '12483700360
Private cnn As New ADODB.Connection
Private rs As New ADODB.Recordset

Private Sub Form_Load()
    cnn.CursorLocation = adUseClient
    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Dev\Serializer\Licenses.mdb;Persist Security Info=False"
    rs.Open "SELECT * FROM Serials;", cnn, adOpenDynamic, adLockOptimistic
    Set Me.LicenseGrid.DataSource = rs
End Sub


Private Sub CmdButtons_Click(Index As Integer)
    Select Case Index
        Case 0:
            End
        Case 1:
            LoadFile
        Case 2:
           'GenerateFile
    End Select
End Sub


Public Sub LoadFile()
    Dim FileName As String, SrcData As String, vData As Variant, idx As Integer, SerNum As String
    FileName = GetFileName("Select Request File", True)
    SrcData = ReadRequestFile(FileName)
    If Len(SrcData) > 0 Then
        vData = Split(SrcData, ",")
        With rs
            .AddNew
            For idx = 0 To UBound(vData)
                .Fields(idx).value = vData(idx)
            Next
            .Update
            .UpdateBatch adAffectAllChapters
            .Requery
            DoEvents
            SerNum = .Fields("SerialNum").value
        End With
        
        SrcData = SerNum & "," & SrcData & "," & SerNum
        WriteLicenseFile Left(FileName, InStrRev(FileName, "\")) & Mid(vData(5), 13, 2) & Right(vData(5), 2) & "_" & Mid(FileName, InStrRev(FileName, "\") + 1) & ".lic", SrcData
    End If
End Sub

Public Function GetFileName(WinTitle As String, OpenFile As Boolean) As String
    With Me.CommonDialog1
        .DialogTitle = WinTitle
        .DefaultExt = "dat"
        .CancelError = False
        If OpenFile Then
            .ShowOpen
        Else
            .ShowSave
        End If
        GetFileName = .FileName
    End With
End Function

Public Function ReadRequestFile(fname As String) As String
    Dim ts As TextStream, LockStr As String
    Set ts = FS.OpenTextFile(fname, ForReading, False, TristateFalse)
    LockStr = ts.ReadAll
    ts.Close
    Set ts = Nothing
    Decipher ReqKey, LockStr, ReadRequestFile
End Function


Public Function WriteLicenseFile(fname As String, UnLockStr As String) As Boolean
    Dim ts As TextStream, LockStr As String
    Set ts = FS.CreateTextFile(fname, True, False)
    Cipher LicKey, UnLockStr, LockStr
    ts.Write LockStr
    ts.Close
    Set ts = Nothing
End Function

VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.UserControl UserControl1 
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   ScaleHeight     =   2385
   ScaleWidth      =   5535
   Begin VB.Frame MainDriveFrame 
      Caption         =   "Drive Frame"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin MSChart20Lib.MSChart MSChart 
         Height          =   1935
         Left            =   2640
         OleObjectBlob   =   "DriveControl.ctx":0000
         TabIndex        =   2
         Top             =   120
         Width           =   2655
      End
      Begin VB.TextBox DriveText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const gb As Long = 2 ^ 30
Public SrcDrive As Drive

Public Sub Reload()
    On Error Resume Next
    With SrcDrive
        MainDriveFrame.Caption = "Drive " & .DriveLetter
        DriveText.Text = _
        "Type=" & GetTypeDescription(.DriveType) & vbCrLf & _
        "Volume Name=" & IIf(Trim(.VolumeName) = "", "N/A", .VolumeName) & vbCrLf & _
        "Drive Letter=" & .DriveLetter & vbCrLf & _
        "S/N=" & .SerialNumber & vbCrLf & _
        "Total Size=" & Format(.TotalSize / gb, "#,##0.00") & "GB" & vbCrLf & _
        "Available=" & Format(.AvailableSpace / gb, "#,##0.00") & "GB" & vbCrLf & _
        "Free=" & Format(.FreeSpace / gb, "#,##0.00") & "GB"
    End With
    DriveText.Refresh
    Dim i As Integer
    With MSChart
        .chartType = VtChChartType2dPie
        .Row = 1
        .Column = 1
        .Data = SrcDrive.TotalSize - SrcDrive.AvailableSpace     'Used = Red
        .Column = 2
        .Data = SrcDrive.AvailableSpace                  'Available = grn
        .Column = 3
        .Data = 0
        With .DataGrid
            .RowLabelCount = 1
            .ColumnCount = 3
            .RowCount = 1
            .ColumnLabel(1, 1) = "Used"
            .ColumnLabel(2, 1) = "Free"
            .ColumnLabel(3, 1) = "N/A"
            .RowLabel(1, 1) = "Drive " & SrcDrive.DriveLetter
        End With
        '--- Percentage Labels
        'For i = 1 To .Plot.SeriesCollection.Count
        '    With .Plot.SeriesCollection(i).DataPoints(-1).DataPointLabel
        '        .LocationType = VtChLabelLocationTypeOutside
        '        .Component = VtChLabelComponentPercent
        '        .PercentFormat = "0%"
        '        .VtFont.Size = 10
        '    End With
        'Next i
        .Refresh
    End With
    
End Sub

Private Function GetTypeDescription(TypeNum As Integer) As String
    Dim typ As String
    typ = "Unknown"
    Select Case TypeNum
        Case 0: typ = "Unknown"
        Case 1: typ = "Removable"
        Case 2: typ = "Fixed"
        Case 3: typ = "Network"
        Case 4: typ = "CD/DVD"
        Case 5: typ = "RAM Disk"
    End Select
    GetTypeDescription = typ
End Function

Public Function IsDriveReady() As Boolean
    IsDriveReady = SrcDrive.IsReady()
End Function


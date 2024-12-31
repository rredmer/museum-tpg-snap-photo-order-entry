VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.UserControl SummaryControl 
   ClientHeight    =   9435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10065
   ScaleHeight     =   9435
   ScaleWidth      =   10065
   Begin FPSpread.vaSpread TotalSpread 
      Height          =   4125
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5235
      Width           =   3795
      _Version        =   393216
      _ExtentX        =   6694
      _ExtentY        =   7276
      _StockProps     =   64
      AutoClipboard   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      MaxRows         =   99
      OperationMode   =   1
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "SummaryControl.ctx":0000
      ClipboardOptions=   0
      CellNoteIndicator=   3
   End
   Begin FPSpread.vaSpread SummarySpread 
      Height          =   5145
      Left            =   30
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   9975
      _Version        =   393216
      _ExtentX        =   17595
      _ExtentY        =   9075
      _StockProps     =   64
      AutoClipboard   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   37
      MaxRows         =   1
      MoveActiveOnFocus=   0   'False
      OperationMode   =   1
      SpreadDesigner  =   "SummaryControl.ctx":070C
      UserResize      =   0
      ClipboardOptions=   0
      CellNoteIndicator=   3
   End
   Begin VB.Label TotalLabel 
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3870
      TabIndex        =   3
      Top             =   9045
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label JobTotalLabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4500
      TabIndex        =   2
      Top             =   9015
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "SummaryControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private OriginalHeight As Integer
Public EditImageFile As String

Public Sub SaveOrderText(SubjectId As String, Pkgs As String, Prc As String, Cust As String, ImageId As String, ImageFileName As String, TookPicture As Boolean, EditMode As Boolean)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim RowNum As Long, SubjectVar As Variant
    'If DemoExceeded = True Then
    '    MsgBox "Demonstration Mode.  Data will not be saved.", vbApplicationModal + vbCritical + vbOKOnly, "WARNING!"
    'Else
        With SummarySpread
            If EditMode = False Then
                .MaxRows = .MaxRows + 1
                .SetText 1, .MaxRows, SubjectId
                .SetText 5, .MaxRows, ImageId
                .SetText 6, .MaxRows, ImageFileName
                .SetText 7, .MaxRows, IIf(TookPicture, "OK", "No Image")
                RowNum = .MaxRows
            Else
                For RowNum = .MaxRows To 1 Step -1
                    .GetText 1, RowNum, SubjectVar
                    If SubjectVar = SubjectId Then
                        .SetActiveCell 1, RowNum
                        .Row = RowNum
                        .Col = 1
                        Exit For
                    End If
                Next
            End If
            
            
            .SetText 2, RowNum, Pkgs
            .SetText 3, RowNum, Prc
            .SetText 4, RowNum, Cust
            
            'Set Text Columns from Edit Form
            .SetText 8, RowNum, frmEdit.txtCode.Text
            .SetText 9, RowNum, frmEdit.txtJobType.Text
            .SetText 10, RowNum, frmEdit.txtSequence.Text
            .SetText 11, RowNum, frmEdit.txtCriteria.Text
            .SetText 12, RowNum, frmEdit.txtFirstName.Text
            .SetText 13, RowNum, frmEdit.txtLastName.Text
            .SetText 14, RowNum, frmEdit.txtAddress.Text
            .SetText 15, RowNum, frmEdit.txtCity.Text
            .SetText 16, RowNum, frmEdit.txtState.Text
            .SetText 17, RowNum, frmEdit.txtZIP.Text
            .SetText 18, RowNum, frmEdit.txtMother.Text
            .SetText 19, RowNum, frmEdit.txtHomePhone.Text
            .SetText 20, RowNum, frmEdit.txtFather.Text
            .SetText 21, RowNum, frmEdit.txtWorkPhone.Text
            .SetText 22, RowNum, frmEdit.txtGender.Text
            .SetText 23, RowNum, frmEdit.txtBirthDate.Text
            .SetText 24, RowNum, frmEdit.txtTeacher.Text
            .SetText 25, RowNum, frmEdit.txtGrade.Text
            .SetText 26, RowNum, frmEdit.txtHomeRoom.Text
            .SetText 27, RowNum, frmEdit.txtTrack.Text
            .SetText 28, RowNum, frmEdit.txtYear.Text
            .SetText 29, RowNum, frmEdit.txtSchool.Text
            .SetText 30, RowNum, frmEdit.txtCustom1.Text
            .SetText 31, RowNum, frmEdit.txtCustom2.Text
            .SetText 32, RowNum, frmEdit.txtCustom3.Text
            .SetText 33, RowNum, frmEdit.txtCustom4.Text
            .SetText 34, RowNum, frmEdit.txtCustom5.Text
            .SetText 35, RowNum, frmEdit.txtCustom6.Text
            .SetText 36, RowNum, frmEdit.txtCustom7.Text
            .SetText 37, RowNum, frmEdit.txtCustom8.Text
            
            .SetActiveCell 1, .MaxRows
        End With
        frmSetup.UpdateSpreadsheet SummarySpread, Trim(frmStartDay.FolderName.Text) & "_PictureFile.txt", True, EditMode
    'End If
    'If DemoExceeded = False Then
    '    If DemoMode = True Then
    '        If SummarySpread.MaxRows >= MAX_DEMO_SHOTS Then
    '            DemoExceeded = True
    '        End If
    '    End If
    'End If
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "SaveOrderText " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

Public Sub FillTotals()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    TotalSpread.Enabled = False
    SummarySpread.Enabled = False
    Dim var As Variant, RowNum As Long, JobTotal As Double
    JobTotal = 0#
    TotalSpread.ClearRange 1, 1, 3, 99, True
    With SummarySpread
        For RowNum = 1 To .MaxRows
            .GetText 2, RowNum, var
            If Trim(var) <> "" Then
                UpdateTotal Trim(var)
            End If
        Next
    End With
    
    With TotalSpread
        For RowNum = 1 To 99
            .GetText 3, RowNum, var
            JobTotal = JobTotal + Val(var)
        Next
    End With
    JobTotalLabel.Caption = Format(JobTotal, "#0.00")
    frmSetup.WriteSpreadsheet TotalSpread, Trim(frmStartDay.FolderName.Text) & "_" & "CashReceipts", True
    TotalSpread.Enabled = True
    SummarySpread.Enabled = True
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "FillTotals " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

Private Sub UpdateTotal(Pkgs As String)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim Pkg As Variant, PkgCnt As Integer, oldQuan As Variant, varPrice As Variant, pkgNum As Long, Quan As Integer
    
    Pkg = Split(Pkgs, ";")
    For PkgCnt = 0 To UBound(Pkg) - 1
        If TotalSpread.RowHeadersAutoText = DispLetters Then        'REMOVED: If frmMain.AlphaMode Then
            pkgNum = Asc(Left(Pkg(PkgCnt), 1)) - Asc("A") + 1
            Quan = Mid(Pkg(PkgCnt), 3)
        Else
            pkgNum = Val(Left(Pkg(PkgCnt), 2))
            Quan = Mid(Pkg(PkgCnt), 4)
        End If
        frmPricing.PriceSpread.GetText 1, pkgNum, varPrice
        TotalSpread.GetText 1, pkgNum, oldQuan
        TotalSpread.SetText 1, pkgNum, Val(oldQuan) + Quan
        If CStr(varPrice) = "" Then
            'no price entered
        Else
            TotalSpread.SetText 2, pkgNum, Format(varPrice, "#.00")
            TotalSpread.SetText 3, pkgNum, Format(Val(varPrice) * (Val(oldQuan) + Quan), "#0.00")
        End If
        varPrice = Null
    Next
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "UpdateTotal " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

Public Sub SetHeaders(Alpha As Boolean, ShowTotal As Boolean)
    TotalSpread.RowHeadersAutoText = IIf(Alpha, DispLetters, DispNumbers)
    TotalSpread.Visible = ShowTotal
    TotalLabel.Visible = ShowTotal
    JobTotalLabel.Visible = ShowTotal
    If ShowTotal = False Then
        SummarySpread.Height = OriginalHeight * 1.8
        SummarySpread.Col = 3
        SummarySpread.ColHidden = True
    Else
        SummarySpread.Height = OriginalHeight
        SummarySpread.Col = 3
        SummarySpread.ColHidden = False
    End If
End Sub

Public Function GetMaxRows() As Long
    GetMaxRows = SummarySpread.MaxRows
End Function

Public Sub ReadSummary()
    If frmSetup.ReadSpreadsheet(SummarySpread, Trim(frmStartDay.FolderName.Text) & "_PictureFile.txt", True) = True Then
        SummarySpread.RowHeadersAutoText = DispNumbers
    Else
        SummarySpread.MaxRows = 0
    End If
End Sub

Public Function MovePkgsFwd() As String
    Dim VarPkgs As Variant
    SummarySpread.GetText 2, SummarySpread.MaxRows, VarPkgs
    SummarySpread.SetText 1, SummarySpread.MaxRows, "OutOfSync"
    SummarySpread.SetText 2, SummarySpread.MaxRows, ""
    MovePkgsFwd = VarPkgs
End Function

Public Function IsEditMode(CurSubjectId As String) As Boolean
    Dim RowNum As Long, SubjectVar As Variant
    IsEditMode = False
    EditImageFile = ""
    With SummarySpread
        For RowNum = .MaxRows To 1 Step -1
            .GetText 1, RowNum, SubjectVar
            If SubjectVar = CurSubjectId Then
                    If MsgBox("Edit existing order?", vbApplicationModal + vbQuestion + vbDefaultButton1 + vbYesNo, "WARNING!  Subject data already in file.") = vbYes Then
                        .Row = RowNum
                        .Col = 6
                        EditImageFile = .Text
                        .Col = 1
                        GetEditValues
                        .SetSelection 0, RowNum, .MaxCols, RowNum
                        IsEditMode = True
                        Exit For
                    End If
            End If
        Next
    End With
End Function

Private Sub SummarySpread_DblClick(ByVal Col As Long, ByVal Row As Long)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler

    'Display image for current row
    With SummarySpread
        .Row = Row
        .Col = 6
        Dim frmView As New frmPhotoView
        If Trim(.Text) <> "" Then
            frmView.ShowPicture .Text
            frmView.Show vbModal
        End If
    End With
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "SummarySpread_DblClick " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

Private Sub UserControl_Initialize()
    OriginalHeight = SummarySpread.Height
End Sub

Public Sub GetEditValues()
        With SummarySpread
            .Col = 1: frmEdit.txtSubject.Text = .Text
            .Col = 2: frmEdit.txtOrder.Text = .Text
            .Col = 4: frmEdit.txtCustom.Text = .Text
            .Col = 8: frmEdit.txtCode.Text = .Text
            .Col = 9: frmEdit.txtJobType = .Text
            .Col = 10: frmEdit.txtSequence.Text = .Text
            .Col = 11: frmEdit.txtCriteria.Text = .Text
            .Col = 12: frmEdit.txtFirstName.Text = .Text
            .Col = 13: frmEdit.txtLastName.Text = .Text
            .Col = 14: frmEdit.txtAddress.Text = .Text
            .Col = 15: frmEdit.txtCity.Text = .Text
            .Col = 16: frmEdit.txtState.Text = .Text
            .Col = 17: frmEdit.txtZIP.Text = .Text
            .Col = 18: frmEdit.txtMother.Text = .Text
            .Col = 19: frmEdit.txtHomePhone.Text = .Text
            .Col = 20: frmEdit.txtFather.Text = .Text
            .Col = 21: frmEdit.txtWorkPhone.Text = .Text
            .Col = 22: frmEdit.txtGender.Text = .Text
            .Col = 23: frmEdit.txtBirthDate.Text = .Text
            .Col = 24: frmEdit.txtTeacher.Text = .Text
            .Col = 25: frmEdit.txtGrade.Text = .Text
            .Col = 26: frmEdit.txtHomeRoom.Text = .Text
            .Col = 27: frmEdit.txtTrack.Text = .Text
            .Col = 28: frmEdit.txtYear.Text = .Text
            .Col = 29: frmEdit.txtSchool.Text = .Text
            .Col = 30: frmEdit.txtCustom1.Text = .Text
            .Col = 31: frmEdit.txtCustom2.Text = .Text
            .Col = 32: frmEdit.txtCustom3.Text = .Text
            .Col = 33: frmEdit.txtCustom4.Text = .Text
            .Col = 34: frmEdit.txtCustom5.Text = .Text
            .Col = 35: frmEdit.txtCustom6.Text = .Text
            .Col = 36: frmEdit.txtCustom7.Text = .Text
            .Col = 37: frmEdit.txtCustom8.Text = .Text
        End With
End Sub

Public Sub ClearEditValues()
    frmEdit.txtCode.Text = ""
    frmEdit.txtSubject.Text = ""
    frmEdit.txtOrder.Text = ""
    frmEdit.txtCustom.Text = ""
    frmEdit.txtSequence.Text = ""
    frmEdit.txtCriteria.Text = ""
    frmEdit.txtFirstName.Text = ""
    frmEdit.txtLastName.Text = ""
    frmEdit.txtAddress.Text = ""
    frmEdit.txtCity.Text = ""
    frmEdit.txtState.Text = ""
    frmEdit.txtZIP.Text = ""
    frmEdit.txtMother.Text = ""
    frmEdit.txtHomePhone.Text = ""
    frmEdit.txtFather.Text = ""
    frmEdit.txtWorkPhone.Text = ""
    frmEdit.txtGender.Text = ""
    frmEdit.txtBirthDate.Text = ""
    frmEdit.txtTeacher.Text = ""
    frmEdit.txtGrade.Text = ""
    frmEdit.txtHomeRoom.Text = ""
    frmEdit.txtTrack.Text = ""
    frmEdit.txtYear.Text = ""
    frmEdit.txtSchool.Text = ""
    frmEdit.txtCustom1.Text = ""
    frmEdit.txtCustom2.Text = ""
    frmEdit.txtCustom3.Text = ""
    frmEdit.txtCustom4.Text = ""
    frmEdit.txtCustom5.Text = ""
    frmEdit.txtCustom6.Text = ""
    frmEdit.txtCustom7.Text = ""
    frmEdit.txtCustom8.Text = ""
    frmEdit.txtJobType.Text = ""
End Sub

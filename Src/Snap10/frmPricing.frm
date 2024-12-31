VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPricing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Price Details"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2160
      Top             =   10140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton SaveButton 
      Caption         =   "Save"
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
      Left            =   7950
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   900
      Width           =   1935
   End
   Begin VB.CommandButton OpenButton 
      Caption         =   "Open"
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
      Left            =   7950
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   90
      Width           =   1935
   End
   Begin VB.Frame TouchPadFrame 
      BackColor       =   &H00808080&
      Caption         =   "Touch Pad"
      Height          =   3825
      Left            =   60
      TabIndex        =   3
      Top             =   1320
      Width           =   4815
      Begin VB.Label KeyPad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   0
         Left            =   150
         TabIndex        =   16
         Top             =   2940
         Width           =   1785
      End
      Begin VB.Label KeyPad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   1
         Left            =   150
         TabIndex        =   15
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label KeyPad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   2
         Left            =   1110
         TabIndex        =   14
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label KeyPad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   3
         Left            =   2070
         TabIndex        =   13
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label KeyPad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   4
         Left            =   150
         TabIndex        =   12
         Top             =   1140
         Width           =   825
      End
      Begin VB.Label KeyPad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   5
         Left            =   1110
         TabIndex        =   11
         Top             =   1140
         Width           =   825
      End
      Begin VB.Label KeyPad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   6
         Left            =   2070
         TabIndex        =   10
         Top             =   1140
         Width           =   825
      End
      Begin VB.Label KeyPad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   7
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   825
      End
      Begin VB.Label KeyPad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   8
         Left            =   1110
         TabIndex        =   8
         Top             =   240
         Width           =   825
      End
      Begin VB.Label KeyPad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   9
         Left            =   2070
         TabIndex        =   7
         Top             =   240
         Width           =   825
      End
      Begin VB.Label KeyPad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   10
         Left            =   3030
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label KeyPad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   11
         Left            =   3030
         TabIndex        =   5
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label KeyPad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   12
         Left            =   2070
         TabIndex        =   4
         Top             =   2940
         Width           =   825
      End
   End
   Begin VB.CommandButton ExitButton 
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
      Left            =   60
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   10140
      Width           =   1935
   End
   Begin VB.TextBox Price 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   4785
   End
   Begin FPSpread.vaSpread PriceSpread 
      Height          =   10815
      Left            =   4920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   2805
      _Version        =   393216
      _ExtentX        =   4948
      _ExtentY        =   19076
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      MaxCols         =   1
      MaxRows         =   99
      MoveActiveOnFocus=   0   'False
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmPricing.frx":0000
      UserResize      =   0
      ClipboardOptions=   0
      CellNoteIndicator=   3
   End
End
Attribute VB_Name = "frmPricing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Private Sub Form_Activate()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Me.PriceSpread.Row = 1
    Me.PriceSpread.SetSelection 1, Me.PriceSpread.Row, 2, Me.PriceSpread.Row
    Me.Price.SetFocus
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":Form_Activate " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

Private Sub ExitButton_Click()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    frmSetup.WriteSpreadsheet Me.PriceSpread, Trim(frmStartDay.FolderName.Text) & "_PriceFile.txt", True
    Me.Hide
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":ExitButton_Click " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub


Private Sub OpenButton_Click()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    With Me.CommonDialog
        .DefaultExt = "txt"
        .ShowOpen
        If .FileName <> "" Then
            frmSetup.ReadSpreadsheet PriceSpread, .FileName, False
            PriceSpread.LoadTextFile .FileName, "", ",", vbCrLf, LoadTextFileClearDataOnly + LoadTextFileColHeaders + LoadTextFileRowHeaders, ""
        End If
    End With
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":OpenButton_Click " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub


Private Sub SaveButton_Click()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    With Me.CommonDialog
        .DefaultExt = "txt"
        .ShowSave
        If .FileName <> "" Then
            PriceSpread.ExportRangeToTextFile 1, 1, PriceSpread.MaxCols, PriceSpread.MaxRows, .FileName, "", ",", vbCrLf, ExportToTextFileCreateNewFile + ExportToTextFileColHeaders + ExportToTextFileRowHeaders, ""
        End If
    End With
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":SaveButton_Click " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

Private Sub Price_KeyPress(KeyAscii As Integer)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    If KeyAscii <> 13 And KeyAscii <> Asc("/") And KeyAscii <> Asc(".") And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyBack And (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
        KeyAscii = 0
    End If
    Select Case KeyAscii                        'Case on the key pressed
        Case 13                                 'If it was the Enter key
            Me.PriceSpread.SetText 1, Me.PriceSpread.ActiveRow, Me.Price.Text
            If Me.PriceSpread.ActiveRow < 99 Then
                Me.PriceSpread.Row = Me.PriceSpread.ActiveRow
                Me.PriceSpread.Row = Me.PriceSpread.Row + 1
                Me.PriceSpread.SetSelection 1, Me.PriceSpread.Row, 2, Me.PriceSpread.Row
            End If
            Me.Price.Text = ""
        Case Asc("/")
            Me.Price.Text = ""
            KeyAscii = 0
    End Select
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":Price_KeyPress " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

Private Sub PriceSpread_Click(ByVal Col As Long, ByVal Row As Long)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Me.Price.Text = ""
    Me.Price.SetFocus
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":PriceSpread_Click " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

Private Sub KeyPad_Click(Index As Integer)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim tmpcolor As Long
    Me.TouchPadFrame.Enabled = False
    tmpcolor = Me.KeyPad(Index).BackColor
    Me.KeyPad(Index).BackColor = vbBlack
    Me.KeyPad(Index).ForeColor = vbWhite
    Select Case Index
        Case 0 To 9
            SendKeys "{" & Trim(Str(Index)) & "}"
        Case 10
            SendKeys "{/}"
        Case 11
            SendKeys "{ENTER}"
        Case 12
            SendKeys "{.}"
    End Select
    Me.KeyPad(Index).BackColor = tmpcolor
    Me.KeyPad(Index).ForeColor = vbBlack
    DoEvents
    Me.TouchPadFrame.Enabled = True
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":KeyPad_Click " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

Public Sub ClearPrices()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim RowNum As Long
    With Me.PriceSpread
        For RowNum = 1 To .MaxRows
            .SetText 1, RowNum, ""
        Next
    End With
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":ClearPrices " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub


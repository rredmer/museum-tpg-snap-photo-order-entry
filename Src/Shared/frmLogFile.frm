VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmLogFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log File"
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
      Left            =   90
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   10170
      Width           =   1935
   End
   Begin FPSpread.vaSpread LogSpread 
      Height          =   10035
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   10215
      _Version        =   393216
      _ExtentX        =   18018
      _ExtentY        =   17701
      _StockProps     =   64
      AutoCalc        =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      MaxCols         =   4
      MaxRows         =   1
      MoveActiveOnFocus=   0   'False
      OperationMode   =   1
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   2
      SpreadDesigner  =   "frmLogFile.frx":0000
      UserResize      =   1
      ClipboardOptions=   0
      CellNoteIndicator=   3
   End
   Begin FPSpread.vaSpread StatSpread 
      Height          =   10035
      Left            =   10290
      TabIndex        =   2
      Top             =   60
      Width           =   5025
      _Version        =   393216
      _ExtentX        =   8864
      _ExtentY        =   17701
      _StockProps     =   64
      AutoCalc        =   0   'False
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
      FormulaSync     =   0   'False
      MaxCols         =   2
      MaxRows         =   5
      MoveActiveOnFocus=   0   'False
      OperationMode   =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmLogFile.frx":0342
      UserResize      =   1
      ClipboardOptions=   0
      CellNoteIndicator=   3
   End
End
Attribute VB_Name = "frmLogFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public UseErrorHandler As Boolean               'Specify whether errors are trapped
Private IsLogFileOpen As Boolean                'Set TRUE when log file is open
Private LogFile As TextStream                   'Pointer to log file
Private OldMainColor As Long

Public Enum LogTypes
    InfoMsg = 0
    DebugMsg = 1
    ErrorMsg = 2
End Enum

Public Enum CounterTypes
    DownloadTime = 0
    FilterTime = 1
    DisplayTime = 2
    SaveTime = 3
    TotalTime = 4
End Enum

Private Sub ExitButton_Click()
    Me.Hide
    ClearError
End Sub

Private Sub Form_Load()
    
    'Set LogFile = FS.OpenTextFile(UsbDrive & BaseFolder & "LogFile.txt", ForAppending, True, TristateTrue)
    UseErrorHandler = True
    IsLogFileOpen = True        'Drive is ready, log file is open
    LogSpread.MaxRows = 0
    LogText LogTypes.InfoMsg, "Program Started. S/W Version: " & App.Major & "." & App.Minor & "." & App.Revision
    OldMainColor = frmMain.OrderFrame.BackColor
    
End Sub


Public Sub ClearError()
    frmMain.OrderFrame.BackColor = OldMainColor
End Sub

'****************************************************************************
'**                                                                        **
'** Subroutine.: LogText                                                   **
'**                                                                        **
'** Description: Appends text to the application log file with time stamp. **
'**                                                                        **
'****************************************************************************
Public Function LogText(EventType As LogTypes, EventMsg As String)
    If UseErrorHandler Then On Error GoTo ErrorHandler
    Dim EvtString As String
    Select Case EventType
        Case 0: EvtString = "Info"
        Case 1: EvtString = "Debug"
        Case 2:
            EvtString = "Error"
            frmMain.OrderFrame.BackColor = vbRed
    End Select
    With Me.LogSpread
        .MaxRows = .MaxRows + 1
        .SetText 1, .MaxRows, Now
        .SetText 2, .MaxRows, Format(Timer, "#,#.00")
        .SetText 3, .MaxRows, EvtString
        .SetText 4, .MaxRows, EventMsg
    End With
    frmSetup.UpdateLogFile Me.LogSpread, "LogFile.txt", False
    Err.Clear
    Exit Function
ErrorHandler:
    Err.Clear
    Resume Next
End Function

Public Sub Update(Stat As CounterTypes, StatValue As Double)
    With Me.StatSpread
        .Col = 2
        .Row = Stat + 1
        .Text = Format(StatValue, "#.00")
    End With
End Sub


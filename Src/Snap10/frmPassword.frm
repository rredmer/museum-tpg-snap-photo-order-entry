VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Password"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPassword 
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
      Left            =   30
      TabIndex        =   15
      Top             =   60
      Width           =   4785
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
      Left            =   30
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   10110
      Width           =   1935
   End
   Begin VB.Frame TouchPadFrame 
      BackColor       =   &H00808080&
      Caption         =   "Touch Pad"
      Height          =   3825
      Left            =   30
      TabIndex        =   0
      Top             =   1290
      Width           =   4815
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
         TabIndex        =   13
         Top             =   2940
         Width           =   825
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
         TabIndex        =   12
         Top             =   2040
         Width           =   1695
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
         TabIndex        =   11
         Top             =   240
         Width           =   1695
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   240
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
         TabIndex        =   8
         Top             =   240
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   1140
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
         TabIndex        =   5
         Top             =   1140
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   2040
         Width           =   825
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
         TabIndex        =   2
         Top             =   2040
         Width           =   825
      End
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
         TabIndex        =   1
         Top             =   2940
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IsPasswordValid As Boolean
Public AdminPassword As String

Private Sub Form_Load()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    IsPasswordValid = False
    Me.txtPassword.Text = ""
    Me.TouchPadFrame.Enabled = True
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":Form_Load " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

Private Sub Form_Activate()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Me.txtPassword.SetFocus
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":Form_Activate " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

Private Sub ExitButton_Click()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    IsPasswordValid = False
    Me.Hide
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":ExitButton_Click " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
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
    Sleep 50
    Me.KeyPad(Index).BackColor = tmpcolor
    Me.KeyPad(Index).ForeColor = vbBlack
    DoEvents
    Me.TouchPadFrame.Enabled = True
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":KeyPad_Click " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

'****************************************************************************
'**                                                                        **
'** Subroutine.: txtPassword_KeyPress                                      **
'**                                                                        **
'** Description: This routine handles hey presses in the folder text box.  **
'**                                                                        **
'****************************************************************************
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Select Case KeyAscii                        'Case on the key pressed
        Case Asc("/"), 27                    'If it was a slash character - this is the clear key
            Me.txtPassword.Text = ""
            IsPasswordValid = False
            Me.Hide
        Case 13                                 'If it was the Enter key
            If Trim(Me.txtPassword.Text) = AdminPassword Then
                IsPasswordValid = True
                Me.Hide
            Else
                IsPasswordValid = False
            End If
    End Select
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":txtPassword_KeyPress " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

Private Sub txtPassword_Validate(Cancel As Boolean)
    'Cancel = True
End Sub


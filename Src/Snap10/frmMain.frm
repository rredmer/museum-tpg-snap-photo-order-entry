VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SNAP '08"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   15360
   Begin VB.Timer DriveTimer 
      Interval        =   1000
      Left            =   14880
      Top             =   10530
   End
   Begin FPSpread.vaSpread StatusSpread 
      Height          =   795
      Left            =   12630
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   10140
      Width           =   2655
      _Version        =   393216
      _ExtentX        =   4683
      _ExtentY        =   1402
      _StockProps     =   64
      Enabled         =   0   'False
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      ButtonDrawMode  =   31
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
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
      MaxRows         =   3
      MoveActiveOnFocus=   0   'False
      OperationMode   =   1
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmMain.frx":0000
      UserResize      =   2
      ClipboardOptions=   0
      CellNoteIndicator=   3
   End
   Begin VB.CommandButton MainButton 
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   3
      Left            =   10380
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   10140
      Width           =   2235
   End
   Begin VB.CommandButton MainButton 
      Caption         =   "Pricing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   2
      Left            =   8100
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   10140
      Width           =   2235
   End
   Begin VB.CommandButton MainButton 
      Caption         =   "End Day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   1
      Left            =   6510
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   10140
      Width           =   1545
   End
   Begin VB.Frame TouchPadFrame 
      BackColor       =   &H00808080&
      Caption         =   "Touch Pad"
      Height          =   3825
      Left            =   60
      TabIndex        =   9
      Top             =   7050
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   2940
         Width           =   1785
      End
   End
   Begin VB.CommandButton MainButton 
      Caption         =   "Start Day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   0
      Left            =   4920
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   10140
      Width           =   1545
   End
   Begin VB.Frame OrderFrame 
      Caption         =   "Order Entry"
      Height          =   6975
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   4875
      Begin VB.TextBox CustomField 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2010
         MaxLength       =   20
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1230
         Width           =   2745
      End
      Begin VB.TextBox FormattedPackages 
         Enabled         =   0   'False
         Height          =   1035
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1770
         Width           =   4605
      End
      Begin VB.TextBox PackageQuantity 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1110
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1230
         Width           =   855
      End
      Begin VB.TextBox Packages 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1230
         Width           =   855
      End
      Begin VB.TextBox LastBarcode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         MaxLength       =   20
         TabIndex        =   1
         Top             =   480
         Width           =   4605
      End
      Begin FPSpread.vaSpread PackageSpread 
         Height          =   3375
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2850
         Width           =   4605
         _Version        =   393216
         _ExtentX        =   8123
         _ExtentY        =   5953
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   3
         MaxRows         =   20
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   2
         SpreadDesigner  =   "frmMain.frx":0345
         UserResize      =   0
         ClipboardOptions=   0
         CellNoteIndicator=   3
      End
      Begin VB.Label BarcodeLabel 
         Caption         =   "Custom (+ to Edit)"
         Height          =   255
         Index           =   4
         Left            =   2010
         TabIndex        =   30
         Top             =   1020
         Width           =   2655
      End
      Begin VB.Label TotalDueLabel 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$0.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1020
         TabIndex        =   26
         Top             =   6270
         Width           =   3735
      End
      Begin VB.Label BarcodeLabel 
         Caption         =   "Total Due:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   6450
         Width           =   825
      End
      Begin VB.Label BarcodeLabel 
         Caption         =   "Quantity"
         Height          =   255
         Index           =   1
         Left            =   1110
         TabIndex        =   7
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label BarcodeLabel 
         Caption         =   "Package"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label BarcodeLabel 
         Caption         =   "Subject"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Width           =   1035
      End
   End
   Begin TabDlg.SSTab MainTab 
      Height          =   10065
      Left            =   4920
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   60
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   17754
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   882
      WordWrap        =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Camera"
      TabPicture(0)   =   "frmMain.frx":0ABF
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MainWIAControl"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Summary"
      TabPicture(1)   =   "frmMain.frx":0ADB
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "MainSummaryControl"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Explorer"
      TabPicture(2)   =   "frmMain.frx":0AF7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ExploreC"
      Tab(2).Control(1)=   "ExploreUSB"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Help"
      TabPicture(3)   =   "frmMain.frx":0B13
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "HelpLabel(3)"
      Tab(3).Control(1)=   "HelpLabel(2)"
      Tab(3).Control(2)=   "HelpLabel(1)"
      Tab(3).Control(3)=   "HelpLabel(0)"
      Tab(3).Control(4)=   "ClearErrorButton"
      Tab(3).ControlCount=   5
      Begin VB.CommandButton ClearErrorButton 
         Caption         =   "Reset Error Indicator"
         Height          =   705
         Left            =   -74910
         TabIndex        =   43
         Top             =   1800
         Width           =   1965
      End
      Begin SNAP10.ExplorerControl ExploreC 
         Height          =   9375
         Left            =   -74910
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   570
         Width           =   4875
         _extentx        =   8599
         _extenty        =   16536
      End
      Begin SNAP10.SummaryControl MainSummaryControl 
         Height          =   9405
         Left            =   30
         TabIndex        =   36
         Top             =   540
         Width           =   10245
         _extentx        =   18071
         _extenty        =   16589
      End
      Begin SNAP10.WIAControl MainWIAControl 
         Height          =   9465
         Left            =   -74970
         TabIndex        =   34
         Top             =   570
         Width           =   10305
         _extentx        =   18177
         _extenty        =   16695
      End
      Begin SNAP10.ExplorerControl ExploreUSB 
         Height          =   9375
         Left            =   -70020
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   570
         Width           =   4875
         _extentx        =   8599
         _extenty        =   16536
      End
      Begin VB.Label HelpLabel 
         Caption         =   $"frmMain.frx":0B2F
         Height          =   285
         Index           =   0
         Left            =   -74910
         TabIndex        =   40
         Top             =   600
         Width           =   10155
      End
      Begin VB.Label HelpLabel 
         Caption         =   "Press [+] to enter custom text."
         Height          =   285
         Index           =   1
         Left            =   -74910
         TabIndex        =   39
         Top             =   870
         Width           =   10155
      End
      Begin VB.Label HelpLabel 
         Caption         =   "Press [SPACEBAR] to clear the current image (if enabled in Setup)."
         Height          =   285
         Index           =   2
         Left            =   -74910
         TabIndex        =   38
         Top             =   1140
         Width           =   10155
      End
      Begin VB.Label HelpLabel 
         Caption         =   "Enter a package quantity of zero [0] to delete a package from the list."
         Height          =   285
         Index           =   3
         Left            =   -74910
         TabIndex        =   37
         Top             =   1410
         Width           =   10155
      End
   End
   Begin VB.Frame UsbFrame 
      Caption         =   "USB DRIVE NOT READY"
      Height          =   6945
      Left            =   30
      TabIndex        =   27
      Top             =   60
      Width           =   4875
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "frmMain.frx":0BC9
         Top             =   240
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'**                                                                          **
'**  Project.....: EzSNAP                                                    **
'**                                                                          **
'**  Module......: frmMain - Application Main Form (Called on Startup)       **
'**                                                                          **
'**  Environment.: Microsoft Visual Basic 6, Service Pack 6                  **
'**                                                                          **
'**  Description.: This form provides the main interface for EzSNAP.         **
'**                Keyboard Notes:                                           **
'**                The [Enter] key is used to change application states from **
'**                waiting for barcode scan to package entry to waiting for a**
'**                picture to be taken.                                      **
'**                The [/] key is used to clear the current data entry.  If  **
'**                the key is pressed twice in a row it will step back to the**
'**                previous field of data entry.                             **
'**                                                                          **
'**  Functions...:                                                           **
'**                Form_Load (Event) Called on application start-up.         **
'**                LastBarcode_KeyPress (Event) Called on key press on field.**
'**                LastBarcode_Change (Event) Called when field value change.**
'**                Packages_KeyPress (Event) Called on key press on field.   **
'**                Packages_Change (Event) Called when field value changes.  **
'**                                                                          **
'**  References..: Microsoft Visual Basic for Applications                   **
'**                Visual Basic runtime objects and procedures               **
'**                Visual Basic object and procedures                        **
'**                OLE Automation                                            **
'**                Microsoft Windows Image Acquisition Library v2.0          **
'**                Microsoft Scripting Runtime                               **
'**                                                                          **
'**  Components..: Microsoft Windows Common Controls 6.0 (SP6)               **
'**                Microsoft Windows Image Acquisition Library v2.0          **
'**                                                                          **
'**  History.....:                                                           **
'**   07/15/05 V1.00 RDR Designed and programmed.                            **
'**                                                                          **
'**  (c) 1992-2010 Technical Products Group Inc.  All rights reserved.       **
'******************************************************************************
'Cashiers need a way to indicate that a package is free or reduced cost because it is a staff or retake 4 conditions here. The photographer needs the same.

'COMPLETED: Put job ID in picture file. We lost it when you added custom field for josephs name.
'COMPLETED: Use alpha in Exif when alpha is enabled.
'COMPLETED: We loose exposure info when we right the exif, any way to keep it.
'COMPLETED: Picture review on camera. Click the line or rescan the card shows photo and edit re-rights the exif package.
'COMPLETED: Added configuration to the edit screen
'Enter student data and print new camera ticket. Would like to use a cheep thermal printer for this.
'Routine to update Options/Preferences using jump drive like Update.
'On Screen crop adjust??? Pushing it I know but how cool if you could click on the top of the head and the chin and auto size it.

Option Explicit                                 'Require explicit variable declaration
Private AlphaMode As Boolean                    'Set TRUE when entering ALPHA packages (A..Z)
Private ShowTotal As Boolean                    'Set TRUE to display the total due for the job
Private SingleDigitEntry As Boolean             'Set TRUE to use single digital package entry (0-9, A-Z).
Private UsbMode As Boolean                      'Set TRUE if using USB Drive
Private EditMode As Boolean                     'Set TRUE if editing an existing record
Private ImageTimeOutVal As Double               'The number of seconds to display an image
Private ImageTimeOutCnt As Double               'Counter used for image preview timeout
Private RepeatSubject As Boolean
Private MainCaption As String                   'The Main Window caption (used for demo mode)
Private LastBarcodeString As String             'The last barcode scanned
Private PackageStringNum As String              'The package string expressed as numbers
Private AdminPassword As String                 'Configurable admin password
Private ReturnFromCustom As Integer             'Field to return to after entering custom text
Private ImgRotate As Integer                    'Image Rotation Setting
Private InTimer As Boolean                      'Set TRUE when in timer routine (avoid recursion)
Private HotFolder As Folder                     'Set to FileSystem Folder Object for monitoring hotfolder
Private fl As File                              'File pointer for processing hotfolder files
Private EditFormFileName As String

Private Sub ClearErrorButton_Click()
    frmLogFile.ClearError
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("@") Then
        With frmEdit
            '.UpdateFields
            .txtSubject.Text = Me.LastBarcode.Text
            .txtOrder.Text = Me.FormattedPackages.Text
            .txtCustom.Text = Me.CustomField.Text
            
            .Show vbModal
            
            'Save Settings File for Edit Form
            SaveEditFormSettings
            
        End With
        KeyAscii = 0
    End If
End Sub

'****************************************************************************
'**                                                                        **
'** Subroutine.: Form_Load                                                 **
'**                                                                        **
'** Description: This routine initializes form controls - called on start. **
'**                                                                        **
'****************************************************************************
Private Sub Form_Load()
'    On Error GoTo ErrorHandler
    
    DriveTimer.Enabled = False
    ExploreUSB.Visible = False
    
    'DemoMode = False
    'DemoExceeded = False
    EditMode = False
    LastBarcodeString = ""                      'Initialize the last barcode scan (used for retakes)
    MainCaption = "SNAP '10 v" & App.Minor & "." & App.Revision
    
    SetKeys "!!@%!(^&", "!@$*#&))#^)", "SNAP08.DAT"
    InTimer = False
    
    Load frmStartDay
    Load frmLogFile
    Load frmPassword
    Load frmSetup
    Load frmPricing                             'Load the pricing form
    Load frmEdit
    
    LoadEditFormSettings
   
    frmSetup.IsLic
    
    MachineId = GetComputerID
    
    If FS.FolderExists("c:\HotFolder\") = False Then
        FS.CreateFolder "C:\HotFolder\"
    End If
    Set HotFolder = FS.GetFolder("c:\hotfolder\")
    
    GetSettings
    MainTab.Tab = 1
    Exit Sub
ErrorHandler:
    'frmLogFile.LogText ErrorMsg, "Form_Load " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

Private Sub GetSettings()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    AlphaMode = IIf(Val(frmSetup.GetSetting(AlphaPackage)) = 1, True, False)
    ShowTotal = IIf(Val(frmSetup.GetSetting(ShowTotals)) = 1, True, False)
    SingleDigitEntry = IIf(Val(frmSetup.GetSetting(SingleDigits)) = 1, True, False)
    UsbMode = IIf(Val(frmSetup.GetSetting(AutoUSB)) = 1, True, False)
    RepeatSubject = IIf(Val(frmSetup.GetSetting(RepeatSubjectSetting)) = 1, True, False)
    ImageTimeOutVal = CDbl(frmSetup.GetSetting(ImageTimeout))
    ImgRotate = Val(frmSetup.GetSetting(RotateImageSetting))
    AdminPassword = frmSetup.GetSetting(AdminPWD)
    'If frmSetup.IsLic = True Then
        Me.Caption = MainCaption
    'Else
    '    Me.Caption = "SNAP DEMO (LIMITED NUMBER OF EXPOSURES PER JOB)"
    'End If
    '---- If there is no folder open, get a new job folder
    If Not frmSetup.IsFolderFileOpen Then
        GetJobFolder False
        frmSetup.CheckDrive UsbMode
        frmSetup.FillGrid
    End If
    If AlphaMode Then
        Me.MainSummaryControl.SetHeaders True, ShowTotal
        frmPricing.PriceSpread.RowHeadersAutoText = DispLetters
    Else
        Me.MainSummaryControl.SetHeaders False, ShowTotal
        frmPricing.PriceSpread.RowHeadersAutoText = DispNumbers
    End If
    Me.MainButton(2).Enabled = ShowTotal
    Me.MainButton(2).Visible = ShowTotal
    
    ExploreC.SetDrive "C:\", "Drive C"
    DriveTimer.Enabled = True
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "GetSettings " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

Private Sub DriveTimer_Timer()                              'MUST be Interval=1000
    If InTimer Then Exit Sub                                'Avoid recursion
    InTimer = True
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    frmSetup.CheckDrive UsbMode
    If UsbMode Then
        SetOrderFrame frmSetup.IsUsbReady
    Else
        SetOrderFrame True
    End If
    
    If frmSetup.IsUsbReady Then
        If ExploreUSB.Visible = False Then
            ExploreUSB.Visible = True
            ExploreUSB.SetDrive frmSetup.UsbDrive.SrcDrive.Path, "USB Drive"
            ExploreUSB.RefreshList
        End If
    Else
        ExploreUSB.Visible = False
    End If
    
    Me.MainTab.TabEnabled(0) = Me.MainWIAControl.IsCameraReady
    Me.MainTab.TabVisible(0) = Me.MainWIAControl.IsCameraReady
    Me.StatusSpread.SetText 2, 1, Me.MainWIAControl.CameraName

    If Me.MainWIAControl.IsPictureReady Then                'If a picture is ready on the camera
        Me.MainWIAControl.IsPictureReady = False            'Clear the flag to avoid recursion
        Me.GetThePicture True, ""
    End If
    
    '---- Check for new image in hotfolder
    If HotFolder.Files.Count > 0 Then
        For Each fl In HotFolder.Files
            If UCase(Right(fl.Name, 3)) = "JPG" Then
                GetThePicture False, fl.Path
                fl.Delete True
            End If
        Next
    End If
    
    If Me.MainTab.Tab = 0 Then                              'If viewing an image
        If ImageTimeOutVal > 0 Then                         'And Timeout is selected
            If ImageTimeOutCnt > ImageTimeOutVal Then
                ImageTimeOutCnt = 0
                MainTab.Tab = 1
                If LastBarcode.Enabled = True Then LastBarcode.SetFocus
            Else
                ImageTimeOutCnt = ImageTimeOutCnt + 1
            End If
        End If
    End If
    
    InTimer = False
    Exit Sub
ErrorHandler:
    Err.Clear
    InTimer = False
End Sub

Private Sub MainButton_Click(Index As Integer)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Select Case Index
        Case 0
            DriveTimer.Enabled = False
            GetJobFolder True
            frmSetup.FoldersValidated = False
            frmSetup.UsbFoldersValidated = False
            frmSetup.CheckDrive UsbMode
            frmSetup.FillGrid
            DriveTimer.Enabled = True
        Case 1
            If MsgBox("Are you sure?", vbApplicationModal + vbQuestion + vbDefaultButton2 + vbYesNo, "EXIT SNAP") = vbYes Then
                End
            Else
                If MsgBox("Are you sure?", vbApplicationModal + vbQuestion + vbDefaultButton2 + vbYesNo, "END DAY AND SHUT DOWN") = vbYes Then
                    ShutDownWindows False, True
                End If
            End If
        Case 2
            frmPricing.PriceSpread.ClearRange 0, 1, 0, 99, False
            frmPricing.Show vbModal
            If ShowTotal Then
                Me.MainSummaryControl.FillTotals
            End If
        Case 3
            frmPassword.AdminPassword = AdminPassword
            If frmPassword.IsPasswordValid = False Or Trim(AdminPassword) = "" Then
                frmPassword.Show vbModal
            End If
            If frmPassword.IsPasswordValid = True Then
                frmSetup.Show vbModal
            End If
            GetSettings
    End Select
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "MainButton_Click " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

Private Sub MainTab_Click(PreviousTab As Integer)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    If Me.LastBarcode.Enabled = True Then
        Me.LastBarcode.SetFocus
    End If
    ExploreC.RefreshList
    If frmSetup.IsUsbReady Then ExploreUSB.RefreshList
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "MainTab_Click " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
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
    frmLogFile.LogText ErrorMsg, "KeyPad_Click " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

'****************************************************************************
'**                                                                        **
'** Subroutine.: LastBarcode_KeyPress                                      **
'**                                                                        **
'** Description: This routine handles hey presses in the barcode text box. **
'**                                                                        **
'****************************************************************************
Private Sub LastBarcode_KeyPress(KeyAscii As Integer)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Select Case KeyAscii                        'Case on the key pressed
        Case 13, 27                             'If it was the Enter key
            If Len(Trim(Me.LastBarcode.Text)) > 1 Then 'And there is something enterred for the barcode
                '---- Search for subject in file
                If Me.MainSummaryControl.IsEditMode(Me.LastBarcode.Text) = True Then
                    EditMode = True
                    Me.OrderFrame.Caption = "*** EDIT ORDER INFORMATION ***"
                    
                    '--- Load the image if it exists
                    If Me.MainSummaryControl.EditImageFile <> "" Then
                        If FS.FileExists(Me.MainSummaryControl.EditImageFile) Then
                            Me.MainWIAControl.LoadPicture Me.MainSummaryControl.EditImageFile
                        End If
                    End If
                Else
                    Me.MainSummaryControl.ClearEditValues
                End If
                
                Me.Packages.Text = ""           'Initialize the package string
                Me.FormattedPackages.Text = ""
                PackageStringNum = ""
                
                Me.LastBarcode.Enabled = False
                Me.Packages.Enabled = True
                Me.Packages.SetFocus
            End If
        Case Asc("/"), 9                        'If it was a slash character - this is the clear key
            KeyAscii = 0                        'Clear the key from the keyboard buffer
            Me.LastBarcode.Text = ""            'Clear the barcode string
            EditMode = False
            Me.OrderFrame.Caption = "Order Entry"
        
        '--- Legacy CODE [-] key was used for a retake.
        'Case Asc("-")                           'If it was a minus key, this is retake of prior picture
        '    If LastBarcodeString <> "" Then
        '        Me.LastBarcode.Text = LastBarcodeString
        '        KeyAscii = 0
        '    Else
        '        'Retake is not an option, there is no prior scan to retake!
        '        KeyAscii = 0
        '    End If
        
        Case Asc(" ")                           'If it was a spacebar, this is to clear the image
            If Val(frmSetup.GetSetting(ClearOnSpacebar)) = 1 Then
                MainTab.Tab = 1
                Me.LastBarcode.SetFocus
            End If
            KeyAscii = 0
        Case Asc("+")                           'Plus key is used to add custom text.
            KeyAscii = 0
            GetCustomField 0
    End Select
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "LastBarcode_KeyPress " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub
'****************************************************************************
'**                                                                        **
'** Subroutine.: Packages_KeyPress                                         **
'**                                                                        **
'** Description: This routine handles hey presses in the packages text box.**
'**                                                                        **
'****************************************************************************
Private Sub Packages_KeyPress(KeyAscii As Integer)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    
    If SingleDigitEntry Then
        Packages.MaxLength = 1
    Else
        'If ShowTotal Then
        '    Packages.MaxLength = 2
        'Else
            Packages.MaxLength = 3
        'End If
    End If
    
    
    If AlphaMode Then
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
            KeyAscii = KeyAscii - 32
        End If
        If KeyAscii <> 13 And KeyAscii <> Asc("/") And KeyAscii <> Asc("+") And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyBack And (KeyAscii < Asc("A") Or KeyAscii > Asc("Z")) Then
            KeyAscii = 0
        End If
    Else
        
        'If ShowTotal Then
        '    If KeyAscii <> 13 And KeyAscii <> Asc("/") And KeyAscii <> Asc("+") And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyBack And (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
        '        KeyAscii = 0
        '    End If
        'Else
            If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
                KeyAscii = KeyAscii - 32
            End If
        
            If KeyAscii <> 13 And KeyAscii <> Asc("/") And KeyAscii <> Asc("+") And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyBack And ((KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And (KeyAscii < Asc("A") Or KeyAscii > Asc("Z"))) Then
                    'If KeyAscii <> 13 And KeyAscii <> Asc("/") And KeyAscii <> Asc("+") And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyBack And (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
                KeyAscii = 0
            End If
        'End If
    End If
    
    Select Case KeyAscii                        'Case on the key pressed
        Case 13                                 'If it was the Enter key
            If Trim(Me.Packages.Text) = "" Then
                '--- If data entry mode, save data and increment to next subject
                If Me.MainWIAControl.IsCameraReady = False Or EditMode = True Then
                    If MsgBox("Save Order?", vbApplicationModal + vbYesNo + vbDefaultButton1 + vbQuestion, "Order Complete") = vbYes Then
                        Me.MainSummaryControl.SaveOrderText LastBarcode.Text, FormattedPackages.Text, TotalDueLabel.Caption, CustomField.Text, "", "", False, EditMode
                        
                        If Me.MainSummaryControl.EditImageFile <> "" Then
                            
                            Me.MainWIAControl.ApplyFilters Trim(LastBarcode.Text), FormattedPackages.Text, Trim(CustomField.Text), 0
                            With Me.MainWIAControl.Img
                                FS.DeleteFile Me.MainSummaryControl.EditImageFile, True
                                .SaveFile Me.MainSummaryControl.EditImageFile   'Save the image file
                                If FS.FileExists(Me.MainSummaryControl.EditImageFile) = False Then         'Error writing file
                                    frmLogFile.LogText ErrorMsg, "GetThePicture Failed to Write Image File."
                                End If
                                If frmSetup.IsUsbReady = True Then
                                    Dim FileName As String
                                    FileName = frmSetup.UsbDrive.SrcDrive.Path & "\" & Mid(Me.MainSummaryControl.EditImageFile, 4)
                                    FS.DeleteFile FileName, True
                                    .SaveFile FileName
                                    If FS.FileExists(FileName) = False Then     'Error writing file
                                        frmLogFile.LogText ErrorMsg, "GetThePicture Failed to Write USB Image File."
                                    End If
                                End If
                            End With
                            
                        End If
                        
                        
                        If EditMode = True Then
                            Me.OrderFrame.Caption = "Order Entry"
                            EditMode = False
                        End If
                        UpdateStatusSpread
                        GetOrder
                    End If
                End If
            Else
                Me.PackageQuantity.Enabled = True
                Me.PackageQuantity.Text = ""
                Me.PackageQuantity.SetFocus
            End If
        Case Asc("/")                           'If it was a slash character - this is the clear key
            If Val(Me.Packages.Text) = 0 Then   'Return to barcode entry because no packages were enterred to clear
                Me.Packages.Enabled = False     'Disable the package text box
                Me.LastBarcode.Enabled = True   'Enable the barcode text box
                Me.PackageSpread.MaxRows = 0
                Me.LastBarcode.SetFocus
            Else                                'Else there were packages enterred to clear
                Me.Packages.Text = ""
            End If
            KeyAscii = 0
        Case Asc("+")
            KeyAscii = 0
            GetCustomField 1
    End Select
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "Packages_KeyPress " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

Private Sub PackageQuantity_KeyPress(KeyAscii As Integer)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    If KeyAscii <> 13 And KeyAscii <> Asc("/") And KeyAscii <> Asc("+") And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyBack And (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
        KeyAscii = 0
    End If
    Select Case KeyAscii                        'Case on the key pressed
        Case 13                                 'If it was the Enter key
            If Trim(Me.PackageQuantity.Text) = "" Then
                Me.PackageQuantity.Text = "1"
            End If
            If Val(Me.PackageQuantity.Text) = 0 Then
                Dim Pkg As Long
               '---- Delete the package from the grid
                With Me.PackageSpread
                    .Col = 1
                    For Pkg = 1 To .MaxRows
                        .Row = Pkg
                        If Trim(.Text) = Trim(Me.Packages.Text) Then
                            .DeleteRows .Row, 1
                            .MaxRows = .MaxRows - 1
                        End If
                    Next
                    .Refresh
                End With
                If ShowTotal Then
                    UpdatePrice
                End If
            Else
                UpdatePackageGrid
            End If
            '--- Update Grid & Package String
            Me.PackageQuantity.Text = ""
            Me.PackageQuantity.Enabled = False
            Me.Packages.Enabled = True
            Me.Packages.Text = ""
            Me.Packages.SetFocus
        Case Asc("/")                           'If it was a slash character - this is the clear key
            If Val(Me.PackageQuantity.Text) = 0 Then
                Me.PackageQuantity.Text = ""
                Me.PackageQuantity.Enabled = False
                Me.Packages.Enabled = True
                Me.Packages.SetFocus
            Else                                'Else there were packages enterred to clear
                Me.PackageQuantity.Text = ""
            End If
            KeyAscii = 0
        Case Asc("+")
            KeyAscii = 0
            GetCustomField 2
    End Select
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "PackageQuantity_KeyPress" & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

Private Sub CustomField_KeyPress(KeyAscii As Integer)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Select Case KeyAscii                        'Case on the key pressed
        Case 13                                 'If it was the Enter key
            Me.CustomField.Enabled = False
            Select Case ReturnFromCustom
                Case 0
                    Me.LastBarcode.Enabled = True
                    Me.LastBarcode.SetFocus
                Case 1
                    Me.Packages.Enabled = True
                    Me.Packages.SetFocus
                Case 2
                    Me.PackageQuantity.Enabled = True
                    Me.PackageQuantity.SetFocus
            End Select
    End Select
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "CustomField_KeyPress" & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

Private Sub UpdatePackageGrid()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim Price As Double, Pkg As Long, TotalDue As Double, pkgrow As Long, tmpprice As Variant, tmppkg As Variant, tmpqty As Variant, FoundPkg As Boolean
    FoundPkg = False
    Price = 0
    If ShowTotal Then
        If AlphaMode Then
            frmPricing.PriceSpread.GetText 1, CLng(Asc(Me.Packages.Text) - Asc("A") + 1), tmpprice
        Else
            Dim tpkg As Long
            
            tpkg = MakeNumeric(Me.Packages.Text)
            frmPricing.PriceSpread.GetText 1, tpkg, tmpprice
        End If
        If tmpprice <> "" Then
            Price = CDbl(tmpprice)
        Else
            Price = 0#
        End If
    End If
    With Me.PackageSpread
        For Pkg = 1 To .MaxRows
            .Row = Pkg
            .Col = 1
            If (.value = Me.Packages.Text) Then
                FoundPkg = True
                Exit For
            End If
        Next
        If Not FoundPkg Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        End If
        .Col = 1
        .Text = Me.Packages.Text
        .Col = 2
        .Text = Me.PackageQuantity.Text
        .Col = 3
        .Text = Format(Val(Me.PackageQuantity.Text) * Price, "###0.00")
    End With
    UpdatePrice
    '--- TO DO:  NEED TO MAKE NUMERIC SORT!!!
    Me.PackageSpread.Sort 1, 1, PackageSpread.MaxCols, PackageSpread.MaxRows, SortByRow, 1
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "UpdatePackageGrid " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

Public Sub UpdatePrice()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    '---- Calculate total price
    
   
    Dim Price As Double, Pkg As Long, TotalDue As Double, pkgrow As Long
    Dim tmpprice As Variant, tmppkg As Variant, tmpqty As Variant, FoundPkg As Boolean
    TotalDue = 0#
    Me.FormattedPackages.Text = ""
    PackageStringNum = ""
    For Pkg = 1 To Me.PackageSpread.MaxRows
        Me.PackageSpread.GetText 1, Pkg, tmppkg
        Me.PackageSpread.GetText 2, Pkg, tmpqty
        Me.PackageSpread.GetText 3, Pkg, tmpprice
        If tmpqty <> "" Then
            If AlphaMode Then
                Me.FormattedPackages.Text = Me.FormattedPackages.Text & CStr(tmppkg) & "-" & Format(CInt(tmpqty), "00") & ";"
                PackageStringNum = PackageStringNum & Format(Asc(tmppkg) - Asc("A") + 1, "00") & "-" & Format(CInt(tmpqty), "00") & ";"
            Else
                
                Me.FormattedPackages.Text = Me.FormattedPackages.Text & tmppkg & "-" & Format(CInt(tmpqty), "00") & ";"
                
                PackageStringNum = PackageStringNum & tmppkg & "-" & Format(CInt(tmpqty), "00") & ";"
                
                'Me.FormattedPackages.Text = Me.FormattedPackages.Text & Format(CInt(tmppkg), "00") & "-" & Format(CInt(tmpqty), "00") & ";"
                'PackageStringNum = PackageStringNum & Format(CInt(tmppkg), "00") & "-" & Format(CInt(tmpqty), "00") & ";"
            
            End If
        End If
        If tmpprice <> "" Then
            TotalDue = TotalDue + CDbl(tmpprice)
        End If
    Next
    Me.TotalDueLabel.Caption = Format(TotalDue, "###0.00")
    
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "UpdatePrice " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
End Sub

Public Sub GetJobFolder(ForceEnter As Boolean)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    
    '---- Loop until valid job folder found
    Do While ReadStatusFile() = False Or ForceEnter
        frmSetup.IsFolderFileOpen = False
        frmStartDay.OldName = frmStartDay.FolderName.Text
        frmStartDay.Show vbModal
        
        StatusSpread.SetText 2, 2, frmStartDay.FolderName.Text
        frmSetup.WriteSpreadsheet StatusSpread, "StatusFile.txt", False
        DoEvents
        ForceEnter = False
    Loop
    
    '---- Load Summary Spreadsheet
    Me.MainSummaryControl.ReadSummary
    UpdateStatusSpread
    
    '---- Load Pricing Data File
    If frmSetup.ReadSpreadsheet(frmPricing.PriceSpread, Trim(frmStartDay.FolderName.Text) & "_PriceFile.txt", True) = False Then
        frmPricing.ClearPrices
    End If
    frmSetup.IsFolderFileOpen = True
    GetOrder
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "GetJobFolder " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

Private Sub GetOrder()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Me.LastBarcode.Text = ""
    Me.PackageQuantity.Text = ""
    Me.Packages.Text = ""
    Me.FormattedPackages.Text = ""
    PackageStringNum = ""
    Me.TotalDueLabel.Caption = "0.00"
    Me.Packages.Enabled = False
    Me.LastBarcode.Enabled = True
    Me.PackageSpread.MaxRows = 0
    If Me.OrderFrame.Visible = True Then
        Me.LastBarcode.SetFocus
    End If
    Me.CustomField.Enabled = False
    Me.CustomField.Text = ""
    If ShowTotal Then
        Me.MainSummaryControl.FillTotals
    End If
    Me.MainSummaryControl.ClearEditValues
    ImageTimeOutCnt = 0
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "GetOrder " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

Private Sub GetCustomField(ReturnField As Integer)
    Me.LastBarcode.Enabled = False
    Me.PackageQuantity.Enabled = False
    Me.Packages.Enabled = False
    Me.CustomField.Enabled = True
    Me.CustomField.SetFocus
    ReturnFromCustom = ReturnField
End Sub

Public Sub UpdateStatusSpread()
    StatusSpread.SetText 2, 3, Me.MainSummaryControl.GetMaxRows + 1
    frmSetup.WriteSpreadsheet StatusSpread, "StatusFile.txt", False
End Sub

Public Sub SetOrderFrame(OrderMode As Boolean)
    OrderFrame.Enabled = OrderMode
    OrderFrame.Visible = OrderMode
    UsbFrame.Visible = Not OrderMode
    UsbFrame.Enabled = Not OrderMode
End Sub

'****************************************************************************
'**                                                                        **
'** Subroutine.: ReadStatusFile                                            **
'**                                                                        **
'** Description: Opens the status file and reads current folder name.      **
'**                                                                        **
'****************************************************************************
Public Function ReadStatusFile() As Boolean
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    ReadStatusFile = False
    Dim varFolder As Variant
    '---- Read the status file.
    If frmSetup.ReadSpreadsheet(StatusSpread, "StatusFile.txt", False) = False Then
        '---- If the status file did not read properly, write a new one
        StatusSpread.SetText 2, 2, frmStartDay.FolderName.Text
        frmSetup.WriteSpreadsheet StatusSpread, "StatusFile.txt", False
    End If
    '---- Get the folder name from the status spreadsheet
    StatusSpread.GetText 2, 2, varFolder
    frmStartDay.FolderName.Text = Trim(varFolder)
    If Len(Trim(frmStartDay.FolderName.Text)) >= 1 Then
        ReadStatusFile = True
    End If
    Exit Function
ErrorHandler:
    frmLogFile.LogText ErrorMsg, Me.Name & ":ReadStatusFile " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Function


Public Sub GetThePicture(PicMode As Boolean, HotName As String)
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler

'    On Error GoTo 0

    Me.LastBarcode.Enabled = False
    Me.Packages.Enabled = False
    Me.PackageQuantity.Enabled = False
    DriveTimer.Enabled = False
    Dim FileName As String, CameraImageName As String
    If Me.MainWIAControl.IsCameraReady = False Then         'If the camera is not ready
        Me.MainTab.TabVisible(0) = True                     'Enable the camera tab - an image is coming
        Me.MainTab.TabEnabled(0) = True
    End If
    Me.MainTab.Tab = 0                                      'Activate the camera tab
    DoEvents
    Me.MainWIAControl.UpdateProgress 0

    If Trim(Me.LastBarcode.Text) <> "" Then                 'If there is a subject enterred
        LastBarcodeString = Me.LastBarcode.Text             'Store it as the last one taken
    Else                                                    'Else a picture was taken before proper data was entered
        frmLogFile.LogText InfoMsg, "Picture taken out of sequence."
        If RepeatSubject = True Then                        'If the repeat subject option is enabled
            If Trim(LastBarcodeString) <> "" Then           'If the last barcode scanned is valid
                Me.LastBarcode.Text = LastBarcodeString     'Store it to current barcode scan
                Me.FormattedPackages.Text = Me.MainSummaryControl.MovePkgsFwd() 'Move packages forward
            Else
                Me.LastBarcode.Text = "OutOfSync"
            End If
        Else
            Me.LastBarcode.Text = "OutOfSync"
        End If
    End If
    Dim ImageNum As Integer, RootPath As String
    
    RootPath = frmSetup.MainDrive.SrcDrive.Path & "\" & frmSetup.BaseFolder & Trim(frmStartDay.GetFolderName) & "\" & Trim(frmStartDay.GetFolderName) & "_" & frmSetup.ImageFolder & Format(Me.MainSummaryControl.GetMaxRows + 1, "000") & "_" & Trim(LastBarcode.Text)
    
    FileName = RootPath & ".JPG"                            'Initialize image file name
    ImageNum = 0                                            'Start with image 0
    Do While FS.FileExists(FileName)                        'If the file already exists, ensure new image is created
        ImageNum = ImageNum + 1                             'Increment the image number and append to file name
        FileName = RootPath & "_" & Trim(Str(ImageNum)) & ".JPG"
    Loop
    
    If PicMode Then
        CameraImageName = Me.MainWIAControl.TransferPicture()
    Else
        CameraImageName = HotName
        Me.MainWIAControl.LoadPicture HotName
    End If
    Me.MainWIAControl.UpdateProgress 40, DownloadTime
    
    Me.MainWIAControl.ApplyFilters Trim(LastBarcode.Text), Trim(PackageStringNum), Trim(CustomField.Text), ImgRotate
    Me.MainWIAControl.UpdateProgress 65, FilterTime
    'If DemoExceeded = False Then
        With Me.MainWIAControl.Img
            .SaveFile FileName                          'Save the image file
            If FS.FileExists(FileName) = False Then     'Error writing file
                frmLogFile.LogText ErrorMsg, "GetThePicture Failed to Write Image File."
            End If
            If frmSetup.IsUsbReady = True Then
                FileName = frmSetup.UsbDrive.SrcDrive.Path & "\" & Mid(FileName, 4)
                .SaveFile FileName
                If FS.FileExists(FileName) = False Then     'Error writing file
                    frmLogFile.LogText ErrorMsg, "GetThePicture Failed to Write USB Image File."
                End If
            End If
            'Me.GetImageInfo Img                        'For Debug only
        End With
    'End If
    Me.MainWIAControl.UpdateProgress 85, SaveTime
    
    Me.MainSummaryControl.SaveOrderText LastBarcode.Text, FormattedPackages.Text, TotalDueLabel.Caption, CustomField.Text, CameraImageName, FileName, True, False
    UpdateStatusSpread
    GetOrder
    Me.MainWIAControl.UpdateProgress 100, TotalTime
    frmSetup.ManageDriveSpace                               'Manage drive space
    frmLogFile.LogText InfoMsg, "GetThePicture," & Trim(Me.MainSummaryControl.GetMaxRows + 1) & "," & Trim(LastBarcode.Text) & "," & Trim(Me.FormattedPackages.Text) & "," & FileName
    DriveTimer.Enabled = True
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "GetThePicture " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

Private Sub LoadEditFormSettings()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim ts As TextStream
    Dim RootPath As String
    Dim i As Integer, tmp As Integer
    RootPath = frmSetup.MainDrive.SrcDrive.Path & "\" & frmSetup.BaseFolder & "Settings\"
    EditFormFileName = RootPath & "EditFormSettings.txt"
    If FS.FileExists(EditFormFileName) = False Then
        SaveEditFormSettings
    Else
        Set ts = FS.OpenTextFile(EditFormFileName, ForReading, False, TristateFalse)
        For i = 0 To frmEdit.chkActive.Count - 1
            tmp = ts.ReadLine
            frmEdit.chkActive(i).value = tmp
        Next
        ts.Close
        Set ts = Nothing
    End If
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "LoadEditFormSettings " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

Private Sub SaveEditFormSettings()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim ts As TextStream
    Dim RootPath As String
    Dim i As Integer
    RootPath = frmSetup.MainDrive.SrcDrive.Path & "\" & frmSetup.BaseFolder & "Settings\"
    EditFormFileName = RootPath & "EditFormSettings.txt"
    Set ts = FS.CreateTextFile(EditFormFileName, True, False)
    For i = 0 To frmEdit.chkActive.Count - 1
        ts.WriteLine frmEdit.chkActive(i).value
    Next
    ts.Close
    Set ts = Nothing
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "SaveEditFormSettings " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

'//------------------------------
'// PLACE THIS CODE INTO A MODULE
'//------------------------------
Public Function MakeNumeric(strText As String) As Long
'//Check For Null String
    If strText = vbNullString Then MakeNumeric = 0

'//Create The Array (Set To 10 - Not Much Chance of it being of 10 Chars Long)
'//And the Counter '(i)' Also sTemp is for the new output after the checks
    Dim Characters(10) As String, i As Integer, sTemp As String

'//Trim Down The Text To Avoid Errors
    If Len(strText) > UBound(Characters) Then
        strText = Left(strText, UBound(Characters))
    End If
    

'//Add Each Character to an Array (Makes Easier Loops for the Checks)
    For i = 1 To Len(strText)
        Characters(i - 1) = Right(Left(strText, i), 1)
    Next

'//Loop To Create the New String Taking Away All The Non Numeric Characters
'//Using the IsNumeric Function
    For i = 0 To UBound(Characters)
    '//Check For Null String
        If Not Characters(i) = vbNullString Then
        '//Check If its Numeric & Not a Space
        '//(After the first Char IsNumeric allows Spaces inside it)
            If IsNumeric(Characters(i)) = True And Not Characters(i) = Chr(32) Then
                sTemp = sTemp & Characters(i)
            End If
        End If
    Next

'//Create The Output
   MakeNumeric = sTemp
End Function



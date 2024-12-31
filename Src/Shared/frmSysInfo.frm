VERSION 5.00
Begin VB.Form frmSysInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SysInfo"
   ClientHeight    =   10950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ReadKey 
      Caption         =   "ReadKey"
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
      Left            =   6120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   10110
      Width           =   1935
   End
   Begin VB.CommandButton SavKey 
      Caption         =   "SavKey"
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
      Left            =   4080
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   10110
      Width           =   1935
   End
   Begin VB.TextBox txtCipher 
      Height          =   2295
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   60
      Width           =   3375
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   5400
      TabIndex        =   5
      Top             =   420
      Width           =   1815
   End
   Begin VB.CommandButton cmdCipher 
      Caption         =   ">>"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   1020
      Width           =   1095
   End
   Begin VB.CommandButton cmdDecipher 
      Caption         =   "<<"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   1860
      Width           =   1095
   End
   Begin VB.CommandButton SrcButton 
      Caption         =   "Get Src"
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
      Left            =   2070
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   10110
      Width           =   1935
   End
   Begin VB.TextBox txtLicSrc 
      Height          =   2235
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   60
      Width           =   5175
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   10110
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Password"
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   7
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "frmSysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

VERSION 5.00
Begin VB.Form frmPhotoView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Photo View"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   9015
      Left            =   120
      ScaleHeight     =   8955
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "frmPhotoView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TmpImg As ImageFile                     'WIA Image File
Private TmpIP As New ImageProcess               'WIA Image process

Public Sub ShowPicture(FileName As String)

    On Error GoTo ErrorHandler
    
    Set TmpImg = New ImageFile
    TmpImg.LoadFile (FileName)

    TmpIP.Filters.Add TmpIP.FilterInfos("Scale").FilterID
    TmpIP.Filters(1).Properties("MaximumWidth") = Picture1.Width / Screen.TwipsPerPixelX
    TmpIP.Filters(1).Properties("MaximumHeight") = Picture1.Height / Screen.TwipsPerPixelY

    Set TmpImg = TmpIP.Apply(TmpImg)

    Set Me.Picture1.Picture = TmpImg.ARGBData.Picture(TmpImg.Width, TmpImg.Height)

    Exit Sub
ErrorHandler:

End Sub


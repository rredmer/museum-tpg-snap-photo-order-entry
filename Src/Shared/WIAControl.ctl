VERSION 5.00
Object = "{94A0E92D-43C0-494E-AC29-FD45948A5221}#1.0#0"; "wiaaut.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl WIAControl 
   ClientHeight    =   9435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10305
   ScaleHeight     =   9435
   ScaleWidth      =   10305
   Begin MSComctlLib.ProgressBar DownloadProgressBar 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   9090
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Enabled         =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9075
      Left            =   0
      ScaleHeight     =   9045
      ScaleWidth      =   10245
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10275
   End
   Begin WIACtl.DeviceManager wia 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "WIAControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IsCameraReady As Boolean                 'Set TRUE when camera is ready
Public IsPictureReady As Boolean                'Set TRUE when image is ready on camera
Public CameraName As String                     'WIA Device Name
Public Img As ImageFile                         'WIA Image File
Private MainTimer As New HighPerfTimer          'Windows high performance timer
Private InWIA As Boolean
Private di As DeviceInfo                        'WIA Device Info Object
Private dev As Device                           'WIA Device
Private itm As Item                             'WIA Image Item
Private IP As New ImageProcess                  'WIA Image process
Private v As New Vector                         'WIA Image vector
Private TmpImg As ImageFile                     'WIA Image File
Private TmpIP As New ImageProcess               'WIA Image process

Private Sub UserControl_Initialize()
    
    InWIA = False
    IsCameraReady = False
    IsPictureReady = False
    CameraName = ""
    
    'Register event for picture taken
    wia.RegisterEvent wiaEventDeviceConnected, "*"    'Register event for camera connected
    wia.RegisterEvent wiaEventDeviceDisconnected, "*" 'Register event for camera disconnected
    wia.RegisterEvent wiaEventItemCreated, wiaAnyDeviceID
    
    'User Comment Tag (Subject ID)
    IP.Filters.Add IP.FilterInfos("Exif").FilterID
    IP.Filters(1).Properties("ID") = 41985         '37510
    IP.Filters(1).Properties("Type") = StringImagePropertyType          'VectorOfBytesImagePropertyType
    
    'Packages
    IP.Filters.Add IP.FilterInfos("Exif").FilterID
    IP.Filters(2).Properties("ID") = 41986
    IP.Filters(2).Properties("Type") = StringImagePropertyType          'VectorOfBytesImagePropertyType
    
    'Sequence #
    IP.Filters.Add IP.FilterInfos("Exif").FilterID
    IP.Filters(3).Properties("ID") = 41987
    IP.Filters(3).Properties("Type") = StringImagePropertyType          'VectorOfBytesImagePropertyType
    
    'Job #
    IP.Filters.Add IP.FilterInfos("Exif").FilterID
    IP.Filters(4).Properties("ID") = 41988
    IP.Filters(4).Properties("Type") = StringImagePropertyType          'VectorOfBytesImagePropertyType
    
    IP.Filters.Add IP.FilterInfos("RotateFlip").FilterID
    IP.Filters.Add IP.FilterInfos("Convert").FilterID

    TmpIP.Filters.Add TmpIP.FilterInfos("Scale").FilterID
    TmpIP.Filters(1).Properties("MaximumWidth") = Picture1.Width / Screen.TwipsPerPixelX
    TmpIP.Filters(1).Properties("MaximumHeight") = Picture1.Height / Screen.TwipsPerPixelY
    BuildTree

End Sub

'****************************************************************************
'**                                                                        **
'** Subroutine.: wia_OnEvent                                               **
'**                                                                        **
'** Description: Process Windows Image Acquisition Events (WIA).  These    **
'**              occur when the camera connects, disconnects, or a picture **
'**              is taken.  The PC has no control over these events.  All  **
'**              events sent to this routine must be Registered - this is  **
'**              performed in the Form_Load Procedure (wia.RegisterEvent). **
'**                                                                        **
'****************************************************************************
Private Sub wia_OnEvent(ByVal EventID As String, ByVal DeviceID As String, ByVal ItemID As String)
    If InWIA Then Exit Sub
    InWIA = True
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    If EventID = wiaEventDeviceConnected Then   'The camera was connected, rebuild the property tree and set flag.
        Sleep 1500
        BuildTree                               'Retrieve camera properties
        IsCameraReady = True
        frmLogFile.LogText InfoMsg, "Camera Connected."
    ElseIf EventID = wiaEventDeviceDisconnected Then 'The camera was disconnected, rebuild the property tree and set flag.
        Sleep 1500
        BuildTree                               'This will simply clear the properties (no devices connected)
        frmLogFile.LogText InfoMsg, "Camera disconnected."
        IsCameraReady = False
    ElseIf EventID = wiaEventItemCreated Then   'A picture was taken, process it
        IsCameraReady = True
        IsPictureReady = True
    End If
    InWIA = False
    Exit Sub
ErrorHandler:
    frmLogFile.LogText ErrorMsg, "wia_OnEvent " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    InWIA = False
    Resume Next
End Sub

'****************************************************************************
'**                                                                        **
'** Subroutine.: BuildTree                                                 **
'**                                                                        **
'** Description: Displays camera information in tree-view control.         **
'**              This routine is only usefull for debugging new WIA Cameras**
'**              and has no real value during production.                  **
'**                                                                        **
'****************************************************************************
Private Sub BuildTree()
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    IsCameraReady = False                       'Assume camera is disconnected
    For Each di In wia.DeviceInfos              'For each device in the WIA Driver
        Set dev = di.Connect                    'Connect to device
        If Not dev Is Nothing Then              'If connected ok
            If dev.Type = CameraDeviceType Then
                CameraName = di.Properties("Name").value
                IsCameraReady = True            'Camera was found
                EnumerateDeviceProperties
            End If
        End If
    Next
    If IsCameraReady = False Then
        CameraName = "Not connected."
    End If
    Exit Sub
ErrorHandler:
    frmLogFile.LogText InfoMsg, "BuildTree " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Sub

Public Function GetImageInfo(Img As ImageFile) As String
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim s As String
    s = "Width = " & Img.Width & vbCrLf & _
        "Height = " & Img.Height & vbCrLf & _
        "Depth = " & Img.PixelDepth & vbCrLf & _
        "HorizontalResolution = " & Img.HorizontalResolution & vbCrLf & _
        "VerticalResolution = " & Img.VerticalResolution & vbCrLf & _
        "FrameCount = " & Img.FrameCount & vbCrLf
    If Img.IsIndexedPixelFormat Then
        s = s & "Pixel data contains palette indexes" & vbCrLf
    End If
    If Img.IsAlphaPixelFormat Then
        s = s & "Pixel data has alpha information" & vbCrLf
    End If
    If Img.IsExtendedPixelFormat Then
        s = s & "Pixel data has extended color information (16 bit/channel)" & vbCrLf
    End If
    If Img.IsAnimated Then
        s = s & "Image is animated" & vbCrLf
    End If
    If Img.Properties.Exists("40091") Then
        Set v = Img.Properties("40091").value
        s = s & "Title = " & v.String & vbCrLf
    End If
    If Img.Properties.Exists("40092") Then
        Set v = Img.Properties("40092").value
        s = s & "Comment = " & v.String & vbCrLf
    End If
    If Img.Properties.Exists("40093") Then
        Set v = Img.Properties("40093").value
        s = s & "Author = " & v.String & vbCrLf
    End If
    If Img.Properties.Exists("40094") Then
        Set v = Img.Properties("40094").value
        s = s & "Keywords = " & v.String & vbCrLf
    End If
    If Img.Properties.Exists("40095") Then
        Set v = Img.Properties("40095").value
        s = s & "Subject = " & v.String & vbCrLf
    End If
    GetImageInfo = s
    Exit Function
ErrorHandler:
    frmLogFile.LogText InfoMsg, "GetImageInfo " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Function

Public Function EnumerateDeviceProperties() As String
    If frmLogFile.UseErrorHandler Then On Error GoTo ErrorHandler
    Dim i As Integer, s As String
    s = ""
    For i = 1 To dev.Properties.Count
        s = dev.Properties(i).Name & "(" & dev.Properties(i).PropertyID & ") = "
        If dev.Properties(i).IsVector Then
            s = s & "[vector of data]"
        Else
            If dev.Properties(i).Type = StringPropertyType Then
                s = s & """" & dev.Properties(i).value & """"
            Else
                s = s & dev.Properties(i).value
            End If
        End If
        frmLogFile.LogText InfoMsg, "Camera: " & s
        s = s & vbCrLf
    Next
    EnumerateDeviceProperties = s
    Exit Function
ErrorHandler:
    frmLogFile.LogText InfoMsg, "EnumerateDeviceProperties " & "Error#" & Err.Number & ",DllErr#" & Err.LastDllError & "," & Err.Source
    Resume Next
End Function

Public Function PropType(id As WiaPropertyType) As String
    Select Case id
    Case BooleanPropertyType
        PropType = "Boolean"
    Case BytePropertyType
        PropType = "Byte"
    Case ClassIDPropertyType
        PropType = "Class ID"
    Case CurrencyPropertyType
        PropType = "Currency"
    Case DatePropertyType
        PropType = "Date"
    Case DoublePropertyType
        PropType = "Double"
    Case ErrorCodePropertyType
        PropType = "Error Code"
    Case FileTimePropertyType
        PropType = "File Time"
    Case HandlePropertyType
        PropType = "Handle"
    Case IntegerPropertyType
        PropType = "Integer"
    Case LargeIntegerPropertyType
        PropType = "Large Integer"
    Case LongPropertyType
        PropType = "Long"
    Case ObjectPropertyType
        PropType = "Object"
    Case SinglePropertyType
        PropType = "Single"
    Case StringPropertyType
        PropType = "String"
    Case UnsignedIntegerPropertyType
        PropType = "Unsigned Integer"
    Case UnsignedLargeIntegerPropertyType
        PropType = "Unsigned Large Integer"
    Case UnsignedLongPropertyType
        PropType = "Unsigned Long"
    Case VariantPropertyType
        PropType = "Variant"
    Case VectorOfBooleansPropertyType
        PropType = "Vector Of Booleans"
    Case VectorOfBytesPropertyType
        PropType = "Vector Of Bytes"
    Case VectorOfClassIDsPropertyType
        PropType = "Vector Of Class IDs"
    Case VectorOfCurrenciesPropertyType
        PropType = "Vector Of Currencies"
    Case VectorOfDatesPropertyType
        PropType = "Vector Of Dates"
    Case VectorOfDoublesPropertyType
        PropType = "Vector Of Doubles"
    Case VectorOfErrorCodesPropertyType
        PropType = "Vector Of Error Codes"
    Case VectorOfFileTimesPropertyType
        PropType = "Vector Of File Times"
    Case VectorOfIntegersPropertyType
        PropType = "Vector Of Integers"
    Case VectorOfLargeIntegersPropertyType
        PropType = "Vector Of Large Integers"
    Case VectorOfLongsPropertyType
        PropType = "Vector Of Longs"
    Case VectorOfSinglesPropertyType
        PropType = "Vector Of Singles"
    Case VectorOfStringsPropertyType
        PropType = "Vector Of Strings"
    Case VectorOfUnsignedIntegersPropertyType
        PropType = "Vector Of Unsigned Integers"
    Case VectorOfUnsignedLargeIntegersPropertyType
        PropType = "Vector Of Unsigned Large Integers"
    Case VectorOfUnsignedLongsPropertyType
        PropType = "Vector Of Unsigned Longs"
    Case VectorOfVariantsPropertyType
        PropType = "Vector Of Variants"
    Case Else
        PropType = "Unsupported"
    End Select
End Function

Public Sub EnumCommands(ByRef cmds As DeviceCommands)
    Dim cmd As DeviceCommand
    If cmds.Count > 0 Then
        For Each cmd In cmds
            'cmd.CommandID,cmd.Name,cmd.Description
        Next
    End If
End Sub

Public Sub EnumEvents(ByRef evts As DeviceEvents)
    Dim evt As DeviceEvent
    If evts.Count > 0 Then
        For Each evt In evts
            'evt.EventID,evt.Name,evt.Description
        Next
    End If
End Sub

Public Sub EnumFormats(ByRef fmts As Formats)
    Dim v As Variant
    If fmts.Count > 0 Then
        For Each v In fmts
            ' v
        Next
    End If
End Sub

Public Sub EnumProperties(ByRef props As Properties, ByRef ThumbnailData As Vector, ByRef ThumbnailWidth As Long, ByRef ThumbnailHeight As Long)
    Dim p As Property
    Dim i As Long
    Dim s As String
    Dim li As ListItem
    Dim v As Vector
    Dim pc As Integer

    On Error GoTo BadDriver
'    ListView1.ListItems.Clear
    pc = 0
    If props.Count > 0 Then
        For Each p In props
            'Set li = ListView1.ListItems.Add(, , p.Name)
            'Set li.Tag = p
            'li.SubItems(1) = p.PropertyID
            'li.SubItems(2) = PropType(p.Type)
            If Not p.IsVector Then
             '   li.SubItems(3) = p.Value
            Else
                s = ""
                Set v = p.value
                For i = 1 To v.Count
                    s = s & CStr(v(i))
                    If i < 50 Then
                        If i <> v.Count Then
                            s = s & ", "
                        End If
                    Else
                        s = s & ", ... (Size = " & v.Count & ")"
                        Exit For
                    End If
                Next
                li.SubItems(3) = s
            End If
            If p.Name = "Thumbnail Data" Then Set ThumbnailData = p.value
            If p.Name = "Thumbnail Width" Then ThumbnailWidth = p.value
            If p.Name = "Thumbnail Height" Then ThumbnailHeight = p.value
            Select Case p.SubType
            Case ListSubType
            '    li.SubItems(4) = "List"
            '    li.SubItems(5) = p.SubTypeDefault
                s = ""
                Set v = p.SubTypeValues
                For i = 1 To v.Count
                    s = s & CStr(v(i))
                    If i < 50 Then
                        If i <> v.Count Then
                            s = s & ", "
                        End If
                    Else
                        s = s & ", ... (Size = " & v.Count & ")"
                        Exit For
                    End If
                Next
                li.SubItems(6) = s
            Case FlagSubType
            '    li.SubItems(4) = "Flags"
            '    li.SubItems(5) = p.SubTypeDefault
                s = ""
                Set v = p.SubTypeValues
                For i = 1 To v.Count
                    s = s & CStr(v(i))
                    If i <> v.Count Then
                        s = s & ", "
                    End If
                Next
                li.SubItems(6) = s
            Case RangeSubType
            '    li.SubItems(4) = "Range"
            '    li.SubItems(5) = p.SubTypeDefault
            '    li.SubItems(6) = "Min = " & p.SubTypeMin & ", Max = " & p.SubTypeMax & ", Step = " & p.SubTypeStep
            Case Else
            '    li.SubItems(4) = "Unspecified"
            End Select
BadDriverResume:
            pc = pc + 1
        Next
    End If
    Exit Sub
BadDriver:
    If Err.Number = -2145320836 Then
        Err.Clear
        Resume BadDriverResume
    Else
        Resume
    End If
End Sub

Public Sub EnumChildren(ByRef itms As Items, ByRef nde As Node)
    Dim itm As Item
    Dim newNode As Node
    For Each itm In itms
        'Set newNode = TreeView1.Nodes.Add(nde.Index, tvwChild, , itm.Properties("Item Name").Value)
        'Set newNode.Tag = itm
        If itm.Items.Count > 0 Then EnumChildren itm.Items, newNode
    Next
End Sub

Public Function TransferPicture() As String
    TransferPicture = False
    Dim TmpItm As Item, f As Variant, UseRoot As Boolean
    Dim CurFolder As Long, CurItem As Long
    Picture1.Visible = False
    UseRoot = False
    For Each TmpItm In dev.Items
        f = TmpItm.Properties("Item Flags")
        If (f And ImageItemFlag) = ImageItemFlag Then
           'Storing items to root folder
           UseRoot = True
        End If
    Next
    If UseRoot Then
        CurItem = dev.Items.Count
        Set Img = dev.Items(CurItem).Transfer
        TransferPicture = dev.Items(CurItem).ItemID
    Else
        CurFolder = dev.Items.Count
        CurItem = dev.Items(CurFolder).Items.Count
        Set Img = dev.Items(CurFolder).Items(CurItem).Transfer
        TransferPicture = dev.Items(CurFolder).Items(CurItem).ItemID
    End If
    
End Function

Public Sub LoadPicture(PicName As String)
    Picture1.Visible = False
    Set Img = New ImageFile
    Img.LoadFile (PicName)
End Sub

Public Sub ApplyFilters(SubjectId As String, PkgString As String, CustString As String, Rot As Integer)
    IP.Filters(1).Properties("Value") = SubjectId
    IP.Filters(2).Properties("Value") = PkgString                  'Trim(Me.FormattedPackages.Text)
    IP.Filters(3).Properties("Value") = CustString               'Trim(Me.SummarySpread.MaxRows + 1)
    IP.Filters(4).Properties("Value") = Trim(frmStartDay.FolderName.Text)
    IP.Filters(5).Properties("RotationAngle") = Rot
    IP.Filters(6).Properties("FormatID").value = wiaFormatJPEG
    IP.Filters(6).Properties("Quality").value = 98
    Set Img = IP.Apply(Img)
    Set TmpImg = TmpIP.Apply(Img)
    Set Picture1.Picture = TmpImg.ARGBData.Picture(TmpImg.Width, TmpImg.Height)
    Picture1.Visible = True
End Sub

Public Sub UpdateProgress(Pct As Integer, Optional Stat As CounterTypes)
    DownloadProgressBar.value = Pct
    Select Case Pct
        Case 0
            MainTimer.StopTimer
            MainTimer.StartTimer True
            Picture1.Visible = False
            DownloadProgressBar.Visible = True
        Case 100
            DownloadProgressBar.Visible = False
            frmLogFile.Update Stat, MainTimer.ElapsedTime
            MainTimer.StopTimer
        Case Else
            frmLogFile.Update Stat, MainTimer.ElapsedTime
    End Select
    DoEvents
End Sub


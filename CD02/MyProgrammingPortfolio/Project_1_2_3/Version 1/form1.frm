VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   BeginProperty Font 
      Name            =   "Bookman Old Style"
      Size            =   9.75
      Charset         =   0
      Weight          =   300
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   8100
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Insall 123 Project"
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   3255
      Begin VB.Label Label2 
         Caption         =   "Project 1 2 3 will be installed on your Computer. Press forward to continue."
         Height          =   1095
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   525
      Index           =   1
      Left            =   3600
      Picture         =   "form1.frx":0000
      Top             =   4440
      Width           =   870
   End
   Begin VB.Image Image1 
      Height          =   525
      Index           =   0
      Left            =   4680
      Picture         =   "form1.frx":1852
      Top             =   4440
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcom to the week 1 2 3 project"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2093
      TabIndex        =   0
      Top             =   360
      Width           =   3915
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Long, lParam As Any) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
    ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Private Declare Function CreateIconIndirect Lib "user32" (icoinfo As ICONINFO) As Long

'Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lppictDesc As _
     pictDesc, riid As Guid, ByVal fown As Long, ipic As IPicture) As Long

'Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, _
     icoinfo As ICONINFO) As Long
'
'Private Type ICONINFO
'    fIcon As Long
'    xHotspot As Long
'    yHotspot As Long
'    hBMMask As Long
'    hBMColor As Long
'End Type
'
'Private Type Guid
'    Data1 As Long
'    Data2 As Integer
'    Data3 As Integer
'    Data4(7) As Byte
'End Type
'
'Private Type pictDesc
'    cbSizeofStruct As Long
'    picType As Long
'    hImage As Long
'    xExt As Long
'    yExt As Long
'End Type
'
'Private Type Rect
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'
'Private Type CharRange
'    firstChar As Long         ' First character of range (0 for start of doc)
'    lastChar As Long          ' Last character of range (-1 for end of doc)
'End Type
'
'Private Type FormatRange
'    hdc As Long               ' Actual DC to draw on
'    hdcTarget As Long         ' Target DC for determining text formatting
'    rectRegion As Rect        ' Region of the DC to draw to (in twips)
'    rectPage As Rect          ' Page size of the entire DC (in twips)
'    mCharRange As CharRange   ' Range of text to draw (see above user type)
'End Type
'
'
'Const WM_USER As Long = &H400
'Const EM_FORMATRANGE As Long = WM_USER + 57
'
'Const PICTYPE_BITMAP = 1
'Const PICTYPE_ICON = 3
'
'Const stdW = 32
'Const stdH = 32
'
'Dim iGuid As Guid
'Dim hdcMono
'Dim bmpMono
'Dim bmpMonoTemp
'Dim allowUndoFlag As Boolean
'Dim mresult
'Dim gcdg As Object
'
'
'
Private Sub Form_Load()
'     ' Create monochrome hDC and bitmap
    hdcMono = CreateCompatibleDC(hdc)
'    bmpMono = CreateCompatibleBitmap(hdcMono, stdW, stdH)
'    bmpMonoTemp = SelectObject(hdcMono, bmpMono)
'    With iGuid
'         .Data1 = &H20400
'         .Data4(0) = &HC0
'         .Data4(7) = &H46
'    End With
'
''    ExtractIconComposite 'picIcon
'    'BuildIcon 'picUndo
''    Set gcdg = 'CommonDialog1
'    allowUndoFlag = False
'     ' Set max no. of chars in rbtText, allowing possible spaces
'     ' in front in order to push typed text into the wanted line.
'     ' If scroll beyond full size of picIcon, lines would overlay
'     ' one another.
'    'rtbText.MaxLength = 18
'    ToggleVisible False
End Sub
'
'
'
'Private Sub ToggleVisible(OnOff As Boolean)
'    rtbText = ""
'    rtbText.Visible = OnOff
'    cmdTextFont.Visible = OnOff
'    cmdTextColor.Visible = OnOff
'    cmdClipboard.Visible = OnOff
'    CmdOverlayText.Visible = OnOff
'    picIcon.Visible = Not OnOff
'End Sub
'
'
'
'Private Sub cmdOpen_Click()
'    On Error GoTo errhandler
'    ToggleVisible False
'    Dim mfilespec As String
'    With gcdg
'        .FileName = ""
'        .CancelError = True
'        .Filter = "Icon files (*.ico)|*.ico|(*.*)|*.*|"
'        .FilterIndex = 1
'        .DefaultExt = "ico"
'        .Flags = cdlOFNFileMustExist
'    End With
'    gcdg.ShowOpen
'    mfilespec = gcdg.FileName
'    If mfilespec = "" Then
'        Exit Sub
'    End If
'
'    picTemp.AutoSize = True
'    picTemp.Picture = LoadPicture(mfilespec)
'    If picTemp.ScaleWidth <> 32 Or picTemp.ScaleHeight <> 32 Then
'         picTemp.Picture = LoadPicture()
'         picTemp.AutoSize = False
'         picTemp.Width = picTemp.Width - picTemp.ScaleWidth + 32
'         picTemp.Height = picTemp.Height - picTemp.ScaleHeight + 32
'         MsgBox "This example program accepts 32 x 32 icon only"
'         Exit Sub
'    End If
'    ExtractIconComposite picTemp
'    BuildIcon picIcon
'    doBackUp
'    allowUndoFlag = False
'    picTemp.AutoSize = False
'    picTemp.Picture = LoadPicture()
'    Exit Sub
'
'errhandler:
'    If Err <> 32755 Then
'         ToggleVisible False
'         ErrMsgProc "cmdOpen_click"
'    End If
'End Sub
'
'
'
'
'Private Sub cmdSave_Click()
'    On Error GoTo errhandler
'    ToggleVisible False
'    Dim mfilespec As String
'    With gcdg
'        .FileName = ""
'        .CancelError = True
'        .Filter = "Icon files (*.ico)|*.ico|(*.*)|*.*|"
'        .FilterIndex = 1
'        .DefaultExt = "ico"
'        .Flags = cdlOFNOverwritePrompt
'    End With
'    gcdg.ShowSave
'    mfilespec = gcdg.FileName
'    If mfilespec = "" Then
'        Exit Sub
'    End If
'
'    SavePicture picIcon.Picture, mfilespec
'    Exit Sub
'
'errhandler:
'    If Err <> 32755 Then
'         ToggleVisible False
'         ErrMsgProc "cmdSave_Click"
'    End If
'End Sub
'
'
'
'
'Private Sub Command1_Click()
'    ToggleVisible True
'    rtbText.SetFocus
'End Sub
'
'
'
'Private Sub doBackUp()
'    ExtractIconComposite picIcon
'    BuildIcon picUndo
'    allowUndoFlag = True
'End Sub
'
'
'
'Private Sub cmdUndo_Click()
'    If allowUndoFlag = False Then
'        If rtbText.Visible Then
'            rtbText = ""
'            ToggleVisible False
'        End If
'        Exit Sub
'    End If
'    ExtractIconComposite picUndo
'    BuildIcon picIcon
'    allowUndoFlag = False
'    ToggleVisible False
'End Sub
'
'
'
'Private Sub ExtractIconComposite(inPic As PictureBox)
'    On Error Resume Next
'    Dim ipic As IPicture
'    Dim icoinfo As ICONINFO
'    Dim pDesc As pictDesc
'    Dim hDCWork
'    Dim hBMOldWork
'    Dim hNewBM
'    Dim hBMOldMono
'    Dim w, h
'    w = inPic.ScaleWidth
'    h = inPic.ScaleHeight
'    GetIconInfo inPic.Picture, icoinfo
    hDCWork = CreateCompatibleDC(0)
'    hNewBM = CreateCompatibleBitmap(inPic.hdc, w, h)
'    hBMOldWork = SelectObject(hDCWork, hNewBM)
'    hBMOldMono = SelectObject(hdcMono, icoinfo.hBMMask)
'    BitBlt hDCWork, 0, 0, w, h, hdcMono, 0, 0, vbSrcCopy
'    SelectObject hdcMono, hBMOldMono
'    SelectObject hDCWork, hBMOldWork
'    With pDesc
'        .cbSizeofStruct = Len(pDesc)
'        .picType = PICTYPE_BITMAP
'        .hImage = hNewBM
'    End With
'    OleCreatePictureIndirect pDesc, iGuid, 1, ipic
'    picMask = ipic
'    Set ipic = Nothing
'
'    pDesc.hImage = icoinfo.hBMColor
'      ' Third parameter set to 1 (true) to let picture be destroyed automatically
'    OleCreatePictureIndirect pDesc, iGuid, 1, ipic
'    picImage = ipic
'    Set ipic = Nothing
'
'    DeleteObject icoinfo.hBMMask
'    DeleteDC hDCWork
'    Set hBMOldWork = Nothing
'    Set hBMOldMono = Nothing
'End Sub
'
'
'
'Private Sub BuildIcon(inPic As PictureBox)
'    On Error Resume Next
'    Dim hOldMonoBM
'    Dim hDCWork
'    Dim hBMOldWork
'    Dim hBMWork
'    Dim ipic As IPicture
'    Dim pDesc As pictDesc
'    Dim icoinfo As ICONINFO
'
'    BitBlt hdcMono, 0, 0, stdW, stdH, picMask.hdc, 0, 0, vbSrcCopy
'    SelectObject hdcMono, bmpMonoTemp
    hDCWork = CreateCompatibleDC(0)
'
'    With inPic
'        hBMWork = CreateCompatibleBitmap(inPic.hdc, stdW, stdH)
'    End With
'
'    hBMOldWork = SelectObject(hDCWork, hBMWork)
'    BitBlt hDCWork, 0, 0, stdW, stdH, picImage.hdc, 0, 0, vbSrcCopy
'    SelectObject hDCWork, hBMOldWork
'
'    With icoinfo
'        .fIcon = 1
'        .xHotspot = 16            ' Doesn't matter here
'        .yHotspot = 16
'        .hBMMask = bmpMono
'        .hBMColor = hBMWork
'    End With
'
'    With pDesc
'        .cbSizeofStruct = Len(pDesc)
'        .picType = PICTYPE_ICON
'        .hImage = CreateIconIndirect(icoinfo)
'    End With
'
'    OleCreatePictureIndirect pDesc, iGuid, 1, ipic
'
'    inPic.Picture = LoadPicture()
'    inPic = ipic
'    bmpMonoTemp = SelectObject(hdcMono, bmpMono)
'    DeleteObject icoinfo.hBMMask
'    DeleteDC hDCWork
'    Set hBMOldWork = Nothing
'End Sub
'
'
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Set gcdg = Nothing
'    SelectObject bmpMono, bmpMonoTemp
'    DeleteObject bmpMono
'    DeleteDC hdcMono
'End Sub
'
'
'
'Private Sub cmdTextFont_Click()
'    On Error GoTo errhandler
'    gcdg.CancelError = True
'    gcdg.Flags = cdlCFBoth
'       ' Display form's current font as default
'    gcdg.FontName = Screen.ActiveForm.FontName
'    gcdg.ShowFont
'    rtbText.SelStart = 0
'    rtbText.SelLength = Len(rtbText.Text)
'    rtbText.SelFontName = gcdg.FontName
'    rtbText.SelFontSize = gcdg.FontSize
'    If rtbText.Visible Then
'        rtbText.SetFocus
'    End If
'    Exit Sub
'errhandler:
'    If Err.Number <> 32755 Then
'        ErrMsgProc "mnuTextFont_click"
'    Else
'        If rtbText.Visible Then
'            rtbText.SetFocus
'        End If
'    End If
'End Sub
'
'
'
'Private Sub cmdTextColor_Click()
'    On Error GoTo errhandler
'    With gcdg
'        .CancelError = True
'        .Flags = cdlCFBoth
'        .Color = Screen.ActiveForm.ForeColor
'    End With
'    gcdg.ShowColor
'    rtbText.SelStart = 0
'    rtbText.SelLength = Len(rtbText.Text)
'    rtbText.SelColor = gcdg.Color
'    If rtbText.Visible Then
'        rtbText.SetFocus
'    End If
'    Exit Sub
'
'errhandler:
'    If Err.Number <> 32755 Then
'        ToggleVisible True
'        ErrMsgProc "mnuFileTextColor"
'    Else
'        If rtbText.Visible Then
'            rtbText.SetFocus
'        End If
'    End If
'End Sub
'
'
'
'Private Sub cmdClipboard_Click()
'    If Clipboard.GetFormat(vbCFText) = False Then
'         MsgBox "No text in clipboard"
'         If rtbText.Visible Then
'             rtbText.SetFocus
'         End If
'         Exit Sub
'    End If
'    rtbText.SelText = Clipboard.GetText()
'    rtbText.SetFocus
'End Sub
'
'
'
'
'Private Sub CmdOverlayText_Click()
'    On Error GoTo errhandler
'    If Len(rtbText.Text) = 0 Then
'         MsgBox "No text yet"
'         If rtbText.Visible = True Then
'              rtbText.SetFocus
'         End If
'         Exit Sub
'    End If
'    Screen.MousePointer = vbHourglass
'
'    Dim mFormatRange As FormatRange
'    Dim rectDrawTo As Rect
'    Dim rectPage As Rect
'    Dim TextLength As Long
'    Dim newStartPos As Long
'    Dim dumpaway As Long
'    Dim origBackColor, i, j
'
'    doBackUp
'
'    origBackColor = picTemp.BackColor
'    picTemp.BackColor = picIcon.BackColor
'    picTemp.Picture = LoadPicture()
'    picTemp.Picture = picIcon.Picture
'
'    rectPage.Left = 0
'    rectPage.Top = 0
'    rectPage.Right = (picIcon.ScaleWidth - 1) * Screen.TwipsPerPixelX
'    rectPage.Bottom = (picIcon.ScaleHeight - 1) * Screen.TwipsPerPixelY
'
'    rectDrawTo.Left = rectPage.Left
'    rectDrawTo.Top = rectPage.Top
'    rectDrawTo.Right = rectPage.Right
'    rectDrawTo.Bottom = rectPage.Bottom
'
'    mFormatRange.hdc = picTemp.hdc
'    mFormatRange.hdcTarget = picTemp.hdc
'
'    newStartPos = 0
'    mFormatRange.rectRegion = rectDrawTo
'    mFormatRange.rectPage = rectPage
'    mFormatRange.mCharRange.firstChar = newStartPos
'    mFormatRange.mCharRange.lastChar = -1
'
'    TextLength = Len(rtbText.Text)
'    Do
'        newStartPos = SendMessage(rtbText.hwnd, EM_FORMATRANGE, True, mFormatRange)
'        If newStartPos >= TextLength Then
'            Exit Do
'        End If
'        mFormatRange.mCharRange.firstChar = newStartPos
'        mFormatRange.hdc = picTemp.hdc
'        mFormatRange.hdcTarget = picTemp.hdc
'        DoEvents
'    Loop
'    mresult = SendMessage(rtbText.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
'
'    For j = 0 To stdW
'         For i = 0 To stdH
'              If picTemp.Point(j, i) <> picIcon.Point(j, i) Then
'                   picImage.PSet (j, i), picTemp.Point(j, i)
'                   picMask.PSet (j, i)
'              End If
'         Next i
'    Next j
'
'    BuildIcon picIcon
'
'    picTemp.BackColor = origBackColor
'    picTemp.Picture = LoadPicture()
'    ToggleVisible False
'    Screen.MousePointer = vbDefault
'    Exit Sub
'
'errhandler:
'    ToggleVisible False
'    Screen.MousePointer = vbDefault
'    ErrMsgProc "cmdOverlaying_click"
'End Sub
'
'
'
'
'Sub ErrMsgProc(mMsg As String)
'    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
'End Sub
'
'
'
'
'
'

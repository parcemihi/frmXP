VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlClock 
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2130
   PropertyPages   =   "ctlClock.ctx":0000
   ScaleHeight     =   360
   ScaleWidth      =   2130
   ToolboxBitmap   =   "ctlClock.ctx":000E
   Begin VB.Timer tmrSec 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   930
      Top             =   510
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   600
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlClock.ctx":0108
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlClock.ctx":021C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlClock.ctx":0330
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlClock.ctx":0444
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlClock.ctx":0558
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlClock.ctx":066C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlClock.ctx":0780
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlClock.ctx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlClock.ctx":09A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlClock.ctx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlClock.ctx":0BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlClock.ctx":0CE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picClock 
      AutoSize        =   -1  'True
      Height          =   345
      Index           =   7
      Left            =   1830
      ScaleHeight     =   285
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picClock 
      AutoSize        =   -1  'True
      Height          =   345
      Index           =   6
      Left            =   1575
      ScaleHeight     =   285
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picClock 
      AutoSize        =   -1  'True
      Height          =   345
      Index           =   5
      Left            =   1320
      ScaleHeight     =   285
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picClock 
      AutoSize        =   -1  'True
      Height          =   345
      Index           =   4
      Left            =   1065
      ScaleHeight     =   285
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picClock 
      AutoSize        =   -1  'True
      Height          =   345
      Index           =   3
      Left            =   795
      ScaleHeight     =   285
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picClock 
      AutoSize        =   -1  'True
      Height          =   345
      Index           =   2
      Left            =   540
      ScaleHeight     =   285
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picClock 
      AutoSize        =   -1  'True
      Height          =   345
      Index           =   1
      Left            =   285
      ScaleHeight     =   285
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picClock 
      AutoSize        =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   30
      ScaleHeight     =   285
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "ctlClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Public Event Alarm(ThisTime As Date)

'Default Property Values:
Const m_def_SetAlarmHr = 0
Const m_def_SetAlarmMin = 0

'Property Variables:
Dim m_SetAlarmHr As Integer
Dim m_SetAlarmMin As Integer
Dim bBlink As Boolean

Private Sub tmrSec_Timer()
    ShowTime
End Sub

Private Sub UserControl_Initialize()
    Dim i As Integer
    For i = 0 To 7
        picClock(i).Picture = imgList.ListImages(1).Picture
    Next i
    picClock(2).Picture = imgList.ListImages(11).Picture
    picClock(5).Picture = imgList.ListImages(11).Picture
          
    m_SetAlarmHr = m_def_SetAlarmHr
    m_SetAlarmMin = m_def_SetAlarmMin
End Sub

Private Sub UserControl_InitProperties()
    ShowTime
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tmrSec,tmrSec,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "General"
    Enabled = tmrSec.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    tmrSec.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    tmrSec.Enabled = PropBag.ReadProperty("Enabled", False)
    m_SetAlarmHr = PropBag.ReadProperty("SetAlarmHr", m_def_SetAlarmHr)
    m_SetAlarmMin = PropBag.ReadProperty("SetAlarmMin", m_def_SetAlarmMin)
    
    ShowTime
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 330
    UserControl.Width = 2140
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", tmrSec.Enabled, False)
    Call PropBag.WriteProperty("SetAlarmHr", m_SetAlarmHr, m_def_SetAlarmHr)
    Call PropBag.WriteProperty("SetAlarmMin", m_SetAlarmMin, m_def_SetAlarmMin)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get SetAlarmHr() As Integer
Attribute SetAlarmHr.VB_ProcData.VB_Invoke_Property = "General"
    SetAlarmHr = m_SetAlarmHr
End Property

Public Property Let SetAlarmHr(ByVal New_SetAlarmHr As Integer)
    m_SetAlarmHr = New_SetAlarmHr
    PropertyChanged "SetAlarmHr"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get SetAlarmMin() As Integer
Attribute SetAlarmMin.VB_ProcData.VB_Invoke_Property = "General"
    SetAlarmMin = m_SetAlarmMin
End Property

Public Property Let SetAlarmMin(ByVal New_SetAlarmMin As Integer)
    m_SetAlarmMin = New_SetAlarmMin
    PropertyChanged "SetAlarmMin"
End Property

Public Sub About()
Attribute About.VB_UserMemId = -552
    frmAbout.Show vbModal
End Sub

Private Sub ShowTime()
    Dim hr As Integer
    Dim min As Integer
    Dim sec As Integer
    
    bBlink = Not (bBlink)
    
    hr = Hour(Now)
    min = Minute(Now)
    sec = Second(Now)
    
    picClock(0).Picture = imgList.ListImages(hr \ 10 + 1).Picture
    picClock(1).Picture = imgList.ListImages(hr Mod 10 + 1).Picture
    picClock(3).Picture = imgList.ListImages(min \ 10 + 1).Picture
    picClock(4).Picture = imgList.ListImages(min Mod 10 + 1).Picture
    picClock(6).Picture = imgList.ListImages(sec \ 10 + 1).Picture
    picClock(7).Picture = imgList.ListImages(sec Mod 10 + 1).Picture
    
    If bBlink Then
        picClock(2).Picture = imgList.ListImages(12).Picture
        picClock(5).Picture = imgList.ListImages(12).Picture
    Else
       picClock(2).Picture = imgList.ListImages(11).Picture
       picClock(5).Picture = imgList.ListImages(11).Picture
    End If
    
    If hr = m_SetAlarmHr And min = m_SetAlarmMin And sec = 0 Then
        'MsgBox "Now it is " & m_SetAlarmHr & " hour and " & m_SetAlarmMin & " minute", vbInformation
        RaiseEvent Alarm(Now)
    End If
End Sub

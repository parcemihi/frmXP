VERSION 5.00
Begin VB.UserControl Scroller 
   AutoRedraw      =   -1  'True
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   ScaleHeight     =   600
   ScaleWidth      =   6120
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Scroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_Caption = "Scroller"
'Property Variables:
Dim m_Caption As Variant
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    RefreshCaption
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    RefreshCaption
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    RefreshCaption
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    RefreshCaption
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub Timer1_Timer()
    Static intPosition As Integer
    Static strMarqueeText As String
    If Not Ambient.UserMode Then
        Exit Sub
    End If
    If Len(strMarqueeText) = 0 Then
        strMarqueeText = Caption & Space(Len(Caption))
        Do While TextWidth(strMarqueeText) < Width * 2
            strMarqueeText = strMarqueeText & " "
        Loop
        intPosition = 1
    End If
    intPosition = intPosition + ((TextWidth("W") \ TextWidth(Mid(strMarqueeText, intPosition, 1))))
    If intPosition > Len(strMarqueeText) Then
        intPosition = 1
    End If
    m_Caption = Mid(strMarqueeText, intPosition) & Left(strMarqueeText, intPosition)
    RefreshCaption
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Timer1,Timer1,-1,Interval
Public Property Get Interval() As Long
Attribute Interval.VB_Description = "Returns/sets the number of milliseconds between calls to a Timer control's Timer event."
    Interval = Timer1.Interval
End Property

Public Property Let Interval(ByVal New_Interval As Long)
    Timer1.Interval() = New_Interval
    PropertyChanged "Interval"
End Property

Public Property Get Caption() As Variant
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As Variant)
    m_Caption = New_Caption
    RefreshCaption
    PropertyChanged "Caption"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    BackColor = vbCyan
    ForeColor = vbBlack
    'BorderStyle = vbSingle
    UserControl.FontName = "Courier New"
    UserControl.FontBold = True
    m_Caption = m_def_Caption
    Set Font = Ambient.Font
    m_Caption = m_def_Caption
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Timer1.Interval = PropBag.ReadProperty("Interval", 100)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
End Sub

Private Sub UserControl_Resize()
    Static sngLastHeight As Single
    Dim sngCurrHeight As Single, sngFont As Single
    If sngLastHeight = 0 Then
        sngLastHeight = Height
    End If
    sngFont = UserControl.FontSize
    sngCurrHeight = Height
    If (sngLastHeight <= sngCurrHeight) And TextHeight("A") < sngCurrHeight Then
        Do While TextHeight("A") <= sngCurrHeight
            sngFont = sngFont + 0.25
            UserControl.FontSize = sngFont
        Loop
    Else
        Do While (TextHeight("A") >= sngCurrHeight) And (sngFont > 1.25)
            sngFont = sngFont - 0.25
            UserControl.FontSize = sngFont
        Loop
    End If
    sngLastHeight = Height
    RefreshCaption
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Interval", Timer1.Interval, 100)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
End Sub


Public Sub RefreshCaption()
    UserControl.Cls
    UserControl.Print m_Caption
End Sub

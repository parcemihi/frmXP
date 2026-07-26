Attribute VB_Name = "API_BAS"
Option Explicit
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal aint As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Const GW_MAX = 5
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Function FindWindowPartial(ByVal TitlePart As String) As Long
'Borrowed from VBPJ and ported to 32-bit
     Dim hWndTmp As Long
     Dim nRet As Integer
     Dim TitleTmp As String
     'We alter the title to compare it case-insensitively.
     TitlePart = UCase$(TitlePart)
     'First find all the open windows so we can
     'loop through them and find the right one.
     hWndTmp = FindWindow(0&, 0&)
     Do Until hWndTmp = 0
          TitleTmp = Space$(256)
          nRet = GetWindowText(hWndTmp, TitleTmp, Len(TitleTmp))
          If nRet Then
               'Let's prepare to compare
               TitleTmp = UCase$(VBA.Left$(TitleTmp, nRet))
               'Now we see if the window we chased down actually
               'has the caption we want.
               If InStr(TitleTmp, TitlePart) Then
                    FindWindowPartial = hWndTmp
                    Exit Do
               End If
          End If
          hWndTmp = GetWindow(hWndTmp, GW_HWNDNEXT)
     Loop
End Function




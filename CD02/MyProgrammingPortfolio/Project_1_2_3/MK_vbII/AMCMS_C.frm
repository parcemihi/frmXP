VERSION 5.00
Begin VB.Form frmAtlanticContact 
   Caption         =   "Atlantic Marketing - Contact Management System"
   ClientHeight    =   4410
   ClientLeft      =   1620
   ClientTop       =   2085
   ClientWidth     =   5865
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4410
   ScaleWidth      =   5865
   Begin VB.Data datContact 
      Caption         =   "Contact Database"
      Connect         =   "Access"
      DatabaseName    =   "C:\vbII\AMCMS.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "qryContact"
      Top             =   3960
      Width           =   5655
   End
   Begin VB.TextBox txtClientID 
      DataField       =   "fldClientID"
      DataSource      =   "datContact"
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtEstimatedSales 
      DataField       =   "fldEstimatedSales"
      DataSource      =   "datContact"
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtNotes 
      DataField       =   "fldNotes"
      DataSource      =   "datContact"
      Height          =   1335
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2400
      Width           =   4695
   End
   Begin VB.TextBox txtDateAdded 
      DataField       =   "fldDateAdded"
      DataSource      =   "datContact"
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtTelephone 
      DataField       =   "fldTelephone"
      DataSource      =   "datContact"
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtFirstName 
      DataField       =   "fldFirstName"
      DataSource      =   "datContact"
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtLastName 
      DataField       =   "fldLastName"
      DataSource      =   "datContact"
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Client ID: "
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Estimated Annual Sales:  "
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Telephone Number:  "
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Notes:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Added:  "
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "First Name:  "
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Name:  "
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuEditRecord 
         Caption         =   "&Edit Record"
      End
      Begin VB.Menu mnuEditUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuEditRefresh 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu mnuFind 
      Caption         =   "F&ind"
      Begin VB.Menu mnuFindNavigate 
         Caption         =   "&Navigate"
         Begin VB.Menu mnuFindNavigateFirst 
            Caption         =   "&First"
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuFindNavigateLast 
            Caption         =   "&Last"
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuFindNavigatePrevious 
            Caption         =   "&Previous"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuFindNavigateNext 
            Caption         =   "&Next"
            Shortcut        =   ^N
         End
      End
      Begin VB.Menu mnuFindSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindLastName 
         Caption         =   "&Last Name..."
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmAtlanticContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrstContact As Recordset


Private Sub datContact_Validate(Action As Integer, Save As Integer)
    Dim pstrmessage As String
    Dim pintReturn As Integer
    If Action = vbDataActionUpdate Then
        If txtLastName.Text = "" Or _
            txtFirstName.Text = "" Then
            pstrmessage = _
                "You must enter both a first and last name."
            Action = vbDataActionCancel
        End If
        If IsDate(txtDateAdded.Text) = False Then
            pstrmessage = pstrmessage & Chr(vbKeyReturn) & _
                "The date " & txtDateAdded.Text & _
                " is not a date."
            Action = vbDataActionCancel
        End If
        If IsNumeric(txtEstimatedSales.Text) = False Then
            pstrmessage = pstrmessage & Chr(vbKeyReturn) & _
            txtEstimatedSales.Text & " is not a number."
            Action = vbDataActionCancel
        End If
        If Action = vbDataActionCancel Then
            pintReturn = _
                MsgBox(pstrmessage, vbOKOnly, "Input error")
        End If
    End If
End Sub

Private Sub mnuEditAdd_Click()
    mrstContact.AddNew
    mnuEditAdd.Enabled = False
    mnuEditRecord.Enabled = False
    mnuEditDelete.Enabled = False
    mnuEditUpdate.Enabled = True
    mnuEditRefresh.Enabled = False
    mnuFind.Enabled = False
End Sub

Private Sub mnuEditDelete_Click()
    Dim pstrmessage As String
    On Error GoTo mnuEditDelete_Error
    mrstContact.Delete
    mrstContact.MoveNext
    Exit Sub
    
mnuEditDelete_Error:
    pstrmessage = "Cannot perform operation #" _
        & Err.Number & ":" & Err.Description
    MsgBox pstrmessage, vbOKOnly & vbExclamation, _
        "Atlantic Contact - Database Error"
        mrstContact.MoveFirst
    
End Sub

Private Sub mnuEditRecord_Click()
    Dim pstrmessage As String
    On Error GoTo mnuEditrecord_Error
    mrstContact.Edit
    mnuEditAdd.Enabled = False
    mnuEditRecord.Enabled = False
    mnuEditDelete.Enabled = False
    mnuEditUpdate.Enabled = True
    mnuEditRefresh.Enabled = False
    mnuFind.Enabled = False
    Exit Sub
mnuEditrecord_Error:
    pstrmessage = "Cannot perform operation #" _
        & Err.Number & ":" & Err.Description
    MsgBox pstrmessage, vbExclamation & vbOKOnly, _
        "Atlantic Contact - Database Error"

End Sub

Private Sub mnuEditRefresh_Click()
    mrstContact.Requery
    mrstContact.MoveFirst
End Sub

Private Sub mnuEditUpdate_Click()
    Dim pstrmessage As String
    On Error GoTo mnuEditUpdate_Error
    mrstContact.Update
    mnuEditAdd.Enabled = True
    mnuEditRecord.Enabled = True
    mnuEditDelete.Enabled = True
    mnuEditUpdate.Enabled = False
    mnuEditRefresh.Enabled = True
    mnuFind.Enabled = True
    Exit Sub
mnuEditUpdate_Error:
    pstrmessage = "Cannot perform operation" & _
        Chr(vbKeyReturn) & "Error #" & _
        Err.Number & ":" & Err.Description
    MsgBox pstrmessage, vbExclamation + vbOKOnly, _
        "Atlantic Contact - Database Error"
    mrstContact.CancelUpdate
End Sub

Private Sub mnuFileExit_Click()
    Unload frmAtlanticContact
    
End Sub

Private Sub mnuFilePrint_Click()
    PrintForm
End Sub

Private Sub mnuFindLastName_Click()
    Dim pstrLastName As String
    pstrLastName = InputBox("Enter the Last Name", "Find")
    If pstrLastName <> "" Then
        mrstContact.FindFirst "fldLastName = " & "'" & _
            pstrLastName & "'"
        If mrstContact.NoMatch = True Then
            MsgBox "Cannot find " & pstrLastName, _
                vbInformation & vbOK, "Altlantic Marketing"
        End If
    End If
End Sub

Private Sub mnuFindNavigateFirst_Click()
    mrstContact.MoveFirst
    mnuFindNavigateFirst.Enabled = False
    mnuFindNavigateLast.Enabled = True
    mnuFindNavigatePrevious.Enabled = False
    mnuFindNavigateNext.Enabled = True
End Sub

Private Sub mnuFindNavigateLast_Click()
    mrstContact.MoveLast
    mnuFindNavigateFirst.Enabled = True
    mnuFindNavigateLast.Enabled = False
    mnuFindNavigatePrevious.Enabled = True
    mnuFindNavigateNext.Enabled = False
End Sub

Private Sub mnuFindNavigateNext_Click()
    mrstContact.MoveNext
    If mrstContact.EOF Then
        mrstContact.MoveLast
        mnuFindNavigateFirst.Enabled = True
        mnuFindNavigateLast.Enabled = False
        mnuFindNavigatePrevious.Enabled = True
        mnuFindNavigateNext.Enabled = False
    Else
        mnuFindNavigateFirst.Enabled = True
        mnuFindNavigateLast.Enabled = True
        mnuFindNavigatePrevious.Enabled = True
        mnuFindNavigateNext.Enabled = True
    End If
End Sub

Private Sub mnuFindNavigatePrevious_Click()
    mrstContact.MovePrevious
    If mrstContact.BOF Then
        mrstContact.MoveFirst
        mnuFindNavigateFirst.Enabled = False
        mnuFindNavigateLast.Enabled = True
        mnuFindNavigatePrevious.Enabled = False
        mnuFindNavigateNext.Enabled = True
    Else
        mnuFindNavigateFirst.Enabled = True
        mnuFindNavigateLast.Enabled = True
        mnuFindNavigatePrevious.Enabled = True
        mnuFindNavigateNext.Enabled = True
    End If
End Sub
Private Sub datContact_Reposition()
    Static pblnFirst As Boolean
    If pblnFirst = False Then
        Set mrstContact = datContact.Recordset
        pblnFirst = True
    End If
    
End Sub


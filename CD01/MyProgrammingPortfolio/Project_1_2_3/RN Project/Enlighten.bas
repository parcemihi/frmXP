Attribute VB_Name = "Enlighten"
Option Explicit

'// Enlighten is the instructional portion of this program.
'// Within the *.enl file, configuration settings may be
'// modified. Each line of the *.enl is constructed similar
'// to that of an *.ini file. That is, comments are struck
'// out with the semi-colon and instructional subsets are
'// preceded by a command header encapsulated in square brackets.

Dim Enlighten As String
Global NumberOfRecords As Integer

Public Function ENL_Settings()

Open App.Path & "\lib.enl" For Input As #1
Do Until EOF(1)
    Input #1, Enlighten
    
    
    Select Case Enlighten
    
    Case "[Records]"
    
    
    
    
    Case "[Settings]"
    
    
    
    
    
    
    
    Case ";"
    
    Case Else
    
    End Select
    
    
Loop
Close #1
End Function

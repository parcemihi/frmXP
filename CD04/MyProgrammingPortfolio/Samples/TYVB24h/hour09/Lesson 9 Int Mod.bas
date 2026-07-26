Attribute VB_Name = "Module1"
Public Function ErrorCheck() As Integer
   ' Error-checking for the form
   If Val(frmInterest.txtRate.Text) <= 0 Then
      intPress = MsgBox("Enter a value for the rate", vbCritical)
      frmInterest.txtRate.SetFocus
      ErrorCheck = 1
      Exit Function
   End If
   If Val(frmInterest.txtTerm.Text) <= 0 Then
      intPress = MsgBox("Enter a value for the term", vbCritical)
      frmInterest.txtTerm.SetFocus
      ErrorCheck = 1
      Exit Function
   End If
   If Val(frmInterest.txtInvest.Text) <= 0 Then
      intPress = MsgBox("Enter a value for the investment", vbCritical)
      frmInterest.txtInvest.SetFocus
      ErrorCheck = 1
      Exit Function
   End If
   ' No error occurred if execution gets here
   ErrorCheck = 0
End Function


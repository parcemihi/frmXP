Attribute VB_Name = "Module1"
Public Function ErrorCheck() As Integer
   ' Error-checking for the form
   If Val(frmInterest.txtRate.Text) <= 0 Then
      ' No do loops around these but they might help some
      frmInterest.txtRate.Text = InputBox("Enter a correct interest rate", "Get rate")
      frmInterest.txtRate.SetFocus
      ErrorCheck = 1
      Exit Function
   End If
   If Val(frmInterest.txtTerm.Text) <= 0 Then
      frmInterest.txtTerm.Text = InputBox("Enter a correct term", "Get term")
      frmInterest.txtTerm.SetFocus
      ErrorCheck = 1
      Exit Function
   End If
   If Val(frmInterest.txtInvest.Text) <= 0 Then
      frmInterest.txtInvest.Text = InputBox("Enter a correct investment", "Get investment")
      frmInterest.txtInvest.SetFocus
      ErrorCheck = 1
      Exit Function
   End If
   ' No error occurred if execution gets here
   ErrorCheck = 0
End Function


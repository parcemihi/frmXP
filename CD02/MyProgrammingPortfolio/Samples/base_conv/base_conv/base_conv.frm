VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Base (Radix) Conversion Function"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4080
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "2"
      Top             =   540
      Width           =   555
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1260
      Width           =   8655
   End
   Begin VB.CommandButton ConvertButton 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Convert"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5340
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   60
      MaxLength       =   28
      TabIndex        =   0
      Text            =   "9999999999999999999999999999"
      Top             =   540
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Base to Convert Into:   2 to 36"
      Height          =   435
      Left            =   3540
      TabIndex        =   6
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label Label3 
      Caption         =   "The value of decimal integer  A  expressed in base B is:"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   1020
      Width           =   4995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Decimal Integer Value To Be Converted:"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   300
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Program written by Jay Tanner
' E-Mail: Jay_Tanner@Bigfoot.com
'
' This simple program converts a decimal integer (A) into
' its base (B) equivalent.
'
' The decimal integer value may be up to 28 digits long and
' the base (radix) may range from 2 to 36.

  Private Sub ConvertButton_Click()

  Dim A, B, C As String

  A = Trim(Text1.Text)       ' Read value of decimal integer
  B = Trim(Text2.Text)       ' Read value of base (radix)
  C = Base_B_EquivOf_A(A, B) ' Convert decimal value (A) into base (B)

  Text3.Text = C

  End Sub

  Public Function Base_B_EquivOf_A(Dec_A, Base_B)
' Compute (Base_B) equivalent of (Dec_A) value with
' up to 28 digits

  Dim DecVal, Accum, Radix, Qi As String
  Dim Q, R
  Dim i As Integer
    
' Define valid digit symbols up to base 36.  For the
' letter "O", the lower case letter "o" was used to avoid
' confusion with the digit zero "0".  All of the other
' letters are rendered in upper case.
  Dim Digits As String
      Digits = "0123456789ABCDEFGHIJKLMNoPQRSTUVWXYZ"
    
' Check for non-numeric decimal argument
  If IsNumeric(Dec_A) = False Then
     Base_B_EquivOf_A = "ERROR: Argument is not an integer."
     Exit Function
  End If
    
' Check if either argument has a zero value
  If Val(Dec_A) = 0 Then Dec_A = 0
     DecVal = CDec(Dec_A)
  If Val(Base_B) = 0 Then Base_B = 0
     Radix = CDec(Base_B)
                  
' Check for non-integer decimal argument
  If IsNumeric(DecVal) = True Then
     If InStr(DecVal, ".") > 0 Then
        Base_B_EquivOf_A = "ERROR: Argument is not an integer."
        Exit Function
     End If
  End If
     
' Check for valid base (radix) argument
  If Base_B < 2 Or Base_B > 36 Then
     Base_B_EquivOf_A = "ERROR: Base (Radix) must be in the range from 2 to 36."
     Exit Function
  Else
   
  End If

   
' Compute and accumulate the digits of the (Base_B)
' equivalent of (Dec_A) one at a time.
        Q = 1
  While Q > 0

      Q = DecVal / Radix
     Qi = Trim(Q)
      i = InStr(Qi, ".")
   If i > 0 Then Qi = Left(Qi, i - 1)
      R = DecVal - Radix * CDec(Qi)
  Accum = Mid(Digits, R + 1, 1) & Accum
 DecVal = CDec(Qi)
   If Val(Qi) = 0 Then Q = 0 ' Check if done yet
   
  Wend

' Return the computed (Base_B) equivalent of (Dec_A)
  Base_B_EquivOf_A = Accum

  End Function



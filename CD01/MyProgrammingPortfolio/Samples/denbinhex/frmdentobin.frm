VERSION 5.00
Begin VB.Form frmdentobin 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3465
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3285
   Icon            =   "frmdentobin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   3285
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraascii 
      Caption         =   "Ascii"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   3015
      Begin VB.TextBox txtascii 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame frahex 
      Caption         =   "Hex"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   3015
      Begin VB.TextBox txthex 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame frabin 
      Caption         =   "Binary"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3015
      Begin VB.TextBox txtbin 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame fraden 
      Caption         =   "Denary"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3015
      Begin VB.TextBox txtden 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmdentobin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' I produced this when I started doing some machine code and assembaly language
' programming, when knowing hex was useful.
' Like the about box says i started work on a binary calculator to deal with
' 2s complement, binary fractions, binary exponential etc but unfortunatly that
' has been put on hold due to a general excess of work.

' If you use this code for anything than e-mail me; grredo@hotmail.com

Private integer1 As Integer
Private integer2 As Integer
Private integer3 As Integer
Private string1 As String
Private string2 As String
Private string3 As String
Private Sub Form_Load()
' keypreview allows the program to check what is being typed before it is displayed
' you'll see its use later
KeyPreview = True
End Sub
Private Sub mnuexit_click()
' unloads the about form then itself
Unload frmAbout
Unload Me
End Sub
Private Sub mnuabout_click()
' when the menu button about is clicked the main form is disabled and the about form
' is shown
frmdentobin.Enabled = False
frmAbout.Show
End Sub
Private Function pad(ByVal binstring As String) As String

' the variables to be used are reset

integer1 = 0
integer2 = 0

' this function will pad out a string with 0s to make it into a byte

integer2 = 8 - Len(binstring)  'the eight in this line is seting the length of the string to be produced
For integer1 = 1 To integer2
    binstring = "0" & binstring     '0 is added to the front of the string
Next integer1
pad = binstring
End Function
Private Sub txtbin_keypress(keyascii As Integer)

' if the length of the text = 8 and the key pressed isn't enter
' backspace or escape then ascii code sent to the txt box is 0 or no charecter

If Len(txtbin.Text) = 8 And keyascii <> 27 And keyascii <> 8 And keyascii <> 13 Then
    keyascii = 0
    Exit Sub
End If

' if the length of the text is 0 and enter is pressed then
' ascii code sent to the txt box is 0 or no charecter

If Len(txtbin.Text) = 0 And keyascii = 13 Then
    keyascii = 0
End If

' this is where keypreviews use is seen, only the ascii codes below can be entered, if another
' key is pressed then the ascii code sent to the txt box is 0 or no charecter
If keyascii = 48 Or keyascii = 49 Or keyascii = 13 Or keyascii = 27 Or keyascii = 8 Then
Else: keyascii = 0
End If

' if enter is pressed then functions are called to calculate the values for the other
' base types

If keyascii = 13 Then
    txtbin.Text = pad(txtbin.Text)
    txtden.Text = calcbin2den(txtbin.Text)
    txtascii.Text = calcden2asci(txtden.Text)
    txthex.Text = calcden2hex(txtden.Text)
End If

' if keyascii = 27, the escape key, and there is no text in the box then the program is
' ended

If keyascii = 27 And Len(txtbin.Text) = 0 Then End

' ifkeyascii = 27, the escape key, and there is text in the box then the text boxes are cleared

If keyascii = 27 And Len(txtbin.Text) <> 0 Then
    txtden.Text = ""
    txtbin.Text = ""
    txthex.Text = ""
    txtascii.Text = ""
End If
End Sub
Private Sub txtascii_keypress(keyascii As Integer)

' if the length of the text is 0 and enter is pressed then
' ascii code sent to the txt box is 0 or no charecter

If Len(txtascii.Text) = 0 And keyascii = 13 Then
    keyascii = 0
End If

' if the length of the text = 1 and the key pressed isn't enter
' backspace or escape then ascii code sent to the txt box is 0 or no charecter

If Len(txtascii.Text) = 1 And keyascii <> 13 And keyascii <> 27 And keyascii <> 8 Then
    keyascii = 0
    Exit Sub
End If

' if enter is pressed then functions are called to calculate the values for the other
' base types

If keyascii = 13 Then
    txtden.Text = calcasci2den(txtascii.Text)
    txtbin.Text = calcden2bin(txtden.Text)
    txthex.Text = calcden2hex(txtden.Text)
End If

' if keyascii = 27, the escape key, and there is no text in the box then the program is
' ended

If keyascii = 27 And Len(txtascii.Text) = 0 Then End

' ifkeyascii = 27, the escape key, and there is text in the box then the text boxes are cleared

If keyascii = 27 And Len(txtascii.Text) <> 0 Then
    txtden.Text = ""
    txtbin.Text = ""
    txthex.Text = ""
    txtascii.Text = ""
End If
End Sub
Private Sub txthex_keypress(keyascii As Integer)

' if the length of the data in the text box is 2 and a key other than escape, backspace
' or enter is pressed then ascii code sent to the txt box is 0 or no charecter

If Len(txthex.Text) = 2 And keyascii <> 27 And keyascii <> 8 And keyascii <> 13 Then
    keyascii = 0
    Exit Sub
End If

' if the length of the text is 0 and enter is pressed then
' ascii code sent to the txt box is 0 or no charecter

If Len(txthex.Text) = 0 And keyascii = 13 Then
    keyascii = 0
    Exit Sub
End If

' because of the range of charecters that can be used in an hexidecimal number
' a select case is used to filter out all unwanted ascii codes

Select Case keyascii
Case 0 To 7
    keyascii = 0
    Exit Sub
Case 9 To 12
    keyascii = 0
    Exit Sub
Case 14 To 26
    keyascii = 0
    Exit Sub
Case 28 To 46
    keyascii = 0
    Exit Sub
Case 59 To 96
    keyascii = 0
    Exit Sub
Case 103 To 255
    keyascii = 0
    Exit Sub
End Select

' if enter is pressed then functions are called to calculate the values for the other
' base types

If keyascii = 13 Then
    txtden.Text = calchex2den(txthex.Text)
    txtbin.Text = calcden2bin(txtden.Text)
    txtascii.Text = calcden2asci(txtden.Text)
End If

' if keyascii = 27, the escape key, and there is no text in the box then the program is
' ended

If keyascii = 27 And Len(txthex.Text) = 0 Then End

' ifkeyascii = 27, the escape key, and there is text in the box then the text boxes are cleared

If keyascii = 27 And Len(txthex.Text) <> 0 Then
    txtden.Text = ""
    txtbin.Text = ""
    txthex.Text = ""
    txtascii.Text = ""
End If
End Sub
Private Sub txtden_keypress(keyascii As Integer)

' if the length of the text is 0 and enter is pressed then
' ascii code sent to the txt box is 0 or no charecter

If Len(txtden.Text) = 0 And keyascii = 13 Then
    keyascii = 0
    Exit Sub
End If

' if the value of the data in the textbox is greater than 255 and the
' enter key is pressed then then ascii code sent to the txt box is 0 or no charecter

If Val(txtden.Text) > 255 And keyascii = 13 Then
    keyascii = 0
    Exit Sub
End If

' if the length of the data in the text box is 3 and a key other than escape, backspace
' or enter is pressed then ascii code sent to the txt box is 0 or no charecter

If Len(txtden.Text) = 3 And keyascii <> 27 And keyascii <> 8 And keyascii <> 13 Then
    keyascii = 0
    Exit Sub
End If

' if the charecter asscoiated with the ascii code isn't numeric then
' if the ascii code isn't enter, escape or backspace then ascii code
' sent to the txt box is 0 or no charecter

If IsNumeric(Chr(keyascii)) = False Then
    If keyascii <> 13 And keyascii <> 27 And keyascii <> 8 Then
        keyascii = 0
        Exit Sub
    End If
End If

' if enter is pressed then functions are called to calculate the values for the other
' base types

If keyascii = 13 Then
    txtbin.Text = calcden2bin(txtden.Text)
    txthex.Text = calcden2hex(txtden.Text)
    txtascii.Text = calcden2asci(txtden.Text)
End If

' if keyascii = 27, the escape key, and there is no text in the box then the program is
' ended

If keyascii = 27 And Len(txtden.Text) = 0 Then End

' ifkeyascii = 27, the escape key, and there is text in the box then the text boxes are cleared

If keyascii = 27 And Len(txtden.Text) <> 0 Then
    txtden.Text = ""
    txtbin.Text = ""
    txthex.Text = ""
    txtascii.Text = ""
End If
End Sub
Private Function calcden2bin(ByVal den As String) As String

' the two variables to be used in this function are reset

string1 = ""
integer1 = 0

' integer1 = mod 2 of den
' den = integer division of itself
' string1 = integer1 and itself

Do Until den = 0
    integer1 = den Mod 2
    den = den \ 2
    string1 = integer1 & string1
Loop

'the string is then padded out using the pad function then returned

string1 = pad(string1)
calcden2bin = string1
End Function
Private Function calcbin2den(ByVal bin As String) As String

' the variables to be used in this function are set to their starting values

integer1 = 0
string1 = ""
integer3 = 0
integer2 = 1


Do While integer3 < Len(bin) 'do while the digit number being checked is 1 less than the total length of the code to be checked
    integer3 = integer3 + 1
    If integer3 > 1 Then integer2 = integer2 * 2 'increase the digno (used for digit ident + weighting) by one
    string1 = Left((Right(bin, integer3)), 1) 'set currentdigit to = the
    integer1 = integer1 + (integer2 * string1) 'total = total + the digno(weighting) * the current digit
Loop

' the end value is returned

calcbin2den = integer1
End Function
Private Function calcden2hex(ByVal den As String) As String

' all the variables are reset

string1 = ""
string2 = ""
string3 = ""

' string1 = the integer division of the input number

string1 = (den \ 16)

' if string 1 is greater than nine then its charecter is calculated using the lcalc function

If string1 > 9 Then string1 = lcalc(string1)
string2 = den - ((den \ 16) * 16)

' if the string2 is greater than 9 then its charecter is calculated using the lcalc function
' if not then string3 = string 2
If string2 > 9 Then
    string3 = lcalc(string2)
Else
    string3 = string2
End If

' the value bearing strings are combined and returned

string2 = string1 & string3
calcden2hex = string2
End Function
Private Function calchex2den(ByVal hex As String) As String

' all the variables to be used in this function are reset

string1 = "0"
string2 = "0"
integer1 = 0
integer2 = 0
integer3 = 0

' string1 = ther rightermost digit of the input hex value

string1 = Right(hex, 1)

' if string1 isn't numeric then integer1 = l2calc of the charecter string1
' if string1 is numeric then integer1 = string1

If IsNumeric(string1) = False Then
    integer1 = l2calc(string1)
    Else:
    integer1 = Val(string1)
End If

' if the length of the input string is 2 then
' string2 = leftermost digit of the input string
' if string2 isnt numeric then string 2 = l2calc of string2
' integer2 = the value held by string2

If Len(hex) = 2 Then
    string2 = Left(hex, 1)
        If IsNumeric(string2) = False Then
            integer2 = l2calc(string2)
        Else:
            integer2 = Val(string2)
        End If
End If

' integer3 = integer1 + (integer2 * base value)
' integer3 is then returned

integer3 = integer1 + (integer2 * 16)
calchex2den = integer3
End Function
Private Function calcden2asci(ByVal den As Integer) As String

' to keep in line with the rest of the program i have set the simple function chr

calcden2asci = Chr(den)
End Function
Private Function calcasci2den(ByVal asci As String) As Integer

' to keep in line with the rest of the program i have set the simple function asc

calcasci2den = Asc(asci)
End Function
Private Function lcalc(ByVal remainder As String) As String

' using select case the letter associated with the value input is calculated

Select Case remainder
    Case 10
        remainder = "a"
    Case 11
        remainder = "b"
    Case 12
        remainder = "c"
    Case 13
        remainder = "d"
    Case 14
        remainder = "e"
    Case 15
        remainder = "f"
End Select

' the letter for the number is returned

lcalc = remainder
End Function
Private Function l2calc(ByVal hexletter As String) As Integer

' using select case the value associated with the letter input is calculated

Select Case hexletter
    Case "a"
        hexletter = 10
    Case "b"
        hexletter = 11
    Case "c"
        hexletter = 12
    Case "d"
        hexletter = 13
    Case "e"
        hexletter = 14
    Case "f"
        hexletter = 15
End Select

' the number for the letter is returned

l2calc = hexletter
End Function

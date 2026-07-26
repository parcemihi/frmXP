Attribute VB_Name = "basReadWriteDAO"
'
'
' This unique sequence of symbols
' was arranged by Steve Bazinet.
'
' Spring 2000
' Copyright (c) 2000 Steven Roland Bazinet
'
Option Explicit
'
'
' !!!CAUTION!!!
Const mcDATABASE As String = "ReadWriteDAO.mdb"
' This database file must be in the same folder that this VB code is in
' due to the use of the App.Path property in the Subs below.
'
' "For the App object, Path specifies the path of the project .VBP file when
' running the application from the development environment or the path of
' the .exe file when running the application as an executable file."
' MSDN Library Visual Studio 6.0
'
'
' ROUTINES HEREIN:
'
' TRAVERSE A TABLE  subTranverseTable
'  UPDATE  A TABLE  subUpdateTable
'  DISPLAY A ROW    subDisplayRow
'  UPDATE  A ROW    subUpdateRow
'
'
'
'                   It is assumed that no errors occur; there is no error trapping.
'     GLOBAL        This is done to simplify the examples.  This document is intended
'                   for beginners to DAO; it contains two extremely easy DAO examples.
'   ASSUMPTIONS     Error trapping is a more advanced topic relative to the ideas being
'                   conveyed here.
'
'
'
'        Name:      TRAVERSE A TABLE
'     Purpose:      This Sub demonstrates in the simpliest possible way opening a Microsoft
'                   Access Database file and a table within then printing every row in
'                   that table.
'      Inputs:      None.
' Assumptions:      See GLOBAL ASSUMPTIONS above.
'     Effects:      Rows of text will be displayed in the Immediate Window.
'     Returns:      Nothing.
'
Public Sub subTranverseTable()

Dim dbs As Database                                     ' Database variable.
Dim rst As Recordset                                    ' Recordset variable.

    Set dbs = OpenDatabase(App.Path & "\" & mcDATABASE) ' Open the database.
    Set rst = dbs.OpenRecordset("tblNames")             ' Open the table.
    With rst                                            ' Start With block.
        Do Until .EOF                                   ' Do until end of file.
            Debug.Print !strFirst, _
                        !strMiddle, _
                        !strLast                        ' Print the columns.
            .MoveNext                                   ' Next row.
        Loop                                            ' Do it again!
        .Close                                          ' Close the recordset.
    End With                                            ' End With block.
    dbs.Close                                           ' Close the database.

End Sub
'
' TRY IT:    subTranverseTable    IN IMMEDIATE WINDOW
'
' BEGINNER, to "TRY IT" follow these steps exactly:
' 1) Select \View\Immediate Window from Menu Bar to open a window where
'    you can execute Subs immediately.  It's also called the Debug Window.
' 2) Double left click the text, "subTranverseTable" above to mark it.
' 3) Left click the Copy Icon on the toolbar to copy that text to the clipboard.
' 4) Left click in the Immediate Window on a blank line.
' 5) Left click the Paste Icon to paste the Sub name into the Immediate Window.
' 6) Finally, press the Enter key and the Sub will execute.
'
'
'
'        Name:      UPDATE A TABLE
'     Purpose:      This Sub demonstrates in the simpliest possible way opening a Microsoft
'                   Access Database file and a table within then updating every row in
'                   that table.
'      Inputs:      None.
' Assumptions:      See GLOBAL ASSUMPTIONS above.
'     Effects:      None.  Arguably it displays text and updates columns but these
'                   effects are done in the called subs.
'     Returns:      Nothing.
Public Sub subUpdateTable()

Dim dbs As Database                                     ' Database variable.
Dim rst As Recordset                                    ' Recordset variable.

    Set dbs = OpenDatabase(App.Path & "\" & mcDATABASE) ' Open the database.
    Set rst = dbs.OpenRecordset("tblRates")             ' Open the table.
    With rst                                            ' Start With block.
        Do Until .EOF                                   ' Do until end of file.
            subDisplayRow rst                           ' Display columns.
            subUpdateRow rst                            ' Bump up rates
            subDisplayRow rst                           ' Display updated columns.
            .MoveNext                                   ' Next row.
        Loop                                            ' Do it again!
        .Close                                          ' Close the recordset.
    End With                                            ' End With block.
    dbs.Close                                           ' Close the database.

End Sub
'
' TRY IT:    subUpdateTable    IN IMMEDIATE WINDOW
'
'
'
' SUPPORT ROUTINES
'
'
'        Name:      DISPLAY A ROW
'     Purpose:      This Sub Debug.Prints the columns in the current row.
'      Inputs:      rst         The Rates table recordset.
' Assumptions:      See GLOBAL ASSUMPTIONS above.
'     Effects:      Rows of text will be displayed in the Immediate Window.
'     Returns:      Nothing.
'
Public Sub subDisplayRow(rst As Recordset)

Static blnBlank As Boolean                              ' Blank line control.

    With rst                                            ' Start With block.
        Debug.Print !pkAgeCode, _
                    Format(!dblRateRiskHigh, "0.00000"), _
                    Format(!dblRateRiskMiddle, "0.00000"), _
                    Format(!dblRateRiskLow, "0.00000")  ' Print the columns.
        If blnBlank Then Debug.Print                    ' Blank line?
        blnBlank = Not blnBlank                         ' Toggle blank line switch.
    End With                                            ' End With block.

End Sub
'
'
'        Name:      UPDATE A ROW
'     Purpose:      This Sub increments the rate columns in the current row.
'      Inputs:      rst         The Rates table recordset.
' Assumptions:      See GLOBAL ASSUMPTIONS above.
'     Effects:      Columns of Rate data are modified.
'     Returns:      Nothing.
'
Public Sub subUpdateRow(rst As Recordset)

Const cBUMP As Double = 1.05                            ' Increase multipler.

    With rst                                            ' Start With block.
        .Edit                                           ' Enter Edit mode.
        !dblRateRiskHigh = !dblRateRiskHigh * cBUMP     ' Bump high risk rate.
        !dblRateRiskMiddle = !dblRateRiskMiddle * cBUMP ' Bump middle.
        !dblRateRiskLow = !dblRateRiskLow * cBUMP       ' Bump low.
        .Update                                         ' Update the recordset.
    End With                                            ' End With block.

End Sub
'
'
' Copyright (c) 2000 Steven Roland Bazinet
'
'

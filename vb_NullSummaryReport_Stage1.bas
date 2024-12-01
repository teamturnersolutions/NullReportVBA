Sub NullSummaryReport_Stage1()

    ' Delete worksheet named "LPN_Level_Data_1" if it exists
    On Error Resume Next
    Application.DisplayAlerts = False ' Suppress confirmation prompts
    Worksheets("LPN_Level_Data_1").Delete
    Application.DisplayAlerts = True

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Perform operations on the current worksheet
    ' 1. Delete rows 1:5
    ws.Rows("1:5").Delete

    ' 2. Delete column B (originally "DIR")
    ws.Columns("B").Delete

    ' 3. Delete column C (now "XREFLPN")
    ws.Columns("C").Delete

    ' 4. Delete columns I:J ("IS_VMI", "HAS_OPEN_TASKS")
    ws.Columns("I:J").Delete

    ' 5. Delete column J ("TASKS")
    ws.Columns("J").Delete

    ' 6. Add 2 columns between C and D
    ws.Columns("D").Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Columns("D").Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    ' 7. Name the new columns with "NULL" + current date
    Dim currentDate As String
    currentDate = Format(Now, "yyyy-mm-dd")
    ws.Cells(1, 4).Value = "NULL " & currentDate
    ws.Cells(1, 5).Value = "NULL " & currentDate

    ' 8. Set the column headers according to the new layout
    ws.Cells(1, 1).Value = "WHSE"
    ws.Cells(1, 2).Value = "LPN"
    ws.Cells(1, 3).Value = "LPN_STATUS"
    ws.Cells(1, 4).Value = "SHIFT"
    ws.Cells(1, 5).Value = "DEPT"
    ws.Cells(1, 6).Value = "LAST_TOUCHED"
    ws.Cells(1, 7).Value = "LAST_TRANSACTION"
    ws.Cells(1, 8).Value = "LAST_USER"
    ws.Cells(1, 9).Value = "CREATED_DTTM"
    ws.Cells(1, 10).Value = "CLUB"
    ws.Cells(1, 11).Value = "PREV_LOCN"
    ws.Cells(1, 12).Value = "PO"
    ws.Cells(1, 13).Value = "ARTICLE"
    ws.Cells(1, 14).Value = "IDESCR"
    ws.Cells(1, 15).Value = "QTY"

    ' 9. Apply formula to column D "SHIFT" using column F "LAST_TOUCHED"
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Determine the last row based on column A

    ws.Range("D2").Formula = _
        "=IF(ISBLANK(F2), ""UNKNOWN"", " & _
        "IF(AND(WEEKDAY(F2, 2) >= 1, WEEKDAY(F2, 2) <= 5, MOD(F2, 1) >= TIME(4, 0, 0), MOD(F2, 1) < TIME(16, 0, 0)), ""1ST"", " & _
        "IF(AND(WEEKDAY(F2, 2) >= 1, WEEKDAY(F2, 2) <= 5), ""2ND"", " & _
        "IF(AND(WEEKDAY(F2, 2) >= 6, WEEKDAY(F2, 2) <= 7, MOD(F2, 1) >= TIME(4, 0, 0), MOD(F2, 1) < TIME(16, 0, 0)), ""3RD"", ""4TH""))))"
    
    ws.Range("D2").AutoFill Destination:=ws.Range("D2:D" & lastRow)

    ' 10. Apply formula to column E "DEPT" using column G "LAST_TRANSACTION" with updated mapping logic
    ws.Range("E2").Formula = _
        "=SWITCH(G2, " & _
        """LPN Disposition     *"", ""PTC"", " & _
        """Pck Cubed Dir     *"", ""STG"", " & _
        """Ptwy iLPN     *"", ""STG"", " & _
        """Ptwy User Non EX01     *"", ""STG"", " & _
        """Recv By ASN     *"", ""REC"", " & _
        """Recv By ASN Shuttle*"", ""REC"", " & _
        """Recv Floor     *"", ""REC"", " & _
        """Recv Mass - Single Sku     *"", ""REC"", " & _
        """Unload LPN     *"", ""SHP"", " & _
        """UNKNOWN"")"
    
    ws.Range("E2").AutoFill Destination:=ws.Range("E2:E" & lastRow)

    ' Rename the current worksheet to include "NULL" and the current date
    ws.Name = "NULL " & currentDate

End Sub


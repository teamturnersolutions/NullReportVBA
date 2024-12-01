Sub NullSummaryReport_Stage2()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim lastTouched As Date
    Dim cutoffDate As Date
    Dim currentTime As Date
    
    ' Set the active worksheet
    Set ws = ActiveSheet

    ' Find the last used row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Get the current date and time
    currentTime = Now
    
    ' Calculate the cutoff date: 3 days ago at 4:00 AM
    cutoffDate = Date - 3 + TimeValue("04:00:00")
    
    ' Loop through rows and delete rows based on conditions
    For i = lastRow To 2 Step -1 ' Start from the bottom and work upwards
        ' Check if value in Column A is 582 or check the date in Column E
        If ws.Cells(i, "A").Value = 582 Then
            ws.Rows(i).Delete
        Else
            ' Check if the date in Column E is under 3 days old and after 4 AM on the 3rd day
            If IsDate(ws.Cells(i, "E").Value) Then
                lastTouched = ws.Cells(i, "E").Value
                ' If the LAST_TOUCHED date is younger than the cutoff date (after 11/27/2024 04:00 AM), delete the row
                If lastTouched > cutoffDate Then
                    ws.Rows(i).Delete
                End If
            End If
        End If
    Next i

    ' Delete column A (WHSE)
    ws.Columns("A").Delete

    ' Rename the active worksheet to "NULL"
    ws.Name = "NULL"

    ' Notify the user
    MsgBox "Data from 582 deleted", vbInformation
End Sub


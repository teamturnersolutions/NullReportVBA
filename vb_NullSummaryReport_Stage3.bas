Sub NullSummaryReport_Stage3()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim wbSource As Workbook
    Dim wbDest As Workbook
    Dim lastRowSource As Long
    Dim lastRowDest As Long
    Dim i As Long
    Dim sourceWBName As String
    Dim destinationWBName As String
    Dim destTable As ListObject
    Dim newTableRange As Range
    
    ' Set the name of the source and destination workbooks
    sourceWBName = "Null Location LPNs.xlsx" ' Modify this to match the new naming convention
    destinationWBName = "Sandbox.xlsm" ' Name of the destination workbook
    
    ' Find the source workbook
    On Error Resume Next ' In case the workbook is not found
    Set wbSource = Workbooks(sourceWBName)
    On Error GoTo 0 ' Reset error handling to default
    
    ' If the source workbook is not open, inform the user and exit
    If wbSource Is Nothing Then
        MsgBox "Source workbook '" & sourceWBName & "' not found.", vbExclamation
        Exit Sub
    End If
    
    ' Set the source worksheet (NULL)
    Set wsSource = wbSource.Sheets("NULL")
    
    ' Find the destination workbook
    On Error Resume Next
    Set wbDest = Workbooks(destinationWBName)
    On Error GoTo 0
    
    ' If the destination workbook is not open, inform the user and exit
    If wbDest Is Nothing Then
        MsgBox "Destination workbook '" & destinationWBName & "' not found.", vbExclamation
        Exit Sub
    End If
    
    ' Set the destination worksheet (NULL)
    Set wsDest = wbDest.Sheets("NULL")
    
    ' Find the last row in the source sheet
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Set the destination start row to 4 (first blank row after headers)
    lastRowDest = 4 ' The first blank row for pasting starts at row 4
    
    ' Copy data from source to destination (starting from row 4 in destination)
    For i = 2 To lastRowSource ' Assuming data starts from row 2 in the source
        wsSource.Rows(i).Copy Destination:=wsDest.Rows(lastRowDest)
        lastRowDest = lastRowDest + 1 ' Move down to the next row for the next transfer
    Next i
    
    ' Notify the user when the transfer is complete
    MsgBox "Data Transfer Complete!", vbInformation
    
 
End Sub


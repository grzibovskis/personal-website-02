Option Explicit


'===================================================== CLEAR TaskCount Cells =====================================================================
Sub clearSTATS()

    ' List of cells in column B to clear
    Dim cellList As Variant
    cellList = Array(2, 3, 4, 5, 7, 8, 11, 12, 14, 15, 16, 17, 18, 20, 21, 22, 23, 24, 25, 26, 27, 28, 34, 35, 36, 37, 41, 43, 48, 49)
    
    Dim i As Integer
    For i = LBound(cellList) To UBound(cellList)
        shTaskCount.Cells(cellList(i), 2).ClearContents ' Column B is the second column
    Next i

    ' Clear cells C21 and D21
    shData.Range("C21, D21").ClearContents

    MsgBox "Selected cells in column B, and cells C21 and D21 have been cleared.", vbInformation
End Sub
'===================================================== CLEAR Data Cells =====================================================================
Sub clearDATA()

    Dim cell As Range
    Dim shape As shape
    
    ' Clear contents and format of columns A to K up to row 350
    With shData.Range("A1:K350")
        .ClearContents ' Remove cell contents
        .ClearFormats ' Clear cell formatting
        .UnMerge ' Unmerge any merged cells
    End With
    
    ' Remove all shapes (images, icons, etc.) within the range A1:K350
    For Each shape In shData.Shapes
        ' Check if the shape is within the specified range
        If Not Intersect(shape.TopLeftCell, shData.Range("A1:K350")) Is Nothing Then
            shape.Delete
        End If
    Next shape

End Sub

'===================================================== KPIinfo call All Subs ================================================================
Sub subCALLER()

    'Initialize output cells to 0 to handle cases where no value is found

    Dim cellRefs As Variant
    Dim cellRef As Variant
    Dim foundGlazed As Boolean, foundMatcha As Boolean

    ' List of cell references in ascending order
    cellRefs = Array("B2", "B3", "B4", "B5", "B7", "B8", "B11", "B12", _
                     "B14", "B15", "B16", "B17", "B18", "B20", "B22", "B23", _
                     "B24", "B25", "B26", "B27", "B28", "B35", "B36", "B37", _
                     "B38", "B48", "B49", "C21", "D21")

    ' Loop through the cell references and set values to 0
    For Each cellRef In cellRefs
        shTaskCount.Range(cellRef).Value = 0
    Next cellRef
    
    foundGlazed = False
    foundMatcha = False

    ' Call each individual subroutine in the desired order
    Call CountSPECIALS
    Call CountITEMS
    Call CountMMAP
    Call CountACK
    Call countTASKS

    ' Display a message when all procedures are complete
    MsgBox "All procedures have completed successfully!", vbInformation
End Sub

'===================================================== Foretak ================================================================
Sub countTASKS()
    Dim appleWorkbook As Workbook
    Dim appleWorksheet As Worksheet
    Dim cell As Range
    Dim lastRow As Long
    Dim i As Long
    Dim checkCell As Range
    Dim countCells As Long
    Dim result As Long
    Dim appleFilePath As String
    Dim extractedNumber As Double
    Dim totalInvest As Double
    Dim totalRerun As Double
    Dim totalSumB40 As Double
    
    ' Define file path based on cell A55 in the shTaskCount sheet
    appleFilePath = shTaskCount.Range("A55").Value
    
    ' Initialize counts
    totalInvest = 0
    totalRerun = 0
    totalSumB40 = 0
    
    ' Case 1: Count and Multiply Cells in Apple2024.xlsm file
    Application.DisplayAlerts = False
    On Error Resume Next
    Set appleWorkbook = Workbooks.Open(appleFilePath, UpdateLinks:=True)
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    If Not appleWorkbook Is Nothing Then
        Set appleWorksheet = appleWorkbook.Sheets(1) ' Adjust if specific sheet is not the first sheet
        
        ' Find the last row with data in column B
        lastRow = appleWorksheet.Cells(appleWorksheet.Rows.count, "B").End(xlUp).Row
    
        ' Initialize countCells
        countCells = 0
    
        ' Loop from row 2 to the lastRow
        For i = 2 To lastRow
            If appleWorksheet.Cells(i, "B").Value <> "" Then
                countCells = countCells + 1
            End If
        Next i
    
        ' Multiply the count by 2 and output result in B34
        result = countCells * 2
        shTaskCount.Range("B34").Value = result
    
        ' Close the opened workbook without saving
        appleWorkbook.Close False
    Else
        MsgBox "Could not open the file at " & appleFilePath, vbExclamation
        Exit Sub
    End If


    ' Case 2 and 3: Search for "INVEST" and "RERUN" tags in "onenote" sheet (Columns A, B, and C)
    For Each checkCell In shData.Range("A1:C" & shData.Cells(shData.Rows.count, "A").End(xlUp).Row)
        Select Case True
            Case InStr(1, checkCell.Value, "INVEST", vbTextCompare) > 0
                ' Move 2 cells to the right and extract numeric value
                If IsNumericInCell(checkCell.Offset(0, 2), extractedNumber) Then
                    totalInvest = totalInvest + extractedNumber
                    totalSumB40 = totalSumB40 + extractedNumber
                End If
                
            Case InStr(1, checkCell.Value, "RERUN", vbTextCompare) > 0
                ' Move 2 cells to the right and extract numeric value
                If IsNumericInCell(checkCell.Offset(0, 2), extractedNumber) Then
                    totalRerun = totalRerun + extractedNumber
                End If
        End Select
    Next checkCell
    
    ' Output results to shTaskCount sheet
    shTaskCount.Range("B41").Value = totalInvest    ' "INVEST" count
    shTaskCount.Range("B43").Value = totalRerun     ' "RERUN" count
End Sub

' Helper function to extract numeric value from cell ignoring non-numeric symbols or characters ===========FORETAK
Function IsNumericInCell(rng As Range, ByRef numericValue As Double) As Boolean
    Dim cellValue As String
    Dim temp As String
    Dim i As Long
    
    cellValue = rng.Value
    temp = ""
    
    ' Extract numeric characters only at the start of the cell
    For i = 1 To Len(cellValue)
        If Mid(cellValue, i, 1) Like "[0-9]" Then
            temp = temp & Mid(cellValue, i, 1)
        Else
            Exit For ' Stop if a non-numeric character is encountered
        End If
    Next i
    
    ' Validate that the next character is not a symbol or text
    If Len(temp) > 0 And Not IsNumeric(Mid(cellValue, Len(temp) + 1, 1)) Then
        numericValue = CDbl(temp)
        IsNumericInCell = True
    Else
        numericValue = 0
        IsNumericInCell = True
    End If
End Function

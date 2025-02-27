Option Explicit
'===================== THIS MODULE INCLUDES TAGS : Latte,Aventador,Porsche,Chevrolet

Sub CountMMAP()
    Dim searchColumns As Variant
    Dim col As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim cell As Range
    Dim latteCount As Long
    Dim aventadorCount As Long
    Dim porscheCount As Long
    Dim chevroletCount As Long

    searchColumns = Array("C", "D", "E", "F", "G", "H")

    ' Loop through each specified column
    For Each col In searchColumns
        lastRow = shData.Cells(shData.Rows.count, col).End(xlUp).Row ' Find the last row in the column

        ' Loop through each cell in the column up to the last row
        For i = 1 To lastRow
            Set cell = shData.Cells(i, col)

            ' Check each tag using a Select Case structure
            Select Case True
                Case InStr(1, cell.Value, "Latte", vbTextCompare) > 0
                    latteCount = latteCount + CountNonEmptyCells(shData, col, i + 1)

                Case InStr(1, cell.Value, "Aventador", vbTextCompare) > 0
                    aventadorCount = aventadorCount + CountNonEmptyCells(shData, col, i + 1)

                Case InStr(1, cell.Value, "Porsche", vbTextCompare) > 0
                    porscheCount = porscheCount + CountNonEmptyCells(shData, col, i + 1)

                Case InStr(1, cell.Value, "Chevrolet", vbTextCompare) > 0
                    chevroletCount = chevroletCount + CountNonEmptyCells(shData, col, i + 1)
            End Select
        Next i
    Next col

    ' Output the counts to the specified cells on KPIinfo sheet
    shTaskCount.Range("B8").Value = latteCount
    shTaskCount.Range("B18").Value = aventadorCount
    shTaskCount.Range("B17").Value = porscheCount
    shTaskCount.Range("B24").Value = chevroletCount
    shTaskCount.Range("B25").Value = chevroletCount
End Sub

'===================================================== NO EVENTS ignor: for Modules mainISINCOUNT02, mainACKCOUNT01 =======================================================================

' Function to count non-empty cells following "Spice" tag, ignoring "no events" entries
Function CountNonEmptyCells(shData As Worksheet, col As Variant, startRow As Long) As Long
    Dim cell As Range
    Dim lastRow As Long
    Dim count As Long
    Dim RowIndex As Long

    ' Find the last non-empty row in the column
    lastRow = shData.Cells(shData.Rows.count, col).End(xlUp).Row
    count = 0

    ' Use a For loop to count consecutive non-empty cells
    For RowIndex = startRow To lastRow
        Set cell = shData.Cells(RowIndex, col)

        ' Ignore cells that contain "no events" (case-insensitive)
        If Len(Trim(cell.Value)) > 0 And InStr(1, cell.Value, "no events", vbTextCompare) = 0 Then
            count = count + 1
        ElseIf Len(Trim(cell.Value)) = 0 Then
            ' Exit if an empty cell is encountered
            Exit For
        End If
    Next RowIndex

    CountNonEmptyCells = count
End Function

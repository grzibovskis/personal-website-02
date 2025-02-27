Option Explicit
'===================== THIS MODULE INCLUDES TAGS : Honey,Honda,Pumpkin,Spice

Sub CountACK()
    Dim searchColumns As Variant
    Dim col As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim cell As Range
    Dim extractedNumber As Double
    Dim SpiceCount As Long

    searchColumns = Array("C", "D", "E", "F", "G", "H")
    
    ' Loop through each specified column
    For Each col In searchColumns
        lastRow = shData.Cells(shData.Rows.count, col).End(xlUp).Row

        ' Loop through each cell in the column up to the last row
        For i = 1 To lastRow
            Set cell = shData.Cells(i, col)

            ' Check each tag using a Select Case structure
            Select Case True
                Case InStr(1, cell.Value, "Honey", vbTextCompare) > 0
                    If SumACKNumbers(shData, col, i + 1, extractedNumber) Then
                        shTaskCount.Range("B7").Value = shTaskCount.Range("B7").Value + extractedNumber
                    End If

                Case InStr(1, cell.Value, "Honda", vbTextCompare) > 0
                    If SumACKNumbers(shData, col, i + 1, extractedNumber) Then
                        shTaskCount.Range("B26").Value = shTaskCount.Range("B26").Value + extractedNumber
                    End If

                Case InStr(1, cell.Value, "Pumpkin", vbTextCompare) > 0
                    If SumACKNumbers(shData, col, i + 1, extractedNumber) Then
                        shTaskCount.Range("B14").Value = shTaskCount.Range("B14").Value + extractedNumber
                    End If

                Case InStr(1, cell.Value, "Spice", vbTextCompare) > 0
                    SpiceCount = CountNonEmptyCells(shData, col, i + 1)
                    shTaskCount.Range("B14").Value = shTaskCount.Range("B14").Value + SpiceCount
            End Select
        Next i
    Next col
End Sub

'===================================================== NO EVENTS ignor: for Modules mainISINCOUNT02, mainACKCOUNT01 =====================

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

'===================================================== ACK n NACK =======================================================================

' Function to sum numbers following "ACK" tag in a column
Function SumACKNumbers(shData As Worksheet, col As Variant, startRow As Long, ByRef total As Double) As Boolean
    Dim cell As Range
    Dim ackValue As String
    Dim nackValue As String
    Dim lastRow As Long
    Dim RowIndex As Long
    
    ' Determine the last non-empty row in the column
    lastRow = shData.Cells(shData.Rows.count, col).End(xlUp).Row
    total = 0

    ' Use a For loop to iterate through rows
    For RowIndex = startRow To lastRow
        Set cell = shData.Cells(RowIndex, col)

        ' Check if the cell contains "ACK" or "NACK" (case-insensitive)
        If InStr(1, cell.Value, "ACK", vbTextCompare) > 0 Or InStr(1, cell.Value, "NACK", vbTextCompare) > 0 Then
            ' Remove any variation of "ACK" or "NACK" (case-insensitive) and "-" characters
            ackValue = Replace(cell.Value, "ACK", "", , , vbTextCompare)
            ackValue = Replace(ackValue, "-", "")
            ackValue = Trim(ackValue)

            nackValue = Replace(cell.Value, "NACK", "", , , vbTextCompare)
            nackValue = Replace(nackValue, "-", "")
            nackValue = Trim(nackValue)

            ' Add to total if value is numeric
            If IsNumeric(ackValue) Then
                total = total + CDbl(ackValue)
            End If
            If IsNumeric(nackValue) Then
                total = total + CDbl(nackValue)
            End If
        Else
            ' Exit loop if a non-"ACK" or "NACK" cell is found
            Exit For
        End If
    Next RowIndex

    ' Return True if any numbers were added to the total
    SumACKNumbers = total > 0
End Function





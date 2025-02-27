Option Explicit

'===================== THIS MODULE INCLUDES TAGS : Cinnamon, Tesla, Mustang, Wrangler, Canyon,
'Greece, Japan, Canada, Iceland, French, Toyota, Italy

Sub CountITEMS()
    Dim searchColumns As Variant
    Dim col As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim cell As Range

    ' Initialize cumulative totals for each tag
    Dim sumCinnamon As Double
    Dim sumTesla As Double
    Dim sumMustang As Double
    Dim sumWrangler As Double
    Dim sumCanyon As Double
    Dim sumGreece As Double
    Dim sumJapan As Double
    Dim sumCanada As Double
    Dim sumIceland As Double
    Dim sumFrench As Double
    Dim sumToyota As Double
    Dim sumItaly As Double

    searchColumns = Array("C", "D", "E", "F", "G", "H") ' Columns to search in

    ' Loop through each specified column
    For Each col In searchColumns
        lastRow = shData.Cells(shData.Rows.count, col).End(xlUp).Row ' Find the last row in the column

        ' Loop through each cell in the column up to the last row
        For i = 1 To lastRow
            Set cell = shData.Cells(i, col)
            
            ' Remove spaces from cell value for checking tags
            Dim cellValue As String
            cellValue = Replace(cell.Value, " ", "")

            ' Check each tag and add the values to the corresponding cumulative sum
            Select Case True
                Case InStr(1, cellValue, "Cinnamon", vbTextCompare) > 0
                    Call AddConsecutiveNumbers(shData, col, i + 1, sumCinnamon)
                
                Case InStr(1, cellValue, "Tesla", vbTextCompare) > 0
                    Call AddConsecutiveNumbers(shData, col, i + 1, sumTesla)
                
                Case InStr(1, cellValue, "Mustang", vbTextCompare) > 0
                    Call AddConsecutiveNumbers(shData, col, i + 1, sumMustang)
                    
                Case InStr(1, cellValue, "Wrangler", vbTextCompare) > 0
                    Call AddConsecutiveNumbers(shData, col, i + 1, sumWrangler)
                    
                Case InStr(1, cellValue, "Canyon", vbTextCompare) > 0
                    Call AddConsecutiveNumbers(shData, col, i + 1, sumCanyon)
                
                Case InStr(1, cellValue, "Greece", vbTextCompare) > 0
                    Call AddConsecutiveNumbers(shData, col, i + 1, sumGreece)
                
                Case InStr(1, cellValue, "Japan", vbTextCompare) > 0
                    Call AddConsecutiveNumbers(shData, col, i + 1, sumJapan)
                
                Case InStr(1, cellValue, "Canada", vbTextCompare) > 0
                    Call AddConsecutiveNumbers(shData, col, i + 1, sumCanada)
                
                Case InStr(1, cellValue, "Iceland", vbTextCompare) > 0
                    Call AddConsecutiveNumbers(shData, col, i + 1, sumIceland)
                
                Case InStr(1, cellValue, "French", vbTextCompare) > 0
                    Call AddConsecutiveNumbers(shData, col, i + 1, sumFrench)
                
                Case InStr(1, cellValue, "Toyota", vbTextCompare) > 0
                    Call AddConsecutiveNumbers(shData, col, i + 1, sumToyota)
                    
                Case InStr(1, cellValue, "Italy", vbTextCompare) > 0
                    Call AddConsecutiveNumbers(shData, col, i + 1, sumItaly)
            End Select
        Next i
    Next col

    ' Output the cumulative sums for each tag to the appropriate cells
    shTaskCount.Range("B11").Value = sumCinnamon
    shTaskCount.Range("B15").Value = sumTesla
    shTaskCount.Range("B16").Value = sumMustang
    shTaskCount.Range("B27").Value = sumWrangler
    shTaskCount.Range("B49").Value = sumCanyon
    shTaskCount.Range("B28").Value = sumGreece
    shTaskCount.Range("B35").Value = sumJapan
    shTaskCount.Range("B36").Value = sumCanada
    shTaskCount.Range("B37").Value = sumIceland
    shTaskCount.Range("B38").Value = sumFrench
    shTaskCount.Range("B20").Value = sumToyota
    shTaskCount.Range("B48").Value = sumItaly
End Sub

' Subroutine to add consecutive numbers to a cumulative total variable, skipping blank cells
Sub AddConsecutiveNumbers(ws As Worksheet, col As Variant, startRow As Long, ByRef cumulativeSum As Double)
    Dim extractedNumber As Double
    Dim lastRow As Long
    Dim RowIndex As Long
    Dim cellValue As String

    ' Determine the last non-empty row in the specified column
    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row

    ' Use a For loop to iterate through the rows
    For RowIndex = startRow To lastRow
        ' Skip blank cells without stopping the loop
        If Not IsEmpty(ws.Cells(RowIndex, col).Value) Then
            ' Remove spaces from cell content before extracting number
            cellValue = Replace(ws.Cells(RowIndex, col).Value, " ", "")
            
            ' Only add if the cell contains a number at the start
            If ExtractLeadingNumber(cellValue, extractedNumber) Then
                cumulativeSum = cumulativeSum + extractedNumber
            Else
                Exit For ' Stop if there's no leading number
            End If
        End If
    Next RowIndex
End Sub


' Helper function to extract leading number from a cell's text, returning True if successful
Function ExtractLeadingNumber(cellText As String, ByRef leadingNumber As Double) As Boolean
    Dim regex As Object
    Dim match As Object

    ' Set up regex to find leading number in the cell's text
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^\d+(\.\d+)?"
    regex.IgnoreCase = True
    regex.Global = False

    ' Check for a match and extract the number if found
    If regex.test(cellText) Then
        Set match = regex.Execute(cellText)(0)
        leadingNumber = CDbl(match.Value)
        ExtractLeadingNumber = True
    Else
        ExtractLeadingNumber = False
    End If
End Function

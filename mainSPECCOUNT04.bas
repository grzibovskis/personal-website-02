Option Explicit

'===================== THIS MODULE INCLUDES TAGS : "Risotto", "Blueberry", "Truffle", "Mango", "Avocado", "EXTM", "ROBOT", "Glazed", "Matcha"

Sub CountSPECIALS()
    Dim startRow As Long
    Dim lastRow As Long
    Dim countRisotto As Long
    Dim countBlueberry As Long
    Dim countTruffle As Long
    Dim countMango As Long
    Dim totalNumberSum As Double
    Dim robotCount As Long
    Dim CountEXTM As Long
    Dim foundGlazedValue As Double
    Dim foundMatchaValue As Double
    Dim cell As Range
    Dim checkCell As Range
    Dim belowCell As Range
    Dim ackValue As String
    Dim extractedNumber As Double
    Dim nackValue As String
    Dim searchColumns As Variant
    Dim i As Long
    Dim col As Variant ' can be number/text
    Dim RowIndex As Long
    Dim foundGlazed As Boolean
    Dim foundMatcha As Boolean
    Dim numberFound As Boolean, cellText As String

    searchColumns = Array("C", "D", "E", "F", "G", "H")

    ' Loop through each specified column
    For Each col In searchColumns
        lastRow = shData.Cells(shData.Rows.count, col).End(xlUp).Row ' Find the last row in the column

        For i = 1 To lastRow
            Set cell = shData.Cells(i, col)
            Select Case True
                Case InStr(1, cell.Value, "Risotto", vbTextCompare) > 0
                    If Not IsNoAck(cell.Offset(0, 1)) Then
                        countRisotto = countRisotto + 1
                    End If

                Case InStr(1, cell.Value, "Blueberry", vbTextCompare) > 0
                    If Not IsNoAck(cell.Offset(0, 1)) Then
                        countBlueberry = countBlueberry + 1
                    End If

                Case InStr(1, cell.Value, "Truffle", vbTextCompare) > 0
                    If Not IsNoAck(cell.Offset(0, 1)) Then
                        countTruffle = countTruffle + 1
                    End If

                Case InStr(1, cell.Value, "Mango", vbTextCompare) > 0
                    If Not IsNoAck(cell.Offset(0, 1)) Then
                        countMango = countMango + 1
                    End If
                ' Case for counting Avocado
                Case InStr(1, cell.Value, "Avocado", vbTextCompare) > 0
                    ' Start checking cells below "Avocado"
                    i = i + 1 ' Move to the next row
                
                    ' Initialize the flag for stopping the loop
                    numberFound = False
                
                    ' Loop through the rows from i to lastRow
                    For RowIndex = i To lastRow
                        Set checkCell = shData.Cells(RowIndex, col)
                        cellText = checkCell.Value
                        
                        ' Sum all numbers in the cell text
                        extractedNumber = SumNumbersInText(cellText)
                        If extractedNumber > 0 Then
                            totalNumberSum = totalNumberSum + extractedNumber
                            numberFound = True
                        Else
                            Exit For ' Stop if no numbers are found in the next cell
                        End If
                    Next RowIndex



                ' Case for counting "EXTM"
                Case InStr(1, cell.Value, "CNTEXTM", vbTextCompare) > 0
                    ' Start checking cells below for "EXTM" substring
                    i = i + 1 ' Move to the next row below "CNTEXTM"
                
                    ' Loop through rows from i to lastRow
                    For RowIndex = i To lastRow
                        Set cell = shData.Cells(RowIndex, col)
                        ' Stop if the cell is empty
                        If IsEmpty(cell.Value) Then Exit For
                        ' Check if the cell contains "EXTM" substring
                        If InStr(1, cell.Value, "EXTM", vbTextCompare) > 0 Then
                            CountEXTM = CountEXTM + 1 ' Increment count if "EXTM" is found
                        End If
                    Next RowIndex

                    
                ' Case for "RODC" followed by "ROBOT"
                Case InStr(1, cell.Value, "RODC", vbTextCompare) > 0
                    i = i + 1 ' Move to the row below "RODC"
                
                    ' Loop through rows from i to lastRow
                    For RowIndex = i To lastRow
                        Set belowCell = shData.Cells(RowIndex, col)
                        ' Stop if cell is empty
                        If IsEmpty(belowCell.Value) Then Exit For
                        ' Check if the cell below "RODC" contains "ROBOT"
                        If InStr(1, belowCell.Value, "ROBOT", vbTextCompare) > 0 Then
                            robotCount = robotCount + 1 ' Increment count if "ROBOT" is found below "RODC"
                        End If
                    Next RowIndex

                ' Case for counting "Glazed" and "Matcha"
                Case InStr(1, cell.Value, "Glazed", vbTextCompare) > 0 And InStr(1, cell.Value, "MIN", vbTextCompare) = 0
                    If IsNumeric(cell.Offset(1, 0).Value) Then
                        foundGlazedValue = cell.Offset(1, 0).Value
                        foundGlazed = True
                    Else
                        foundGlazedValue = 0 ' Set to 0 if no numeric value is found
                    End If
                
                Case InStr(1, cell.Value, "Matcha", vbTextCompare) > 0
                    ' Initialize a variable to keep track of the starting row for the "Matcha" section
                    startRow = i + 1 ' Start processing the rows directly below "Matcha"
                
                    ' Loop through consecutive rows to check for "ACK" and "NACK" values
                    For RowIndex = startRow To lastRow
                        Set belowCell = shData.Cells(RowIndex, col)
                        
                        ' Exit the loop if an empty cell is encountered
                        If IsEmpty(belowCell.Value) Then Exit For
                
                        ' Extract and process both "ACK" and "NACK" values in the cell
                        ackValue = Replace(belowCell.Value, "ACK", "", , , vbTextCompare)
                        ackValue = Replace(ackValue, "-", "")
                        ackValue = Trim(ackValue)
                        
                        nackValue = Replace(belowCell.Value, "NACK", "", , , vbTextCompare)
                        nackValue = Replace(nackValue, "-", "")
                        nackValue = Trim(nackValue)
                        
                        ' Add numeric "ACK" values to the total
                        If IsNumeric(ackValue) Then
                            foundMatchaValue = foundMatchaValue + CDbl(ackValue)
                        End If
                        
                        ' Add numeric "NACK" values to the total
                        If IsNumeric(nackValue) Then
                            foundMatchaValue = foundMatchaValue + CDbl(nackValue)
                        End If
                    Next RowIndex
                    
                    foundMatcha = True

            End Select
        Next i
    Next col
    ' Output the counts to the KPIinfo sheet
    shTaskCount.Range("B2").Value = countRisotto
    shTaskCount.Range("B3").Value = countBlueberry
    shTaskCount.Range("B4").Value = countTruffle
    shTaskCount.Range("B5").Value = countMango
    shTaskCount.Range("B12").Value = totalNumberSum
    shTaskCount.Range("B22").Value = robotCount    ' "ROBOT" count
    shTaskCount.Range("B23").Value = CountEXTM     ' "EXTM" count
    shTaskCount.Range("C21").Value = foundGlazedValue
    shTaskCount.Range("D21").Value = foundMatchaValue
    shTaskCount.Range("B21").Value = foundGlazedValue - foundMatchaValue
End Sub

'===================================================== FOR Limpo - MANUAL =======================================================================

' Function to extract the first number found within text
Function ExtractNumberFromText(text As String, ByRef Number As Double) As Boolean
    Dim regex As Object
    Dim matches As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\d+(\.\d+)?" ' Pattern to match integer or decimal numbers
    regex.Global = False
    regex.IgnoreCase = True
    If regex.test(text) Then
        Set matches = regex.Execute(text)
        Number = CDbl(matches(0)) ' Convert the first match to a number
        ExtractNumberFromText = True
    Else
        ExtractNumberFromText = False
    End If
End Function

'===================================================== For CNTDVCACLAIM count all numbers in cells and ignore unnecessary characters ======
Function SumNumbersInText(text As String) As Double
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim sum As Double
    
    ' Create the RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "\d+(\.\d+)?"
    
    ' Initialize the sum
    sum = 0
    
    ' Find all matches (numeric values) in the text
    If regex.test(text) Then
        Set matches = regex.Execute(text)
        For Each match In matches
            sum = sum + CDbl(match.Value)
        Next match
    End If
    
    ' Return the total sum of all numbers found
    SumNumbersInText = sum
End Function

' Helper function to check if a cell contains "no ack" (case insensitive)
Function IsNoAck(rng As Range) As Boolean
    IsNoAck = (InStr(1, rng.Value, "no ack", vbTextCompare) > 0)
End Function
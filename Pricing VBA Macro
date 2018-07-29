Sub EODPricing()

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

Dim templateWorkbook As Workbook
Dim PriceDeviationWorkbook As Workbook
Dim exclusionWorkbook As Workbook

Set PriceDeviationWorkbook = ActiveWorkbook

'Create three empty columns'

Columns("J:M").Insert Shift:=xlToRight

'Give the columns headers'
PriceDeviationWorkbook.Sheets("High Deviation").Range("A:BC").AutoFilter
PriceDeviationWorkbook.Sheets("Missing Price").Range("A:BC").AutoFilter


PriceDeviationWorkbook.Sheets("High Deviation").Range("J1").Value = "CUSIP"
PriceDeviationWorkbook.Sheets("High Deviation").Range("K1").Value = "PRCUSD"
PriceDeviationWorkbook.Sheets("High Deviation").Range("L1").Value = "IEMID"
PriceDeviationWorkbook.Sheets("High Deviation").Range("M1").Value = "RESULT"

'Formatting the columns to text so they can be analyzed'

PriceDeviationWorkbook.Sheets("Missing Price").Range("M1").Value = "CUSIP"
PriceDeviationWorkbook.Sheets("Missing Price").Range("N1").Value = "PRCUSD"
PriceDeviationWorkbook.Sheets("Missing Price").Range("O1").Value = "IEMID"

PriceDeviationWorkbook.Sheets("High Deviation").Range("I:I").TextToColumns
PriceDeviationWorkbook.Sheets("High Deviation").Range("A:A").TextToColumns
PriceDeviationWorkbook.Sheets("Missing Price").Range("A:A").TextToColumns

PriceDeviationWorkbook.Sheets("High Deviation").Range("K:K").NumberFormat = "General"
PriceDeviationWorkbook.Sheets("High Deviation").Range("L:L").NumberFormat = "General"
PriceDeviationWorkbook.Sheets("High Deviation").Range("M:M").NumberFormat = "General"
PriceDeviationWorkbook.Sheets("High Deviation").Range("J:J").NumberFormat = "General"


For i = 2 To 2333

    If PriceDeviationWorkbook.Sheets("Missing Price").Range("A" & i).Value = "" Then Exit For
    
Next i

For lastValueOfFirstSheet = 2 To 2333

    If PriceDeviationWorkbook.Sheets("High Deviation").Range("A" & lastValueOfFirstSheet).Value = "" Then Exit For

Next lastValueOfFirstSheet


PriceDeviationWorkbook.Sheets("Missing Price").Range("A2" & ":A" & i).Copy PriceDeviationWorkbook.Sheets("High Deviation").Range("A" & lastValueOfFirstSheet)

PriceDeviationWorkbook.Application.SendKeys "%", True
PriceDeviationWorkbook.Application.SendKeys "Y3", True
PriceDeviationWorkbook.Application.SendKeys "Y4", True
PriceDeviationWorkbook.Application.SendKeys "~", True

'We paste the true value into row a2333 for the api to read'
PriceDeviationWorkbook.Sheets(1).Range("A2333").Value = "True"

PriceDeviationWorkbook.Application.SendKeys "~", True

'Wait until code is done'
DoEvents

'Move cusip'
PriceDeviationWorkbook.Sheets("High Deviation").Range("J2" & ":J" & i).ClearContents
PriceDeviationWorkbook.Sheets("High Deviation").Range("A2" & ":A" & lastValueOfFirstSheet - 1).Copy PriceDeviationWorkbook.Sheets("High Deviation").Range("J2")

'move cusip for missing price'

PriceDeviationWorkbook.Sheets("Missing Price").Range("A2" & ":A" & i).Copy PriceDeviationWorkbook.Sheets("Missing Price").Range("M2")


'remove the values and move them into missing price'

PriceDeviationWorkbook.Sheets("High Deviation").Range("A" & lastValueOfFirstSheet & ":A" & i).ClearContents

PriceDeviationWorkbook.Sheets("High Deviation").Range("K" & lastValueOfFirstSheet & ":K" & i).Cut PriceDeviationWorkbook.Sheets("Missing Price").Range("N2" & ":N" & i)

PriceDeviationWorkbook.Sheets("High Deviation").Range("L" & lastValueOfFirstSheet & ":L" & i).Cut PriceDeviationWorkbook.Sheets("Missing Price").Range("O2" & ":O" & i)

'Apply filters to High Deviation'

PriceDeviationWorkbook.Sheets("High Deviation").Range("F1").AutoFilter Field:=6, Criteria1:="<>OPEQINEX", Operator:=xlAnd, Criteria2:="<>OPEQTYEX"
PriceDeviationWorkbook.Sheets("High Deviation").Range("N1").AutoFilter Field:=14, Criteria1:="=USD"
PriceDeviationWorkbook.Sheets("High Deviation").Range("O1").AutoFilter Field:=15, Criteria1:=" =IDC", Operator:=xlOr, Criteria2:="=IDCPS"

PriceDeviationWorkbook.Sheets("Missing Price").Range("F1").AutoFilter Field:=6, Criteria1:="<>OPEQINEX", Operator:=xlAnd, Criteria2:="<>OPEQTYEX"

'Remove true from A2333 cell'

PriceDeviationWorkbook.Sheets("High Deviation").Range("A2333").Value = ""

'We compare the prices inside of the High Deviation"

comparePrices PriceDeviationWorkbook.Sheets("High Deviation").Range("K:K")

Set templateWorkbook = Workbooks.Open("L:\Wells Fargo Prime Services\Operations\Pricing\EOD pricing.xlsx")

'Clear any data in templateWorkbook'

templateWorkbook.Sheets(1).Range("A2:F500").ClearContents

'start looping and passing values to template'

For lastValue = 2 To 2333
    If PriceDeviationWorkbook.Sheets("Missing Price").Range("A" & lastValue).Value = "" Then Exit For

            If PriceDeviationWorkbook.Sheets("Missing Price").Range("N" & lastValue).Value = "#N/A" Then
                
                PriceDeviationWorkbook.Sheets("Missing Price").Range("P" & lastValue).Value = PriceDeviationWorkbook.Sheets("Missing Price").Range("O" & lastValue).Value

                Else
                    
                    If PriceDeviationWorkbook.Sheets("Missing Price").Range("O" & lastValue) = "#N/A" Then
                            
                            PriceDeviationWorkbook.Sheets("Missing Price").Range("P" & lastValue).Value = PriceDeviationWorkbook.Sheets("Missing Price").Range("N" & lastValue).Value
                            
                       Else
                        
                        If PriceDeviationWorkbook.Sheets("Missing Price").Range("N" & lastValue).Value <> "#N/A" And PriceDeviationWorkbook.Sheets("Missing Price").Range("O" & lastValue).Value <> "#N/A" Then
                            
                            PriceDeviationWorkbook.Sheets("Missing Price").Range("P" & lastValue).Value = PriceDeviationWorkbook.Sheets("Missing Price").Range("N" & lastValue).Value

                    End If

            End If
               
    End If
    
Next lastValue


'Moving exclusion list into Missing Price worksheet '

Set exclusionWorkbook = Workbooks.Open("L:\Wells Fargo Prime Services\Operations\Pricing\Exclude List.xlsx")

'Find last row'

lastRow = exclusionWorkbook.Sheets(1).Range("A2").End(xlDown).Row

PriceDeviationWorkbook.Sheets("Missing Price").Range("R2" & ":R" & lastRow).Value = exclusionWorkbook.Sheets(1).Range("A2" & ":A" & lastRow).Value


'Pass Missing prices values to template'

For valuesInP = 2 To i
    
    If IsError(PriceDeviationWorkbook.Sheets("Missing Price").Range("P" & valuesInP).Value) = True Then
        
        PriceDeviationWorkbook.Sheets("Missing Price").Range("P" & valuesInP).Clear
        
        Else
        
        If PriceDeviationWorkbook.Sheets("Missing Price").Range("P" & valuesInP).Value <> "" Then
            
            If checkIfExcluded(PriceDeviationWorkbook, CInt(valuesInP)) = False Then
                       
                    moveToTemplate templateWorkbook, CInt(valuesInP), PriceDeviationWorkbook, "Missing Price"
        
        End If
    
    End If

End If

Next valuesInP

'Remove values excluded List from the R range'

PriceDeviationWorkbook.Sheets("Missing Price").Range("R:R").ClearContents

'Close workbook'

exclusionWorkbook.Close

'Remove any 'N/A Values'

For removeNA = 2 To lastValueOfFirstSheet
    
    If IsError(PriceDeviationWorkbook.Sheets("High Deviation").Range("M" & removeNA).Value) = True Then PriceDeviationWorkbook.Sheets("High Deviation").Range("M" & removeNA).Clear
        
        If PriceDeviationWorkbook.Sheets("High Deviation").Range("M" & removeNA).Value = "N/A" Then PriceDeviationWorkbook.Sheets("High Deviation").Range("M" & removeNA).Clear

Next removeNA

'Pass High Deviation values to template'
For Each cl In PriceDeviationWorkbook.Sheets("High Deviation").Range("M2" & ":M" & lastValueOfFirstSheet).SpecialCells(xlCellTypeVisible)
    
    If cl <> "" Then
    
    moveToTemplate templateWorkbook, CInt(cl.Row), PriceDeviationWorkbook, "High Deviation"
    
    End If
    
Next cl


Application.ScreenUpdating = True
Application.DisplayStatusBar = True


'Next Update: improve speed of macro by reducing copy and paste along with for statements if possible'


End Sub
Private Function checkIfExcluded(PriceDeviationWorkbook As Workbook, locationOfCell As Integer) As Boolean

    For i = 2 To 2333
        If PriceDeviationWorkbook.Sheets("Missing Price").Range("R" & i).Value = "" Then Exit For
        
            If PriceDeviationWorkbook.Sheets("Missing Price").Range("R" & i).Value = PriceDeviationWorkbook.Sheets("Missing Price").Range("A" & locationOfCell).Value Then
        
                checkIfExcluded = True
        
        End If
    
    Next i
    
End Function

Private Function moveToTemplate(templateWorkbook As Workbook, locationOfCell As Integer, priceWorkbook As Workbook, sheetName As String)

'Find next empty row'

For nextEmptyRow = 2 To 2333
    
    If templateWorkbook.Sheets(1).Range("A" & nextEmptyRow).Value = "" Then Exit For

Next nextEmptyRow


'Start passing the values'
templateWorkbook.Sheets(1).Range("A" & nextEmptyRow).Value = priceWorkbook.Sheets(sheetName).Range("A" & locationOfCell).Value
templateWorkbook.Sheets(1).Range("B" & nextEmptyRow).Value = "CUSIP"

If sheetName = "High Deviation" Then

templateWorkbook.Sheets(1).Range("C" & nextEmptyRow).Value = priceWorkbook.Sheets(sheetName).Range("M" & locationOfCell).Value

Else

templateWorkbook.Sheets(1).Range("C" & nextEmptyRow).Value = priceWorkbook.Sheets(sheetName).Range("P" & locationOfCell).Value

End If

templateWorkbook.Sheets(1).Range("D" & nextEmptyRow).Value = "USD"
templateWorkbook.Sheets(1).Range("E" & nextEmptyRow).Value = "IDC"
templateWorkbook.Sheets(1).Range("F" & nextEmptyRow).Value = Format(Date, "M/DD/YY")


End Function

Private Function comparePrices(target As Range)

                target.Worksheet.Range("K:K").NumberFormat = "General"
                target.Worksheet.Range("L:L").NumberFormat = "General"
                target.Worksheet.Range("M:M").NumberFormat = "General"
                target.Worksheet.Range("J:J").NumberFormat = "General"
                
                For i = 2 To 2333
                    
                    'Exit loop if there's no values'
                    If target.Worksheet.Range("I" & i).Value = "" Then Exit For
                    
                    'If both are N/A'
                        If IsError(target.Worksheet.Range("K" & i).Value) = True And IsError(target.Worksheet.Range("L" & i).Value) = True Then
                
                             target.Worksheet.Range("M" & i).Value = ""
            
                            Else
            'PRCUSD came up N/A so we take the iemid price'
                
                If IsError(target.Worksheet.Range("K" & i).Value) = True Then
                   
                    If target.Worksheet.Range("L" & i).Value <> target.Worksheet.Range("I" & i) Then
                        
                        target.Worksheet.Range("M" & i).Value = target.Worksheet.Range("L" & i).Value


                        End If
                    
                    Else
                
                'IEMID came up N/A so we take the PRCUSD price'

                        If IsError(target.Worksheet.Range("L" & i).Value) = True Then

                            If target.Worksheet.Range("K" & i).Value <> target.Worksheet.Range("I" & i).Value Then
                                
                                target.Worksheet.Range("M" & i).Value = target.Worksheet.Range("K" & i).Value

                            End If
                        
                            Else
                            
                            'We check if it is equal to either price is equal'
                            If target.Worksheet.Range("I" & i).Value = target.Worksheet.Range("K" & i).Value Or target.Worksheet.Range("I" & i).Value = target.Worksheet.Range("L" & i).Value Then
                                
                                target.Worksheet.Range("M" & i).Value = ""

      
                                Else
                                
                                'If price is not equal to neither of them'
                                
                                If target.Worksheet.Range("I" & i).Value <> target.Worksheet.Range("K" & i).Value And target.Worksheet.Range("I" & i).Value <> target.Worksheet.Range("L" & i).Value Then
                                        
                                    target.Worksheet.Range("M" & i).Value = target.Worksheet.Range("K" & i).Value

                                        End If
                                        
                                    End If
                                    
                                End If
                                
                            End If
                            
                        End If
                        
                    Next i
                    
End Function
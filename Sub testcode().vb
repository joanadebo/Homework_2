Sub testcode()

Dim iRows As Integer
Dim Total As Double  'variable for each credit card brand total
Dim Count As Integer
Dim Brand As String 'variable for distinct brand name
Dim Openstartprice As Double
Dim Closeendprice As Double
Dim SumTable As Integer
Dim j As Integer
Dim i As Double
Dim greatestval As Double
Dim lowestval As Double
Dim greatesttotal As Double



iRows = ActiveWorkbook.Worksheets.Count
greatestval = 0
lowestval = 0
greatesttotal = 0
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest Percent Increase"
Cells(3, 15).Value = "Greatest Percent Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Columns("I:Q").AutoFit     'Autofit text to column

'For loop for worksheet with number of rows = iRows
For Count = 1 To iRows

  Total = 0
  Openstartprice = 0
  Closeendprice = 0
   j = 0
  SumTable = 2
  lastRow = Cells(Rows.Count, 1).End(xlUp).Row


    'For loop for credit card purchases
     For i = 2 To lastRow
    

    'List each distinct brand, open start price, close start price when ticker symbol changes and add total for the same brand

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Brand = Cells(i, 1).Value
        Openstartprice = Cells(i, 3).Value
        Closeendprice = Cells(i, 6).Value
        Total = Total + Cells(i, 7).Value
        'Percent Yearly Stock
        Range("J" & SumTable).Value = (Closeendprice - Openstartprice)
       'Percent change
        Range("K" & SumTable).Value = ((Closeendprice - Openstartprice) / Openstartprice)
        
      'Conditional Formatting Yearly Stock
      If ((Closeendprice - Openstartprice) < 0) Then
        Range("J" & SumTable).Interior.ColorIndex = 3
    
      Else
      Range("J" & SumTable).Interior.ColorIndex = 4
  End If
  
      'Percent change
      Range("K" & SumTable).Value = ((Closeendprice - Openstartprice) / Openstartprice)
        
      ' Print Ticker Brand
      Range("I" & SumTable).Value = Brand

      ' Print Total by Ticker Brand
      Range("L" & SumTable).Value = Total

      ' Add one to the summary table row
      SumTable = SumTable + 1

      
      ' Reset the Brand Total
      Total = 0
      j = 0

    ' If brand name in cell  i+1 and i are the same brand...
    Else

      Total = Total + Cells(i, 7).Value
      j = j + 1
        
    End If
    
    If Range("K" & SumTable).Value > greatestval Then
    Cells(2, 16).Value = Range("I" & SumTable).Value
    Cells(2, 17).Value = Range("K" & SumTable).Value
    ElseIf Range("K" & SumTable).Value < lowestestval Then
    Cells(3, 16).Value = Range("I" & SumTable).Value
    Cells(3, 17).Value = Range("K" & SumTable).Value
    End If
    
    If Range("L" & SumTable).Value > greatesttotal Then
    Cells(4, 16).Value = Range("I" & SumTable).Value
    Cells(4, 17).Value = Range("L" & SumTable).Value
    End If
    
  Next i
  
Next Count

End Sub

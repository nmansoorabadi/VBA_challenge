Option Explicit
Sub stock_Analysis()

'loop through worksheet
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    ws.Activate



   ' Set variables
   Dim total As Double
   Dim i As Long
   Dim change As Single
   Dim j As Integer
   Dim start As Long
   Dim lastrow As Long
   Dim percentChange As Single
   Dim days As Integer
   Dim dailyChange As Single
   Dim averageChange As Single
   Dim find_value As Double
   Dim increase_number As Double
   Dim decrease_number As Double
   Dim volume_number  As Double
   
   
   
   
   
   ' Determine the Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "vol"

    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

   
   ' Set initial values
   j = 0
   total = 0
   change = 0
   start = 2
   
  ' Run through each worksheet
  
   lastrow = Cells(Rows.Count, "A").End(xlUp).Row
   For i = 2 To lastrow
       ' If ticker changes then print results
       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           
           total = total + Cells(i, 7).Value
           
           If total = 0 Then
               ' Insert Values
               Range("I" & 2 + j).Value = Cells(i, 1).Value
               Range("J" & 2 + j).Value = 0
               Range("K" & 2 + j).Value = "%" & 0
               Range("L" & 2 + j).Value = 0
           Else
               ' Find First non zero starting value
               If Cells(start, 3) = 0 Then
                   For find_value = start To i
                       If Cells(find_value, 3).Value <> 0 Then
                           start = find_value
                           Exit For
                       End If
                    Next find_value
               End If
               ' Calculate Change
               change = (Cells(i, 6) - Cells(start, 3))
               percentChange = Round((change / Cells(start, 3) * 100), 2)
              
               start = i + 1
               ' print the results
               Range("I" & 2 + j).Value = Cells(i, 1).Value
               Range("J" & 2 + j).Value = Round(change, 2)
               Range("K" & 2 + j).Value = "%" & percentChange
               Range("L" & 2 + j).Value = total
               ' colors formatting
               Select Case change
                   Case Is > 0
                       Range("J" & 2 + j).Interior.ColorIndex = 4
                   Case Is < 0
                       Range("J" & 2 + j).Interior.ColorIndex = 3
                   Case Else
                       Range("J" & 2 + j).Interior.ColorIndex = 0
               End Select
           End If
           ' reset variables
           total = 0
           change = 0
           j = j + 1
           days = 0
       Else
           total = total + Cells(i, 7).Value
       End If
   Next i
   
   ' CHALLENGES
   Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastrow)) * 100
   Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastrow)) * 100
   Range("Q4") = WorksheetFunction.Max(Range("L2:L" & lastrow))
 
   increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
   decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
   volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)
  

   ' Insert values
   Range("P2") = Cells(increase_number + 1, 9)
   Range("P3") = Cells(decrease_number + 1, 9)
   Range("P4") = Cells(volume_number + 1, 9)
   
Next ws

End Sub




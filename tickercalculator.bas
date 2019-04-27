Attribute VB_Name = "Module1"
'Brandon Coleman
'Data Analytics
'Homework #2
'4-27-2019


Sub tickercalculator()
  
  Dim year_open, year_close, year_change, Total_Volume, percent, lastrow As Double
  Dim Ticker_Name As String
  Dim Summary_Table_Row As Integer
  Dim Current As Worksheet
    
  For Each Current In Worksheets
    'reset variables for each worksheet
    Total_Volume = 0
    Summary_Table_Row = 2
    year_open = Current.Range("C2").Value
    lastrow = Current.Cells(Rows.Count, 1).End(xlUp).Row
  
    'set column headers
    Current.Range("I1").Value = "Ticker"
    Current.Range("J1").Value = "Yearly Change"
    Current.Range("K1").Value = "Percent Change"
    Current.Range("L1").Value = "Total Stock Volume"
    Current.Range("P1").Value = "Ticker"
    Current.Range("Q1").Value = "Value"
    Current.Range("O2").Value = "Greatest % Increase"
    Current.Range("O3").Value = "Greatest % Decrease"
    Current.Range("O4").Value = "Greatest Total Volume"
    Current.Columns("I:P").AutoFit
   
    
        For i = 2 To lastrow 'iterate through existing worksheet
      
            If Current.Cells(i + 1, 1).Value <> Current.Cells(i, 1).Value Then
            
                'Calculate Year-Change
                
                year_close = Current.Cells(i, 6).Value
                year_change = year_close - year_open
                Current.Range("J" & Summary_Table_Row).Value = year_change
                
                'set cell colors red or green
                If (year_change < 0) Then
                    Current.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                Else
                    Current.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                End If
                    
                'set percent change when year_open is not zero

                If (year_open <= 0) Then
                   'not applicable dividing by zero
                   
                   Current.Range("K" & Summary_Table_Row).Value = "n/a"
                   Current.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
               
                Else
                 
                  percent = (year_change / year_open)
                  Current.Range("K" & Summary_Table_Row).Value = percent
                  Current.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                  
               
                End If
        
                'set new year open
                year_open = Current.Cells(i + 1, 3)
                
                'grab and set stock ticker
                Ticker_Name = Current.Cells(i, 1).Value

                Current.Range("I" & Summary_Table_Row).Value = Ticker_Name
               
                'grab and set total volume
                Current.Range("L" & Summary_Table_Row).Value = Total_Volume

                Summary_Table_Row = Summary_Table_Row + 1

                Total_Volume = 0
                
            Else
                'Still Same Ticker Symbol Add Onto Volume
                Total_Volume = Total_Volume + Current.Cells(i, 7).Value

            End If

        Next i 'move to next cell
        'Set Greatest Increase Decrease and Volume By Passing Current Worksheet
        GetDecrease Current
        GetIncrease Current
        GetVolume Current
   
  Current.Columns("A:Q").AutoFit
  Next Current 'move to next worksheet

End Sub

Sub GetDecrease(sheetname As Worksheet)
Dim MinIncrease As Double
Dim FndRng, Workrange As Range

'find min in column K
MinIncrease = sheetname.Application.WorksheetFunction.Min(sheetname.Range("$K:$K"))
sheetname.Range("Q3") = MinIncrease
sheetname.Range("Q3").NumberFormat = "0.00%"
MinIncrease = MinIncrease * 100

'find fow of min and corresponding ticker
Set FndRng = sheetname.Range("K:K").Find(what:=MinIncrease, LookIn:=xlFormulas)
sheetname.Range("P3") = sheetname.Range("I" & FndRng.Row).Value
End Sub
Sub GetIncrease(sheetname As Worksheet)
Dim MaxIncrease, MinIncrease As Double
Dim FndRng As Range

'find max in column K
MaxIncrease = sheetname.Application.WorksheetFunction.Max(sheetname.Range("$K:$K"))
sheetname.Range("Q2") = MaxIncrease
sheetname.Range("Q2").NumberFormat = "0.00%"
MaxIncrease = MaxIncrease * 100

'find row of max and corresponding ticker
Set FndRng = sheetname.Range("K:K").Find(what:=MaxIncrease, LookIn:=xlFormulas)
sheetname.Range("P2") = sheetname.Range("I" & FndRng.Row).Value
End Sub

Sub GetVolume(sheetname As Worksheet)
Dim Volume As Double
Dim FndRng As Range

'Find Max In Column L
Volume = sheetname.Application.WorksheetFunction.Max(sheetname.Range("$L:$L"))
sheetname.Range("Q4") = Volume

'find row of max and corresponding ticker
Set FndRng = sheetname.Range("L:L").Find(what:=Volume, LookIn:=xlFormulas)
sheetname.Range("P4") = sheetname.Range("I" & FndRng.Row).Value
End Sub

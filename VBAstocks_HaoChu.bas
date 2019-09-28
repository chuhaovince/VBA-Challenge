Attribute VB_Name = "Module1"
Sub Stock():

    'To fasten the run-time
    'Turn off screenupdating, pagebreak,automatic calculatioins,events, at the beginning of the code
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    ActiveSheet.DisplayPageBreaks = False
    
    ' Declaration in memory
    ' use cell_loc variable to locate another row of ticker starts at
    'ws_count: the number of worksheets
    '---------------------------
    Dim tol_rows As Long, ticker_num As Integer, ws_count As Integer, first_row As Long, last_row As Long, index_max_increase As Long, index_max_decrease As Long, index_max_vol As Long
    
    'count the number of worksheets
    ws_count = ActiveWorkbook.Worksheets.count
    
    For x = 1 To ws_count
    
        ThisWorkbook.Worksheets(x).Activate

        ' Using Range Function
        '----------------------------
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("I1").Value = "<ticker>"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
                        
        'Count the number of tickers in <tikcer> column
        tol_rows = Cells(Rows.count, "A").End(xlUp).Row

        'Sort the ticker by labels first in case if we get some datasets with everything in a mess, then sort the date corresponding to those ticker labels
        'This part is redundent for this assignment
        With ActiveSheet.Sort
            .SortFields.Add Key:=Range("A1"), Order:=xlAscending
            .SortFields.Add Key:=Range("B1"), Order:=xlAscending
            .SetRange Range("A1", "C" & tol_rows)
            .Header = xlYes
            .Apply
        End With
        
        ' List of Ticker labels via advanced filter
        '---------------------------
        ActiveSheet.Range("A1", "A" & tol_rows).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("I1"), Unique:=True
                
        ' Using Range function to rename the header of Ticker
        Range("I1").Value = "Ticker"
    
        'Count the number of rows in Ticker column
        ticker_num = Cells(Rows.count, "I").End(xlUp).Row
        
        For i = 2 To ticker_num 'Ticker/tic_label
            
            'find the first and last row with the same ticker label
            With ActiveSheet
                first_row = Application.Match(Range("I" & i).Value, Range("A1", "A" & tol_rows), 0)
                last_row = Application.Match(Range("I" & i).Value, Range("A2", "A" & tol_rows), 0) + (Application.CountIf(Range("A2", "A" & tol_rows), Range("I" & i).Value))
            End With

            ' Calculate the yearly change, Calculate the totle volume of a specific ticker
            '----------------------------
            Range("J" & i).Value = Range("F" & last_row).Value - Range("C" & first_row).Value: Range("L" & i).Value = WorksheetFunction.Sum(Range("G" & first_row, "G" & last_row))
            
            ' Calculate the percent change of the year
            '------------------------------------------
            If Range("C" & first_row).Value = 0 Then 'If the open value is 0, we cant calculate the percentage mathmatically
                Range("K" & i).Value = 0
            Else
                Range("K" & i).Value = Range("J" & i).Value / Range("C" & first_row).Value * 100 & "%"
            End If
        
            ' Assign color to the yearly change
            With Range("J" & i)
                If Range("J" & i).Value > 0 Then
                    .Interior.ColorIndex = 4              'Green color
                Else
                    .Interior.ColorIndex = 3              'Red color
                End If
            End With
                    
        Next i
        
        'Find the max and min value of percentage change
        'Find the maximum value of stock volume
        Range("Q2").Value = WorksheetFunction.Max(Range("K2", "K" & ticker_num)) * 100 & "%"
        Range("Q3").Value = WorksheetFunction.Min(Range("K2", "K" & ticker_num)) * 100 & "%"
        Range("Q4").Value = WorksheetFunction.Max(Range("L2", "L" & ticker_num))
        
        'Find each of the corresponding ticker label
        
        index_max_increase = Application.Match(Range("Q2").Value, Range("K1", "K" & ticker_num), 0)
        index_max_decrease = Application.Match(Range("Q3").Value, Range("K1", "K" & ticker_num), 0)
        index_max_vol = Application.Match(Range("Q4").Value, Range("L1", "L" & ticker_num), 0)
              
        Range("P2").Value = Range("I" & index_max_increase).Value
        Range("P3").Value = Range("I" & index_max_decrease).Value
        Range("P4").Value = Range("I" & index_max_vol).Value
        
    Next x
    
    'Turn on screenupdating, pagebreak,automatic calculatioins,events, at the beginning of the code
    ActiveSheet.DisplayPageBreaks = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub


  
        
    
        


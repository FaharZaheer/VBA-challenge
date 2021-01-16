Attribute VB_Name = "Module1"

Sub Run_wall_street_On_all_sheets() ' got this from https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        'MsgBox ("running sheet # " + xSh.Name)
        Call Wall_Street
    Next
    Application.ScreenUpdating = True
End Sub

Sub Wall_Street()




Dim Range_str As String
Dim Ticker_counter_Start As Long
Dim Ticker_counter_stop As Long
Dim counter As Long
Dim Ticker_counter As Long
Dim Run_loop As Boolean
Dim num_of_rows As Long

Dim G_Increase_index As Integer
Dim G_Decrease_index As Integer
Dim G_TValue_index As Integer

Ticker_counter_Start = 0
Ticker_counter_stop = 0
num_of_rows = 0
Run_loop = True
Range_str = "$"

G_Increase_index = 0
G_Decrease_index = 0
G_TValue_index = 0

'''''''''''''''''''''''''''''''''''
    counter = 2
    Ticker_counter = 2
    Ticker_counter_Start = Ticker_counter
    
   
    ' Setting up the header
    Cells(1, 9).Value = "Ticker"
    Range("I1").EntireColumn.AutoFit
    Cells(1, 10).Value = "Yearly Change"
    Range("J1").EntireColumn.AutoFit
    Cells(1, 11).Value = "Percent Change"
    Range("K1").EntireColumn.AutoFit
    Cells(1, 12).Value = "Total Stock Volume"
    Range("L1").EntireColumn.AutoFit
        
    
    
    

   While Run_loop
 '   While counter < 10
           
             While Cells(Ticker_counter, 1).Value = Cells(Ticker_counter + 1, 1).Value
            
                 Ticker_counter = Ticker_counter + 1
                          
             Wend
        
            Ticker_counter_stop = Ticker_counter
            
            ' Ticker
                Cells(counter, 9).Value = Cells(Ticker_counter_Start, 1).Value
            
            
            ' Yearly Change
            
                Cells(counter, 10).Value = Cells(Ticker_counter_stop, 6).Value - Cells(Ticker_counter_Start, 3).Value
                
                    If Cells(counter, 10).Value >= 0 Then
                        Cells(counter, 10).Interior.ColorIndex = 4  ' for the green
                    Else
                        Cells(counter, 10).Interior.ColorIndex = 3  ' for the red
                    End If
            
            'Percent Change
            
                Cells(counter, 11).NumberFormat = "0.00%"
                If Cells(Ticker_counter_Start, 3).Value = 0 Then
                    'MsgBox ("Division by zero is not allow")
                    Cells(counter, 11).Value = "Undefined"
                Else
                     Cells(counter, 11).Value = (Cells(counter, 10).Value / Cells(Ticker_counter_Start, 3).Value)
                End If
            
            
            ' Total Stock Volume
            
                Range_str = "G" + CStr(Ticker_counter_Start) + ":G" + CStr(Ticker_counter_stop)
                Cells(counter, 12).Value = Excel.WorksheetFunction.Sum(Range(Range_str))
            
                  
           If IsEmpty(Cells(Ticker_counter_stop + 1, 1).Value) = True Then
              Run_loop = False
           End If
           
           counter = counter + 1
           Ticker_counter_Start = Ticker_counter_stop + 1
           Ticker_counter = Ticker_counter_Start
    
    Wend
    
    '''''''''''''Bonus ''''''''''''''
   
    Cells(1, 16).Value = "Ticker"
    Range("P1").EntireColumn.AutoFit
    
    Cells(1, 17).Value = "Value"
    Range("Q1").EntireColumn.AutoFit
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Range("O1").EntireColumn.AutoFit
    
    
    num_of_rows = Cells(Rows.Count, 9).End(xlUp).Row
    Range_str = "K2:K" + CStr(num_of_rows)
        
    ' Greatest % Increase
        
    G_Increase_index = Excel.WorksheetFunction.Match(Excel.WorksheetFunction.Max(Range(Range_str)), Range(Range_str), 0)
    Cells(2, 16).Value = Cells(G_Increase_index + 1, 9).Value
    Cells(2, 17).Value = Cells(G_Increase_index + 1, 11).Value
    Cells(2, 17).NumberFormat = "0.00%"
    
    ' Greatest % Decrease
   
    G_Decrease_index = Excel.WorksheetFunction.Match(Excel.WorksheetFunction.Min(Range(Range_str)), Range(Range_str), 0)
    Cells(3, 16).Value = Cells(G_Decrease_index + 1, 9).Value
    Cells(3, 17).Value = Cells(G_Decrease_index + 1, 11).Value
    Cells(3, 17).NumberFormat = "0.00%"
    
    ' Greatest Total Volume
    Range_str = "L2:L" + CStr(num_of_rows)
    G_TValue_index = Excel.WorksheetFunction.Match(Excel.WorksheetFunction.Max(Range(Range_str)), Range(Range_str), 0)
    Cells(4, 16).Value = Cells(G_TValue_index + 1, 9).Value
    Cells(4, 17).Value = Cells(G_TValue_index + 1, 12).Value
    
    
    'MsgBox (num_of_rows)
    '''''''''''''''''''''''''''''''''
    
    


End Sub

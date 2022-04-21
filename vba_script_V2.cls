VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub submission_v2()

For Each ws In Worksheets
    
    'headers
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    'set up ticker column
    Dim Ticker As String
    Dim lastrow As Double

    
    'volume
    Dim volume As Double
    
    'stock changes
    Dim stock_close As Double
    Dim stock_open As Double
    Dim Yearly_change As Double
    Dim Percent_change As Double
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
volume = 0

Dim Summary_table As Double
Summary_table = 2
   

   
For n = 2 To lastrow
    
    
    If ws.Cells(n + 1, 1).Value <> ws.Cells(n, 1) Then
        
        Ticker = ws.Cells(n, 1).Value
            ws.Range("I" & Summary_table).Value = Ticker

        volume = volume + ws.Cells(n, 7).Value
            ws.Range("L" & Summary_table).Value = volume
        volume = 0

        stock_close = ws.Cells(n, 6)
        
        'column yearly change and percent change
        If stock_open = 0 Then
            
            Yearly_change = 0
            Percent_change = 0
        
        Else
            
            Yearly_change = stock_close - stock_open
            Percent_change = (stock_close - stock_open) / stock_open
            
        End If
            
                'input yearly_change
                ws.Range("J" & Summary_table).Value = Yearly_change
                'input percent_change
                ws.Range("K" & Summary_table).Value = Percent_change
                
                'formatting percent_change for appropriate output of values
                ws.Range("k" & Summary_table).Style = "Percent"
                ws.Range("k" & Summary_table).NumberFormat = "0.00%"
                      
                Summary_table = Summary_table + 1
        
        'set up system to continue gatherin values for other condition
        ElseIf ws.Cells(n - 1, 1).Value <> ws.Cells(n, 1) Then
            
            stock_open = ws.Cells(n, 3)
        
        Else: volume = volume + ws.Cells(n, 7).Value
        
    End If
        
Next n
    
'color in column Yearly_change

    For i = 2 To lastrow
        
        If ws.Range("J" & i).Value > 0 Then
            
            ws.Range("J" & i).Interior.ColorIndex = 4
            
        ElseIf ws.Range("J" & i).Value < 0 Then
            
            ws.Range("J" & i).Interior.ColorIndex = 3
            
        End If
        
    Next i
    
    'bonus section
        
        'column headers
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        
        'row headers
        ws.Cells("2,15") = "Greatest % Increase"
        ws.Cells("3,15") = "Greatest % Decrease"
        ws.Cells("4,15") = "Greatest Total Volume"
        
        Dim great_increase As Double
        Dim great_decrease As Double
        Dim great_total_volume As Double
        
        great_increase = 0
        great_decrease = 0
        great_total_volume = 0
        
        'greatest increase
        For a = 2 To lastrow
            
            If ws.Cells(a, 11).Value > great_increase Then
                
                great_increase = ws.Cells(a, 11).Value
                ws.Range("q2").Value = great_increase
                ws.Range("q2").Style = "Percent"
                ws.Range("q2").NumberFormat = "0.00%"
                ws.Range("p2").Value = ws.Cells(a, 9).Value
            
            End If
            
        Next a
        
        'greatest decrease
        For b = 2 To lastrow
            
            If ws.Cells(b, 11).Value < great_decrease Then
                
                great_decrease = ws.Cells(b, 11).Value
                ws.Range("Q3").Value = great_decrease
                ws.Range("Q3").Style = "Percent"
                ws.Range("Q3").NumberFormat = "0.00%"
                ws.Range("P3").Value = ws.Cells(b, 9).Value
            
            End If
            
        Next b
        
        'greatest total volume
        For c = 2 To lastrow
            
            If ws.Cells(c, 12).Value > great_total_volume Then
                
                great_total_volume = ws.Cells(c, 12).Value
                ws.Cells(4, 17).Value = great_total_volume
                ws.Cells(4, 16).Value = ws.Cells(c, 9).Value
                
            End If
        
        Next c
    
ws.Range("A:Q").Columns.AutoFit



Next ws

End Sub

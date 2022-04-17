VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub final_submission()


    Dim n As Integer
    Dim sum_data As Integer
    Dim outrow As Integer
    Dim ychange As Double
    Dim openprice As Double
    Dim closeday As Double
    Dim perchange As Double
    Dim yearly_change As Double
    Dim daily_change As Double
    Dim stock_vol As Double
    Dim day_vol As Double
    Dim lastrow As Long
    
    Dim ws As Worksheet
    Dim wb As Workbook
    
    
    Set wb = ActiveWorkbook
    
    
    For Each ws In Worksheets
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
       
    
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        sum_data = 2
        outrow = 2
        perchange = 0
        yearly_change = 0
        openprice = ws.Cells(2, 3).Value
        daily_change = o
        stock_vol = 0
        day_vol = 0
        
            For n = 2 To lastrow
                
                If ws.Cells(n + 1, "A").Value <> ws.Cells(n, "A").Value Then
                    ws.Cells(outrow, "I").Value = ws.Cells(n, "A").Value
                    outrow = outrow + 1
                
                    closeday = ws.Cells(n, "F").Value
                    openday = ws.Cells(n, "c").Value
                    daily_change = closeday - openday
                    yearly_change = yearly_change + daily_change
                    
                    ws.Cells(n, "J").Value = yearly_change
                    yearly_change = yearly_change + 1
                    
                    perchange = (daily_change / openprice) * 100
                    ws.Range("k" & sum_data).Value = perchange
                    
                    
                    'total volume: value in the column "g" added to each other in same ticker.
                    stock_vol = stock_vol + ws.Cells(n, "g").Value
                    'print stock volume to table on column g
                    ws.Range("L" & sum_data).Value = stock_vol
                    
                    sum_data = sum_data + 1
                    yearly_change = 0
                    perchange = 0
                    stock_vol = 0
                 
                Else
                    stock_vol = stock_vol + mainws.Cells(n, "g").Value
                
                    
                End If
                     
        
                If ws.Cells(ychange, "J").Value > 0 Then
                        ws.Cells(ychange, "J").Interior.ColorIndex = 4
                        
                ElseIf ws.Cells(ychange, "J").Value < 0 Then
                        ws.Cells(ychange, "J").Interior.ColorIndex = 3
                    
                End If
        
               
            Next n
        
        ws.Columns("I:L").AutoFit
        
    Next ws
       
End Sub




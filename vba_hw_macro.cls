VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub test_run()


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
    Dim day_vol As Integer
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    

    sum_data = 2
    outrow = 2
    perchange = 0
    yearly_change = 0
    openprice = Cells(2, 3).Value
    daily_change = o
    stock_vol = 0
    day_vol = 0
    
    For n = 2 To lastrow
        
        If Cells(n + 1, "A").Value <> Cells(n, "A").Value Then
            Cells(outrow, "I").Value = Cells(n, "A").Value
            outrow = outrow + 1
        
            closeday = Cells(n, "F").Value
            openday = Cells(n, "c").Value
            daily_change = closeday - openday
            yearly_change = yearly_change + daily_change
            
            Cells(n, "J").Value = yearly_change
            yearly_change = yearly_change + 1
            
            perchange = (daily_change / openprice) * 100
            Range("k" & sum_data).Value = perchange
            
            
            'total volume: value in the column "g" added to each other in same ticker.
            stock_vol = stock_vol + Cells(n, "g").Value
            'print stock volume to table on column g
            Range("L" & sum_data).Value = stock_vol
            
            sum_data = sum_data + 1
            yearly_change = 0
            perchange = 0
            stock_vol = 0
         
        Else
            stock_vol = stock_vol + Cells(n, "g").Value
        
            
        End If
             

        If Cells(ychange, "J").Value > 0 Then
                Cells(ychange, "J").Interior.ColorIndex = 4
                
        ElseIf Cells(ychange, "J").Value < 0 Then
                Cells(ychange, "J").Interior.ColorIndex = 3
            
        End If

       
    Next n
    
       
End Sub



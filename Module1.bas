Attribute VB_Name = "Module1"
Sub VBA_Challange_2()
'    For Each ws In Worksheets
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        Dim WorksheetName As String
        Dim Ticker_Name As String
        Dim Ticker_vol As Double
        Ticker_vol = 0
        
        Dim open_price As Double
        open_price = Cells(2, 3).Value
        
        Dim Ticker_Uniq As Integer
        Ticker_Uniq = 2

        Dim close_price As Double
        Dim Yearly_change As Double
        Dim Percent_change As Double
        
        WorksheetName = ws.Name
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Change_Yearly"
        ws.Cells(1, 11).Value = "Percentage_Change"
        ws.Cells(1, 12).Value = "Total_Volume"
        
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
        Cells(2, 14).Value = "Greatest_Increase%"
        Cells(3, 14).Value = "Greatest_Decrease%"
        Cells(4, 14).Value = "Total_Volume"

        
        Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To Lastrow
            
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker_Name = Cells(i, 1).Value
            
            Ticker_vol = Ticker_vol + Cells(i, 7).Value
            Range("i" & Ticker_Uniq).Value = Ticker_Name
       
            Range("L" & Ticker_Uniq).Value = Ticker_vol
            Ticker_vol = Ticker_vol + 1
        
            close_price = Cells(i, 6).Value
         
            Yearly_change = (close_price - open_price)
            Range("J" & Ticker_Uniq).Value = Yearly_change

            Ticker_Uniq = Ticker_Uniq + 1

            summary_ticker_row = summary_ticker_row + 1
          
            Ticker_vol = 0
            open_price = Cells(i + 1, 3)


         Else

               Ticker_vol = Ticker_vol + Cells(i, 7).Value

               Percent_change = Yearly_change / open_price

               Range("K" & Ticker_Uniq).Value = Percent_change
               Range("K" & Ticker_Uniq).NumberFormat = "0.0%"

                                     
            If Cells(i, 10).Value >= 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
               If Cells(i, 10).Value < 0 Then
               Cells(i, 10).Interior.ColorIndex = 3
             
                          
                                            
            End If
   
          End If
            
            
             End If
             
             Next i
             Next ws

        
End Sub


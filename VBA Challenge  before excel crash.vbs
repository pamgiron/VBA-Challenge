Attribute VB_Name = "Module1"

Sub Stockdata():

    For Each WS In Worksheets
        WS.Activate
        
        WS.range("I1").Value = "Ticker"
        WS.range("J1").Value = "Yearly Change"
        WS.range("K1").Value = "Percent Change"
        WS.range("L1").Value = "Total Stock Volume"
        
        total_vol = 0
        ipointer = 2
        

        cpointer = 2
        fpointer = 2
        
        lastrow = WS.Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To lastrow
        
            If WS.Cells(i + 1, "A").Value <> WS.Cells(i, "A").Value Then
            
                total_vol = total_vol + Cells(i, "G").Value
                
                Cells(ipointer, "I").Value = WS.Cells(i, "A").Value
                
                Cells(ipointer, "L").Value = total_vol
                
                ipointer = ipointer + 1
                
            Else
                total_vol = total_vol + Cells(i, "G").Value
                
                
            End If
            
        
      Next i
    Next WS
    
        
                    
                    
    
            
   
End Sub

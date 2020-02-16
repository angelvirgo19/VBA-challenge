Sub Stock_AnalysisTest()

    
    Dim ws As Worksheet
    Dim Summary_TableHeader As Boolean
    Dim COMMAND_SPREADSHEET As Boolean
    
    Summary_TableHeader = True
    COMMAND_SPREADSHEET = True
    
    
    For Each ws In Worksheets
    
        
        Dim Ticker_Name As String
        Ticker_Name = " "
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        Dim Year_OpenPrice As Double
        Year_OpenPrice = 0
        Dim Year_ClosePrice As Double
        Year_ClosePrice = 0
        Dim Yearly_change As Double
        Yearly_change = 0
        Dim Percent_change As Double
        Percent_change = 0
        Dim Max_Ticker As String
        Max_Ticker = " "
        Dim Min_Ticker As String
        Min_Ticker = " "
        Dim Max_percent As Double
        Max_percent = 0
        Dim Min_percent As Double
        Min_percent = 0
        Dim Max_Vol_Ticker As String
        Max_Vol_Ticker = " "
        Dim Max_Vol As Double
        Max_Vol = 0
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        Dim Lastrow As Long
        Dim i As Long
        
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        
        If Summary_TableHeader Then
            
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
        Else
            
            Summary_TableHeader = True
        End If
        
        
        Year_OpenPrice = ws.Cells(2, 3).Value
        
        For i = 2 To Lastrow
        
      
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                
                Ticker_Name = ws.Cells(i, 1).Value
                
               
                Year_ClosePrice = ws.Cells(i, 6).Value
                Yearly_change = Year_ClosePrice - Year_OpenPrice
            
                If Year_OpenPrice <> 0 Then
                    Percent_change = (Yearly_change / Year_OpenPrice) * 100
                Else
         
                    Percent_change = 0
                End If
                
       
                Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
              
            
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                ws.Range("J" & Summary_Table_Row).Value = Yearly_change
               
                If (Yearly_change > 0) Then
                    
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Yearly_change <= 0) Then
                   
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                 
                ws.Range("K" & Summary_Table_Row).Value = (CStr(Percent_change) & "%")
               
                ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                
               
                Summary_Table_Row = Summary_Table_Row + 1
               
                Yearly_change = 0
                
                Year_ClosePrice = 0
                
                Year_OpenPrice = ws.Cells(i + 1, 3).Value
              
                
                
                If (Percent_change > Max_percent) Then
                    Max_percent = Percent_change
                    Max_Ticker = Ticker_Name
                ElseIf (Percent_change < Min_percent) Then
                    Min_percent = Percent_change
                    Min_Ticker = Ticker_Name
                End If
                       
                If (Total_Ticker_Volume > Max_Vol) Then
                    Max_Vol = Total_Ticker_Volume
                    Max_Vol_Ticker = Ticker_Name
                End If
                
               
                Percent_change = 0
                Total_Ticker_Volume = 0
                
          
            Else
                
                Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
            End If
          
      
        Next i

            If COMMAND_SPREADSHEET Then
            
                ws.Range("Q2").Value = (CStr(Max_percent) & "%")
                ws.Range("Q3").Value = (CStr(Min_percent) & "%")
                ws.Range("P2").Value = Max_Ticker
                ws.Range("P3").Value = Min_Ticker
                ws.Range("Q4").Value = Max_Vol
                ws.Range("P4").Value = Max_Vol_Ticker
                
            Else
                COMMAND_SPREADSHEET = False
            End If
        
     Next ws
End Sub



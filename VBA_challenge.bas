Attribute VB_Name = "Module1"
Sub VBA_loop_Ticker()

    'set variables for problem
    Dim ws As Worksheet
    
    'loop through each worksheet in the workbook
    For Each ws In Worksheets
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("L:L").NumberFormat = "#,##0.00"
    
    'Declare ticker for ticker name and create place holder
    Dim ticker As String
    ticker = ""
    
    'Declare variables for figures
    Dim openingprice As Double
    openingprice = 0
    Dim closingprice As Double
    closingprice = 0
    Dim tickertotal As Double
    tickertotal = 0
    Dim yearlychange As Double
    yearlychange = 0
    Dim percentchange As Double
    percentchange = 0
    
    
    'Location for each ticker
    Dim summarytablerow As Long
    summarytablerow = 2
    
    'Set the count for rows from first to last row
    Dim lastrow As Long
    Dim i As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'set headings for summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly change"
    ws.Range("K1").Value = "Percent change"
    ws.Range("L1").Value = "Total stock volume"
    
   ' setting location for opening price
    openingprice = ws.Cells(2, 3).Value
    
   'For loop from the beginning of each row until the last row
    For i = 2 To lastrow
       
    
        'Check if the ticker code is different to ticker name below
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Output ticker and insert ticker data into summary table
            ticker = ws.Cells(i, 1).Value
            
            'setting location of closing price
            closingprice = ws.Cells(i, 6).Value
            
            'Calculate closing price and yearly change in price (from opening price to closing price)
            yearlychange = (closingprice - openingprice)
            
            'Calculate the percentage change + Check that open price is not 0 to avoid error
            If openingprice <> 0 Then
                percentchange = (yearlychange / openingprice)
        
            Else
            MsgBox ("For " & ticker & "the opening price is <= 0. Please manually fix the spreadsheet.")
            End If
    
                'calculation for the total of the ticker volume
                    tickertotal = tickertotal + ws.Cells(i, 7).Value
        
                'output the ticker name and total volume to summary table column I
                    ws.Range("I" & summarytablerow).Value = ticker
        
                'output the yearly change in price to summary table column J
                    ws.Range("J" & summarytablerow).Value = yearlychange
                 
                'output percentage change to summary table to column K
                    ws.Range("K" & summarytablerow).Value = percentchange
                   
                
                'output total stock volume to summary table to columm L
                    ws.Range("L" & summarytablerow).Value = tickertotal
                       
    
                'highlight numbers depending on positive change (green) or negative change( red)
                 If (yearlychange > 0) Then
                  ws.Range("J" & summarytablerow).Interior.ColorIndex = 4
                        
                ElseIf (yearlychange < 0) Then
                 ws.Range("J" & summarytablerow).Interior.ColorIndex = 3
                        
                
                End If
                
                If (percentchange > 0) Then
                    ws.Range("K" & summarytablerow).Interior.ColorIndex = 4
                
                ElseIf (percentchange < 0) Then
                    ws.Range("K" & summarytablerow).Interior.ColorIndex = 3
                
                
                End If
                  
    
                'Add row to summary table for next ticker
                    summarytablerow = summarytablerow + 1
                 
            
                'reset counter for yearly change, percent change and ticker total
                    yearlychange = 0
                    percentchange = 0
                    tickertotal = 0
            
                'Find next ticker's opening price
                    openingprice = ws.Cells(i + 1, 3).Value
         
                Else
                    tickertotal = tickertotal + ws.Cells(i, 7).Value
         
             End If
         
        Next i
        
    
    Next ws
          
End Sub


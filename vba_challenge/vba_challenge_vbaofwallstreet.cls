VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ticker()

'----Cycle Though all Worksheets
    
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Select
        

'----initialize our variables
    
        'Set variable for ticker symbol
        Dim tickerSymbol As String
    
        'Initialize the Total Stock Volume variable
        Dim totalStock As LongLong
        totalStock = 0
  
        'Initialize the percent change variable
        Dim percentChange As Double
  
        'Initialize variable for first opening value of ticker
        Dim firstValue As Double
  
        'Initialize varible for last closing value of ticker
        Dim lastValue As Double
  
        'Initialize yearly change
        Dim yearlyChange As Double
  
        'Keep track of the location for each entry in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        
        'Count the number of rows for our loop
        Dim RowCount As Long
        RowCount = Cells(Rows.Count, 2).End(xlUp).Row
    
    
    
    
'----Make the formatting appropriate
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Columns("K").NumberFormat = "###.##%"
    
    
    
'----execute output stuff here

        'Set first opening value outside of the loop
        firstValue = Cells(2, 3).Value
    
    
        'Loop through all rows
        For I = 2 To RowCount

            'Check if we are still within the same ticker, if not then...
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

                'Set the ticker symbol
                tickerSymbol = Cells(I, 1).Value
            
                'Add to the stock volume to the total
                totalStock = totalStock + Cells(I, 7).Value
      
                'Store last value for the particular ticker
                lastValue = Cells(I, 6).Value
        
                'Calculate the yearly change
                yearlyChange = lastValue - firstValue
            
            
                'Calculate the percent change
                If firstValue = 0 Then
                    percentChange = 0
                Else
                    percentChange = (yearlyChange / firstValue)
                End If



                'Print the ticker in the Summary Table
                Range("I" & Summary_Table_Row).Value = tickerSymbol
      
                'Print the yearly change in the Summary Table
                Range("J" & Summary_Table_Row).Value = yearlyChange
                
                
                If Range("J" & Summary_Table_Row).Value < 0 Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
                
      
                'Print the percent change in summary Table
                Range("K" & Summary_Table_Row).Value = percentChange
    
                'Print the total stock volume for the ticker in the summary table
                Range("L" & Summary_Table_Row).Value = totalStock

                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1

                'Reset initial value of year
                firstValue = Cells(I + 1, 3).Value
            
                'Reset total stock volume
                totalStock = 0
    
            'If the cell immediately following a row has the same ticker then...
            Else

                'Just add the stock volume to the total
                totalStock = totalStock + Cells(I, 7).Value

            End If

        Next I
        
    Next ws
   
End Sub

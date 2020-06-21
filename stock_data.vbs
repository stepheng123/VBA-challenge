Sub year_stock_data()

' Set new variables for Moderate Solution Part
Dim Open_Price As Double
Dim Close_Price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim ws As Worksheet
Dim Spreadsheet As Boolean

Spreadsheet = True
Open_Price = 0
Close_Price = 0
yearly_change = 0
percent_change = 0

' Loop through all of the worksheets in the active workbook.
    For Each ws In Worksheets
    
' Set variable for holding the ticker name
   Dim Ticker As String
   Ticker = " "
   Dim Total_Volume As Double
   Total_Volume = 0
        
        
' location of ticker name in the summary tabe
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
        
  ' Set row count
   Dim Lastrow As Long
   Dim i As Long
        
      ' Find the last row in the list
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

            ' Set Titles for the Summary Table
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"


        ' Set value of Open Price
        Open_Price = ws.Cells(2, 3).Value
        
        ' Loop from the beginning to last row
        For i = 2 To Lastrow
        
      
            ' Check if we are still within the same ticker name

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' find the values in each worksheet
                Ticker = ws.Cells(i, 1).Value
                Close_Price = ws.Cells(i, 6).Value
                yearly_change = Close_Price - Open_Price
                ' Check Division by 0 condition
                If Open_Price <> 0 Then
                    percent_change = (yearly_change / Open_Price) * 100
                Else

                End If
                
                ' Add to the Ticker name total volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
              
                
                ' Print the Ticker in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ' Print the Year change in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = yearly_change
                ' Set "Yearly Change" to highlight positive and negative results
                If (yearly_change > 0) Then
                    'Fill column with GREEN for positive
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (yearly_change <= 0) Then
                    'Fill column with RED for negative
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                 ' Print the Percentage change in the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = (CStr(percent_change) & "%")
                ' Print the Total stock volume in the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Total_Volume
                'autofit columns
                ws.Columns("J:L").EntireColumn.AutoFit
                
                ' Add 1 to the summary table row count
                Summary_Table_Row = Summary_Table_Row + 1
                ' Reset
                yearly_change = 0
                Close_Price = 0
                ' Capture next Ticker's Open_Price
                Open_Price = ws.Cells(i + 1, 3).Value
                    
            
            'Else - If a row is still the same ticker name,
            
            Else
                ' Encrease the Total Ticker Volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            End If
            
      
        Next i

        
     Next ws
End Sub


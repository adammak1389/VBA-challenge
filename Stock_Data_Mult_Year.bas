Attribute VB_Name = "Module1"
Sub Stock_Data()

    ' Declare all of your variable types here
    Dim ws As Worksheet
    Dim Ticker_symbol As String
    Dim yearly_Change As Double
    Dim open_Price As Double
    Dim Close_Price As Double
    Dim Percent_Change As Double
    Dim Vol As Double
    Dim LastRow As Long
    
    ' Loop through all sheets
    For Each ws In Worksheets
        ' Determine the last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' Loop through all the rows in the worksheet
        For i = 2 To LastRow
        For j = 1 To 7
            j = Column
        
            
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                Ticker_symbol = Cells(i, Column).Value
            End If
            
            yearly_Change = (Close_Price - open_Price)
            Percent_Change = (Close_Price - open_Price) / open_Price
            Vol = Cells(i + 1, 7).Value
            

            Cells(1, 9) = (yearly_Change.Value)
            Cells(1, 10) = (Percent_Change.Value)
            Cells(1, 11) = (Vol.Value)
            
            
        ' Go to the next row
        Next i
        
        
    
    ' Go to next worksheet
    Next ws
    
End Sub

End Sub

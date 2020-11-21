Sub dotoallworksheets()
'Declare the variable to iterate the formula for each worksheet

    Dim ws As Worksheet
'for worksheet
    For Each ws In Worksheets
        ws.Activate

'set initial variables to hold ticker, yearly change, opening price and closure price
        Dim ticker_name As String
        Dim yearly_change As Double
        Dim opening_price As Double
        Dim closure_price As Double


        ' Set an initial variable for holding the total volume per ticker
        Dim volume_Total As Double
        volume_Total = 0

        Dim tmp_volume_max_Total As Double
        tmp_volume_max_Total = 0

        Dim current_percentage As Double

 'Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly change"
        Range("K1").Value = "Percent change"
        Range("L1").Value = "Total stock volume"


'find the last row of active excel worksheet
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' Loop through all ticker, and print yearly change
        opening_price = Cells(2, 3)

        Dim tmp_increase_percentage As Double
        tmp_increase_percentage = 0
        Dim tmp_decrease_percentage As Double
        tmp_decrease_percentage = 0

        Dim tmp_increase_ticker As String
        Dim tmp_decrease_ticker As String
        Dim tmp_volume_max_ticker As String
    'For single sheet
        For i = 2 To lastrow
' 'Check if it is not the same ticker...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
' 'Set the ticker name
                ticker_name = Cells(i, 1).Value
'      'add to the volume total
                volume_Total = volume_Total + Cells(i, 7).Value

      'Print the ticker e volume total in the Summary Table
                Range("I" & Summary_Table_Row).Value = ticker_name
                Range("L" & Summary_Table_Row).Value = volume_Total
                Range("J" & Summary_Table_Row).Value = opening_price - Cells(i, 6).Value

                If opening_price <> 0 Then


                    current_percentage = Round((opening_price - Cells(i, 6).Value) / opening_price * 100, 2)
                    Range("K" & Summary_Table_Row).Value = current_percentage

                    If (current_percentage > tmp_increase_percentage) Then
                        tmp_increase_percentage = current_percentage
                        tmp_increase_ticker = ticker_name
                    End If

                    If (tmp_decrease_percentage > current_percentage) Then
                        tmp_decrease_percentage = current_percentage
                        tmp_decrease_ticker = ticker_name
                    End If

                    If (volume_Total > tmp_volume_max_Total) Then
                        tmp_volume_max_Total = volume_Total
                        tmp_volume_max_ticker = ticker_name
                    End If
            'Set color of the cells
                    If Range("K" & Summary_Table_Row) >= 0 Then
                        Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                    Else
                        Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                Else
                    Range("K" & Summary_Table_Row).Value = nan
                End If
      'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
       ' Reset the ticker name
                opening_price = Cells(i + 1, 3).Value
                volume_Total = 0
            Else
'      ' Add to the volume total
                volume_Total = volume_Total + Cells(i, 7).Value
            End If
        Next i


'MAX MIN quantities

        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % increase"
        Range("O3").Value = "Greatest % decrease"
        Range("O4").Value = "Greatest Total Volume decrease"
        Range("Q2").Value = tmp_increase_percentage
        Range("Q3").Value = tmp_decrease_percentage
        Range("Q4").Value = tmp_volume_max_Total
        Range("P2").Value = tmp_increase_ticker
        Range("P3").Value = tmp_decrease_ticker
        Range("P4").Value = tmp_volume_max_ticker
    Next ws

End Sub

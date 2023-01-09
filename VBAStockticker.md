  # VBA-challenge
  
  Sub Ticker_count()
  'Loop through all the worksheets
      
      Dim ws As Worksheet

      For Each ws In ThisWorkbook.Worksheets

          Dim Worksheetname As String
          Dim Lastrow As Double
          Dim Tickername As String
          Dim open_price As Double
          Dim close_price As Double
          Dim Volume As Integer
          Dim Summary_Table_Row As Long

         'When I was running the script through all the worksheets, it was giving the same output to all the sheets. By doing google search, came to know the reason due to active sheet. 'ws.Activate command make the each sheet run separately.

         ws.Activate

               'Assign each cell name in Summary table

          ws.Cells(1, 9).Value = "Tickername"
          ws.Cells(1, 10).Value = "Yearly change"
          ws.Cells(1, 11).Value = "Percent Change"
          ws.Cells(1, 12).Value = "Total Stock Vol"
          ws.Cells(1, 16).Value = "Value"
          ws.Cells(2, 16).Value = "Greatest % Increase"
          ws.Cells(3, 16).Value = "Greatest % Decrease"
          ws.Cells(4, 16).Value = "Greatest Total Volume"
          ws.Cells(1, 17).Value = "Ticker"

          'Last row info
          Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

          Worksheetname = ws.Name

          SummaryTableRow = 2
          Yearly_change = 0
          Percent_change = 0
          TotalVolume = 0
          
          
          'Calculations: Yearly change = Close price - open price
          ' % change = (close price- open price/ open price)* 100
          
           For Row_counter = 2 To Lastrow

              If ws.Cells(Row_counter + 1, 1).Value <> ws.Cells(Row_counter, 1).Value Then

                  Yearly_change = Yearly_change + (Cells(Row_counter, 6).Value - Cells(Row_counter, 3).Value)
                  Percent_change = Percent_change + ((Cells(Row_counter, 6).Value - Cells(Row_counter, 3).Value) / Cells(Row_counter, 3).Value) * 100
                  TotalVolume = TotalVolume + Cells(Row_counter, 7)
                  Tickername = ws.Cells(Row_counter, 1).Value

                  ws.Range("I" & SummaryTableRow).Value = Tickername
                  ws.Range("J" & SummaryTableRow).Value = Yearly_change
                  ws.Range("L" & SummaryTableRow).Value = TotalVolume
                  ws.Range("K" & SummaryTableRow).Value = Percent_change
                  SummaryTableRow = SummaryTableRow + 1


                  Yearly_change = 0
                  TotalVolume = 0
                  Percent_change = 0
           Else
                  Yearly_change = Yearly_change + (Cells(Row_counter, 6).Value - Cells(Row_counter, 3).Value)
                  TotalVolume = TotalVolume + Cells(Row_counter, 7)
                  Percent_change = Percent_change + ((Cells(Row_counter, 6).Value - Cells(Row_counter, 3).Value) / Cells(Row_counter, 3).Value) * 100

          End If
              'To assign the color code to Yearly Increse column
              
              If ws.Cells(Row_counter, 10).Value < 0 Then
                  ws.Cells(Row_counter, 10).Interior.ColorIndex = 3
              Else:
                  ws.Cells(Row_counter, 10).Interior.ColorIndex = 4

              End If

          'Bonus activity
          'To compare the highest or lowest percent value, I took first cell value of the series to test for condition true/false. Same with Greatest stock volume
          
          Next Row_counter

              Lastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row

              Dim tName As String

              Max_percentvalue = Cells(2, 11).Value
              Min_percentvalue = Cells(2, 11).Value
              Greatest_TotalValue = Cells(2, 12).Value
              tName = Cells(2, 9).Value


              For Rownum = 2 To Lastrow

                      If ws.Cells(Rownum, 11).Value > Max_value Then
                          Max_percentvalue = ws.Cells(Rownum, 11).Value
                          tName_increase = Cells(Rownum, 9).Value

                      End If

                      If ws.Cells(Rownum, 11).Value < Min_value Then
                          Min_percentvalue = ws.Cells(Rownum, 11).Value
                          tName_decrease = Cells(Rownum, 9).Value

                      End If

                      If ws.Cells(Rownum, 12).Value > Greatest_TotalValue Then
                          Greatest_TotalValue = ws.Cells(Rownum, 12).Value
                          tName_GreatestTotalVol = Cells(Rownum, 9).Value

                      End If

              Next Rownum

              ws.Cells(2, 18).Value = Max_percentvalue
              ws.Cells(2, 17).Value = tName_increase
              ws.Cells(3, 18).Value = Min_percentvalue
              ws.Cells(3, 17).Value = tName_decrease
              ws.Cells(4, 17).Value = tName_GreatestTotalVol
              ws.Cells(4, 18).Value = Greatest_TotalValue

      Next ws

      MsgBox ("End")

      End Sub

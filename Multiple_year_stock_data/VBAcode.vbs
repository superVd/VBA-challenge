
Sub StockMarket_Alphabet()


'**************************Declare Variables*******************************
Dim Ticker As String
Dim year_start As Double
Dim year_end As Double
Dim Yearly_Diff As Double
Dim Total_Stock As Double
Dim Percent_Change As Double
Dim start_data As Integer

'Define variable for worksheet
Dim ws As Worksheet


For Each ws In Worksheets

    'Assign column headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Diff"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock "

    'Assign intiger to loop
    data_start = 2
    previous_i = 1
    Total_Stock = 0


  'Label and insert last row of coumn A
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        'For loop finds the  yearly change, percent change, and total stock volume
        For i = 2 To LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Find Tickersymbol
            Ticker = ws.Cells(i, 1).Value

            'Initiate the variable to move the next alphabet
            previous_i = previous_i + 1


            ' Get the value for  column C and column F
            year_start = ws.Cells(previous_i, 3).Value
            year_end = ws.Cells(i, 6).Value

            'For loop to sum the total in column G
            For j = previous_i To i

                Total_Stock = Total_Stock + ws.Cells(j, 7).Value

            Next j

            'When the loop value is "0"
            If year_start = 0 Then

                Percent_Change = year_end

            Else
                Yearly_Diff = year_end - year_start

                Percent_Change = Yearly_Diff / year_start

            End If
            
            
         '********************worksheet summary******************

            ws.Cells(data_start, 9).Value = Ticker
            ws.Cells(data_start, 10).Value = Yearly_Diff
            ws.Cells(data_start, 11).Value = Percent_Change

            'Use for percentage format
            ws.Cells(data_start, 11).NumberFormat = "0.00%"
            ws.Cells(data_start, 12).Value = Total_Stock

            'first row task completed go to the next row
            data_start = data_start + 1

            'Get back the variable to zero
            Total_Stock = 0
            Yearly_Diff = 0
            Percent_Change = 0

            'Move i number
            previous_i = i

        End If

    'Done the loop

    Next i


  '**********************Second Summary Table*****************************

    'Calculate column k
    kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row

    'Define variable to initiate the second summery table value

    increase = 0
    decrease = 0
    greatest = 0

        'find max/min for percentage change and the max volume Loop
        For k = 3 To kEndRow

            Thelast_k = k - 1

            'Define percentage for current row
            current_k = ws.Cells(k, 11).Value

            'Define percentage for previous row
            previous_k = ws.Cells(Thelast_k, 11).Value

            'greatest total volume row
            volume = ws.Cells(k, 12).Value

            'Prevous greatest volume row
            previous_vol = ws.Cells(Thelast_k, 12).Value

   '************************Find the Increase*************************

            If increase > current_k And increase > previous_k Then

                increase = increase


            ElseIf current_k > increase And current_k > previous_k Then

                increase = current_k

                'define name for increase percentage
                increase_name = ws.Cells(k, 9).Value

            ElseIf previous_k > increase And previous_k > current_k Then

                increase = previous_k

                'define name for increase percentage
                increase_name = ws.Cells(Thelast_k, 9).Value

            End If

       '***************************Find the drecrease**********************

            If decrease < current_k And decrease < previous_k Then

                decrease = decrease

 
            ElseIf current_k < increase And current_k < previous_k Then

                decrease = current_k

                decrease_name = ws.Cells(k, 9).Value

            ElseIf previous_k < increase And previous_k < current_k Then

                decrease = previous_k

                decrease_name = ws.Cells(Thelast_k, 9).Value

            End If

       '***********************Find the greatest volume*******************

            If greatest > volume And greatest > previous_vol Then

                greatest = greatest

            ElseIf volume > greatest And volume > previous_vol Then

                greatest = volume

                'define name for greatest volume
                greatest_name = ws.Cells(k, 9).Value

            ElseIf previous_vol > greatest And previous_vol > volume Then

                greatest = previous_vol

                'define name for greatest volume
                greatest_name = ws.Cells(last_k, 9).Value

            End If

        Next k

  '*********************calculate the greatest increase, greatest decrease, and greatest volume*****************

    ws.Range("N1").Value = "Column Name"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker Name"
    ws.Range("P1").Value = "Value"

    '***********Get the greatest increase name, greatest increase name, and  greatest volume Ticker name
    ws.Range("O2").Value = increase_name
    ws.Range("O3").Value = decrease_name
    ws.Range("O4").Value = greatest_name
    ws.Range("P2").Value = increase
    ws.Range("P3").Value = decrease
    ws.Range("P4").Value = greatest

    'Greatest increase and decrease in percentage format

    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"


'--------------------------------------------------
' Conditional formatting columns colors

'The end row for column J

    jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row


        For j = 2 To jEndRow

            'if greater than or less than zero
            If ws.Cells(j, 10) > 0 Then

                ws.Cells(j, 10).Interior.ColorIndex = 4

            Else

                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If

        Next j

'Excute to next worksheet
Next ws
'--------------------------------------------------
End Sub

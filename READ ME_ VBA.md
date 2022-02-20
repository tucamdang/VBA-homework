'Part 1: I mainly focus on summarising the table and total ticker volume (same as credit charge to see how it works)
" I started to run test same as credit charge to see how it works.
'I manually test ticker AAB and AAF to
'I firstly dim (current worksheet and workbook), and loop all formula to go through all sheets.

	Dim Currentws As Worksheet
   	Dim wd As Workbook 
	For Each Currentws In Worksheets

'then dim all related information (ticker name and ticker total, open/ close price...) to recognise their identification.
	        
    	Dim Ticker_name As String
    	Ticker_name = " "
    	Dim Ticker_total As Double
    	Ticker_total = 0
    	Dim Open_price As Double
    	Open_price = 0
    	Dim Close_price As Double
    	Close_price = 0
    	Dim Yearly_change As Double
   	Yearly_change = 0
    	Dim Percent_change As Double
    	Percent_change = 0

	Dim Summary_table_row As Long
    	Summary_table_row = 2

'I started set up the column for ticker name and ticker total 
	For i = 2 To Lastrow

    If Currentws.Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker_name = Cells(i, 1).Value

		Ticker_total = Ticker_total + Currentws.Cells(i, 7).Value
        
        'print and check
        	Currentws.Range("I" & Summary_table_row).Value = Ticker_name
        	Currentws.Range("L" & Summary_table_row).Value = Ticker_total

' Add 1 to the summary table row
		Summary_table_row = Summary_table_row + 1

'Reset ticker total
		Ticker_total = 0

'Add to Ticker total, else if in next ticker name, enter new ticker stock volume
	Else
    	Ticker_total = Ticker_total + Currentws.Cells(i, 7).Value
   End If

'Next step (yearly change and percent change):

'Because I already dim type of yearly change and percent change, I don't need to repeat.
'After counting the certain type of ticker name, get next beggining price
Open_price = Currentws.Cells(i + 1, 3).Value

'Due to long list of data, I started add last row after Dim summary table row 
	Lastrow = Currentws.Cells(Rows.Count, 1).End(xlUp).Row
    	Open_price = Currentws.Cells(2, 3).Value

'Set up open/ close price in which cells, then edit formula to calculate the yearly change and percent change:

        Close_price = Currentws.Cells(i, 6).Value
        Yearly_change = Close_price - Open_price
        'Percent change
        If Open_price <> 0 Then
        Percent_change = (Yearly_change / Open_price) * 100
        End If

'Print the calculation in Yearly change and decide the colour (positive green, negative red)
Currentws.Range("J" & Summary_table_row).Value = Yearly_change
            'change colour 4 is green, 3 is red
            If (Yearly_change > 0) Then
            Currentws.Range("J" & Summary_table_row).Interior.ColorIndex = 4
            
            ElseIf (Yearly_change <= 0) Then
            Currentws.Range("J" & Summary_table_row).Interior.ColorIndex = 3
            End If

'Print percent change %
        Currentws.Range("K" & Summary_table_row).Value = (CStr(Percent_change) & "%")
        
'BONUS PART
'I set value for max and min data required
    Dim Max_ticker_name As String
    Max_ticker_name = " "
    Dim Min_ticker_name As String
    Min_ticker_name = " "
    Dim Max_percent As Double
    Max_percent = 0
    Dim Min_percent As Double
    Min_percent = 0
    Dim Max_volume_ticker_name As String
    Max_volume_ticker_name = " "
    Dim Max_volume As Double
    Max_volume = 0

'Do calculation 
If (Percent_change > Max_percent) Then
                Max_percent = Percent_change
                Max_ticker_name = Ticker_name
                
                ElseIf (Percent_change < Min_percent) Then
                Min_percent = Percent_change
                Min_ticker_name = Ticker_name
                
                End If
                
                If (Ticker_total > Max_volume) Then
                Max_volume = Ticker_total
                Max_volume_ticker_name = Ticker_name
                
                End If
       
        
        'reset value 
        Percent_change = 0
        Ticker_total = 0

'THEN print information into row or cells
Currentws.Range("I1").Value = "Ticker"
		Currentws.Range("J1").Value = "Yearly Change"
		Currentws.Range("K1").Value = "Percent Change"
		Currentws.Range("L1").Value = "Total Stock Volume"
		Currentws.Range("O1") = "Ticker name"
		Currentws.Range("P1") = "Value"

		Currentws.Range("N2").Value = "Greatest % Increase"
		Currentws.Range("N3").Value = "Greatest % Decrease"
		Currentws.Range("N4").Value = "Greatest Total Volume"
	
		Currentws.Range("P2").Value = (CStr(Max_percent) & "%")
		Currentws.Range("P3").Value = (CStr(Min_percent) & "%")
		Currentws.Range("O2").Value = Max_ticker_name
		Currentws.Range("O3").Value = Min_ticker_name
		Currentws.Range("P4").Value = Max_volume
		Currentws.Range("O4").Value = Max_volume_ticker_name
    	Next Currentws
	End 

'NOTE: I copied codes into three different sheets to achieve the correct value.
	

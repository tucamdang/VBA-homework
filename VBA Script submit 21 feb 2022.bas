Attribute VB_Name = "Module1"

Sub data_A()
    Dim Currentws As Worksheet
    Dim wd As Workbook
    
    Set wb = ActiveWorkbook

    'loop all sheets
For Each Currentws In Worksheets
        
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
    
    'bonus set value
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
       
    Dim Summary_table_row As Long
    Summary_table_row = 2
   
    Dim Lastrow As Long
    'Dim i As Long
    Lastrow = Currentws.Cells(Rows.Count, 1).End(xlUp).Row
    Open_price = Currentws.Cells(2, 3).Value

For i = 2 To Lastrow

    If Currentws.Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker_name = Cells(i, 1).Value
        
        'Yearly_change
        Close_price = Currentws.Cells(i, 6).Value
        Yearly_change = Close_price - Open_price
        'Percent change
        If Open_price <> 0 Then
        Percent_change = (Yearly_change / Open_price) * 100
        End If
        
        'ticker total
        Ticker_total = Ticker_total + Currentws.Cells(i, 7).Value
        
        'print and check
        Currentws.Range("I" & Summary_table_row).Value = Ticker_name
        Currentws.Range("J" & Summary_table_row).Value = Yearly_change
            
	'change colour
            If (Yearly_change > 0) Then
            Currentws.Range("J" & Summary_table_row).Interior.ColorIndex = 4
            
            ElseIf (Yearly_change <= 0) Then
            Currentws.Range("J" & Summary_table_row).Interior.ColorIndex = 3
            End If
        
        'Print
        Currentws.Range("K" & Summary_table_row).Value = (CStr(Percent_change) & "%")
        Currentws.Range("L" & Summary_table_row).Value = Ticker_total
        Summary_table_row = Summary_table_row + 1
    
    Open_price = Currentws.Cells(i + 1, 3).Value
    
                
       'bonus part
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
    Else
    Ticker_total = Ticker_total + Currentws.Cells(i, 7).Value
        End If
    Next i
    
        'Print
        Currentws.Range("N2").Value = "Greatest % Increase"
        Currentws.Range("N3").Value = "Greatest % Decrease"
        Currentws.Range("N4").Value = "Greatest Total Volume"
	
	Currentws.Range("P2").Value = (CStr(Max_percent) & "%")
        Currentws.Range("P3").Value = (CStr(Min_percent) & "%")
        Currentws.Range("O2").Value = Max_ticker_name
        Currentws.Range("O3").Value = Min_ticker_name
        Currentws.Range("P4").Value = Max_volume
        
    Next Currentws
End Sub











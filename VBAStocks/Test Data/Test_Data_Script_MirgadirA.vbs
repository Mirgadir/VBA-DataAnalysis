Sub stock():
	'assign variables
	Dim total As Double
	Dim rownum As Integer
	Dim oopen As Double
	Dim cloose As Double
	Dim ticker As String
	Dim ychange As Double
	Dim perchange As Double
	Dim maxPer As Double
	Dim minPer As Double
	Dim maxTotal As Double
	Dim tagMax As String
	Dim tagMin As String
	Dim tagTot As String

	'Main loop through worksheets
	For Each ws In Worksheets
		LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
		'To activate loop for all worksheets
		'ws.Activate
		'Start total from 0
		total = 0
		'First row for calculations
		rownum = 2

		'oopen as open year value
		oopen = ws.Cells(2, 3).Value
		'assign headings
		ws.Range("I1").Value = "Ticker"
		ws.Range("J1").Value = "Yearly Change"
		ws.Range("K1").Value = "Percent Change"
		ws.Range("L1").Value = "Total Stock Volume"
		'Second loop from 2 to last row
			For i = 2 To LastRow
				'check when stock changes
				If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
					'cloose as close year value
					cloose = ws.Cells(i, 6).Value
					'year change calculation and assign cell
					ychange = cloose - oopen
					ws.Range("J" & rownum).Value = ychange
					'Stock name called ticker
					ticker = ws.Cells(i, 1).Value
					'Total stock volume calculation and assign cell/format
					total = total + Cells(i, 7).Value
					ws.Range("L" & rownum).Value = total
					ws.Range("L" & rownum).NumberFormat = "0"
							'Year Change Cell color change as per value
							If ychange < 0 Then
							ws.Cells(rownum, 10).Interior.ColorIndex = 3
							Else
							ws.Cells(rownum, 10).Interior.ColorIndex = 4
							End If

							'Percentage change calculation in order to avoid /0 error
							If oopen <> 0 Then
							perchange = ychange / oopen
							Else
							perchange = 0
							End If
					ws.Range("K" & rownum).Value = perchange
					ws.Range("K" & rownum).NumberFormat = "0.00%"
					ws.Range("I" & rownum).Value = ticker
					'add row number for next loop
					rownum = rownum + 1
					'total = 0 for next loop
					total = 0
					'open year value from next stock
					oopen = ws.Cells(i + 1, 3).Value
				Else
					'if stock continious add total stock volume
					total = total + ws.Cells(i, 7).Value
				End If
			Next i
	Next ws

	'new loop for faster calculations
	For Each ws In Worksheets
		ws.Range("O1").Value = "Greatest % Increase"
		ws.Range("O2").Value = "Greatest % Decrease"
		ws.Range("O3").Value = "Greatest Total Volume"
	
		maxTotal = WorksheetFunction.Max(ws.Range("L:L"))
		maxPer = WorksheetFunction.Max(ws.Range("K:K"))
		minPer = WorksheetFunction.Min(ws.Range("K:K"))
	
		ws.Range("Q1").Value = maxPer
		ws.Range("Q2").Value = minPer
		ws.Range("Q3").Value = maxTotal
	
		ws.Range("P1").Value = tagMax
		ws.Range("P2").Value = tagMin
		ws.Range("P3").Value = tagTot
		
		ws.Range("Q1").NumberFormat = "0.00%"
		ws.Range("Q2").NumberFormat = "0.00%"
		ws.Range("Q3").NumberFormat = "0.0000E+00"
	
		For i = 2 To 3200
			If ws.Cells(i, 12).Value = maxTotal Then
			tagTot = ws.Cells(i, 12).Offset(0, -3).Value
			ElseIf ws.Cells(i,11).Value = maxPer Then
			tagMax = ws.Cells(i, 11).Offset(0, -2).Value
			ElseIf ws.Cells(i, 11).Value = minPer Then
			tagMin = ws.Cells(i, 11).Offset(0, -2).Value
			End If
		Next i
	Next ws
End Sub
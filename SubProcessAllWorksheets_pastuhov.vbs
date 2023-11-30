Sub ProcessAllWorksheets()
    Dim ws As Worksheet

    ' Loop through all worksheets
    For Each ws In Worksheets
        ' Call the main subroutine for each worksheet
        ProcessWorksheet ws
    Next ws
End Sub

Sub ProcessWorksheet(ws As Worksheet)
    ' Workout last row
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Make some variables
    Dim ticker As String
    Dim Summary_table_row As Integer
    Summary_table_row = 2
    Dim volume As LongLong
    volume = 0
    Dim openprice As Double
    openprice = 0
    Dim closeprice As Double
    closeprice = 0
    Dim yearlychange As Double
    yearlychange = 0
    Dim percentchange As Double
    percentchange = 0
    Dim GreatestIncrease As Double
    GreatestIncrease = 0
    Dim GreatestIncreaseTicker As String
    GreatestIncreaseTicker = " "
    Dim GreatestDecrease As Double
    GreatestDecrease = 0
    Dim GreatestDecreaseTicker As String
    GreatestDecreaseTicker = " "
    Dim GreatestTotalVolume As LongLong
    GreatestTotalVolume = 0
    Dim GreatestVolumeTicker As String
    GreatestVolumeTicker = " "

    ' Start loop
    ' Check for unique values in Ticker column A
    For i = 2 To lastRow
        If i = 2 Then
            openprice = ws.Cells(i, 3).Value
        End If

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value

            ' Put it in the summary table
            ws.Range("I" & Summary_table_row).Value = ticker

            closeprice = ws.Cells(i, 6).Value
            yearlychange = closeprice - openprice
            ws.Range("J" & Summary_table_row).Value = yearlychange
            percentchange = yearlychange / openprice
            ws.Range("K" & Summary_table_row).Value = percentchange

            ' Add up all the volumes
            volume = volume + ws.Cells(i, 7).Value
            ws.Range("L" & Summary_table_row).Value = volume

            If percentchange > GreatestIncrease Then
                GreatestIncrease = percentchange
                GreatestIncreaseTicker = ticker
            Else
                If percentchange < GreatestDecrease Then
                    GreatestDecrease = percentchange
                    GreatestDecreaseTicker = ticker
                End If
            End If

            If volume > GreatestTotalVolume Then
                GreatestTotalVolume = volume
                GreatestVolumeTicker = ticker
            End If

            ' Move to the next row in summary_table
            Summary_table_row = Summary_table_row + 1
            ' Reset volume for new ticker
            volume = 0
            ' Set new openprice
            openprice = ws.Cells(i + 1, 3).Value
        Else
            ' If the next ticker is the same, then add it up
            volume = volume + ws.Cells(i, 7).Value
        End If
    Next i

    ws.Range("I1").Value = "ticker"
    ws.Range("P1").Value = "ticker"
    ws.Range("J1").Value = "YearlyChange"
    ws.Range("K1").Value = "PercentChange"
    ws.Range("L1").Value = "TotalVolumeStock"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest%increase"
    ws.Range("O3").Value = "Greatest%decrease"
    ws.Range("O4").Value = "GreatestTotalVolume"
   
   
    ws.Range("Q2").Value = GreatestIncrease
    ws.Range("P2").Value = GreatestIncreaseTicker
    ws.Range("Q3").Value = GreatestDecrease
    ws.Range("P3").Value = GreatestDecreaseTicker
    ws.Range("Q4").Value = GreatestTotalVolume
    ws.Range("P4").Value = GreatestVolumeTicker
    ws.Range("Q2").Style = "Percent"
ws.Range("Q3").Style = "Percent"

    

    ' Color Column J and K based on Percent Change
    Dim lastRowsummarytable As Long
    lastRowsummarytable = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

    For i = 2 To lastRowsummarytable
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4 ' Green
        ElseIf ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3 ' Red
        End If
        
    
    Next i
    
     For i = 2 To lastRowsummarytable
     ws.Cells(i, 11).Style = "percent"
     Next i
     
       For i = 2 To lastRowsummarytable
     ws.Cells(i, 10).Style = "currency"
     Next i
    
  
End Sub



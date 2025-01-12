VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
 
Sub StockAnalysis()
    
    Dim ws As Worksheet
    Dim lastRow As Long
 
    
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
 
    
    Dim summaryRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim increaseTickerName As String
    Dim decreaseTickerName As String
    Dim volumeTickerName As String
 
    
    For Each ws In ThisWorkbook.Worksheets
        
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        summaryRow = 2
        totalVolume = 0
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
 
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
 
        
        ticker = ws.Cells(2, 1).Value
        openPrice = ws.Cells(2, 3).Value
 
        
        Dim i As Long
        For i = 2 To lastRow
            
            totalVolume = totalVolume + ws.Cells(i, 7).Value
 
            
            If ws.Cells(i + 1, 1).Value <> ticker Then
                
                closePrice = ws.Cells(i, 6).Value
 
                
                yearlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = yearlyChange / openPrice
                Else
                    percentChange = 0
                End If
 
                
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
 
                
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
 
                
                If yearlyChange >= 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
 
                
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    increaseTickerName = ticker
                End If
 
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    decreaseTickerName = ticker
                End If
 
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    volumeTickerName = ticker
                End If
 
                
                summaryRow = summaryRow + 1
                totalVolume = 0
                ticker = ws.Cells(i + 1, 1).Value
                openPrice = ws.Cells(i + 1, 3).Value
            End If
        Next i
 
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
 
        
        ws.Cells(2, 16).Value = increaseTickerName
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 16).Value = decreaseTickerName
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 16).Value = volumeTickerName
        ws.Cells(4, 17).Value = greatestVolume
 
        
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
 
        
        ws.Columns("I:Q").AutoFit
    Next ws
 
    MsgBox "Analysis Complete!", vbInformation
End Sub

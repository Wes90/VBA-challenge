VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockdata_WRA()
    'confirm variables
    Dim ticker As String
    Dim lastrow As Long
    Dim secondrow As Long
    Dim stockvolume As Double
    Dim openprice As Double
    Dim closedprice As Double
    Dim totalvolume As Long
    Dim yearlychange As Double
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        
        ws.Range("I" & 1) = "Ticker"
        ws.Range("J" & 1) = "Yearly change"
        ws.Range("K" & 1) = "Percent change"
        ws.Range("L" & 1) = "Total stock volume"
        
        
        secondrow = 2
        stockvolume = 0
        openprice = 0
        closedprice = 0
        yearlychange = 0
        
        j = 0
        k = 0
        
     
    
        lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
     
     'loopthrough tickers
     For i = 2 To lastrow
        
    'check tickers if they are not alike, then...
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            
'            If Total = 0 Then
'                ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
'                ws.Range("J" & 2 + j).Value = 0
'                ws.Range("K" & 2 + j).Value = "%" & 0
'                ws.Range("L" & 2 + j).Value = 0
'            Else
'                If ws.Cells(secondrow, 3).Value = 0 Then
'                    For find_value = secondrow To i
'                        If ws.Cells(find_value, 3).Value <> 0 Then
'                            secondrow = find_value
'                    Exit For
'                        End If
'                Next find_value
'                End If
                
               
                
'                secondrow = i + 1
                
                
            
            ticker = ws.Cells(i, 1).Value
            openprice = ws.Cells(i, 3).Value
            yearlychange = (ws.Cells(i, 6).Value - ws.Cells(secondrow, 3).Value)
            percentchange = Round((yearlychange / ws.Cells(secondrow, 3) * 100), 2)
'            closedprice = ws.Cells(i, 6).Value
'            stockvolume = stockvolume + ws.Cells(i, 7).Value
'            yearlychange = yearlychange + (closedprice - openprice)



            ws.Range("I" & secondrow).Value = ticker
            ws.Range("L" & secondrow).Value = stockvolume
            ws.Range("J" & secondrow).Value = yearlychange
            ws.Range("K" & secondrow).Value = "%" & percentchange


            secondrow = secondrow + 1
            stockvolume = 0
            yearlychange = 0
        Else
            stockvolume = stockvolume + ws.Cells(i, 7).Value
    'yearly change from opening price to closing at each year
            openprice = ws.Cells(i, 3).Value
            closedprice = ws.Cells(i, 6).Value
            yearlychange = yearlychange + (closedprice - openprice)








'
'            End If


        End If
      Next i
    Next ws
 
End Sub



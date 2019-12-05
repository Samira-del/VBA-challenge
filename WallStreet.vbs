Sub WallStreet():
    Dim Ticker As String
    Dim lastRow As double
    Dim k As integer
    Dim yearlychange as double
    Dim lastprice As double
    Dim firstprice As double
    Dim percentage as double
    Dim totalStock as double
    Dim GreatestIncrease As double
    Dim GreatestDecrease As double
    Dim GreatestValume As double

    For Each ws In Worksheets
        k = 1
        p = 1
        m = 1
        lastRow = cells(rows.count, 1).end(xlUp).row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage"
        ws.Cells(1, 12).Value = "Total Stock Valume"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Valume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"

        For i = 2 To lastRow
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                Ticker = ws.Cells(i, 1).Value
                k = k + 1
                ws.Cells(k, 9).Value = Ticker
            End If      
        Next i
        
        For j = 2 To lastRow
            totalStock = ws.cells(j, 7) + totalStock
            
            If (ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value) Then
                m = j - m + 1
                firstprice = ws.Cells(m, 3).value
                lastprice = ws.Cells(j, 6).Value
                p = p + 1
                ws.Cells(p, 10).NumberFormat = "0.00"
                yearlychange = lastprice - firstprice  
                ws.Cells(p, 10).Value = yearlychange
                if yearlychange >= 0 then
                        ws.Cells(p, 10).interior.colorindex = 4
                else
                    ws.Cells(p, 10).Value = yearlychange
                    ws.Cells(p, 10).interior.colorindex = 3
                end if 
                ws.Cells(p, 11).NumberFormat = "0.00"
                if firstprice <> 0 then
                    percentage = (yearlychange / firstprice) * 100
                    ws.Cells(p, 11).value = percentage 
                    m = 0
                else
                    ws.Cells(p, 11).value = 0 
                    m = 0
                end if
                ws.cells(p, 12).value = totalStock
                totalStock = 0
            End If
            
            m = m + 1
            yearlychange = 0  

        Next j

        GreatestIncrease = ws.cells(2, 11).value
        Ti = ws.Cells(2, 9).value
        GreatestDecrease = ws.cells(2, 11).value
        Td = ws.Cells(2, 9).value
        GreatestValume = ws.cells(2, 12).value
        Tv = ws.Cells(2, 9).value
        for index = 3 to lastRow:
            if GreatestIncrease < ws.Cells(index, 11).value Then
                GreatestIncrease = ws.cells(index, 11).value
                Ti = ws.Cells(index, 9)
            end If
            if GreatestDecrease > ws.Cells(index, 11).value Then
                GreatestDecrease = ws.cells(index, 11).value
                Td = ws.Cells(index, 9)
            end If
            if GreatestValume < ws.Cells(index, 12).value Then
                GreatestValume = ws.cells(index, 12).value
                Tv = ws.Cells(index, 9)
            end If
        next index

        ws.Cells(2, 15).value = Ti
        ws.Cells(2, 16).NumberFormat = "0.00"
        ws.Cells(2, 16).value = GreatestIncrease
        ws.Cells(3, 15).value = Td
        ws.Cells(3, 16).NumberFormat = "0.00"
        ws.Cells(3, 16).value = GreatestDecrease
        ws.Cells(4, 15).value = Tv
        ws.Cells(4, 16).NumberFormat = "0.00"
        ws.Cells(4, 16).value = GreatestValume
        
    next ws

end sub
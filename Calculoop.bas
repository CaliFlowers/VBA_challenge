Attribute VB_Name = "Module1"
Sub Calculoop()

    Dim Current As Worksheet
    
For Each Current In Worksheets
    Dim i As Long
    Dim summaryrow As Integer
    Dim ticker As String
    Dim netchange As Double
    Dim openprice As Double
    Dim closeprice As Double
    Dim percentchange As Double
    Dim lastrow As Long
    Dim gpi As Double
    Dim lpi As Double
    Dim greatestvalue As Variant
    Dim thisvolume As Variant
    Dim totalvolume As Variant

        
        Current.Cells(1, 8).Value = ""
        Current.Cells(1, 9).Value = "Ticker"
        Current.Cells(1, 10).Value = "Yearly_Net_Change"
        Current.Cells(1, 11).Value = "Percent_Annual_Change"
        Current.Cells(1, 12).Value = "Total Volume"
        Current.Cells(1, 15).Value = "Distinct Performers"
        Current.Cells(2, 15).Value = "Highest Volume"
        Current.Cells(3, 15).Value = "GPI"
        Current.Cells(3, 17).NumberFormat = "0.00%"
        Current.Cells(4, 15).Value = "LPI"
        Current.Cells(4, 17).NumberFormat = "0.00%"
        Current.Cells(1, 16).Value = "Ticker"
        Current.Cells(1, 17).Value = "Value"
       
        
        lastrow = Current.Cells(Rows.Count, 1).End(xlUp).Row
        totalvolume = 0
        summaryrow = 2
        openprice = Current.Cells(2, 3).Value
        
        
        
        
        For i = 2 To lastrow
              thisvolume = Current.Cells(i, 7).Value
              If Current.Cells(i + 1, 1).Value <> Current.Cells(i, 1).Value Then
            
              ticker = Current.Cells(i, 1).Value
             
              totalvolume = totalvolume + thisvolume
            
              closeprice = Current.Cells(i, 6).Value
              netchange = closeprice - openprice
              
              If netchange = 0 Then
              percentchange = 0
              ElseIf openprice = 0 Then
              percentchange = 0
              Else
              percentchange = (netchange / openprice) * 100
              End If
            
              
              summaryrow = summaryrow + 1
              totalvolume = 0
              openprice = Current.Cells(i + 1, 3).Value
              
            Else
                
                thisvolume = Current.Cells(i, 7).Value
                totalvolume = totalvolume + thisvolume
            
            End If
        Next i
        
        lastrow = Current.Cells(Rows.Count, 1).End(xlUp).Row
        lpi = 99.99
        gpi = -99.99
        gretestvolume = 0
        
        For i = 2 To summaryrow
              If Current.Cells(i, 11).Value > gpi Then
              gpi = Current.Cells(i, 11).Value
              Current.Cells(3, 17).Value = gpi
              Current.Cells(3, 16).Value = Current.Cells(i, 9).Value
              ElseIf Current.Cells(i, 11).Value < lpi Then
              lpi = Current.Cells(i, 11).Value
              Current.Cells(4, 17).Value = lpi
              Current.Cells(4, 16).Value = Current.Cells(i, 9).Value

              End If
              
              If Current.Cells(i, 12).Value > greatestvolume Then
              greatestvolume = Current.Cells(i, 12).Value
              Current.Range("Q2").Value = greatestvolume
              Current.Cells(2, 16).Value = Current.Cells(i, 9).Value
              
              End If
        Next i
              
        
             For i = 2 To lastrow
                Current.Cells(i, 11).NumberFormat = "0.00%"
                Current.Cells(i, 12).ColumnWidth = 25
                Current.Cells(i, 15).ColumnWidth = 20
                Current.Cells(i, 17).ColumnWidth = 25
                
            If Current.Cells(i, 11).Value > 0 Then
                Current.Cells(i, 11).Interior.ColorIndex = 4
            ElseIf Current.Cells(i, 11).Value < 0 Then Current.Cells(i, 11).Interior.ColorIndex = 3
            Else: Current.Cells(i, 11).Interior.ColorIndex = 0
            
            End If
        
        Next i
        
       
        
      Current.Range("I" & summaryrow).Value = ticker
              Current.Range("L" & summaryrow).Value = totalvolume
              Current.Range("J" & summaryrow).Value = netchange
              Current.Range("K" & summaryrow).Value = percentchange
        
Next Current
End Sub

Attribute VB_Name = "Module1"
Sub ticker():

    Dim innerloop As Long
    Dim outerloop As Long
    Dim lrow As Long
    Dim currentticker, lastticker, ultimateticker As String
    Dim runningtotal As Double
    
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    currentticker = Cells(2, 1).Value  'get the first ticker symbol
    lastticker = Cells(2, 1).Value 'set the last ticker to the same as first in the beginning
    ultimateticker = Cells(lrow - 1, 1).Value 'get the last ticker symbol
    innerloop = 2
        
    For outerloop = 2 To lrow
        runningtotal = 0
        
        While (currentticker = lastticker) And (lastticker <> "ZZZ")  ' run the loop for each ticker symbol
            runningtotal = runningtotal + Cells(innerloop, 7).Value ' 7->I
            'If IsEmpty(Cells(innerloop + 1, 1).Value) = False Then
            If currentticker <> ultimateticker Then
                lastticker = currentticker
                currentticker = Cells(innerloop + 1, 1).Value
            Else
                lastticker = "ZZZ"
            End If
            innerloop = innerloop + 1
        Wend
        
        If lastticker <> "ZZZ" Then ' if the ultimate ticker is not flagged
            Cells(outerloop, 9).Value = lastticker
            Cells(outerloop, 10).Value = runningtotal
            lastticker = currentticker
        Else
            Cells(outerloop, 9).Value = ultimateticker
            While Not IsEmpty((Cells(innerloop, 1).Value))  'run the loop one more time for the ultimate ticker
            runningtotal = runningtotal + Cells(innerloop, 7)
            innerloop = innerloop + 1
            Wend
            Cells(outerloop, 10).Value = runningtotal
            Exit For
        End If
                
    Next outerloop
    
    
    
End Sub

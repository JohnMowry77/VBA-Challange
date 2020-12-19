Attribute VB_Name = "Module1"
Sub headers()

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
 End Sub
 
Sub UniqueTicker_List()
    Dim LastRow As Double
    Dim i As Double
    Dim ticker As String
    Dim i_summary As Double
    Dim totalVolume As LongLong
    Dim start As Double
    Dim ticker_start As Double
    Dim change As Double
    Dim percentChange As Double
    
    i_summary = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    totalVolume = 0
    start = 2
    change = 0
    
    For i = 2 To LastRow:
        ticker = Cells(i, 1).Value
        
        If ticker <> Cells(i + 1, 1).Value Then
            If Cells(start, 3).Value = 0 Then
                For ticker_start = start To i
                    If Cells(ticker_start, 3).Value <> 0 Then
                        start = ticker_start
                        Exit For
                    End If
                Next ticker_start
            End If
            
            change = (Cells(i, 6).Value - Cells(start, 3).Value)
            percentChange = Round((change / Cells(start, 3) * 100), 2)
            
            Cells(i_summary, 9).Value = ticker
            Cells(i_summary, 10).Value = Round(change, 2)
            Cells(i_summary, 11).Value = "%" & percentChange
            Cells(i_summary, 12).Value = totalVolume + Cells(i, 7).Value
            
            start = i + 1
            totalVolume = 0
            i_summary = i_summary + 1
        Else
            totalVolume = totalVolume + Cells(i, 7).Value
        End If
        
     Next i

End Sub

Sub format()
Dim YearlyChg As Double
Dim LastUniqueRow As Double
Dim i As Long


    For i = 2 To Rows.Count
        If Not IsEmpty(Cells(i, 11).Value) Then
            YearlyChg = Cells(i, 11).Value
            
            If YearlyChg >= 0 Then
                Cells(i, 11).Interior.ColorIndex = 4
            Else
                Cells(i, 11).Interior.ColorIndex = 3
            End If
         End If
    Next i
 
End Sub

Sub max()
Dim Maximum As String
Dim Minimum As String
Dim Total As String
Dim MaxValue  As Long
Dim Count As Integer
Dim LastRow As Double


LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    
    Maximum = "Greatest % increase"
    Cells(2, 15) = ("Greatest % Increase")
     Cells(1, 16).Value = "Ticker"
     Cells(1, 17).Value = "Total"
     
    Minimum = "Greatest % decrease"
    Cells(3, 15) = ("Greatest % decrease")

     Total = "Greatest Total Volume"
        Cells(4, 15) = ("Greatest Total Volume")
        Cells(2, 16) = ("I Couldn't Get this")
        Cells(2, 17) = ("No Bonus Points?")
        Cells(3, 16) = ("No WS Worksheets Loop")
        Cells(3, 17) = ("Can you show me, please?")
        Cells(4, 16) = ("expression.Max(Ar1,Arg2,Arg3)")
        Cells(4, 17) = ("expression.Min(Ar1,Arg2,Arg3)")
        
End Sub

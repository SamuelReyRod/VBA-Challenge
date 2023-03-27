# VBA-Challenge
Week 2 Challenge



Sub MasterLooper()
Dim ws As Integer
ws = Application.Worksheets.Count
For i = 1 To ws
Worksheets(i).Activate
StocksFinal

Next

End Sub





Sub StocksFinal()


Range("I1,S1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("T1").Value = "Value"
Range("R2").Value = "Greatest % Increase"
Range("R3").Value = "Greatest % Decrease"
Range("R4").Value = "Greatest Total Volume"

Dim Summary_Table_Row As Integer
Dim lastrow As Long
Summary_Table_Row = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Brand_Total = 0

For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Brand_Name = Cells(i, 1).Value
    Brand_Total = Brand_Total + Cells(i, 7).Value
    closed_Bid = Cells(i, 6).Value
    Open_Bid = Application.WorksheetFunction.VLookup(Brand_Name, Range("A:C"), 3, 0)
    Prcnt = (closed_Bid - Open_Bid) / Open_Bid
    Percent_Change = FormatPercent(Prcnt, 2)
    Yearly_Change = (closed_Bid - Open_Bid)
    
    Range("J" & Summary_Table_Row).Value = Yearly_Change
           
        If Yearly_Change >= 0 Then
        Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
        Else
        Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
        End If
    
    Range("K" & Summary_Table_Row).Value = Percent_Change
    Range("I" & Summary_Table_Row).Value = Brand_Name
    Range("L" & Summary_Table_Row).Value = Brand_Total
    Summary_Table_Row = Summary_Table_Row + 1
    
    Brand_Total = 0
    
    Else
    Brand_Total = Brand_Total + Cells(i, 7).Value
    
    End If
   
Next i

MaxPC = Application.WorksheetFunction.Max(Range("k:k"))
MinPC = Application.WorksheetFunction.Min(Range("k:k"))
MaxPCT = FormatPercent(MaxPC, 2)
MinPCT = FormatPercent(MinPC, 2)

Range("T2").Value = MaxPCT
Range("T3").Value = MinPCT
Range("T4").Value = Application.WorksheetFunction.Max(Range("l:l"))
Range("S2").Value = Application.WorksheetFunction.XLookup(Range("T2"), Range("K:K"), Range("I:I"))
Range("S3").Value = Application.WorksheetFunction.XLookup(Range("T3"), Range("K:K"), Range("I:I"))
Range("S4").Value = Application.WorksheetFunction.XLookup(Range("T4"), Range("L:L"), Range("I:I"))


End Sub


![image](https://user-images.githubusercontent.com/125604132/228079514-f1d17668-28a1-4104-a6c3-c3cfb4026af4.png)
![image](https://user-images.githubusercontent.com/125604132/228079563-24a14a48-3857-4fd9-b1c2-20e7359d312a.png)
![image](https://user-images.githubusercontent.com/125604132/228079593-874565ef-bf02-47a9-adce-eee152e645df.png)




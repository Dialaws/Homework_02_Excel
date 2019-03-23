Attribute VB_Name = "Module1"
Sub Stock()

Dim totalvolume As LongLong

Dim rowcount As Long

Dim tickercount As Integer
Dim i As Long
Dim ticker As String
Dim summarytablerow As Integer



For Each ws In Worksheets
ws.Activate



totalvolume = 0

summarytablerow = 2
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Total Stock Volume"


rowcount = ws.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowcount

        If ws.Cells((i + 1), 1).Value <> ws.Cells(i, 1).Value Then
   
         ticker = ws.Cells(i, 1).Value
         
         ws.Range("I" & summarytablerow).Value = ticker
         ws.Range("J" & summarytablerow).Value = totalvolume
         
         summarytablerow = summarytablerow + 1
         
         totalvolume = 0
         Else
         
         totalvolume = totalvolume + ws.Cells(i, 7).Value
         
         End If
         Next i
         Next ws
         
         
         
         
End Sub


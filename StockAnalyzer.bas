Attribute VB_Name = "Module1"
Sub StockAnalyzer()

Dim row_number As Long
Dim column As Long
Dim ticker As String
Dim day As Date
Dim open_price As Long
Dim high_price As Long
Dim low_price As Long
Dim close_price As Long
Dim volume As Double
Dim end_of_data As Long
Dim summary_table_row As Long
Dim end_summary_data As Long


end_of_data = Cells(Rows.Count, 1).End(xlUp).Row
summary_table_row = 2
volume = 0
row_number = 2

'MsgBox (end_of_data)

For row_number = 2 To end_of_data

    If Cells(row_number + 1, 1).Value <> Cells(row_number, 1).Value Then
        ticker = Cells(row_number, 1).Value
        Range("I" & summary_table_row).Value = ticker
        summary_table_row = summary_table_row + 1
        
    End If

Next row_number

end_summary_data = Cells(Rows.Count, 9).End(xlUp).Row

summary_table_row = 2
volume = 0
row_number = 2

For summary_table_row = 2 To end_summary_data

    ticker = Range("I" & summary_table_row).Value
    
    For row_number = 2 To end_of_data
        If ticker = Range("A" & row_number).Value Then
            volume = volume + Cells(row_number, 7).Value
        End If
    Next row_number

    Range("L" & summary_table_row).Value = volume
    
    volume = 0

Next summary_table_row


'Dim DateRange()
'Dim first_day As Variant
'Dim last_day As Variant


'summary_table_row = 2
'row_number = 2

'For summary_table_row = 2 To end_summary_data

'    ticker = Range("I" & summary_table_row).Value
    
'    For row_number = 2 To end_of_data
    
'        If Cells(row_number, 1).Value = ticker And Cells(row_number + 1, 2) < Cells(row_number, 2) Then
'        first_day = Cells(row_number + 1, 2).Value
'        Else
'            first_day = Cells(row_number, 2).Value
              
'        End If
        
'        If Cells(row_number, 1).Value = ticker And Cells(row_number + 1, 2) > Cells(row_number, 2) Then
'        last_day = Cells(row_number + 1, 2).Value
        
'        End If
                
'    Next row_number
   
'    MsgBox (first_day)
'    MsgBox (last_day)
          
'   Range("M" & summary_table_row).Value = first_day
'   Range("N" & summary_table_row).Value = last_day

'Next summary_table_row


End Sub


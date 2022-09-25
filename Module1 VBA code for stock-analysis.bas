Attribute VB_Name = "Module1"
Sub alphabetical_testing():

'Declare variables and values
Dim ticker As String
Dim yearly_change As Double
    yearly_change = 0
Dim percentage_change As Integer
    percentage_change = 0
Dim total_stocks As Long
       

' Keep track of the location for each column in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2


'last row formula components
Dim lrow As Integer
lrow = Cells(Rows.Count, 1).End(xlUp).Row


'Loop through all sheets
'For Each ws In Worksheets not working?
worksheet_count = ActiveWorkbook.Worksheets.Count
For w = 1 To worksheet_count
Worksheets(w).Activate

'Headings
        Range("I1").Value = "ticker"
        Range("J1").Value = "yearly_change"
        Range("K1").Value = "percentage_change"
        Range("L1").Value = "total_stocks"

'Set opening value
opening_value = Cells(2, 3).Value

'Set Summary Table Row
Summary_Table_Row = 2

'forloop
For i = 2 To lrow
    '22771 (A)

'Check if ticker has changed. If it has...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            'Set Ticker Name
            ticker = Cells(i - 1, 1).Value
            'And print Ticker in Summary Column
            Range("I" & Summary_Table_Row).Value = ticker

            ' Calculate Yearly Change
            '******How to get to the first row of that ticker and pull from column 3 (opening)???
            yearly_change = Cells(i, 6).Value - opening_value
            
            'Print yearly_change in summary column
            Range("J" & Summary_Table_Row).Value = yearly_change

            '*********Calculate percentage change (formula for this?)
            percentage_change = (yearly_change / opening_value) * 100
            
            
            'percentage_change (Cells(i - 1, 6).Value / Cells(i - 1, 3).Value)
            'Print percentage_change in summary column
            Range("K" & Summary_Table_Row).Value = percentage_change
             
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1

            opening_value = Cells(i + 1, 3).Value
        
   
    Else
       
            'Add to the Yearly Total (Should this be calculated first?)
            yearly_total = yearly_total + Cells(i, 7).Value
    
    
            'Print yearly total in summary column
            Range("L" & Summary_Table_Row).Value = yearly_total
    End If

        'do autofit
        'Range("L1").EntireColumn.AutoFit
        'But this keeps breaking excel so nevermind
    
'continue loop
Next i

'Next ws
Next w

End Sub


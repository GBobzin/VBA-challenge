
Sub TickerOpenandclose()

 Dim current_sheet As Worksheet
 
 For Each current_sheet In Worksheets
 
Sheets(current_sheet.Name).Select


'find last row in sheet
Dim last_row As Long

last_row = current_sheet.Cells(current_sheet.Rows.Count, "A").End(xlUp).Row

'set all data on column B as date
Range("A1:a" & last_row).NumberFormat = "yyyy/mm/dd"

'Find Open on day 1 and Close on last day values for each Ticker

' Ticker Code
Dim Ticker As String

'Ticker open value
Dim Ticker_open As Double

'Ticker close value
Dim Ticker_close As Double

'Sum of total volume for a given ticker
 Dim Ticker_Total As Double
 
'Indetify first and last entry for a ticker


Dim ticker_count As Integer
Dim yearly_change As Double


'Variable to assist in creating summary table
 
 Dim summary_table_Row As Integer
     summary_table_Row = 2
     
     
'Identify Columns

Range("J1") = "Ticker"
Range("M1") = "Total Stock Volume"
Range("K1") = "Yearly Change"
Range("L1") = "Percent Change"

 ' Loop through daily data
  For i = 2 To last_row

    ' Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    
      ' Set the Ticker
      Ticker = Cells(i, 1).Value
       
      
       'Determine Ticker Close Value
      Ticker_close = Cells(i, 6).Value
      
     
      'Determine Ticker open value
      Ticker_open = Cells(i - ticker_count, 3)
   
      'Calculate Yearly change
      
      yearly_change = Ticker_close - Ticker_open
          
      
      'Calculate Percent change
               
               If yearly_change = 0 Or Ticker_open = 0 Then
               Percent_change = "NIL"
               Else
               
                      Percent_change = (yearly_change / Ticker_open)
     
               End If
 
   
      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value
    
    
 

      ' Print the Ticker Total in a Summary -                               CONSIDER REMOVING THE WORKSHEET REFERENCE - EACH YEAR SUMMARISED INDIVIDUALLY
      Range("J" & summary_table_Row).Value = Ticker

      ' Print the Ticker Amount to the Summary Table
     Range("M" & summary_table_Row).Value = Ticker_Total
      

      
      'print yearly change
       Range("K" & summary_table_Row).Value = yearly_change
       
       ' Conditional formatting for yearly change
            If Percent_change > 0 Then
                Range("K" & summary_table_Row).Interior.ColorIndex = 43
            Else
                Range("K" & summary_table_Row).Interior.ColorIndex = 3
            End If
      
      'print Percent change W conditional formatting
      Range("L" & summary_table_Row).Value = Percent_change
    Range("L" & summary_table_Row).NumberFormat = "0.00%"
        

      


      ' Add one to the summary table row
      summary_table_Row = summary_table_Row + 1
               
         
      ' Reset the Ticker Total
      Ticker_Total = 0
      ticker_count = 0
      
      'Reset the yearly change
      yearly_change = 0
      

    ' If the cell immediately following a row is the same brand...
    
    Else

      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value
      ticker_count = ticker_count + 1
    
    
    End If

  Next i
  
  
  Call increase_decrease_total
Next



End Sub


Sub increase_decrease_total()


' Variable definitions

' Variables to hold values

Dim GRT_Increase As Double
Dim GRT_Decrease As Double
Dim GRT_Total As Double

'Variables to find Ticker

Dim GRT_Increase_Ticker As String
Dim GRT_Decrease_Ticker As String
Dim GRT_Total_Ticker As String

'Variables to set ranges for Lookup
Dim GRT_Increase_RNG As Range
Dim GRT_Decrease_RNG As Range
Dim GRT_Total_RNG As Range
Dim Ticker_RNG As Range

Set GRT_Total_RNG = ActiveSheet.Range("M:M")
Set GRT_Increase_RNG = ActiveSheet.Range("L:L")
Set Ticker_RNG = ActiveSheet.Range("J:J")


'Find Greatest Values

GRT_Total = WorksheetFunction.Max(GRT_Total_RNG)

GRT_Increase = WorksheetFunction.Max(GRT_Increase_RNG)

GRT_Decrease = WorksheetFunction.Min(GRT_Increase_RNG)

'Find Ticker for each value

GRT_Increase_Ticker = WorksheetFunction.Index(Ticker_RNG, WorksheetFunction.Match(GRT_Increase, GRT_Increase_RNG, 0))

Range("S2").Value = GRT_Increase_Ticker


GRT_Decrease_Ticker = WorksheetFunction.Index(Ticker_RNG, WorksheetFunction.Match(GRT_Decrease, GRT_Increase_RNG, 0))

Range("S3").Value = GRT_Decrease_Ticker


GRT_Total_Ticker = WorksheetFunction.Index(Ticker_RNG, WorksheetFunction.Match(GRT_Total, GRT_Total_RNG, 0))

Range("S4").Value = GRT_Total_Ticker

'print Headers & Format

Range("R2").Value = "Greatest % Increase"
Range("R3").Value = "Greatest % Decrease"
Range("R4").Value = "Greatest Total Volume"

Range("S1").Value = "Ticker"
Range("t1").Value = "Value"

Range("t2:t3").NumberFormat = "0.00%"

Columns("a:t").AutoFit

'Print Values

Range("t2").Value = GRT_Increase
Range("t3").Value = GRT_Decrease

Range("t4").Value = GRT_Total


End Sub

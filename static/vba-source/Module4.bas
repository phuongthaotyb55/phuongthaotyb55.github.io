Attribute VB_Name = "Module4"
Option Explicit

Sub Update_Final_Fcst()
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Dim startTime As Double
    Dim endTime As Double
    Dim runningTime As Double
    
    ' Record the start time
    startTime = Timer
    
    Dim i As Long, j As Long
    Dim npd_ws As Worksheet, splt_ws As Worksheet, fcst_ws As Worksheet, erp_ws As Worksheet
    With ThisWorkbook
        Set npd_ws = .Worksheets("1. NPD Info")
        Set splt_ws = .Worksheets("3. Split by SKUs")
        Set fcst_ws = .Worksheets("4. Final Fcst")
        Set erp_ws = .Worksheets("5. Demand Planning Export")
    End With
    
    With npd_ws
        'Define NPD's rows count exclude 2 first row
        Dim npd_row As Long: npd_row = .Range("C6").End(xlDown).Row - 5
    End With
    
    With fcst_ws
        .Range("A:A").ColumnWidth = 30
    'Create Check box table
        ' Set the range to count
        Dim Customer_Rng As Range: Set Customer_Rng = .Range(.Cells(npd_row + 10, 2), .Cells(npd_row + 24, 2))
        
        ' Count the number of values in the range
        Dim Customer_Cnt As Integer
        Customer_Cnt = Application.WorksheetFunction.CountA(Customer_Rng)
        
        
        'Delete any existing checkboxes in the range
        .CheckBoxes.Delete
        
        'Clear existing data
        With .Range("F" & npd_row + 7 & ":BM" & npd_row * 2 + 40 + Customer_Cnt)
            .ClearFormats 'Clear IRC checking to Fcst per Cus VAl range
            .ClearContents
        End With
    
        'Clear existing Database & consolidated Tables & filtered by Cus
        With .Range("A" & npd_row + 24 + Customer_Cnt & ":BM1000")
            .ClearFormats
            .ClearContents
        End With
        
        ' Copy Customers' Names
        Dim c As Integer, r As Integer
        If Customer_Cnt > 0 Then
            For c = 9 To 8 + Customer_Cnt
                .Cells(npd_row + 9, c).Formula = "=IF(" & .Cells(npd_row + c + 1, 2).Address & "<>""""," & _
                    .Cells(npd_row + c + 1, 2).Address & ","""")"
                .Cells(npd_row + 9, c + 28).Formula = "=IF(" & .Cells(npd_row + c + 1, 2).Address & "<>""""," & _
                    .Cells(npd_row + c + 1, 2).Address & ","""")"
            Next c
        Else:
            'Detect if any input of Customer's Name
            MsgBox "Error: No data of Customer's Name found", vbCritical, "Error"
            Exit Sub
        End If
        
        'Detect if the total weight of customers equal to 100%
        If Application.WorksheetFunction.Sum(.Range(.Cells(npd_row + 10, 4), .Cells(npd_row + 24, 4))) <> 1 Then
            MsgBox "Error: The Total Weight(%) for all Customers does not equal to 100%. Please adjust it", vbCritical, "Error"
            Exit Sub
        End If
        
        'Copy IRC, Shade/ML, Article Type for both Check box and List Price Tbls
        Dim arrCols() As Variant
        arrCols = Array(3, 5, 7)
        
        For i = 0 To UBound(arrCols)
            .Range(.Cells(5, arrCols(i)), .Cells(npd_row + 4, arrCols(i))).Copy Destination:=.Range("AH" & (npd_row + 9)).Offset(0, i)
            .Range(.Cells(5, arrCols(i)), .Cells(npd_row + 4, arrCols(i))).Copy Destination:=.Range("F" & (npd_row + 9)).Offset(0, i)
        Next i
        
        
    'Create fcst per customer overview (volume) table
        'Copy Customer's Input table
        For c = 2 To 4
            For r = npd_row + 9 To npd_row + 9 + Customer_Cnt
                .Cells(r + npd_row + 3, c + 4).Formula = "=" & .Cells(r, c).Address 'Checkbox
                .Cells(r + npd_row + 3, c + 32) = "=" & .Cells(r, c).Address 'List Price
            Next r
        Next c
        
        .Cells(npd_row * 2 + 13 + Customer_Cnt, 8).Value = "Total Vol"
        
        'Create date for the table based on Customer's ICW
        Dim month_horizon As Integer, month_max As Integer, month_cnvrt As Integer
        month_max = Month(.Cells(5, 9).Value)
        For r = npd_row * 2 + 13 To npd_row * 2 + 12 + Customer_Cnt
            month_cnvrt = Month(.Cells(r, 7).Value)
            If month_cnvrt > month_max Then
                month_max = month_cnvrt
                month_horizon = Application.WorksheetFunction.Max(month_max) - Month(.Cells(5, 9).Value)
            End If
        Next r
        
        .Cells(npd_row * 2 + 12, 9).Value = "=" & npd_ws.Cells(5, 9).Address(External:=True)
        For c = 10 To 20 + month_horizon
            .Cells(npd_row * 2 + 12, c).Formula = "=EDATE(" & .Cells(npd_row * 2 + 12, c - 1).Address & ", 1)" 'Using Edate function by increasing 1 month from Month 1
        Next c
        
        .Cells(npd_row * 2 + 12, 21 + month_horizon).Value = "Total Vol 12M"
        .Range(.Cells(npd_row * 2 + 12, 9), .Cells(npd_row * 2 + 12, 20 + month_horizon)).NumberFormat = "mmm-yy"
        
        'Calculate sum each column
        For c = 9 To 21 + month_horizon
            .Cells(npd_row * 2 + 13 + Customer_Cnt, c).Formula = "=SUM(" & .Cells(npd_row * 2 + 13, c).Address & _
                                    ":" & .Cells(npd_row * 2 + 12 + Customer_Cnt, c).Address & ")"
        Next c
    
        'Calculate fcst per customer for both Volume and Value table
        For r = npd_row * 2 + 13 To npd_row * 2 + 12 + Customer_Cnt
            'Sum each row
            .Cells(r, 21 + month_horizon).Formula = "=SUM(" & .Cells(r, 9).Address & ":" & _
                                        .Cells(r, 20 + month_horizon).Address & ")" 'Vol
    
            .Cells(r, 49 + month_horizon).Formula = "=SUM(" & .Cells(r, 37).Address & ":" & _
                                        .Cells(r, 48 + month_horizon).Address & ")" 'Val
                     
            For c = 9 To 20
                Dim month_var As Integer
                month_var = Abs(Month(.Cells(r, 7).Value) - Month(.Cells(npd_row * 2 + 12, 9).Value))
                'Vol
                '.Cells(r, c + month_var).Formula = "=SUMIFS(" & .Cells(6, c).Address & ":" & .Cells(npd_row + 4, c).Address & "," & _
                    .Cells(npd_row + 10, r - (npd_row * 2 + 13) + 9).Address & ":" & _
                                        .Cells(npd_row * 2 + 8, r - (npd_row * 2 + 13) + 9).Address & ", TRUE) *" & _
                    .Cells(r, 8).Address
                .Cells(r, c + month_var).Formula = "=SUM(" & .Cells(6, c).Address & ":" & .Cells(npd_row + 4, c).Address & "," & _
                    .Cells(npd_row + 10, r - (npd_row * 2 + 13) + 9).Address & ":" & _
                                        .Cells(npd_row * 2 + 8, r - (npd_row * 2 + 13) + 9).Address & ") *" & _
                    .Cells(r, 8).Address
            Next c
            For c = 37 To 48
                'Val
                '.Cells(r, c + month_var).Formula = "=SUMPRODUCT(" & .Cells(6, c - 28).Address & ":" & .Cells(npd_row + 4, c - 28).Address & _
                    ",--(" & _
                    .Cells(npd_row + 10, r - (npd_row * 2 + 13) + 9).Address & ":" & _
                                        .Cells(npd_row * 2 + 8, r - (npd_row * 2 + 13) + 9).Address & " = TRUE)," & _
                    .Cells(npd_row + 10, r - (npd_row * 2 + 13) + 37).Address & ":" & .Cells(npd_row * 2 + 8, r - (npd_row * 2 + 13) + 37).Address & _
                    ") *" & .Cells(r, 8).Address
                .Cells(r, c + month_var).Formula = "=SUMPRODUCT(" & .Cells(6, c - 28).Address & ":" & .Cells(npd_row + 4, c - 28).Address & _
                    "," & _
                    .Cells(npd_row + 10, r - (npd_row * 2 + 13) + 37).Address & ":" & .Cells(npd_row * 2 + 8, r - (npd_row * 2 + 13) + 37).Address & _
                    ") *" & .Cells(r, 8).Address
            Next c
        Next r
        
        
          'Create Name of Checkbox table and List Price Table, and all other Tables
        .Cells(npd_row + 8, 6).Value = "IRCs checking for each Customer"
        .Cells(npd_row + 8, 34).Value = "List Price checking for each Customer"
        .Cells(npd_row * 2 + 11, 6).Value = "Topline Forecast by Customer (in Units)"
        .Cells(npd_row * 2 + 11, 34).Value = "Topline Forecast by Customer (in Gross Sales)"
        .Cells(npd_row * 2 + 19 + Customer_Cnt, 2).Value = "(2) Forecast by IRC including Customer Weight/ICW (in Units)"
        .Cells(npd_row * 2 + 19 + Customer_Cnt, 30).Value = "Forecast by IRC including Customer Weight/ICW (in Gross Sales)"
        
        Dim name_rng() As Range
        ReDim name_rng(1 To 6)
        Set name_rng(1) = .Range(.Cells(npd_row + 8, 6), .Cells(npd_row + 8, 8 + Customer_Cnt))
        Set name_rng(2) = .Range(.Cells(npd_row + 8, 34), .Cells(npd_row + 8, 36 + Customer_Cnt))
        Set name_rng(3) = .Range(.Cells(npd_row * 2 + 11, 6), .Cells(npd_row * 2 + 11, 21 + month_horizon))
        Set name_rng(4) = .Range(.Cells(npd_row * 2 + 11, 34), .Cells(npd_row * 2 + 11, 49 + month_horizon))
        Set name_rng(5) = .Range(.Cells(npd_row * 2 + 19 + Customer_Cnt, 2), .Cells(npd_row * 2 + 19 + Customer_Cnt, 21 + month_horizon))
        Set name_rng(6) = .Range(.Cells(npd_row * 2 + 19 + Customer_Cnt, 30), .Cells(npd_row * 2 + 19 + Customer_Cnt, 49 + month_horizon))
        For i = 1 To 6
            With name_rng(i)
                .Merge
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .Font.Color = vbWhite
                .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
            End With
        Next i
        
        
    'Create List Price Table
        .Range(.Cells(6, 36), .Cells(npd_row + 4, 36)).Copy Destination:=.Range(.Cells(npd_row + 10, 37), _
                                .Cells(npd_row + 10, 36 + Customer_Cnt))
        
        'Format
        Dim format_rng() As Range
        ReDim format_rng(1 To 25)
        Set format_rng(1) = .Range(.Cells(npd_row + 9, 9), .Cells(npd_row + 9, 8 + Customer_Cnt)) 'Customer's Name of IRC checkbox table
        Set format_rng(2) = .Range(.Cells(npd_row + 9, 37), .Cells(npd_row + 9, 36 + Customer_Cnt)) 'Customer's Names of List Price Table
        Set format_rng(3) = .Range(.Cells(npd_row * 2 + 12, 9), .Cells(npd_row * 2 + 12, 20 + month_horizon)) '12M date of Fcst per Cus VOL Table
        Set format_rng(4) = .Range(.Cells(npd_row * 2 + 12, 37), .Cells(npd_row * 2 + 12, 48 + month_horizon)) '12M date of Fcst per Cus VAL Table
        Set format_rng(5) = .Range(.Cells(npd_row + 9, 6), .Cells(npd_row * 2 + 8, 8)) 'NPD Info of IRC checkbox table incl Headers
        Set format_rng(6) = .Range(.Cells(npd_row + 9, 34), .Cells(npd_row * 2 + 8, 36)) 'NPD Info of List Price table incl Headers
        Set format_rng(7) = .Range(.Cells(npd_row * 2 + 13, 6), .Cells(npd_row * 2 + 12 + Customer_Cnt, 8)) 'Excl Header of Customer Info of Fcst per Cus VOL Tbl
        Set format_rng(8) = .Range(.Cells(npd_row * 2 + 12, 34), .Cells(npd_row * 2 + 12 + Customer_Cnt, 36)) 'Excl Header of Customer Info of Fcst per Cus VAL Tbl
        Set format_rng(9) = .Range(.Cells(npd_row + 9, 9), .Cells(npd_row * 2 + 8, 8 + Customer_Cnt))
        Set format_rng(10) = .Range(.Cells(npd_row * 6 + 48, 1), _
                .Cells(npd_row * 6 + 49 + (Customer_Cnt - 1) * (npd_row - 1) + (npd_row - 2), 21 + month_horizon)) 'Database VOl
        Set format_rng(11) = .Range(.Cells(npd_row * 6 + 48, 29), _
                .Cells(npd_row * 6 + 49 + (Customer_Cnt - 1) * (npd_row - 1) + (npd_row - 2), 49 + month_horizon)) 'Database VAL
        Set format_rng(12) = .Range(.Cells(npd_row * 2 + 12, 6), .Cells(npd_row * 2 + 12, 8)) 'Header of Customer Info of Fcst per Cus VOL Tbl
        Set format_rng(13) = .Range(.Cells(npd_row * 2 + 12, 34), .Cells(npd_row * 2 + 12, 36)) 'Header of Customer Info of Fcst per Cus VAL Tbl
        Set format_rng(14) = .Range(.Cells(npd_row * 2 + 13 + Customer_Cnt, 36), _
                        .Cells(npd_row * 2 + 13 + Customer_Cnt, 49 + month_horizon)) 'Col SUm of Fcst per Cus VAL tbl
        Set format_rng(15) = .Range(.Cells(npd_row * 2 + 13 + Customer_Cnt, 8), _
                        .Cells(npd_row * 2 + 13 + Customer_Cnt, 21 + month_horizon)) 'Col SUm of Fcst per Cus VOL tbl
        Set format_rng(16) = .Range(.Cells(npd_row * 2 + 12, 49 + month_horizon), _
                        .Cells(npd_row * 2 + 12 + Customer_Cnt, 49 + month_horizon)) 'Row SUm of Fcst per Cus VAL tbl
        Set format_rng(17) = .Range(.Cells(npd_row * 2 + 12, 21 + month_horizon), _
                        .Cells(npd_row * 2 + 12 + Customer_Cnt, 21 + month_horizon)) 'Row SUm of Fcst per Cus VOL tbl
        Set format_rng(18) = .Range(.Cells(npd_row * 6 + 48, 21 + month_horizon), _
                    .Cells(npd_row * 6 + 49 + (Customer_Cnt - 1) * (npd_row - 1) + (npd_row - 2), 21 + month_horizon)) 'Total VOl 12M col of database table
        Set format_rng(19) = .Range(.Cells(npd_row * 6 + 48, 49 + month_horizon), _
                    .Cells(npd_row * 6 + 49 + (Customer_Cnt - 1) * (npd_row - 1) + (npd_row - 2), 49 + month_horizon)) 'Total VAl 12M col of database table
        Set format_rng(20) = .Range(.Cells(npd_row * 6 + 49, 9), _
                    .Cells(npd_row * 6 + 49 + (Customer_Cnt - 1) * (npd_row - 1) + (npd_row - 2), 21 + month_horizon)) 'All numerics range of VOl database table
        Set format_rng(21) = .Range(.Cells(npd_row * 6 + 49, 37), _
                    .Cells(npd_row * 6 + 49 + (Customer_Cnt - 1) * (npd_row - 1) + (npd_row - 2), 49 + month_horizon)) 'All numeric range of Val database table
        Set format_rng(22) = .Range(.Cells(npd_row * 2 + 13, 9), _
                        .Cells(npd_row * 2 + 13 + Customer_Cnt + 1, 22 + month_horizon)) 'All numeric values in Fcst per Customer VOL Tbl
        Set format_rng(23) = .Range(.Cells(npd_row * 2 + 13, 37), _
                        .Cells(npd_row * 2 + 13 + Customer_Cnt + 1, 49 + month_horizon)) 'All numeric values in Fcst per Customer VAL Tbl
        Set format_rng(24) = .Range(.Cells(npd_row + 10, 9), .Cells(npd_row * 2 + 8, 8 + Customer_Cnt)) 'checkboxes of IRC checkbox table excl Header
        Set format_rng(25) = .Range(.Cells(npd_row + 10, 37), .Cells(npd_row * 2 + 8, 36 + Customer_Cnt)) 'List Price of List Price table excl Header
        
        For i = 1 To 4
            With format_rng(i)
                .NumberFormat = "mmm-yy"
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .Font.Color = vbYellow
                .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
            End With
        Next i
        For i = 5 To 11
            format_rng(i).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        Next i
        For i = 12 To 19
            With format_rng(i)
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .Font.Color = vbWhite
                .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
            End With
        Next i
        For i = 20 To 23
            format_rng(i).NumberFormat = "#,##0"
        Next i
        For i = 24 To 25
            With format_rng(i)
                .Interior.Color = RGB(255, 242, 204) 'Gold, accent 4 , lighter 80%
            End With
        Next i
        
        .Cells.HorizontalAlignment = xlCenter


    'Create fcst per customer overview (value) table
        'Copy-paste 12-month date
        .Range(.Cells(npd_row * 2 + 12, 9), .Cells(npd_row * 2 + 12, 20 + month_horizon)).Copy Destination:=.Range("AK" & _
                                                                                                (npd_row * 2 + 12))
        
        .Cells(npd_row * 2 + 13 + Customer_Cnt, 36).Value = "Total Val"
        .Cells(npd_row * 2 + 12, 49 + month_horizon).Value = "Total Val 12M"
        
        'Calculate sum each column
        For c = 37 To 49 + month_horizon
            .Cells(npd_row * 2 + 13 + Customer_Cnt, c).Formula = "=SUM(" & .Cells(npd_row * 2 + 13, c).Address & _
                                    ":" & .Cells(npd_row * 2 + 12 + Customer_Cnt, c).Address & ")"
        Next c
    
        'Create database for Pivot Table for Fcst per Customer (VOL & VAL)
        Dim t As Long, k As Long
        
        Dim copyRange As Range, NPDVolRange As Range, NPDValRange As Range, VOLmonthRange As Range, VALmonthRange As Range
        
        Set VOLmonthRange = .Range(.Cells(npd_row * 2 + 12, 9), .Cells(npd_row * 2 + 12, 21 + month_horizon)) 'Month horizon
        Set VALmonthRange = .Range(.Cells(npd_row * 2 + 12, 37), .Cells(npd_row * 2 + 12, 49 + month_horizon)) 'Month horizon
        Set NPDVolRange = .Range(.Cells(6, 2), .Cells(npd_row + 4, 8))
        Set NPDValRange = .Range(.Cells(6, 30), .Cells(npd_row + 4, 36))
        
    
        'Copy header of NPD Vol for NPD Vol Database
        .Range(.Cells(5, 2), .Cells(5, 8)).Copy Destination:=.Range("B" & npd_row * 6 + 48) 'Vol
        .Range(.Cells(5, 30), .Cells(5, 36)).Copy Destination:=.Range("AD" & npd_row * 6 + 48) ''Val
        VOLmonthRange.Copy Destination:=.Range("I" & (npd_row * 6 + 48)) '12M of Vol
        VALmonthRange.Copy Destination:=.Range("AK" & (npd_row * 6 + 48)) '12M of Val
        .Cells(npd_row * 6 + 48, 1).Value = "Customer"
        .Cells(npd_row * 6 + 48, 29).Value = "Customer"
  
    
        For t = 0 To Customer_Cnt - 1
            Dim dest As Long
            dest = npd_row * 6 + 49 + t * (npd_row - 1)
            
            NPDVolRange.Copy Destination:=.Range("B" & dest) 'Vol
            NPDValRange.Copy Destination:=.Range("AD" & dest) 'Val
           
            month_var = Abs(Month(.Cells(npd_row * 2 + 13 + t, 7).Value) - Month(.Cells(npd_row * 2 + 12, 9).Value))
    
            For r = dest To dest + (npd_row - 2)
                
                .Cells(r, 1).Formula = "=IF(" & .Cells(npd_row + 10 + t, 2).Address & "<>""""," & _
                      .Cells(npd_row + 10 + t, 2).Address & ","""")" 'Create Customer Name for Vol
                .Cells(r, 29).Formula = "=IF(" & .Cells(npd_row + 10 + t, 2).Address & "<>""""," & _
                      .Cells(npd_row + 10 + t, 2).Address & ","""")" 'Create Customer Name for Val
                
                .Cells(r, 21 + month_horizon) = "=SUM(" & .Cells(r, 9).Address & ":" & _
                                            .Cells(r, 20 + month_horizon).Address & ")" 'Sum each row for Vol
                .Cells(r, 49 + month_horizon) = "=SUM(" & .Cells(r, 37).Address & ":" & _
                                            .Cells(r, 48 + month_horizon).Address & ")" 'Sum each row for Val
                                            
                
                'Calculate Weight of each IRC of each Customer
               .Cells(r, 8).Formula = "=IF(AND(" & .Cells(r - (npd_row * 5 + 39 + t * (npd_row - 1)), t + 9).Address & _
                          "=TRUE," & .Cells(r - (npd_row * 5 + 43 + t * (npd_row - 1) + npd_row), 8).Address & "<>0)," & _
                          .Cells(r - (npd_row * 5 + 43 + t * (npd_row - 1) + npd_row), 8).Address & "*" & _
                          .Cells(npd_row * 2 + 13 + t, 8).Address & "+ SUMPRODUCT(--(" & _
                          .Range(.Cells(npd_row + 10, t + 9), .Cells(npd_row * 2 + 8, t + 9)).Address & _
                          "=FALSE)," & .Range(.Cells(6, 8), .Cells(4 + npd_row, 8)).Address & ")*" & _
                          .Cells(npd_row * 2 + 13 + t, 8).Address & "/COUNTIF(" & _
                          .Range(.Cells(npd_row + 10, t + 9), .Cells(npd_row * 2 + 8, t + 9)).Address & _
                          ",TRUE),0)"
    
                
                'Calculate fcst per customer (VOl)
                For c = 9 To 20
                    .Cells(r, c + month_var).Formula = "=" & .Cells(r, 8).Address & "*" & _
                            .Cells(2, c).Address
                Next c
                
                'Copy List Price
                .Cells(r, 36).Formula = "=" & .Cells(r - (npd_row * 5 + 39 + t * (npd_row - 1)), t + 37).Address
               
                
                'Calculate fcst per customer (VAl)
                For c = 37 To 48
                    .Cells(r, c + month_var).Formula = "=" & .Cells(r, c + month_var - 28).Address & _
                        "*" & .Cells(r, 36).Address
                Next c
            Next r
        Next t
 
    'Create table for consolidated forecast & forecast filtered by customer
        Dim NPDVolRange_hd As Range, NPDValRange_hd As Range
        
        Set NPDVolRange_hd = .Range(.Cells(5, 2), .Cells(npd_row + 4, 8))
        Set NPDValRange_hd = .Range(.Cells(5, 30), .Cells(npd_row + 4, 36))
        
        NPDVolRange_hd.Copy Destination:=.Range("B" & (npd_row * 2 + 20 + Customer_Cnt)) 'Consolidated VOL
        NPDValRange_hd.Copy Destination:=.Range("AD" & (npd_row * 2 + 20 + Customer_Cnt)) 'Consolidated VAL
        NPDVolRange_hd.Copy Destination:=.Range("B" & (npd_row * 3 + 25 + Customer_Cnt)) 'Forecast filtered by Cus VOL
        NPDValRange_hd.Copy Destination:=.Range("AD" & (npd_row * 3 + 25 + Customer_Cnt)) 'Forecast filtered by Cus VAL
        VOLmonthRange.Copy Destination:=.Range("I" & (npd_row * 2 + 20 + Customer_Cnt)) 'Month horizon for Consolidated VOL
        VALmonthRange.Copy Destination:=.Range("AK" & (npd_row * 2 + 20 + Customer_Cnt)) 'Month horizon for Consolidated VAL
        VOLmonthRange.Copy Destination:=.Range("I" & (npd_row * 3 + 25 + Customer_Cnt)) 'Month horizon for Forecast filtered by Cus VOL
        VALmonthRange.Copy Destination:=.Range("AK" & (npd_row * 3 + 25 + Customer_Cnt)) 'Month horizon for Forecast filtered by Cus VAL
        
        Dim formulaStr As String
        
        For r = npd_row * 2 + 21 + Customer_Cnt To npd_row * 3 + 19 + Customer_Cnt
            'VOL Consolidated Table
            For c = 9 To 21 + month_horizon
                formulaStr = "="
                For i = 0 To Customer_Cnt - 1
                    formulaStr = formulaStr & IIf(i = 0, "", "+") & _
                                    .Cells(r + npd_row * 4 + 28 - Customer_Cnt + i * (npd_row - 1), c).Address
                Next i
                .Cells(r, c).Formula = formulaStr
            Next c
            
            'VAL Consolidated Table
            For c = 37 To 49 + month_horizon
                formulaStr = "="
                For i = 0 To Customer_Cnt - 1
                    formulaStr = formulaStr & IIf(i = 0, "", "+") & _
                                    .Cells(r + npd_row * 4 + 28 - Customer_Cnt + i * (npd_row - 1), c).Address
                Next i
                .Cells(r, c).Formula = formulaStr
            Next c
        Next r
        
        'Sum up columns for filter VOL Cus
        .Cells(npd_row * 3 + 20 + Customer_Cnt, 8).Value = "Total Vol"
        For c = 9 To 21 + month_horizon
            .Cells(npd_row * 3 + 20 + Customer_Cnt, c).Formula = "=SUM(" & .Cells(npd_row * 2 + 21 + Customer_Cnt, c).Address & ":" & _
                                            .Cells(npd_row * 3 + 19 + Customer_Cnt, c).Address & ")"
        Next c
        
        'Sum up columns for filter VAL Cus
        .Cells(npd_row * 3 + 20 + Customer_Cnt, 36).Value = "Total Val"
        For c = 37 To 49 + month_horizon
            .Cells(npd_row * 3 + 20 + Customer_Cnt, c).Formula = "=SUM(" & .Cells(npd_row * 2 + 21 + Customer_Cnt, c).Address & ":" & _
                                            .Cells(npd_row * 3 + 19 + Customer_Cnt, c).Address & ")"
        Next c
        
        'Create drop-down menu for Customer filtered
        Dim targetCellVOL As Range, targetCellVAL As Range
        Dim listSource As Range
        
        ' Define the target cell
        Set targetCellVOL = .Cells(npd_row * 3 + 23 + Customer_Cnt, 4)
        Set targetCellVAL = .Cells(npd_row * 3 + 23 + Customer_Cnt, 32)
        
        ' Define the source of the list
        Set listSource = .Range(.Cells(npd_row + 10, 2), .Cells(npd_row + 9 + Customer_Cnt, 2))
        
        ' Add data validation
        With targetCellVOL.Validation
            .Delete ' Remove any previous data validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="=" & listSource.Address
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .errorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
        
        With targetCellVAL.Validation
            .Delete ' Remove any previous data validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="=" & listSource.Address
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .errorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
        
        targetCellVAL.Formula = "=IF(" & targetCellVOL.Address & "<>""""," & _
                      targetCellVOL.Address & ","""")"
    
               
        ' Filter Values based on Customer and IRC VOL table
        Dim VOLlookup_range As Range
        Dim VOLcriteria_range1 As Range
        Dim VOLcriteria1 As Range
        Dim VOLcriteria_range2 As Range
        Dim VOLcriteria2 As Range
        
        For r = npd_row * 3 + 26 + Customer_Cnt To npd_row * 4 + 24 + Customer_Cnt
            For c = 9 To 21 + month_horizon
                ' Set ranges based on current column
                Set VOLlookup_range = .Range(.Cells(npd_row * 6 + 49, c), .Cells(npd_row * 6 + 49 + (Customer_Cnt - 1) * (npd_row - 1) + (npd_row - 2), c))
                Set VOLcriteria_range1 = .Range(.Cells(npd_row * 6 + 49, 1), .Cells(npd_row * 6 + 49 + (Customer_Cnt - 1) * (npd_row - 1) + (npd_row - 2), 1))
                Set VOLcriteria1 = .Cells(npd_row * 3 + 23 + Customer_Cnt, 4)
                Set VOLcriteria_range2 = .Range(.Cells(npd_row * 6 + 49, 3), .Cells(npd_row * 6 + 49 + (Customer_Cnt - 1) * (npd_row - 1) + (npd_row - 2), 3))
                Set VOLcriteria2 = .Cells(r, 3)
                        
                
                ' Build INDEX MATCH formula
                .Cells(r, c).FormulaArray = "=IFERROR(INDEX(" & VOLlookup_range.Address & _
                    ",MATCH(1,(" & VOLcriteria_range1.Address & "=" & VOLcriteria1.Address & _
                    ")*(" & VOLcriteria_range2.Address & "=" & VOLcriteria2.Address & "),0)), """")"

            Next c
        Next r
        
        'Sum up columns for filter VOL Cus
        .Cells(npd_row * 4 + 25 + Customer_Cnt, 8).Value = "Total Vol"
        For c = 9 To 21 + month_horizon
            .Cells(npd_row * 4 + 25 + Customer_Cnt, c).Formula = "=SUM(" & .Cells(npd_row * 3 + 26 + Customer_Cnt, c).Address & ":" & _
                                            .Cells(npd_row * 4 + 24 + Customer_Cnt, c).Address & ")"
        Next c
        
        
        ' Filter Values based on Customer and IRC VOL table
        Dim VALlookup_range As Range
        Dim VALcriteria_range1 As Range
        Dim VALcriteria1 As Range
        Dim VALcriteria_range2 As Range
        Dim VALcriteria2 As Range
        
        For r = npd_row * 3 + 26 + Customer_Cnt To npd_row * 4 + 24 + Customer_Cnt
            For c = 37 To 49 + month_horizon
                ' Set ranges based on current column
                Set VALlookup_range = .Range(.Cells(npd_row * 6 + 49, c), .Cells(npd_row * 6 + 49 + (Customer_Cnt - 1) * (npd_row - 1) + (npd_row - 2), c))
                Set VALcriteria_range1 = .Range(.Cells(npd_row * 6 + 49, 29), .Cells(npd_row * 6 + 49 + (Customer_Cnt - 1) * (npd_row - 1) + (npd_row - 2), 29))
                Set VALcriteria1 = .Cells(npd_row * 3 + 23 + Customer_Cnt, 32)
                Set VALcriteria_range2 = .Range(.Cells(npd_row * 6 + 49, 31), .Cells(npd_row * 6 + 49 + (Customer_Cnt - 1) * (npd_row - 1) + (npd_row - 2), 31))
                Set VALcriteria2 = .Cells(r, 3)
                        
                
                ' Build INDEX MATCH formula
                .Cells(r, c).FormulaArray = "=IFERROR(INDEX(" & VALlookup_range.Address & _
                    ",MATCH(1,(" & VALcriteria_range1.Address & "=" & VALcriteria1.Address & _
                    ")*(" & VALcriteria_range2.Address & "=" & VALcriteria2.Address & "),0)), """")"
            Next c
        Next r
        
        'Sum up columns for filter VAL Cus
        .Cells(npd_row * 4 + 25 + Customer_Cnt, 36).Value = "Total Val"
        For c = 37 To 49 + month_horizon
            .Cells(npd_row * 4 + 25 + Customer_Cnt, c).Formula = "=SUM(" & .Cells(npd_row * 3 + 26 + Customer_Cnt, c).Address & ":" & _
                                            .Cells(npd_row * 4 + 24 + Customer_Cnt, c).Address & ")"
        Next c
    
    'Format Consolidated and Filtered by Customer Tables
    With targetCellVOL
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204) 'Gold, accent 4 , lighter 80%
    End With
        
    With targetCellVAL
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204) 'Gold, accent 4 , lighter 80%
    End With
    
    'Stripe format for filter by Cus Tbl
    For r = npd_row * 3 + 26 + Customer_Cnt To npd_row * 4 + 24 + Customer_Cnt Step 2
        .Range(.Cells(r, 2), .Cells(r, 20 + month_horizon)).Interior.Color = RGB(142, 170, 219) 'Blue, Accent 1, Lighter 60%
        .Range(.Cells(r, 30), .Cells(r, 48 + month_horizon)).Interior.Color = RGB(142, 170, 219) 'Blue, Accent 1, Lighter 60%
    Next r
    
    'Stripe format for Consolidate
    For r = npd_row * 2 + 21 + Customer_Cnt To npd_row * 3 + 19 + Customer_Cnt Step 2
        .Range(.Cells(r, 2), .Cells(r, 20 + month_horizon)).Interior.Color = RGB(142, 170, 219) 'Blue, Accent 1, Lighter 60%
        .Range(.Cells(r, 30), .Cells(r, 48 + month_horizon)).Interior.Color = RGB(142, 170, 219) 'Blue, Accent 1, Lighter 60%
    Next r
    
    
    Dim format_tbl() As Range
        ReDim format_tbl(1 To 25)
        Set format_tbl(1) = .Range(.Cells(npd_row * 3 + 26 + Customer_Cnt, 9), .Cells(npd_row * 4 + 24 + Customer_Cnt, 20 + month_horizon)) 'All numeric of filtered by Cus VOL table
        Set format_tbl(2) = .Range(.Cells(npd_row * 3 + 26 + Customer_Cnt, 37), .Cells(npd_row * 4 + 24 + Customer_Cnt, 49 + month_horizon)) 'All numeric of filtered by Cus VAL table
        Set format_tbl(3) = .Range(.Cells(npd_row * 2 + 21 + Customer_Cnt, 9), .Cells(npd_row * 3 + 19 + Customer_Cnt, 20 + month_horizon)) 'All numeric of consolidated VOL table
        Set format_tbl(4) = .Range(.Cells(npd_row * 2 + 21 + Customer_Cnt, 37), .Cells(npd_row * 3 + 19 + Customer_Cnt, 49 + month_horizon)) 'All numeric of consolidated VAL table
        Set format_tbl(5) = .Range(.Cells(npd_row * 3 + 26 + Customer_Cnt, 21 + month_horizon), .Cells(npd_row * 4 + 24 + Customer_Cnt, 21 + month_horizon)) 'Total VOL 12M of filtered by Cus VOL table
        Set format_tbl(6) = .Range(.Cells(npd_row * 3 + 26 + Customer_Cnt, 49 + month_horizon), .Cells(npd_row * 4 + 24 + Customer_Cnt, 49 + month_horizon)) 'Total VAL 12M of filtered by Cus VAL table
        Set format_tbl(7) = .Range(.Cells(npd_row * 2 + 21 + Customer_Cnt, 21 + month_horizon), .Cells(npd_row * 3 + 19 + Customer_Cnt, 21 + month_horizon)) 'Total VOL 12M of consolidated VOL table
        Set format_tbl(8) = .Range(.Cells(npd_row * 2 + 21 + Customer_Cnt, 49 + month_horizon), .Cells(npd_row * 3 + 19 + Customer_Cnt, 49 + month_horizon)) 'Total VAL 12M of consolidated Cus VAL table
        Set format_tbl(9) = .Range(.Cells(npd_row * 4 + 25 + Customer_Cnt, 8), .Cells(npd_row * 4 + 25 + Customer_Cnt, 21 + month_horizon)) 'Total VOL row of filtered by Cus VOL table
        Set format_tbl(10) = .Range(.Cells(npd_row * 4 + 25 + Customer_Cnt, 36), .Cells(npd_row * 4 + 25 + Customer_Cnt, 49 + month_horizon)) 'Total VAL row of filtered by Cus VAL table
        Set format_tbl(11) = .Range(.Cells(npd_row * 3 + 20 + Customer_Cnt, 8), .Cells(npd_row * 3 + 20 + Customer_Cnt, 21 + month_horizon)) 'Total VOL row of consolidated VOL table
        Set format_tbl(12) = .Range(.Cells(npd_row * 3 + 20 + Customer_Cnt, 36), .Cells(npd_row * 3 + 20 + Customer_Cnt, 49 + month_horizon)) 'Total VAL row of consolidated Cus VAL table
        Set format_tbl(13) = .Cells(npd_row * 3 + 23 + Customer_Cnt, 2)
        Set format_tbl(14) = .Cells(npd_row * 3 + 23 + Customer_Cnt, 32)
        Set format_tbl(15) = .Cells(npd_row * 3 + 23 + Customer_Cnt, 3)
        Set format_tbl(16) = .Cells(npd_row * 3 + 23 + Customer_Cnt, 33)
        
        Dim row_base As Long
        row_base = npd_row * 3 + 23 + Customer_Cnt
        
        Dim text() As Variant
        text = Array("Forecast by IRC by Customer (in Units)", "Forecast by IRC by Customer (in Gross Sales)", "------------>", "------------>")
        
        Dim columns() As Variant
        columns = Array(2, 30, 3, 31)
        
        For i = 1 To 4
            Set format_tbl(i + 12) = .Cells(row_base, columns(i - 1))
            format_tbl(i + 12).Value = text(i - 1)
            With format_tbl(i)
                .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
                .NumberFormat = "#,##0"
            End With
        Next i
        
        For i = 5 To 12
            With format_tbl(i)
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .Font.Color = vbWhite
                .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
                .NumberFormat = "#,##0"
            End With
        Next i

        For i = 13 To 16
            With format_tbl(i)
                .Font.Color = vbRed
                .Interior.Color = RGB(255, 217, 102) 'Gold, Accent 4, Lighter 60%
                .VerticalAlignment = xlCenter
                .Font.Bold = True
            End With
        Next i
            
    End With
    
    'Update DemandPlanner Converted Sheet
    With erp_ws
        'To tackle order of rows in Pivot Table
        For r = 7 To 5 + npd_row
            .Cells(r, 4).Formula = "=IF(" & .Cells(2, 3).Address & "=""(1) Final Forecast by IRC with common ICW (in Units)"", " & _
                     fcst_ws.Cells(r - 1, 3).Address(External:=True) & _
                    ", IF(" & .Cells(2, 3).Address & "=""(2) Forecast by IRC including Customer Weight/ICW (in Units)"", " & _
                    fcst_ws.Cells(npd_row * 2 + 14 + r + Customer_Cnt, 3).Address(External:=True) & ", ""Pick table""))"
   
            For c = 8 To 19
                .Cells(r, c).Formula = "=IF(" & .Cells(2, 3).Address & "=""(1) Final Forecast by IRC with common ICW (in Units)"", " & _
                        "ROUND(" & fcst_ws.Cells(r - 1, c + 1).Address(External:=True) & ",0)" & _
                        ", IF(" & .Cells(2, 3).Address & "=""(2) Forecast by IRC including Customer Weight/ICW (in Units)"", " & _
                        "ROUND(" & fcst_ws.Cells(npd_row * 2 + 14 + r + Customer_Cnt, c + 1).Address(External:=True) & _
                        ",0),""""))"
            Next c
        Next r
        
        'Font Name, Font Size And General Setting
        .Cells.Font.Size = 8
        .Cells.Font.Name = "Verdana"
        .columns.AutoFit
        .Rows.AutoFit
        
        .Range("C:C").ColumnWidth = 55
        .Range("E:G").ColumnWidth = 12
        
    End With
    
    With fcst_ws
        .columns.AutoFit
        .Rows.AutoFit
    
        'Loop through the range and add a checkbox to each cell (Bring checkbox after autofit to prevent format destroyed)
        For i = npd_row + 10 To npd_row * 2 + 8
            For j = 9 To 8 + Customer_Cnt
                With .CheckBoxes.Add(.Cells(i, j).left, .Cells(i, j).Top + (.Cells(i, j).Height - 30) / 2, 6, 6)
                    If .Parent.Cells(i - (npd_row + 4), 8).Value <> 0 Then
                        .Value = xlOn
                    Else
                        .Value = xlOff
                    End If
                    .Display3DShading = True
                    .LinkedCell = .Parent.Cells(i, j).Address(True, True)
                    .Name = ""
                    .Caption = ""
                End With
            Next j
        Next i
        
        With .Range("B1:C1")
            .Value = Array("Macro last run at:", Format(Now(), "mm/dd/yyyy hh:mm:ss"))
            .Font.Color = vbRed
            .Font.Italic = True
        End With
        
        'Font Name, Font Size And General Setting
        .Cells.Font.Size = 8
        .Cells.Font.Name = "Verdana"
        .columns.AutoFit
        .Rows.AutoFit
        
        ' Record the end time
        endTime = Timer
        ' Calculate the running time in seconds
        runningTime = endTime - startTime
        ' Display the running time in cell B2
        With .Range("B2:C2")
            .Value = Array("Macro/VBA running Time:", Format(runningTime, "0.00") & " seconds")
            .Font.Color = vbRed
            .Font.Italic = True
        End With
    End With
    
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    
    MsgBox "Please check all of the data again." & vbNewLine & _
        "In case of modification in IRCs checking Table or List Price Table , please click Refresh button", vbInformation
    
End Sub



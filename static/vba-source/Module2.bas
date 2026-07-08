Attribute VB_Name = "Module2"
Option Explicit
Sub Generate_FinalFcst()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Dim db_ws As Worksheet, tl_ws As Worksheet, npd_ws As Worksheet, spl_ws As Worksheet, fcst_ws As Worksheet, erp_ws As Worksheet
    With ThisWorkbook
        Set db_ws = .Worksheets("0. Database")
        Set npd_ws = .Worksheets("1. NPD Info")
        Set tl_ws = .Worksheets("2. Topline Fcst")
        Set spl_ws = .Worksheets("3. Split by SKUs")
        Set fcst_ws = .Worksheets("4. Final Fcst")
        Set erp_ws = .Worksheets("5. Demand Planning Export")
    End With
    
    Dim r As Integer, c As Integer
    
    With npd_ws
        Dim npd_row As Long: npd_row = .Range("C6").End(xlDown).Row - 5 ' Subtract 5 to exclude the header row
    End With
    
    With fcst_ws
        'Clear Final Forecast WorkSheet
        .Cells.Clear
        'Delete any existing checkboxes in the range
        .CheckBoxes.Delete
   
        .Cells(2, 8) = "Topline Forecast"
        
        'Copy-paste NPD's Info except Concatenate Column into Final Forecast Worksheet
        For c = 2 To 7
            For r = 5 To (npd_row + 4)
                .Cells(r, c).Formula = "=" & npd_ws.Cells(r + 1, c).Address(External:=True)
            Next r
        Next c
    'Create final fcst per SKU (Volume) Table
        .Cells(5, 9).Value = "=" & npd_ws.Range("I5").Address(External:=True) 'Copy-paste the Month 1 date from NPD's input
        .Range(.Cells(5, 9), .Cells(5, 20)).NumberFormat = "mmm-yy" 'Set the date format for 12 months
        For c = 10 To 20
            .Cells(5, c).Formula = "=EDATE(" & .Cells(5, c - 1).Address & ", 1)" 'Using Edate function by increasing 1 month from Month 1
        Next c
        
        'Creat IRC's Weight Column
        .Cells(5, 8).Value = "IRC's Weight (%)"
        For r = 6 To npd_row + 4
            .Cells(r, 21).Formula = "=SUM(" & .Cells(r, 9).Address & ":" & .Cells(r, 20).Address & ")" 'Sum each row of Volume Tbl
        Next r
        
        .Cells(5, 21).Value = "Total Vol 12M"
        .Cells(npd_row + 5, 8).Value = "Total Vol"
        
        'Calculate each IRC demand in each month
        For r = 6 To npd_row + 4
            For c = 9 To 20
                .Cells(r, c).Formula = "=IFERROR(" & .Cells(2, c).Address & "*" & .Cells(r, 8).Address & ","""")"
            Next c
        Next r
        
        'Sum for each column
        For c = 9 To 21
            .Cells(npd_row + 5, c).Formula = "=IFERROR(SUM(" & .Cells(6, c).Address & ":" & .Cells(npd_row + 4, c).Address & "),"""")" 'Set the SUM formula for each column
            .Cells(1, c).Formula = "=" & .Cells(5, c).Address 'Date for Topline Forecast
        Next c
        
    'Create final fcst per SKU (Value) Table
        'Copy NPD's INfo table
        For c = 30 To 36
            For r = 5 To (npd_row + 4)
                .Cells(r, c).Formula = "=" & npd_ws.Cells(r + 1, c - 28).Address(External:=True)
            Next r
        Next c
        
        .Cells(5, 49).Value = "Total Val 12M"
        .Cells(npd_row + 5, 36).Value = "Total Val"
        For c = 37 To 48
            'Copy-paste 12 months
            .Cells(5, c).Formula = "=" & .Cells(5, c - 28).Address
            'SUm each column
            .Cells(npd_row + 5, c).Formula = "=SUM(" & .Cells(6, c).Address & ":" & .Cells(npd_row + 4, c).Address & ")"
        Next c
        
        'Sum each row
         For r = 6 To npd_row + 5
            .Cells(r, 49).Formula = "=IFERROR(SUM(" & .Cells(r, 37).Address & ":" & .Cells(r, 48).Address & "),"""")" 'Sum each row of Value Tbl
        Next r
        
        'Calculate each IRC demand in value in each month
        For r = 6 To npd_row + 4
            For c = 37 To 48
                .Cells(r, c).Formula = "=IFERROR(" & .Cells(r, c - 28).Address & "*" & .Cells(r, 36).Address & ","""")"
            Next c
        Next r
        
        'Create customer input table
        .Cells(npd_row + 9, 2).Value = "Customer"
        With .Range(.Cells(npd_row + 9, 3), .Cells(npd_row + 24, 4))
            .Cells(1, 1).Value = "Customer's ICW"
            .NumberFormat = "mmm-yy"
        End With
        
        Dim dataRange As Range
        Set dataRange = .Range(.Cells(5, 9), .Cells(5, 20))
        
        Dim dropdownRange As Range
        Set dropdownRange = .Range(.Cells(npd_row + 10, 3), .Cells(npd_row + 24, 3))
        
        With dropdownRange.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & dataRange.Address
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = "Invalid Entry"
            .InputMessage = ""
            .errorMessage = "Please select a value from the list."
            .ShowInput = True
            .ShowError = True
        End With
                
        With .Range(.Cells(npd_row + 9, 4), .Cells(npd_row + 24, 4))
            .Cells(1, 1).Value = "Weight (%)"
            .NumberFormat = "0%"
        End With
        
         'Copy overwrite extrapolated Weight
        For r = 6 To npd_row + 4
            .Cells(r, 8).Formula = "=" & spl_ws.Cells(r + 6, 31).Address(External:=True)
        Next r
    
        'Copy Final Topline Fcst into Final Fcst Sheet
        With tl_ws
            Dim newStartRow As Long, begin_cell As Long
            newStartRow = .Range("B11").CurrentRegion.Rows.count + 14 'First row of Option 2 Table
            begin_cell = .Range("B" & newStartRow).CurrentRegion.Rows.count * 2 + 17 'First row of Assumption Table
            For c = 7 To 19
                fcst_ws.Cells(2, c + 2).Formula = "=" & .Cells(begin_cell + 25, c).Address(External:=True)
            Next c
        End With
        
        'Create Update button
        Dim btn As Button
        Dim btn_position As Range
        Dim btnToDelete As Button
        
        ' Define the button position
        Set btn_position = .Range("A" & (npd_row + 10))
        
        
        ' Delete existing Update Customer Info button (if any)
        On Error Resume Next
        .Shapes("Update").Delete
        On Error GoTo 0
        
        ' Delete existing Refresh all Pivot Table button (if any)
        On Error Resume Next
        .Shapes("Refresh").Delete
        On Error GoTo 0
        
        ' Add the new button
        Set btn = .Buttons.Add(btn_position.left, btn_position.Top, 45, 45)
        
        ' Customize the button properties
        With btn
            .Font.Bold = True
            .Font.Size = 7
            .Font.Name = "Verdana"
        End With
        
        ' Customize the button properties
        With btn
            .Name = "Update"
            .Caption = "Update Customer Info"
            .OnAction = "Update_Final_Fcst" ' Name of the macro to run when the button is clicked
        End With
        
        
        .Cells(4, 2).Value = "(1) Final Forecast by IRC with common ICW (in Units)"
        .Cells(4, 30).Value = "(2) Forecast by IRC including Customer Weight/ICW (in Gross Sales)"
        .Cells(npd_row + 8, 2).Value = "Customer Info"
        
    'Format Fcst Split by SKUs
        'Stripe format
        For r = 6 To npd_row + 4 Step 2
            .Range(.Cells(r, 2), .Cells(r, 20)).Interior.Color = RGB(142, 170, 219) 'Blue, Accent 1, Lighter 60%
            .Range(.Cells(r, 30), .Cells(r, 48)).Interior.Color = RGB(142, 170, 219) 'Blue, Accent 1, Lighter 60%
        Next r
        
        Dim i As Long
        Dim format_rng() As Range
        ReDim format_rng(1 To 20)
        Set format_rng(1) = .Range(.Cells(2, 8), .Cells(2, 21)) 'Topline value of 12M row
        Set format_rng(2) = .Range(.Cells(1, 9), .Cells(1, 21)) 'Topline date of 12M row
        Set format_rng(3) = .Range(.Cells(5, 9), .Cells(5, 21)) '12M row of fcst split by SKUs VOL Table
        Set format_rng(4) = .Range(.Cells(5, 37), .Cells(5, 48)) '12M row of fcst split by SKUs VAL Table
        Set format_rng(5) = .Range(.Cells(npd_row + 9, 2), .Cells(npd_row + 9, 4)) 'Header of Customer Input Tbl
        Set format_rng(6) = .Range(.Cells(5, 2), .Cells(5, 8)) 'NPD Info row fcst split by SKUs VOl Table
        Set format_rng(7) = .Range(.Cells(5, 30), .Cells(5, 36)) 'NPD Info row fcst split by SKUs VAL Table
        Set format_rng(8) = .Range(.Cells(5, 21), .Cells(npd_row + 4, 21)) 'SUm of row fcst split by SKUs VOL Table
        Set format_rng(9) = .Range(.Cells(5, 49), .Cells(npd_row + 4, 49)) 'SUm of row fcst split by SKUs VAL Table
        Set format_rng(10) = .Range(.Cells(npd_row + 5, 8), .Cells(npd_row + 5, 21)) 'SUm of col fcst split by SKUs VOL Table
        Set format_rng(11) = .Range(.Cells(npd_row + 5, 36), .Cells(npd_row + 5, 49)) 'SUm of col fcst split by SKUs VAL Table
        Set format_rng(12) = .Range(.Cells(npd_row + 8, 2), .Cells(npd_row + 8, 4)) 'Title of Customer Input Tbl
        Set format_rng(13) = .Range(.Cells(npd_row + 10, 2), .Cells(npd_row + 24, 4)) 'Customer Input Value of 15 row
        Set format_rng(14) = .Range(.Cells(5, 2), .Cells(npd_row + 4, 8)) 'NPD Info of Split by SKUs VOL table
        Set format_rng(15) = .Range(.Cells(5, 30), .Cells(npd_row + 4, 36)) 'NPD Info of Split by SKUs VAL table
        Set format_rng(16) = .Range(.Cells(6, 9), .Cells(npd_row + 5, 21)) 'All numerics value range of Split by SKU VOL Table
        Set format_rng(17) = .Range(.Cells(6, 37), .Cells(npd_row + 5, 49)) 'All numerics value range of Split by SKU VAL Table
        Set format_rng(18) = .Range(.Cells(6, 8), .Cells(npd_row + 4, 8)) 'IRC's Weight column
        Set format_rng(19) = .Range(.Cells(4, 2), .Cells(4, 21)) 'Tittle of first Table in units
        Set format_rng(20) = .Range(.Cells(4, 30), .Cells(4, 49)) 'Tittle of first Table in value
        
        With format_rng(1)
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .NumberFormat = "#,##0"
            .Interior.Color = RGB(142, 170, 219) 'Blue, Accent 1, Lighter 60%
            .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        End With
        For i = 2 To 5
            With format_rng(i)
                .NumberFormat = "mmm-yy"
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .Font.Color = vbYellow
                .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
            End With
        Next i
        
        For i = 6 To 12
            With format_rng(i)
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .Font.Color = vbWhite
                .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
            End With
        Next i
        
        With format_rng(12)
            .Merge
            .HorizontalAlignment = xlCenter
        End With
        
        With format_rng(13)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(255, 242, 204) 'Gold, accent 4 , lighter 80%
            .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        End With
        
        For i = 14 To 15
            format_rng(i).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        Next i
        For i = 16 To 17
            format_rng(i).NumberFormat = "#,##0"
        Next i
        
        format_rng(18).NumberFormat = "0.00%"
         
        For i = 19 To 20
            With format_rng(i)
                .Merge
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Font.Color = vbWhite
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
            End With
        Next i
        
        'Font Name, Font Size And General Setting
        .Cells.Font.Size = 8
        .Cells.Font.Name = "Verdana"
        .columns.AutoFit
        .Rows.AutoFit
    End With
    
    With erp_ws
        .Cells.Clear
        
        .Range("B6:D6").Value = Array("Location", "Dmd Group", "IRC Code")
        Dim format_erp() As Range
        ReDim format_erp(1 To 4)
        Set format_erp(1) = .Cells(2, 3)
        Set format_erp(2) = .Cells(4, 3)
        Set format_erp(3) = .Range(.Cells(7, 2), .Cells(npd_row + 5, 2))
        Set format_erp(4) = .Range(.Cells(7, 3), .Cells(npd_row + 5, 3))
        format_erp(1).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:="(1) Final Forecast by IRC with common ICW (in Units), (2) Forecast by IRC including Customer Weight/ICW (in Units)"
        
        For i = 1 To 4
            With format_erp(i)
                .Interior.Color = RGB(255, 242, 204) 'Gold, accent 4 , lighter 80%
            End With
        Next i
        
        With .Range("E5")
            .Value = "Add pre-shipment values if needed below"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        With .Range("E5:G5")
            .Merge
        End With
        With .Range(.Cells(5, 5), .Cells(5 + npd_row, 7))
            .Interior.Color = vbYellow
            .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        End With
        
        
        With .Range("B2")
            .Value = "Pick Forecast Table to be uploaded to DemandPlanner/LegacyERP"
            .Interior.Color = RGB(0, 36, 84) ' Brand Blue
            .Font.Color = vbWhite
            .Font.Bold = True
        End With
        
        With .Range("B4")
            .Value = "Forecast ID"
            .Interior.Color = RGB(0, 36, 84) ' Brand Blue
            .Font.Color = vbWhite
            .Font.Bold = True
        End With
           
        With .Range("B6:S6")
            .Interior.Color = RGB(0, 36, 84) ' Brand Blue
            .Font.Color = vbYellow
            .Font.Bold = True
        End With
        
        'Copy IRCs from NPD Info
        For r = 7 To (npd_row + 5)
            .Cells(r, 4).Formula = "=" & npd_ws.Cells(r, 3).Address(External:=True)
        Next r
            
        'Copy ICW and generate 12M
        .Cells(6, 8).Value = "=" & npd_ws.Range("I5").Address(External:=True) 'Copy-paste the Month 1 date from NPD's input
        .Range(.Cells(6, 5), .Cells(6, 19)).NumberFormat = "mmm-yy" 'Set the date format for 12 months
        For c = 7 To 5 Step -1
            .Cells(6, c).Formula = "=EDATE(" & .Cells(6, c + 1).Address & ", -1)" 'Using Edate function by increasing 1 month from Month 1
        Next c
        For c = 9 To 19
            .Cells(6, c).Formula = "=EDATE(" & .Cells(6, c - 1).Address & ", 1)" 'Using Edate function by increasing 1 month from Month 1
        Next c
        
        For r = 7 To 5 + npd_row
            For c = 8 To 19
                .Cells(r, c).Formula = "=IF(" & .Cells(2, 3).Address & "=""(1) Final Forecast by IRC with common ICW (in Units)"", " & _
                        "ROUND(" & fcst_ws.Cells(r - 1, c + 1).Address(External:=True) & ",0), IF(" & _
                        .Cells(2, 3).Address & "=""(2) Forecast by IRC including Customer Weight/ICW (in Units)"", " & _
                        """The Customers' Inputs are not updated""" & _
                        ",""""))"
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
    
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    
    MsgBox "The Final Forecast is generated." & vbNewLine & _
        "If there are different ICW date amongst customers. Please input the Customer Info Table and" & vbNewLine & _
        "then click Update Customer Info button", vbInformation
    
End Sub



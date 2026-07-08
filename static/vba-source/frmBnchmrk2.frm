Attribute VB_Name = "frmBnchmrk2"
Attribute VB_Base = "0{94B95225-9D43-4A73-B36F-F658E999C5FF}{3C9A849A-2DB7-4432-8311-2A7064FB9991}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Sub SetUpApplicationSettings()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
End Sub

Private Sub ResetApplicationSettings()
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub

Private Sub quickSort(Arr As Variant, ByVal left As Long, ByVal right As Long)
    Dim i As Long, j As Long, pivot As Variant, temp As Variant
    
    i = left
    j = right
    pivot = Arr((left + right) \ 2)
    
    While (i <= j)
        While (Arr(i) < pivot And i < right)
            i = i + 1
        Wend
        
        While (pivot < Arr(j) And j > left)
            j = j - 1
        Wend
        
        If (i <= j) Then
            temp = Arr(i)
            Arr(i) = Arr(j)
            Arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Wend
    
    If (left < j) Then
        Call quickSort(Arr, left, j)
    End If
    
    If (i < right) Then
        Call quickSort(Arr, i, right)
    End If
End Sub
Private Function SortDictionary(Dict As Object) As Variant
    Dim Arr() As Variant
    Dim i As Long
    
    'Convert dictionary to array
    ReDim Arr(1 To Dict.count)
    For i = 1 To Dict.count
        Arr(i) = Dict.Keys()(i - 1)
    Next i
    
    'Sort array
    Call quickSort(Arr, 1, UBound(Arr))
    
    'Return sorted array
    SortDictionary = Arr
End Function
Private Function FindColumnNumber(sheet As Worksheet, headerName As String) As Long
    Dim col As Range
    Dim msg As String

    For Each col In sheet.Rows(1).Cells
        If col.Value = headerName Then
            FindColumnNumber = col.Column
            msg = "Column " & headerName & " found at position " & col.Column
            Exit Function
        End If
    Next col
    FindColumnNumber = 0

End Function

Private Function GetDataArray(data_ws As Worksheet, colNumber As Long) As Variant
    If colNumber > 0 Then
        Dim lastRow As Long
        lastRow = data_ws.Cells(data_ws.Rows.count, colNumber).End(xlUp).Row
        GetDataArray = data_ws.Range(data_ws.Cells(2, colNumber), data_ws.Cells(lastRow, colNumber)).Value
    Else
        GetDataArray = Array() 'Return an empty array
    End If
End Function


Private Sub UserForm_Initialize()
    
    SetUpApplicationSettings
    
    Dim Region() As Variant, House() As Variant
    Dim UniqueRegion As Object, UniqueHouse As Object
    Dim i As Long
    Dim data_ws As Worksheet: Set data_ws = Worksheets("0. Database")
    
    Dim colRegion As Long, colHouse As Long
    colRegion = FindColumnNumber(data_ws, "Region")
    colHouse = FindColumnNumber(data_ws, "House")
    
    'Read data into arrays
    Region = GetDataArray(data_ws, colRegion)
    House = GetDataArray(data_ws, colHouse)

    'Check if arrays are not empty
    If UBound(Region) < 1 Then
        MsgBox "Column headers of Region were not found. No Region/Sub Region/Market could be selected."
    Else
        'Get unique values using a dictionary object
        Set UniqueRegion = CreateObject("Scripting.Dictionary")
        For i = 1 To UBound(Region)
            If Not UniqueRegion.Exists(Region(i, 1)) Then
                UniqueRegion.Add Region(i, 1), ""
            End If
        Next i
        
        'Sort the unique values
        Dim SortedRegion() As Variant
        SortedRegion = SortDictionary(UniqueRegion)
        
        'Populate the combo boxes with the unique values
        With Me.Region
            .list = SortedRegion
        End With
        
    End If
    
    If UBound(House) < 1 Then
        MsgBox "Column headers of House were not found. Please check the header names."
        Exit Sub
    Else
        Set UniqueHouse = CreateObject("Scripting.Dictionary")
        For i = 1 To UBound(House)
            If Not UniqueHouse.Exists(House(i, 1)) Then
                UniqueHouse.Add House(i, 1), ""
            End If
        Next i
        
        Dim SortedHouse() As Variant
        SortedHouse = SortDictionary(UniqueHouse)
        
        With Me.House
            .list = SortedHouse
        End With
    
    End If
    
    ResetApplicationSettings
End Sub

Private Sub Region_Change()
    SetUpApplicationSettings
    
    'Generate Sub-Region Combo Box based on the selected Region
    Dim sub_region() As Variant
    Dim unique_sub_region As Object
    Set unique_sub_region = CreateObject("Scripting.Dictionary")
    
    Dim data_ws As Worksheet
    Set data_ws = Worksheets("0. Database")
    
    Dim colSubRegion As Long, colRegion As Long
    colSubRegion = FindColumnNumber(data_ws, "Sub Region")
    colRegion = FindColumnNumber(data_ws, "Region")
    
    sub_region = GetDataArray(data_ws, colSubRegion)
    
    If UBound(sub_region) < 1 Then
        MsgBox "Column headers of Sub Region was not found. No Sub Region/Market could be selected"
    Else
        Dim i As Long
        For i = 1 To UBound(sub_region)
            If data_ws.Cells(i + 1, colRegion).Value = Me.Region.Value Then
                unique_sub_region(sub_region(i, 1)) = ""
            End If
        Next i
        
        Dim sorted_sub_region() As Variant
        sorted_sub_region = SortDictionary(unique_sub_region)
        
        With Me.SubRegion
            .Clear
            .list = sorted_sub_region
        End With
    End If
    
    ResetApplicationSettings
End Sub

Private Sub SubRegion_Change()
    SetUpApplicationSettings
    
    Dim selected_sub_region As String
    selected_sub_region = Me.SubRegion.Value
    
    Dim data_ws As Worksheet
    Set data_ws = Worksheets("0. Database")
    
    Dim markets() As Variant
    Dim colMarkets As Long, colSubRegion As Long
    colMarkets = FindColumnNumber(data_ws, "Market")
    colSubRegion = FindColumnNumber(data_ws, "Sub Region")
    
    markets = GetDataArray(data_ws, colMarkets)
    
    If UBound(markets) < 1 Then
        MsgBox "Column headers of Market was not found. No Market can be selected"
    Else
        Dim unique_markets As Object
        Set unique_markets = CreateObject("Scripting.Dictionary")
        Dim i As Long
        For i = 1 To UBound(markets)
            If data_ws.Cells(i + 1, colSubRegion).Value = selected_sub_region Then
                unique_markets(markets(i, 1)) = True
            End If
        Next i
        
        With Me.Market
            .Clear
            .list = unique_markets.Keys
            .MultiSelect = fmMultiSelectMulti
        End With
    End If
    
    ResetApplicationSettings
    
End Sub

Private Sub House_Change()

    SetUpApplicationSettings

    'Generate Product Line Combo Box based on the selected House
    Dim product_line() As Variant
    Dim unique_product_line As Object
    Set unique_product_line = CreateObject("Scripting.Dictionary")
    
    Dim data_ws As Worksheet
    Set data_ws = Worksheets("0. Database")
    
    Dim colHouse As Long, colProductLine As Long
    colHouse = FindColumnNumber(data_ws, "House")
    colProductLine = FindColumnNumber(data_ws, "Product Line")
    
    'Read all the product lines from the database worksheet into an array
    product_line = GetDataArray(data_ws, colProductLine)
    
    If UBound(product_line) < 1 Then
        MsgBox "Column headers of Product Line was not found. Please check the header names."
        Exit Sub
    Else
        'Loop through the product lines array to find the unique values for the selected house
        Dim i As Long
        For i = 1 To UBound(product_line)
            If data_ws.Cells(i + 1, colHouse).Value = Me.House.Value Then
                unique_product_line(product_line(i, 1)) = ""
            End If
        Next i
        
        'Sort the unique values
        Dim sorted_product_line() As Variant
        sorted_product_line = SortDictionary(unique_product_line)
        
        'Populate the Product Line list box with the unique product lines for the selected house
        With Me.ProductLine
            .Clear 'Clear the previous items in the combo box
            .list = sorted_product_line
            .MultiSelect = fmMultiSelectMulti ' Enable multiselect
        End With
    End If
    
    ResetApplicationSettings
    
End Sub
Private Sub ProductLine_Change()
    SetUpApplicationSettings
    
    Dim selected_product_line As String
    Dim data_ws As Worksheet
    Set data_ws = Worksheets("0. Database")
    
    Dim market_category() As Variant
    Dim colMarketCategory As Long, colProductLine As Long
    colMarketCategory = FindColumnNumber(data_ws, "Market Category")
    colProductLine = FindColumnNumber(data_ws, "Product Line")
    
    market_category = GetDataArray(data_ws, colMarketCategory)
    
    If UBound(market_category) < 1 Then
        MsgBox "Column headers of Market Category was not found. Please check the header names."
        Exit Sub
    Else
        Dim unique_category As Object
        Set unique_category = CreateObject("Scripting.Dictionary")
        
        Dim i As Long
        For i = 1 To UBound(market_category)
            ' Loop through each selected item in the listbox
            Dim j As Integer
            For j = 0 To Me.ProductLine.ListCount - 1
                If Me.ProductLine.Selected(j) And data_ws.Cells(i + 1, colProductLine).Value = Me.ProductLine.list(j) Then
                    unique_category(market_category(i, 1)) = True
                End If
            Next j
        Next i
        
        With Me.MarketCategory
            .Clear
            .list = unique_category.Keys
            .MultiSelect = fmMultiSelectMulti
        End With
    End If
    
    ResetApplicationSettings
    
End Sub
Private Sub ToggleButton_Click()

    SetUpApplicationSettings

    If ToggleButton.Value = True Then
        'Display the name of the new benchmark
        ToggleButton.Caption = "New Benchmark Selection "
    Else
        'Display Topline's benchmark
        ToggleButton.Caption = "Used Benchmark in Topline Fcst"
    End If
    
    ResetApplicationSettings
    
End Sub
Function IsValidDateFormat(val As String) As Boolean
    On Error Resume Next
    Dim temp As Date
    temp = Application.WorksheetFunction.text(val, "m/d/yyyy")
    If Err.Number = 0 Then
        IsValidDateFormat = True
        Exit Function
    End If
    Err.Clear
    temp = Application.WorksheetFunction.text(val, "d/m/yyyy")
    If Err.Number = 0 Then
        IsValidDateFormat = True
        Exit Function
    End If
    Err.Clear
    IsValidDateFormat = False
End Function
Private Sub ExtractData_Click()

    SetUpApplicationSettings
    
    Dim startTime As Double
    Dim endTime As Double
    Dim runningTime As Double
    
    ' Record the start time
    startTime = Timer
    
    'Define all Worksheets
    Dim db_ws As Worksheet, tl_ws As Worksheet, npd_ws As Worksheet, spl_ws As Worksheet, fcst_ws As Worksheet
    With ThisWorkbook
        Set db_ws = .Worksheets("0. Database")
        Set npd_ws = .Worksheets("1. NPD Info")
        Set tl_ws = .Worksheets("2. Topline Fcst")
        Set spl_ws = .Worksheets("3. Split by SKUs")
        Set fcst_ws = .Worksheets("4. Final Fcst")
    End With
    
    With spl_ws
        .Cells.Clear 'Clear the whole Split by SKUs Sheet
        
        On Error Resume Next 'Ignore errors temporarily
        .columns("G:R").Ungroup
        On Error GoTo 0 'Reset error handling
        
    End With
    
    If Yes1.Value = False And No1.Value = False Then
        MsgBox "Please Tick Yes or No for Use Topline Benchmark for the Split by SKU", vbCritical
        Exit Sub
    End If
    
    If Yes1.Value = True And Option1.Value = False And Option2.Value = False Then
        MsgBox "Please Pick Option 1 or 2 for the Split by Month approach", vbCritical
        Exit Sub
    End If
    
    If Yes2.Value = False And No2.Value = False Then
        MsgBox "Please Tick Yes or No for Include pipefill period from Benchmark ", vbCritical
        Exit Sub
    End If
    
    If Yes1.Value = True And No1.Value = True Then
        MsgBox "Please Tick only 1 option Yes or No for Use Topline Benchmark for the Split by SKU", vbCritical
        Exit Sub
    End If
    
    If Yes2.Value = True And No2.Value = True Then
        MsgBox "Please Tick only 1 option Yes or No for Include pipefill period from Benchmark ", vbCritical
        Exit Sub
    End If
    
    With tl_ws
        
        Dim newStartRow As Long, begin_cell As Long, last_row As Long, last_col As Long
        newStartRow = .Range("B11").CurrentRegion.Rows.count + 14 'First row of Option 2 Table
        begin_cell = .Range("B" & newStartRow).CurrentRegion.Rows.count * 2 + 17 'First row of Assumption Table
        last_row = .Range("B11").CurrentRegion.Rows.count
        last_col = .Range("B11").CurrentRegion.columns.count
            
        'Detect if Option 1 or 2 is selected in Topline Forecast
        If .Cells(begin_cell + 1, 5).Value <> "Option 1" And .Cells(begin_cell + 1, 5).Value <> "Option 2" Then
            MsgBox "Please select Option 1 or 2 in cell" & Cells(begin_cell + 1, 5).Address & " in Topline Fcst Worksheet", vbExclamation
            Exit Sub
        End If
         
'If Yes selected for Topline Benchmark for the Split by SKU
        If Yes1.Value = True Then
            'Detect specific rows to be blank, otherwise it will affect .CurrentRegion syntax
            Dim blank_row(1 To 5) As Long
            blank_row(1) = 10
            blank_row(2) = newStartRow - 3
            blank_row(3) = newStartRow - 1
            blank_row(4) = begin_cell - 3
            blank_row(5) = begin_cell - 1
            
            Dim i As Long
            For i = 1 To 5
                 If WorksheetFunction.CountA(.Rows(blank_row(i))) > 0 Then
                    MsgBox "Row " & blank_row(i) & " must be left blank. Please move it to other position", vbExclamation
                    Exit Sub
                End If
            Next i
            
            'Detect specific columns to be blank
            Dim blank_col() As Range
            Dim col_rng As Variant
            
            ReDim blank_col(1 To 4)
            Set blank_col(1) = .Range(.Cells(10, 1), .Cells(newStartRow - 3, 1))
            Set blank_col(2) = .Range(.Cells(newStartRow - 1, 1), .Cells(begin_cell - 3, 1))
            Set blank_col(3) = .Range(.Cells(newStartRow - 1, 20), .Cells(begin_cell - 3, 20))
            Set blank_col(4) = .Range(.Cells(10, last_col + 2), .Cells(newStartRow - 3, last_col + 2))
            
            For Each col_rng In blank_col
                If WorksheetFunction.CountA(col_rng) > 0 Then
                    MsgBox "Please leave the range " & col_rng.Address & " blank Please move it to other position.", vbExclamation
                    Exit Sub
                End If
            Next col_rng
        
            'Copy-paste Option 1 or 2 into Split by SKUs Worksheet
            If Option1.Value = True Then
            .Range("B11").Resize(last_row - 1, 17).Copy Destination:=spl_ws.Range("B11")
            ElseIf Option2.Value = True Then
               .Range("B" & newStartRow).CurrentRegion.Resize(last_row - 1, _
                    .Range("B" & newStartRow).CurrentRegion.columns.count - 1).Copy Destination:=spl_ws.Range("B11")
            End If
        End If
    End With
    
'If No selected for Topline Benchmark for the Split by SKU
    If No1.Value = True Then 'Check if New Benchmark selected
        With db_ws
            'Clear the filters from the database worksheet + Clear Toplie Forecast worksheet
            .AutoFilterMode = False
        
            'Get the selected values from the combo boxes
            Dim Region As String: Region = ""
            If Not IsNull(Me.Region.Value) Then Region = Me.Region.Value
            
            Dim sub_region As String: sub_region = ""
            If Not IsNull(Me.SubRegion.Value) Then sub_region = Me.SubRegion.Value
            
            Dim House As String: House = ""
            If Not IsNull(Me.House.Value) Then House = Me.House.Value
            
            Dim RegionCol As Long: RegionCol = FindColumnNumber(db_ws, "Region") - 1
            Dim SubRegionCol As Long: SubRegionCol = FindColumnNumber(db_ws, "Sub Region") - 1
            Dim HouseCol As Long: HouseCol = FindColumnNumber(db_ws, "House") - 1
        
            'Filter the database worksheet based on the selected values
            If Region <> "" Then .Range("A1").AutoFilter Field:=RegionCol, criteria1:=Region
            If sub_region <> "" Then .Range("A1").AutoFilter Field:=SubRegionCol, criteria1:=sub_region
            If House <> "" Then .Range("A1").AutoFilter Field:=HouseCol, criteria1:=House
            
            'If no selection in House, pop up an error message
            If House = "" Then
                MsgBox "Please filtering House/Product Line/Market Category boxes" & vbNewLine & _
                        "because the selection No is ticked for Use Topline Benchmark for the Splits by SKUs", vbCritical
                Exit Sub 'End the procedure to prevent further execution
            End If
            
            'Filter based on Market List Box
            Dim markets_selected As String
            Dim markets_array() As String
            Dim selected_count As Long
            
            selected_count = 0
            For i = 0 To Me.Market.ListCount - 1
                If Me.Market.Selected(i) Then
                    selected_count = selected_count + 1
                End If
            Next i
    
            If selected_count > 0 Then
                'Convert the selected values in the ListBox to a comma-separated string
                For i = 0 To Me.Market.ListCount - 1
                    If Me.Market.Selected(i) Then
                        markets_selected = markets_selected & Me.Market.list(i) & ","
                    End If
                Next i
                
                'Remove the last comma from the string
                markets_selected = left(markets_selected, Len(markets_selected) - 1)
                
                'Split the comma-separated string into an array
                markets_array = Split(markets_selected, ",")
                
                'Apply filter to database worksheet based on selected Markets
                Dim MarketCol As Long: MarketCol = FindColumnNumber(db_ws, "Market") - 1
                With .Range("A1")
                    .AutoFilter Field:=MarketCol, criteria1:=markets_array, Operator:=xlFilterValues
                End With
            End If
        
        'Filter based on Product Line List Box
            Dim product_line_selected As String
            Dim product_line_array() As String
            Dim selected_line_count As Long
        
        'Check if any values are selected in the ListBox
            selected_line_count = 0
            For i = 0 To Me.ProductLine.ListCount - 1
                If Me.ProductLine.Selected(i) Then
                    selected_line_count = selected_line_count + 1
                End If
            Next i
    
            If selected_line_count > 0 Then
                'Convert the selected values in the ListBox to a comma-separated string
                For i = 0 To Me.ProductLine.ListCount - 1
                    If Me.ProductLine.Selected(i) Then
                        product_line_selected = product_line_selected & Me.ProductLine.list(i) & ","
                    End If
                Next i
                
                'Remove the last comma from the string
                product_line_selected = left(product_line_selected, Len(product_line_selected) - 1)
                
                'Split the comma-separated string into an array
                product_line_array = Split(product_line_selected, ",")
                
                'Apply filter to database worksheet based on selected Market Categories
                Dim ProductLineCol As Long: ProductLineCol = FindColumnNumber(db_ws, "Product Line") - 1
                With .Range("A1")
                    .AutoFilter Field:=ProductLineCol, criteria1:=product_line_array, Operator:=xlFilterValues
                End With
            End If
        
        'Filter based on Market Category List Box
            Dim market_category_selected As String
            Dim market_category_array() As String
            Dim selected_category_count As Long
        
        'Check if any values are selected in the ListBox
            selected_category_count = 0
            For i = 0 To Me.MarketCategory.ListCount - 1
                If Me.MarketCategory.Selected(i) Then
                    selected_category_count = selected_category_count + 1
                End If
            Next i
    
            If selected_category_count > 0 Then
                'Convert the selected values in the ListBox to a comma-separated string
                For i = 0 To Me.MarketCategory.ListCount - 1
                    If Me.MarketCategory.Selected(i) Then
                        market_category_selected = market_category_selected & Me.MarketCategory.list(i) & ","
                    End If
                Next i
                
                'Remove the last comma from the string
                market_category_selected = left(market_category_selected, Len(market_category_selected) - 1)
                
                'Split the comma-separated string into an array
                market_category_array = Split(market_category_selected, ",")
                
                'Apply filter to database worksheet based on selected Market Categories
                Dim CategoryCol As Long: CategoryCol = FindColumnNumber(db_ws, "Market Category") - 1
                With .Range("A1")
                    .AutoFilter Field:=CategoryCol, criteria1:=market_category_array, Operator:=xlFilterValues
                End With
            End If

            'Get the filtered range and copy/paste to Split by SKUs worksheet
            .AutoFilter.Range.Copy Destination:=spl_ws.Range("B11")
            
            'Clear the filters from the database worksheet
            .AutoFilterMode = False
        End With 'End db_ws
    
        'Section below is the VBA code for Option 2
            
        Dim rng As Range
        Dim small_value As Double
        
        'Define rng
        Set rng = spl_ws.Range("B11").CurrentRegion
        
        With rng
            
            'Delete specific columns except date and ...
             For i = .columns.count To 1 Step -1
                Dim col As Range: Set col = .columns(i)
                If Not IsValidDateFormat(col.Cells(1, 1).Value2) Then
                    Select Case col.Cells(1, 1).Value2
                        Case "GTIN", "IRC", "IRC Description", "Market Category", "Article Type"
                            ' Do nothing
                        Case Else
                            col.Delete
                    End Select
                End If
            Next i
            
            last_row = .Rows.count
            last_col = .columns.count
                
            For i = last_col To 6 Step -1
                Dim blank_count As Long: blank_count = WorksheetFunction.CountBlank(.columns(i).Resize(last_row - 1).Offset(1))
                If blank_count = last_row - 1 Then
                    .columns(i).Delete
                End If
            Next i
                
            last_col = .columns.count
            'Sum up rows of the same EANs
            Dim r As Long, c As Long
            For r = 2 To last_row ' Loop through rows 2 to last row
                Dim sum_value As Double
                sum_value = 0 ' Initialize sum value to zero
                For i = r + 1 To last_row ' Loop through remaining rows below the current row
                    If .Cells(r, 1).Value = .Cells(i, 1).Value Then ' If cell value in column 1 of current row matches with any cell value in column 1 of remaining rows
                        For c = 6 To last_col ' Loop through columns 6 to last column
                            .Cells(r, c).Value = .Cells(r, c).Value + .Cells(i, c).Value ' Add the value of the current cell in row r to the corresponding cell in the matching row i
                        Next c
                        .Rows(i).Delete ' Delete the matching row
                        last_row = last_row - 1 ' Decrement last row as one row is deleted
                        i = i - 1 ' Decrement i as one row is deleted
                    End If
                Next i
            Next r
                
            Dim row_num As Long, num_count As Long
            Dim col_sum As Double, avg_sum As Double
            last_row = .Rows.count
    
            'Replace 0 value by blank
            Dim cells_cleared As Boolean ' flag to check if any cells have been cleared
            cells_cleared = False
            
            For r = 2 To last_row
                For c = 6 To last_col
                    If .Cells(r, c).Value = 0 Then
                        .Cells(r, c).ClearContents
                        cells_cleared = True
                    End If
                Next c
            Next r
            If cells_cleared = True Then
                .SpecialCells(xlCellTypeBlanks).Delete Shift:=xlToLeft
                cells_cleared = False
            End If
            
            
            For r = 2 To last_row
                'Reset num_count and col_sum for each row
                num_count = 0
                col_sum = 0
                avg_sum = 0
                For c = 6 To 17
                    If IsNumeric(.Cells(r, c).Value2) And Not IsEmpty(.Cells(r, c).Value2) Then
                        num_count = num_count + 1 'Increment the count of numeric values
                        col_sum = col_sum + .Cells(r, c).Value2 'Add value to the column sum
                    End If
                Next c
                
                If num_count <> 0 Then
                    avg_sum = col_sum / num_count
               
                    For c = 6 To 17
                        If IsNumeric(.Cells(r, c).Value2) And Not IsEmpty(.Cells(r, c).Value2) And .Cells(r, c).Value2 < avg_sum Then
                            .Cells(r, c + 1).Value2 = .Cells(r, c + 1).Value2 + .Cells(r, c).Value2
                            .Cells(r, c).ClearContents
                            cells_cleared = True
                        Else
                            ' If the value is bigger than avg_sum, exit the loop
                            Exit For
                        End If
                    Next c
                    If cells_cleared = True Then
                        .SpecialCells(xlCellTypeBlanks).Delete Shift:=xlToLeft
                        cells_cleared = False
                    End If
                End If
                    
            Next r
            
            'Delete columns to have 12M left
            Dim col2 As Long
            For col2 = last_col To 18 Step -1
                .columns(col2).Delete Shift:=xlToLeft
            Next col2
                
            last_col = .columns.count 'Re-define columns count after deleting
            For col2 = 6 To 17
                .Cells(1, col2) = "Month" & " " & col2 - 5
            Next col2
                
            'Format split by SKUs Table
            .Offset(1, 0).Resize(last_row - 1, last_col).Cells.ClearFormats 'Clear format except the header row
                
            'Create Stripe Format
            Dim r2 As Long
            For r2 = 2 To last_row Step 2
                .Rows(r2).Interior.Color = RGB(142, 170, 219) 'Blue, Accent 1, Lighter 60%
            Next r2
                
            .columns(1).Offset(1, 0).Resize(last_row - 1, 1).NumberFormat = "0" 'Format GTIN to display full string
            .Offset(1, 5).NumberFormat = "#,##0" 'Format number to have comma and rounded
                
            Dim format_rng() As Range
            ReDim format_rng(1 To 2)
            Set format_rng(1) = .Rows(1).Resize(1, 5) 'code apply to the first 5 columns of row 1
            Set format_rng(2) = .Rows(1).Resize(1, last_col - 5).Offset(0, 5) 'code to apply to row 1, except first 5 columns
                
            For i = 1 To 2
                With format_rng(i)
                    .Font.Bold = True
                    .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                    .Font.Color = vbWhite
                    .VerticalAlignment = xlCenter
                    .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
                End With
            Next i
            format_rng(2).Font.Color = vbYellow
            
            With .columns(5).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
                    
            .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
            
        End With 'End With rng
    End If 'of when New Benchmark selected

'Define split table
    Dim split_tbl As Range: Set split_tbl = spl_ws.Range("B11").CurrentRegion
    With split_tbl
        Dim split_row As Long: split_row = .Rows.count
        Dim split_col As Long: split_col = .columns.count
        
    'Calculate Total 12M excl pipe & Weight (%) of IRC compared to Total 12M excl pipe
        .Cells(1, split_col + 2).Value = "Weight (%)"
        
        For r = 2 To split_row
            'Total 12M excl pipe or incl pip
            If No2.Value = True Then 'Exclude the pipe
                .Cells(1, split_col + 1).Value = "Total 12M (excl pipe)"
                num_count = 0
                col_sum = 0
                avg_sum = 0
                Dim split_c As Long
                For c = 6 To 17
                    If IsNumeric(.Cells(r, c).Value2) And Not IsEmpty(.Cells(r, c).Value2) Then
                        num_count = num_count + 1 'Increment the count of numeric values
                        col_sum = col_sum + .Cells(r, c).Value2 'Add value to the column sum
                    End If
                Next c
                
                If num_count <> 0 Then
                    avg_sum = col_sum / num_count
                
                    For c = 6 To 17
                        If .Cells(r, c).Value2 < avg_sum Then
                            split_c = c
                            Exit For
                        End If
                    Next c
                    
                    If split_c > 0 Then
                        .Cells(r, split_col + 1).Formula = "=SUM(" & .Cells(r, split_c).Address & ":" & .Cells(r, 17).Address & ")"
                    End If
                End If
            End If
            
            If Yes2.Value = True Then
                .Cells(1, split_col + 1).Value = "Total 12M (incl pipe)"
                .Cells(r, split_col + 1).Formula = "=SUM(" & .Cells(r, 6).Address & ":" & .Cells(r, 17).Address & ")"
            End If
            
            'Weight (%) of IRC compared to Total 12M excl pipe
            .Cells(r, split_col + 2).Formula = "=" & .Cells(r, split_col + 1).Address & _
                "/" & "SUM(" & .Cells(2, split_col + 1).Address & ":" & .Cells(split_row, split_col + 1).Address & ")"
        Next r
             
    
    'Calculate total for whole column Total 12M excl pipe & Weight(%) of IRC compared to Total 12M excl pipe
        With .Cells(split_row + 1, 5)
            .Value = "Total"
            .Font.Bold = True
            .Interior.Color = RGB(0, 36, 84) 'Brand Blue
            .Font.Color = vbWhite
            .VerticalAlignment = xlCenter
            .NumberFormat = "#,##0"
        End With
        For c = 6 To split_col + 2
            .Cells(split_row + 1, c).Formula = "=SUM(" & .Cells(2, c).Address & ":" & .Cells(split_row, c).Address & ")"
            With .Cells(split_row + 1, c)
                .Font.Bold = True
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .Font.Color = vbWhite
                .VerticalAlignment = xlCenter
                .NumberFormat = "#,##0"
            End With
        Next c
    End With
    
    With spl_ws
    'Create headers for Split by SKUs WorkSheet
        Dim split_headers() As Variant
        split_headers = Array("Match to NPD IRCs", " ", "NPD's New IRCs", "Line-up", "NPD's Shade/ML", _
                     "Weight based on Benchmark (%)", "Extrapolated Weight", "Topline Final Fcst * Extrapolated Weight", _
                     "Weight (%) (Overwrite here if desired)", "Overwrite Extrapolated Weight", _
                     "Topline Final Fcst * Overwrite Extrapolated Weight")
        With .Range(.Cells(11, 22), .Cells(11, 32))
            .Value = split_headers
        End With
        
        Dim npd_count As Long
        npd_count = npd_ws.Range("C6").End(xlDown).Row - 6 'Count the number of New Products
        
        With npd_ws
        'Create Concatenate formula for the last column exclude the header
            For r = 7 To 6 + npd_count
                .Cells(r, 9).Formula = "=CONCATENATE(" & .Cells(r, 6).Address & ",  "" "" ," & _
                .Cells(r, 5).Address & ",  "" "" ," & .Cells(r, 7).Address & ")"
            Next r
        End With
        
    'Calculation
        For r = 12 To (npd_count + 11)
            'Copy-paste NPD Info Tab into Split by SKUs Tab
            .Cells(r, 24).Formula = "=" & npd_ws.Cells(r - 5, 3).Address(External:=True)
            .Cells(r, 25).Formula = "=" & npd_ws.Cells(r - 5, 9).Address(External:=True)
            .Cells(r, 26).Formula = "=" & npd_ws.Cells(r - 5, 5).Address(External:=True)
             
            'Weight based on Benchmark
            .Cells(r, 27).Formula = "=SUMIF(" & .Cells(12, 22).Address & ":" & .Cells(split_row + 10, 22).Address & "," & _
                .Cells(r, 25).Address & "," & _
                .Cells(12, 20).Address & ":" & .Cells(split_row + 10, 20).Address & ")"

            'Extrapolated Weight
            .Cells(r, 28).Formula = "=IFERROR(" & .Cells(r, 27).Address & "*" & 1 & "/" & .Cells(npd_count + 12, 27).Address & ","""")"

            'Topline Final Fcst * Extrapolated Weight
            .Cells(r, 29).Formula = "=IFERROR(" & .Cells(r, 28).Address & "*" & _
                                        tl_ws.Cells(begin_cell + 25, 19).Address(External:=True) & ","""")"

            'Overwrite Weight
            .Cells(r, 30).Formula = "=" & .Cells(r, 27).Address
            
            'Overwrite Extrapolated Weight
            .Cells(r, 31).Formula = "=IFERROR(IF(" & .Cells(r, 30).Address & "<>"""", " & _
                .Cells(r, 30).Address & "*" & 1 & "/" & .Cells(npd_count + 12, 30).Address & "," & _
                .Cells(r, 28).Address & "),"""")"


            'Topline Final Fcst * Overwrite Extrapolated Weight
            .Cells(r, 32).Formula = "=IFERROR(" & .Cells(r, 31).Address & "*" & _
                                        tl_ws.Cells(begin_cell + 25, 19).Address(External:=True) & ","""")"

        Next r
         
    'Calculate the total of each column
        .Cells(npd_count + 12, 26).Value = "Total"
        For c = 27 To 32
            .Cells(npd_count + 12, c).Formula = "=IFERROR(SUM(" & .Cells(12, c).Address & ":" & .Cells(npd_count + 11, c).Address & "),"""")"
        Next c
        
    'Format Split by SKU Worksheet
        Dim split_format() As Range
        ReDim split_format(1 To 18)
        
        'Format percentage for Split by SKUs Sheet
        Set split_format(1) = .Range(.Cells(12, 20), .Cells(split_row + 11, 20)) 'Weight(%) column
        Set split_format(2) = .Range(.Cells(12, 27), .Cells(npd_count + 12, 28)) 'Weight based on Benchmark & Weight Extrapolated column
        Set split_format(3) = .Range(.Cells(12, 30), .Cells(npd_count + 12, 31)) 'Overwrite Weight & Overwrite Extrapolated Weight COlumn
        For i = 1 To 3
            With split_format(i)
                .NumberFormat = "0%"
            End With
        Next i
    
        'Outside border
        Set split_format(4) = .Range(.Cells(11, 22), .Cells(split_row + 10, 22)) 'Allocate into NDP's new IRCs column
        Set split_format(5) = .Range(.Cells(11, 24), .Cells(npd_count + 11, 32)) 'Whole Weight Table
        Set split_format(6) = .Range(.Cells(11, 19), .Cells(split_row + 10, 20)) 'Total 12M & Weight(%) Columns
        For i = 4 To 6
            With split_format(i)
                .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
            End With
        Next i
        
        Set split_format(7) = .Range(.Cells(npd_count + 12, 26), .Cells(npd_count + 12, 32)) 'Total Row of Weight Table
        For i = 6 To 7
            With split_format(i)
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .Font.Color = vbWhite
                .Font.Bold = True
            End With
        Next i
        
        'Bottom border & Color & Bold (font)
        Set split_format(8) = .Range(.Cells(11, 19), .Cells(11, 20))
        Set split_format(9) = .Range(.Cells(11, 24), .Cells(11, 26))
        For i = 8 To 9
            With split_format(i)
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlMedium
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .Font.Color = vbWhite
            End With
        Next i
        
        Set split_format(10) = .Range(.Cells(11, 27), .Cells(11, 29))
        Set split_format(11) = .Range(.Cells(11, 31), .Cells(11, 32))
        For i = 10 To 11
            With split_format(i)
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlMedium
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .Font.Color = vbYellow
            End With
        Next i
        
        Set split_format(12) = .Cells(11, 22) 'Allocate into new IRC Cell
        Set split_format(13) = .Cells(11, 30) 'Weight(Overwrite if desired) Cell
        For i = 12 To 13
            With split_format(i)
                .Font.Color = vbRed
                .Interior.Color = RGB(255, 217, 102) 'Gold, Accent 4, Lighter 60%
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlMedium
            End With
        Next i
        
        Set split_format(14) = .Range(.Cells(11, 26), .Cells(npd_count + 12, 26)) 'NPD's Shade/Ml Col
        Set split_format(15) = .Range(.Cells(11, 29), .Cells(npd_count + 12, 29)) 'Topline Final Fcst * Extrapolated Weight Col
        For i = 14 To 15
            With split_format(i).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        Next i
        
        Set split_format(16) = .Range(.Cells(11, 32), .Cells(npd_count + 12, 32)) 'Topline Final Fcst * OverWrited Extrapolated Weight Col
        For i = 15 To 16
            split_format(i).NumberFormat = "#,##0"
        Next i
        
        Set split_format(17) = .Range(.Cells(12, 22), .Cells(split_row + 10, 22)) 'Allocate into new IRCs excl cell header
        Set split_format(18) = .Range(.Cells(12, 30), .Cells(npd_count + 11, 30)) 'Weight (%) (Overwrite if desired) excl cell header
        For i = 17 To 18
            With split_format(i)
                .Interior.Color = RGB(255, 242, 204) 'Gold, accent 4 , lighter 80%
                .Font.Bold = True
            End With
        Next i
    
        'Automatically Allocation
        
    
        'Create drop-down menu for Allocate into NDP's new IRCs
        Dim source_range As Range: Set source_range = npd_ws.Range(npd_ws.Cells(7, 9), npd_ws.Cells(npd_count + 6, 9))
        Dim dropdown_range As Range: Set dropdown_range = .Range(.Cells(12, 22), .Cells(split_row + 10, 22))
        
        With dropdown_range.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="=" & source_range.Address(External:=True)
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .errorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    
        With .Range("B1:C1")
            .Value = Array("Macro/VBA last run at:", Format(Now(), "mm/dd/yyyy hh:mm:ss"))
            .Font.Color = vbRed
            .Font.Italic = True
        End With
        
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
        
        .Cells.Font.Size = 8
        .Cells.Font.Name = "Verdana"
        .columns.AutoFit
        .Rows.AutoFit
        .Range("W:W").ColumnWidth = 10
        .Range("V:V").ColumnWidth = 30
        
        On Error Resume Next 'Ignore errors temporarily
        .columns("G:R").Group
        On Error GoTo 0 'Reset error handling

        
    End With
    
    ResetApplicationSettings
    
    MsgBox "Done", vbInformation
    
End Sub


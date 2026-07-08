Attribute VB_Name = "frmBnchmrk1"
Attribute VB_Base = "0{8451CC14-8128-4D5D-938F-4F6383E370ED}{E4CF1206-D618-44DC-A0DA-9287AD49414C}"
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
    
    Dim i As Long, r As Long, j As Long, k As Long, c As Long
    
    'Define Worksheets
    Dim db_ws As Worksheet, tl_ws As Worksheet, npd_ws As Worksheet, spl_ws As Worksheet, fcst_ws As Worksheet
    With ThisWorkbook
        Set db_ws = .Worksheets("0. Database")
        Set npd_ws = .Worksheets("1. NPD Info")
        Set tl_ws = .Worksheets("2. Topline Fcst")
        Set spl_ws = .Worksheets("3. Split by SKUs")
        Set fcst_ws = .Worksheets("4. Final Fcst")
    End With
    
    With npd_ws
        'Detect if ICW date is typed in
        If IsEmpty(.Range("I5").Value) Then
            MsgBox "Error: Please enter the ICW date of new product in cell I5 in 1. NPD Info Worksheet" & vbNewLine & _
                        "And then come back to this Worksheet", vbCritical, "Error"
            Exit Sub
        End If
    End With
    
    With db_ws
        'Clear the filters from the database worksheet + Clear Toplie Forecast worksheet
        .AutoFilterMode = False
        tl_ws.Cells.Clear
        
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
        
        'Filter based on Market List Box
        Dim markets_selected As String
        Dim markets_array() As String
        Dim selected_count As Long
        
        'Count the number of selected items in the ListBox to check if any values are selected
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

'Section below is the VBA code for Option 1
    'Get the filtered range and copy/paste to Topline Fcst worksheet
        .AutoFilter.Range.Copy Destination:=tl_ws.Range("B11")
    
    'Clear the filters from the database worksheet
        .AutoFilterMode = False
    End With
    
    'Define the current region range for the copied data
    Dim filteredTable As Range: Set filteredTable = tl_ws.Range("B11").CurrentRegion
    
    With filteredTable
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
        
        Dim last_row As Long: last_row = .Rows.count
        Dim last_col As Long: last_col = .columns.count
        
        For i = last_col To 6 Step -1
            Dim blank_count As Long: blank_count = Application.WorksheetFunction.CountBlank(.columns(i).Resize(last_row - 1).Offset(1))
            If blank_count = last_row - 1 Then
                ' Check if column is blank before attempting to delete it
                If Not IsEmpty(.columns(i)) Then
                    .columns(i).Delete
                End If
            End If
        Next i
        
        Set filteredTable = tl_ws.Range("B11").CurrentRegion
        last_row = .Rows.count
        last_col = .columns.count
        
        'Sum up rows of the same EANs
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
        
        
        For c = 6 To last_col + 1
            .Cells(last_row + 1, c).Formula = "=SUM(" & .Cells(2, c).Address & ":" & .Cells(last_row, c).Address & ")"
            With .Cells(last_row + 1, c)
                .Font.Bold = True
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .Font.Color = vbWhite
                .VerticalAlignment = xlCenter
            End With
        Next c
        
        For r = 2 To last_row
            .Cells(r, last_col + 1).Formula = "=SUM(" & .Cells(r, 6).Address & ":" & .Cells(r, last_col).Address & ")"
        Next r
        
        'Format filteredTable (Option 1)
        .Offset(1, 0).Resize(last_row - 1, last_col).Cells.ClearFormats 'Clear format except the header row
        
        'Create Stripe Format
        For r = 2 To last_row Step 2
            .Rows(r).Interior.Color = RGB(142, 170, 219) 'Blue, Accent 1, Lighter 60%
        Next r
        
        .columns(1).Offset(1, 0).Resize(last_row - 1, 1).NumberFormat = "0"
        
        Dim format_filteredTable() As Range
        ReDim format_filteredTable(1 To 5)
        Set format_filteredTable(1) = .Rows(1).Resize(1, 5) 'code apply to the first 5 columns of row 1
        Set format_filteredTable(2) = .Rows(1).Resize(1, last_col - 5).Offset(0, 5) 'code to apply to row 1, except first 5 columns
        Set format_filteredTable(3) = .columns(last_col + 1) 'code apply to the last column+1
        Set format_filteredTable(4) = .Resize(last_row, last_col + 1).Cells(1, last_col + 1) 'code apply to the last column's cell
        Set format_filteredTable(5) = .Resize(last_row + 1, last_col).Cells(last_row + 1, 5) 'code apply to the last row's Cell
        
        For i = 1 To 5
            With format_filteredTable(i)
                .Font.Bold = True
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .Font.Color = vbWhite
                .VerticalAlignment = xlCenter
                .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
            End With
        Next i
        
        format_filteredTable(4).Value = "Total"
        format_filteredTable(5).Value = "Total (Option 1)"
        format_filteredTable(2).Font.Color = vbYellow
        
        With .columns(5).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        
        .Offset(1, 5).NumberFormat = "#,##0"
        
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With

'Section below is the VBA code for Option 2
    Dim newStartRow As Long
    Dim rng As Range
    Dim small_value As Double
    
    'Define rng As Option 2 Table
    With tl_ws
        newStartRow = .Range("B11").CurrentRegion.Rows.count + 14
    
        .Range("B11").CurrentRegion.Resize(.Range("B11").CurrentRegion.Rows.count - 1, _
        .Range("B11").CurrentRegion.columns.count - 1).Copy .Range("B" & newStartRow)
        
        last_row = .Range("B" & newStartRow).CurrentRegion.Rows.count
        last_col = .Range("B" & newStartRow).CurrentRegion.columns.count
        Set rng = .Range("B" & newStartRow).CurrentRegion.Resize(last_row, last_col)
    End With
    
    With rng
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
        
        Dim row_num As Long, num_count As Long
        Dim col_sum As Double, avg_sum As Double
        
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
            
            
                'Replace the values as average
                If Yes.Value = True Then
                    For i = 6 To 17
                        If IsNumeric(.Cells(r, i).Value2) And Not IsEmpty(.Cells(r, i).Value2) And .Cells(r, i).Value2 < avg_sum Then
                            Dim sum_dmd As Double
                            Dim count_dmd As Long
                            Dim coefficient As Integer
                            
                            sum_dmd = 0
                            count_dmd = 0
                            coefficient = 17 - i + 1
                            
                            ' calculate the sum and count of the values from i to 17
                            For c = i To 17
                                If Not IsEmpty(.Cells(r, c).Value2) Then
                                    sum_dmd = sum_dmd + .Cells(r, c).Value2
                                    count_dmd = count_dmd + 1
                                End If
                            Next c
                            
                            ' calculate the average
                            Dim avg_dmd As Double
                            If count_dmd > 0 Then
                                avg_dmd = sum_dmd / coefficient
                                ' replace the values from i to 17 with the average
                                For c = i To 17
                                    .Cells(r, c).Value2 = avg_dmd
                                Next c
                            End If
                            
                            Exit For
                            
                        End If
                    Next i
                End If
            End If
        Next r
       
    
        'Delete columns to have 12M left
        Dim col2 As Long
        For col2 = last_col To 18 Step -1
            .columns(col2).Delete Shift:=xlToLeft
        Next col2
        
        For col2 = 6 To 17
            .Cells(1, col2) = "Month" & " " & col2 - 5
        Next col2
        
        'Re-define last_row and last_col after deleting
        last_row = .Rows.count
        last_col = .columns.count

        For c = 6 To last_col + 1
            .Cells(last_row + 1, c).Formula = "=SUM(" & .Cells(2, c).Address & ":" & .Cells(last_row, c).Address & ")"
            With .Cells(last_row + 1, c)
                .Font.Bold = True
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .Font.Color = vbWhite
                .VerticalAlignment = xlCenter
            End With
        Next c

        For r = 2 To last_row
            .Cells(r, last_col + 1).Formula = "=SUM(" & .Cells(r, 6).Address & ":" & .Cells(r, last_col).Address & ")"
        Next r
        
    'Format Rng (Option 2)
        'Create Stripe Format
        For r = 2 To last_row Step 2
            .Rows(r).Interior.Color = RGB(142, 170, 219) 'Blue, Accent 1, Lighter 60%
        Next r
        
        For r = 2 To last_row + 1
            For c = 6 To last_col + 1
                .Cells(r, c).NumberFormat = "#,##0"
            Next c
        Next r
        
        Dim format_rng() As Range
        ReDim format_rng(1 To 4)
        Set format_rng(1) = .columns(last_col + 1) 'code apply to the last column+1
        Set format_rng(2) = .Resize(last_row, last_col + 1).Cells(1, last_col + 1) 'code apply to the last column's cell
        Set format_rng(3) = .Resize(last_row + 1, last_col).Cells(last_row + 1, 5) 'code apply to the last row's Cell
       
        For i = 1 To 3
            With format_rng(i)
                .Font.Bold = True
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .Font.Color = vbWhite
                .VerticalAlignment = xlCenter
                .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
            End With
        Next i
        
        format_rng(2).Value = "Total 12M"
        format_rng(3).Value = "Total (Option 2)"
        
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With
    
'Section below is the VBA for Assumption Table
    Dim begin_cell As Long
    begin_cell = rng.Rows.count * 2 + 19 'Note: rng (Option 2 Table) at this moment is defined without the Total(Option 2) row
    
    With tl_ws
        With .Range("A9")
            .Value = "Option 1 (As-is in market data):"
            .Font.Color = RGB(0, 36, 84) 'Brand Blue
            .Interior.Color = RGB(191, 191, 191) 'Gray, Accent 3, Lighter 60%
            .Font.Bold = True
        End With

        With .Range("A" & newStartRow - 2)
            .Value = "Option 2 (Automatically Adjusted):"
            .Font.Color = RGB(0, 36, 84) 'Brand Blue
            .Interior.Color = RGB(191, 191, 191) 'Gray, Accent 3, Lighter 60%
            .Font.Bold = True
        End With
        
        .Cells(begin_cell, 7).Value = npd_ws.Range("I5").Value
        For i = 8 To 18
            .Cells(begin_cell, i).Value = Application.WorksheetFunction.EDate(.Cells(begin_cell, i - 1), 1)
        Next i
    
        Dim tl_row As Range: Set tl_row = .Range(.Cells(begin_cell, 7), .Cells(begin_cell, 19))
        With tl_row
            .NumberFormat = "mmm-yy"
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Interior.Color = RGB(0, 36, 84) 'Brand Blue
            .Font.Color = vbYellow
        End With
        
        
        With .Cells(begin_cell + 1, 5)
            .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Option 1, Option 2"
            .Interior.Color = RGB(255, 242, 204) 'Gold, accent 4 , lighter 80%
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        End With
        For i = 7 To 18
            .Cells(begin_cell + 1, i).Formula = "=IF(" & .Cells(begin_cell + 1, 5).Address & "=""Option 1"", " & _
            filteredTable.Cells(last_row + 1, i - 1).Address & ", IF(" & .Cells(begin_cell + 1, 5).Address & "=""Option 2"", " & _
            rng.Cells(last_row + 1, i - 1).Address & ", ""Select Option""))"
        Next i
        
    'Generate string text for cells
        
        .Cells(begin_cell, 19).Value = "Total"
        
        Dim headerPositions4 As Variant
        headerPositions4 = Array(1, 3, 14)
        For i = LBound(headerPositions4) To UBound(headerPositions4)
            .Cells(begin_cell + headerPositions4(i), 4).Value = Array( _
                "Pick Option 1 or 2 -------->", "Overall Impact" & vbNewLine & "(Affects all the months)", _
                "Promotional Activities" & vbNewLine & "(Affects only selected months)")(i - LBound(headerPositions4))
            .Cells(begin_cell + headerPositions4(i), 4).Font.Bold = True
        Next i
        
        Dim headerPositions5 As Variant
        headerPositions5 = Array(3, 4, 5, 6, 7, 14, 15, 16, 17)
        For i = LBound(headerPositions5) To UBound(headerPositions5)
            .Cells(begin_cell + headerPositions5(i), 5).Value = Array( _
            "Market Growth", "Price Elasticity", "% doors increase", "% customers increase", "Overall A&CP support", _
            "TV", "Digital support", "KCP 1", "KCP 2")(i - LBound(headerPositions5))
        Next i
        
        Dim headerPositions6 As Variant
        headerPositions6 = Array(1, 13, 24, 25)
        For i = LBound(headerPositions6) To UBound(headerPositions6)
            .Cells(begin_cell + headerPositions6(i), 6).Value = Array("---------------------->", _
                "Sub Total", "Sub Total", "Final Topline Forecast")(i - LBound(headerPositions6))
            .Cells(begin_cell + headerPositions6(i), 6).Font.Bold = True
        Next i
        
        'Set formula for Overall Impact and Total for each row
        For r = 3 To 12
            For c = 7 To 18
                .Cells(begin_cell + r, c).Formula = "=IFERROR(IF(" & .Cells(begin_cell + r, 6).Address & _
                    "="""",""""," & .Cells(begin_cell + r, 6).Address & "*" & .Cells(begin_cell + 1, c).Address & "),"""")"
            Next c
            .Cells(begin_cell + r, 19).Formula = "=IFERROR(SUM(" & .Cells(begin_cell + r, 7).Address & ":" & .Cells(begin_cell + r, 18).Address & "),"""")"
        Next r
        
        'Calculate Sub Total for Overall Impact
        For c = 7 To 19
            .Cells(begin_cell + 13, c).Formula = "=IFERROR(SUM(" & .Cells(begin_cell + 3, c).Address & ":" & _
                .Cells(begin_cell + 12, c).Address & "),"""")"
        Next c

        
        'Create drop-down menu for Promotional Activities
        .Range(.Cells(begin_cell + 14, 6), .Cells(begin_cell + 23, 6)).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Percentage, Absolute Value"
        

        'Calculate Sub Total for Promotional Activities & Final Topline Forecast
        For c = 7 To 18
            'Sub Total
            .Cells(begin_cell + 24, c).Formula = "=IFERROR(SUMIF(" & .Cells(begin_cell + 14, 6).Address & ":" & _
                .Cells(begin_cell + 23, 6).Address & ",""=Absolute Value""," & .Cells(begin_cell + 14, c).Address & ":" & _
                .Cells(begin_cell + 23, c).Address & ") + SUMIF(" & .Cells(begin_cell + 14, 6).Address & ":" & _
                .Cells(begin_cell + 23, 6).Address & ",""Percentage""," & .Cells(begin_cell + 14, c).Address & ":" & _
                .Cells(begin_cell + 23, c).Address & ")*" & .Cells(begin_cell + 1, c).Address & ","""")"

            
            'Final Topline Forecast
            .Cells(begin_cell + 25, c).Formula = "=IFERROR(" & .Cells(begin_cell + 1, c).Address & "+" & _
                        .Cells(begin_cell + 13, c).Address & "+" & .Cells(begin_cell + 24, c).Address & ","""")"
        Next c
           
        
        'Calculate Total for each row of Promotional Activities
        For r = 14 To 23
            .Cells(begin_cell + r, 19).Formula = "=IFERROR(CHOOSE(MATCH(" & .Cells(begin_cell + r, 6).Address & _
                ",{""Percentage"",""Absolute value""},0),SUMPRODUCT(" & .Cells(begin_cell + 1, 7).Address & ":" & _
                .Cells(begin_cell + 1, 18).Address & "," & .Cells(begin_cell + r, 7).Address & ":" & _
                .Cells(begin_cell + r, 18).Address & "),SUM(" & .Cells(begin_cell + r, 7).Address & ":" & _
                .Cells(begin_cell + r, 18).Address & ")," & """Select Percentage or Absolute Value""" & "),"""")"
        Next r
        
        'Calculate Total for Option Row and Sub Total (Overall Impact) Row and Sub Total (Promotional activities) Row and Final Topline Forecast Row
        Dim row_nums As Variant, ind As Variant
        row_nums = Array(1, 13, 24, 25)
        
        For Each ind In row_nums
            .Cells(begin_cell + ind, 19).Formula = "=IFERROR(SUM(" & .Cells(begin_cell + ind, 7).Address & ":" & _
                                                         .Cells(begin_cell + ind, 18).Address & "),"""")"
        Next ind
        
        
    'Format Assumptions Table
        Dim assump_rng() As Range
        ReDim assump_rng(1 To 15)
        Set assump_rng(1) = .Range(.Cells(begin_cell + 3, 5), .Cells(begin_cell + 12, 5))
        Set assump_rng(2) = .Range(.Cells(begin_cell + 3, 6), .Cells(begin_cell + 12, 6))
        Set assump_rng(3) = .Range(.Cells(begin_cell + 14, 5), .Cells(begin_cell + 23, 5))
        Set assump_rng(4) = .Range(.Cells(begin_cell + 14, 6), .Cells(begin_cell + 23, 6))
        Set assump_rng(5) = .Range(.Cells(begin_cell + 3, 7), .Cells(begin_cell + 12, 18))
        Set assump_rng(6) = .Range(.Cells(begin_cell + 14, 7), .Cells(begin_cell + 23, 18))
        Set assump_rng(7) = .Range(.Cells(begin_cell + 25, 7), .Cells(begin_cell + 25, 18)) 'Final Topline Forecast Row
        Set assump_rng(8) = .Range(.Cells(begin_cell + 1, 7), .Cells(begin_cell + 1, 18))
        Set assump_rng(9) = .Range(.Cells(begin_cell + 1, 19), .Cells(begin_cell + 25, 19))
        Set assump_rng(10) = .Cells(begin_cell + 25, 19) 'Total of Topline Forecast Cell
        Set assump_rng(11) = .Cells(begin_cell + 1, 4) 'Pick Option 1 or 2 Cell
        Set assump_rng(12) = .Range(.Cells(begin_cell + 13, 7), .Cells(begin_cell + 13, 18)) 'Sub Total row for Overall Impact
        Set assump_rng(13) = .Range(.Cells(begin_cell + 24, 7), .Cells(begin_cell + 24, 18)) 'Sub Total row for Promotional Activities
        Set assump_rng(14) = .Range(.Cells(begin_cell + 3, 4), .Cells(begin_cell + 12, 4)) 'Overall Impact cell
        Set assump_rng(15) = .Range(.Cells(begin_cell + 14, 4), .Cells(begin_cell + 23, 4)) 'Promotional Activities cell
        
        For i = 1 To 4
            With assump_rng(i)
                .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        Next i
        With assump_rng(2)
            .NumberFormat = "0.00%"
            .Interior.Color = RGB(255, 242, 204) 'Gold, accent 4 , lighter 80%
            .Font.Bold = True
            .VerticalAlignment = xlCenter
        End With
        With assump_rng(4)
            .Interior.Color = RGB(255, 242, 204) 'Gold, accent 4 , lighter 80%
            .Font.Bold = True
            .VerticalAlignment = xlCenter
        End With
        assump_rng(5).NumberFormat = "#,##0"
        For i = 5 To 6
            With assump_rng(i)
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlCenter
                .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
            End With
        Next i
        assump_rng(6).Interior.Color = RGB(255, 242, 204)  'Gold, accent 4 , lighter 80%
        With assump_rng(7)
            .Font.Color = vbRed
            .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .NumberFormat = "#,##0"
        End With
        For i = 8 To 9
            With assump_rng(i)
                .Font.Bold = True
                .NumberFormat = "#,##0"
                .Interior.Color = RGB(142, 170, 219) 'Blue, Accent 1, Lighter 60%
            End With
        Next i
        For i = 10 To 11
            With assump_rng(i)
                .Font.Color = vbRed
                .Interior.Color = RGB(255, 217, 102) 'Gold, Accent 4, Lighter 60%
                .VerticalAlignment = xlCenter
            End With
        Next i
        For i = 12 To 13
            With assump_rng(i)
                .Font.Bold = True
                .Interior.Color = RGB(142, 170, 219) 'Blue, Accent 1, Lighter 60%
                .NumberFormat = "#,##0"
                .VerticalAlignment = xlCenter
            End With
        Next i
        For i = 14 To 15 'Format Overall Impact cell & Promotional Activities cell
            With assump_rng(i)
                .Merge
                .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
                .Interior.Color = RGB(0, 36, 84) 'Brand Blue
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Font.Color = vbWhite
            End With
        Next i
        
    End With
    
'Create a chart for Topline Option 1, 2 and Final Topline Fcst
    
    Dim chartObj As ChartObject
    'Delete any existing chart
    On Error Resume Next
    For Each chartObj In tl_ws.ChartObjects
        chartObj.Delete
    Next chartObj
    On Error GoTo 0
    
    Set chartObj = tl_ws.ChartObjects.Add(left:=tl_ws.Range("A" & begin_cell).left, Top:=tl_ws.Range("A" & begin_cell + 3).Top, Width:=310, Height:=310)
    
    With chartObj.Chart
    ' Set font and font size for axis and legend
        Dim fontName As String: fontName = "Verdana"
        Dim fontSize As Integer: fontSize = 7
        
        
        With .Axes(xlCategory).TickLabels
            .Orientation = 45 ' Rotate labels by 45 degrees
            .Font.Name = fontName
            .Font.Size = fontSize
        End With
        With .Axes(xlValue).TickLabels.Font
            .Name = fontName
            .Size = fontSize
        End With
        With .Legend.Font
            .Name = fontName
            .Size = fontSize
        End With
        
        'Set horizontal axis
        Dim xAxisRange As Range
        Set xAxisRange = rng.Range(rng.Cells(last_row - last_row - last_row - 13, 5), rng.Cells(last_row - last_row - last_row - 13, 16))
        
        With .SeriesCollection.NewSeries
            .Name = "Option 1"
            .values = filteredTable.Range(filteredTable.Cells(last_row - 9, 5), filteredTable.Cells(last_row - 9, 16))
            .ChartType = xlLine
            .AxisGroup = xlPrimary
            .Format.Line.Weight = 2
            .XValues = xAxisRange
        End With
        
        With .SeriesCollection.NewSeries
            .Name = "Option 2"
            .values = rng.Range(rng.Cells(last_row - last_row - 13, 5), rng.Cells(last_row - last_row - 13, 16))
            .ChartType = xlLine
            .AxisGroup = xlPrimary
            .Format.Line.Weight = 2
            .XValues = xAxisRange
        End With
    
        With .SeriesCollection.NewSeries
            .Name = "Topline Final Fcst"
            .values = tl_ws.Range(tl_ws.Cells(begin_cell + 25, 7), tl_ws.Cells(begin_cell + 25, 18))
            .ChartType = xlLine
            .AxisGroup = xlPrimary
            .Format.Line.Weight = 2
            .XValues = xAxisRange
        End With
    
        .Legend.Position = xlLegendPositionBottom
    End With

'Section is Reminder Table for Topline Fcst Worksheet
    With tl_ws
    'Quick User Guide
        
        'With .Range("C5")
            '.value = "Click into this cell for quick user guide"
            '.Font.Bold = True
            '.Font.Underline = True
        'End With
        
        'With .Range("D5").Validation
            '.Delete
            '.Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
           
            '.ShowInput = True
        'End With


        With .Range("B1:C1")
            .Value = Array("Macro last run at:", Format(Now(), "mm/dd/yyyy hh:mm:ss"))
            .Font.Color = vbRed
            .Font.Italic = True
        End With
        
        ' Record the end time
        endTime = Timer
        ' Calculate the running time in seconds
        runningTime = endTime - startTime
        ' Display the running time in cell B2
        With .Range("D1:E1")
            .Value = Array("Macro/VBA running Time:", Format(runningTime, "0.00") & " seconds")
            .Font.Color = vbRed
            .Font.Italic = True
        End With
        
        .Cells.Font.Size = 8
        .Cells.Font.Name = "Verdana"
        .columns.AutoFit
        .Rows.AutoFit
    End With
    
    ResetApplicationSettings
    
    MsgBox "The benchmarked data is extracted." & vbNewLine & _
        "Feel free to overwrite or delete (not remove the whole row) in Option 1 or 2" & vbNewLine & _
        "Then select Option 1 or 2 in the dropdown menu in Cell" & tl_ws.Cells(begin_cell + 1, 5).Address, vbInformation
    
End Sub



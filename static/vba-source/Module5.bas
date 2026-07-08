Attribute VB_Name = "Module5"
Sub Convert_to_LegacyERP()

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Dim erp_ws As Worksheet: Set erp_ws = ThisWorkbook.Worksheets("5. Demand Planning Export")
    
    With erp_ws
        
        Dim legacy_wb As Workbook
        Dim legacy_ws As Worksheet
        Dim headerRow As Range
        Dim lastRow As Long
        Dim lastCol As Long
        Dim filePath As String
        

' Open File Explorer to select a file

        With Application.FileDialog(msoFileDialogFilePicker)
            .Title = "Select legacy ERP export file"
            .Filters.Clear
            .AllowMultiSelect = False
            If .Show = -1 Then ' If file is selected
                filePath = .selectedItems(1)
            Else
                Exit Sub ' Exit if no file is selected
            End If
        End With


        ' Open the selected file
        Set legacy_wb = Workbooks.Open(filePath)
        Set legacy_ws = legacy_wb.Sheets(1)

        'Define NPD's rows count exclude 2 first row
        Dim npd_row As Long: npd_row = .Range("D6").End(xlDown).Row - 6

        Dim datesArray() As Variant
        Dim outputArray() As String
        Dim i As Integer
    
        ' Assign the range to an array
        datesArray = .Range("E6:S6").Value
    
        ' Initialize the output array
        ReDim outputArray(1 To UBound(datesArray, 2))
    
        ' Convert each date to text and store it in the output array
        For i = 1 To UBound(datesArray, 2)
            If IsDate(datesArray(1, i)) Then
                outputArray(i) = Format(datesArray(1, i), "mmm - yy")
            End If
        Next i
        
        For i = LBound(outputArray) To UBound(outputArray)
            ' Loop through each column in the first row of legacy_ws
            For c = 1 To legacy_ws.Cells(1, legacy_ws.columns.count).End(xlToLeft).Column
                If left(outputArray(i), 3) = Mid(legacy_ws.Cells(1, c).Value, 2, 3) And _
                   right(outputArray(i), 2) = right(legacy_ws.Cells(1, c).Value, 2) Then
                    ' If the condition matches, iterate over rows in the 5th column of erp_ws
                    For r = 2 To legacy_ws.Cells(.Rows.count, 5).End(xlUp).Row
                        For m = 7 To 6 + npd_row
                            If legacy_ws.Cells(r, 5).Value = .Cells(m, 4).Value Then
                                ' If a match is found, set .Cells(r,c).Value = .Cells(r+4+npd_row, i).Value
                                Debug.Print .Cells(m, 4)
                                legacy_ws.Cells(r, c).Value = .Cells(m, i + 4).Value
                            End If
                        Next m
                    Next r
                End If
            Next c
        Next i

        ' Close the legacy system's workbook without saving changes
        'legacy_wb.Close SaveChanges:=True

     End With
     
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    
End Sub

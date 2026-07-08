Attribute VB_Name = "Module3"
Option Explicit 'This syntax is for VBA to make sure all the variables are defined before using so that VBA would run faster
Sub Convert_to_DemandPlanner()
        
    'Turn off all the Excel Updating to make the script run faster
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Dim npd_ws As Worksheet, erp_ws As Worksheet
    With ThisWorkbook
        Set npd_ws = .Worksheets("1. NPD Info")
        Set erp_ws = .Worksheets("5. Demand Planning Export")
    End With
    
    Dim npd_row As Long: npd_row = npd_ws.Range("C6").End(xlDown).Row

    With erp_ws
        Dim loc As Range
        Set loc = .Range(.Cells(7, 2), .Cells(npd_row, 2))
        
        Dim dmdGroup As Range
        Set dmdGroup = .Range(.Cells(7, 3), .Cells(npd_row, 3))
        
        Dim FcstID As String
        FcstID = .Cells(4, 3).Value
        
        Dim IRC As Range
        Set IRC = .Range(.Cells(7, 4), .Cells(npd_row, 4))
        
        Dim qty As Range
        Set qty = .Range(.Cells(7, 5), .Cells(npd_row, 19))
        
        Dim start_date As Range
        Set start_date = .Range(.Cells(6, 5), .Cells(6, 19))
        
        
        If IsEmpty(.Cells(2, 3).Value) Then
             MsgBox "Error: Please Pick Forecast Table submitted to DemandPlanner in Cell C2", vbCritical, "Error"
             Exit Sub
        End If
        If IsEmpty(.Cells(4, 3).Value) Then
            MsgBox "Error: Please manually input the Forecast ID in Cell C4", vbCritical, "Error"
             Exit Sub
        End If
        If Application.WorksheetFunction.CountA(loc) <> npd_row - 6 Then
            MsgBox "Error: Please input all the Location", vbCritical, "Error"
            Exit Sub
        End If
        If Application.WorksheetFunction.CountA(dmdGroup) <> npd_row - 6 Then
            MsgBox "Error: Please input all the Dmd Group", vbCritical, "Error"
            Exit Sub
        End If
    End With
    
    Dim wbNew As Workbook: Set wbNew = Workbooks.Add
    With wbNew.Sheets(1)
        Dim loc_arr() As Variant
        loc_arr = loc.Value
        
        Dim dmdGroup_arr() As Variant
        dmdGroup_arr = dmdGroup.Value
        
        Dim IRC_arr() As Variant
        IRC_arr = IRC.Value
        
        Dim qty_arr() As Variant
        qty_arr = qty.Value

        
        Dim start_date_arr() As Variant
        start_date_arr = start_date.Value
        
        Dim erp_row As Long
        erp_row = 1
        
        Dim i As Long
        For i = 1 To 15
            Dim j As Integer
            For j = LBound(loc_arr) To UBound(loc_arr)
                If qty_arr(j, i) > 0 And Not IsEmpty(qty_arr(j, i)) Then 'Set condition to only take into quantity amount > 0
                    .Cells(erp_row, 1).Value = loc_arr(j, 1)
                    .Cells(erp_row, 2).Value = dmdGroup_arr(j, 1)
                    .Cells(erp_row, 3).Value = IRC_arr(j, 1)
                    .Cells(erp_row, 4).Value = "5"
                    .Cells(erp_row, 5).Value = FcstID
                    .Cells(erp_row, 6).Value = "SALES-STD"
                    .Cells(erp_row, 7).Value = Round(qty_arr(j, i), 0)
                    .Cells(erp_row, 8).Value = start_date_arr(1, i)
                    .Cells(erp_row, 8).NumberFormat = "m/d/yyyy"
                    
                    Select Case Month(.Cells(erp_row, 8).Value)
                        Case 1, 3, 5, 7, 8, 10, 12
                            .Cells(erp_row, 9).Value = "31D"
                        Case 4, 6, 9, 11
                            .Cells(erp_row, 9).Value = "30D"
                        Case 2
                            ' Check if the year is a leap year
                            If year(.Cells(erp_row, 8).Value) Mod 4 = 0 And (year(.Cells(erp_row, 8).Value) Mod 100 <> 0 Or year(.Cells(erp_row, 8).Value) Mod 400 = 0) Then
                                .Cells(erp_row, 9).Value = "29D"
                            Else
                                .Cells(erp_row, 9).Value = "28D"
                            End If
                    End Select
                    
                    erp_row = erp_row + 1
                End If
            Next j
        Next i
        
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
    End With
    
    'Syntax to pop up Disk D: and get the file as csv
    Dim fileSaveName As Variant
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:="C:\Forecast_Export.csv", FileFilter:="CSV Files (*.csv), *.csv")
    If fileSaveName <> False Then
        wbNew.SaveAs Filename:=fileSaveName, FileFormat:=xlCSV
    End If
    
    With erp_ws.Range("O1:P1")
        .Value = Array("Macro last run at:", Format(Now(), "mm/dd/yyyy hh:mm:ss"))
        .Font.Color = vbRed
        .Font.Italic = True
    End With
    
    'Turn on screen updating again
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With

    
End Sub

    
    
   



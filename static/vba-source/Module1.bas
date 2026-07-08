Attribute VB_Name = "Module1"
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
Sub CleanAllCells(ws As Worksheet, rng As Range)
    SetUpApplicationSettings
    
    Dim cell As Range
    For Each cell In rng
        ' Replace line breaks with an empty string
        cell.Value = Replace(cell.Value, vbCrLf, "")
        cell.Value = Replace(cell.Value, vbCr, "")
        cell.Value = Replace(cell.Value, vbLf, "")
        ' Trim leading and trailing spaces
        cell.Value = Trim(cell.Value)
    Next cell
    
    ResetApplicationSettings
End Sub

Sub Update_database()

    SetUpApplicationSettings
    
    Dim startTime As Double
    Dim endTime As Double
    Dim runningTime As Double
    
    ' Record the start time
    startTime = Timer
    
    Dim data_ws As Worksheet: Set data_ws = ThisWorkbook.Worksheets("0. Database")
       
    With data_ws
        
        Dim src_wb As Workbook
        Dim src_ws As Worksheet
        Dim headerRow As Range
        Dim lastRow As Long
        Dim lastCol As Long
        Dim filePath As String
        
        ' Open File Explorer to select a file
        With Application.FileDialog(msoFileDialogFilePicker)
            .Title = "Select market-data export file"
            .Filters.Clear
            .AllowMultiSelect = False
             If .Show = -1 Then ' If file is selected
                filePath = .selectedItems(1)
            Else
                Exit Sub ' Exit if no file is selected
            End If
        End With
        
        ' Open the selected file
        Set src_wb = Workbooks.Open(filePath)
        Set src_ws = src_wb.Sheets(1) ' Assuming data is in the first sheet
        
        ' Find the header row with GTIN, IRC, IRC Description
        Dim r As Long, headerRowNumber As Long
        For r = 1 To 50
            Set headerRow = src_ws.Rows(r).Find(What:="GTIN", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=True)
            If Not headerRow Is Nothing Then
                Exit For ' Exit the loop if the header row is found
            End If
        Next r
        
        If headerRow Is Nothing Then
            MsgBox "Header row with GTIN, IRC, IRC Description, etc is not found in the selected file.", vbExclamation
            src_wb.Close SaveChanges:=False
            Exit Sub
        End If
        
        'Check if the target market is selected when downloading the source data
        'Dim MarketFound As Range
        'Set MarketFound = src_ws.Rows(headerRow.Row).Find("Market", LookIn:=xlValues, LookAt:=xlPart)
        
        'If MarketFound Is Nothing Then
            'MsgBox "Error: Total Market level is not selected in Figures Details when downloading the source data", vbCritical, "Error"
            'Exit Sub
        'End If
        
        ' Find the last row and last column with data in the selected file
        lastRow = src_ws.Cells(src_ws.Rows.count, headerRow.Column).End(xlUp).Row
        lastCol = src_ws.Cells(headerRow.Row, src_ws.columns.count).End(xlToLeft).Column
        
        'Clear the existing data in Database workbook
        .Cells.Clear
        
        ' Copy the data from the selected file to this workbook
        Dim destinationRange As Range
        Set destinationRange = .Range("B1").Resize(lastRow - 1, lastCol - 1)
        
        ' Copy values and format from src_ws range to destinationRange
        src_ws.Range(headerRow, src_ws.Cells(lastRow - 1, lastCol - 1)).Copy Destination:=destinationRange
        src_ws.Range(src_ws.Cells(1, 1), src_ws.Cells(headerRow.Row - 1, lastCol - 1)).Copy Destination:=.Cells(1, lastCol + 3)
        
        ' After copying the data:
        CleanAllCells data_ws, data_ws.Range(data_ws.Cells(1, 2), data_ws.Cells(1, lastCol))

        ' Clear clipboard
        Application.CutCopyMode = False
    
        ' Close the selected file without saving
        src_wb.Close SaveChanges:=False
           

        ' Identify the first column if the first three characters are 'Jul'
        Dim firstCol As Long, c As Long
        firstCol = 0 ' Initialize to 0
        For c = 2 To lastCol
                If left(.Cells(1, c).Value, 3) = "Jul" Then
                firstCol = c ' This is the column number
                Exit For
            End If
        Next c
        
        ' If the loop completes without finding a "Jul" header, inform the user
        If firstCol = 0 Then
            MsgBox "Cannot find any value of the header row start with 'Jul'", vbExclamation
            Exit Sub
        End If

        
        'Read the market data into an array
        Dim data() As Variant: data = .Range(.Cells(1, firstCol), .Cells(1, lastCol)).Value
    
        'Check if month is Jan-Jun or Jul-Dec and convert the text to a date accordingly
        Dim year As Integer
        If left(data(1, 1), 3) = "Jan" Or left(data(1, 1), 3) = "Feb" Or left(data(1, 1), 3) = "Mar" Or left(data(1, 1), 3) = "Apr" Or left(data(1, 1), 3) = "May" Or left(data(1, 1), 3) = "Jun" Then
            year = 2000
        Else
            year = 1999
        End If
        data(1, 1) = DateSerial(right(data(1, 1), 2) + year, Month("1/" & left(data(1, 1), 3)), 1)
        
        'Convert the date to the desired format
        Dim dateFormat As String
        dateFormat = "mmm-yy"
        
        'Loop through the headers from the second date column to the last column and increment the dates by 1 month
        Dim i As Long
        For i = 2 To UBound(data, 2)
            data(1, i) = Application.WorksheetFunction.EDate(data(1, i - 1), 1)
        Next i
    
        'Write the data back to the worksheet
        .Range(.Cells(1, firstCol), .Cells(1, lastCol)).Value = data
        .Range(.Cells(1, firstCol), .Cells(1, lastCol)).NumberFormat = dateFormat

        With .Range(.Cells(18, lastCol + 3), .Cells(18, lastCol + 4))
            .Value = Array("Macro last run at:", Format(Now(), "mm/dd/yyyy hh:mm:ss"))
            .Font.Color = vbRed
            .Font.Italic = True
        End With
        
        ' Record the end time
        endTime = Timer
        ' Calculate the running time in seconds
        runningTime = endTime - startTime
        ' Display the running time in cell B2
        With .Range(.Cells(19, lastCol + 3), .Cells(19, lastCol + 4))
            .Value = Array("Macro/VBA running Time:", Format(runningTime, "0.00") & " seconds")
            .Font.Color = vbRed
            .Font.Italic = True
        End With
        
        'Format
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
    End With
    
    ResetApplicationSettings
    
    MsgBox "Market data has been copied from the selected file to this workbook." & vbNewLine & _
    "Please move to 1. NPD Info Worksheet to manually input the information of new products", vbInformation
    
End Sub




Attribute VB_Name = "Sheet4"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Activate()
    
    ' Store the clipboard contents
    Dim DataObj As Object
    Set DataObj = GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    DataObj.GetFromClipboard
    Dim clipboardContent As String
    On Error Resume Next ' Ignore errors if clipboard is empty
    clipboardContent = DataObj.GetText(1)
    
    
    ' Set common formatting properties
    
    With Me.Range("B:C")
        .NumberFormat = "#0"
    End With
    
    With Me.Range("H:H")
        .NumberFormat = "#.#0"
    End With
    
    With Me.Range("B:I")
        .Font.Name = "Verdana"
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With
    With Range("I5").Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
        .InputTitle = "Enter Date"
        .InputMessage = "Date format of d/m/yyyy or m/d/yyyy"
        .ShowInput = True
    End With
    ' Format cell I5
    With Range("I5")
        .NumberFormat = "mmm-yy"
        .Interior.Color = RGB(255, 242, 204) ' Gold, accent 4, lighter 80%
        .Font.Color = vbRed
        .Font.Bold = True
    End With
    
    
    ' Format merged cell B5:H5
    With Range("B5:H5")
        .Merge
        .Value = "NPD Information"
        .Interior.Color = RGB(0, 36, 84) ' Brand Blue
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' Format cell range B6:I6
    With Range("B6:I6")
        .Value = Array("GTIN", "IRC", "IRC Description", "Shade/ML", "Market Category", "Article Type", "List Price (GSP/NSP)", "Concatenate")
        .Interior.Color = RGB(0, 36, 84) ' Brand Blue
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' Format cell I4
    With Range("I4")
        .Value = "Put ICW date in the cell below"
        .Font.Color = vbRed
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With Range("B7:H300")
        .Interior.Color = RGB(255, 242, 204) ' Gold, accent 4, lighter 80%
    End With
    
    With Range("B2")
        .Value = "Click into this cell for quick user guide"
        .Font.Bold = True
        .Font.Underline = True
        .Validation.Delete
        .Validation.Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
        .Validation.InputMessage = "Please input the info of NPD in cells with Yellow color" & vbNewLine & _
                                   "Please input numbers only for List Price"
        .Validation.ShowInput = True
    End With
    
    ' Restore clipboard contents
    If clipboardContent <> "" Then
        DataObj.SetText clipboardContent
        DataObj.PutInClipboard
    End If
    
    columns.AutoFit
    Rows.AutoFit
End Sub


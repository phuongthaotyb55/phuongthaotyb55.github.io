Attribute VB_Name = "Sheet7"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "BnchmrkButton2, 8, 0, MSForms, CommandButton"
Private Sub BnchmrkButton2_Click()
    frmBnchmrk2.Show
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("A:A")) Is Nothing Then
        columns("A:A").ColumnWidth = 30
    End If
End Sub


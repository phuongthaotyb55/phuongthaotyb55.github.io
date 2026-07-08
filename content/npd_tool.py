TITLE = "NPD Demand Forecasting Automation"

TAGS = ["Excel", "VBA", "Macros", "Demand Planning", "Automation"]

INTRO = (
    "During my Global Demand Planning internship, new-product demand forecasts were being "
    "built by hand in Excel for every launch: pulling the latest market-data export, splitting "
    "a topline number down to SKU level, weighting it by customer, and reformatting it for "
    "upload to the demand-planning system. I built a VBA tool that turned that into a "
    "handful of one-click macros, and it was adopted by multiple business units."
)

CONFIDENTIALITY_NOTE = (
    "The real workbook holds no confidential data itself — it's a blank template the macros "
    "populate — but the automation is wired to my former employer's specific systems (an ERP "
    "demand-planning platform, a market-data export format, internal naming conventions). "
    "Everything shown and downloadable here is a version I rebuilt with those references "
    "replaced by generic equivalents and filled with synthetic sample data, so the logic and "
    "engineering are real and unchanged — only the employer-specific parts are gone."
)

ARCHITECTURE = [
    {
        "sheet": "0. Database",
        "role": "Raw market-data landing zone. One macro opens a file picker, pulls in the "
        "latest export, and normalises its date headers regardless of which half of the year "
        "it starts in.",
    },
    {
        "sheet": "1. NPD Info",
        "role": "The only sheet a planner types into: new product codes, descriptions, list "
        "prices, and the in-consumer-window date for the launch.",
    },
    {
        "sheet": "2. Topline Fcst / 3. Split by SKUs",
        "role": "Working sheets — a reconciled topline number and a per-SKU weighting "
        "extrapolated from historical splits.",
    },
    {
        "sheet": "4. Final Fcst",
        "role": "One click (Generate Final Forecast) builds a 12-month volume and value "
        "forecast per SKU, then a second click (Update Customer Info) re-splits it across "
        "however many customers are entered, each with their own weight and launch date.",
    },
    {
        "sheet": "5. Demand Planning Export",
        "role": "Reformats the final forecast into the row-per-month-per-location shape the "
        "demand-planning system expects, and exports it as CSV.",
    },
]

CODE_HIGHLIGHTS = [
    {
        "title": "Column lookup by header name, not position",
        "note": (
            "The benchmarking screen reads from whatever columns the market-data export "
            "happens to contain that week — so instead of hardcoding column numbers, it finds "
            "them by header text. Any export format change and the tool doesn't have to."
        ),
        "code": """Private Function FindColumnNumber(sheet As Worksheet, headerName As String) As Long
    Dim col As Range
    For Each col In sheet.Rows(1).Cells
        If col.Value = headerName Then
            FindColumnNumber = col.Column
            Exit Function
        End If
    Next col
    FindColumnNumber = 0
End Function""",
    },
    {
        "title": "Custom QuickSort for the cascading filter dropdowns",
        "note": (
            "The benchmark form's Region → Sub-Region → Market dropdowns need their "
            "unique values sorted before display. Rather than round-tripping through a "
            "worksheet, it's a plain in-memory QuickSort over the values pulled from a "
            "Scripting.Dictionary."
        ),
        "code": """Private Sub quickSort(Arr As Variant, ByVal left As Long, ByVal right As Long)
    Dim i As Long, j As Long, pivot As Variant, temp As Variant
    i = left: j = right
    pivot = Arr((left + right) \\ 2)
    While (i <= j)
        While (Arr(i) < pivot And i < right): i = i + 1: Wend
        While (pivot < Arr(j) And j > left): j = j - 1: Wend
        If (i <= j) Then
            temp = Arr(i): Arr(i) = Arr(j): Arr(j) = temp
            i = i + 1: j = j - 1
        End If
    Wend
    If (left < j) Then Call quickSort(Arr, left, j)
    If (i < right) Then Call quickSort(Arr, i, right)
End Sub""",
    },
    {
        "title": "One-click forecast generation",
        "note": (
            "Generate_FinalFcst() rebuilds the entire per-SKU, 12-month forecast table from "
            "scratch — formulas, formatting, and all — driven entirely by how many "
            "product rows are currently in 1. NPD Info. Add a 6th product and the sheet "
            "just grows."
        ),
        "code": """With npd_ws
    Dim npd_row As Long: npd_row = .Range("C6").End(xlDown).Row - 5
End With

With fcst_ws
    .Cells.Clear
    .CheckBoxes.Delete
    For c = 2 To 7
        For r = 5 To (npd_row + 4)
            .Cells(r, c).Formula = "=" & npd_ws.Cells(r + 1, c).Address(External:=True)
        Next r
    Next c
    ' ...generates the 12-month date row, per-SKU weight column,
    ' and the volume/value formulas beneath it entirely from npd_row
End With""",
    },
]

DOWNLOADS = {
    "workbook": "downloads/npd_forecasting_sample.xlsx",
    "source": "downloads/npd_forecasting_vba_source.zip",
}

DEMO_PRODUCTS = [
    {"irc": "LUM-HG-050", "desc": "Hydra Glow Serum 50ml", "price": 24.99, "weight": 25},
    {"irc": "LUM-HG-030", "desc": "Hydra Glow Serum 30ml", "price": 16.99, "weight": 20},
    {"irc": "LUM-VM-201", "desc": "Velvet Matte Lipstick — Rosewood", "price": 19.99, "weight": 20},
    {"irc": "LUM-VM-202", "desc": "Velvet Matte Lipstick — Terracotta", "price": 19.99, "weight": 20},
    {"irc": "LUM-BR-100", "desc": "Bloom Radiance Highlighter", "price": 22.50, "weight": 15},
]

DEMO_CUSTOMERS = [
    {"name": "Retailer Alpha", "weight": 40, "icw_offset": 0},
    {"name": "Retailer Beta", "weight": 35, "icw_offset": 1},
    {"name": "Retailer Gamma", "weight": 25, "icw_offset": 2},
]

DEMO_TOTAL_ANNUAL_VOLUME = 30000

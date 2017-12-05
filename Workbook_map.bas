Attribute VB_Name = "Workbook_map"
Option Explicit

Sub AUX_clean_shapes()
''AUX delete all shapes
Dim shp As Shape
For Each shp In ActiveSheet.Shapes
   shp.Delete
   'shp.Select
Next
End Sub

Sub create_wb_map()
Sheets_to_boxes
add_dependency_arrows_to_boxes
End Sub

Sub Sheets_to_boxes()
''Creates a text box for every worksheet of the current Workbook in the activesheet

Dim n As Integer, oldStatusBar
Dim i As Integer, t As Integer
Dim l As Integer
Dim Colour As Long, prevColour As Long
Dim ThemeColour As Long, prevThemeColour As Long
Dim boxw As Integer, maxw As Integer
Dim xsht As Worksheet

i = 0
Const TxtSize = 10
l = (Cells(1, 1).width * 1.5)

n = 1
oldStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = True

Set xsht = ActiveSheet

For Each xsht In ActiveWorkbook.Sheets
    
    Application.StatusBar = n & "/" & ActiveWorkbook.Sheets.Count & "   " & xsht.Name
    
    Colour = xsht.Tab.color
    'TODO: Fix default Colour
    If Colour = 0 Then Colour = 15132390 '= RGB(230, 230, 230) 'default
    'Every Colour change advances one column and resets the row counter
    
    ThemeColour = xsht.Tab.ThemeColor
    If (ThemeColour <> prevThemeColour) Then
            i = 0
            l = l + (maxw + TxtSize * 3) 'increasing variable with
            maxw = 0
    End If
    
    prevColour = Colour
    prevThemeColour = ThemeColour
    
    t = (Cells(1, 1).Height * 4.5) + (TxtSize * 2 + TxtSize) * i

    boxw = insert_box(xsht.Name, l, t, _
                        Colour:=Colour, _
                        TintAndShade:=xsht.Tab.TintAndShade)
                        
    If boxw > maxw Then maxw = boxw
    
    i = i + 1
    n = n + 1
    
Next

Application.StatusBar = False
Application.DisplayStatusBar = oldStatusBar

End Sub


Function insert_box(Optional txt As String = "TEXT", _
               Optional Left As Integer = 1, _
               Optional Top As Integer = 1, _
               Optional StrSize As Integer = 11, _
               Optional Bold As Boolean = False, _
               Optional Colour As Long = 15132390, _
               Optional TintAndShade As Integer = 0)
    ' 15132390 = RGB(230, 230, 230)
    ' 16777215 = RGB(255, 255, 255)
    
    '' Returns the final box width.
    
    Dim i As Integer
    Dim RGBarr
    Dim txtColour As Long
    Dim shp As Shape
    
    'Irrelevant, as this will be overwritten by AutoSize
    Const w = 10
    Const h = 10

    Set shp = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Left, Top, w, h)
    
    shp.Name = txt  'This allows easy identification later (AVOID DUPS)
    ActiveSheet.Hyperlinks.Add Anchor:=shp, _
                               Address:="", _
                               SubAddress:="'" & txt & "'!A1"
    shp.TextFrame2.TextRange.Characters.Text = txt
    
    shp.Placement = xlFreeFloating
    With shp.TextFrame2
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeShapeToFitText
    End With

    RGBarr = ColourToRGB(Colour)
    'Empirical approximation to plane RGB Colourspace used by Excel func
    If (RGBarr(0) * 20132) _
        + (RGBarr(1) * 64005) _
        + (RGBarr(2) * 6630) <= 11675430 Then
        txtColour = RGB(255, 255, 255) 'white
    Else
        txtColour = RGB(0, 0, 0) 'black
    End If

    With shp.TextFrame2.TextRange.Font
        .Size = StrSize
        .Fill.ForeColor.RGB = txtColour
        If Bold Then .Bold = msoTrue
    End With
    
    With shp.Fill
        .ForeColor.RGB = Colour
        '.ForeColor.TintAndShade = TintAndShade
    End With

    ''Line custom black:
    With shp.Line
        .ForeColor.RGB = RGB(0, 0, 0)
        .Weight = 1
    End With
    
    insert_box = shp.width
    
End Function


Private Function ColourToRGB(Colour As Long) As Variant
'original: https://www.office-forums.com/threads/need-inverse-of-rgb-r-g-b.1886634/
Dim strColour As String
Dim hexColour As String
Dim nColour As Long
Dim nR As Long, nB As Long, nG As Long
Dim RGB(2) As Integer

hexColour = Hex(Colour) ' convert decimal number to hex string
While Len(hexColour) < 6 ' pad on left to 6 hex digits
hexColour = "0" & hexColour
Wend

nB = CLng("&H" & Mid(hexColour, 1, 2))
nG = CLng("&H" & Mid(hexColour, 3, 2))
nR = CLng("&H" & Mid(hexColour, 5, 2))

RGB(0) = nR
RGB(1) = nG
RGB(2) = nB

ColourToRGB = RGB
End Function

Sub AUX_clean_dependecy_arrows()
''AUX delete all connectors
Dim shp As Shape
For Each shp In ActiveSheet.Shapes
    If shp.Connector = msoTrue Then shp.Delete
Next
End Sub

Sub AUX_change_dependecy_arrows()
''AUX Apply a change to all connectors
Dim shp As Shape
For Each shp In ActiveSheet.Shapes
    If shp.Connector = msoTrue Then
        'shp.ConnectorFormat.Type = msoConnectorStraight
        shp.ConnectorFormat.Type = msoConnectorCurve
        'shp.ConnectorFormat.Type = msoConnectorElbow
   End If
Next
End Sub

Private Sub add_dependency_arrows_to_boxes()
'' Connects the boxes created by macro "Sheets_to_boxes", based on each tabs' formulae
'' Provides a visualization of the workbook structure

Dim t As Integer, oldStatusBar
Dim tshts
Dim trng As Range
Dim ishp As Shape, fshp As Shape
Dim sht As Worksheet
Dim ishpn, n As Long, thickness As Double
Dim d As Object, x

t = 1
oldStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = True

Set tshts = ActiveWorkbook.Worksheets
'Set tshts = ActiveWindow.SelectedSheets

On Error Resume Next
For Each sht In tshts
    
    'Debug.Print sht.Name
    Application.StatusBar = t & "/" & tshts.Count & "   " & sht.Name
    
    Set fshp = ActiveSheet.Shapes(sht.Name)
    If fshp Is Nothing Then GoTo Next_fshp
    Set trng = sht.Cells.SpecialCells(xlCellTypeFormulas)
    If trng Is Nothing Then GoTo Next_fshp
    
    Set d = precedent_sheetnames_count(trng)
    For Each ishpn In d
        
        If ishpn = sht.Name Then GoTo Next_ishp
        Set ishp = ActiveSheet.Shapes(ishpn)
        If ishp Is Nothing Then GoTo Next_ishp
        
        n = d(ishpn)
        thickness = Log(n) + 0.25 'in vba log = ln
        If thickness >= 1 Then insert_connector ishp, fshp, thickness
        
Next_ishp:
        Set ishp = Nothing
        
    Next
    
    'sht.Parent.Save
    
Next_fshp:
    Set d = Nothing
    Set sht = Nothing
    Set fshp = Nothing
    Set trng = Nothing
    Set ishpn = Nothing
    Set ishp = Nothing
    t = t + 1
Next

On Error GoTo 0

Application.StatusBar = False
Application.DisplayStatusBar = oldStatusBar
End Sub


Private Sub insert_connector(ishp As Shape, fshp As Shape, Optional thickness As Double = 1)
    
    Dim shp As Shape, Name As String
    Dim Colour As Long, RGBarr
    Dim iConnectPt As Integer, fConnectPt As Integer
    
    ''points
    Const l = 0
    Const t = 0
    Const w = 10
    Const h = 10
    
    Name = ishp.Name & " to " & fshp.Name
    If shape_exists(Name) Then Exit Sub 'do not overwrite
    
    'Set shp = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, l, t, w, h) 'straight
    'Set shp = ActiveSheet.Shapes.AddConnector(msoConnectorElbow, l, t, w, h) ' elbows
    Set shp = ActiveSheet.Shapes.AddConnector(msoConnectorCurve, l, t, w, h) 'straight
    
    shp.Name = Name
    
    shp.Placement = xlFreeFloating
    
    shp.Line.EndArrowheadStyle = msoArrowheadTriangle
    
    iConnectPt = 4 'boxes' right side
    fConnectPt = 2 'boxes' left side
    
    If ishp.Left + ishp.width > fshp.Left Then fConnectPt = 4
    
    With shp.ConnectorFormat
        .BeginConnect ishp, iConnectPt
        .EndConnect fshp, fConnectPt
    End With
    'shp.RerouteConnections 'shortest path might change previous points
    shp.Line.Weight = thickness
    
    RGBarr = ColourToRGB(Colour)
    'Empirical approximation to plane RGB Colourspace used by Excel func
    'https://stackoverflow.com/a/47208623/2802352
    If (RGBarr(0) * 225) _
        + (RGBarr(1) * 225) _
        + (RGBarr(2) * 225) <= 168750 Then
        Colour = ishp.Fill.ForeColor.RGB
    End If
    shp.Line.ForeColor.RGB = Colour
    
    shp.ZOrder msoSendToBack
    
End Sub

Private Function shape_exists(Name As String)
''Reyurns True if a shape exists, False otherwise
Dim shp As Shape
On Error Resume Next
Set shp = ActiveSheet.Shapes(Name)
shape_exists = Not shp Is Nothing
End Function


Sub AUX_print_precedent_sheetnames_count(rng As Range)
''AUX
''Prints the name of every sheet used in sheet_name cells' formulas
Dim d As Object, x
On Error GoTo FinishThis
Set d = precedent_sheetnames_count(rng)
For Each x In d
    Debug.Print x, d(x)
Next
FinishThis:
Set d = Nothing
End Sub


Private Function precedent_sheetnames_count(trng As Range)
''Returns a dictionary of {sheet_name: count_cells_using_sheet_name}
''of the sheets used in rng's formulas.

Dim rng As Range
Dim shtns, shtn
Dim d As Object
Set d = CreateObject("Scripting.Dictionary")

On Error GoTo FinishThis
Set trng = trng.Cells.SpecialCells(xlCellTypeFormulas)

For Each rng In trng
    shtns = precedent_sheetnames(rng)
    For Each shtn In shtns
        d(shtn) = d(shtn) + 1
    Next
Next

Set precedent_sheetnames_count = d

FinishThis:
If Not d Is Nothing Then Set d = Nothing
End Function


Private Function precedent_sheetnames(rng As Range) As Variant
'Returns an array with the unique names of all the sheets used in rng.Formula

Dim RE As Object, matches, m, sm
Set RE = CreateObject("vbscript.regexp")
Dim d As Object
Set d = CreateObject("Scripting.Dictionary")

''Either formula with "'"s (easy), or without "'"s
RE.pattern = "'(.+?)'!|\b([^/\*-+ =&<>\[\]""\(\)]+)!"
''This approach is weak, but sufficient for this purpose

RE.Global = True
Set matches = RE.Execute(rng.Formula)

For Each m In matches
    For Each sm In m.submatches
        d(sm) = 1
    Next
Next

precedent_sheetnames = d.keys()
Set RE = Nothing
Set d = Nothing
End Function



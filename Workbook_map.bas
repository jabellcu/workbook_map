Attribute VB_Name = "Workbook_map"
Option Explicit

Sub AUX_clean_shapes()
    '' AUX delete all shapes
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        shp.Delete
        'shp.Select
    Next
End Sub

Sub create_wb_map()
    
    Sheets_to_boxes
    add_dependency_arrows_to_boxes
    
    '' Alternative for faster calculation of links between sheets (relevant for
    '' larger workbooks):
    ''  1) Run output_formulae (in this module) on the target workbook;
    ''  2) Run output_names (in this module) on the target workbook;
    ''  3) Run process_formulas.ipynb (jupyter notebook - requires python 3.7+).
    ''     Ammend names and paths as necessary. This will produce the file:
    ''     "Workbook_map_EXAMPLE_formulas_count.csv"
    '' Then use the following line of code instead of the previous one:
    
    'add_dependency_arrows_to_boxes precedents_filepath:="Workbook_map_EXAMPLE_formulas_count.csv"
    
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
    
        Application.StatusBar = n & "/" & ActiveWorkbook.Sheets.Count & "   " & xsht.name
    
        Colour = xsht.Tab.color
        'TODO: Fix default Colour
        If Colour = 0 Then Colour = 15132390     '= RGB(230, 230, 230) 'default
        'Every Colour change advances one column and resets the row counter
    
        ThemeColour = xsht.Tab.ThemeColor
        If (ThemeColour <> prevThemeColour) Then
            i = 0
            l = l + (maxw + TxtSize * 3)         'increasing variable with
            maxw = 0
        End If
    
        prevColour = Colour
        prevThemeColour = ThemeColour
    
        t = (Cells(1, 1).Height * 4.5) + (TxtSize * 2 + TxtSize) * i

        boxw = insert_box(xsht.name, l, t, _
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
    
    shp.name = txt                               'This allows easy identification later (AVOID DUPS)
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
        txtColour = RGB(255, 255, 255)           'white
    Else
        txtColour = RGB(0, 0, 0)                 'black
    End If

    With shp.TextFrame2.TextRange.Font
        .size = StrSize
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

    hexColour = Hex(Colour)                      ' convert decimal number to hex string
    While Len(hexColour) < 6                     ' pad on left to 6 hex digits
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

Sub Select_Shapes()

    Dim i As Long, condition As Boolean
    Dim shp As Shape, ishp As Long
    Dim shpn_arr() As String

    For i = 0 To ActiveSheet.Shapes.Count - 1
        ishp = i + 1
        Set shp = ActiveSheet.Shapes(ishp)
    
        'condition = (shp.Connector = msoTrue) 'arrows
        'condition = (shp.Connector = msoFalse) 'boxes
        'condition = (shp.Connector = msoFalse) And (shp.Fill.ForeColor <> 15132390) 'Non-grey boxes
        'condition = (shp.Connector = msoFalse) And (shp.Fill.ForeColor = 192)
        condition = (shp.Connector = msoFalse) And (shp.Left > 900)
        'condition = (shp.Connector = msoTrue) And (shp.Name Like "*Control*")
    
        If condition Then
            ReDim Preserve shpn_arr(i)
            shpn_arr(i) = shp.name
        End If
    Next
    ActiveSheet.Shapes.Range(shpn_arr).Select
End Sub

Sub linkfy_boxes()

    Dim shp As Shape, txt As String

    For Each shp In ActiveSheet.Shapes
    
        If Not (shp.Connector = msoFalse) Then GoTo NextIteration
        txt = shp.TextFrame2.TextRange.Characters.Text
        ActiveSheet.Hyperlinks.Add Anchor:=shp, _
                                   Address:="", _
                                   SubAddress:="'" & txt & "'!A1"

NextIteration:
    Next

End Sub

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
    'For Each shp In ActiveSheet.Shapes
    For Each shp In Selection.ShapeRange
        If shp.Connector = msoTrue Then
            'shp.ConnectorFormat.Type = msoConnectorStraight
            'shp.ConnectorFormat.Type = msoConnectorCurve
            shp.ConnectorFormat.Type = msoConnectorElbow
        
            'shp.Line.Weight = 3
        
            'shp.Adjustments.Item(1) = 0.5
        
        End If
   
        'shp.Top = shp.Top + 30 * 1
        'shp.Top = 548
        'shp.Left = shp.Left - 50
   
    Next
End Sub

Private Sub add_dependency_arrows_to_boxes( _
        Optional max As Long = 0, _
        Optional thick As Double = 0, _
        Optional always_connect_right_to_left As Boolean = False, _
        Optional max_time As Integer = 0, _
        Optional sample_every As Integer = 0, _
        Optional precedents_filepath As String = "", _
        Optional precedents_filepath_sep As String = ",")
    '' Connects the boxes created by macro "Sheets_to_boxes", based on each tabs' formulae
    '' Provides a visualization of the relationships between sheets.
    '' Assumes the active sheet contains a box shape for each sheet to be linked,
    '' and that the box's shape.name is the sheet name for the connectors to be
    '' created. Optional arguments:
    ''
    ''  - max: controls the maximum number of connectors per box to avoid clutter.
    ''
    ''  - thick: use a constant thinkness for the connectors. If 0 (default) then the
    ''    thinkess is based on the number of references between sheets.
    ''
    ''  - always_connect_right_to_left: if True, connectors always go from the right
    ''    edge of a box to the left of the next one.
    ''
    ''  - max_time: limits the calculation time spent on each sheet calculating the
    ''    references (links) between sheets.
    ''
    ''  - sample_every: the calculation of the references (links) between sheets is
    ''    done on a sample of the cells only. This is used to limit the calculation
    ''    time spent on each sheet.
    ''
    ''  - precedents_filepath: if specified, the calculation of the references (links)
    ''    between sheets is overriden, and taken from the file path specified. The
    ''    input file must be a CSV file with three columns:
    ''
    ''      + "sheet_name": each sheet's name;
    ''      + "sheet_name_precedent": each sheet's reference (link) to previous sheet
    ''        (eg. as in "Trace precedents"); and
    ''      + "count": number of references or links found between the two sheets.
    ''
    ''    To produce such file:
    ''      1) Run output_formulae (in this module) on the target workbook;
    ''      2) Run output_names (in this module) on the target workbook; and
    ''      3) Run process_formulas.ipynb (jupyter notebook - requires python 3.7+).
    ''
    ''  - precedents_filepath_sep: sparator used in precedents_filepath

    Dim i As Long
    Dim t As Integer, oldStatusBar
    Dim tshts
    Dim trng As Range
    Dim ishp As Shape, fshp As Shape
    Dim sht As Worksheet
    Dim ishpn, n As Long, thickness As Double
    Dim d As Object, x
    Dim iFileNum As Integer, ifileLine As String, iFileRowNum As Integer
    Dim iFileArr() As String
    Dim shtn As String, precedent_shtn As String, ref_count As Long

    t = 1
    oldStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True

    Set tshts = ActiveWorkbook.Worksheets
    'Set tshts = ActiveWindow.SelectedSheets
    
    If precedents_filepath <> "" Then
    
        ' Read the input file
        iFileNum = FreeFile
        Open precedents_filepath For Input As #iFileNum
        iFileRowNum = 1
            
        Do Until EOF(iFileNum)
            Line Input #iFileNum, ifileLine
            ReDim Preserve iFileArr(1 To 3, 1 To iFileRowNum)
            i = 1
            For Each x In Split(ifileLine, precedents_filepath_sep)
                iFileArr(i, iFileRowNum) = x
                i = i + 1
            Next x
            iFileRowNum = iFileRowNum + 1
        Loop
        
        Close #iFileNum
        
    End If

    On Error Resume Next
    For Each sht In tshts
    
        'Debug.Print sht.Name
        Application.StatusBar = t & "/" & tshts.Count & "   " & sht.name
    
        Set fshp = ActiveSheet.Shapes(sht.name)
        If fshp Is Nothing Then GoTo Next_fshp
        Set trng = sht.Cells.SpecialCells(xlCellTypeFormulas)
        If trng Is Nothing Then GoTo Next_fshp
        
        If precedents_filepath <> "" Then
            'Create the reference count dictionary from the input file
            Set d = CreateObject("Scripting.Dictionary")
            For i = 1 To UBound(iFileArr, 2)
                
                shtn = iFileArr(1, i)
                precedent_shtn = iFileArr(2, i)
                ref_count = CInt(iFileArr(3, i))
                
                If sht.name = shtn Then d(precedent_shtn) = ref_count
            Next
        Else
            'Calculate the reference count dictionary
            Set d = precedent_sheetnames_count(trng, max_time, sample_every)
        End If
        If d Is Nothing Then GoTo Next_fshp
    
        i = 0
        If max = 0 Then max = d.Count            'max=0 means no max
        Do While d.Count > 0 And i < max

            For Each ishpn In dict_keys_with_max_values(d)
            
                If ishpn = sht.name Then GoTo Next_ishp
                Set ishp = ActiveSheet.Shapes(ishpn)
                If ishp Is Nothing Then GoTo Next_ishp
            
                n = d(ishpn)
                If thick > 0 Then
                    thickness = thick            'fixed thickness!
                    insert_connector ishp, fshp, thickness, always_connect_right_to_left
                Else
                    thickness = Log(n) + 0.25    'in vba log = ln
                    If thickness >= 1 Then
                        insert_connector ishp, fshp, thickness, always_connect_right_to_left
                    End If
                End If
            
                i = i + 1
            
Next_ishp:
                d.Remove ishpn
                Set ishpn = Nothing
                Set ishp = Nothing
            
            Next
        Loop
    
Next_fshp:
        Set d = Nothing
        Set sht = Nothing
        Set fshp = Nothing
        Set trng = Nothing
        t = t + 1
    Next

    On Error GoTo 0

    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
End Sub

Private Sub insert_connector(ishp As Shape, fshp As Shape, _
                             Optional thickness As Double = 1, _
                             Optional always_connect_right_to_left As Boolean = False)
    
    Dim shp As Shape, name As String
    Dim Colour As Long, RGBarr
    Dim iConnectPt As Integer, fConnectPt As Integer
    
    ''points
    Const l = 0
    Const t = 0
    Const w = 10
    Const h = 10
    
    name = ishp.name & " to " & fshp.name
    If shape_exists(name) Then Exit Sub          'do not overwrite
    
    'Set shp = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, l, t, w, h) 'straight
    'Set shp = ActiveSheet.Shapes.AddConnector(msoConnectorElbow, l, t, w, h) ' elbows
    Set shp = ActiveSheet.Shapes.AddConnector(msoConnectorCurve, l, t, w, h) 'straight
    
    shp.name = name
    
    shp.Placement = xlFreeFloating
    
    shp.Line.EndArrowheadStyle = msoArrowheadTriangle
    
    iConnectPt = 4                               'boxes' right side
    fConnectPt = 2                               'boxes' left side
    
    If Not always_connect_right_to_left Then
        If ishp.Left > fshp.Left Then fConnectPt = 4
    End If
    
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

Private Function shape_exists(name As String)
    ''Reyurns True if a shape exists, False otherwise
    Dim shp As Shape
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(name)
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

Private Function precedent_sheetnames_count(trng As Range, _
                                            Optional max_time As Integer = 180, _
                                            Optional sample_every As Integer = 10)
    ''Returns a dictionary of {sheet_name: count_cells_using_sheet_name}
    ''of the sheets used in rng's formulas. max_time and sample_every can be
    ''used to limit the time spent counting precedents:
    '' max_time: maximum total time spent
    '' sample_every: sampling on trng (e.g. 1 out of every 10 cells)

    Dim rng As Range
    Dim shtns, shtn
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    On Error GoTo FinishThis

    Dim Start As Double, Duration As Double
    Start = Timer

    Dim i As Long
    For i = 1 To trng.Cells.Count
        If (sample_every > 0) Then If ((i Mod sample_every) <> 1) Then GoTo NextIteration
        Set rng = trng.Cells(i)
        Duration = Timer - Start
        If (max_time = 0) Or (Duration < max_time) Then
            shtns = precedent_sheetnames(rng)
            For Each shtn In shtns
                d(shtn) = d(shtn) + 1
            Next
        Else
            Exit For
        End If
NextIteration:
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
    RE.Pattern = "'(.+?)'!|\b([^/\*-+ =&<>\[\]""\(\)]+)!"
    ''This approach is weak, but sufficient for this purpose

    RE.Global = True
    Set matches = RE.Execute(rng.Formula)

    For Each m In matches
        For Each sm In m.submatches
            If CStr(sm) <> "" Then d(sm) = 1
        Next
    Next

    precedent_sheetnames = d.Keys()
    Set RE = Nothing
    Set d = Nothing
End Function

Private Function precedent_sheetnames_unreliable_alternative(rng As Range) As Variant
    'This version explores the use of .ShowPrecedents, but it is slow and unreliable
    'LEFT HERE AS A WARNING. DO NOT USE.
    'Returns an array with the unique names of all the sheets used in rng.Formula
    
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    
    Dim xarrow As Long, xlink As Long, prng As Range
    
    rng.ShowPrecedents
    ActiveWindow.SmallScroll
    Application.WindowState = Application.WindowState
    On Error Resume Next
    xarrow = 1
    Do
        xlink = 1
        Do
            Set prng = Nothing
            Set prng = Selection.NavigateArrow(True, xarrow, xlink)
            ' Go back to input range
            rng.Parent.Select
            rng.Select
            If (prng Is Nothing) _
                Or ((prng.Parent.name = rng.Parent.name) _
                    And (prng.Address = rng.Address)) Then
                Exit Do
            End If
            ' Avoid internal precedents
            If Not prng.Parent.name = rng.Parent.name Then d(prng.Parent.name) = 1
            xlink = xlink + 1
        Loop
        If Not prng Is Nothing Then
            If ((prng.Parent.name = rng.Parent.name) _
                And (prng.Address = rng.Address)) Then
                Exit Do
            End If
        End If
        xarrow = xarrow + 1
    Loop
    On Error GoTo 0
    rng.Parent.ClearArrows

    precedent_sheetnames = d.Keys()
    Set d = Nothing
    
End Function


Private Function dict_keys_with_max_values(d) As Variant
    'Returns an array of dict keys whose values are the dict's maximum

    Dim i As Long
    Dim arr()
    Dim max As Long
    Dim key As Variant

    i = 0
    max = Application.max(d.items)
    For Each key In d.Keys
        If d(key) = max Then
            ReDim Preserve arr(i)
            arr(i) = key
            i = i + 1
        End If
    Next key
    dict_keys_with_max_values = arr
End Function


Sub output_formulae()

    Dim dfile As Integer
    Dim dfilep As String
    Dim wpath As String
    
    dfile = FreeFile ''Assigns the next free file number
    wpath = ActiveWorkbook.path & "\"
    dfilep = wpath & Left(ActiveWorkbook.name, (InStrRev(ActiveWorkbook.name, ".", -1, vbTextCompare) - 1)) & "_formulas.tsv"
    
    ''Delete the file if it exists:
    If Len(Dir$(dfilep)) > 1 Then
        SetAttr dfilep, vbNormal
        Kill dfilep
    End If
    
    ''Open the file for appending
    Open dfilep For Append As #dfile
    
    Dim sht As Worksheet, rng As Range, trng As Range
    Dim rowtxt As String
    Const sep = vbTab
    
    rowtxt = """sheet_name""" & sep & """cell_address""" & sep & """cell_formula"""
    Print #dfile, rowtxt
    
    For Each sht In ActiveWorkbook.Worksheets
        On Error Resume Next
        Set trng = sht.Cells.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not trng Is Nothing Then
            For Each rng In trng
                rowtxt = Join(Array("""" & sht.name & """", """" & rng.Address(False, False) & """", """'" & rng.Formula & """"), sep)
                Print #dfile, rowtxt
            Next
        End If
        Set trng = Nothing
    Next
    
    Close #dfile
    
End Sub


Sub output_names()

    Dim dfile As Integer
    Dim dfilep As String
    Dim wpath As String
    
    dfile = FreeFile ''Assigns the next free file number
    wpath = ActiveWorkbook.path & "\"
    dfilep = wpath & Left(ActiveWorkbook.name, (InStrRev(ActiveWorkbook.name, ".", -1, vbTextCompare) - 1)) & "_names.tsv"
    
    ''Delete the file if it exists:
    If Len(Dir$(dfilep)) > 1 Then
        SetAttr dfilep, vbNormal
        Kill dfilep
    End If
    
    ''Open the file for appending
    Open dfilep For Append As #dfile
    
    Dim rowtxt As String
    Const sep = vbTab
    
    rowtxt = """name""" & sep & """range"""
    Print #dfile, rowtxt
    
    Dim name As name
    
    For Each name In ActiveWorkbook.Names
        rowtxt = Join(Array("""" & name.name & """", """'" & name.RefersTo & """"), sep)
        Print #dfile, rowtxt
    Next
    
    Set name = Nothing
    Close #dfile
    
End Sub

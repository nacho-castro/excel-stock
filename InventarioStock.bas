Attribute VB_Name = "Módulo1"
Sub BORRARinventario()
Attribute BORRARinventario.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BORRARinventario Macro
'

'
    Range("B6").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B7").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B8").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B9").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B10").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B11").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B6").Select
    ActiveCell.FormulaR1C1 = ""
End Sub
Sub GUARDARinventario()
Attribute GUARDARinventario.VB_ProcData.VB_Invoke_Func = " \n14"
'
' GUARDARinventario Macro
'

'
    Application.ScreenUpdating = False

    Application.DisplayAlerts = False

    ActiveSheet.Unprotect "124"
    
    Rows("16:16").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Range("K16").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("B6:B12").Select
    Selection.Copy
    Range("A16").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("INVENTARIO").ListObjects("inventario").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("INVENTARIO").ListObjects("inventario").Sort. _
        SortFields.Add Key:=Range("inventario[[#All],[CÓDIGO]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("INVENTARIO").ListObjects("inventario").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("INVENTARIO").ListObjects("inventario").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("INVENTARIO").ListObjects("inventario").Sort. _
        SortFields.Add Key:=Range("inventario[[#All],[CÓDIGO]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("INVENTARIO").ListObjects("inventario").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("H16:J16").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("B6").Select
    
    Application.DisplayAlerts = True

    Application.ScreenUpdating = True

    ActiveSheet.Protect "124"
End Sub
Sub GUARDARentrada()
Attribute GUARDARentrada.VB_ProcData.VB_Invoke_Func = " \n14"
'
' GUARDARentrada Macro
'

'
    Application.ScreenUpdating = False

    Application.DisplayAlerts = False

    ActiveSheet.Unprotect "124"
    
    Rows("16:16").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.Font.Bold = False
    Range("B6:B13").Select
    Selection.Copy
    Range("A16").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=True
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveWorkbook.Worksheets("ENTRADA").ListObjects("entrada").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("ENTRADA").ListObjects("entrada").Sort.SortFields. _
        Add Key:=Range("entrada[[#Headers],[#Data],[FECHA]]"), SortOn:= _
        xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ENTRADA").ListObjects("entrada").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("ENTRADA").ListObjects("entrada").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("ENTRADA").ListObjects("entrada").Sort.SortFields. _
        Add Key:=Range("entrada[[#Headers],[#Data],[FECHA]]"), SortOn:= _
        xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ENTRADA").ListObjects("entrada").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B6").Select
    
    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = True
    
    ActiveSheet.Protect "124"
    
End Sub
Sub BORRARentrada()
Attribute BORRARentrada.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BORRARentrada Macro
'

'
    Range("B6").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B7").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B6").Select
    ActiveCell.FormulaR1C1 = ""
End Sub
Sub GUARDARsalida()
Attribute GUARDARsalida.VB_ProcData.VB_Invoke_Func = " \n14"
'
' GUARDARsalida Macro
'

'
    Application.ScreenUpdating = False

    Application.DisplayAlerts = False

    ActiveSheet.Unprotect "124"
    
    Rows("16:16").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Selection.Font.Bold = False
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B6:B13").Select
    Selection.Copy
    Range("A16:H16").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("SALIDA").ListObjects("salida").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SALIDA").ListObjects("salida").Sort.SortFields.Add _
        Key:=Range("salida[[#Headers],[#Data],[FECHA]]"), SortOn:=xlSortOnValues, _
        Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("SALIDA").ListObjects("salida").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("SALIDA").ListObjects("salida").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SALIDA").ListObjects("salida").Sort.SortFields.Add _
        Key:=Range("salida[[#Headers],[#Data],[FECHA]]"), SortOn:=xlSortOnValues, _
        Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("SALIDA").ListObjects("salida").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B6").Select
    
    Application.DisplayAlerts = True

    Application.ScreenUpdating = True

    ActiveSheet.Protect "124"
    
End Sub
Sub BORRARsalida()
Attribute BORRARsalida.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BORRARsalida Macro
'

'
    Range("B6").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B7").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B6").Select
    ActiveCell.FormulaR1C1 = ""
End Sub

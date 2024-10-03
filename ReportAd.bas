Function Sh_Exist(wb As Workbook, sName As String) As Boolean
    Dim wsSh As Worksheet
    On Error Resume Next
    Set wsSh = wb.Sheets(sName)
    Sh_Exist = Not wsSh Is Nothing
End Function


Sub copy_past(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(2, 1), Cells(lLastRow, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Sheets(nameOfSheet1).Activate
        Range(Cells(1, 1), Cells(lLastRow, 5)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close

End Sub



Sub copy_past_GA(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(1, 1), Cells(lLastRow + 1, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close

End Sub


Sub copy_past_gDE(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(2, 1), Cells(lLastRow, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(7, 1), Cells(lLastRow, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close

End Sub

Sub copy_past_gDE_bot(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(2, 1), Cells(lLastRow, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 3).End(xlUp).Row
        Range(Cells(13, 3), Cells(lLastRow, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(2, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close

End Sub
Sub copy_past_gDE_geo(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
        Range(Cells(1, 2), Cells(lLastRow - 1, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(5, 1), Cells(lLastRow, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(1, 2).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
        Range("C31").Select ' протягивание формулы, значения
        Selection.AutoFill Destination:=Range(Cells(31, 3), Cells(31, lLastCol + 1)), Type:=xlFillDefault
    Windows(nameOfFile).Close

End Sub
Sub copy_past_gDE_realization(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(4, 1), Cells(lLastRow, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(5, 1), Cells(lLastRow - 1, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(2, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
        Range("C1").Select ' протягивание формулы, значения
        Selection.AutoFill Destination:=Range(Cells(1, 3), Cells(1, lLastCol)), Type:=xlFillDefault
    Windows(nameOfFile).Close

End Sub

Sub copy_past_gDE_TA_placement(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
        Range(Cells(3, 1), Cells(lLastRow, 9)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(8, 1), Cells(lLastRow, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(2, 2).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close

End Sub

Sub copy_past_gDE_TA_total(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)

    lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    Windows(nameOfGeneralFile).Activate
    ActiveWorkbook.Save
    'ChDir (pathDir & "\Svod")
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastrowA = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(8, 1), Cells(lLastrowA, 1)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(lLastRow + 1, 2).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Activate
        Sheets(nameOfSheet1).Activate
        lLastrowA = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(8, 2), Cells(lLastrowA, 7)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(lLastRow + 1, 4).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
        Cells(lLastRow, 3).Select
        Selection.AutoFill Destination:=Range(Cells(lLastRow, 3), Cells(lLastRow + lLastrowA - 7, 3)) ' протягивание значения
        Range("A2").Select ' протягивание формулы, значения
        Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(lLastRow + lLastrowA - 7, 1)), Type:=xlFillDefault
        Range("J1").Select ' протягивание формулы, значения
        Selection.AutoFill Destination:=Range(Cells(1, 10), Cells(lLastRow + lLastrowA - 7, 10)), Type:=xlFillDefault
    Windows(nameOfFile).Close
        
End Sub

Sub copy_past_gDE_Complete_Views(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(1, 2), Cells(lLastRow, 4)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(7, 1), Cells(lLastRow, 4)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(1, 2).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
        Range("A2").Select ' протягивание формулы, значения
        Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(lLastRow - 8, 1)), Type:=xlFillDefault
    Windows(nameOfFile).Close

End Sub
Sub copy_past_Frequency_week_1(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        Range(Cells(2, 4), Cells(16, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        Range(Cells(6, 1), Cells(20, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(2, 4).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close

End Sub
Sub copy_past_Frequency_week_2(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        Range(Cells(19, 4), Cells(33, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        Range(Cells(6, 1), Cells(20, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(19, 4).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close

End Sub
Sub copy_past_Frequency_week_3(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        Range(Cells(36, 4), Cells(50, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        Range(Cells(6, 1), Cells(20, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(36, 4).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close

End Sub

Sub copy_past_Frequency_week_4(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        Range(Cells(53, 4), Cells(67, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        Range(Cells(6, 1), Cells(20, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(53, 4).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close

End Sub
Sub copy_past_gde_socdem_feature_placement(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(6, 2), Cells(lLastRow + 1, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close

End Sub
Sub copy_past_gde_socdem_TG_placement(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(7, 1), Cells(lLastRow + 1, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Cells(lLastRow + 1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close

End Sub
Sub copy_past_gde_socdem_TG_total(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Cells(lLastRow, lLastCol).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
        Cells(lLastRow, 7).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close

End Sub
Sub copy_past_gde_socdem_total(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Cells(7, lLastCol).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
        Cells(lLastRow, 13).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Activate
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        Cells(8, lLastCol).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
        Cells(lLastRow, 12).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Activate
        Sheets(nameOfSheet1).Activate
        Range(Cells(9, 4), Cells(14, 4)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
        Cells(lLastRow, 17).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=True
    Windows(nameOfFile).Close

End Sub

Sub copy_past_SD_TNS(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim XCell As Object, YCell As Object
    Dim XCol, XRow, YCol, YRow, ZCol, ZRow As Integer
    txtCol1 = "По полу"
    txtCol2 = "По возрасту"
    txtCol3 = "По полу/возрасту"
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        Set XCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol1)
        Set YCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol2)
        Set ZCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol3)
    If XCell Is Nothing Then
    MsgBox ("Не найдены ячейки в TNS" & txtCol1 & txtCol1 & txtCol1)
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    YCol = YCell.Column
    YRow = YCell.Row
    ZCol = YCell.Column
    ZRow = YCell.Row
    Cells(1, 1).Value = 100
    Cells(1, 1).Copy
    Range("I" & (XRow + 3) & ":" & "I" & (YCell.Row - 2)).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlDivide, SkipBlanks _
        :=False, Transpose:=False
    ActiveWorkbook.Worksheets(nameOfSheet1).Range("A" & (XRow + 3) & ":" & "I" & (YCell.Row - 2)).Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(2, 1).Select
        ActiveSheet.Paste
    Windows(nameOfFile).Activate
    Cells(1, 1).Copy
    Range("I" & (YRow + 3) & ":" & "I" & (ZCell.Row - 2)).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlDivide, SkipBlanks _
        :=False, Transpose:=False
    ActiveWorkbook.Worksheets(nameOfSheet1).Range("A" & (YRow + 3) & ":" & "I" & (ZCell.Row - 2)).Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Range("A" & (YRow - XRow - 2)).Select
        ActiveSheet.Paste
    End If
    
    Windows(nameOfFile).Close

End Sub

Sub copy_past_DCM(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Dim XCell As Object
    Dim XRow, XCol As Integer
    
    txtCol1 = "Поля отчета"
    
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(3, 1), Cells(lLastRow, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        
        '-----поиск начала диапазона данных---------
                
        Set XCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol1)
        XCol = XCell.Column
        XRow = XCell.Row
        '------копирование диапазона------
        
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(XRow + 1, 1), Cells(lLastRow, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(1, 3).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
        lLastRow = Cells(Rows.Count, 3).End(xlUp).Row
        Range("A2:B2").Select ' протягивание формулы, значения
        Selection.AutoFill Destination:=Range("A2" & ":" & "B" & lLastRow), Type:=xlFillDefault
    Windows(nameOfFile).Close

End Sub
Sub copy_past_DCM_frequency(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Dim XCell As Object
    Dim XRow, XCol As Integer
    
    txtCol1 = "Поля отчета"
    
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(5, 1), Cells(lLastRow, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        
        '-----поиск начала диапазона данных---------
                
        Set XCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol1)
        XCol = XCell.Column
        XRow = XCell.Row
        '------копирование диапазона------
        
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(XRow + 1, 1), Cells(lLastRow, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(3, 4).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
        lLastRow = Cells(Rows.Count, 4).End(xlUp).Row
        Range("A4:C4").Select ' протягивание формулы, значения
        Selection.AutoFill Destination:=Range("A4" & ":" & "C" & lLastRow), Type:=xlFillDefault
        Range("AB3:AD3").Select
        Selection.AutoFill Destination:=Range("AB3:AD14"), Type:=xlFillDefault
    Windows(nameOfFile).Close

End Sub



Sub copy_past_DCM_reach(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Dim XCell, YCell, ZCell, IdYCell, IdZCell, nonCell As Object
    Dim XRow, XCol, YCol, YRow, ZCol, ZRow, IdYCol, IdZCol As Integer
    
    txtCol1 = "Поля отчета"
    txtCol2 = "Место размещения"
    txtCol3 = "Объявление"
    txtCol4 = "Идентификатор места размещения"
    txtCol5 = "Идентификатор объявления"
    txtCol6 = "(not set)"
    
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(3, 1), Cells(lLastRow, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        
    '------вставить столбец--------
    
        Cells(1, 1).EntireColumn.Insert
        
    '-----поиск начала диапазона данных---------
    
        Set XCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol1)
        Set YCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol2)
        Set ZCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol3)
        Set IdYCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol4)
        Set IdZCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol5)
        Set nonCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol6)
        YCol = YCell.Column
        YRow = YCell.Row
        XCol = XCell.Column
        XRow = XCell.Row
        ZCol = ZCell.Column
        ZRow = ZCell.Row
        IdYCol = IdYCell.Column
        IdZCol = IdZCell.Column
       
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
        
    '------проверка на наличие нон-сет ---------
    
    
        If nonCell Is Nothing Then
        Else
    
    '------сцепляем ключ---------
        With Range(Cells(XRow + 1, IdYCol), Cells(lLastRow, IdYCol))
                NumberFormat = "0"
                .Value = .Value
        End With
        With Range(Cells(XRow + 1, IdZCol), Cells(lLastRow, IdZCol))
                NumberFormat = "0"
                .Value = .Value
        End With
    
        Cells(XRow + 1, 1).Select
        For i = lLastRow To XRow Step -1
                Sheets("Данные").Cells(i, 1).Value = Cells(i, IdYCol).Value & "@" & Cells(i, IdZCol).Value & "@"
        Next
        
        Range(Cells(XRow + 1, 1), Cells(lLastRow, lLastCol)).Select
        Selection.Copy
        Sheets.Add After:=ActiveSheet
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Sheets("Лист1").Activate
        
    '-----поиск ключевого диапазона(для фильтрации)-------
    
        Set YCell = Workbooks(nameOfFile).Worksheets("Лист1").Cells.Find(txtCol2)
        YCol = YCell.Column
        YRow = YCell.Row
    '-----чистка ненужных данных-------
    
        lLastRow = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row
                
            For i = lLastRow To 2 Step -1
                If Sheets("Лист1").Cells(i, YCol) = "(not set)" Then Sheets("Лист1").Rows(i).Delete Shift:=xlUp
            Next
        Sheets("Данные").Select
        
    '-----протяжка ВПР------
    
        Set XCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol1)
        Set YCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol2)
        Set ZCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol3)
        XCol = XCell.Column
        XRow = XCell.Row
        YCol = YCell.Column
        YRow = YCell.Row
        ZCol = ZCell.Column
        ZRow = ZCell.Row

        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row
        
        For i = lLastRow To XRow + 2 Step -1
            If Sheets("Данные").Cells(i, YCol) <> "(not set)" Then Rows(i).Delete Shift:=xlUp
        Next
        

        lLastRow = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        
        Cells(XRow + 1, YCol).Select
        
            For i = lLastRow To XRow + 1 Step -1
        
                Cells(i, YCol) = Application.VLookup(Cells(i, 1), Workbooks(nameOfFile).Sheets("Лист1").Range( _
                                                                    Workbooks(nameOfFile).Sheets("Лист1").Cells(1, 1), _
                                                                    Workbooks(nameOfFile).Sheets("Лист1").Cells(lLastRow + XRow + 1, YCol)), YCol, False)
                                                                    
        
            Next
            
        Cells(XRow + 1, ZCol).Select
            For i = lLastRow To XRow + 1 Step -1
    
                Cells(i, ZCol) = Application.VLookup(Cells(i, 1), Workbooks(nameOfFile).Sheets("Лист1").Range( _
                                                                    Workbooks(nameOfFile).Sheets("Лист1").Cells(1, 1), _
                                                                    Workbooks(nameOfFile).Sheets("Лист1").Cells(lLastRow + XRow + 1, ZCol)), ZCol, False)
    
            Next
        
    
    End If
    '-------копирование готовых данных в отчет--------
    
        lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
        lLastCol = Cells.SpecialCells(xlLastCell).Column
    
        Range(Cells(XRow + 1, 2), Cells(lLastRow, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(1, 2).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
        Range("A2").Select ' протягивание формулы, значения
        Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(lLastRow - XRow, 1)), Type:=xlFillDefault
    Windows(nameOfFile).Close

End Sub

Sub copy_past_DCM_TA(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Dim XCell As Object
    Dim YCell As Object
    Dim XRow, YRow, YCol As Integer
    
    txtCol1 = "Поля отчета"
    txtCol = "Дней с момента связанного взаимодействия"
    
    Dim arrayOfPVC(1 To 100000) As Boolean
    
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 4).End(xlUp).Row
        Range(Cells(4, 1), Cells(lLastRow, lLastCol)).Clear
        Sheets("Total").Select
        Range("K7").Copy
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
        
        Set XCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol)
        Set YCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol1)
        XCol = XCell.Column
        XRow = XCell.Row
        YCol = YCell.Column
        YRow = YCell.Row
        
        Range("c1").Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
            
            If XCell Is Nothing Then
                MsgBox ("Не найдено количество дней для анализа PVC")
                Else
                With Range(Cells(XRow + 1, XCol), Cells(lLastRow, XCol))
                NumberFormat = "0"
                .Value = .Value
                End With
            End If
         
            For i = YRow + 2 To lLastRow Step 1
                arrayOfPVC(i) = False
            Next
            
            For i = lLastRow To YRow + 2 Step -1
            
            If Not arrayOfPVC(i) And Cells(i, XCol).Value > Range("c1").Value Then arrayOfPVC(i) = True
            Next
    
            For i = lLastRow To YRow + 2 Step -1
                If arrayOfPVC(i) Then Rows(i).Delete
         
            Next
            
        Range(Cells(YRow + 1, 1), Cells(lLastRow, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(1, 4).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
      
        lLastRow = Cells(Rows.Count, 4).End(xlUp).Row
            
        Range("A2:C2").Select ' протягивание формулы, значения
        Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(lLastRow, 3)), Type:=xlFillDefault
        Range("S1:U1").Select ' протягивание формулы, значения
        Selection.AutoFill Destination:=Range(Cells(1, 19), Cells(lLastRow, 21)), Type:=xlFillDefault
        Range("W3").Select ' протягивание формулы, значения
        Selection.AutoFill Destination:=Range(Cells(3, 23), Cells(lLastRow, 23)), Type:=xlFillDefault
    Windows(nameOfFile).Close

End Sub



Sub copy_past_DCM_geo_data(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Dim XCell As Object
    Dim XRow As Integer
    
    txtCol1 = "Поля отчета"
    
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 3).End(xlUp).Row
        Range(Cells(3, 1), Cells(lLastRow, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
        
        '-----поиск начала диапазона данных---------
                
        Set XCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol1)
        XRow = XCell.Row
        '------копирование диапазона------
        
        
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(XRow + 1, 1), Cells(lLastRow - 1, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(1, 3).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
        lLastRow = Cells(Rows.Count, 3).End(xlUp).Row
        Range("A2").Select ' протягивание формулы, значения
        Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(lLastRow, 1)), Type:=xlFillDefault
        Range("B2").Select ' протягивание формулы, значения
        Selection.AutoFill Destination:=Range(Cells(2, 2), Cells(lLastRow, 2)), Type:=xlFillDefault
    Windows(nameOfFile).Close

End Sub

Sub copy_past_DCM_bot(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
    Dim lLastRow As Long
    Dim lLastCol As Long
    Dim XCell As Object
    Dim XRow As Integer
    
    txtCol1 = "Поля отчета"
    
    Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    'ChDir (pathDir & "\Svod")
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
        Range(Cells(3, 1), Cells(lLastRow, lLastCol)).Clear
    Workbooks.Open (pathDir & "\Data\" & nameOfFile)  'Открытие файла
        Sheets(nameOfSheet1).Activate
                
        '-----поиск начала диапазона данных---------
                
        Set XCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Cells.Find(txtCol1)
        XRow = XCell.Row
        '------копирование диапазона------
        
        lLastCol = Cells.SpecialCells(xlLastCell).Column
        lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(XRow + 1, 1), Cells(lLastRow, lLastCol)).Select
        Selection.Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(1, 2).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
        lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
        Range("A2").Select ' протягивание формулы, значения
        Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(lLastRow, 1)), Type:=xlFillDefault
    Windows(nameOfFile).Close

End Sub




Sub generalPast()    'Replace "Formuls" with the name of the sheet to be copied.
    Dim iTimer As Single
        iTimer = Timer
    Dim nameOfGeneralFile As String
    Dim nameOfPathGeneralFile As String
    nameOfPathGeneralFile = ActiveWorkbook.Path
    nameOfGeneralFile = ActiveWorkbook.Name
    Application.ScreenUpdating = False 'отключение обновления экрана
    Workbooks.Application.DisplayAlerts = False ' отключение всплывающих окон
    Workbooks(nameOfGeneralFile).Save
    
    If ActiveWorkbook.Worksheets("Total").Range("K24").Value = "+" Then Call copy_past_gDE("gDE.xlsx", "Рейтинги", "gDE", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K25").Value = "+" Then Call copy_past_gDE_Complete_Views("gDe Complete Views.xlsx", "Рейтинги", "gDE Complete Views", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K26").Value = "+" Then Call copy_past_gDE_TA_placement("gDE TA_placement.xlsx", "Рейтинги", "gDE_TA", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K26").Value = "+" Then Call copy_past_gDE_TA_total("gDE TA_total.xlsx", "Рейтинги", "gDE_TA", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K27").Value = "+" Then Call copy_past_gDE_realization("gDE Realization.xlsx", "Временные рамки", "gDE Realization", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K28").Value = "+" Then Call copy_past_gDE_geo("gDE Geo.xlsx", "Technical", "gDE Geo", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K29").Value = "+" Then Call copy_past_gDE_bot("gDE Bot.xlsx", "gde2", "gDE Bot", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K30").Value = "+" Then Call copy_past_Frequency_week_1("1week.xlsx", "Распределение", "Frequency", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K30").Value = "+" Then Call copy_past_Frequency_week_2("2week.xlsx", "Распределение", "Frequency", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K30").Value = "+" Then Call copy_past_Frequency_week_3("3week.xlsx", "Распределение", "Frequency", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K30").Value = "+" Then Call copy_past_Frequency_week_4("4week.xlsx", "Распределение", "Frequency", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K32").Value = "+" Then Call copy_past_GA("GA.xlsx", "Набор данных1", "Analytics", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K33").Value = "+" Then Call copy_past("TNS.xlsx", "Sheet1", "TNS", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K34").Value = "+" Then Call copy_past_SD_TNS("TNS_socdem_pixels.xlsx", "Отчет по профилю", "SD TNS", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K31").Value = "+" Then Call copy_past_gde_socdem_feature_placement("gde_socdem_feature_placement.xlsx", "Целевые группы", "SD gDE", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K31").Value = "+" Then Call copy_past_gde_socdem_TG_placement("gde_socdem_TG_placement.xlsx", "Целевые группы", "SD gDE", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K31").Value = "+" Then Call copy_past_gde_socdem_TG_total("gde_socdem_TG_total.xlsx", "Целевые группы", "Audience", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K31").Value = "+" Then Call copy_past_gde_socdem_total("gde_socdem_total.xlsx", "Соц-дем.", "Audience", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K36").Value = "+" Then Call copy_past_DCM("DCM.xlsx", "Данные", "DCM", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K36").Value = "+" Then Call copy_past_DCM_frequency("DCM_frequency.xlsx", "Данные", "DCM_frequency", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K37").Value = "+" Then Call copy_past_DCM_reach("DCM_reach_place_creativ_platfor.xlsx", "Данные", "DCM_reach_place_creativ_platfor", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K37").Value = "+" Then Call copy_past_DCM_reach("DCM_reach_place_creativ.xlsx", "Данные", "DCM_reach_place_creativ", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K37").Value = "+" Then Call copy_past_DCM("DCM_reach_place.xlsx", "Данные", "DCM_reach_place", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K37").Value = "+" Then Call copy_past_DCM("DCM_reach_place_platforms.xlsx", "Данные", "DCM_reach_place_platforms", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K37").Value = "+" Then Call copy_past_DCM("DCM_reach_platforms.xlsx", "Данные", "DCM_reach_platforms", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K38").Value = "+" Then Call copy_past_DCM_TA("DCM_TA.xlsx", "Данные", "DCM_TA", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K39").Value = "+" Then Call copy_past_DCM_geo_data("DCM_geo_data.xlsx", "Данные", "DCM_geo_data", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Total").Range("K40").Value = "+" Then Call copy_past_DCM_bot("DCM_bot.xlsx", "Данные", "DCM_bot", nameOfPathGeneralFile, nameOfGeneralFile)
    

    Windows(nameOfGeneralFile).Activate ' активация книги откуда запущен макрос
    Sheets("Plan-Fact").Select
    
    Application.DisplayAlerts = 1
    Application.ScreenUpdating = True
    MsgBox "Time finish makros: " & Format((Timer - iTimer) / 86400, "Long Time"), vbExclamation, "" ' таймер час-мин-сек
    
End Sub
